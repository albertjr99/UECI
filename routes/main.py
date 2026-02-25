from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify, send_file
from flask_login import login_required, current_user
from models import db
from models.planilha import PlanilhaData, AbaConfig
from models.configuracoes import DropdownConfig
from models.user import User
from models.link_temporario import LinkTemporario  # ✅ ADICIONAR ESTA LINHA
from datetime import datetime, timedelta
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
import secrets
import hashlib
import unicodedata

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
@login_required
def index():
    primeira_aba = AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).first()
    if primeira_aba:
        return redirect(url_for('main.view_planilha', aba_name=primeira_aba.aba_name))
    return render_template('error.html', message='Nenhuma planilha configurada')


@main_bp.route('/dashboard')
@login_required
def dashboard():
    abas = AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).all()
    total_registros = PlanilhaData.query.count()
    usuarios_ativos = User.query.filter_by(is_active=True).count()

    return render_template('dashboard.html',
                         abas=abas,
                         total_registros=total_registros,
                         usuarios_ativos=usuarios_ativos)

@main_bp.route('/calendario')
@login_required
def calendario():
    """Página do calendário de prazos"""
    return render_template('calendario.html')


# ✅ REMOVER A FUNÇÃO DUPLICADA "gerar_link_preenchimento" (linha 48-82)
# ✅ MANTER APENAS A FUNÇÃO "gerar_link_area" (linha 569-604)


@main_bp.route('/preencher/<token>', methods=['GET', 'POST'])
def preenchimento_externo(token):
    """Página de preenchimento externo (sem login)"""

    # Buscar link temporário
    link = LinkTemporario.query.filter_by(token=token).first()

    if not link:
        return render_template('preenchimento_externo.html', erro='Link inválido')

    # Verificar expiração
    if datetime.utcnow() > link.expiracao:
        return render_template('preenchimento_externo.html',
                             link_expirado=True,
                             expiracao=link.expiracao.strftime('%d/%m/%Y às %H:%M'))

    # Buscar registro
    registro = PlanilhaData.query.get(link.registro_id)
    if not registro:
        return render_template('preenchimento_externo.html', erro='Registro não encontrado')

    # Buscar configuração da aba
    aba_config = AbaConfig.query.filter_by(aba_name=link.aba_name).first()
    if not aba_config:
        return render_template('preenchimento_externo.html', erro='Configuração não encontrada')

    # Pegar TODAS as colunas e separar editáveis
    colunas_todas = aba_config.get_columns()
    orange_columns = aba_config.get_orange_columns()
    colunas_editaveis = [col for col in colunas_todas if (col.get('name') if isinstance(col, dict) else col) in orange_columns]
    colunas_editaveis_nomes = [col.get('name') if isinstance(col, dict) else col for col in colunas_editaveis]

    if request.method == 'POST':
        # Processar preenchimento APENAS das colunas editáveis
        dados_atuais = registro.get_data()

        for col in colunas_editaveis:
            col_name = col.get('name') if isinstance(col, dict) else col
            valor = request.form.get(col_name, '').strip()

            # Converter data se necessário
            col_type = col.get('type', 'text') if isinstance(col, dict) else 'text'
            if col_type == 'date' and valor:
                try:
                    if '/' in valor:
                        date_parts = valor.split('/')
                        if len(date_parts) == 3:
                            valor = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
                except:
                    pass

            dados_atuais[col_name] = valor

        registro.set_data(dados_atuais)
        registro.updated_at = datetime.utcnow()

        # Marcar link como usado
        link.usado_em = datetime.utcnow()

        db.session.commit()

        return render_template('preenchimento_externo.html',
                             sucesso=True,
                             mensagem='✅ Dados salvos com sucesso! Você já pode fechar esta página.')

    # GET: Mostrar formulário
    return render_template('preenchimento_externo.html',
                         colunas_todas=colunas_todas,
                         colunas_editaveis=colunas_editaveis,
                         colunas_editaveis_nomes=colunas_editaveis_nomes,
                         dados_atuais=registro.get_data())

@main_bp.route('/preencher/<token>/gerar-pdf')
def gerar_pdf_preenchimento(token):
    """Gera PDF do preenchimento"""
    import re
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_JUSTIFY
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from io import BytesIO
    from xml.sax.saxutils import escape

    # Buscar link
    link = LinkTemporario.query.filter_by(token=token).first()
    if not link:
        flash('Link inválido', 'danger')
        return redirect(url_for('main.index'))

    # Buscar registro
    registro = PlanilhaData.query.get(link.registro_id)
    if not registro:
        flash('Registro não encontrado', 'danger')
        return redirect(url_for('main.index'))

    # Buscar aba
    aba_config = AbaConfig.query.filter_by(aba_name=link.aba_name).first()
    if not aba_config:
        flash('Configuração não encontrada', 'danger')
        return redirect(url_for('main.index'))

    # Criar PDF em memória
    buffer = BytesIO()

    # Configurar documento em paisagem
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=1*cm,
        leftMargin=1*cm,
        topMargin=2*cm,
        bottomMargin=1.5*cm
    )

    # Elementos do PDF
    elements = []
    styles = getSampleStyleSheet()

    # Estilo para título
    titulo_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=14,
        textColor=colors.HexColor('#1e293b'),
        spaceAfter=10,
        alignment=1
    )

    # Estilo para subtítulo
    subtitulo_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#64748b'),
        spaceAfter=15,
        alignment=1
    )

    # Cabeçalho
    campo_style = ParagraphStyle(
        'CampoPDF',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.whitesmoke
    )

    valor_style = ParagraphStyle(
        'ValorPDF',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#1e293b'),
        alignment=TA_JUSTIFY,
        wordWrap='CJK'
    )

    def limpar_texto_pdf(valor):
        """Remove espaços/quebras excessivas sem alterar o conteúdo."""
        texto = str(valor).replace('\r', ' ').replace('\n', ' ')
        texto = re.sub(r'\s+', ' ', texto).strip()
        return texto or '-'

    elements.append(Paragraph("UNIDADE EXECUTORA DE CONTROLE INTERNO - UECI", titulo_style))
    elements.append(Paragraph(f"Plano de Respostas - {aba_config.aba_name}", subtitulo_style))
    elements.append(Spacer(1, 0.3*cm))

    # Preparar dados da tabela
    dados_registro = registro.get_data()
    colunas = aba_config.get_columns()

    # Criar lista de dados para a tabela (formato vertical)
    table_data = []

    for col in colunas:
        col_name = col.get('name') if isinstance(col, dict) else col
        valor = dados_registro.get(col_name, '-')

        col_name_safe = escape(limpar_texto_pdf(col_name))
        valor_safe = escape(limpar_texto_pdf(valor))

        # Adicionar linha [Campo | Valor]
        table_data.append([
            Paragraph(f"<b>{col_name_safe}</b>", campo_style),
            Paragraph(valor_safe, valor_style)
        ])

    # Criar tabela (2 colunas: campo e valor)
    table = Table(table_data, colWidths=[5.8*cm, 20.2*cm], splitByRow=1)

    # Estilo da tabela
    table.setStyle(TableStyle([
        # Coluna de campos (esquerda)
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#667eea')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.whitesmoke),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (0, -1), 8),
        ('LEFTPADDING', (0, 0), (0, -1), 8),
        ('RIGHTPADDING', (0, 0), (0, -1), 8),
        ('TOPPADDING', (0, 0), (0, -1), 6),
        ('BOTTOMPADDING', (0, 0), (0, -1), 6),

        # Coluna de valores (direita)
        ('BACKGROUND', (1, 0), (1, -1), colors.HexColor('#f8fafc')),
        ('TEXTCOLOR', (1, 0), (1, -1), colors.HexColor('#1e293b')),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTSIZE', (1, 0), (1, -1), 8),
        ('RIGHTPADDING', (1, 0), (1, -1), 8),
        ('LEFTPADDING', (1, 0), (1, -1), 8),
        ('TOPPADDING', (1, 0), (1, -1), 5),
        ('BOTTOMPADDING', (1, 0), (1, -1), 5),
        ('WORDWRAP', (1, 0), (1, -1), 'CJK'),

        # Bordas
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    elements.append(table)

    # Rodapé
    elements.append(Spacer(1, 0.5*cm))

    rodape_style = ParagraphStyle(
        'Rodape',
        parent=styles['Normal'],
        fontSize=7,
        textColor=colors.HexColor('#94a3b8'),
        alignment=1
    )

    elements.append(Paragraph(
        f"Documento gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} | UECI",
        rodape_style
    ))

    # Construir PDF
    try:
        doc.build(elements)
        buffer.seek(0)

        # Nome do arquivo
        nota_num = dados_registro.get('Nº Nota Recomendatória', 'SemNumero')
        if nota_num:
            nota_num = str(nota_num).replace('/', '_').replace(' ', '_')

        filename = f'Preenchimento_Nota_{nota_num}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'

        return send_file(
            buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Erro ao gerar PDF: {str(e)}")
        flash(f'Erro ao gerar PDF: {str(e)}', 'danger')
        return redirect(url_for('main.index'))


@main_bp.route('/preencher/<token>/finalizar', methods=['POST'])
def finalizar_preenchimento(token):
    """Marca o preenchimento como finalizado"""
    link = LinkTemporario.query.filter_by(token=token).first()

    if not link:
        return jsonify({'success': False, 'message': 'Link inválido'}), 404

    # Marcar como usado/finalizado
    link.usado_em = datetime.utcnow()
    db.session.commit()

    return jsonify({'success': True, 'message': 'Preenchimento finalizado com sucesso!'})

@main_bp.route('/api/calendario/prazos')
@login_required
def get_calendario_prazos():
    """Retorna eventos para o calendário"""
    from datetime import date
    import unicodedata

    def remove_acentos(texto):
        nfkd = unicodedata.normalize('NFKD', texto)
        return ''.join([c for c in nfkd if not unicodedata.combining(c)])

    print("\n" + "=" * 70)
    print("CARREGANDO CALENDÁRIO")
    print("=" * 70)

    eventos = []
    hoje = date.today()

    abas = AbaConfig.query.filter_by(is_active=True).all()

    for aba in abas:
        print(f"\nABA: {aba.aba_name}")
        registros = PlanilhaData.query.filter_by(aba_name=aba.aba_name).all()

        for registro in registros:
            dados = registro.get_data()

            for campo_nome, valor in dados.items():
                if not valor:
                    continue

                campo_normalizado = remove_acentos(campo_nome.lower())

                # Verificar prazo de término OU data limite de retorno
                eh_prazo_termino = 'prazo' in campo_normalizado and 'termino' in campo_normalizado
                eh_data_limite_retorno = 'data' in campo_normalizado and 'limite' in campo_normalizado and 'retorno' in campo_normalizado

                if eh_prazo_termino or eh_data_limite_retorno:
                    tipo_prazo = "Término" if eh_prazo_termino else "Retorno da Área"
                    icone = "📅" if eh_prazo_termino else "📬"

                    try:
                        data_prevista = None
                        data_formatada = str(valor).strip()

                        if '/' in data_formatada:
                            parts = data_formatada.split('/')
                            if len(parts) == 3 and len(parts[2]) == 4:
                                data_prevista = date(int(parts[2]), int(parts[1]), int(parts[0]))
                        elif '-' in data_formatada:
                            data_limpa = data_formatada.split(' ')[0]
                            parts = data_limpa.split('-')
                            if len(parts) == 3 and len(parts[0]) == 4:
                                data_prevista = date(int(parts[0]), int(parts[1]), int(parts[2]))
                                data_formatada = f"{parts[2]}/{parts[1]}/{parts[0]}"

                        if data_prevista:
                            diferenca = (data_prevista - hoje).days

                            nota_recomendatoria = dados.get('Nº Nota Recomendatória', '')
                            identificacao = f"Nota {nota_recomendatoria}" if nota_recomendatoria else f'Reg #{registro.id}'

                            recomendacao = dados.get('Recomendação', '')[:100] or dados.get('Constatação', '')[:100]
                            if recomendacao:
                                if len(recomendacao) > 100:
                                    recomendacao = recomendacao[:100] + '...'

                            if diferenca < 0:
                                classe = 'fc-event-vencido'
                                status = f'Vencido há {abs(diferenca)} dia(s)'
                                badge = 'danger'
                            elif diferenca == 0:
                                classe = 'fc-event-hoje'
                                status = 'Vence Hoje'
                                badge = 'warning'
                            elif diferenca <= 7:
                                classe = 'fc-event-proximo'
                                status = f'Em {diferenca} dia(s)'
                                badge = 'info'
                            else:
                                classe = 'fc-event-futuro'
                                status = f'Em {diferenca} dias'
                                badge = 'success'

                            evento = {
                                'title': f'{icone} {tipo_prazo}: {identificacao}',
                                'start': data_prevista.isoformat(),
                                'className': classe,
                                'extendedProps': {
                                    'aba_name': aba.aba_name,
                                    'registro_id': registro.id,
                                    'descricao': recomendacao or identificacao,
                                    'data_formatada': data_formatada,
                                    'status_texto': status,
                                    'badge_color': badge,
                                    'dias': diferenca,
                                    'tipo_prazo': tipo_prazo
                                }
                            }
                            eventos.append(evento)
                            print(f"  ✓ Evento: {data_formatada} - {tipo_prazo} - {status}")
                    except:
                        pass

    print(f"\nTOTAL EVENTOS: {len(eventos)}")

    return jsonify({
        'success': True,
        'eventos': eventos
    })


@main_bp.route('/planilha/<aba_name>')
@login_required
def view_planilha(aba_name):
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()
    registros = PlanilhaData.query.filter_by(aba_name=aba_name).order_by(PlanilhaData.row_order).all()
    dropdown_rows = DropdownConfig.query.filter_by(
        aba_name=aba_name,
        is_active=True
    ).order_by(DropdownConfig.ordem).all()

    # Converter JSON para lista de dicionários
    data = []
    for reg in registros:
        row = reg.get_data()
        row['id'] = reg.id

        # Converter datas para formato brasileiro (dd/mm/aaaa)
        columns = aba.get_columns()
        for col in columns:
            col_name = col.get('name', col) if isinstance(col, dict) else col
            col_type = col.get('type', 'text') if isinstance(col, dict) else 'text'

            if col_type == 'date' and col_name in row and row[col_name]:
                try:
                    valor = str(row[col_name])
                    if ' ' in valor:
                        valor = valor.split(' ')[0]

                    if '-' in valor:
                        date_parts = valor.split('-')
                        if len(date_parts) == 3:
                            row[col_name] = f"{date_parts[2]}/{date_parts[1]}/{date_parts[0]}"
                except:
                    pass

        data.append(row)

    def normalize_field_name(name):
        text = str(name or '').strip().lower()
        text = ''.join(
            c for c in unicodedata.normalize('NFKD', text)
            if not unicodedata.combining(c)
        )
        return ' '.join(text.split())

    columns = aba.get_columns()
    orange_columns = aba.get_orange_columns()

    dropdown_index = {}
    for dropdown in dropdown_rows:
        dropdown_index[normalize_field_name(dropdown.campo_nome)] = dropdown.get_opcoes()

    dropdown_configs = {}
    for col in columns:
        col_name = col.get('name', col) if isinstance(col, dict) else col
        matched_options = dropdown_index.get(normalize_field_name(col_name))
        if matched_options:
            dropdown_configs[col_name] = matched_options

    return render_template('planilha.html',
                         aba=aba,
                         data=data,
                         columns=columns,
                         orange_columns=orange_columns,
                         dropdown_configs=dropdown_configs,
                         aba_name=aba_name)


@main_bp.route('/planilha/<aba_name>/add', methods=['POST'])
@login_required
def add_row(aba_name):
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()

    form_data = {}
    columns = aba.get_columns()

    for col in columns:
        col_name = col.get('name', col) if isinstance(col, dict) else col
        col_type = col.get('type', 'text') if isinstance(col, dict) else 'text'
        valor = request.form.get(col_name, '')

        if col_type == 'date' and valor:
            try:
                if '/' in valor:
                    date_parts = valor.split('/')
                    if len(date_parts) == 3:
                        valor = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
                elif ' ' in valor:
                    valor = valor.split(' ')[0]
            except:
                pass

        form_data[col_name] = valor

    max_order = db.session.query(db.func.max(PlanilhaData.row_order)).filter_by(aba_name=aba_name).scalar() or 0

    novo_registro = PlanilhaData(
        aba_name=aba_name,
        row_order=max_order + 1,
        created_by=current_user.id,
        updated_by=current_user.id
    )
    novo_registro.set_data(form_data)

    db.session.add(novo_registro)
    db.session.commit()

    flash('Registro adicionado com sucesso!', 'success')
    return redirect(url_for('main.view_planilha', aba_name=aba_name))


@main_bp.route('/api/alertas/prazos')
@login_required
def get_alertas_prazos():
    """Retorna alertas de prazos próximos ou vencidos"""
    from datetime import date
    import unicodedata

    def remove_acentos(texto):
        nfkd = unicodedata.normalize('NFKD', texto)
        return ''.join([c for c in nfkd if not unicodedata.combining(c)])

    print("\n" + "=" * 70)
    print("VERIFICANDO ALERTAS DE PRAZO")
    print("=" * 70)

    alertas = []
    hoje = date.today()
    print(f"Data de hoje: {hoje}")

    abas = AbaConfig.query.filter_by(is_active=True).all()
    print(f"\nAbas ativas: {len(abas)}")

    for aba in abas:
        print(f"\n{'=' * 70}")
        print(f"ABA: {aba.aba_name}")
        print("=" * 70)

        registros = PlanilhaData.query.filter_by(aba_name=aba.aba_name).all()
        print(f"Registros: {len(registros)}")

        for idx, registro in enumerate(registros, 1):
            dados = registro.get_data()

            for campo_nome, valor in dados.items():
                if not valor:
                    continue

                campo_normalizado = remove_acentos(campo_nome.lower())

                eh_prazo_termino = 'prazo' in campo_normalizado and 'termino' in campo_normalizado
                eh_data_limite_retorno = 'data' in campo_normalizado and 'limite' in campo_normalizado and 'retorno' in campo_normalizado

                if eh_prazo_termino or eh_data_limite_retorno:
                    tipo_prazo = "Prazo de Término" if eh_prazo_termino else "Data Limite de Retorno da Área"

                    print(f"\nReg #{idx} (ID:{registro.id}) - Campo: '{campo_nome}'")
                    print(f"  Tipo: {tipo_prazo}")
                    print(f"  Valor: '{valor}'")

                    try:
                        data_prevista = None
                        data_formatada = str(valor).strip()

                        if '/' in data_formatada:
                            parts = data_formatada.split('/')
                            if len(parts) == 3 and len(parts[2]) == 4:
                                data_prevista = date(int(parts[2]), int(parts[1]), int(parts[0]))
                        elif '-' in data_formatada:
                            data_limpa = data_formatada.split(' ')[0]
                            parts = data_limpa.split('-')
                            if len(parts) == 3 and len(parts[0]) == 4:
                                data_prevista = date(int(parts[0]), int(parts[1]), int(parts[2]))
                                data_formatada = f"{parts[2]}/{parts[1]}/{parts[0]}"

                        if data_prevista:
                            diferenca = (data_prevista - hoje).days
                            print(f"  Data: {data_prevista}, Dif: {diferenca} dias")

                            if diferenca <= 7:
                                nota_recomendatoria = dados.get('Nº Nota Recomendatória', '')
                                identificacao = f"Nota {nota_recomendatoria}" if nota_recomendatoria else f'Registro #{registro.id}'

                                recomendacao = dados.get('Recomendação', '')[:100] or dados.get('Constatação', '')[:100]
                                if recomendacao:
                                    identificacao += f" - {recomendacao}"
                                    if len(recomendacao) > 100:
                                        identificacao += '...'

                                if diferenca < 0:
                                    tipo = 'vencido'
                                    titulo = f'⚠️ {tipo_prazo} Vencido há {abs(diferenca)} dia(s)!'
                                elif diferenca == 0:
                                    tipo = 'hoje'
                                    titulo = f'🔔 {tipo_prazo} Vence Hoje!'
                                else:
                                    tipo = 'proximo'
                                    titulo = f'📅 {tipo_prazo} em {diferenca} dia(s)'

                                alerta = {
                                    'tipo': tipo,
                                    'titulo': titulo,
                                    'mensagem': identificacao,
                                    'data_prevista': data_formatada,
                                    'aba_name': aba.aba_name,
                                    'registro_id': registro.id,
                                    'dias': diferenca,
                                    'tipo_prazo': tipo_prazo
                                }
                                alertas.append(alerta)
                                print(f"  ✅ ALERTA: {titulo}")

                    except Exception as e:
                        print(f"  ❌ Erro: {e}")

    alertas.sort(key=lambda x: x['dias'])

    print(f"\n{'=' * 70}")
    print(f"TOTAL DE ALERTAS: {len(alertas)}")
    print("=" * 70)

    return jsonify({
        'success': True,
        'alertas': alertas[:15]
    })


@main_bp.route('/planilha/<aba_name>/edit/<int:row_id>', methods=['POST'])
@login_required
def edit_row(aba_name, row_id):
    registro = PlanilhaData.query.filter_by(id=row_id, aba_name=aba_name).first_or_404()
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()

    form_data = {}
    columns = aba.get_columns()

    for col in columns:
        col_name = col.get('name', col) if isinstance(col, dict) else col
        col_type = col.get('type', 'text') if isinstance(col, dict) else 'text'
        valor = request.form.get(col_name, '')

        if col_type == 'date' and valor:
            try:
                if '/' in valor:
                    date_parts = valor.split('/')
                    if len(date_parts) == 3:
                        valor = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
                elif ' ' in valor:
                    valor = valor.split(' ')[0]
            except:
                pass

        form_data[col_name] = valor

    registro.set_data(form_data)
    registro.updated_by = current_user.id
    registro.updated_at = datetime.utcnow()

    db.session.commit()

    flash('Registro atualizado com sucesso!', 'success')
    return redirect(url_for('main.view_planilha', aba_name=aba_name))


@main_bp.route('/planilha/<aba_name>/delete/<int:row_id>', methods=['POST'])
@login_required
def delete_row(aba_name, row_id):
    registro = PlanilhaData.query.filter_by(id=row_id, aba_name=aba_name).first_or_404()
    db.session.delete(registro)
    db.session.commit()

    flash('Registro excluído com sucesso!', 'success')
    return redirect(url_for('main.view_planilha', aba_name=aba_name))


@main_bp.route('/planilha/<aba_name>/export/excel')
@login_required
def export_planilha_excel(aba_name):
    """Exporta a planilha para Excel com formatação profissional"""
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()
    registros = PlanilhaData.query.filter_by(aba_name=aba_name).order_by(PlanilhaData.row_order).all()
    columns = aba.get_columns()

    data = []
    for reg in registros:
        row_data = reg.get_data()

        for col in columns:
            col_name = col.get('name', col) if isinstance(col, dict) else col
            col_type = col.get('type', 'text') if isinstance(col, dict) else 'text'

            if col_type == 'date' and col_name in row_data and row_data[col_name]:
                try:
                    valor = str(row_data[col_name])
                    if ' ' in valor:
                        valor = valor.split(' ')[0]

                    if '-' in valor:
                        date_parts = valor.split('-')
                        if len(date_parts) == 3:
                            row_data[col_name] = f"{date_parts[2]}/{date_parts[1]}/{date_parts[0]}"
                except:
                    pass

        data.append(row_data)

    df = pd.DataFrame(data)
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=aba_name[:31], index=False)

        workbook = writer.book
        worksheet = writer.sheets[aba_name[:31]]

        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    cell_value = str(cell.value) if cell.value is not None else ''
                    if len(cell_value) > max_length:
                        max_length = len(cell_value)

                    if cell.row > 1:
                        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                except:
                    pass

            adjusted_width = min(max(max_length + 2, 10), 60)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        worksheet.freeze_panes = 'A2'
        worksheet.row_dimensions[1].height = 30

    output.seek(0)
    filename = f'{aba_name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


# ✅ FUNÇÃO CORRIGIDA - Gerar Link (usando modelo LinkTemporario)
@main_bp.route('/planilha/<aba_name>/gerar-link', methods=['POST'])
@login_required
def gerar_link_area(aba_name):
    """Gera link temporário para área preencher apenas colunas laranjas"""
    data = request.get_json()
    registro_id = data.get('registro_id')

    if not registro_id:
        return jsonify({'success': False, 'message': 'ID do registro não fornecido'}), 400

    # Verificar se o registro existe
    registro = PlanilhaData.query.filter_by(id=registro_id, aba_name=aba_name).first()
    if not registro:
        return jsonify({'success': False, 'message': 'Registro não encontrado'}), 404

    # Gerar token único
    token = secrets.token_urlsafe(32)
    expiracao = datetime.utcnow() + timedelta(days=10)

    # Salvar no banco usando o modelo LinkTemporario
    link_temp = LinkTemporario(
        token=token,
        aba_name=aba_name,
        registro_id=registro_id,
        expiracao=expiracao,
        criado_por=current_user.id
    )
    db.session.add(link_temp)
    db.session.commit()

    # Gerar link público
    link = url_for('main.preenchimento_externo', token=token, _external=True)
    expiracao_formatada = expiracao.strftime('%d/%m/%Y às %H:%M')

    return jsonify({
        'success': True,
        'link': link,
        'expiracao': expiracao_formatada
    })


# ✅ MANTER ROTAS DE ÁREA (sistema antigo de hash)
@main_bp.route('/area/<aba_name>/<int:registro_id>/<token>')
def preencher_area(aba_name, registro_id, token):
    """Página pública para área preencher apenas colunas laranjas (sistema legado)"""
    registro = PlanilhaData.query.filter_by(id=registro_id, aba_name=aba_name).first_or_404()

    token_hash = hashlib.sha256(token.encode()).hexdigest()
    dados_registro = registro.get_data()

    token_valido = False
    if '_temp_links' in dados_registro:
        for link_info in dados_registro['_temp_links']:
            if link_info['token_hash'] == token_hash:
                expiracao = datetime.fromisoformat(link_info['expiracao'])
                if datetime.utcnow() < expiracao:
                    token_valido = True
                    break

    if not token_valido:
        flash('Link inválido ou expirado. Entre em contato com a UECI para obter um novo link.', 'danger')
        return render_template('error.html', message='Link de acesso inválido ou expirado')

    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()
    columns = aba.get_columns()
    orange_columns = aba.get_orange_columns()
    data = registro.get_data()

    for col in columns:
        col_name = col.get('name', col) if isinstance(col, dict) else col
        col_type = col.get('type', 'text') if isinstance(col, dict) else 'text'

        if col_type == 'date' and col_name in data and data[col_name]:
            try:
                valor = str(data[col_name])
                if ' ' in valor:
                    valor = valor.split(' ')[0]
                if '-' in valor:
                    date_parts = valor.split('-')
                    if len(date_parts) == 3:
                        data[col_name] = f"{date_parts[2]}/{date_parts[1]}/{date_parts[0]}"
            except:
                pass

    return render_template('area_preencher.html',
                         aba_name=aba_name,
                         registro_id=registro_id,
                         token=token,
                         columns=columns,
                         orange_columns=orange_columns,
                         data=data)


@main_bp.route('/area/<aba_name>/<int:registro_id>/<token>/salvar', methods=['POST'])
def salvar_area(aba_name, registro_id, token):
    """Salva preenchimento da área (apenas colunas laranjas)"""
    registro = PlanilhaData.query.filter_by(id=registro_id, aba_name=aba_name).first_or_404()

    token_hash = hashlib.sha256(token.encode()).hexdigest()
    dados_registro = registro.get_data()

    token_valido = False
    if '_temp_links' in dados_registro:
        for link_info in dados_registro['_temp_links']:
            if link_info['token_hash'] == token_hash:
                expiracao = datetime.fromisoformat(link_info['expiracao'])
                if datetime.utcnow() < expiracao:
                    token_valido = True
                    break

    if not token_valido:
        flash('Link inválido ou expirado.', 'danger')
        return redirect(url_for('main.preencher_area',
                               aba_name=aba_name,
                               registro_id=registro_id,
                               token=token))

    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()
    orange_columns = aba.get_orange_columns()
    columns = aba.get_columns()
    data_atual = registro.get_data()

    for col in columns:
        col_name = col.get('name', col) if isinstance(col, dict) else col
        col_type = col.get('type', 'text') if isinstance(col, dict) else 'text'

        if col_name in orange_columns:
            valor = request.form.get(col_name, '')

            if col_type == 'date' and valor:
                try:
                    if '/' in valor:
                        date_parts = valor.split('/')
                        if len(date_parts) == 3:
                            valor = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
                except:
                    pass

            data_atual[col_name] = valor

    registro.set_data(data_atual)
    registro.updated_at = datetime.utcnow()

    try:
        db.session.commit()
        flash('✅ Dados salvos com sucesso! Obrigado pelo preenchimento.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'❌ Erro ao salvar: {str(e)}', 'danger')

    return redirect(url_for('main.preencher_area',
                           aba_name=aba_name,
                           registro_id=registro_id,
                           token=token))


@main_bp.route('/admin/users')
@login_required
def admin_users():
    if not current_user.is_admin:
        flash('Acesso negado.', 'danger')
        return redirect(url_for('main.dashboard'))

    users = User.query.all()
    return render_template('admin_users.html', users=users)
