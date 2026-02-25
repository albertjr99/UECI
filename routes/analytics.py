from flask import Blueprint, render_template, request, jsonify, send_file
from flask_login import login_required, current_user
from models import db
from models.planilha import PlanilhaData, AbaConfig
from datetime import datetime, timedelta
import pandas as pd
import json
from io import BytesIO
import matplotlib
matplotlib.use('Agg')  # Backend sem interface gráfica
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
import tempfile
import os

analytics_bp = Blueprint('analytics', __name__)

@analytics_bp.route('/analytics')
@login_required
def analytics_dashboard():
    abas = AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).all()
    return render_template('analytics.html', abas=abas)


@analytics_bp.route('/analytics/data/<aba_name>')
@login_required
def get_analytics_data(aba_name):
    """Gera dados analíticos gerenciais"""

    registros = PlanilhaData.query.filter_by(aba_name=aba_name).all()

    # Converter para DataFrame
    data = []
    for reg in registros:
        row_data = reg.get_data()
        row_data['created_at'] = reg.created_at
        row_data['id'] = reg.id
        data.append(row_data)

    df = pd.DataFrame(data)

    if df.empty:
        return jsonify({
            'total_registros': 0,
            'charts': {},
            'kpis': {}
        })

    total_registros = len(df)
    charts = {}
    kpis = {}

    # KPIs Principais
    kpis['total_registros'] = total_registros

    # Análise de Status da Recomendação
    status_col = None
    for col in df.columns:
        if 'STATUS' in col.upper() and 'RECOMENDA' in col.upper():
            status_col = col
            break

    if status_col and not df[status_col].isna().all():
        status_counts = df[status_col].value_counts()

        # Mapear status para categorias
        cumpridas = 0
        nao_cumpridas = 0
        em_andamento = 0
        a_cumprir = 0

        for status, count in status_counts.items():
            status_str = str(status).lower()
            if 'cumprida' in status_str and 'não' not in status_str:
                cumpridas += count
            elif 'não cumprida' in status_str:
                nao_cumpridas += count
            elif 'andamento' in status_str:
                em_andamento += count
            elif 'cumprir' in status_str:
                a_cumprir += count

        kpis['cumpridas'] = cumpridas
        kpis['nao_cumpridas'] = nao_cumpridas
        kpis['em_andamento'] = em_andamento
        kpis['a_cumprir'] = a_cumprir
        kpis['taxa_cumprimento'] = round((cumpridas / total_registros * 100), 1) if total_registros > 0 else 0

        # Dados para gráfico de status consolidado
        charts['status_consolidado'] = {
            'labels': ['Cumpridas', 'Não Cumpridas', 'Em Andamento', 'A Cumprir'],
            'data': [cumpridas, nao_cumpridas, em_andamento, a_cumprir],
            'colors': ['#10b981', '#ef4444', '#f59e0b', '#3b82f6']
        }

        # Dados para gráfico detalhado de status
        charts['status_detalhado'] = {
            'labels': [str(s)[:50] for s in status_counts.index.tolist()],
            'data': status_counts.values.tolist()
        }

    # Análise por Setor Responsável
    setor_col = None
    for col in df.columns:
        if 'SETOR' in col.upper() and 'RESPONSÁVEL' in col.upper():
            setor_col = col
            break

    if setor_col and not df[setor_col].isna().all():
        setor_counts = df[setor_col].value_counts().head(10)

        charts['por_setor'] = {
            'labels': [str(s) for s in setor_counts.index.tolist()],
            'data': setor_counts.values.tolist()
        }

    # Análise por Unidade Gestora
    ug_col = None
    for col in df.columns:
        if 'UNIDADE' in col.upper() and 'GESTORA' in col.upper() or col.upper() == 'UG':
            ug_col = col
            break

    if ug_col and not df[ug_col].isna().all():
        ug_counts = df[ug_col].value_counts()

        charts['por_ug'] = {
            'labels': [str(s) for s in ug_counts.index.tolist()],
            'data': ug_counts.values.tolist()
        }

    # Análise por Origem
    origem_col = None
    for col in df.columns:
        if col.upper() == 'ORIGEM':
            origem_col = col
            break

    if origem_col and not df[origem_col].isna().all():
        origem_counts = df[origem_col].value_counts()

        charts['por_origem'] = {
            'labels': [str(s) for s in origem_counts.index.tolist()],
            'data': origem_counts.values.tolist()
        }

    # Análise por Tipo de Ação
    tipo_col = None
    for col in df.columns:
        if 'TIPO' in col.upper() and 'AÇÃO' in col.upper():
            tipo_col = col
            break

    if tipo_col and not df[tipo_col].isna().all():
        tipo_counts = df[tipo_col].value_counts()

        charts['por_tipo'] = {
            'labels': [str(s) for s in tipo_counts.index.tolist()],
            'data': tipo_counts.values.tolist()
        }

    # Análise de Prazos
    prazo_fim_col = None
    for col in df.columns:
        if 'PRAZO' in col.upper() and ('TÉRMINO' in col.upper() or 'CONCLUSÃO' in col.upper()):
            prazo_fim_col = col
            break

    if prazo_fim_col and not df[prazo_fim_col].isna().all():
        df_temp = df.copy()
        df_temp[prazo_fim_col] = pd.to_datetime(df_temp[prazo_fim_col], errors='coerce')
        df_temp = df_temp.dropna(subset=[prazo_fim_col])

        if not df_temp.empty:
            hoje = pd.Timestamp.now()
            df_temp['dias_ate_prazo'] = (df_temp[prazo_fim_col] - hoje).dt.days

            atrasadas = len(df_temp[df_temp['dias_ate_prazo'] < 0])
            no_prazo = len(df_temp[df_temp['dias_ate_prazo'] >= 0])

            kpis['atrasadas'] = atrasadas
            kpis['no_prazo'] = no_prazo

            charts['situacao_prazos'] = {
                'labels': ['No Prazo', 'Atrasadas'],
                'data': [no_prazo, atrasadas],
                'colors': ['#10b981', '#ef4444']
            }

    return jsonify({
        'total_registros': total_registros,
        'charts': charts,
        'kpis': kpis
    })


def criar_grafico_pizza(labels, data, colors, titulo):
    """Cria um gráfico de pizza profissional"""
    fig, ax = plt.subplots(figsize=(8, 6), facecolor='white')

    # Calcular percentuais
    total = sum(data)
    percentages = [(x/total)*100 for x in data]

    # Criar pizza
    wedges, texts, autotexts = ax.pie(
        data,
        labels=None,
        autopct='%1.1f%%',
        colors=colors,
        startangle=90,
        textprops={'fontsize': 11, 'weight': 'bold', 'color': 'white'}
    )

    # Legenda
    legend_labels = [f'{label}: {val} ({perc:.1f}%)' for label, val, perc in zip(labels, data, percentages)]
    ax.legend(legend_labels, loc='center left', bbox_to_anchor=(1, 0, 0.5, 1), fontsize=10)

    ax.set_title(titulo, fontsize=14, weight='bold', pad=20)

    # Salvar em buffer
    buf = BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf


def criar_grafico_barras(labels, data, titulo, cor='#667eea', horizontal=False):
    """Cria um gráfico de barras profissional"""
    fig, ax = plt.subplots(figsize=(10, 6), facecolor='white')

    # Calcular percentuais
    total = sum(data)
    percentages = [(x/total)*100 for x in data]

    if horizontal:
        bars = ax.barh(labels, data, color=cor, edgecolor='black', linewidth=0.7)
        ax.set_xlabel('Quantidade', fontsize=11, weight='bold')
        ax.set_ylabel('')

        # Adicionar valores e percentuais nas barras
        for i, (bar, val, perc) in enumerate(zip(bars, data, percentages)):
            width = bar.get_width()
            ax.text(width + max(data)*0.01, bar.get_y() + bar.get_height()/2,
                   f'{val} ({perc:.1f}%)',
                   ha='left', va='center', fontsize=10, weight='bold')
    else:
        bars = ax.bar(labels, data, color=cor, edgecolor='black', linewidth=0.7)
        ax.set_ylabel('Quantidade', fontsize=11, weight='bold')
        ax.set_xlabel('')
        plt.xticks(rotation=45, ha='right')

        # Adicionar valores e percentuais nas barras
        for bar, val, perc in zip(bars, data, percentages):
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + max(data)*0.01,
                   f'{val}\n({perc:.1f}%)',
                   ha='center', va='bottom', fontsize=10, weight='bold')

    ax.set_title(titulo, fontsize=14, weight='bold', pad=20)
    ax.grid(axis='x' if horizontal else 'y', alpha=0.3, linestyle='--')
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)

    # Salvar em buffer
    buf = BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf


@analytics_bp.route('/analytics/export/<aba_name>')
@login_required
def export_report(aba_name):
    """Exporta relatório executivo PROFISSIONAL em PDF com gráficos reais"""

    registros = PlanilhaData.query.filter_by(aba_name=aba_name).all()

    data = []
    for reg in registros:
        row_data = reg.get_data()
        data.append(row_data)

    df = pd.DataFrame(data)

    if df.empty:
        return jsonify({'error': 'Sem dados para gerar relatório'}), 400

    # Criar PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
        leftMargin=0.5*inch,
        rightMargin=0.5*inch
    )

    elements = []
    styles = getSampleStyleSheet()

    # Estilos personalizados PROFISSIONAIS
    title_style = ParagraphStyle(
        'ExecutiveTitle',
        parent=styles['Heading1'],
        fontSize=26,
        textColor=colors.HexColor('#1a237e'),
        spaceAfter=10,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )

    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#455a64'),
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica'
    )

    heading_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontSize=16,
        textColor=colors.HexColor('#283593'),
        spaceAfter=15,
        spaceBefore=25,
        fontName='Helvetica-Bold',
        borderColor=colors.HexColor('#283593'),
        borderWidth=0,
        borderPadding=5,
        leftIndent=0
    )

    normal_style = ParagraphStyle(
        'ExecutiveNormal',
        parent=styles['Normal'],
        fontSize=11,
        alignment=TA_JUSTIFY,
        spaceAfter=12,
        fontName='Helvetica',
        leading=16
    )

    # ========== CAPA ==========
    elements.append(Spacer(1, 1.5*inch))
    elements.append(Paragraph("RELATÓRIO EXECUTIVO GERENCIAL", title_style))
    elements.append(Paragraph(f"{aba_name}", title_style))
    elements.append(Spacer(1, 0.3*inch))
    elements.append(Paragraph(f"Período de Análise: {datetime.now().strftime('%B de %Y')}", subtitle_style))
    elements.append(Paragraph(f"Data de Geração: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", subtitle_style))
    elements.append(Spacer(1, 0.5*inch))

    # Linha decorativa
    line_table = Table([['']], colWidths=[10*inch])
    line_table.setStyle(TableStyle([
        ('LINEABOVE', (0, 0), (-1, 0), 3, colors.HexColor('#667eea')),
    ]))
    elements.append(line_table)

    elements.append(Spacer(1, 1*inch))
    elements.append(Paragraph("<para alignment='center'><i>Documento Confidencial - Uso Interno</i></para>", normal_style))
    elements.append(Paragraph("<para alignment='center'><b>UECI - Unidade Executiva de Controle Interno</b></para>", normal_style))

    elements.append(PageBreak())

    # ========== ANÁLISE DOS DADOS ==========
    total_registros = len(df)

    # Buscar colunas
    status_col = None
    for col in df.columns:
        if 'STATUS' in col.upper() and 'RECOMENDA' in col.upper():
            status_col = col
            break

    cumpridas = nao_cumpridas = em_andamento = a_cumprir = 0

    if status_col:
        status_counts = df[status_col].value_counts()
        for status, count in status_counts.items():
            status_str = str(status).lower()
            if 'cumprida' in status_str and 'não' not in status_str:
                cumpridas += count
            elif 'não cumprida' in status_str:
                nao_cumpridas += count
            elif 'andamento' in status_str:
                em_andamento += count
            elif 'cumprir' in status_str:
                a_cumprir += count

    taxa_cumprimento = round((cumpridas / total_registros * 100), 1) if total_registros > 0 else 0

    # ========== SUMÁRIO EXECUTIVO ==========
    elements.append(Paragraph("SUMÁRIO EXECUTIVO", heading_style))

    summary_data = [
        ['', '', '', ''],
        ['INDICADOR', 'VALOR', 'PERCENTUAL', 'STATUS'],
        ['Total de Recomendações', str(total_registros), '100%', '●'],
        ['Recomendações Cumpridas', str(cumpridas), f'{taxa_cumprimento}%', '✓' if taxa_cumprimento >= 70 else '○'],
        ['Recomendações Não Cumpridas', str(nao_cumpridas), f'{round((nao_cumpridas/total_registros*100), 1)}%', '✗' if nao_cumpridas > total_registros*0.2 else '○'],
        ['Em Andamento', str(em_andamento), f'{round((em_andamento/total_registros*100), 1)}%', '⧗'],
        ['A Cumprir', str(a_cumprir), f'{round((a_cumprir/total_registros*100), 1)}%', '○'],
    ]

    summary_table = Table(summary_data, colWidths=[3.5*inch, 1.5*inch, 1.5*inch, 1*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#283593')),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#5c6bc0')),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.whitesmoke),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, -1), 11),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 12),
        ('TOPPADDING', (0, 1), (-1, 1), 12),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#e8f5e9')),
        ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#c8e6c9')),
        ('BACKGROUND', (0, 4), (-1, 4), colors.HexColor('#ffebee')),
        ('BACKGROUND', (0, 5), (-1, 5), colors.HexColor('#fff3e0')),
        ('BACKGROUND', (0, 6), (-1, 6), colors.HexColor('#e3f2fd')),
        ('GRID', (0, 1), (-1, -1), 1, colors.grey),
        ('LINEBELOW', (0, 1), (-1, 1), 2, colors.HexColor('#283593')),
    ]))

    elements.append(summary_table)
    elements.append(Spacer(1, 0.3*inch))

    # Análise textual
    if taxa_cumprimento >= 70:
        analise = f"<b>DESEMPENHO SATISFATÓRIO:</b> A taxa de cumprimento de {taxa_cumprimento}% indica que a maioria das recomendações foi implementada com sucesso. Este resultado demonstra comprometimento e efetividade na gestão de riscos."
        cor_destaque = 'green'
    elif taxa_cumprimento >= 50:
        analise = f"<b>DESEMPENHO MODERADO:</b> A taxa de cumprimento de {taxa_cumprimento}% está na média esperada, mas há margem significativa para melhoria. Recomenda-se intensificar o acompanhamento das recomendações pendentes."
        cor_destaque = 'orange'
    else:
        analise = f"<b>ATENÇÃO NECESSÁRIA:</b> A taxa de cumprimento de {taxa_cumprimento}% está abaixo do esperado e requer atenção imediata da alta gestão. É fundamental estabelecer plano de ação para reverter este cenário."
        cor_destaque = 'red'

    elements.append(Paragraph(f"<para><font color='{cor_destaque}'>{analise}</font></para>", normal_style))

    elements.append(PageBreak())

    # ========== GRÁFICO 1: STATUS CONSOLIDADO ==========
    if cumpridas + nao_cumpridas + em_andamento + a_cumprir > 0:
        elements.append(Paragraph("1. DISTRIBUIÇÃO POR STATUS DAS RECOMENDAÇÕES", heading_style))

        img_buf = criar_grafico_pizza(
            ['Cumpridas', 'Não Cumpridas', 'Em Andamento', 'A Cumprir'],
            [cumpridas, nao_cumpridas, em_andamento, a_cumprir],
            ['#10b981', '#ef4444', '#f59e0b', '#3b82f6'],
            'Status Consolidado das Recomendações'
        )

        img = RLImage(img_buf, width=6*inch, height=4*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))

        interpretacao = f"""
        <para>
        <b>Interpretação:</b> O gráfico apresenta a distribuição consolidada das {total_registros} recomendações por status de implementação.
        Observa-se que {cumpridas} recomendações ({taxa_cumprimento}%) foram plenamente cumpridas, enquanto {nao_cumpridas} permanecem não implementadas.
        Adicionalmente, {em_andamento} recomendações estão em fase de implementação e {a_cumprir} aguardam início das ações.
        </para>
        """
        elements.append(Paragraph(interpretacao, normal_style))

        elements.append(PageBreak())

    # ========== GRÁFICO 2: STATUS DETALHADO ==========
    if status_col:
        elements.append(Paragraph("2. ANÁLISE DETALHADA POR CATEGORIA DE STATUS", heading_style))

        status_counts = df[status_col].value_counts()
        labels = [str(s)[:40] for s in status_counts.index.tolist()]
        values = status_counts.values.tolist()

        img_buf = criar_grafico_barras(
            labels,
            values,
            'Distribuição Detalhada por Status',
            cor='#667eea',
            horizontal=True
        )

        img = RLImage(img_buf, width=7*inch, height=5*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))

        top_status = labels[0]
        top_count = values[0]
        top_perc = round((top_count/total_registros*100), 1)

        interpretacao = f"""
        <para>
        <b>Interpretação:</b> A categoria "{top_status}" concentra o maior volume de recomendações,
        totalizando {top_count} ocorrências ({top_perc}% do total). Esta distribuição permite identificar
        padrões de conformidade e direcionar esforços gerenciais de forma mais assertiva.
        </para>
        """
        elements.append(Paragraph(interpretacao, normal_style))

        elements.append(PageBreak())

    # ========== GRÁFICO 3: POR UNIDADE GESTORA ==========
    ug_col = None
    for col in df.columns:
        if 'UNIDADE' in col.upper() and 'GESTORA' in col.upper() or col.upper() == 'UG':
            ug_col = col
            break

    if ug_col and not df[ug_col].isna().all():
        elements.append(Paragraph("3. DISTRIBUIÇÃO POR UNIDADE GESTORA", heading_style))

        ug_counts = df[ug_col].value_counts()
        labels = [str(s) for s in ug_counts.index.tolist()]
        values = ug_counts.values.tolist()
        colors_ug = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'][:len(labels)]

        img_buf = criar_grafico_pizza(
            labels,
            values,
            colors_ug,
            'Distribuição por Unidade Gestora'
        )

        img = RLImage(img_buf, width=6*inch, height=4*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))

        top_ug = labels[0]
        top_ug_count = values[0]
        top_ug_perc = round((top_ug_count/total_registros*100), 1)

        interpretacao = f"""
        <para>
        <b>Interpretação:</b> A Unidade Gestora "{top_ug}" apresenta o maior volume de recomendações,
        com {top_ug_count} registros ({top_ug_perc}% do total). Esta concentração pode indicar necessidade de
        fortalecimento dos controles internos ou maior exposição a riscos operacionais nesta unidade.
        </para>
        """
        elements.append(Paragraph(interpretacao, normal_style))

        elements.append(PageBreak())

    # ========== GRÁFICO 4: POR ORIGEM ==========
    origem_col = None
    for col in df.columns:
        if col.upper() == 'ORIGEM':
            origem_col = col
            break

    if origem_col and not df[origem_col].isna().all():
        elements.append(Paragraph("4. ANÁLISE POR ORIGEM DAS RECOMENDAÇÕES", heading_style))

        origem_counts = df[origem_col].value_counts()
        labels = [str(s) for s in origem_counts.index.tolist()]
        values = origem_counts.values.tolist()

        img_buf = criar_grafico_barras(
            labels,
            values,
            'Distribuição por Fonte/Origem',
            cor='#10b981',
            horizontal=False
        )

        img = RLImage(img_buf, width=6*inch, height=4*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))

        top_origem = labels[0]
        top_origem_count = values[0]
        top_origem_perc = round((top_origem_count/total_registros*100), 1)

        interpretacao = f"""
        <para>
        <b>Interpretação:</b> A origem "{top_origem}" é responsável por {top_origem_count} recomendações
        ({top_origem_perc}% do total), sendo a principal fonte de apontamentos. Este dado subsidia
        planejamento de auditorias futuras e priorização de áreas de maior criticidade.
        </para>
        """
        elements.append(Paragraph(interpretacao, normal_style))

        elements.append(PageBreak())

    # ========== GRÁFICO 5: POR TIPO DE AÇÃO ==========
    tipo_col = None
    for col in df.columns:
        if 'TIPO' in col.upper() and 'AÇÃO' in col.upper():
            tipo_col = col
            break

    if tipo_col and not df[tipo_col].isna().all():
        elements.append(Paragraph("5. CLASSIFICAÇÃO POR TIPO DE AÇÃO", heading_style))

        tipo_counts = df[tipo_col].value_counts()
        labels = [str(s) for s in tipo_counts.index.tolist()]
        values = tipo_counts.values.tolist()
        colors_tipo = ['#8b5cf6', '#f59e0b'][:len(labels)]

        img_buf = criar_grafico_pizza(
            labels,
            values,
            colors_tipo,
            'Tipo de Ação: Melhoria vs Regularização'
        )

        img = RLImage(img_buf, width=6*inch, height=4*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))

        melhoria = 0
        regularizacao = 0
        for tipo, count in tipo_counts.items():
            if 'MELHORIA' in str(tipo).upper():
                melhoria = count
            elif 'REGULARIZAÇÃO' in str(tipo).upper():
                regularizacao = count

        if melhoria > regularizacao:
            analise_tipo = f"predominam ações de <b>MELHORIA</b> ({melhoria} registros), indicando postura proativa de aprimoramento contínuo"
        else:
            analise_tipo = f"predominam ações de <b>REGULARIZAÇÃO</b> ({regularizacao} registros), sinalizando necessidade de correção de não conformidades"

        interpretacao = f"""
        <para>
        <b>Interpretação:</b> Na classificação por tipo de ação, {analise_tipo}.
        Esta proporção reflete o nível de maturidade dos controles internos e a efetividade preventiva da gestão.
        </para>
        """
        elements.append(Paragraph(interpretacao, normal_style))

        elements.append(PageBreak())

    # ========== GRÁFICO 6: SITUAÇÃO DE PRAZOS ==========
    prazo_fim_col = None
    for col in df.columns:
        if 'PRAZO' in col.upper() and ('TÉRMINO' in col.upper() or 'CONCLUSÃO' in col.upper()):
            prazo_fim_col = col
            break

    if prazo_fim_col and not df[prazo_fim_col].isna().all():
        elements.append(Paragraph("6. ANÁLISE DE CUMPRIMENTO DE PRAZOS", heading_style))

        df_temp = df.copy()
        df_temp[prazo_fim_col] = pd.to_datetime(df_temp[prazo_fim_col], errors='coerce')
        df_temp = df_temp.dropna(subset=[prazo_fim_col])

        if not df_temp.empty:
            hoje = pd.Timestamp.now()
            df_temp['dias_ate_prazo'] = (df_temp[prazo_fim_col] - hoje).dt.days

            atrasadas = len(df_temp[df_temp['dias_ate_prazo'] < 0])
            no_prazo = len(df_temp[df_temp['dias_ate_prazo'] >= 0])

            img_buf = criar_grafico_pizza(
                ['No Prazo', 'Atrasadas'],
                [no_prazo, atrasadas],
                ['#10b981', '#ef4444'],
                'Situação dos Prazos de Implementação'
            )

            img = RLImage(img_buf, width=6*inch, height=4*inch)
            elements.append(img)
            elements.append(Spacer(1, 0.2*inch))

            perc_atrasadas = round((atrasadas/(atrasadas+no_prazo)*100), 1)

            if perc_atrasadas > 30:
                alert_tipo = "CRÍTICO"
                alert_cor = "red"
                alert_msg = "requer ação imediata da alta gestão para regularização"
            elif perc_atrasadas > 15:
                alert_tipo = "ALERTA"
                alert_cor = "orange"
                alert_msg = "demanda atenção e acompanhamento intensificado"
            else:
                alert_tipo = "CONTROLADO"
                alert_cor = "green"
                alert_msg = "situação sob controle, manter monitoramento"

            interpretacao = f"""
            <para>
            <b>Interpretação:</b> <font color='{alert_cor}'><b>STATUS {alert_tipo}:</b></font>
            Do total de {atrasadas + no_prazo} recomendações com prazos definidos,
            {atrasadas} ({perc_atrasadas}%) encontram-se com prazos vencidos, enquanto {no_prazo}
            ({round((no_prazo/(atrasadas+no_prazo)*100), 1)}%) estão dentro do prazo.
            Esta situação {alert_msg}.
            </para>
            """
            elements.append(Paragraph(interpretacao, normal_style))

    elements.append(PageBreak())

    # ========== CONCLUSÕES E RECOMENDAÇÕES ==========
    elements.append(Paragraph("CONCLUSÕES E RECOMENDAÇÕES GERENCIAIS", heading_style))

    conclusoes = []

    conclusoes.append(f"<b>1. Efetividade da Implementação:</b> Com taxa de cumprimento de {taxa_cumprimento}%, " +
                     ("o desempenho está acima da média, evidenciando comprometimento institucional." if taxa_cumprimento >= 70 else
                      "há margem significativa para melhoria na implementação das recomendações." if taxa_cumprimento >= 50 else
                      "urge estabelecimento de plano de ação com prazos definidos e responsáveis designados."))

    if nao_cumpridas > (total_registros * 0.2):
        conclusoes.append(f"<b>2. Recomendações Não Cumpridas:</b> O volume de {nao_cumpridas} recomendações não implementadas ({round((nao_cumpridas/total_registros*100), 1)}%) requer análise das justificativas e eventuais impedimentos estruturais.")

    if em_andamento > 0:
        conclusoes.append(f"<b>3. Recomendações em Andamento:</b> As {em_andamento} recomendações em fase de implementação necessitam de monitoramento próximo para garantir conclusão dentro dos prazos estabelecidos.")

    if 'top_ug' in locals():
        conclusoes.append(f"<b>4. Concentração por Unidade:</b> A concentração de {top_ug_perc}% das recomendações na unidade '{top_ug}' sugere necessidade de reforço nos controles internos ou capacitação específica da equipe.")

    conclusoes.append("<b>5. Governança e Transparência:</b> Recomenda-se institucionalizar reuniões trimestrais de acompanhamento com os gestores responsáveis, documentando progressos e dificuldades.")

    conclusoes.append("<b>6. Boas Práticas:</b> Identificar e documentar as práticas exitosas dos setores com melhores índices de cumprimento para disseminação institucional.")

    for i, conclusao in enumerate(conclusoes, 1):
        elements.append(Paragraph(conclusao, normal_style))
        elements.append(Spacer(1, 0.1*inch))

    elements.append(Spacer(1, 0.4*inch))

    # Assinatura
    elements.append(Paragraph("<para alignment='center'>___________________________________________</para>", normal_style))
    elements.append(Paragraph("<para alignment='center'><b>UECI - Unidade Executiva de Controle Interno</b></para>", normal_style))
    elements.append(Paragraph(f"<para alignment='center'>Gerado automaticamente em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</para>", normal_style))

    # Rodapé institucional
    elements.append(Spacer(1, 0.3*inch))
    footer_table = Table([['Documento de uso interno - Confidencial']], colWidths=[10*inch])
    footer_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.grey),
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.grey),
    ]))
    elements.append(footer_table)

    # Construir PDF
    doc.build(elements)

    buffer.seek(0)
    filename = f'Relatorio_Executivo_{aba_name.replace(" ", "_")}_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf'

    return send_file(
        buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename
    )