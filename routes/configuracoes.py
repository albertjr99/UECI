from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify
from flask_login import login_required, current_user
from models import db
from models.configuracoes import ConfiguracaoSistema, DropdownConfig
from models.planilha import AbaConfig
from functools import wraps

config_bp = Blueprint('configuracoes', __name__, url_prefix='/configuracoes')

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin:
            flash('Acesso negado. Apenas administradores podem acessar esta área.', 'danger')
            return redirect(url_for('main.index'))
        return f(*args, **kwargs)
    return decorated_function


@config_bp.route('/zoom')
@login_required
def zoom_config():
    """Página de configuração de zoom"""
    config = ConfiguracaoSistema.query.filter_by(chave='zoom_padrao').first()
    zoom_atual = 80  # Padrão

    if config:
        valor = config.get_valor()
        zoom_atual = valor.get('zoom', 80)

    return render_template('configuracoes/zoom.html', zoom_atual=zoom_atual)


@config_bp.route('/zoom/salvar', methods=['POST'])
@login_required
def salvar_zoom():
    """Salva a configuração de zoom"""
    zoom = request.form.get('zoom', 80, type=int)

    if zoom < 50 or zoom > 150:
        return jsonify({'success': False, 'message': 'Zoom deve estar entre 50% e 150%'}), 400

    config = ConfiguracaoSistema.query.filter_by(chave='zoom_padrao').first()

    if not config:
        config = ConfiguracaoSistema(chave='zoom_padrao')
        db.session.add(config)

    config.set_valor({'zoom': zoom})
    config.updated_by = current_user.id

    try:
        db.session.commit()
        return jsonify({'success': True, 'message': 'Zoom salvo com sucesso!', 'zoom': zoom})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao salvar: {str(e)}'}), 500


@config_bp.route('/dropdowns')
@login_required
def dropdowns_config():
    """Página de configuração de dropdowns"""
    abas = AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).all()

    # Buscar dropdowns configurados
    dropdowns = DropdownConfig.query.order_by(
        DropdownConfig.aba_name,
        DropdownConfig.ordem
    ).all()

    # Organizar por aba
    dropdowns_por_aba = {}
    for dropdown in dropdowns:
        if dropdown.aba_name not in dropdowns_por_aba:
            dropdowns_por_aba[dropdown.aba_name] = []
        dropdowns_por_aba[dropdown.aba_name].append(dropdown)

    return render_template('configuracoes/dropdowns.html',
                         abas=abas,
                         dropdowns_por_aba=dropdowns_por_aba)


@config_bp.route('/dropdowns/campos/<aba_name>')
@login_required
def get_campos_aba(aba_name):
    """Retorna os campos de uma aba"""
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first()

    if not aba:
        return jsonify({'success': False, 'message': 'Aba não encontrada'}), 404

    columns_data = aba.get_columns()

    # Extrair apenas os nomes das colunas
    campos = []
    if isinstance(columns_data, list):
        for col in columns_data:
            if isinstance(col, dict):
                # Se for dicionário, pegar a chave 'name' ou 'nome'
                campo_nome = col.get('name') or col.get('nome') or col.get('campo') or str(col)
                campos.append(campo_nome)
            else:
                # Se for string diretamente
                campos.append(str(col))

    return jsonify({'success': True, 'campos': campos})


@config_bp.route('/dropdowns/criar', methods=['POST'])
@login_required
def criar_dropdown():
    """Cria um novo dropdown"""
    data = request.get_json()

    aba_name = data.get('aba_name')
    campo_nome = data.get('campo_nome')
    opcoes = data.get('opcoes', [])

    if not aba_name or not campo_nome:
        return jsonify({'success': False, 'message': 'Aba e campo são obrigatórios'}), 400

    # Verificar se já existe
    dropdown_existente = DropdownConfig.query.filter_by(
        aba_name=aba_name,
        campo_nome=campo_nome
    ).first()

    if dropdown_existente:
        return jsonify({'success': False, 'message': 'Já existe um dropdown para este campo'}), 400

    # Criar novo dropdown
    dropdown = DropdownConfig(
        aba_name=aba_name,
        campo_nome=campo_nome,
        is_active=True,
        created_by=current_user.id,
        updated_by=current_user.id
    )
    dropdown.set_opcoes(opcoes)

    try:
        db.session.add(dropdown)
        db.session.commit()
        return jsonify({
            'success': True,
            'message': 'Dropdown criado com sucesso!',
            'dropdown': {
                'id': dropdown.id,
                'campo_nome': dropdown.campo_nome,
                'opcoes': dropdown.get_opcoes(),
                'is_active': dropdown.is_active
            }
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao criar: {str(e)}'}), 500


@config_bp.route('/dropdowns/editar/<int:dropdown_id>', methods=['POST'])
@login_required
def editar_dropdown(dropdown_id):
    """Edita um dropdown existente"""
    dropdown = DropdownConfig.query.get_or_404(dropdown_id)
    data = request.get_json()

    opcoes = data.get('opcoes', [])

    dropdown.set_opcoes(opcoes)
    dropdown.updated_by = current_user.id

    try:
        db.session.commit()
        return jsonify({
            'success': True,
            'message': 'Dropdown atualizado com sucesso!',
            'dropdown': {
                'id': dropdown.id,
                'campo_nome': dropdown.campo_nome,
                'opcoes': dropdown.get_opcoes(),
                'is_active': dropdown.is_active
            }
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao atualizar: {str(e)}'}), 500


@config_bp.route('/dropdowns/toggle/<int:dropdown_id>', methods=['POST'])
@login_required
def toggle_dropdown(dropdown_id):
    """Ativa/desativa um dropdown"""
    dropdown = DropdownConfig.query.get_or_404(dropdown_id)

    dropdown.is_active = not dropdown.is_active
    dropdown.updated_by = current_user.id

    try:
        db.session.commit()
        status = 'ativado' if dropdown.is_active else 'desativado'
        return jsonify({
            'success': True,
            'message': f'Dropdown {status} com sucesso!',
            'is_active': dropdown.is_active
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao alterar status: {str(e)}'}), 500


@config_bp.route('/dropdowns/excluir/<int:dropdown_id>', methods=['POST'])
@login_required
def excluir_dropdown(dropdown_id):
    """Exclui um dropdown"""
    dropdown = DropdownConfig.query.get_or_404(dropdown_id)

    try:
        db.session.delete(dropdown)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Dropdown excluído com sucesso!'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao excluir: {str(e)}'}), 500


@config_bp.route('/dropdowns/obter/<aba_name>')
@login_required
def obter_dropdowns_aba(aba_name):
    """Retorna todos os dropdowns ativos de uma aba"""
    dropdowns = DropdownConfig.query.filter_by(
        aba_name=aba_name,
        is_active=True
    ).order_by(DropdownConfig.ordem).all()

    resultado = {}
    for dropdown in dropdowns:
        resultado[dropdown.campo_nome] = dropdown.get_opcoes()

    return jsonify({'success': True, 'dropdowns': resultado})

    # ADICIONE ESTAS NOVAS ROTAS NO FINAL DO ARQUIVO routes/configuracoes.py

@config_bp.route('/colunas')
@login_required
def colunas_config():
    """Página de configuração de colunas"""
    abas = AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).all()
    return render_template('configuracoes/colunas.html', abas=abas)


@config_bp.route('/colunas/get/<aba_name>')
@login_required
def get_colunas(aba_name):
    """Retorna as colunas de uma aba"""
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()
    colunas = aba.get_columns()
    return jsonify({'success': True, 'colunas': colunas})


@config_bp.route('/colunas/salvar/<aba_name>', methods=['POST'])
@login_required
def salvar_colunas(aba_name):
    """Salva as colunas de uma aba"""
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()
    data = request.get_json()
    colunas = data.get('colunas', [])

    if not colunas:
        return jsonify({'success': False, 'message': 'Nenhuma coluna fornecida'}), 400

    try:
        aba.set_columns(colunas)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Colunas atualizadas com sucesso!'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao salvar: {str(e)}'}), 500

@config_bp.route('/colunas/editar/<aba_name>/<int:index>', methods=['POST'])
@login_required
def editar_coluna(aba_name, index):
    """Edita uma coluna específica"""
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()
    data = request.get_json()

    nome = data.get('nome', '').strip()
    tipo = data.get('tipo', 'text')

    if not nome:
        return jsonify({'success': False, 'message': 'Nome da coluna é obrigatório'}), 400

    colunas = aba.get_columns()

    if index < 0 or index >= len(colunas):
        return jsonify({'success': False, 'message': 'Índice inválido'}), 400

    # Atualizar coluna
    nome_antigo = colunas[index].get('name') if isinstance(colunas[index], dict) else colunas[index]
    colunas[index] = {'name': nome, 'type': tipo}

    # Atualizar lista de colunas laranjas se necessário
    orange_cols = aba.get_orange_columns()
    if nome_antigo in orange_cols:
        orange_cols = [nome if col == nome_antigo else col for col in orange_cols]
        aba.set_orange_columns(orange_cols)

    try:
        aba.set_columns(colunas)
        db.session.commit()
        return jsonify({'success': True, 'message': 'Coluna atualizada com sucesso!'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao atualizar: {str(e)}'}), 500


@config_bp.route('/colunas/orange/<aba_name>', methods=['GET', 'POST'])
@login_required
def manage_orange_columns(aba_name):
    """Gerencia quais colunas são laranjas"""
    aba = AbaConfig.query.filter_by(aba_name=aba_name).first_or_404()

    if request.method == 'POST':
        data = request.get_json()
        orange_names = data.get('orange_columns', [])

        try:
            aba.set_orange_columns(orange_names)
            db.session.commit()
            return jsonify({'success': True, 'message': 'Colunas laranjas atualizadas!'})
        except Exception as e:
            db.session.rollback()
            return jsonify({'success': False, 'message': str(e)}), 500

    return jsonify({'success': True, 'orange_columns': aba.get_orange_columns()})
