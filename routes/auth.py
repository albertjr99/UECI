from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify
from flask_login import login_user, logout_user, login_required, current_user
from models import db
from models.user import User, Token
from datetime import datetime, timedelta
import secrets

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        from models.planilha import AbaConfig
        primeira_aba = AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).first()
        if primeira_aba:
            return redirect(url_for('main.view_planilha', aba_name=primeira_aba.aba_name))
        return redirect(url_for('main.index'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        user = User.query.filter_by(username=username).first()
        
        if user:
            if not user.password_hash:
                flash('Sua senha ainda não foi definida. Entre em contato com o administrador para receber um link de definição de senha.', 'warning')
                return render_template('login.html')
            
            if not user.is_active:
                flash('Sua conta está desativada. Entre em contato com o administrador.', 'danger')
                return render_template('login.html')
            
            if user.check_password(password):
                login_user(user, remember=True)
                user.last_login = datetime.utcnow()
                db.session.commit()
                
                next_page = request.args.get('next')
                if next_page:
                    return redirect(next_page)
                
                from models.planilha import AbaConfig
                primeira_aba = AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).first()
                if primeira_aba:
                    return redirect(url_for('main.view_planilha', aba_name=primeira_aba.aba_name))
                return redirect(url_for('main.index'))
            else:
                flash('Senha incorreta.', 'danger')
        else:
            flash('Usuário não encontrado.', 'danger')
    
    return render_template('login.html')


@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Você saiu do sistema.', 'info')
    return redirect(url_for('auth.login'))


@auth_bp.route('/admin/generate-token', methods=['POST'])
@login_required
def generate_token():
    if not current_user.is_admin:
        return jsonify({'error': 'Acesso negado'}), 403
    
    data = request.get_json()
    user_id = data.get('user_id')
    token_type = data.get('token_type', 'password_reset')
    send_method = data.get('send_method', 'email')  # 'email' or 'whatsapp'
    
    user = User.query.get(user_id)
    if not user:
        return jsonify({'error': 'Usuário não encontrado'}), 404
    
    # Criar token
    token = Token()
    token.user_id = user_id
    token.token = Token.generate_token()
    token.token_type = token_type
    token.expires_at = datetime.utcnow() + timedelta(hours=24)
    
    db.session.add(token)
    db.session.commit()
    
    # Gerar link
    reset_link = url_for('auth.reset_password', token=token.token, _external=True)
    
    # Aqui você pode integrar com API de email ou WhatsApp
    # Por enquanto, apenas retornamos o link
    
    return jsonify({
        'success': True,
        'token': token.token,
        'link': reset_link,
        'expires_at': token.expires_at.isoformat(),
        'send_method': send_method,
        'user_email': user.email,
        'user_phone': user.phone
    })


@auth_bp.route('/reset-password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    token_obj = Token.query.filter_by(token=token, token_type='password_reset').first()
    
    if not token_obj or not token_obj.is_valid():
        flash('Token inválido ou expirado.', 'danger')
        return redirect(url_for('auth.login'))
    
    if request.method == 'POST':
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')
        
        if password != confirm_password:
            flash('As senhas não coincidem.', 'danger')
            return render_template('reset_password.html', token=token)
        
        if len(password) < 6:
            flash('A senha deve ter no mínimo 6 caracteres.', 'danger')
            return render_template('reset_password.html', token=token)
        
        user = token_obj.user
        user.set_password(password)
        token_obj.used = True
        
        db.session.commit()
        
        flash('Senha definida com sucesso! Faça login.', 'success')
        return redirect(url_for('auth.login'))
    
    return render_template('reset_password.html', token=token, user=token_obj.user)
