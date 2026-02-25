import sys
import os

# Tentar usar venv_novo se .venv não estiver funcionando
venv_paths = [
    r'C:\Users\Computador\OneDrive\UECI\venv_novo\Lib\site-packages',
    r'C:\Users\Computador\OneDrive\UECI\.venv\Lib\site-packages'
]

for venv_path in venv_paths:
    if os.path.exists(venv_path) and venv_path not in sys.path:
        sys.path.insert(0, venv_path)

from flask import Flask
from flask_login import LoginManager
from models import db, login_manager
from models.user import User
import os

# Criar a aplicação Flask
app = Flask(__name__)

# Configurações
app.config['SECRET_KEY'] = 'sua-chave-secreta-super-segura-aqui-2025'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///ueci_monitoramento.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Inicializar extensões
db.init_app(app)
login_manager.init_app(app)
login_manager.login_view = 'auth.login'
login_manager.login_message = 'Por favor, faça login para acessar esta página.'
login_manager.login_message_category = 'info'

# User loader para Flask-Login
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Registrar blueprints
from routes.auth import auth_bp
from routes.main import main_bp
from routes.analytics import analytics_bp
from routes.configuracoes import config_bp

app.register_blueprint(auth_bp)
app.register_blueprint(main_bp)
app.register_blueprint(analytics_bp)
app.register_blueprint(config_bp)

# Criar tabelas e importar dados iniciais
def init_database():
    """Inicializa o banco de dados e importa dados"""
    with app.app_context():
        # Criar todas as tabelas
        db.create_all()
        print("✓ Banco de dados criado!")
        
        # Importar usuários iniciais
        from utils.import_data import create_initial_users, import_excel_data
        from models.planilha import PlanilhaData
        
        create_initial_users(app)
        
        # Importar dados do Excel apenas se não houver dados
        if PlanilhaData.query.count() == 0:
            if os.path.exists('MONITORAMENTO UECI - CONSOLIDADO.xlsx'):
                import_excel_data(app)
            else:
                print("\n⚠ Arquivo Excel não encontrado. Pule a importação de dados.")
        else:
            print("\nDados já existem no banco. Pulando importação.")

if __name__ == '__main__':
    # Verificar se o banco de dados existe
    if not os.path.exists('ueci_monitoramento.db'):
        print("Inicializando banco de dados pela primeira vez...")
        init_database()
    
    print("\n" + "="*60)
    print("🚀 UECI Monitoramento - Sistema Iniciado!")
    print("="*60)
    print("\n📊 Acesse o sistema em: http://localhost:5000")
    print("\n👤 Login de Administrador:")
    print("   Username: admin")
    print("   Senha: admin123")
    print("\n" + "="*60 + "\n")
    
    # Executar aplicação
    app.run(debug=True, host='0.0.0.0', port=5000)
