from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager

db = SQLAlchemy()
login_manager = LoginManager()

# Import models
from models.user import User
from models.planilha import PlanilhaData, AbaConfig
from models.configuracoes import ConfiguracaoSistema, DropdownConfig
