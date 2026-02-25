from datetime import datetime
from models import db
import json

class PlanilhaData(db.Model):
    __tablename__ = 'planilha_data'

    id = db.Column(db.Integer, primary_key=True)
    aba_name = db.Column(db.String(100), nullable=False)  # Nome da aba
    row_data = db.Column(db.Text, nullable=False)  # JSON com os dados da linha
    row_order = db.Column(db.Integer, default=0)  # Ordem da linha
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_by = db.Column(db.Integer, db.ForeignKey('users.id'))

    creator = db.relationship('User', foreign_keys=[created_by])
    updater = db.relationship('User', foreign_keys=[updated_by])

    def set_data(self, data_dict):
        """Converte dicionário em JSON"""
        self.row_data = json.dumps(data_dict, ensure_ascii=False)

    def get_data(self):
        """Retorna os dados como dicionário"""
        return json.loads(self.row_data) if self.row_data else {}

    def __repr__(self):
        return f'<PlanilhaData {self.aba_name} - Row {self.row_order}>'


class AbaConfig(db.Model):
    __tablename__ = 'aba_config'

    id = db.Column(db.Integer, primary_key=True)
    aba_name = db.Column(db.String(100), unique=True, nullable=False)
    columns_config = db.Column(db.Text)
    orange_columns = db.Column(db.Text)  # ADICIONE ESTA LINHA
    display_order = db.Column(db.Integer, default=0)
    is_active = db.Column(db.Boolean, default=True)

    def set_columns(self, columns_list):
        """Define as colunas da aba"""
        self.columns_config = json.dumps(columns_list, ensure_ascii=False)

    def get_columns(self):
        """Retorna as colunas como lista"""
        return json.loads(self.columns_config) if self.columns_config else []

    # ADICIONE ESTES DOIS MÉTODOS:
    def set_orange_columns(self, column_names_list):
        """Define quais colunas ficam laranjas (por nome)"""
        self.orange_columns = json.dumps(column_names_list, ensure_ascii=False)

    def get_orange_columns(self):
        """Retorna lista de nomes das colunas laranjas"""
        return json.loads(self.orange_columns) if self.orange_columns else []

    def __repr__(self):
        return f'<AbaConfig {self.aba_name}>'