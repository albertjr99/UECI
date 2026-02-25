from datetime import datetime
from models import db
import json

class ConfiguracaoSistema(db.Model):
    """Configurações globais do sistema"""
    __tablename__ = 'configuracao_sistema'
    
    id = db.Column(db.Integer, primary_key=True)
    chave = db.Column(db.String(100), unique=True, nullable=False)
    valor = db.Column(db.Text)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    updated_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    
    updater = db.relationship('User', foreign_keys=[updated_by])
    
    def set_valor(self, valor_dict):
        """Converte dicionário em JSON"""
        self.valor = json.dumps(valor_dict, ensure_ascii=False)
    
    def get_valor(self):
        """Retorna o valor como dicionário"""
        if self.valor:
            try:
                return json.loads(self.valor)
            except:
                return {}
        return {}
    
    def __repr__(self):
        return f'<ConfiguracaoSistema {self.chave}>'


class DropdownConfig(db.Model):
    """Configuração de listas suspensas por campo"""
    __tablename__ = 'dropdown_config'
    
    id = db.Column(db.Integer, primary_key=True)
    aba_name = db.Column(db.String(100), nullable=False)
    campo_nome = db.Column(db.String(200), nullable=False)
    opcoes = db.Column(db.Text)  # JSON com as opções
    is_active = db.Column(db.Boolean, default=True)
    ordem = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    
    creator = db.relationship('User', foreign_keys=[created_by])
    updater = db.relationship('User', foreign_keys=[updated_by])
    
    def set_opcoes(self, opcoes_list):
        """Define as opções do dropdown"""
        self.opcoes = json.dumps(opcoes_list, ensure_ascii=False)
    
    def get_opcoes(self):
        """Retorna as opções como lista"""
        if self.opcoes:
            try:
                return json.loads(self.opcoes)
            except:
                return []
        return []
    
    def __repr__(self):
        return f'<DropdownConfig {self.aba_name} - {self.campo_nome}>'
