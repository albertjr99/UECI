
from app import db
from datetime import datetime

class LinkTemporario(db.Model):
    __tablename__ = 'links_temporarios'
    
    id = db.Column(db.Integer, primary_key=True)
    token = db.Column(db.String(200), unique=True, nullable=False, index=True)
    aba_name = db.Column(db.String(100), nullable=False)
    registro_id = db.Column(db.Integer, nullable=False)
    expiracao = db.Column(db.DateTime, nullable=False)
    criado_por = db.Column(db.Integer, db.ForeignKey('users.id'))
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    usado_em = db.Column(db.DateTime)
    
    def __repr__(self):
        return f'<LinkTemporario {self.token[:10]}... para registro #{self.registro_id}>'

