import sys
sys.path.insert(0, '/home/aicsj/ueci_sistema')

from app import app, db
from sqlalchemy import text

with app.app_context():
    # Criar tabela usando text() para SQLAlchemy 2.x
    db.session.execute(text('''
        CREATE TABLE IF NOT EXISTS links_temporarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            token VARCHAR(200) NOT NULL UNIQUE,
            aba_name VARCHAR(100) NOT NULL,
            registro_id INTEGER NOT NULL,
            expiracao DATETIME NOT NULL,
            criado_por INTEGER,
            criado_em DATETIME,
            usado_em DATETIME,
            FOREIGN KEY(criado_por) REFERENCES users (id)
        )
    '''))
    
    db.session.commit()
    print("✅ Tabela links_temporarios criada com sucesso!")
