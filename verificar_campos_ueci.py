from app import app
from models.planilha import AbaConfig

app.app_context().push()

ueci = AbaConfig.query.filter_by(aba_name='Plano de Ação - UECI').first()
campos = ueci.get_columns()

print('Campos atuais UECI:')
for i, c in enumerate(campos):
    print(f'{i}: {c["name"]}')
