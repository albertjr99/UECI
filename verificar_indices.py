from app import app, db
from models.planilha import AbaConfig

app.app_context().push()

print("\n=== VERIFICAÇÃO DE ÍNDICES ===\n")

# UECI - campos 14 a 19
ueci = AbaConfig.query.filter_by(aba_name='Plano de Ação - UECI').first()
print("UECI - Campos 14-19 (total:", len(ueci.get_columns()), "campos):")
for i, col in enumerate(ueci.get_columns()):
    if i >= 14 and i <= 19:
        print(f"  {i}: {col['name']}")

# SECONT - campos 10 a 16
secont = AbaConfig.query.filter_by(aba_name='Plano de Ação - SECONT').first()
print("\nSECONT - Campos 10-16 (total:", len(secont.get_columns()), "campos):")
for i, col in enumerate(secont.get_columns()):
    if i >= 10 and i <= 16:
        print(f"  {i}: {col['name']}")

# TCEES - campos 10 a 16
tcees = AbaConfig.query.filter_by(aba_name='Plano de Ação - TCEES').first()
print("\nTCEES - Campos 10-16 (total:", len(tcees.get_columns()), "campos):")
for i, col in enumerate(tcees.get_columns()):
    if i >= 10 and i <= 16:
        print(f"  {i}: {col['name']}")
