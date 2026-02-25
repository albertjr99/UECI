from app import app, db
from models.planilha import PlanilhaData

app.app_context().push()

# Deletar todos os registros da UECI
registros = PlanilhaData.query.filter_by(aba_name='Plano de Ação - UECI').all()
print(f"Encontrados {len(registros)} registros na UECI")

for registro in registros:
    db.session.delete(registro)

db.session.commit()
print("Todos os registros da UECI foram removidos!")
