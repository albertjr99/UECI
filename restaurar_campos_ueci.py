from app import app, db
from models.planilha import AbaConfig

app.app_context().push()

# Campos originais da UECI (18 campos)
campos_ueci_originais = [
    {'name': 'Exercício', 'type': 'text'},
    {'name': 'Data do Envio', 'type': 'date'},
    {'name': 'Data da Ciência', 'type': 'date'},
    {'name': 'Responsável pela Análise', 'type': 'text'},
    {'name': 'Origem', 'type': 'select'},
    {'name': 'Tipo de Ação', 'type': 'select'},
    {'name': 'E-docs', 'type': 'text'},
    {'name': 'Ponto de Controle', 'type': 'text'},
    {'name': 'Recomendação', 'type': 'textarea'},
    {'name': 'Riscos Envolvidos', 'type': 'textarea'},
    {'name': 'STATUS DA RECOMENDAÇÃO', 'type': 'select'},
    {'name': 'Servidor(es) responsável(is)', 'type': 'text'},
    {'name': 'Iniciativa da área', 'type': 'textarea'},
    {'name': 'Data da Resposta', 'type': 'date'},
    {'name': 'Prazo previsto de início', 'type': 'date'},
    {'name': 'Prazo previsto de conclusão', 'type': 'date'},
    {'name': 'Status para a UECI', 'type': 'select'},
    {'name': 'Análise do retorno da área', 'type': 'textarea'}
]

# Atualizar configuração
ueci = AbaConfig.query.filter_by(aba_name='Plano de Ação - UECI').first()
if ueci:
    ueci.set_columns(campos_ueci_originais)
    db.session.commit()
    print(f"✅ Configuração da UECI restaurada para {len(campos_ueci_originais)} campos!")
    print("\nCampos:")
    for i, campo in enumerate(campos_ueci_originais):
        print(f"{i}: {campo['name']}")
else:
    print("❌ UECI não encontrada!")
