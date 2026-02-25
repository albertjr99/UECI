from app import app, db
from models.planilha import PlanilhaData
import pandas as pd

app.app_context().push()

# Ler o arquivo Excel
excel_file = r'c:\Users\albert.junior\OneDrive\UECI\planilha ueci - dados.xlsx'
df = pd.read_excel(excel_file)

# Pegar primeira linha do Excel
primeira_linha = df.iloc[0]

# Mapeamento dos campos
mapeamento = {
    'Exercício': 'Exercício',
    'Data do Envio': 'Data do Envio',
    'Data da Ciência': 'Data da Ciência',
    'Responsável pela Análise': 'Responsável pela Análise',
    'Origem': 'Origem',
    'Tipo de Ação': 'Tipo de Ação',
    'E-docs': 'E-docs',
    'Ponto de Controle': 'Ponto de Controle',
    'Recomendação': 'Recomendação',
    'Riscos Envolvidos': 'Riscos Envolvidos',
    'STATUS DA RECOMENDAÇÃO\nQual é a situação atual da recomendação?': 'STATUS DA RECOMENDAÇÃO',
    'Servidor (es) responsável': 'Servidor(es) responsável(is)',
    'Iniciativa da área': 'Iniciativa da área',
    'Data da Resposta': 'Data da Resposta',
    'Prazo previsto de ínicio': 'Prazo previsto de início',
    'Prazo previsto de término': 'Prazo previsto de conclusão',
    'Status para a UECI': 'Status para a UECI',
    'Análise do retorno da área': 'Análise do retorno da área'
}

# Construir dados restaurados
dados_restaurados = {}
for col_excel, col_ueci in mapeamento.items():
    if col_excel in df.columns:
        valor = primeira_linha[col_excel]
        
        if pd.isna(valor):
            valor = ''
        elif isinstance(valor, pd.Timestamp):
            valor = valor.strftime('%Y-%m-%d')
        else:
            valor = str(valor)
            if ' ' in valor and ':' in valor:
                valor = valor.split(' ')[0]
        
        dados_restaurados[col_ueci] = valor

print("Dados originais do Excel:")
for campo, valor in dados_restaurados.items():
    print(f"{campo}: {valor}")

# Restaurar no banco
primeiro_reg = PlanilhaData.query.filter_by(aba_name='Plano de Ação - UECI').order_by(PlanilhaData.row_order).first()
if primeiro_reg:
    primeiro_reg.set_data(dados_restaurados)
    db.session.commit()
    print("\n✅ Primeiro registro restaurado com sucesso!")
else:
    print("\n❌ Registro não encontrado!")
