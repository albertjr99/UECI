from app import app, db
from models.planilha import PlanilhaData, AbaConfig
import pandas as pd
from datetime import datetime

app.app_context().push()

# Ler o arquivo Excel
excel_file = r'c:\Users\albert.junior\OneDrive\UECI\planilha ueci - dados.xlsx'
df = pd.read_excel(excel_file)

print(f"Arquivo lido: {len(df)} linhas")

# Campos esperados da UECI (18 campos)
campos_ueci = [
    'Exercício',
    'Data do Envio',
    'Data da Ciência',
    'Responsável pela Análise',
    'Origem',
    'Tipo de Ação',
    'E-docs',
    'Ponto de Controle',
    'Recomendação',
    'Riscos Envolvidos',
    'STATUS DA RECOMENDAÇÃO',
    'Servidor(es) responsável(is)',
    'Iniciativa da área',
    'Data da Resposta',
    'Prazo previsto de início',
    'Prazo previsto de conclusão',
    'Status para a UECI',
    'Análise do retorno da área'
]

# Mapeamento de colunas do Excel para campos UECI
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

# Inserir dados
contador = 0
for idx, row in df.iterrows():
    row_data = {}
    
    # Mapear apenas os campos que existem
    for col_excel, col_ueci in mapeamento.items():
        if col_excel in df.columns:
            valor = row[col_excel]
            
            # Converter NaN e valores vazios para string vazia
            if pd.isna(valor):
                valor = ''
            # Converter datas para string YYYY-MM-DD (sem hora)
            elif isinstance(valor, pd.Timestamp):
                valor = valor.strftime('%Y-%m-%d')
            else:
                valor = str(valor)
                # Remover timestamp se existir
                if ' ' in valor and ':' in valor:
                    valor = valor.split(' ')[0]
            
            row_data[col_ueci] = valor
    
    # Garantir que todos os campos existam (mesmo que vazios)
    for campo in campos_ueci:
        if campo not in row_data:
            row_data[campo] = ''
    
    # Criar novo registro
    novo_registro = PlanilhaData()
    novo_registro.aba_name = 'Plano de Ação - UECI'
    novo_registro.set_data(row_data)
    novo_registro.row_order = idx + 1
    
    db.session.add(novo_registro)
    contador += 1

db.session.commit()
print(f"\n✅ Importação concluída! {contador} registros inseridos com {len(campos_ueci)} campos cada.")
