from app import app, db
from models.planilha import PlanilhaData, AbaConfig
import pandas as pd
from datetime import datetime

app.app_context().push()

# Ler o arquivo Excel
excel_file = r'c:\Users\albert.junior\OneDrive\UECI\planilha ueci - dados.xlsx'
df = pd.read_excel(excel_file)

print(f"Arquivo lido com sucesso: {len(df)} linhas encontradas")
print(f"Colunas: {list(df.columns)}")
print("\nPrimeiras linhas:")
print(df.head())

# Obter configuração da aba UECI
ueci_config = AbaConfig.query.filter_by(aba_name='Plano de Ação - UECI').first()
campos_esperados = [col['name'] for col in ueci_config.get_columns()]

print(f"\nCampos esperados ({len(campos_esperados)}):")
for i, campo in enumerate(campos_esperados):
    print(f"{i}: {campo}")

# Inserir dados
contador = 0
for idx, row in df.iterrows():
    row_data = {}
    
    # Mapear cada coluna do Excel para os campos esperados
    for i, col_name in enumerate(df.columns):
        if i < len(campos_esperados):
            campo = campos_esperados[i]
            valor = row[col_name]
            
            # Converter NaN e valores vazios para string vazia
            if pd.isna(valor):
                valor = ''
            # Converter datas para string no formato brasileiro
            elif isinstance(valor, pd.Timestamp):
                valor = valor.strftime('%Y-%m-%d')
            else:
                valor = str(valor)
            
            row_data[campo] = valor
    
    # Criar novo registro
    novo_registro = PlanilhaData()
    novo_registro.aba_name = 'Plano de Ação - UECI'
    novo_registro.set_data(row_data)
    novo_registro.row_order = idx + 1
    
    db.session.add(novo_registro)
    contador += 1
    
    if (idx + 1) % 5 == 0:
        print(f"Processadas {idx + 1} linhas...")

db.session.commit()
print(f"\n✅ Importação concluída! {contador} registros inseridos na planilha UECI.")
