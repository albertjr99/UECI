"""Script para reimportar dados do Excel para UECI com a nova estrutura de 21 campos"""

import pandas as pd
from app import app, db
from models.planilha import PlanilhaData, AbaConfig
from datetime import datetime

def limpar_dados_ueci():
    """Remove todos os registros da UECI"""
    with app.app_context():
        print("Limpando dados antigos da UECI...")
        PlanilhaData.query.filter_by(aba_name='Plano de Ação - UECI').delete()
        db.session.commit()
        print("✅ Dados antigos removidos\n")

def reimportar_excel():
    """Reimporta dados do Excel com a nova estrutura"""
    
    with app.app_context():
        print("Importando dados do Excel...\n")
        
        # Ler o arquivo Excel
        try:
            df = pd.read_excel('planilha ueci - dados.xlsx')
        except:
            df = pd.read_excel('MONITORAMENTO UECI - CONSOLIDADO.xlsx')
        
        print(f"Total de linhas no Excel: {len(df)}")
        print(f"Colunas encontradas: {list(df.columns)}\n")
        
        # Buscar configuração da UECI
        aba = AbaConfig.query.filter_by(aba_name='Plano de Ação - UECI').first()
        if not aba:
            print("❌ Aba UECI não encontrada!")
            return
        
        campos_ueci = aba.get_columns()
        print(f"Campos configurados na UECI: {len(campos_ueci)}\n")
        
        # Mapear colunas do Excel para campos da UECI
        mapeamento = {
            'Exercício': ['Exercício', 'EXERCÍCIO', 'Ano'],
            'Data do Envio': ['Data do Envio', 'DATA DO ENVIO', 'Data Envio'],
            'Data da Ciência': ['Data da Ciência', 'DATA DA CIÊNCIA', 'Data Ciencia', 'Data da Ciencia'],
            'Responsável pela Análise': ['Responsável pela Análise', 'RESPONSÁVEL PELA ANÁLISE', 'Responsavel pela Analise'],
            'Origem': ['Origem', 'ORIGEM'],
            'Tipo de Ação': ['Tipo de Ação', 'TIPO DE AÇÃO', 'Tipo de Acao'],
            'E-docs': ['E-docs', 'E-DOCS', 'Edocs', 'E docs'],
            'Ponto de Controle': ['Ponto de Controle', 'PONTO DE CONTROLE'],
            'Unidade Gestora': ['UG', 'Unidade Gestora', 'UNIDADE GESTORA'],
            'Constatação': ['Constatação', 'CONSTATAÇÃO', 'Constatacao'],
            'Recomendação': ['Recomendação', 'RECOMENDAÇÃO', 'Recomendacao'],
            'Riscos Envolvidos': ['Riscos Envolvidos', 'RISCOS ENVOLVIDOS'],
            'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': ['STATUS DA RECOMENDAÇÃO', 'Status da Recomendação', 'Status'],
            'Setor(es) responsável(is)': ['Setor(es) responsável(is)', 'Setor responsável', 'Setor'],
            'Servidor (es) responsável': ['Servidor(es) responsável(is)', 'Servidor (es) responsável', 'Servidor responsável', 'Servidor'],
            'Iniciativa da área': ['Iniciativa da área', 'INICIATIVA DA ÁREA', 'Retorno da área', 'Retorno'],
            'Data da Resposta': ['Data da Resposta', 'DATA DA RESPOSTA', 'Data Resposta'],
            'Prazo previsto de início': ['Prazo previsto de início', 'PRAZO PREVISTO DE INÍCIO', 'Prazo inicio'],
            'Prazo previsto de término': ['Prazo previsto de término', 'PRAZO PREVISTO DE TÉRMINO', 'Prazo termino', 'Prazo previsto de conclusão'],
            'Status para a UECI': ['Status para a UECI', 'STATUS PARA A UECI'],
            'Análise do retorno da área': ['Análise do retorno da área', 'ANÁLISE DO RETORNO DA ÁREA', 'Analise']
        }
        
        def encontrar_coluna(campo_destino):
            """Encontra a coluna do Excel que corresponde ao campo"""
            if campo_destino in mapeamento:
                for variacao in mapeamento[campo_destino]:
                    if variacao in df.columns:
                        return variacao
            return None
        
        def formatar_data(valor):
            """Converte data para formato YYYY-MM-DD"""
            if pd.isna(valor) or valor == '':
                return ''
            
            try:
                if isinstance(valor, pd.Timestamp):
                    return valor.strftime('%Y-%m-%d')
                
                valor_str = str(valor)
                if ' ' in valor_str:
                    valor_str = valor_str.split(' ')[0]
                
                if '-' in valor_str and len(valor_str.split('-')[0]) == 4:
                    return valor_str
                
                if '/' in valor_str:
                    partes = valor_str.split('/')
                    if len(partes) == 3:
                        if len(partes[2]) == 4:
                            return f"{partes[2]}-{partes[1].zfill(2)}-{partes[0].zfill(2)}"
                        else:
                            return f"{partes[0]}-{partes[1].zfill(2)}-{partes[2].zfill(2)}"
                
                return ''
            except:
                return ''
        
        # Importar cada linha
        registros_importados = 0
        for idx, row in df.iterrows():
            dados = {}
            
            for campo in campos_ueci:
                campo_nome = campo['name']
                campo_tipo = campo['type']
                coluna_excel = encontrar_coluna(campo_nome)
                
                if coluna_excel and coluna_excel in df.columns:
                    valor = row[coluna_excel]
                    
                    if pd.isna(valor):
                        dados[campo_nome] = ''
                    elif campo_tipo == 'date':
                        dados[campo_nome] = formatar_data(valor)
                    else:
                        dados[campo_nome] = str(valor).strip()
                else:
                    dados[campo_nome] = ''
            
            # Criar registro
            novo_registro = PlanilhaData(
                aba_name='Plano de Ação - UECI',
                row_order=idx + 1
            )
            novo_registro.set_data(dados)
            db.session.add(novo_registro)
            registros_importados += 1
            
            print(f"Registro {registros_importados}: Exercício={dados.get('Exercício', 'N/A')}, UG={dados.get('Unidade Gestora', 'N/A')}")
        
        db.session.commit()
        print(f"\n✅ {registros_importados} registros importados com sucesso!")

if __name__ == '__main__':
    limpar_dados_ueci()
    reimportar_excel()
