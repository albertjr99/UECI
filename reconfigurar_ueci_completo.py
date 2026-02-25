"""Script para reconfigurar completamente a planilha UECI"""

from app import app, db
from models.planilha import AbaConfig

def reconfigurar_ueci():
    """Reconfigura a planilha UECI do zero"""
    
    with app.app_context():
        print("Reconfigurando planilha UECI...\n")
        
        # Buscar a aba UECI
        aba_ueci = AbaConfig.query.filter_by(aba_name='Plano de Ação - UECI').first()
        
        if not aba_ueci:
            print("❌ Planilha UECI não encontrada!")
            return
        
        # Campos da parte azul (ANTES da laranja) - 12 campos
        campos_azul_inicio = [
            {'name': 'Exercício', 'type': 'text'},
            {'name': 'Data do Envio', 'type': 'date'},
            {'name': 'Data da Ciência', 'type': 'date'},
            {'name': 'Responsável pela Análise', 'type': 'select'},
            {'name': 'Origem', 'type': 'select'},
            {'name': 'Tipo de Ação', 'type': 'select'},
            {'name': 'E-docs', 'type': 'text'},
            {'name': 'Ponto de Controle', 'type': 'text'},
            {'name': 'Unidade Gestora', 'type': 'select'},
            {'name': 'Constatação', 'type': 'textarea'},
            {'name': 'Recomendação', 'type': 'textarea'},
            {'name': 'Riscos Envolvidos', 'type': 'textarea'}
        ]
        
        # Campos da parte LARANJA (PREENCHIMENTO PELA ÁREA RESPONSÁVEL) - 7 campos
        campos_laranja = [
            {'name': 'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?', 'type': 'select'},
            {'name': 'Setor(es) responsável(is)', 'type': 'text'},
            {'name': 'Servidor (es) responsável', 'type': 'text'},
            {'name': 'Iniciativa da área', 'type': 'textarea'},
            {'name': 'Data da Resposta', 'type': 'date'},
            {'name': 'Prazo previsto de início', 'type': 'date'},
            {'name': 'Prazo previsto de término', 'type': 'date'}
        ]
        
        # Campos da parte azul (DEPOIS da laranja) - 2 campos
        campos_azul_fim = [
            {'name': 'Status para a UECI', 'type': 'select'},
            {'name': 'Análise do retorno da área', 'type': 'textarea'}
        ]
        
        # Juntar todos os campos na ordem correta
        todos_campos = campos_azul_inicio + campos_laranja + campos_azul_fim
        
        print(f"Total de campos: {len(todos_campos)}")
        print(f"  - Azul inicial: {len(campos_azul_inicio)} campos (1-12)")
        print(f"  - Laranja: {len(campos_laranja)} campos (13-19)")
        print(f"  - Azul final: {len(campos_azul_fim)} campos (20-21)")
        print()
        
        # Atualizar a configuração
        aba_ueci.set_columns(todos_campos)
        db.session.commit()
        
        print("✅ Planilha UECI reconfigurada com sucesso!")
        print("\nOrdem dos campos:")
        for i, campo in enumerate(todos_campos, 1):
            cor = "LARANJA" if i >= 13 and i <= 19 else "AZUL"
            print(f"{i:2d}. [{cor:7s}] {campo['name']}")

if __name__ == '__main__':
    reconfigurar_ueci()
