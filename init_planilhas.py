"""Script para criar as planilhas iniciais no banco de dados"""

from app import app, db
from models.planilha import AbaConfig

def create_planilhas():
    """Cria as 3 planilhas principais do sistema"""
    
    with app.app_context():
        print("Criando planilhas no banco de dados...")
        
        # Definir os campos padrão baseado na estrutura moderna
        campos_padrao = [
            {'name': 'Exercício', 'type': 'text'},
            {'name': 'Data do Envio', 'type': 'date'},
            {'name': 'Data da Ciência', 'type': 'date'},
            {'name': 'Responsável pela Análise', 'type': 'text'},
            {'name': 'Origem', 'type': 'text'},
            {'name': 'Tipo de Ação', 'type': 'text'},
            {'name': 'E-docs', 'type': 'text'},
            {'name': 'Ponto de Controle', 'type': 'text'},
            {'name': 'UG', 'type': 'text'},
            {'name': 'Constatação', 'type': 'textarea'},
            {'name': 'Recomendação', 'type': 'textarea'},
            {'name': 'Riscos Envolvidos', 'type': 'textarea'},
            {'name': 'STATUS DA RECOMENDAÇÃO', 'type': 'select'},
            {'name': 'Servidor(es) responsável(is)', 'type': 'text'},
            {'name': 'Observações', 'type': 'textarea'},
            {'name': 'Iniciativa da área', 'type': 'textarea'},
            {'name': 'Data da Resposta', 'type': 'date'},
            {'name': 'Prazo previsto de conclusão', 'type': 'date'},
            {'name': 'Status para a UECI', 'type': 'select'},
            {'name': 'Análise do retorno da área', 'type': 'textarea'}
        ]
        
        # Definir as 3 planilhas
        planilhas = [
            {
                'aba_name': 'Plano de Ação - UECI',
                'display_order': 1,
                'description': 'Plano de Ação - UECI'
            },
            {
                'aba_name': 'Plano de Ação - SECONT',
                'display_order': 2,
                'description': 'Plano de Ação - SECONT'
            },
            {
                'aba_name': 'Plano de Ação - TCEES',
                'display_order': 3,
                'description': 'Plano de Ação - TCEES'
            }
        ]
        
        # Criar ou atualizar cada planilha
        for planilha_info in planilhas:
            aba_config = AbaConfig.query.filter_by(aba_name=planilha_info['aba_name']).first()
            
            if aba_config:
                print(f"✓ Atualizando planilha: {planilha_info['aba_name']}")
            else:
                print(f"✓ Criando planilha: {planilha_info['aba_name']}")
                aba_config = AbaConfig(
                    aba_name=planilha_info['aba_name'],
                    display_order=planilha_info['display_order'],
                    is_active=True
                )
                db.session.add(aba_config)
            
            # Configurar as colunas
            aba_config.set_columns(campos_padrao)
            aba_config.is_active = True
        
        # Salvar no banco
        db.session.commit()
        print("\n✅ Planilhas criadas com sucesso!")
        print("\nPlanilhas disponíveis:")
        
        for aba in AbaConfig.query.order_by(AbaConfig.display_order).all():
            print(f"  {aba.display_order}. {aba.aba_name}")
        
        print("\n🚀 Você já pode acessar o sistema!")

if __name__ == '__main__':
    create_planilhas()
