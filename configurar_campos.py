"""Script para configurar campos específicos para cada planilha"""

from app import app, db
from models.planilha import AbaConfig

def configurar_campos_planilhas():
    """Define campos específicos para cada planilha"""
    
    with app.app_context():
        print("Configurando campos específicos para cada planilha...\n")
        
        # Campos específicos da UECI (parte azul - antes da laranja)
        campos_ueci = [
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
        
        # Campos compartilhados (parte laranja - PREENCHIMENTO PELA ÁREA RESPONSÁVEL) para UECI
        campos_laranja_ueci = [
            {'name': 'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?', 'type': 'select'},
            {'name': 'Setor(es) responsável(is)', 'type': 'text'},
            {'name': 'Servidor (es) responsável', 'type': 'text'},
            {'name': 'Iniciativa da área', 'type': 'textarea'},
            {'name': 'Data da Resposta', 'type': 'date'},
            {'name': 'Prazo previsto de início', 'type': 'date'},
            {'name': 'Prazo previsto de término', 'type': 'date'}
        ]
        
        # Campos finais da UECI (parte azul - depois da laranja)
        campos_ueci_finais = [
            {'name': 'Status para a UECI', 'type': 'select'},
            {'name': 'Análise do retorno da área', 'type': 'textarea'}
        ]
        
        # Campos compartilhados (parte amarela) para SECONT e TCEES
        campos_amarelos_secont_tcees = [
            {'name': 'STATUS DA RECOMENDAÇÃO', 'type': 'select'},
            {'name': 'Setor(es) responsável(is)', 'type': 'text'},
            {'name': 'Servidor(es) responsável', 'type': 'text'},
            {'name': 'Retorno da área', 'type': 'textarea'},
            {'name': 'Data da Resposta', 'type': 'date'},
            {'name': 'Prazo previsto de início', 'type': 'date'},
            {'name': 'Prazo previsto de término', 'type': 'date'},
            {'name': 'Status para a UECI', 'type': 'select'}
        ]
        
        # Campos específicos da UECI (parte azul - 12 colunas)
        campos_ueci = [
            {'name': 'Exercício', 'type': 'text'},
            {'name': 'Data do Envio', 'type': 'date'},
            {'name': 'Data da Ciência', 'type': 'date'},
            {'name': 'Responsável pela Análise', 'type': 'select'},
            {'name': 'Origem', 'type': 'select'},
            {'name': 'Tipo de Ação', 'type': 'select'},
            {'name': 'UNIDADE GESTORA', 'type': 'select'},
            {'name': 'E-docs', 'type': 'text'},
            {'name': 'Ponto de Controle', 'type': 'text'},
            {'name': 'Recomendação', 'type': 'textarea'},
            {'name': 'Riscos Envolvidos', 'type': 'textarea'}
        ]
        
        # Campos específicos da SECONT e TCEES (parte azul - 9 colunas)
        campos_secont_tcees = [
            {'name': 'Exercício', 'type': 'text'},
            {'name': 'Documento nº', 'type': 'text'},
            {'name': 'Data do encaminhamento', 'type': 'date'},
            {'name': 'Data da ciência', 'type': 'date'},
            {'name': 'Responsável pela Análise', 'type': 'text'},
            {'name': 'Tipo de Ação', 'type': 'text'},
            {'name': 'E-docs', 'type': 'text'},
            {'name': 'Constatação', 'type': 'textarea'},
            {'name': 'Recomendação', 'type': 'textarea'}
        ]
        
        # Configurar UECI
        aba_ueci = AbaConfig.query.filter_by(aba_name='Plano de Ação - UECI').first()
        if aba_ueci:
            campos_completos_ueci = campos_ueci + campos_laranja_ueci + campos_ueci_finais
            aba_ueci.set_columns(campos_completos_ueci)
            print(f"✅ Plano de Ação - UECI: {len(campos_completos_ueci)} campos configurados")
        
        # Configurar SECONT
        aba_secont = AbaConfig.query.filter_by(aba_name='Plano de Ação - SECONT').first()
        if aba_secont:
            campos_completos_secont = campos_secont_tcees + campos_amarelos_secont_tcees
            aba_secont.set_columns(campos_completos_secont)
            print(f"✅ Plano de Ação - SECONT: {len(campos_completos_secont)} campos configurados")
        
        # Configurar TCEES
        aba_tcees = AbaConfig.query.filter_by(aba_name='Plano de Ação - TCEES').first()
        if aba_tcees:
            campos_completos_tcees = campos_secont_tcees + campos_amarelos_secont_tcees
            aba_tcees.set_columns(campos_completos_tcees)
            print(f"✅ Plano de Ação - TCEES: {len(campos_completos_tcees)} campos configurados")
        
        db.session.commit()
        print("\n✅ Configuração concluída com sucesso!")

if __name__ == '__main__':
    configurar_campos_planilhas()
