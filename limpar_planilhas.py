"""Script para limpar planilhas antigas e manter apenas as 3 principais"""

from app import app, db
from models.planilha import AbaConfig, PlanilhaData

def limpar_planilhas():
    """Remove planilhas antigas e mantém apenas as 3 principais"""
    
    with app.app_context():
        print("Limpando planilhas antigas do banco de dados...\n")
        
        # Planilhas que devem ser mantidas
        planilhas_manter = [
            'Plano de Ação - UECI',
            'Plano de Ação - SECONT',
            'Plano de Ação - TCEES'
        ]
        
        # Buscar todas as planilhas
        todas_abas = AbaConfig.query.all()
        
        print(f"Total de planilhas encontradas: {len(todas_abas)}\n")
        
        # Desativar ou remover planilhas que não estão na lista
        for aba in todas_abas:
            if aba.aba_name not in planilhas_manter:
                print(f"❌ Desativando: {aba.aba_name}")
                aba.is_active = False
                
                # Remover dados associados a esta aba
                dados_removidos = PlanilhaData.query.filter_by(aba_name=aba.aba_name).delete()
                if dados_removidos > 0:
                    print(f"   Removidos {dados_removidos} registros")
            else:
                print(f"✅ Mantendo: {aba.aba_name}")
                aba.is_active = True
        
        # Reordenar as planilhas ativas
        print("\n\nReorganizando ordem das planilhas...")
        for idx, nome in enumerate(planilhas_manter, start=1):
            aba = AbaConfig.query.filter_by(aba_name=nome).first()
            if aba:
                aba.display_order = idx
                print(f"{idx}. {nome}")
        
        # Salvar alterações
        db.session.commit()
        
        print("\n✅ Limpeza concluída!")
        print("\nPlanilhas ativas:")
        
        for aba in AbaConfig.query.filter_by(is_active=True).order_by(AbaConfig.display_order).all():
            total_registros = PlanilhaData.query.filter_by(aba_name=aba.aba_name).count()
            print(f"  {aba.display_order}. {aba.aba_name} ({total_registros} registros)")

if __name__ == '__main__':
    limpar_planilhas()
