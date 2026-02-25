"""Script para resetar a planilha TCEES com apenas 2 registros corretos"""

from app import app, db
from models.planilha import PlanilhaData

def resetar_tcees():
    """Remove todos os registros TCEES e deixa apenas 2"""
    
    with app.app_context():
        print("Resetando planilha TCEES...\n")
        
        # Remover TODOS os registros TCEES
        total_removido = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').delete()
        print(f"Removidos {total_removido} registros antigos\n")
        
        db.session.commit()
        
        # Verificar
        total_final = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').count()
        print(f"✅ Planilha TCEES resetada!")
        print(f"Total de registros: {total_final}")
        print("\nAgora reinicie o app.py para importar os 2 registros corretos do Excel.")

if __name__ == '__main__':
    resetar_tcees()
