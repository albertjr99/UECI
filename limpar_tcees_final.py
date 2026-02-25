"""Script para limpar completamente a planilha TCEES e deixar apenas 2 registros"""

from app import app, db
from models.planilha import PlanilhaData

def limpar_tcees_completo():
    """Remove TODOS os registros e depois verifica o que restou"""
    
    with app.app_context():
        print("Verificando planilha TCEES...\n")
        
        # Buscar todos os registros do TCEES
        total_antes = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').count()
        print(f"Total de registros ANTES: {total_antes}\n")
        
        # Buscar todos os registros
        registros = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').order_by(PlanilhaData.id).all()
        
        # Manter apenas os 2 primeiros IDs mais baixos (originais)
        ids_manter = []
        exercicios_vistos = set()
        
        for reg in registros:
            data = reg.get_data()
            exercicio = data.get('Exercício', '')
            
            if exercicio not in exercicios_vistos and len(ids_manter) < 2:
                ids_manter.append(reg.id)
                exercicios_vistos.add(exercicio)
                print(f"✅ Mantendo ID {reg.id} - Exercício: {exercicio}")
            else:
                print(f"❌ Removendo ID {reg.id} - Exercício: {exercicio} (duplicado)")
                db.session.delete(reg)
        
        # Salvar alterações
        db.session.commit()
        
        # Verificar novamente
        total_depois = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').count()
        print(f"\n✅ Limpeza concluída!")
        print(f"Total de registros DEPOIS: {total_depois}")
        
        # Mostrar os registros finais
        print("\nRegistros finais:")
        registros_finais = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').order_by(PlanilhaData.row_order).all()
        for i, reg in enumerate(registros_finais, 1):
            data = reg.get_data()
            print(f"{i}. Exercício: {data.get('Exercício', '')} - Documento: {data.get('Documento nº', '')}")

if __name__ == '__main__':
    limpar_tcees_completo()
