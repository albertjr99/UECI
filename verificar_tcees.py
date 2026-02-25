"""Script para verificar e limpar dados duplicados da planilha TCEES"""

from app import app, db
from models.planilha import PlanilhaData

def verificar_tcees():
    """Verifica os registros da planilha TCEES"""
    
    with app.app_context():
        print("Verificando registros da planilha TCEES...\n")
        
        # Buscar todos os registros do TCEES
        registros = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').all()
        
        print(f"Total de registros encontrados: {len(registros)}\n")
        
        # Mostrar informações de cada registro
        registros_unicos = {}
        
        for i, reg in enumerate(registros, 1):
            data = reg.get_data()
            exercicio = data.get('Exercício', '')
            documento = data.get('Documento nº', '')
            responsavel = data.get('Responsável pela Análise', '')
            
            # Criar uma chave única baseada no conteúdo
            chave = f"{exercicio}_{documento}_{responsavel}"
            
            print(f"Registro {i} (ID: {reg.id}):")
            print(f"  Exercício: {exercicio}")
            print(f"  Documento: {documento}")
            print(f"  Responsável: {responsavel}")
            print(f"  Row Order: {reg.row_order}")
            
            if chave in registros_unicos:
                print(f"  ⚠️  DUPLICADO do registro {registros_unicos[chave]['num']}")
                print(f"  ❌ Marcado para remoção")
                # Remover este registro duplicado
                db.session.delete(reg)
            else:
                registros_unicos[chave] = {'num': i, 'id': reg.id}
                print(f"  ✓ Registro único")
            
            print()
        
        # Salvar alterações
        db.session.commit()
        
        # Verificar novamente
        total_final = PlanilhaData.query.filter_by(aba_name='Plano de Ação - TCEES').count()
        print(f"\n✅ Limpeza concluída!")
        print(f"Total de registros após limpeza: {total_final}")

if __name__ == '__main__':
    verificar_tcees()
