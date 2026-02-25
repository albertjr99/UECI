"""
Script para limpar e reimportar os dados do Excel
"""
import os
import sys

# Verificar se o banco existe
if os.path.exists('ueci_monitoramento.db'):
    print("🗑️  Removendo banco de dados antigo...")
    os.remove('ueci_monitoramento.db')
    print("✓ Banco de dados removido!")

print("\n🔄 Execute 'python app.py' para recriar o banco com os dados corretos.\n")
