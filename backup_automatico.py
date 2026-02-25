"""
Sistema de Backup Automático
Realiza backup do banco de dados em intervalos regulares
"""
import os
import shutil
from datetime import datetime
from pathlib import Path

def criar_backup():
    """Cria backup do banco de dados com timestamp"""
    
    # Diretório de backups
    backup_dir = Path(__file__).parent / 'backups'
    backup_dir.mkdir(exist_ok=True)
    
    # Arquivo do banco de dados
    db_file = Path(__file__).parent / 'instance' / 'ueci_monitoramento.db'
    
    if not db_file.exists():
        print(f"❌ Banco de dados não encontrado: {db_file}")
        return None
    
    # Nome do arquivo de backup com timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_file = backup_dir / f'backup_{timestamp}.db'
    
    try:
        # Copiar arquivo
        shutil.copy2(db_file, backup_file)
        
        # Obter tamanho do arquivo
        tamanho_mb = backup_file.stat().st_size / (1024 * 1024)
        
        print(f"✅ Backup criado com sucesso!")
        print(f"   Arquivo: {backup_file.name}")
        print(f"   Tamanho: {tamanho_mb:.2f} MB")
        
        # Limpar backups antigos (manter apenas os últimos 30)
        limpar_backups_antigos(backup_dir)
        
        return backup_file
    
    except Exception as e:
        print(f"❌ Erro ao criar backup: {e}")
        return None

def limpar_backups_antigos(backup_dir, manter=30):
    """Remove backups antigos, mantendo apenas os mais recentes"""
    
    backups = sorted(backup_dir.glob('backup_*.db'), key=lambda x: x.stat().st_mtime, reverse=True)
    
    if len(backups) > manter:
        removidos = 0
        for backup in backups[manter:]:
            try:
                backup.unlink()
                removidos += 1
            except Exception as e:
                print(f"⚠️  Erro ao remover backup antigo {backup.name}: {e}")
        
        if removidos > 0:
            print(f"🗑️  {removidos} backup(s) antigo(s) removido(s)")

def listar_backups():
    """Lista todos os backups disponíveis"""
    
    backup_dir = Path(__file__).parent / 'backups'
    
    if not backup_dir.exists():
        print("Nenhum backup encontrado.")
        return
    
    backups = sorted(backup_dir.glob('backup_*.db'), key=lambda x: x.stat().st_mtime, reverse=True)
    
    if not backups:
        print("Nenhum backup encontrado.")
        return
    
    print(f"\n📦 Backups disponíveis ({len(backups)}):")
    print("-" * 60)
    
    for i, backup in enumerate(backups, 1):
        tamanho_mb = backup.stat().st_size / (1024 * 1024)
        data_mod = datetime.fromtimestamp(backup.stat().st_mtime)
        print(f"{i:2d}. {backup.name}")
        print(f"    Data: {data_mod.strftime('%d/%m/%Y %H:%M:%S')} | Tamanho: {tamanho_mb:.2f} MB")

def restaurar_backup(backup_file):
    """Restaura um backup específico"""
    
    db_file = Path(__file__).parent / 'instance' / 'ueci_monitoramento.db'
    backup_path = Path(backup_file)
    
    if not backup_path.exists():
        print(f"❌ Arquivo de backup não encontrado: {backup_file}")
        return False
    
    try:
        # Criar backup do estado atual antes de restaurar
        if db_file.exists():
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_antes = db_file.parent.parent / 'backups' / f'backup_antes_restauracao_{timestamp}.db'
            shutil.copy2(db_file, backup_antes)
            print(f"💾 Backup do estado atual criado: {backup_antes.name}")
        
        # Restaurar backup
        shutil.copy2(backup_path, db_file)
        print(f"✅ Backup restaurado com sucesso!")
        print(f"   De: {backup_path.name}")
        print(f"   Para: {db_file}")
        
        return True
    
    except Exception as e:
        print(f"❌ Erro ao restaurar backup: {e}")
        return False

if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        comando = sys.argv[1]
        
        if comando == 'criar':
            criar_backup()
        elif comando == 'listar':
            listar_backups()
        elif comando == 'restaurar' and len(sys.argv) > 2:
            restaurar_backup(sys.argv[2])
        else:
            print("Uso:")
            print("  python backup_automatico.py criar     - Cria um novo backup")
            print("  python backup_automatico.py listar    - Lista todos os backups")
            print("  python backup_automatico.py restaurar <arquivo> - Restaura um backup")
    else:
        # Sem argumentos, criar backup
        criar_backup()
