"""
Agendador de Backups Automáticos
Roda em background e cria backups em intervalos regulares
"""
import time
import schedule
from backup_automatico import criar_backup
from datetime import datetime

def job_backup():
    """Tarefa de backup agendada"""
    print(f"\n{'='*60}")
    print(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - Iniciando backup automático...")
    print(f"{'='*60}")
    criar_backup()
    print(f"{'='*60}\n")

def iniciar_agendador(intervalo_horas=6):
    """
    Inicia o agendador de backups
    
    Args:
        intervalo_horas: Intervalo em horas entre backups (padrão: 6 horas)
    """
    
    print(f"\n🤖 Agendador de Backups Iniciado")
    print(f"📅 Backups serão criados a cada {intervalo_horas} hora(s)")
    print(f"💾 Backups mantidos: últimos 30")
    print(f"⏰ Próximo backup: {datetime.now().strftime('%d/%m/%Y')} às {datetime.now().hour + intervalo_horas}:00")
    print("\nPressione Ctrl+C para parar\n")
    
    # Criar backup inicial
    job_backup()
    
    # Agendar backups
    schedule.every(intervalo_horas).hours.do(job_backup)
    
    # Loop principal
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Verificar a cada minuto
    except KeyboardInterrupt:
        print("\n\n⏹️  Agendador de backups interrompido pelo usuário.")

if __name__ == '__main__':
    import sys
    
    # Permitir personalizar intervalo via linha de comando
    intervalo = 6  # padrão: 6 horas
    
    if len(sys.argv) > 1:
        try:
            intervalo = int(sys.argv[1])
        except ValueError:
            print("⚠️  Intervalo inválido. Usando padrão de 6 horas.")
    
    iniciar_agendador(intervalo)
