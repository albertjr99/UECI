"""
Script para adicionar as tabelas de configuração ao banco de dados
"""
import sys
import os

# Adicionar o diretório ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db
from models.configuracoes import ConfiguracaoSistema, DropdownConfig
from models.planilha import AbaConfig

def criar_tabelas():
    print("=" * 60)
    print("CRIANDO TABELAS DE CONFIGURAÇÃO")
    print("=" * 60)
    
    with app.app_context():
        try:
            # Criar tabelas
            db.create_all()
            print("✓ Tabelas criadas com sucesso!")
            
            # Configurar zoom padrão
            config_zoom = ConfiguracaoSistema.query.filter_by(chave='zoom_padrao').first()
            if not config_zoom:
                config_zoom = ConfiguracaoSistema(chave='zoom_padrao')
                config_zoom.set_valor({'zoom': 80})
                db.session.add(config_zoom)
                print("✓ Configuração de zoom padrão criada (80%)")
            
            # Criar dropdowns padrão para UECI
            abas = AbaConfig.query.all()
            
            for aba in abas:
                print(f"\nProcessando planilha: {aba.aba_name}")
                
                if 'UECI' in aba.aba_name:
                    # Responsável pela Análise
                    if not DropdownConfig.query.filter_by(
                        aba_name=aba.aba_name,
                        campo_nome='Responsável pela Análise'
                    ).first():
                        dropdown = DropdownConfig(
                            aba_name=aba.aba_name,
                            campo_nome='Responsável pela Análise',
                            is_active=True,
                            ordem=1
                        )
                        dropdown.set_opcoes(['Carla Zambi', 'Gabriela Salgado', 'Larissa Janiques'])
                        db.session.add(dropdown)
                        print(f"  ✓ Dropdown 'Responsável pela Análise' criado")
                    
                    # Origem
                    if not DropdownConfig.query.filter_by(
                        aba_name=aba.aba_name,
                        campo_nome='Origem'
                    ).first():
                        dropdown = DropdownConfig(
                            aba_name=aba.aba_name,
                            campo_nome='Origem',
                            is_active=True,
                            ordem=2
                        )
                        dropdown.set_opcoes(['RELUCI', 'PRÓ-GESTÃO', 'OUTRAS DEMANDAS'])
                        db.session.add(dropdown)
                        print(f"  ✓ Dropdown 'Origem' criado")
                    
                    # Tipo de Ação
                    if not DropdownConfig.query.filter_by(
                        aba_name=aba.aba_name,
                        campo_nome='Tipo de Ação'
                    ).first():
                        dropdown = DropdownConfig(
                            aba_name=aba.aba_name,
                            campo_nome='Tipo de Ação',
                            is_active=True,
                            ordem=3
                        )
                        dropdown.set_opcoes(['MELHORIA', 'REGULARIZAÇÃO'])
                        db.session.add(dropdown)
                        print(f"  ✓ Dropdown 'Tipo de Ação' criado")
                    
                    # Unidade Gestora
                    if not DropdownConfig.query.filter_by(
                        aba_name=aba.aba_name,
                        campo_nome='Unidade Gestora'
                    ).first():
                        dropdown = DropdownConfig(
                            aba_name=aba.aba_name,
                            campo_nome='Unidade Gestora',
                            is_active=True,
                            ordem=4
                        )
                        dropdown.set_opcoes([
                            '600201 - IPAJM',
                            '600210 - FF',
                            '600211 - FP',
                            '600212 - FPS'
                        ])
                        db.session.add(dropdown)
                        print(f"  ✓ Dropdown 'Unidade Gestora' criado")
                
                # STATUS DA RECOMENDAÇÃO - para todas as planilhas
                if not DropdownConfig.query.filter_by(
                    aba_name=aba.aba_name,
                    campo_nome='STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?'
                ).first():
                    dropdown = DropdownConfig(
                        aba_name=aba.aba_name,
                        campo_nome='STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?',
                        is_active=True,
                        ordem=5
                    )
                    dropdown.set_opcoes([
                        'Cumprida – Implementada conforme recomendação.',
                        'Cumprida com Ressalvas – Implementada, mas com limitações ou pequenas adequações.',
                        'A cumprir – A partir das próximas demandas.',
                        'Não Cumprida – Ante justificativa.',
                        'Em Andamento – A implementação está em progresso na área.',
                        'Aguardando Aprovação – Medidas foram propostas, mas aguardam validação superior.',
                        'Risco Assumido – A Administração optou por não implementar a recomendação, assumindo os riscos envolvidos.'
                    ])
                    db.session.add(dropdown)
                    print(f"  ✓ Dropdown 'STATUS DA RECOMENDAÇÃO' criado")
                
                # Status para a UECI - para todas as planilhas
                if not DropdownConfig.query.filter_by(
                    aba_name=aba.aba_name,
                    campo_nome='Status para a UECI'
                ).first():
                    dropdown = DropdownConfig(
                        aba_name=aba.aba_name,
                        campo_nome='Status para a UECI',
                        is_active=True,
                        ordem=6
                    )
                    dropdown.set_opcoes([
                        'Implementada',
                        'Atrasada',
                        'Em monitoramento',
                        'Não acatada',
                        'Sem evidências de ação',
                        'Será implementada a partir das próximas ocorrências'
                    ])
                    db.session.add(dropdown)
                    print(f"  ✓ Dropdown 'Status para a UECI' criado")
            
            # Salvar alterações
            db.session.commit()
            print("\n✓ Todas as configurações foram salvas com sucesso!")
            
        except Exception as e:
            print(f"\n✗ Erro ao criar tabelas: {str(e)}")
            db.session.rollback()
            import traceback
            traceback.print_exc()
            return False
    
    print("\n" + "=" * 60)
    print("MIGRAÇÃO CONCLUÍDA!")
    print("=" * 60)
    return True

if __name__ == '__main__':
    criar_tabelas()
