"""Script para atualizar status e setores dos registros UECI"""

from app import app, db
from models.planilha import PlanilhaData

def atualizar_registros():
    """Atualiza os status e setores dos 18 registros"""
    
    with app.app_context():
        print("Atualizando registros da UECI...\n")
        
        # Buscar todos os registros da UECI
        registros = PlanilhaData.query.filter_by(aba_name='Plano de Ação - UECI').order_by(PlanilhaData.row_order).all()
        
        if len(registros) != 18:
            print(f"⚠️ Esperado 18 registros, encontrado {len(registros)}")
            return
        
        # Dados para atualizar (baseado nas imagens)
        atualizacoes = [
            # Registro 1
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 2
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 3
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 4
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'SCP'
            },
            # Registro 5
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'SCP'
            },
            # Registro 6
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'SCP'
            },
            # Registro 7
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'SCP'
            },
            # Registro 8
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Cumprida – Implementada conforme recomendação.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 9
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 10
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 11
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 12
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 13
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Não Cumprida – Ante justificativa.',
                'Setor(es) responsável(is)': 'GCO'
            },
            # Registro 14
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'Em Andamento – A implementação está em progresso na área.',
                'Setor(es) responsável(is)': 'GFI'
            },
            # Registro 15
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GFI'
            },
            # Registro 16
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GFI'
            },
            # Registro 17
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GCO',
                'Servidor (es) responsável': 'MARIA DE FÁTIMA AGNEZ DE OLIVEIRA'
            },
            # Registro 18
            {
                'STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?': 'A cumprir – A partir das próximas demandas.',
                'Setor(es) responsável(is)': 'GCO',
                'Servidor (es) responsável': 'MARIA DE FÁTIMA AGNEZ DE OLIVEIRA'
            }
        ]
        
        # Atualizar cada registro
        for i, registro in enumerate(registros):
            dados = registro.get_data()
            
            # Aplicar atualizações
            for campo, valor in atualizacoes[i].items():
                dados[campo] = valor
            
            registro.set_data(dados)
            
            status = atualizacoes[i].get('STATUS DA RECOMENDAÇÃO - Qual é a situação atual da recomendação?', '')
            setor = atualizacoes[i].get('Setor(es) responsável(is)', '')
            servidor = atualizacoes[i].get('Servidor (es) responsável', '')
            
            print(f"Registro {i+1:2d}: Setor={setor:3s}, Status={status[:30]}{'...' if len(status) > 30 else ''}")
            if servidor:
                print(f"            Servidor={servidor}")
        
        db.session.commit()
        print(f"\n✅ 18 registros atualizados com sucesso!")

if __name__ == '__main__':
    atualizar_registros()
