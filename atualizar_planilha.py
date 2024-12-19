import pandas as pd
from datetime import datetime
import schedule
import time

# Função para atualizar a planilha
def atualizar_planilha():
    caminho = 'sua_planilha.xlsx'  # Caminho da planilha
    
    try:
        # Carregar a planilha
        df = pd.read_excel(caminho)

        # Adicionar/atualizar a coluna de data de atualização
        df['Última Atualização'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

         # Verificar se a coluna 'Contador de Atualizações' existe
        if 'Contador de Atualizações' not in df.columns:
            df['Contador de Atualizações'] = 0  # Criar a coluna caso não exista

        # Incrementar o contador
        df['Contador de Atualizações'] += 1

        # Exemplo: Adicionar uma nova linha (se necessário)
        nova_linha = {'Nome': 'Maria', 'Idade': 80, 'Cidade': 'Maringá'}
        df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)

        # Salvar a planilha atualizada
        df.to_excel(caminho, index=False)
        print("Planilha atualizada com sucesso!")

    except FileNotFoundError:
        print("Erro: A planilha não foi encontrada. Verifique o caminho.")
    except Exception as e:
        print(f"Erro inesperado: {e}")

# Agendar a tarefa para 17:59
schedule.every().day.at("18:24").do(atualizar_planilha)

# Loop infinito para manter o script rodando
print("Automação iniciada. Aguardando horário programado...")
while True:
    schedule.run_pending()
    time.sleep(1)
