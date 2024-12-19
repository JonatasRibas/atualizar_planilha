import pandas as pd
from datetime import datetime
import schedule
import time

# Função para atualizar a planilha
def atualizar_planilha():
    # Caminho da planilha
    caminho = 'sua_planilha.xlsx'  # Coloque o nome ou o caminho do seu arquivo aqui
    
    try:
        # Carregar a planilha
        df = pd.read_excel(caminho)

        # Adicionar ou atualizar uma coluna com a data de atualização
        df['Última Atualização'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Salvar a planilha atualizada
        df.to_excel(caminho, index=False)
        print("Planilha atualizada com sucesso!")

    except FileNotFoundError:
        print("Erro: A planilha não foi encontrada. Verifique o caminho.")
    except Exception as e:
        print(f"Erro inesperado: {e}")

# Agendar a tarefa
schedule.every().day.at("17:59").do(atualizar_planilha)

# Loop infinito para manter o script rodando
print("Automação iniciada. Aguardando horário programado...")
while True:
    schedule.run_pending()
    time.sleep(1)
