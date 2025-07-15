import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

CREDS_FILE = 'credentials.json' 
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
]

CREDS = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
GSPREAD_CLIENT = gspread.authorize(CREDS)

# --- FUNÇÕES DE LEITURA ---

def ler_de_google_sheet(nome_planilha, nome_aba=None):
    print(f"Lendo da Planilha Google: '{nome_planilha}'...")
    planilha = GSPREAD_CLIENT.open(nome_planilha)
    if nome_aba:
        aba = planilha.worksheet(nome_aba)
    else:
        aba = planilha.sheet1
    
    dados = aba.get_all_records()
    return pd.DataFrame(dados)

# --- FUNÇÃO DE PROCESSAMENTO ---

def processar_dados(df):
    """Aplica a lógica de negócio ao DataFrame."""
    print("\n--- Aplicando lógica de negócio ---")
    # Exemplo: Selecionar apenas linhas onde Status é 'Aprovado' e adicionar coluna
    df_processado = df[df['Status'] == 'Aprovado'].copy()
    df_processado['Observacao'] = 'Processado via automação completa'
    return df_processado


# --- FUNÇÃO DE ESCRITA ---

def escrever_em_google_sheet(df_final, nome_planilha_destino):
    """Escreve os dados de um DataFrame em uma Planilha Google, substituindo o conteúdo."""
    print(f"\nEscrevendo na planilha de destino: '{nome_planilha_destino}'...")
    try:
        planilha_destino = GSPREAD_CLIENT.open(nome_planilha_destino).sheet1
        planilha_destino.clear()
        # Converte o DataFrame para o formato que o gspread aceita (lista de listas)
        dados_para_escrever = [df_final.columns.values.tolist()] + df_final.values.tolist()
        planilha_destino.update('A1', dados_para_escrever, value_input_option='USER_ENTERED')
        print("Dados inseridos com sucesso!")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ERRO: A planilha de destino '{nome_planilha_destino}' não foi encontrada.")
        print("Verifique o nome e se ela foi compartilhada com o email da conta de serviço.")


# --- EXECUÇÃO PRINCIPAL ---

if __name__ == "__main__":
    # Lê de uma Planilha Google
    df_origem = ler_de_google_sheet("Nome da sua Planilha de Origem", "Nome da Aba")
    
    print("\n--- Dados Originais Extraídos ---")
    print(df_origem)
    
    # Aplica a lógica de negócio
    df_processado = processar_dados(df_origem)
    
    print("\n--- Dados Finais para Escrever ---")
    print(df_processado)

    # Escreve o resultado na planilha de destino
    escrever_em_google_sheet(df_processado, "Nome da sua Planilha de Destino")
