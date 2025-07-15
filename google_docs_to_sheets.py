import gspread
import pandas as pd
import re
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# --- CONFIGURAÇÃO INICIAL ---
# Arquivo de credenciais JSON da sua conta de serviço
CREDS_FILE = 'credentials.json' 
# Escopos necessários para as APIs do Sheets e Docs
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents.readonly',
]

# --- Autenticação ---
CREDS = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
# Cliente para o Gspread (Sheets)
GSPREAD_CLIENT = gspread.authorize(CREDS)
# Cliente para as APIs do Google (Docs)
SERVICE = build('docs', 'v1', credentials=CREDS)


# --- FUNÇÕES DE LEITURA ---

def ler_de_google_doc(document_id):
    """Lê o texto de um Google Doc e retorna como uma string única."""
    print(f"Lendo do Google Doc ID: '{document_id}'...")
    doc = SERVICE.documents().get(documentId=document_id).execute()
    # Extrai o texto de cada parágrafo no corpo do documento
    conteudo_doc = doc.get('body').get('content')
    texto_completo = ""
    for valor in conteudo_doc:
        if 'paragraph' in valor:
            elements = valor.get('paragraph').get('elements')
            for elem in elements:
                texto_completo += elem.get('textRun', {}).get('content', '')
    return texto_completo


# --- FUNÇÃO DE PROCESSAMENTO ---

def parse_texto_para_dataframe(texto_extraido):
    """
    Esta é uma função CRÍTICA que você precisa adaptar.
    Ela transforma o texto bruto de um Google Doc em dados estruturados (DataFrame).
    """
    print("\n--- Convertendo texto bruto para dados estruturados ---")
    # Exemplo de lógica: encontrar "Nome: [algum nome]" e "Valor: [algum valor]"
    # Esta parte é altamente dependente do formato do seu texto.
    # Use expressões regulares (regex) para extrair os dados.
    nomes = re.findall(r"Nome do Cliente: (.*)", texto_extraido)
    valores = re.findall(r"Valor do Pedido: R\$ ([\d,.]+)", texto_extraido)
    status = re.findall(r"Status: (.*)", texto_extraido)

    # Cria um dicionário com os dados extraídos
    dados_estruturados = {
        "Cliente": nomes,
        "Valor": valores,
        "Status": status
    }
    return pd.DataFrame(dados_estruturados)


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
    # Lê de um Google Doc
    # O ID do documento está na URL. Ex: docs.google.com/document/d/DOC_ID_AQUI/edit
    texto_doc = ler_de_google_doc("SEU_ID_DO_GOOGLE_DOC_AQUI")
    df_origem = parse_texto_para_dataframe(texto_doc)
    
    print("\n--- Dados Originais Extraídos ---")
    print(df_origem)
    
    # Para dados de texto, o DataFrame já vem "processado" da função de parse
    df_processado = df_origem
    
    print("\n--- Dados Finais para Escrever ---")
    print(df_processado)

    # Escreve o resultado na planilha de destino
    escrever_em_google_sheet(df_processado, "Nome da sua Planilha de Destino")
