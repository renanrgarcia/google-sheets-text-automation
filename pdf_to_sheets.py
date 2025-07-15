import gspread
import pandas as pd
import pdfplumber
import io
import re
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# --- CONFIGURAÇÃO INICIAL ---
# Arquivo de credenciais JSON da sua conta de serviço
CREDS_FILE = 'credentials.json' 
# Escopos necessários para as APIs do Sheets e Drive
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'
]

# --- Autenticação ---
CREDS = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
# Cliente para o Gspread (Sheets)
GSPREAD_CLIENT = gspread.authorize(CREDS)
# Cliente para as APIs do Google (Drive)
DRIVE_SERVICE = build('drive', 'v3', credentials=CREDS)


# --- FUNÇÕES DE LEITURA ---

def ler_de_pdf_local(caminho_arquivo_pdf):
    """Lê o texto de um arquivo PDF local e retorna como uma string única."""
    print(f"Lendo do arquivo PDF local: '{caminho_arquivo_pdf}'...")
    texto_completo = ""
    with pdfplumber.open(caminho_arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"
    return texto_completo

def ler_pdf_do_drive(nome_arquivo_pdf):
    """Encontra um PDF no Google Drive pelo nome, lê e retorna seu texto."""
    print(f"Procurando e lendo o PDF '{nome_arquivo_pdf}' no Google Drive...")
    # Procura pelo arquivo no Drive
    query = f"name = '{nome_arquivo_pdf}' and mimeType = 'application/pdf'"
    results = DRIVE_SERVICE.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print("Arquivo PDF não encontrado no Google Drive.")
        return ""
    
    # Pega o ID do primeiro arquivo encontrado
    pdf_id = items[0]['id']
    
    # Baixa o conteúdo do arquivo
    request = DRIVE_SERVICE.files().get_media(fileId=pdf_id)
    fh = io.BytesIO()
    downloader = io.BytesIO(request.execute())
    
    # Lê o conteúdo baixado com pdfplumber
    texto_completo = ""
    with pdfplumber.open(downloader) as pdf:
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"
    return texto_completo


# --- FUNÇÃO DE PROCESSAMENTO ---

def parse_texto_para_dataframe(texto_extraido):
    """
    Esta é uma função CRÍTICA que você precisa adaptar.
    Ela transforma o texto bruto de um PDF em dados estruturados (DataFrame).
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
    # --- ESCOLHA SUA FONTE DE DADOS ---
    
    # Opção 1: Ler de um PDF local
    texto_pdf_local = ler_de_pdf_local("caminho/para/seu/arquivo.pdf")
    df_origem = parse_texto_para_dataframe(texto_pdf_local)

    # Opção 2: Ler de um PDF no Google Drive
    # texto_pdf_drive = ler_pdf_do_drive("nome_do_arquivo_no_drive.pdf")
    # df_origem = parse_texto_para_dataframe(texto_pdf_drive)
    
    print("\n--- Dados Originais Extraídos ---")
    print(df_origem)
    
    # Para dados de texto, o DataFrame já vem "processado" da função de parse
    df_processado = df_origem
    
    print("\n--- Dados Finais para Escrever ---")
    print(df_processado)

    # Escreve o resultado na planilha de destino
    escrever_em_google_sheet(df_processado, "Nome da sua Planilha de Destino")
