import sys
import os
import glob
import base64
import openpyxl
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from dateutil.relativedelta import relativedelta
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# --- 1. Função para garantir que os caminhos funcionem no .EXE ---
def obter_caminho_base():
    """Retorna o diretório onde o executável ou o script está rodando."""
    if getattr(sys, 'frozen', False):
        # Se for um executável (PyInstaller)
        return os.path.dirname(sys.executable)
    # Se for um script Python normal
    return os.path.dirname(os.path.abspath(__file__))

# Define o caminho base
BASE_DIR = obter_caminho_base()

# Configura as pastas RELATIVAS ao executável
PASTA_ENTRADA = os.path.join(BASE_DIR, "Dados de entrada") 
PASTA_DOWNLOAD = os.path.join(BASE_DIR, "downloads_pdf")
PASTA_OUTPUT = os.path.join(BASE_DIR, "output")
os.makedirs(PASTA_DOWNLOAD, exist_ok=True)
os.makedirs(PASTA_OUTPUT, exist_ok=True)


# --- Funções do Script ---

def ler_dados():
    # 1. Tenta encontrar arquivos
    arquivos_excel = glob.glob(os.path.join(PASTA_ENTRADA, "*.xlsx"))
    arquivos_csv = glob.glob(os.path.join(PASTA_ENTRADA, "*.csv"))
    
    caminho_arquivo = None
    if arquivos_excel:
        caminho_arquivo = arquivos_excel[0]
        tipo_arquivo = "xlsx"
        # Mapeamento Padrão para Excel (use os nomes reais das suas colunas)
        mapa_colunas = {'EFISCO': 'efisco', 'VALOR': 'valor', 'DATA': 'data_base'}
    elif arquivos_csv:
        caminho_arquivo = arquivos_csv[0]
        tipo_arquivo = "csv"
        # Mapeamento Flexível para CSV
        mapa_colunas = {'Código do Item': 'efisco', 'Preço Unitário': 'valor', 'Data/Hora da Compra': 'data_base'}
    else:
        print(f"ERRO CRÍTICO: Não encontrei nenhum arquivo .xlsx ou .csv na pasta '{PASTA_ENTRADA}'.")
        return []
    
    nome_arquivo_usado = os.path.basename(caminho_arquivo)
    print(f"Lendo o primeiro arquivo encontrado ({tipo_arquivo}): {nome_arquivo_usado}")

    try:
        # 2. Leitura com Pandas
        if tipo_arquivo == "xlsx":
            df = pd.read_excel(caminho_arquivo)
        else: # csv
            df = pd.read_csv(caminho_arquivo, encoding='latin-1', skiprows= 2, sep=';', usecols=["Código do Item", "Preço Unitário", "Data/Hora da Compra"])
            print(df)
        # 3. Normalização e Mapeamento de Colunas
        df.columns = df.columns.str.upper().str.strip()
        
        # Invertemos o mapa para verificar se a coluna original (key) está no DF
        colunas_para_renomear = {}
        for nome_original, nome_novo in mapa_colunas.items():
            if nome_original in df.columns:
                colunas_para_renomear[nome_original] = nome_novo
            
        df.rename(columns=colunas_para_renomear, inplace=True)
        
        # Filtra apenas as colunas que conseguimos mapear para evitar erros
        colunas_finais = ['efisco', 'valor', 'data_base']
        df = df[colunas_finais]

        # 4. Processamento dos Dados
        lista_itens = []
        for index, row in df.iterrows():
            try:
                efisco_raw = row['efisco']
                valor_raw = row['valor']
                data_raw = row['data_base']
                
                if pd.isna(efisco_raw) or pd.isna(valor_raw):
                    continue

                # Normalização de tipos
                efisco = str(int(efisco_raw)).strip()
                valor = float(valor_raw)

                if isinstance(data_raw, datetime):
                    data_objeto = data_raw.date()
                elif isinstance(data_raw, str):
                    data_objeto = datetime.strptime(data_raw, '%d/%m/%Y').date()
                else:
                    # Se for um Timestamp do Pandas
                    data_objeto = pd.to_datetime(data_raw).date()
                
                lista_itens.append({
                    'efisco': efisco,
                    'valor': valor,
                    'data_base': data_objeto
                })
            except Exception as e:
                print(f"Erro ao processar linha {index + 2}: {e}")
                
        return lista_itens
        
    except Exception as e:
        print(f"Erro ao abrir/processar arquivo: {e}")
        return []

def verificar_necessidade_atualizacao(dados):
    data_hoje = datetime.now().date()
    limite_dias = timedelta(days=180)
    itens_para_atualizar = []
    
    for item in dados:
        diferenca = data_hoje - item['data_base']
        if diferenca > limite_dias:
            item['status'] = 'Atualizar'
            item['dias_atraso'] = diferenca.days
            itens_para_atualizar.append(item)
        else:
            item['status'] = 'OK'
            item['dias_atraso'] = diferenca.days
            
    return itens_para_atualizar, dados

def gerar_pdf_cdp(driver, efisco, data_base, pasta_destino, item_id):
    try:
        params = {
            'landscape': False,
            'displayHeaderFooter': False,
            'printBackground': True,
            'paperWidth': 8.27,
            'paperHeight': 11.69,
            'marginTop': 0.4,
            'marginBottom': 0.4,
            'marginLeft': 0.4,
            'marginRight': 0.4
        }
        
        resultado = driver.execute_cdp_cmd("Page.printToPDF", params)
        
        data_formatada = data_base.strftime('%d%m%Y')
        nome_arquivo = f"EFISCO_{efisco}_item_{item_id}Correcao_IPCA_{data_formatada}.pdf"
        caminho_completo = os.path.join(pasta_destino, nome_arquivo)
        
        with open(caminho_completo, 'wb') as f:
            f.write(base64.b64decode(resultado['data']))
            
        print(f"   -> PDF SALVO: {nome_arquivo}")
        return True
        
    except Exception as e:
        print(f"   -> ERRO ao gerar PDF via CDP: {e}")
        return False

def corrigir_valor_ipca_selenium(item, item_id):
    service = Service(ChromeDriverManager().install())
    
    opcoes = Options()
    
    #opcoes.add_argument("--headless=new") 
    
    driver = webdriver.Chrome(service=service, options=opcoes)
    driver.implicitly_wait(3) 
    url_calculadora = "https://www3.bcb.gov.br/CALCIDADAO/publico/exibirFormCorrecaoValores.do?method=exibirFormCorrecaoValores" 
    
    data_origem_str = item['data_base'].strftime('%m%Y')
    
    valor_a_enviar = f"{item['valor']:.2f}".replace('.', ',')
    
    print(f"Processando EFISCO {item['efisco']}...")

    tentativas = 0
    max_tentativas = 2
    try:
        driver.get(url_calculadora)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'selIndice'))
        )
        while tentativas < max_tentativas:
            Select(driver.find_element(By.ID, 'selIndice')).select_by_value("00433IPCA")
            driver.find_element(By.NAME, 'dataInicial').send_keys(data_origem_str)
            
            data_hoje = datetime.now().date()
            data_final_str = (data_hoje - relativedelta(months=1+tentativas)).strftime('%m%Y')
            driver.find_element(By.NAME, 'dataFinal').send_keys(data_final_str)
            
            campo_valor = driver.find_element(By.NAME, 'valorCorrecao')
            campo_valor.clear()
            campo_valor.send_keys(valor_a_enviar)
            
            btn_corrigir = driver.find_element(By.CSS_SELECTOR, "input[value='Corrigir valor']")
            btn_corrigir.click()

            try:
                WebDriverWait(driver, 3).until( 
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[value='Imprimir']"))
                )
                gerar_pdf_cdp(driver, item['efisco'], item['data_base'], PASTA_DOWNLOAD, item_id)
                break
            except TimeoutException:
                print("   -> ERRO: O carregamento da página de resultados demorou mais de 5 segundos.")
                print("   -> Tentando buscar atualização para o mês anterior.")
                tentativas += 1
                driver.get(url_calculadora) 
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'selIndice'))
                )
        return False


    except Exception as e:
        print(f"   -> Erro Selenium: {e}")
        return False
    finally:
        driver.quit()

from PyPDF2 import PdfWriter

from PyPDF2 import PdfWriter

def concatena_pdf(catmat: str):
    """
    Concatena o relatório original (se existir) com todos os PDFs de preço gerados
    para o EFISCO (catmat) especificado.
    """
    
    # 1. Procura pelo relatório original
    relatorio = [arq for arq in os.listdir(PASTA_ENTRADA) if arq.startswith(catmat) and arq.endswith(".pdf")]
    
    # 2. Procura pelos PDFs gerados 
    arquivos_gerados = [arq for arq in os.listdir(PASTA_DOWNLOAD) if arq.startswith(f"EFISCO_{catmat}") and arq.endswith(".pdf")]
    
    # Falha apenas no caso de ambos não existirem. i.e, gera um output se um dos dois existir.
    if not relatorio and not arquivos_gerados:
        print(f"   -> ATENÇÃO: Nenhuma arquivo PDF encontrado para o EFISCO {catmat} nas pastas de entrada/downloads. Nenhuma concatenação foi feita.")
        return False
        
    merger = PdfWriter()
    
    # Tenta adicionar o relatório original primeiro
    if relatorio:
        caminho_relatorio = os.path.join(PASTA_ENTRADA, relatorio[0])
        print(f"   -> Adicionando Relatório Base: {relatorio[0]}")
        merger.append(caminho_relatorio)
    else:
        print(f"   -> ATENÇÃO: Não foi encontrado um relatório PDF que comece com '{catmat}' na pasta de entrada. Apenas os PDFs gerados serão concatenados.")

    # Adiciona todos os PDFs gerados
    for arquivo in arquivos_gerados:
        caminho_arquivo = os.path.join(PASTA_DOWNLOAD, arquivo)
        merger.append(caminho_arquivo)
    
    # Finaliza a escrita
    caminho_saida = os.path.join(PASTA_OUTPUT, f"{catmat}_COMPLETO.pdf")
    merger.write(caminho_saida)
    merger.close()
    
    print(f"   -> SUCESSO: Arquivo final salvo em: {caminho_saida}")
    return True


# --- Bloco Principal de Execução ---
if __name__ == "__main__":
    print("--- INICIANDO AUTOMACAO IPCA ---")
    print(f"Diretório base: {BASE_DIR}")
    
    dados_completos = ler_dados()
    
    if dados_completos:
        # eu fiz usando a logica de que um codigo ja identifica unicamente
        itens_a_corrigir, dados_completos = verificar_necessidade_atualizacao(dados_completos)
        
        if itens_a_corrigir:
            print(f"\nEncontrados {len(itens_a_corrigir)} itens para atualizar.")
            
            for i, item in enumerate(dados_completos):
                item_id = i+1
                if item['status'] == 'Atualizar':
                    corrigir_valor_ipca_selenium(item, item_id)
            
            print("\n--- PROCESSO FINALIZADO COM SUCESSO ---")
        else:
            print("\nNenhum item precisou de atualização (todos < 180 dias).")
        
        codigos_efisco_unicos = set(item['efisco'] for item in dados_completos)
        
        print(f"\n--- INICIANDO CONCATENAÇÃO DE PDFs para {len(codigos_efisco_unicos)} códigos ---")
        
        # Iterar sobre CADA código EFISCO e chamar a função
        for codigo in codigos_efisco_unicos:
            concatena_pdf(codigo)
    
    else:
        print("\nNenhum dado lido ou arquivo não encontrado.")

    print(f"Verifique a pasta: {PASTA_OUTPUT}")
    # 2. IMPEDE O FECHAMENTO DO CONSOLE
    print("\n")
    input("Pressione ENTER para fechar o programa...")