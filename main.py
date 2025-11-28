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
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from dateutil.relativedelta import relativedelta
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
PASTA_EXCEL = os.path.join(BASE_DIR, "excel") 
PASTA_DOWNLOAD = os.path.join(BASE_DIR, "downloads_pdf")
os.makedirs(PASTA_DOWNLOAD, exist_ok=True)

# --- Funções do Script ---

def ler_dados():
    lista_itens = []
    arquivos_excel = glob.glob(os.path.join(PASTA_EXCEL, "*.xlsx"))

    if not arquivos_excel:
        print(f"ERRO CRÍTICO: Não encontrei nenhum arquivo .xlsx na pasta '{PASTA_EXCEL}'.")
        return []
    caminho_arquivo = arquivos_excel[0] 
    nome_arquivo_usado = os.path.basename(caminho_arquivo) 

    print(f"Lendo o primeiro arquivo encontrado: {nome_arquivo_usado}")

    try:
        workbook = openpyxl.load_workbook(caminho_arquivo)
        sheet = workbook.active
        
        for row in sheet.iter_rows(min_row=2):
            try:
                efisco_raw = row[0].value
                valor_raw = row[1].value
                data_raw = row[2].value
                
                if efisco_raw is None or valor_raw is None:
                    continue

                efisco = str(int(efisco_raw)).strip()
                valor_str = str(valor_raw).replace(',', '.')
                valor = float(valor_str)

                if isinstance(data_raw, datetime):
                    data_objeto = data_raw.date()
                else:
                    data_objeto = datetime.strptime(str(data_raw), '%d/%m/%Y').date()
                
                lista_itens.append({
                    'efisco': efisco,
                    'valor': valor,
                    'data_base': data_objeto
                })
            except Exception as e:
                print(f"Erro ao processar linha: {e}")
                
        return lista_itens
    except FileNotFoundError:
        print(f"ERRO CRÍTICO: Não encontrei o arquivo '{nome_arquivo_usado}' na pasta '{PASTA_EXCEL}'.")
        return []
    except Exception as e:
        print(f"Erro ao abrir Excel: {e}")
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

def gerar_pdf_cdp(driver, efisco, data_base, pasta_destino):
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
        
        data_formatada = data_base.strftime('%Y%m%d')
        nome_arquivo = f"EFISCO_{efisco}_Correcao_IPCA_{data_formatada}.pdf"
        caminho_completo = os.path.join(pasta_destino, nome_arquivo)
        
        with open(caminho_completo, 'wb') as f:
            f.write(base64.b64decode(resultado['data']))
            
        print(f"   -> PDF SALVO: {nome_arquivo}")
        return True
        
    except Exception as e:
        print(f"   -> ERRO ao gerar PDF via CDP: {e}")
        return False

def corrigir_valor_ipca_selenium(item):
    service = Service(ChromeDriverManager().install())
    
    opcoes = Options()
    
    opcoes.add_argument("--headless=new") 
    
    driver = webdriver.Chrome(service=service, options=opcoes)
    driver.implicitly_wait(3) 
    url_calculadora = "https://www3.bcb.gov.br/CALCIDADAO/publico/exibirFormCorrecaoValores.do?method=exibirFormCorrecaoValores" 
    
    data_origem_str = item['data_base'].strftime('%m%Y')
    
    valor_a_enviar = f"{item['valor']:.2f}".replace('.', ',')
    
    print(f"Processando EFISCO {item['efisco']}...")

    try:
        driver.get(url_calculadora)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'selIndice'))
        )
        
        Select(driver.find_element(By.ID, 'selIndice')).select_by_value("00433IPCA")
        driver.find_element(By.NAME, 'dataInicial').send_keys(data_origem_str)
        
        data_hoje = datetime.now().date()
        data_final_str = (data_hoje - relativedelta(months=1)).strftime('%m%Y')
        driver.find_element(By.NAME, 'dataFinal').send_keys(data_final_str)
        
        campo_valor = driver.find_element(By.NAME, 'valorCorrecao')
        campo_valor.clear()
        campo_valor.send_keys(valor_a_enviar)
        
        btn_corrigir = driver.find_element(By.CSS_SELECTOR, "input[value='Corrigir valor']")
        btn_corrigir.click()

        try:
            WebDriverWait(driver, 15).until( 
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[value='Imprimir']"))
            )
    
        except TimeoutException:
            print("   -> ERRO: O carregamento da página de resultados demorou mais de 15 segundos.")
            print("   -> Verifique se os dados de entrada estão válidos.")
            return False

        gerar_pdf_cdp(driver, item['efisco'], item['data_base'], PASTA_DOWNLOAD)
        

    except Exception as e:
        print(f"   -> Erro Selenium: {e}")
        return False
    finally:
        driver.quit()

# --- Bloco Principal de Execução ---
if __name__ == "__main__":
    print("--- INICIANDO AUTOMACAO IPCA ---")
    print(f"Diretório base: {BASE_DIR}")
    
    dados_completos = ler_dados()
    
    if dados_completos:
        itens_a_corrigir, dados_completos = verificar_necessidade_atualizacao(dados_completos)
        
        if itens_a_corrigir:
            print(f"\nEncontrados {len(itens_a_corrigir)} itens para atualizar.")
            
            for item in dados_completos:
                if item['status'] == 'Atualizar':
                    corrigir_valor_ipca_selenium(item)
            
            print("\n--- PROCESSO FINALIZADO COM SUCESSO ---")
            print(f"Verifique a pasta: {PASTA_DOWNLOAD}")
        else:
            print("\nNenhum item precisou de atualização (todos < 180 dias).")
    else:
        print("\nNenhum dado lido ou arquivo não encontrado.")

    # 2. IMPEDE O FECHAMENTO DO CONSOLE
    print("\n")
    input("Pressione ENTER para fechar o programa...")