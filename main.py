import sys
import os
import glob
import base64
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from dateutil.relativedelta import relativedelta
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import camelot
import numpy as np
from PyPDF2 import PdfWriter


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
    arquivos_pdf = glob.glob(os.path.join(PASTA_ENTRADA, "*.pdf"))
    
    caminho_arquivo = None
    if arquivos_excel:
        caminho_arquivo = arquivos_excel[0]
        tipo_arquivo = "xlsx"
        # Mapeamento Padrão para Excel
        mapa_colunas = {'EFISCO': 'efisco', 'VALOR': 'valor', 'DATA': 'data_base'}
    elif arquivos_csv:
        caminho_arquivo = arquivos_csv[0]
        tipo_arquivo = "csv"
        # Mapeamento Flexível para CSV
        mapa_colunas = {'Código do Item': 'efisco', 'Preço Unitário': 'valor', 'Data/Hora da Compra': 'data_base'}
    elif arquivos_pdf:
        caminho_arquivo = arquivos_pdf[0]
        nome_arquivo_usado = os.path.basename(caminho_arquivo)
        try:
            tabelas = camelot.io.read_pdf(
                caminho_arquivo, 
                pages='all', 
                flavor='stream', # 'stream' é bom para tabelas com poucas linhas visíveis (como a sua)
                # O parâmetro column_names pode ser necessário para forçar a identificação correta das colunas
            )

            if not tabelas:
                raise ValueError("Nenhuma tabela encontrada no PDF. Verifique o formato do arquivo.")
            
            print(f"Encontradas {tabelas.n} tabelas. Combinando dados...")
            # 2. Combinar todas as tabelas em um único DataFrame
            # No seu relatório, a tabela que lista os preços está dividida em várias páginas.
            df_lista = [t.df for t in tabelas if t.df.shape[1] > 5] # Filtra apenas tabelas com mais de 5 colunas (a principal)


            df_lista[0] = df_lista[0].drop(index=[0,1]) 
            # Limpa a sujeira do ultimo df
            df_ultimo = df_lista[-1]
            for index, row in df_ultimo.iterrows():
                if row.astype(str).str.contains('Legenda:', case=False, na=False).any():
                    linha_legenda = index
                    break
            else:
                linha_legenda = None # Nenhuma Legenda encontrada

            # 3. Limpar o DataFrame
            if linha_legenda is not None:
                df_lista[-1] = df_ultimo.iloc[:linha_legenda] # Mantém apenas as linhas ANTES da Legenda
            
            df_lista_final = []
            # corrige as colunas das tabelas
            for i, tabela in enumerate(df_lista):
                df_temp = tabela.copy() # Obtém o DataFrame da tabela atual

                if df_temp.shape[1] == 8:
                    # Tabela da Página 1 (ou se o Camelot acertou) - Assumimos que está correta

                    # Renomeamos as colunas para facilitar a concatenação e limpeza posterior
                    df_temp.columns = ['N°', 'Inciso', 'Nome', 'Quantidade', 'Unidade', 'Preço unitário', 'Data', 'Compõe']
                    
                    df_temp[['Quantidade', 'Unidade']] = df_temp['Quantidade'].str.extract(r'(\d+)\s*(.*)') # inutil, tava viajando demais
                    df_temp = df_temp[['Preço unitário', 'Data']]
                    
                    for coluna in df_temp.columns:
                        df_temp[coluna] = df_temp[coluna].replace(r'^\s*$', np.nan, regex=True)
                    
                    df_temp.dropna(inplace=True)
                    df_temp["efisco"] = nome_arquivo_usado.strip('.pdf')
                    df_lista_final.append(df_temp)
                    
                elif df_temp.shape[1] == 7:
                    # Tabela da Página 2 em diante (onde Quantidade e Unidade se juntaram)
                    print(f"   -> Reestruturando Tabela {i} (Colunas agrupadas)...")
                    
                    # 1. Definir os nomes das colunas atuais (7 colunas)
                    # O agrupamento provavelmente ocorreu na 4ª coluna (índice 3)
                    df_temp.columns = ['N°', 'Inciso', 'Nome', 'Quant_Unid_Agrupada', 'Preço unitário', 'Data', 'Compõe']
                    
                    # 2. Aplicar o SPLIT e extrair os dados
                    # Usamos uma expressão regular para capturar o número da quantidade (o primeiro valor)
                    # e o restante (a unidade)
                    
                    # Cria duas novas colunas: 'Quantidade' (o número) e 'Unidade' (o restante)
                    df_temp[['Quantidade', 'Unidade']] = df_temp['Quant_Unid_Agrupada'].str.extract(r'(\d+)\s*(.*)') # inutil, tava viajando demais
                    
                    # 3. Reordenar e Renomear para o formato de 8 colunas (igual ao da Tabela 1)
                    
                    # Remove a coluna agrupada original
                    df_temp.drop(columns=['Quant_Unid_Agrupada'], inplace=True)
                    
                    # A coluna 'Unidade' foi inserida no final. Vamos reordenar para o formato padrão:
                    df_temp = df_temp[['Preço unitário', 'Data']]
                    
                    for coluna in df_temp.columns:
                        df_temp[coluna] = df_temp[coluna].replace(r'^\s*$', np.nan, regex=True)
                    
                    df_temp.dropna(inplace=True)
                    df_temp["efisco"] = nome_arquivo_usado.strip('.pdf')
                    # Adiciona o DataFrame reestruturado
                    df_lista_final.append(df_temp)
                    
                else:
                    print(f"   -> Ignorando Tabela {i} ({df_temp.shape[1]} colunas), estrutura não mapeada.")

            if not df_lista:
                 raise ValueError("Nenhuma tabela de cotação principal encontrada.")

        
            df = pd.concat(df_lista_final, ignore_index=True)
            
            # --- FASE 3: Mapeamento e Processamento ---
        
            df.rename(columns={
                'Preço unitário': 'valor_raw', # Coluna 6
                'Data': 'data_base_raw'      # Coluna 7
            }, inplace=True)
            
            
            # 4. Processamento de Dados (Convertendo e Limpando)
            lista_itens = []
            codigo_item = nome_arquivo_usado.strip('.pdf')
            # Limpeza de strings e conversão para float/date
            for index, row in df.iterrows():
                try:
                    # Tenta extrair a data e valor
                    valor_str = str(row['valor_raw']).replace('R$', '').replace('.', '').replace(',', '.').strip()
                    valor = float(valor_str)
                    
                    # Converte a data para objeto date
                    data_objeto = datetime.strptime(str(row['data_base_raw']).strip(), '%d/%m/%Y').date()
                    
                    lista_itens.append({
                        'efisco': codigo_item,
                        'valor': valor,
                        'data_base': data_objeto
                    })
                except Exception as e:
                    pass 
                    
            return lista_itens
            
        except ValueError as ve:
            print(f"Erro na extração de PDF (Valor): {ve}")
            return []
        except Exception as e:
            print(f"Erro geral ao processar PDF: {e}")
            return []
    
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
    limite_dias = timedelta(days=60) #mudei pra testar usando o pdf
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
            
    print(itens_para_atualizar, dados)
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

def corrigir_valor_ipca_selenium(item, item_id, mostrar_browser=True):
    service = Service(ChromeDriverManager().install())
    
    opcoes = Options()
    
    if not mostrar_browser:
        opcoes.add_argument("--headless=new") 
    
    driver = webdriver.Chrome(service=service, options=opcoes)
    driver.implicitly_wait(3) 
    url_calculadora = "https://www3.bcb.gov.br/CALCIDADAO/publico/exibirFormCorrecaoValores.do?method=exibirFormCorrecaoValores" 
    
    data_origem_str = item['data_base'].strftime('%m%Y')
    
    valor_a_enviar = f"{item['valor']:.2f}".replace('.', ',')
    
    print(f"Processando codigo {item['efisco']}...")

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
            
            # checa se o mês para o qual está tentando atualizar é o mesmo de referencia
            data_final_str_mes = datetime.strptime(data_final_str,'%m%Y').month
            data_origem_str_mes = item['data_base'].month
            if data_final_str_mes == data_origem_str_mes:
                print(f"   -> AVISO: A data final do codigo {item["efisco"]} atingiu o mesmo mês da data base. Não é possível atualizar.")
                break
            
            campo_data = driver.find_element(By.NAME, 'dataFinal')
            campo_data.clear()
            campo_data.send_keys(data_final_str)

            campo_valor = driver.find_element(By.NAME, 'valorCorrecao')
            campo_valor.clear()
            campo_valor.send_keys(valor_a_enviar)
            
            btn_corrigir = driver.find_element(By.CSS_SELECTOR, "input[value='Corrigir valor']")
            btn_corrigir.click()

            try:
                elementos_erro = driver.find_elements(By.CLASS_NAME, "msgErro")
                if elementos_erro:
                    tentativas += 1
                    print(f"   -> ERRO: {elementos_erro[0].text} para data final {data_final_str}.")
                    print(f"\n Data alterada automaticamente para {data_hoje - relativedelta(months=1+tentativas)}.")
                    continue
                

                WebDriverWait(driver, 3).until( 
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[value='Imprimir']"))
                )
                gerar_pdf_cdp(driver, item['efisco'], item['data_base'], PASTA_DOWNLOAD, item_id)
                break
            except TimeoutException:
                print("   -> ERRO: O carregamento da página de resultados demorou mais de 3 segundos.")
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

def concatena_pdf(catmat: str, todos_dados: list): 
    """
    Concatena o relatório original (se existir) com todos os PDFs de preço gerados
    para o EFISCO (catmat) especificado, na ordem do Excel.
    """
    
    # 1. Filtra a ordem dos item_id (1, 2, 3...) do Excel para este EFISCO
    # Pega apenas os índices (item_id) dos itens que pertencem a este EFISCO
    ordem_item_ids = [
        i + 1 for i, item in enumerate(todos_dados) 
        if item['efisco'] == catmat and item['status'] == 'Atualizar'
    ]
    
    # 2. Constrói a lista de caminhos na ORDEM CORRETA
    arquivos_ordenados_caminho = []

    # Itera pelos item_id na ordem do Excel (e, portanto, da lista todos_dados)
    for item_id in ordem_item_ids:
        
        padrao_busca = os.path.join(PASTA_DOWNLOAD, f"EFISCO_{catmat}_item_{item_id}Correcao_IPCA_*.pdf")
        
        arquivos_encontrados = glob.glob(padrao_busca)
        
        if arquivos_encontrados:
            # Adiciona o primeiro arquivo encontrado para aquele item_id
            arquivos_ordenados_caminho.append(arquivos_encontrados[0])
        else:
            print(f"   -> AVISO: PDF de correção para codigo {catmat} (Item {item_id}) não encontrado.")


    # 3. Procura pelo relatório original (não precisa de ordem, é só o primeiro)
    relatorio = [arq for arq in os.listdir(PASTA_ENTRADA) if arq.startswith(catmat) and arq.endswith(".pdf")]
    
    if not relatorio and not arquivos_ordenados_caminho:
        print(f"   -> ATENÇÃO: Nenhuma arquivo PDF encontrado para o codigo {catmat} nas pastas de entrada/downloads. Nenhuma concatenação foi feita.")
        return False
        
    merger = PdfWriter()
    
    # Tenta adicionar o relatório original primeiro
    if relatorio:
        caminho_relatorio = os.path.join(PASTA_ENTRADA, relatorio[0])
        print(f"   -> Adicionando Relatório Base: {relatorio[0]}")
        merger.append(caminho_relatorio)
    else:
        print(f"   -> ATENÇÃO: Não foi encontrado um relatório PDF que comece com '{catmat}' na pasta de entrada.")

    # Adiciona TODOS os PDFs gerados na ordem estabelecida pelo loop
    for caminho_arquivo in arquivos_ordenados_caminho:
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
            concatena_pdf(codigo, dados_completos)
    
    else:
        print("\nNenhum dado lido ou arquivo não encontrado.")

    print(f"Verifique a pasta: {PASTA_OUTPUT}")
    # 2. IMPEDE O FECHAMENTO DO CONSOLE
    print("\n")
    input("Pressione ENTER para fechar o programa...")