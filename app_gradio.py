import gradio as gr
import os
import shutil
from datetime import datetime
import glob
# Importa todas as funções de automação
from automacao_core import (
    PASTA_ENTRADA, PASTA_DOWNLOAD, PASTA_OUTPUT, 
    ler_dados, verificar_necessidade_atualizacao, 
    corrigir_valor_ipca_selenium, concatena_pdf,
    obter_caminho_base 
) 

# Garante que as pastas estejam prontas
os.makedirs(PASTA_ENTRADA, exist_ok=True)
os.makedirs(PASTA_DOWNLOAD, exist_ok=True)
os.makedirs(PASTA_OUTPUT, exist_ok=True)


# --- Funções de Wrapper para a Interface Gradio ---

def limpar_pastas_temp():
    """Limpa as pastas de entrada e download antes de cada execução."""
    # NÃO use shutil.rmtree no BASE_DIR, apenas nas subpastas!
    for pasta in [PASTA_ENTRADA, PASTA_DOWNLOAD]:
        for arquivo in os.listdir(pasta):
            os.remove(os.path.join(pasta, arquivo))

def executar_automacao(arquivo_principal, lista_pdfs_base, mostrar_browser=True, periodo_atualizacao=60):
    """
    Executa a automação baseada no tipo de arquivo principal e usa a lista_pdfs_base 
    para concatenar os resultados.
    """
    
    limpar_pastas_temp()
    yield "Iniciando automação... Limpando pastas temporárias", None
    # 1. Copiar Arquivos para a PASTA_ENTRADA (Ambiente de Trabalho)
    
    # A. Arquivo Principal (Excel ou PDF Cotação)
    caminho_principal = os.path.join(PASTA_ENTRADA, os.path.basename(arquivo_principal))
    shutil.copy(arquivo_principal, caminho_principal)
    
    # B. PDFs Base (Relatórios que serão concatenados)
    efiscos_com_pdf_base = set()
    for pdf_file in lista_pdfs_base:
        # Renomeia para o nome original no Gradio e salva.
        # EXIGÊNCIA: O nome do arquivo DEVE ser o código EFISCO.pdf
        nome_base = os.path.basename(pdf_file)
        caminho_pdf_base = os.path.join(PASTA_ENTRADA, nome_base)
        shutil.copy(pdf_file, caminho_pdf_base)
        efiscos_com_pdf_base.add(nome_base.replace('.pdf', ''))

    yield "Arquivos de entrada copiados. Lendo dados do arquivo principal...", None

    # 2. Ler Dados e Obter Estrutura (Dados a serem corrigidos)
    
    # A função ler_dados agora aceita apenas o caminho do arquivo principal
    dados_a_corrigir = ler_dados(caminho_principal)
    
    if not dados_a_corrigir:
        yield "ERRO: Falha ao ler dados do arquivo principal ou arquivo vazio/inválido.", None
        return None, None

    # 3. Processar e Gerar Atualizações de Preço
    
    itens_a_corrigir, dados_completos = verificar_necessidade_atualizacao(dados_a_corrigir, periodo_atualizacao)

    total_a_atualizar = len(itens_a_corrigir)
    if total_a_atualizar > 0:
        yield f"Encontrados {total_a_atualizar} itens para atualizar. Iniciando correção de IPCA...", None
        
        itens_restantes = total_a_atualizar
        # O índice 'i' deve ser único em todos os dados lidos
        for i, item in enumerate(dados_completos):
            item_id = i + 1
            if item['status'] == 'Atualizar':
                yield f"Atualizando item {item_id}/{len(dados_completos)} (Codigo {item['efisco']}). Restantes: {itens_restantes - 1}.", None
                corrigir_valor_ipca_selenium(item, item_id, mostrar_browser)
                itens_restantes -= 1

    else:
        print("\nNenhum item precisou de atualização (todos < 180 dias).")

    
    # 4. Concatenar Resultados
    
    
    # Se o arquivo principal for PDF, o EFISCO é único. Se for Excel, são múltiplos.
    codigos_para_concatenar = set(item['efisco'] for item in dados_completos)

    
    arquivos_finais_gerados = []
    
    yield f"\nIniciando concatenação de PDFs para {len(codigos_para_concatenar)} códigos...", None    
    for codigo in codigos_para_concatenar:
        if codigo in efiscos_com_pdf_base:
            # Chama a função de concatenação com todos os dados para obter a ordem correta
            concatena_pdf(codigo, dados_completos)
            yield f"Concatenando PDF completo para EFISCO {codigo}...", None
            # Adiciona o caminho do arquivo gerado para o retorno do Gradio
            caminho_saida = os.path.join(PASTA_OUTPUT, f"{codigo}_COMPLETO.pdf")
            if os.path.exists(caminho_saida):
                arquivos_finais_gerados.append(caminho_saida)
        else:
            yield f"AVISO: PDF base '{codigo}.pdf' não fornecido. Concatenação ignorada.", None

    # 5. Retorno Final
    if arquivos_finais_gerados:
        yield f"SUCESSO! {len(arquivos_finais_gerados)} arquivos completos gerados na pasta de saída.", arquivos_finais_gerados
        return True
    else:
        yield "Concluído, mas nenhum arquivo PDF final foi gerado.", None
        return  None


# --- Interface Gradio ---

with gr.Blocks(title="Automação de Correção de IPCA") as demo:
    gr.Markdown("# 🤖 Automação de Correção Monetária (IPCA)")

    with gr.Tab("Principal"):
        gr.Markdown("#### 📁 Entrada de Dados")

        mostrar_browser = gr.Checkbox(label="Mostrar Navegador Durante a Execução", value=False)
        periodo_atualizacao = gr.Number(label="Atualizar a partir de (dias)", value=60, interactive=True)

        # Entrada do Excel
        main_file = gr.File(label="1. Carregar Arquivo Excel (efisco, valor e data) ou PDF (cotação resumida)", file_types=[".xlsx", ".pdf"])
        
        # Entrada dos PDFs (Múltipla Seleção)
        pdf_reports = gr.Files(label="2. Carregar PDFs de Relatório (Nomear como EFISCO.pdf)", file_types=[".pdf"])

        btn_excel_run = gr.Button("🚀 Executar Automação")
        
        # Saída do Modo 1
        output_text = gr.Textbox(label="Status da Execução / Log")
        output_files_text = gr.Files(label="Arquivos PDF Completos Gerados")

        btn_excel_run.click(
            fn=executar_automacao, 
            inputs=[main_file, pdf_reports, mostrar_browser, periodo_atualizacao], 
            outputs=[output_text, output_files_text]
        )


if __name__ == "__main__":
    demo.launch()