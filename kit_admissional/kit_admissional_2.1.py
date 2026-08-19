import re
import fitz  # PyMuPDF
import pandas as pd
import unicodedata
from pathlib import Path
from datetime import datetime


# ============================================================
# CONFIGURAÇÕES
# ============================================================

NOME_PDF = "Geral.pdf"
NOME_EXCEL = "Dados.xlsx"
PASTA_SAIDA_PRINCIPAL = "Kits_Gerados"

MARCADOR_INICIO_KIT = "REGISTRO DE EMPREGADO"

# Colunas da planilha (0-based)
COLUNA_NOME = 2   # Coluna C
COLUNA_PASTA = 7  # Coluna H

# Se a planilha tiver cabeçalho na primeira linha, deixe True
PLANILHA_TEM_CABECALHO = True


# ============================================================
# FUNÇÕES DE APOIO
# ============================================================

def normalizar_texto(texto):
    """
    Normaliza texto para comparação:
    - remove acentos
    - converte para minúsculas
    - remove espaços duplicados
    """
    if texto is None:
        return ""

    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    return texto.lower().strip()


def nome_arquivo_seguro(nome):
    """
    Remove caracteres inválidos para nomes de arquivo/pasta no Windows.
    """
    nome = str(nome).strip()
    nome = re.sub(r'[<>:"/\\|?*]', '', nome)
    nome = re.sub(r"\s+", " ", nome)
    return nome.strip()


def gerar_nome_sem_sobrescrever(caminho_arquivo):
    """
    Se o arquivo já existir, cria um novo nome com sufixo _2, _3, etc.
    """
    caminho_arquivo = Path(caminho_arquivo)

    if not caminho_arquivo.exists():
        return caminho_arquivo

    stem = caminho_arquivo.stem
    suffix = caminho_arquivo.suffix
    pasta = caminho_arquivo.parent

    contador = 2
    while True:
        novo = pasta / f"{stem}_{contador}{suffix}"
        if not novo.exists():
            return novo
        contador += 1


def extrair_nome_da_pagina(texto):
    """
    Extrai o nome do colaborador da primeira página do kit.
    
    Trata estes formatos:
    1) Nome: João da Silva
    2) Nome : João da Silva
    3) Nome:
       João da Silva
    4) Nome :
       João da Silva
    """
    if not texto:
        return None

    # Normaliza quebras de linha e espaços
    texto = texto.replace("\r", "\n")
    linhas = [re.sub(r"\s+", " ", linha).strip() for linha in texto.split("\n")]

    # Remove linhas vazias
    linhas = [linha for linha in linhas if linha]

    # --------------------------------------------------------
    # CASO 1 e 2:
    # Nome: Fulano
    # Nome : Fulano
    # --------------------------------------------------------
    for linha in linhas:
        match = re.match(r"(?i)^nome\s*:\s*(.+)$", linha)
        if match:
            nome = match.group(1).strip()

            # Evita pegar "Nome:" vazio
            if nome:
                return nome

    # --------------------------------------------------------
    # CASO 3 e 4:
    # Nome:
    # Fulano
    # --------------------------------------------------------
    for i, linha in enumerate(linhas):
        if re.match(r"(?i)^nome\s*:\s*$", linha):
            if i + 1 < len(linhas):
                proxima = linhas[i + 1].strip()

                # Evita capturar títulos ou campos vazios
                if proxima and len(proxima) > 2:
                    return proxima

    return None


def carregar_base_excel(caminho_excel):
    """
    Carrega a base do Excel e retorna um dicionário normalizado:
    {
        nome_normalizado: {
            "nome_original": "...",
            "pasta": "..."
        }
    }
    """
    header = 0 if PLANILHA_TEM_CABECALHO else None
    df = pd.read_excel(caminho_excel, header=header)

    base = {}
    linhas_invalidas = []

    for idx, row in df.iterrows():
        numero_linha_excel = idx + 2 if PLANILHA_TEM_CABECALHO else idx + 1

        try:
            nome = row.iloc[COLUNA_NOME]
            pasta = row.iloc[COLUNA_PASTA]

            if pd.isna(nome) or pd.isna(pasta):
                continue

            nome = str(nome).strip()
            pasta = str(pasta).strip()

            if not nome or not pasta:
                continue

            nome_norm = normalizar_texto(nome)

            base[nome_norm] = {
                "nome_original": nome,
                "pasta": pasta
            }

        except Exception as e:
            linhas_invalidas.append({
                "linha_excel": numero_linha_excel,
                "erro": str(e)
            })

    return base, linhas_invalidas


def localizar_kits_no_pdf(doc):
    """
    Localiza os intervalos de páginas de cada kit.
    Um kit começa quando encontra MARCADOR_INICIO_KIT.
    """
    inicios = []

    for i in range(len(doc)):
        texto = doc[i].get_text("text")
        if MARCADOR_INICIO_KIT.upper() in texto.upper():
            inicios.append(i)

    kits = []
    for i, inicio in enumerate(inicios):
        if i < len(inicios) - 1:
            fim = inicios[i + 1] - 1
        else:
            fim = len(doc) - 1

        kits.append((inicio, fim))

    return kits


def salvar_kit(doc, pagina_inicial, pagina_final, caminho_saida):
    """
    Salva um novo PDF contendo apenas as páginas do kit.
    """
    novo_pdf = fitz.open()

    for p in range(pagina_inicial, pagina_final + 1):
        novo_pdf.insert_pdf(doc, from_page=p, to_page=p)

    novo_pdf.save(caminho_saida)
    novo_pdf.close()


def escrever_log(caminho_log, mensagem):
    """
    Escreve mensagem no arquivo de log e também imprime na tela.
    """
    print(mensagem)
    with open(caminho_log, "a", encoding="utf-8") as f:
        f.write(mensagem + "\n")


# ============================================================
# PROCESSO PRINCIPAL
# ============================================================

def main():
    downloads = Path.home() / "Downloads"
    caminho_pdf = downloads / NOME_PDF
    caminho_excel = downloads / NOME_EXCEL

    # Data da execução para criar subpasta do dia
    data_execucao = datetime.now().strftime("%Y-%m-%d")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    pasta_saida_principal = downloads / PASTA_SAIDA_PRINCIPAL
    pasta_saida = pasta_saida_principal / data_execucao

    pasta_saida.mkdir(parents=True, exist_ok=True)

    caminho_log = pasta_saida / f"log_processamento_{timestamp}.txt"
    caminho_relatorio = pasta_saida / f"Relatorio_Processamento_{timestamp}.xlsx"

    # --------------------------------------------------------
    # Validação dos arquivos
    # --------------------------------------------------------
    if not caminho_pdf.exists():
        print(f"[ERRO] PDF não encontrado: {caminho_pdf}")
        return

    if not caminho_excel.exists():
        print(f"[ERRO] Excel não encontrado: {caminho_excel}")
        return

    escrever_log(caminho_log, "=" * 80)
    escrever_log(caminho_log, "INÍCIO DO PROCESSAMENTO")
    escrever_log(caminho_log, "=" * 80)
    escrever_log(caminho_log, f"PDF: {caminho_pdf}")
    escrever_log(caminho_log, f"Excel: {caminho_excel}")
    escrever_log(caminho_log, f"Pasta de saída principal: {pasta_saida_principal}")
    escrever_log(caminho_log, f"Pasta de saída da execução: {pasta_saida}")

    # --------------------------------------------------------
    # Carrega Excel
    # --------------------------------------------------------
    escrever_log(caminho_log, "\n[1/4] Carregando base do Excel...")
    base_colaboradores, linhas_invalidas = carregar_base_excel(caminho_excel)

    if not base_colaboradores:
        escrever_log(caminho_log, "[ERRO] Nenhum colaborador válido encontrado na planilha.")
        return

    escrever_log(caminho_log, f"Total de colaboradores carregados: {len(base_colaboradores)}")

    # --------------------------------------------------------
    # Abre PDF e localiza kits
    # --------------------------------------------------------
    escrever_log(caminho_log, "\n[2/4] Abrindo PDF e localizando kits...")
    doc = fitz.open(caminho_pdf)
    kits = localizar_kits_no_pdf(doc)

    if not kits:
        escrever_log(caminho_log, "[ERRO] Nenhum kit encontrado no PDF.")
        doc.close()
        return

    escrever_log(caminho_log, f"Total de kits encontrados no PDF: {len(kits)}")

    # --------------------------------------------------------
    # Processa kits
    # --------------------------------------------------------
    escrever_log(caminho_log, "\n[3/4] Processando kits...")

    processados = []
    nao_encontrados = []
    sem_nome = []
    erros = []

    total_kits = len(kits)

    for idx, (inicio, fim) in enumerate(kits, start=1):
        try:
            escrever_log(
                caminho_log,
                f"\n--- Kit {idx}/{total_kits} | páginas {inicio + 1} a {fim + 1} ---"
            )

            texto_primeira_pagina = doc[inicio].get_text("text")
            nome_pdf = extrair_nome_da_pagina(texto_primeira_pagina)

            if not nome_pdf:
                msg = f"[AVISO] Nome não identificado no kit {idx} (página inicial {inicio + 1})."
                escrever_log(caminho_log, msg)

                sem_nome.append({
                    "kit": idx,
                    "pagina_inicial": inicio + 1,
                    "pagina_final": fim + 1
                })
                continue

            nome_pdf = re.sub(r"\s+", " ", nome_pdf).strip()
            nome_pdf_norm = normalizar_texto(nome_pdf)

            escrever_log(caminho_log, f"Nome identificado no PDF: {nome_pdf}")

            if nome_pdf_norm not in base_colaboradores:
                msg = f"[AVISO] Nome não encontrado na planilha: {nome_pdf}"
                escrever_log(caminho_log, msg)

                nao_encontrados.append({
                    "kit": idx,
                    "nome_pdf": nome_pdf,
                    "pagina_inicial": inicio + 1,
                    "pagina_final": fim + 1
                })
                continue

            dados_colaborador = base_colaboradores[nome_pdf_norm]
            nome_excel = dados_colaborador["nome_original"]
            nome_pasta = dados_colaborador["pasta"]

            # Cria a pasta do colaborador dentro da pasta da execução
            pasta_destino = pasta_saida / nome_arquivo_seguro(nome_pasta)
            pasta_destino.mkdir(parents=True, exist_ok=True)

            # Nome final do arquivo
            nome_pdf_saida = f"Kit Admissional - {nome_arquivo_seguro(nome_excel)}.pdf"
            caminho_pdf_saida = pasta_destino / nome_pdf_saida
            caminho_pdf_saida = gerar_nome_sem_sobrescrever(caminho_pdf_saida)

            salvar_kit(doc, inicio, fim, caminho_pdf_saida)

            escrever_log(caminho_log, f"[OK] Kit salvo em: {caminho_pdf_saida}")

            processados.append({
                "kit": idx,
                "nome_pdf_extraido": nome_pdf,
                "nome_excel": nome_excel,
                "pasta_coluna_h": nome_pasta,
                "pasta_destino": str(pasta_destino),
                "arquivo_salvo": str(caminho_pdf_saida),
                "pagina_inicial": inicio + 1,
                "pagina_final": fim + 1
            })

        except Exception as e:
            msg = f"[ERRO] Falha ao processar kit {idx}: {str(e)}"
            escrever_log(caminho_log, msg)

            erros.append({
                "kit": idx,
                "pagina_inicial": inicio + 1,
                "pagina_final": fim + 1,
                "erro": str(e)
            })

    doc.close()

    # --------------------------------------------------------
    # Relatório final
    # --------------------------------------------------------
    escrever_log(caminho_log, "\n[4/4] Gerando relatório final...")

    resumo = [{
        "data_execucao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "kits_encontrados_pdf": len(kits),
        "kits_processados": len(processados),
        "kits_sem_nome": len(sem_nome),
        "nomes_nao_encontrados": len(nao_encontrados),
        "erros_processamento": len(erros),
        "linhas_invalidas_excel": len(linhas_invalidas)
    }]

    with pd.ExcelWriter(caminho_relatorio, engine="openpyxl") as writer:
        pd.DataFrame(resumo).to_excel(writer, sheet_name="Resumo", index=False)

        if processados:
            pd.DataFrame(processados).to_excel(writer, sheet_name="Processados", index=False)

        if sem_nome:
            pd.DataFrame(sem_nome).to_excel(writer, sheet_name="Sem_Nome", index=False)

        if nao_encontrados:
            pd.DataFrame(nao_encontrados).to_excel(writer, sheet_name="Nao_Encontrados", index=False)

        if erros:
            pd.DataFrame(erros).to_excel(writer, sheet_name="Erros", index=False)

        if linhas_invalidas:
            pd.DataFrame(linhas_invalidas).to_excel(writer, sheet_name="Excel_Invalidas", index=False)

    escrever_log(caminho_log, "\n" + "=" * 80)
    escrever_log(caminho_log, "PROCESSAMENTO FINALIZADO")
    escrever_log(caminho_log, "=" * 80)
    escrever_log(caminho_log, f"Kits encontrados no PDF: {len(kits)}")
    escrever_log(caminho_log, f"Kits processados com sucesso: {len(processados)}")
    escrever_log(caminho_log, f"Kits sem nome identificado: {len(sem_nome)}")
    escrever_log(caminho_log, f"Nomes não encontrados na planilha: {len(nao_encontrados)}")
    escrever_log(caminho_log, f"Erros de processamento: {len(erros)}")
    escrever_log(caminho_log, f"Linhas inválidas no Excel: {len(linhas_invalidas)}")
    escrever_log(caminho_log, f"\nRelatório Excel: {caminho_relatorio}")
    escrever_log(caminho_log, f"Log TXT: {caminho_log}")

    print("\nConcluído.")


if __name__ == "__main__":
    main()