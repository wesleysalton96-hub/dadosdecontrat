import os
import re
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path


# ============================================================
# CONFIGURAÇÕES
# ============================================================

NOME_PDF = "Geral.pdf"
NOME_EXCEL = "Dados.xlsx"
PASTA_SAIDA = "Kits_Gerados"

# Texto que identifica o início de um novo kit
MARCADOR_INICIO_KIT = "REGISTRO DE EMPREGADO"

# Colunas da planilha
COLUNA_NOME = 2   # Coluna C (índice 0-based)
COLUNA_PASTA = 7  # Coluna H (índice 0-based)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def extrair_nome_da_pagina(texto):
    """
    Extrai o nome do colaborador a partir da linha:
    Nome: Fulano de Tal
    """
    # Procura especificamente a linha iniciada por "Nome:"
    padrao = r"(?im)^\s*Nome:\s*(.+?)\s*$"
    match = re.search(padrao, texto)

    if match:
        nome = match.group(1).strip()
        # Remove espaços duplicados
        nome = re.sub(r"\s+", " ", nome)
        return nome

    return None


def carregar_base_excel(caminho_excel):
    """
    Lê a planilha Dados.xlsx e cria um dicionário:
    {
        "NOME COLABORADOR": "PASTA_DESTINO"
    }
    usando a coluna C para nome e H para pasta.
    """
    df = pd.read_excel(caminho_excel, header=0)

    base = {}

    for idx, row in df.iterrows():
        try:
            nome = row.iloc[COLUNA_NOME]
            pasta = row.iloc[COLUNA_PASTA]

            if pd.isna(nome) or pd.isna(pasta):
                continue

            nome = str(nome).strip()
            pasta = str(pasta).strip()

            if nome:
                base[nome] = pasta

        except Exception as e:
            print(f"[AVISO] Erro ao ler linha {idx + 2} da planilha: {e}")

    return base


def localizar_kits_no_pdf(doc):
    """
    Percorre o PDF e identifica os intervalos de páginas de cada kit.
    Um novo kit começa sempre que encontra 'REGISTRO DE EMPREGADO'.
    
    Retorna lista no formato:
    [
        (pagina_inicial, pagina_final),
        ...
    ]
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
    Salva um novo PDF com as páginas do kit.
    """
    novo_pdf = fitz.open()

    for p in range(pagina_inicial, pagina_final + 1):
        novo_pdf.insert_pdf(doc, from_page=p, to_page=p)

    novo_pdf.save(caminho_saida)
    novo_pdf.close()


def nome_arquivo_seguro(nome):
    """
    Remove caracteres inválidos para nome de arquivo no Windows.
    """
    nome = re.sub(r'[<>:"/\\|?*]', '', nome)
    nome = nome.strip()
    return nome


# ============================================================
# PROCESSO PRINCIPAL
# ============================================================

def main():
    # Descobre a pasta Downloads do usuário
    downloads = Path.home() / "Downloads"

    caminho_pdf = downloads / NOME_PDF
    caminho_excel = downloads / NOME_EXCEL
    pasta_saida = downloads / PASTA_SAIDA

    # Valida arquivos
    if not caminho_pdf.exists():
        print(f"[ERRO] Arquivo PDF não encontrado: {caminho_pdf}")
        return

    if not caminho_excel.exists():
        print(f"[ERRO] Arquivo Excel não encontrado: {caminho_excel}")
        return

    # Cria pasta principal de saída
    pasta_saida.mkdir(parents=True, exist_ok=True)

    print("Carregando planilha...")
    base_colaboradores = carregar_base_excel(caminho_excel)

    if not base_colaboradores:
        print("[ERRO] Nenhum colaborador válido foi encontrado na planilha.")
        return

    print(f"Total de colaboradores carregados da planilha: {len(base_colaboradores)}")

    print("Abrindo PDF...")
    doc = fitz.open(caminho_pdf)

    print("Localizando kits no PDF...")
    kits = localizar_kits_no_pdf(doc)

    if not kits:
        print("[ERRO] Nenhum kit foi encontrado no PDF.")
        doc.close()
        return

    print(f"Total de kits encontrados no PDF: {len(kits)}")

    processados = 0
    nao_encontrados = []
    sem_nome = []
    erros = []

    for idx, (inicio, fim) in enumerate(kits, start=1):
        try:
            texto_primeira_pagina = doc[inicio].get_text("text")
            nome_colaborador = extrair_nome_da_pagina(texto_primeira_pagina)

            if not nome_colaborador:
                sem_nome.append({
                    "kit": idx,
                    "pagina_inicial": inicio + 1
                })
                print(f"[AVISO] Kit {idx}: nome não encontrado na página {inicio + 1}.")
                continue

            print(f"\nKit {idx}: {nome_colaborador}")

            if nome_colaborador not in base_colaboradores:
                nao_encontrados.append({
                    "kit": idx,
                    "nome_pdf": nome_colaborador,
                    "pagina_inicial": inicio + 1
                })
                print(f"[AVISO] Nome não encontrado na planilha: {nome_colaborador}")
                continue

            nome_pasta = base_colaboradores[nome_colaborador]

            # Cria pasta do colaborador dentro de Kits_Gerados
            pasta_colaborador = pasta_saida / nome_arquivo_seguro(nome_pasta)
            pasta_colaborador.mkdir(parents=True, exist_ok=True)

            # Define nome do arquivo PDF
            nome_pdf_saida = f"{nome_arquivo_seguro(nome_colaborador)}.pdf"
            caminho_pdf_saida = pasta_colaborador / nome_pdf_saida

            # Salva kit
            salvar_kit(doc, inicio, fim, caminho_pdf_saida)

            processados += 1
            print(f"[OK] Kit salvo em: {caminho_pdf_saida}")

        except Exception as e:
            erros.append({
                "kit": idx,
                "pagina_inicial": inicio + 1,
                "erro": str(e)
            })
            print(f"[ERRO] Falha ao processar kit {idx}: {e}")

    doc.close()

    # ========================================================
    # RELATÓRIO FINAL
    # ========================================================
    print("\n" + "=" * 60)
    print("RELATÓRIO FINAL")
    print("=" * 60)
    print(f"Kits encontrados no PDF: {len(kits)}")
    print(f"Kits processados com sucesso: {processados}")
    print(f"Kits sem nome identificado: {len(sem_nome)}")
    print(f"Nomes não encontrados na planilha: {len(nao_encontrados)}")
    print(f"Erros de processamento: {len(erros)}")

    # Salvar relatório em Excel
    relatorio_path = pasta_saida / "Relatorio_Processamento.xlsx"

    with pd.ExcelWriter(relatorio_path, engine="openpyxl") as writer:
        pd.DataFrame(
            [{"kits_encontrados": len(kits),
              "kits_processados": processados,
              "kits_sem_nome": len(sem_nome),
              "nomes_nao_encontrados": len(nao_encontrados),
              "erros": len(erros)}]
        ).to_excel(writer, sheet_name="Resumo", index=False)

        if sem_nome:
            pd.DataFrame(sem_nome).to_excel(writer, sheet_name="Sem_Nome", index=False)

        if nao_encontrados:
            pd.DataFrame(nao_encontrados).to_excel(writer, sheet_name="Nao_Encontrados", index=False)

        if erros:
            pd.DataFrame(erros).to_excel(writer, sheet_name="Erros", index=False)

    print(f"\nRelatório salvo em: {relatorio_path}")
    print("\nProcesso finalizado.")


if __name__ == "__main__":
    main()