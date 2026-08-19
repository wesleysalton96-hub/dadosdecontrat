import re
import fitz  # PyMuPDF
import pandas as pd
import unicodedata
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ============================================================
# CONFIGURAÇÕES
# ============================================================

NOME_SAIDA_PRINCIPAL = "Kits_Gerados"

# Cabeçalhos sugeridos para seleção automática
POSSIVEIS_CABECALHOS_NOME = [
    "nome",
    "nome do colaborador",
    "colaborador",
    "funcionario",
    "funcionário",
    "nome colaborador",
    "nome completo"
]

POSSIVEIS_CABECALHOS_PASTA = [
    "pasta",
    "pasta destino",
    "nome da pasta",
    "destino",
    "pasta do kit",
    "arquivo destino"
]


# ============================================================
# FUNÇÕES DE APOIO
# ============================================================

def normalizar_texto(texto):
    if texto is None:
        return ""

    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    return texto.lower().strip()


def nome_arquivo_seguro(nome):
    nome = str(nome).strip()
    nome = re.sub(r'[<>:"/\\|?*]', '', nome)
    nome = re.sub(r"\s+", " ", nome)
    return nome.strip()


def gerar_nome_sem_sobrescrever(caminho_arquivo):
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

    Trata:
    1) Nome: João da Silva
    2) Nome : João da Silva
    3) Nome:
       João da Silva
    4) Nome :
       João da Silva
    """
    if not texto:
        return None

    texto = texto.replace("\r", "\n")
    linhas = [re.sub(r"\s+", " ", linha).strip() for linha in texto.split("\n")]
    linhas = [linha for linha in linhas if linha]

    # Caso 1 e 2
    for linha in linhas:
        match = re.match(r"(?i)^nome\s*:\s*(.+)$", linha)
        if match:
            nome = match.group(1).strip()
            if nome:
                return nome

    # Caso 3 e 4
    for i, linha in enumerate(linhas):
        if re.match(r"(?i)^nome\s*:\s*$", linha):
            if i + 1 < len(linhas):
                proxima = linhas[i + 1].strip()
                if proxima and len(proxima) > 2:
                    return proxima

    return None


def pagina_e_ficha_registro(texto):
    """
    Identifica a página inicial do kit.
    """
    if not texto:
        return False

    texto_upper = texto.upper()

    tem_registro = "REGISTRO DE EMPREGADO" in texto_upper
    tem_empregador = "EMPREGADOR" in texto_upper
    tem_funcionario = "FUNCIONÁRIO" in texto_upper or "FUNCIONARIO" in texto_upper
    tem_nome = "NOME" in texto_upper

    return tem_registro and tem_empregador and tem_funcionario and tem_nome


def localizar_kits_no_pdf(doc):
    """
    Um kit começa somente na Ficha de Registro.
    Todas as páginas seguintes pertencem ao mesmo kit até a próxima Ficha.
    """
    inicios = []

    for i in range(len(doc)):
        texto = doc[i].get_text("text")
        if pagina_e_ficha_registro(texto):
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
    novo_pdf = fitz.open()

    for p in range(pagina_inicial, pagina_final + 1):
        novo_pdf.insert_pdf(doc, from_page=p, to_page=p)

    novo_pdf.save(caminho_saida)
    novo_pdf.close()


def encontrar_coluna_sugerida(colunas, lista_possiveis):
    """
    Tenta sugerir automaticamente a coluna com base no cabeçalho.
    """
    mapa = {normalizar_texto(c): c for c in colunas}

    for nome in lista_possiveis:
        nome_norm = normalizar_texto(nome)
        if nome_norm in mapa:
            return mapa[nome_norm]

    return None


# ============================================================
# APLICAÇÃO TKINTER
# ============================================================

class AppKitsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Separador de Kits Admissionais - Versão 2.4")
        self.root.geometry("980x720")
        self.root.minsize(900, 650)

        self.caminho_pdf = tk.StringVar()
        self.caminho_excel = tk.StringVar()
        self.pasta_saida = tk.StringVar()

        self.coluna_nome_var = tk.StringVar()
        self.coluna_pasta_var = tk.StringVar()

        self.df_excel = None
        self.colunas_excel = []

        self._montar_tela()
        self._preencher_padroes()

    # --------------------------------------------------------
    # UI
    # --------------------------------------------------------
    def _montar_tela(self):
        frame_principal = ttk.Frame(self.root, padding=12)
        frame_principal.pack(fill="both", expand=True)

        titulo = ttk.Label(
            frame_principal,
            text="Separador de Kits Admissionais - v2.4",
            font=("Segoe UI", 15, "bold")
        )
        titulo.pack(anchor="w", pady=(0, 12))

        subtitulo = ttk.Label(
            frame_principal,
            text="Selecione o PDF, a planilha e as colunas correspondentes. O programa irá separar os kits e salvar os PDFs nas pastas de destino.",
            wraplength=900
        )
        subtitulo.pack(anchor="w", pady=(0, 16))

        # ----------------------------------------------------
        # Frame de arquivos
        # ----------------------------------------------------
        frame_arquivos = ttk.LabelFrame(frame_principal, text="Arquivos e pasta de saída", padding=12)
        frame_arquivos.pack(fill="x", pady=(0, 12))

        # PDF
        ttk.Label(frame_arquivos, text="PDF Geral:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(frame_arquivos, textvariable=self.caminho_pdf, width=90).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(frame_arquivos, text="Selecionar...", command=self.selecionar_pdf).grid(row=0, column=2, padx=(8, 0), pady=6)

        # Excel
        ttk.Label(frame_arquivos, text="Excel Dados:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(frame_arquivos, textvariable=self.caminho_excel, width=90).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(frame_arquivos, text="Selecionar...", command=self.selecionar_excel).grid(row=1, column=2, padx=(8, 0), pady=6)

        # Pasta de saída
        ttk.Label(frame_arquivos, text="Pasta de saída:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(frame_arquivos, textvariable=self.pasta_saida, width=90).grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Button(frame_arquivos, text="Selecionar...", command=self.selecionar_pasta_saida).grid(row=2, column=2, padx=(8, 0), pady=6)

        frame_arquivos.columnconfigure(1, weight=1)

        # ----------------------------------------------------
        # Frame Excel / colunas
        # ----------------------------------------------------
        frame_colunas = ttk.LabelFrame(frame_principal, text="Configuração das colunas do Excel", padding=12)
        frame_colunas.pack(fill="x", pady=(0, 12))

        ttk.Label(frame_colunas, text="Coluna do Nome:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        self.combo_nome = ttk.Combobox(frame_colunas, textvariable=self.coluna_nome_var, state="readonly", width=50)
        self.combo_nome.grid(row=0, column=1, sticky="w", pady=6)

        ttk.Label(frame_colunas, text="Coluna da Pasta:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        self.combo_pasta = ttk.Combobox(frame_colunas, textvariable=self.coluna_pasta_var, state="readonly", width=50)
        self.combo_pasta.grid(row=1, column=1, sticky="w", pady=6)

        botoes_colunas = ttk.Frame(frame_colunas)
        botoes_colunas.grid(row=0, column=2, rowspan=2, padx=(20, 0), sticky="ns")

        ttk.Button(botoes_colunas, text="Ler cabeçalhos do Excel", command=self.ler_cabecalhos_excel).pack(fill="x", pady=(0, 8))
        ttk.Button(botoes_colunas, text="Sugerir colunas", command=self.sugerir_colunas).pack(fill="x")

        # ----------------------------------------------------
        # Frame ações
        # ----------------------------------------------------
        frame_acoes = ttk.Frame(frame_principal)
        frame_acoes.pack(fill="x", pady=(0, 12))

        self.btn_processar = ttk.Button(frame_acoes, text="Processar Kits", command=self.processar_kits)
        self.btn_processar.pack(side="left")

        self.lbl_status = ttk.Label(frame_acoes, text="Pronto.")
        self.lbl_status.pack(side="left", padx=12)

        # ----------------------------------------------------
        # Log
        # ----------------------------------------------------
        frame_log = ttk.LabelFrame(frame_principal, text="Log de execução", padding=8)
        frame_log.pack(fill="both", expand=True)

        self.txt_log = tk.Text(frame_log, wrap="word", height=25)
        self.txt_log.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(frame_log, orient="vertical", command=self.txt_log.yview)
        scroll.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=scroll.set)

    def _preencher_padroes(self):
        downloads = Path.home() / "Downloads"
        self.pasta_saida.set(str(downloads / NOME_SAIDA_PRINCIPAL))

    # --------------------------------------------------------
    # Utilidades da UI
    # --------------------------------------------------------
    def log(self, mensagem):
        self.txt_log.insert("end", mensagem + "\n")
        self.txt_log.see("end")
        self.root.update_idletasks()

    def limpar_log(self):
        self.txt_log.delete("1.0", "end")

    def set_status(self, texto):
        self.lbl_status.config(text=texto)
        self.root.update_idletasks()

    # --------------------------------------------------------
    # Seleção de arquivos/pasta
    # --------------------------------------------------------
    def selecionar_pdf(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o PDF Geral",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if arquivo:
            self.caminho_pdf.set(arquivo)

    def selecionar_excel(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o Excel Dados",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivo:
            self.caminho_excel.set(arquivo)

    def selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de saída")
        if pasta:
            self.pasta_saida.set(pasta)

    # --------------------------------------------------------
    # Leitura de cabeçalhos / sugestões
    # --------------------------------------------------------
    def ler_cabecalhos_excel(self):
        caminho_excel = self.caminho_excel.get().strip()

        if not caminho_excel:
            messagebox.showwarning("Atenção", "Selecione primeiro o arquivo Excel.")
            return

        try:
            self.df_excel = pd.read_excel(caminho_excel, header=0)
            self.colunas_excel = list(self.df_excel.columns)

            self.combo_nome["values"] = self.colunas_excel
            self.combo_pasta["values"] = self.colunas_excel

            self.log("Cabeçalhos do Excel carregados com sucesso:")
            for col in self.colunas_excel:
                self.log(f" - {col}")

            self.set_status("Cabeçalhos carregados.")
            self.sugerir_colunas()

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o Excel:\n{e}")

    def sugerir_colunas(self):
        if not self.colunas_excel:
            messagebox.showwarning("Atenção", "Primeiro clique em 'Ler cabeçalhos do Excel'.")
            return

        sugestao_nome = encontrar_coluna_sugerida(self.colunas_excel, POSSIVEIS_CABECALHOS_NOME)
        sugestao_pasta = encontrar_coluna_sugerida(self.colunas_excel, POSSIVEIS_CABECALHOS_PASTA)

        if sugestao_nome:
            self.coluna_nome_var.set(sugestao_nome)
            self.log(f"Sugestão de coluna para Nome: {sugestao_nome}")

        if sugestao_pasta:
            self.coluna_pasta_var.set(sugestao_pasta)
            self.log(f"Sugestão de coluna para Pasta: {sugestao_pasta}")

        if not sugestao_nome:
            self.log("Não foi possível sugerir automaticamente a coluna do Nome.")

        if not sugestao_pasta:
            self.log("Não foi possível sugerir automaticamente a coluna da Pasta.")

    # --------------------------------------------------------
    # Processamento
    # --------------------------------------------------------
    def processar_kits(self):
        caminho_pdf = self.caminho_pdf.get().strip()
        caminho_excel = self.caminho_excel.get().strip()
        pasta_saida_base = self.pasta_saida.get().strip()
        coluna_nome = self.coluna_nome_var.get().strip()
        coluna_pasta = self.coluna_pasta_var.get().strip()

        if not caminho_pdf:
            messagebox.showwarning("Atenção", "Selecione o PDF Geral.")
            return

        if not caminho_excel:
            messagebox.showwarning("Atenção", "Selecione o Excel Dados.")
            return

        if not pasta_saida_base:
            messagebox.showwarning("Atenção", "Selecione a pasta de saída.")
            return

        if not coluna_nome:
            messagebox.showwarning("Atenção", "Selecione a coluna do Nome.")
            return

        if not coluna_pasta:
            messagebox.showwarning("Atenção", "Selecione a coluna da Pasta.")
            return

        try:
            self.btn_processar.config(state="disabled")
            self.limpar_log()
            self.set_status("Processando...")

            self._executar_processamento(
                caminho_pdf=caminho_pdf,
                caminho_excel=caminho_excel,
                pasta_saida_base=pasta_saida_base,
                coluna_nome=coluna_nome,
                coluna_pasta=coluna_pasta
            )

            self.set_status("Processamento concluído.")
            messagebox.showinfo("Concluído", "Processamento finalizado com sucesso.")

        except Exception as e:
            self.set_status("Erro durante o processamento.")
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

        finally:
            self.btn_processar.config(state="normal")

    def _executar_processamento(self, caminho_pdf, caminho_excel, pasta_saida_base, coluna_nome, coluna_pasta):
        caminho_pdf = Path(caminho_pdf)
        caminho_excel = Path(caminho_excel)
        pasta_saida_base = Path(pasta_saida_base)

        data_execucao = datetime.now().strftime("%Y-%m-%d")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        pasta_saida_execucao = pasta_saida_base / data_execucao
        pasta_saida_execucao.mkdir(parents=True, exist_ok=True)

        caminho_log = pasta_saida_execucao / f"log_processamento_{timestamp}.txt"
        caminho_relatorio = pasta_saida_execucao / f"Relatorio_Processamento_{timestamp}.xlsx"

        def log_geral(msg):
            self.log(msg)
            with open(caminho_log, "a", encoding="utf-8") as f:
                f.write(msg + "\n")

        # ----------------------------------------------------
        # Valida arquivos
        # ----------------------------------------------------
        if not caminho_pdf.exists():
            raise FileNotFoundError(f"PDF não encontrado: {caminho_pdf}")

        if not caminho_excel.exists():
            raise FileNotFoundError(f"Excel não encontrado: {caminho_excel}")

        log_geral("=" * 90)
        log_geral("INÍCIO DO PROCESSAMENTO")
        log_geral("=" * 90)
        log_geral(f"PDF: {caminho_pdf}")
        log_geral(f"Excel: {caminho_excel}")
        log_geral(f"Pasta de saída base: {pasta_saida_base}")
        log_geral(f"Pasta de saída da execução: {pasta_saida_execucao}")
        log_geral(f"Coluna Nome selecionada: {coluna_nome}")
        log_geral(f"Coluna Pasta selecionada: {coluna_pasta}")

        # ----------------------------------------------------
        # Lê Excel
        # ----------------------------------------------------
        log_geral("\n[1/4] Carregando base do Excel...")
        df = pd.read_excel(caminho_excel, header=0)

        if coluna_nome not in df.columns:
            raise ValueError(f"A coluna '{coluna_nome}' não existe no Excel.")

        if coluna_pasta not in df.columns:
            raise ValueError(f"A coluna '{coluna_pasta}' não existe no Excel.")

        base_colaboradores = {}
        linhas_invalidas = []

        for idx, row in df.iterrows():
            numero_linha_excel = idx + 2

            try:
                nome = row[coluna_nome]
                pasta = row[coluna_pasta]

                if pd.isna(nome) or pd.isna(pasta):
                    continue

                nome = str(nome).strip()
                pasta = str(pasta).strip()

                if not nome or not pasta:
                    continue

                nome_norm = normalizar_texto(nome)

                base_colaboradores[nome_norm] = {
                    "nome_original": nome,
                    "pasta": pasta
                }

            except Exception as e:
                linhas_invalidas.append({
                    "linha_excel": numero_linha_excel,
                    "erro": str(e)
                })

        if not base_colaboradores:
            raise ValueError("Nenhum colaborador válido encontrado na planilha.")

        log_geral(f"Total de colaboradores carregados: {len(base_colaboradores)}")

        # ----------------------------------------------------
        # Lê PDF e encontra kits
        # ----------------------------------------------------
        log_geral("\n[2/4] Abrindo PDF e localizando kits...")
        doc = fitz.open(caminho_pdf)
        kits = localizar_kits_no_pdf(doc)

        if not kits:
            doc.close()
            raise ValueError("Nenhum kit encontrado no PDF.")

        log_geral(f"Total de kits encontrados no PDF: {len(kits)}")

        # ----------------------------------------------------
        # Processa kits
        # ----------------------------------------------------
        log_geral("\n[3/4] Processando kits...")

        processados = []
        nao_encontrados = []
        sem_nome = []
        erros = []

        total_kits = len(kits)

        for idx, (inicio, fim) in enumerate(kits, start=1):
            try:
                log_geral(f"\n--- Kit {idx}/{total_kits} | páginas {inicio + 1} a {fim + 1} ---")

                texto_primeira_pagina = doc[inicio].get_text("text")
                nome_pdf = extrair_nome_da_pagina(texto_primeira_pagina)

                if not nome_pdf:
                    msg = f"[AVISO] Nome não identificado no kit {idx} (página inicial {inicio + 1})."
                    log_geral(msg)

                    sem_nome.append({
                        "kit": idx,
                        "pagina_inicial": inicio + 1,
                        "pagina_final": fim + 1
                    })
                    continue

                nome_pdf = re.sub(r"\s+", " ", nome_pdf).strip()
                nome_pdf_norm = normalizar_texto(nome_pdf)

                log_geral(f"Nome identificado no PDF: {nome_pdf}")

                if nome_pdf_norm not in base_colaboradores:
                    msg = f"[AVISO] Nome não encontrado na planilha: {nome_pdf}"
                    log_geral(msg)

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

                pasta_destino = pasta_saida_execucao / nome_arquivo_seguro(nome_pasta)
                pasta_destino.mkdir(parents=True, exist_ok=True)

                nome_pdf_saida = f"Kit Admissional - {nome_arquivo_seguro(nome_excel)}.pdf"
                caminho_pdf_saida = pasta_destino / nome_pdf_saida
                caminho_pdf_saida = gerar_nome_sem_sobrescrever(caminho_pdf_saida)

                salvar_kit(doc, inicio, fim, caminho_pdf_saida)

                log_geral(f"[OK] Kit salvo em: {caminho_pdf_saida}")

                processados.append({
                    "kit": idx,
                    "nome_pdf_extraido": nome_pdf,
                    "nome_excel": nome_excel,
                    "pasta_destino_base": nome_pasta,
                    "pasta_destino_completa": str(pasta_destino),
                    "arquivo_salvo": str(caminho_pdf_saida),
                    "pagina_inicial": inicio + 1,
                    "pagina_final": fim + 1,
                    "qtde_paginas_kit": (fim - inicio + 1)
                })

            except Exception as e:
                msg = f"[ERRO] Falha ao processar kit {idx}: {str(e)}"
                log_geral(msg)

                erros.append({
                    "kit": idx,
                    "pagina_inicial": inicio + 1,
                    "pagina_final": fim + 1,
                    "erro": str(e)
                })

        doc.close()

        # ----------------------------------------------------
        # Relatório
        # ----------------------------------------------------
        log_geral("\n[4/4] Gerando relatório final...")

        resumo = [{
            "data_execucao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "coluna_nome_utilizada": coluna_nome,
            "coluna_pasta_utilizada": coluna_pasta,
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

        log_geral("\n" + "=" * 90)
        log_geral("PROCESSAMENTO FINALIZADO")
        log_geral("=" * 90)
        log_geral(f"Kits encontrados no PDF: {len(kits)}")
        log_geral(f"Kits processados com sucesso: {len(processados)}")
        log_geral(f"Kits sem nome identificado: {len(sem_nome)}")
        log_geral(f"Nomes não encontrados na planilha: {len(nao_encontrados)}")
        log_geral(f"Erros de processamento: {len(erros)}")
        log_geral(f"Linhas inválidas no Excel: {len(linhas_invalidas)}")
        log_geral(f"Relatório Excel: {caminho_relatorio}")
        log_geral(f"Log TXT: {caminho_log}")


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = AppKitsGUI(root)
    root.mainloop()