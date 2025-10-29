import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# -----------------------------
# Utilidades
# -----------------------------
def excel_col_index(col_letters: str) -> int:
    """Converte letras de coluna do Excel (ex.: 'J', 'AR') para índice zero-based."""
    col_letters = col_letters.strip().upper()
    total = 0
    for ch in col_letters:
        total = total * 26 + (ord(ch) - ord('A') + 1)
    return total - 1  # zero-based

COL_J  = excel_col_index('J')   # 9
COL_S  = excel_col_index('S')   # 18
COL_T  = excel_col_index('T')   # 19
COL_U  = excel_col_index('U')   # 20
COL_V  = excel_col_index('V')   # 21
COL_AR = excel_col_index('AR')  # 43

def ensure_min_columns(df: pd.DataFrame, min_cols_index: int) -> pd.DataFrame:
    """Garante que o DataFrame tenha pelo menos (min_cols_index+1) colunas."""
    needed = (min_cols_index + 1) - df.shape[1]
    if needed > 0:
        for _ in range(needed):
            df[df.shape[1]] = ""
    return df

def only_digits(s) -> str:
    """Retorna apenas os dígitos de s; '' para NaN."""
    if pd.isna(s):
        return ""
    s = str(s).strip()
    return "".join(ch for ch in s if ch.isdigit())

def ncm8(x: object) -> str:
    """
    Normaliza qualquer entrada para NCM com 8 dígitos:
    - remove tudo que não for dígito
    - se tiver 9+ dígitos (ex.: '27101230.0' -> '271012300'), corta para os primeiros 8
    - se tiver <8, completa com zeros à esquerda
    """
    d = only_digits(x)
    if not d:
        return ""
    if len(d) >= 8:
        return d[:8]
    return d.zfill(8)

# -----------------------------
# App Tkinter
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Processar NCM/CEST - produto.xls")
        self.geometry("700x420")
        self.resizable(False, False)

        # Estados
        self.ncm_path = tk.StringVar()
        self.prod_path = tk.StringVar()
        self.ncm_padrao = tk.StringVar()
        self.lookup_ncm = tk.StringVar()
        self.lookup_desc = tk.StringVar(value="—")

        # Cache ncm.xlsx
        self.ncm_map = None  # dict NCM(8 dígitos) -> (CEST, DESCR)

        self._build_ui()

    def _build_ui(self):
        pad = 8

        frm_files = tk.LabelFrame(self, text="Arquivos")
        frm_files.pack(fill="x", padx=pad, pady=(pad, 4))

        # ncm.xlsx
        row1 = tk.Frame(frm_files); row1.pack(fill="x", padx=pad, pady=(pad, 4))
        tk.Label(row1, text="ncm.xlsx:").pack(side="left")
        tk.Entry(row1, textvariable=self.ncm_path, width=70).pack(side="left", padx=(6, 6))
        tk.Button(row1, text="Selecionar", command=self.sel_ncm).pack(side="left")

        # produto.xls
        row2 = tk.Frame(frm_files); row2.pack(fill="x", padx=pad, pady=(0, pad))
        tk.Label(row2, text="produto.xls:").pack(side="left")
        tk.Entry(row2, textvariable=self.prod_path, width=70).pack(side="left", padx=(6, 6))
        tk.Button(row2, text="Selecionar", command=self.sel_prod).pack(side="left")

        # Parâmetros
        frm_defaults = tk.LabelFrame(self, text="Parâmetros")
        frm_defaults.pack(fill="x", padx=pad, pady=4)
        row3 = tk.Frame(frm_defaults); row3.pack(fill="x", padx=pad, pady=(pad, pad))
        tk.Label(row3, text="NCM padrão (usado só quando J estiver vazio):").pack(side="left")
        tk.Entry(row3, textvariable=self.ncm_padrao, width=20).pack(side="left", padx=(6, 12))
        tk.Label(row3, text="(8 dígitos, ex.: 00000000)").pack(side="left")

        # Lookup NCM -> Descrição
        frm_lookup = tk.LabelFrame(self, text="Buscar descrição do NCM em ncm.xlsx (coluna C)")
        frm_lookup.pack(fill="x", padx=pad, pady=4)
        row4 = tk.Frame(frm_lookup); row4.pack(fill="x", padx=pad, pady=(pad, 4))
        tk.Label(row4, text="NCM:").pack(side="left")
        tk.Entry(row4, textvariable=self.lookup_ncm, width=20).pack(side="left", padx=(6, 6))
        tk.Button(row4, text="Buscar descrição", command=self.do_lookup).pack(side="left")

        row5 = tk.Frame(frm_lookup); row5.pack(fill="x", padx=pad, pady=(0, pad))
        tk.Label(row5, text="Descrição:").pack(side="left")
        tk.Label(row5, textvariable=self.lookup_desc, fg="#333").pack(side="left", padx=(6, 0))

        # Ações
        frm_actions = tk.Frame(self)
        frm_actions.pack(fill="x", padx=pad, pady=(8, pad))
        tk.Button(frm_actions, text="Processar e salvar (produto_atualizado.xlsx)", command=self.processar).pack(side="left")
        tk.Label(frm_actions, text="   ").pack(side="left")
        tk.Button(frm_actions, text="Sair", command=self.destroy).pack(side="left")

        lbl = tk.Label(self, text="Regras: só preenche J quando vazio; S/T/U/V/AR via CEST (se J tiver 8 dígitos).", fg="#666")
        lbl.pack(side="bottom", pady=(0, 6))

    # -------------------------
    # Seletores
    # -------------------------
    def sel_ncm(self):
        path = filedialog.askopenfilename(
            title="Selecione ncm.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        if path:
            self.ncm_path.set(path)
            self._load_ncm_df()

    def sel_prod(self):
        path = filedialog.askopenfilename(
            title="Selecione produto.xls",
            filetypes=[("Excel 97-2003", "*.xls"), ("Excel", "*.xlsx;*.xls"), ("Todos", "*.*")]
        )
        if path:
            self.prod_path.set(path)

    # -------------------------
    # Carregar ncm.xlsx
    # -------------------------
    def _load_ncm_df(self):
        p = self.ncm_path.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showerror("Erro", "Arquivo ncm.xlsx não encontrado.")
            return
        try:
            df = pd.read_excel(p, engine="openpyxl")  # Espera A=NCM, B=CEST, C=Descrição
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler '{os.path.basename(p)}'.\n{e}")
            return

        if df.shape[1] < 3:
            messagebox.showerror("Erro", "A planilha ncm.xlsx deve ter ao menos 3 colunas (A=NCM, B=CEST, C=Descrição).")
            return

        # Normaliza NCM (8 dígitos)
        df = df.copy()
        df.iloc[:, 0] = df.iloc[:, 0].apply(ncm8)

        # Cria mapa NCM -> (CEST, DESCR)
        ncm_map = {}
        for _, r in df.iterrows():
            key = r.iloc[0]  # NCM normalizado (8 dígitos)
            if key:
                cest = "" if pd.isna(r.iloc[1]) else str(r.iloc[1]).strip()
                desc = "" if pd.isna(r.iloc[2]) else str(r.iloc[2]).strip()
                ncm_map[key] = (cest, desc)

        self.ncm_map = ncm_map
        messagebox.showinfo("OK", f"ncm.xlsx carregado ({len(ncm_map)} NCMs).")

    # -------------------------
    # Lookup descrição
    # -------------------------
    def do_lookup(self):
        if self.ncm_map is None:
            self._load_ncm_df()
            if self.ncm_map is None:
                return
        key = ncm8(self.lookup_ncm.get()) if self.lookup_ncm.get().strip() else ""
        if not key:
            self.lookup_desc.set("—")
            return
        tupla = self.ncm_map.get(key)
        self.lookup_desc.set("(NCM não encontrado na ncm.xlsx)" if tupla is None else (tupla[1] or "(sem descrição)"))

    # -------------------------
    # Processamento principal
    # -------------------------
    def processar(self):
        if not self.ncm_path.get().strip():
            messagebox.showerror("Erro", "Selecione o arquivo ncm.xlsx.")
            return
        if not self.prod_path.get().strip():
            messagebox.showerror("Erro", "Selecione o arquivo produto.xls.")
            return
        if self.ncm_map is None:
            self._load_ncm_df()
            if self.ncm_map is None:
                return

        # NCM padrão (string com zeros à esquerda)
        ncm_default_input = self.ncm_padrao.get().strip()
        ncm_default = ncm8(ncm_default_input) if ncm_default_input else ""
        if ncm_default_input and len(ncm_default) != 8:
            messagebox.showerror("Erro", "NCM padrão deve ter exatamente 8 dígitos (ex.: 00000000).")
            return

        # Ler produto.xls/.xlsx
        prod_path = self.prod_path.get().strip()
        try:
            df = pd.read_excel(prod_path, header=0)
        except Exception as e:
            messagebox.showerror(
                "Erro ao ler produto",
                "Não foi possível abrir o arquivo de produtos.\n"
                "Se for .xls, instale 'xlrd==1.2.0' (pip install xlrd==1.2.0) ou salve como .xlsx e tente novamente.\n\n"
                f"Detalhes:\n{e}"
            )
            return

        # Garantir colunas até AR e trabalhar em cópia
        df = ensure_min_columns(df, COL_AR).copy()

        # ===== Loop: SÓ altera J se estiver vazio; lookup com ncm8 =====
        for idx in range(len(df)):
            ncm_raw = df.iat[idx, COL_J]

            # vazio? (NaN ou string em branco)
            is_empty = (pd.isna(ncm_raw) or str(ncm_raw).strip() == "")
            if is_empty and ncm_default:
                df.iat[idx, COL_J] = ncm_default
                ncm_use = ncm_default
            else:
                # não altera J
                ncm_use = str(ncm_raw).strip() if not pd.isna(ncm_raw) else ""

            # Para lookup de CEST, normaliza pra 8 dígitos com ncm8
            ncm_digits = ncm8(ncm_use)
            if ncm_digits:
                tupla = self.ncm_map.get(ncm_digits)
                cest = tupla[0] if tupla is not None else "Não se aplica"
            else:
                cest = "Não se aplica"

            nao_aplica = (str(cest).strip().lower() in ("não se aplica", "nao se aplica"))

            # S e T (strings)
            df.iat[idx, COL_S] = "0101" if nao_aplica else "0500"
            df.iat[idx, COL_T] = "0101" if nao_aplica else "0500"

            # U e V (inteiros)
            df.iat[idx, COL_U] = 102 if nao_aplica else 403
            df.iat[idx, COL_V] = 102 if nao_aplica else 405

            # AR (vazio ou o CEST da ncm.xlsx)
            df.iat[idx, COL_AR] = "" if nao_aplica else str(cest)

        # Salvar como XLSX (não sobrescreve o .xls original)
        out_dir = os.path.dirname(prod_path) or "."
        out_path = os.path.join(out_dir, "produto_atualizado.xlsx")
        try:
            df.to_excel(out_path, index=False, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Falha ao salvar '{out_path}'.\n{e}")
            return

        messagebox.showinfo("Concluído", f"Arquivo salvo:\n{out_path}")

# -----------------------------
# Main
# -----------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
