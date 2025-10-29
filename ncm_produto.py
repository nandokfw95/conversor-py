# ajusta_ncm_petshop.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os

# ------------------ Helpers ------------------
def clean_ncm(n):
    if pd.isna(n): return ""
    s = re.sub(r"\D", "", str(n))
    return s

def is_valid_ncm(ncm_str: str) -> bool:
    return bool(re.fullmatch(r"\d{8}", ncm_str or ""))

def norm_text(s):
    return str(s).strip().lower()

# ------------------ App ------------------
class NCMApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ajuste de NCM - Petshop")
        self.geometry("1280x680")

        self.df = pd.DataFrame(columns=["codigo","descricao","ncm","alterado","marcado"])
        self.df_view = self.df.copy()
        self.current_file = None

        self.checkbox_unchecked = "☐"
        self.checkbox_checked = "☑"

        self._build_ui()

    def _build_ui(self):
        # Top bar: Load / Save
        bar = ttk.Frame(self, padding=6)
        bar.pack(fill="x")

        ttk.Button(bar, text="Carregar planilha de produtos…", command=self.load_products).pack(side="left", padx=4)
        ttk.Button(bar, text="Salvar Excel…", command=self.save_excel).pack(side="left", padx=4)

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=4)

        # Search + Filtro Alterado
        search_frame = ttk.Frame(self, padding=6)
        search_frame.pack(fill="x")

        ttk.Label(search_frame, text="Buscar (código/descrição):").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.apply_filter())
        ttk.Entry(search_frame, textvariable=self.search_var, width=50).pack(side="left", padx=8)
        ttk.Button(search_frame, text="Limpar", command=lambda: self.search_var.set("")).pack(side="left", padx=8)

        ttk.Label(search_frame, text="Mostrar:").pack(side="left", padx=(16, 4))
        self.alterado_filter = tk.StringVar(value="Todos")
        alterado_cb = ttk.Combobox(
            search_frame, textvariable=self.alterado_filter,
            values=["Todos", "Alterados", "Não alterados"], width=18, state="readonly"
        )
        alterado_cb.pack(side="left")
        alterado_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

        # Table
        table_frame = ttk.Frame(self, padding=6)
        table_frame.pack(fill="both", expand=True)

        cols = ("marcar","codigo","descricao","ncm","alterado")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("marcar", text="Marcar")
        self.tree.heading("codigo", text="Código")
        self.tree.heading("descricao", text="Descrição")
        self.tree.heading("ncm", text="NCM")
        self.tree.heading("alterado", text="Alterado")

        self.tree.column("marcar", width=80, anchor="center")
        self.tree.column("codigo", width=140, anchor="w")
        self.tree.column("descricao", width=820, anchor="w")
        self.tree.column("ncm", width=100, anchor="center")
        self.tree.column("alterado", width=100, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Click handlers
        self.tree.bind("<Button-1>", self._on_click_checkbox)  # clique simples na coluna "Marcar"
        self.tree.bind("<Double-1>", self._on_double_click)    # duplo clique na coluna NCM para editar

        # Apply bar
        apply_frame = ttk.Frame(self, padding=6)
        apply_frame.pack(fill="x")

        ttk.Label(apply_frame, text="Novo NCM (8 dígitos):").pack(side="left")
        self.ncm_var = tk.StringVar()
        ttk.Entry(apply_frame, textvariable=self.ncm_var, width=15).pack(side="left", padx=6)

        ttk.Button(apply_frame, text="Aplicar ao(s) selecionado(s)", command=self.apply_to_selected).pack(side="left", padx=6)
        ttk.Button(apply_frame, text="Aplicar a TODOS os filtrados", command=self.apply_to_filtered).pack(side="left", padx=6)
        ttk.Button(apply_frame, text="Aplicar ao(s) MARCADOS (☑)", command=self.apply_to_marked).pack(side="left", padx=12)

        ttk.Button(apply_frame, text="Marcar/Desmarcar filtrados (toggle)", command=self.toggle_mark_filtered).pack(side="left", padx=6)

        # Status bar
        self.status = tk.StringVar(value="Carregue sua planilha de produtos.")
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w").pack(side="bottom", fill="x")

    # ------------------ Data I/O ------------------
    def load_products(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha de produtos",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if not path:
            return
        self.current_file = path
        try:
            xls = pd.ExcelFile(path)
            df_raw = pd.read_excel(path, sheet_name=xls.sheet_names[0])
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o arquivo:\n{e}")
            return

        df_raw.columns = [str(c).strip().lower() for c in df_raw.columns]

        code_col = self._pick_col(df_raw.columns, ["codigo","código","cod","id","sku"])
        desc_col = self._pick_col(df_raw.columns, ["descricao","descrição","nome","produto","item"])
        ncm_col  = self._pick_col(df_raw.columns, ["cod_ncm","ncm","codigo_ncm","c_ncm"])

        if not code_col or not desc_col:
            messagebox.showerror("Erro", "Não encontrei colunas de 'código' e 'descrição' no arquivo.")
            return

        df = pd.DataFrame({
            "codigo": df_raw[code_col],
            "descricao": df_raw[desc_col],
            "ncm": df_raw[ncm_col] if ncm_col else "",
        })
        df["ncm"] = df["ncm"].map(clean_ncm)

        # Coluna 'alterado'
        df["alterado"] = df_raw.get("alterado", "NÃO")
        df.loc[~df["alterado"].astype(str).isin(["SIM","NÃO"]), "alterado"] = "NÃO"

        # Nova coluna 'marcado' (checkbox)
        df["marcado"] = False

        df["codigo"] = df["codigo"].astype(str)
        df["descricao"] = df["descricao"].astype(str)

        self.df = df.copy()
        self.apply_filter()
        self._update_status_loaded(path)

    def save_excel(self):
        if self.df.empty:
            messagebox.showwarning("Aviso", "Nada para salvar. Carregue produtos primeiro.")
            return
        path = filedialog.asksaveasfilename(
            title="Salvar planilha resultante",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return
        try:
            out = self.df[["codigo","descricao","ncm","alterado"]].copy()
            out.to_excel(path, index=False)
            self.status.set(f"Salvo em: {path}")
            messagebox.showinfo("OK", "Planilha salva com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar:\n{e}")

    # ------------------ Filtering & UI ------------------
    def apply_filter(self):
        term = norm_text(self.search_var.get())
        dfv = self.df.copy()
        if term:
            dfv = dfv[
                dfv["codigo"].astype(str).str.lower().str.contains(term, na=False) |
                dfv["descricao"].astype(str).str.lower().str.contains(term, na=False)
            ].copy()

        f = self.alterado_filter.get()
        if f == "Alterados":
            dfv = dfv[dfv["alterado"] == "SIM"].copy()
        elif f == "Não alterados":
            dfv = dfv[dfv["alterado"] == "NÃO"].copy()

        self.df_view = dfv
        self._refresh_table()

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        # iid único por linha (codigo pode repetir -> adicionar índice)
        for i, row in self.df_view.reset_index(drop=True).iterrows():
            iid = f"{row['codigo']}|{i}"
            chk = self.checkbox_checked if bool(row.get("marcado", False)) else self.checkbox_unchecked
            values = (chk, row["codigo"], row["descricao"], row["ncm"] or "", row["alterado"])
            self.tree.insert("", "end", iid=iid, values=values)

    # ------------------ Checkbox handling ------------------
    def _on_click_checkbox(self, event):
        # Detecta clique na coluna "Marcar" e alterna o checkbox da linha
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)  # '#1' == 'marcar'
        row_id = self.tree.identify_row(event.y)
        if not row_id or col != "#1":
            return

        vals = self.tree.item(row_id, "values")
        # mapear linha no df a partir de (codigo, descricao) exibidos
        codigo, descricao = vals[1], vals[2]
        idx = self.df.index[(self.df["codigo"] == str(codigo)) & (self.df["descricao"] == str(descricao))]
        if len(idx):
            current = bool(self.df.loc[idx, "marcado"].iloc[0])
            self.df.loc[idx, "marcado"] = not current
            self.apply_filter()  # refresh visual

    def toggle_mark_filtered(self):
        if self.df_view.empty:
            messagebox.showwarning("Sem resultados", "Não há itens filtrados.")
            return
        # Se mais da metade dos filtrados estiver marcada, desmarca todos; senão, marca todos
        marked_count = int(self.df_view["marcado"].sum())
        mark_all = marked_count < (len(self.df_view) / 2)

        keys = self.df_view[["codigo","descricao"]].astype(str)
        merged = self.df.merge(keys, on=["codigo","descricao"], how="left", indicator=True)
        idx = merged.index[merged["_merge"] == "both"]
        self.df.loc[idx, "marcado"] = mark_all

        self.apply_filter()

    # ------------------ Apply NCM ------------------
    def apply_to_selected(self):
        new_ncm = clean_ncm(self.ncm_var.get())
        if not is_valid_ncm(new_ncm):
            messagebox.showwarning("NCM inválido", "Informe um NCM com 8 dígitos (apenas números).")
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Seleção vazia", "Selecione uma ou mais linhas na tabela.")
            return

        selected_rows = []
        for iid in sel:
            vals = self.tree.item(iid, "values")
            selected_rows.append((vals[1], vals[2]))  # (codigo, descricao)

        cnt = 0
        for codigo, descricao in selected_rows:
            idx = self.df.index[(self.df["codigo"] == str(codigo)) & (self.df["descricao"] == str(descricao))]
            if len(idx):
                self.df.loc[idx, "ncm"] = new_ncm
                self.df.loc[idx, "alterado"] = "SIM"
                cnt += len(idx)

        self.apply_filter()
        self.status.set(f"NCM aplicado a {cnt} item(ns) selecionado(s).")

    def apply_to_filtered(self):
        new_ncm = clean_ncm(self.ncm_var.get())
        if not is_valid_ncm(new_ncm):
            messagebox.showwarning("NCM inválido", "Informe um NCM com 8 dígitos (apenas números).")
            return
        if self.df_view.empty:
            messagebox.showwarning("Sem resultados", "Não há itens filtrados.")
            return
        if not messagebox.askyesno("Confirmar", f"Aplicar NCM {new_ncm} a TODOS os {len(self.df_view)} itens filtrados?"):
            return

        merged = self.df.merge(self.df_view[["codigo","descricao"]], on=["codigo","descricao"], how="left", indicator=True)
        idx = merged.index[merged["_merge"] == "both"]
        self.df.loc[idx, "ncm"] = new_ncm
        self.df.loc[idx, "alterado"] = "SIM"

        self.apply_filter()
        self.status.set(f"NCM aplicado a {len(idx)} item(ns) filtrado(s).")

    def apply_to_marked(self):
        new_ncm = clean_ncm(self.ncm_var.get())
        if not is_valid_ncm(new_ncm):
            messagebox.showwarning("NCM inválido", "Informe um NCM com 8 dígitos (apenas números).")
            return
        # Só dentro do conjunto filtrado + marcados
        if self.df_view.empty:
            messagebox.showwarning("Sem resultados", "Não há itens filtrados.")
            return

        keys = self.df_view[self.df_view["marcado"] == True][["codigo","descricao"]].astype(str)
        if keys.empty:
            messagebox.showwarning("Nada marcado", "Marque ao menos um item (☑) nos filtrados.")
            return

        merged = self.df.merge(keys, on=["codigo","descricao"], how="left", indicator=True)
        idx = merged.index[merged["_merge"] == "both"]
        self.df.loc[idx, "ncm"] = new_ncm
        self.df.loc[idx, "alterado"] = "SIM"

        # opcional: desmarcar após aplicar
        self.df.loc[idx, "marcado"] = False

        self.apply_filter()
        self.status.set(f"NCM aplicado a {len(idx)} item(ns) MARCADO(S).")

    # ------------------ Inline edit (double-click) ------------------
    def _on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)  # '#1'..'#5'
        if not item or column != "#4":  # só permite edição na coluna NCM (marcar=#1, codigo=#2, descricao=#3, ncm=#4)
            return
        x, y, width, height = self.tree.bbox(item, column)
        vals = self.tree.item(item, "values")
        current_ncm = vals[3]

        editor = ttk.Entry(self.tree, width=12)
        editor.place(x=x, y=y, width=width, height=height)
        editor.insert(0, current_ncm)
        editor.focus()

        def on_return(event=None):
            new_val = clean_ncm(editor.get())
            if not is_valid_ncm(new_val):
                messagebox.showwarning("NCM inválido", "Informe um NCM com 8 dígitos (apenas números).")
            else:
                codigo, descricao = vals[1], vals[2]
                idx = self.df.index[(self.df["codigo"] == str(codigo)) & (self.df["descricao"] == str(descricao))]
                if len(idx):
                    self.df.loc[idx, "ncm"] = new_val
                    self.df.loc[idx, "alterado"] = "SIM"
                    self.apply_filter()
            editor.destroy()

        def on_escape(event=None):
            editor.destroy()

        editor.bind("<Return>", on_return)
        editor.bind("<Escape>", on_escape)

    @staticmethod
    def _pick_col(cols, candidates):
        cols = list(cols)
        for c in candidates:
            if c in cols:
                return c
        return None

    def _update_status_loaded(self, path):
        total = len(self.df)
        alt = int((self.df["alterado"] == "SIM").sum())
        self.status.set(f"Carregado: {os.path.basename(path)} — {total} itens (Alterados: {alt})")

# ------------------ Run ------------------
if __name__ == "__main__":
    app = NCMApp()
    app.mainloop()
