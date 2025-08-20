# main.py
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict

# módulos ativos
import clientes
import fornecedores

APP_TITLE = "Correção de Planilhas – Clientes / Fornecedores"
CONFIG_FILE = "config.json"

# ---------------------- Persistência ----------------------

def carregar_config() -> Dict[str, str]:
    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def salvar_config(config: Dict[str, str]):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Erro ao salvar config:", e)

# ---------------------- Execuções ----------------------

def run_clientes(defaults: Dict[str, str], arquivos: Dict[str, str]) -> str:
    return clientes.processar_clientes(
        clientes_path=arquivos["clientes"],
        municipios_csv_path=arquivos["municipios"],
        cidade_default=defaults["cidade"],
        uf_default=defaults["uf"],
        ibge_default=defaults["ibge"],
        cep_default=defaults["cep"],
        ddd_default=defaults["ddd"],
    )

def run_fornecedores(defaults: Dict[str, str], arquivos: Dict[str, str]) -> str:
    return fornecedores.processar_fornecedores(
        fornecedores_path=arquivos["fornecedores"],
        municipios_path=arquivos["municipios"],
        cidade_default=defaults["cidade"],
        uf_default=defaults["uf"],
        ibge_default=defaults["ibge"],
        cep_default=defaults["cep"],
        ddd_default=defaults["ddd"],
    )

# ---------------------- UI ----------------------

def main():
    root = tk.Tk()
    root.title(APP_TITLE)

    cfg = carregar_config()

    # --- Defaults (lembrados) ---
    var_cidade = tk.StringVar(value=cfg.get("cidade", ""))
    var_uf = tk.StringVar(value=cfg.get("uf", ""))
    var_ibge = tk.StringVar(value=cfg.get("ibge", ""))
    var_cep = tk.StringVar(value=cfg.get("cep", ""))
    var_ddd = tk.StringVar(value=cfg.get("ddd", ""))

    # --- Arquivos (lembrados) ---
    var_municipios = tk.StringVar(value=cfg.get("municipios", r"C:\Conversor py\municipios.xlsx"))
    var_clientes = tk.StringVar(value=cfg.get("clientes", ""))
    var_fornecedores = tk.StringVar(value=cfg.get("fornecedores", ""))

    # --- Seleção de módulos ---
    var_run_clientes = tk.BooleanVar(value=cfg.get("run_clientes", True))
    var_run_fornecedores = tk.BooleanVar(value=cfg.get("run_fornecedores", False))

    # --- Helpers: pickers ---
    def pick_municipios():
        path = filedialog.askopenfilename(
            title="Selecione municipios (CSV ou Excel)",
            initialdir=r"C:\Conversor py",
            filetypes=[("Planilhas", "*.csv *.xls *.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if path:
            var_municipios.set(path)

    def pick_clientes():
        path = filedialog.askopenfilename(
            title="Selecione a planilha de CLIENTES (.xls ou .xlsx)",
            filetypes=[("Planilhas Excel", "*.xls *.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if path:
            var_clientes.set(path)

    def pick_fornecedores():
        path = filedialog.askopenfilename(
            title="Selecione a planilha de FORNECEDORES (.xls ou .xlsx)",
            filetypes=[("Planilhas Excel", "*.xls *.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if path:
            var_fornecedores.set(path)

    # --- Execução ---
    def executar():
        # ao menos um tipo
        if not (var_run_clientes.get() or var_run_fornecedores.get()):
            messagebox.showwarning("Atenção", "Selecione pelo menos um tipo para corrigir (Clientes e/ou Fornecedores).")
            return

        # arquivo de municípios é compartilhado
        if not var_municipios.get().strip():
            messagebox.showerror("Erro", "Selecione o arquivo de municípios (CSV/Excel).")
            return

        # valida arquivos de cada tipo selecionado
        if var_run_clientes.get() and not var_clientes.get().strip():
            messagebox.showerror("Erro", "Selecione o arquivo de CLIENTES.")
            return
        if var_run_fornecedores.get() and not var_fornecedores.get().strip():
            messagebox.showerror("Erro", "Selecione o arquivo de FORNECEDORES.")
            return

        defaults = {
            "cidade": var_cidade.get().strip(),
            "uf": var_uf.get().strip().upper(),
            "ibge": var_ibge.get().strip(),
            "cep": var_cep.get().strip(),
            "ddd": var_ddd.get().strip(),
        }
        arquivos = {
            "municipios": var_municipios.get().strip(),
            "clientes": var_clientes.get().strip(),
            "fornecedores": var_fornecedores.get().strip(),
        }

        resultados = []
        try:
            if var_run_clientes.get():
                out_cli = run_clientes(defaults, arquivos)
                resultados.append(f"Clientes corrigido: {out_cli}")

            if var_run_fornecedores.get():
                out_forn = run_fornecedores(defaults, arquivos)
                resultados.append(f"Fornecedores corrigido: {out_forn}")

            if resultados:
                messagebox.showinfo("Concluído", "\n".join(resultados))
            else:
                messagebox.showwarning("Atenção", "Nenhuma saída gerada.")

            # salvar config após sucesso
            salvar_config({
                **defaults,
                **arquivos,
                "run_clientes": var_run_clientes.get(),
                "run_fornecedores": var_run_fornecedores.get(),
            })

        except Exception as e:
            messagebox.showerror("Erro durante o processamento", str(e))

    # --- Layout ---
    PADX, PADY = 6, 4
    col_lbl, col_inp = 0, 1
    row = 0

    # Defaults
    tk.Label(root, text="Cidade padrão:").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_cidade, width=35).grid(row=row, column=col_inp, padx=PADX, pady=PADY); row += 1

    tk.Label(root, text="UF padrão:").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_uf, width=10).grid(row=row, column=col_inp, sticky="w", padx=PADX, pady=PADY); row += 1

    tk.Label(root, text="Código IBGE padrão:").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_ibge, width=20).grid(row=row, column=col_inp, sticky="w", padx=PADX, pady=PADY); row += 1

    tk.Label(root, text="CEP padrão:").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_cep, width=20).grid(row=row, column=col_inp, sticky="w", padx=PADX, pady=PADY); row += 1

    tk.Label(root, text="DDD padrão (YY):").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_ddd, width=10).grid(row=row, column=col_inp, sticky="w", padx=PADX, pady=PADY); row += 1

    # Municípios (compartilhado)
    tk.Label(root, text="Arquivo municípios:").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_municipios, width=60).grid(row=row, column=col_inp, padx=PADX, pady=PADY)
    tk.Button(root, text="Procurar...", command=pick_municipios).grid(row=row, column=2, padx=PADX, pady=PADY); row += 1

    # Seletores de tipo
    tk.Label(root, text="O que corrigir:").grid(row=row, column=col_lbl, sticky="ne", padx=PADX, pady=PADY)
    box = tk.Frame(root)
    box.grid(row=row, column=col_inp, sticky="w", padx=PADX, pady=PADY)
    tk.Checkbutton(box, text="Clientes", variable=var_run_clientes).grid(row=0, column=0, padx=(0, 12))
    tk.Checkbutton(box, text="Fornecedores", variable=var_run_fornecedores).grid(row=0, column=1)
    row += 1

    # Arquivo clientes
    tk.Label(root, text="Arquivo CLIENTES:").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_clientes, width=60).grid(row=row, column=col_inp, padx=PADX, pady=PADY)
    tk.Button(root, text="Procurar...", command=pick_clientes).grid(row=row, column=2, padx=PADX, pady=PADY); row += 1

    # Arquivo fornecedores
    tk.Label(root, text="Arquivo FORNECEDORES:").grid(row=row, column=col_lbl, sticky="e", padx=PADX, pady=PADY)
    tk.Entry(root, textvariable=var_fornecedores, width=60).grid(row=row, column=col_inp, padx=PADX, pady=PADY)
    tk.Button(root, text="Procurar...", command=pick_fornecedores).grid(row=row, column=2, padx=PADX, pady=PADY); row += 1

    # Botão executar
    tk.Button(root, text="Executar correções", command=executar).grid(row=row, column=0, columnspan=3, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
