# fornecedores.py
import os
import re
import unicodedata
import warnings
from typing import Dict, List, Tuple
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

# ===================== Utilitários =====================

def _strip_accents(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFD", s)
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")

def _only_digits(s: str) -> str:
    # None ou NaN
    if s is None:
        return ""
    if isinstance(s, float):
        if pd.isna(s):
            return ""
        # Converte float para inteiro sem casas decimais (evita ".0")
        s = f"{s:.0f}"

    s = str(s).strip()

    # Se vier como string "123... .0", remove o sufixo ".0"
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    # (Opcional) tratar notação científica em string, ex: "2.0579080544e10"
    if re.fullmatch(r"[+-]?\d+(\.\d+)?([eE][+-]?\d+)?", s):
        try:
            from decimal import Decimal
            d = Decimal(s)
            s = format(d.quantize(0), "f")  # sem decimais
        except Exception:
            pass

    return re.sub(r"\D", "", s)

def _mask_cpf(num11: str) -> str:
    return f"{num11[0:3]}.{num11[3:6]}.{num11[6:9]}-{num11[9:11]}"

def _mask_cnpj(num14: str) -> str:
    return f"{num14[0:2]}.{num14[2:5]}.{num14[5:8]}/{num14[8:12]}-{num14[12:14]}"

def _mask_cep(num8: str) -> str:
    return f"{num8[0:5]}-{num8[5:8]}"

def _fix_mojibake_pt(s: str) -> str:
    """
    Corrige mojibake comum em PT (ex.: 'VIT¢RIA', 'JUNDIA¡').
    Tenta latin1<->cp1252 e escolhe a melhor heurística.
    """
    if s is None:
        return ""
    s = str(s)
    cands = [s]
    for enc, dec in (("latin1", "cp1252"), ("cp1252", "latin1")):
        try:
            cands.append(s.encode(enc, errors="ignore").decode(dec, errors="ignore"))
        except Exception:
            pass

    def score(t: str):
        bad = sum(t.count(x) for x in ("Ã", "¢", "¡", "Â", "§", "¥"))
        letters = sum(ch.isalpha() for ch in t)
        return (letters - 2 * bad, -len(t))

    return max(cands, key=score)

def _sanitize_text(s: str) -> str:
    # remove acentos e caracteres especiais; normaliza espaços
    s = _strip_accents(_fix_mojibake_pt(s))
    s = re.sub(r"[^A-Za-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _read_excel_any(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path.lower())[1]
    if ext == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd")
        except Exception as e:
            raise RuntimeError("Falha ao ler .xls. Instale xlrd==1.2.0. Erro: " + str(e))
    else:
        return pd.read_excel(path, engine="openpyxl")

def _best_guess_columns(df: pd.DataFrame):
    cols_lower = {c.lower(): c for c in df.columns}

    def find_any(cands: List[str]):
        for c in cands:
            if c.lower() in cols_lower:
                return cols_lower[c.lower()]
        for c in df.columns:
            cl = c.lower()
            if any(x in cl for x in [cand.lower() for cand in cands]):
                return c
        return None

    colmap = {
        "uf": find_any(["uf", "estado", "sigla_uf"]),
        "cpf_cnpj": find_any(["cpf_cnpj", "cpfcnpj", "documento", "doc", "cnpjcpf"]),
        "cep": find_any(["cep", "codigo_postal"]),
        "cidade": find_any(["cidade", "municipio", "municipío", "município"]),
        "codigo_cidade": find_any(["codigo_cidade", "cod_cidade", "codigo_municipio", "cod_municipio", "ibge"]),
        "nome": find_any(["nome", "razao", "razao_social", "razão_social"]),
        "fantasia": find_any(["fantasia", "apelido", "nome_fantasia"]),
        "endereco": find_any(["endereco", "endereço", "logradouro"]),
        "bairro": find_any(["bairro"]),
        "contato": find_any(["contato", "responsavel", "responsável"]),
        "fone": find_any(["fone", "telefone", "telefone1", "tel1"]),
        # coluna de número
        "numero": find_any(["numero", "número", "num"]),
    }

    # Fallbacks posicionais comuns: F (cidade) e G (codigo_cidade)
    if colmap["cidade"] is None and len(df.columns) > 5:
        colmap["cidade"] = df.columns[5]  # F
    if colmap["codigo_cidade"] is None and len(df.columns) > 6:
        colmap["codigo_cidade"] = df.columns[6]  # G
    if colmap["uf"] is None and len(df.columns) > 1:
        colmap["uf"] = df.columns[1]  # B

    created = []
    for k in ["uf", "cpf_cnpj", "cep", "cidade", "codigo_cidade",
              "nome", "fantasia", "endereco", "bairro", "contato", "fone", "numero"]:
        if colmap.get(k) is None:
            new_name = k
            base = new_name
            i = 2
            while new_name in df.columns:
                new_name = f"{base}_{i}"
                i += 1
            df[new_name] = ""
            colmap[k] = new_name
            created.append(new_name)

    return colmap, created

def _normalize_city_for_match(s: str) -> str:
    """
    Normaliza nome da cidade para comparação:
    - Corrige mojibake, remove acentos
    - UPPER
    - Remove sufixos de UF: " - MG", "(MG)", ", MG"
    - Remove não alfanum (mantém espaço)
    - Normaliza espaços
    """
    if s is None:
        return ""
    t = _fix_mojibake_pt(str(s))
    t = _strip_accents(t).upper().strip()
    t = re.sub(r"\s*[-,]?\s*\([A-Z]{2}\)$", "", t)   # "(MG)"
    t = re.sub(r"\s*[-,]?\s+[A-Z]{2}$", "", t)       # " - MG" ou ", MG"
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def _normalize_city(s: str) -> str:
    return _strip_accents(s).upper().strip()

def _smart_cpf_cnpj_mask(raw: str) -> str:
    digits = _only_digits(raw)
    if len(digits) == 0:
        return ""
    if len(digits) == 11:
        return _mask_cpf(digits)
    if len(digits) == 14:
        return _mask_cnpj(digits)
    if 8 < len(digits) < 11:
        return _mask_cpf(digits.zfill(11))
    if 11 < len(digits) < 14:
        return _mask_cnpj(digits.zfill(14))
    return digits

def _smart_cep_mask(raw: str, cep_padrao: str) -> str:
    digits = _only_digits(raw)
    if len(digits) == 0:
        digits = _only_digits(cep_padrao)
    if len(digits) < 8:
        digits = digits.zfill(8)
    elif len(digits) > 8:
        digits = digits[-8:]
    return _mask_cep(digits)

def _format_phone_with_ddd(ddd: str, local: str) -> str:
    ddd = _only_digits(ddd).zfill(2)[-2:]
    if len(local) == 8:
        return f"({ddd}) {local[:4]}-{local[4:]}"
    elif len(local) == 9:
        return f"({ddd}) {local[:5]}-{local[5:]}"
    return local

def _smart_phone_mask(raw: str, ddd_default: str) -> str:
    digits = _only_digits(raw)
    if not digits:
        return ""
    if len(digits) > 11:
        digits = digits[-11:]
    if len(digits) < 7:
        return ""  # decide deixar vazio se muito curto
    if len(digits) in (8, 9):
        return _format_phone_with_ddd(ddd_default, digits)
    if len(digits) == 10:
        ddd, local = digits[:2], digits[2:]
        return f"({ddd}) {local[:4]}-{local[4:]}"
    if len(digits) == 11:
        ddd, local = digits[:2], digits[2:]
        return f"({ddd}) {local[:5]}-{local[5:]}"
    return digits

def _extract_num_from_endereco(endereco: str) -> str:
    """
    Extrai o número do FINAL do endereço ou 'SN' (sem número).
    Aceita variações: 'S N', 'SN', 'S/N', com ou sem pontuação/espacos finais.
    """
    if not endereco:
        return ""
    s = str(endereco).strip().upper()
    s = re.sub(r"[.,;:\s]+$", "", s)
    if re.search(r"\bS[\/\s]*N\.?$", s):
        return "SN"
    m = re.search(r"(\d+)$", s)
    if m:
        return m.group(1)
    return ""

# ===================== Municípios (autodetect) =====================

from typing import Tuple  # no topo do arquivo, se ainda não tiver

def _load_municipios_auto(muni_path: str) -> Tuple[Dict[str, str], Dict[str, str], Dict[str, str], Dict[str, str]]:
    """
    Lê CSV/Excel de municípios.
    Preferências de colunas (com autodetecção fallback):
      - Nome: D (índice 3) ou coluna mais 'alfabética'
      - IBGE: B (índice 1) ou coluna com maioria 7 dígitos
      - UF:   E (índice 4) ou coluna com maioria de siglas (2 letras)
    Retorna 4 dicionários:
      - name_to_code: NOME_NORMALIZADO -> IBGE
      - code_to_name: IBGE -> Nome Title Case
      - code_to_uf:   IBGE -> UF (ex.: 'GO')
      - name_to_uf:   NOME_NORMALIZADO -> UF
    """
    if not os.path.isfile(muni_path):
        raise FileNotFoundError(f"Arquivo municipios não encontrado: {muni_path}")

    ext = os.path.splitext(muni_path.lower())[1]
    df = None
    if ext in [".xls", ".xlsx", ".xlsm"]:
        df = pd.read_excel(muni_path, header=None, dtype=str, engine="openpyxl")
    else:
        for sep in (",", ";"):
            try:
                tmp = pd.read_csv(muni_path, header=None, sep=sep, dtype=str, encoding="utf-8", engine="python")
                if tmp.shape[1] >= 2:
                    df = tmp
                    break
            except Exception:
                continue
    if df is None:
        raise ValueError("Não foi possível ler o arquivo de municípios.")

    # Nome: tenta D(3); senão, coluna mais 'alfabética'
    if df.shape[1] > 3:
        nome_series = df.iloc[:, 3]
    else:
        scores = []
        for i in range(min(df.shape[1], 8)):
            col = df.iloc[:, i].fillna("").astype(str)
            alpha = col.str.count(r"[A-Za-z]").sum()
            digits = col.str.count(r"\d").sum()
            scores.append((alpha - digits, i))
        scores.sort(reverse=True)
        nome_series = df.iloc[:, scores[0][1]]
    nome_col = nome_series.fillna("").astype(str)

    def _as_digits(s: pd.Series) -> pd.Series:
        return s.fillna("").astype(str).map(lambda x: _only_digits(x))

    # IBGE: preferir B(1); senão coluna com maioria de 7 dígitos
    ibge_idx, ibge_hits = (1 if df.shape[1] > 1 else None), -1
    if ibge_idx is not None:
        col = _as_digits(df.iloc[:, ibge_idx])
        ibge_hits = (col.str.match(r"^\d{7}$")).sum()

    best_idx, best_hits = None, -1
    for i in range(min(df.shape[1], 12)):
        if i == nome_series.name:
            continue
        col = _as_digits(df.iloc[:, i])
        hits = (col.str.match(r"^\d{7}$")).sum()
        if hits > best_hits:
            best_hits, best_idx = hits, i
    if best_hits > ibge_hits:
        ibge_idx = best_idx

    if ibge_idx is None:
        raise ValueError("Não foi possível identificar a coluna do IBGE.")
    ibge_col = _as_digits(df.iloc[:, ibge_idx])

    # UF: preferir E(4); senão coluna com maioria de siglas 2 letras
    def _is_uf_series(s: pd.Series) -> int:
        s = s.fillna("").astype(str).str.strip().str.upper()
        return (s.str.match(r"^[A-Z]{2}$")).sum()

    uf_idx, uf_hits = (4 if df.shape[1] > 4 else None), -1
    if uf_idx is not None:
        uf_hits = _is_uf_series(df.iloc[:, uf_idx])

    best_uf_idx, best_uf_hits = None, -1
    for i in range(min(df.shape[1], 12)):
        if i in (nome_series.name, ibge_idx):
            continue
        hits = _is_uf_series(df.iloc[:, i])
        if hits > best_uf_hits:
            best_uf_hits, best_uf_idx = hits, i
    if best_uf_hits > uf_hits:
        uf_idx = best_uf_idx

    uf_col = df.iloc[:, uf_idx].fillna("").astype(str).str.strip().str.upper() if uf_idx is not None else pd.Series([""]*len(df))

    name_to_code, code_to_name, code_to_uf, name_to_uf = {}, {}, {}, {}
    for nome, cod, uf in zip(nome_col, ibge_col, uf_col):
        nome_norm = _normalize_city(nome)
        cod_norm = cod.strip()
        uf_norm  = (uf or "").strip().upper()
        if nome_norm and re.fullmatch(r"\d{7}", cod_norm):
            name_to_code[nome_norm] = cod_norm
            code_to_name[cod_norm] = str(nome).strip().title()
            if uf_norm and re.fullmatch(r"^[A-Z]{2}$", uf_norm):
                code_to_uf[cod_norm] = uf_norm
                name_to_uf[nome_norm] = uf_norm

    return name_to_code, code_to_name, code_to_uf, name_to_uf


# ===================== Núcleo (impl) =====================

def _processar_fornecedores_impl(
    *,
    fornecedores_path: str,
    municipios_path: str,
    cidade_default: str,
    uf_default: str,
    ibge_default: str,
    cep_default: str,
    ddd_default: str,
    modo: str,  # "cidade_para_codigo" | "codigo_para_cidade"
) -> str:
    """
    Implementação interna com 'modo':
      - "cidade_para_codigo": preenche IBGE a partir da cidade (original)
      - "codigo_para_cidade": se cidade vazia e código presente, preenche cidade a partir do IBGE
    """
    df = _read_excel_any(fornecedores_path)
    colmap, created_cols = _best_guess_columns(df)

    name_to_code, code_to_name, code_to_uf, name_to_uf = _load_municipios_auto(municipios_path)


    uf_default = (uf_default or "").strip().upper()
    cidade_default_norm = _normalize_city_for_match(cidade_default or "")
    cep_default_digits = _only_digits(cep_default or "")
    if len(cep_default_digits) < 8:
        cep_default_digits = cep_default_digits.zfill(8)
    cep_default_masked = _mask_cep(cep_default_digits)
    ddd_default = (_only_digits(ddd_default) or "00")[-2:]

    for idx in range(len(df)):
        # CPF/CNPJ
        val_doc = df.at[idx, colmap["cpf_cnpj"]]
        if pd.notna(val_doc) and str(val_doc).strip() != "":
            new_doc = _smart_cpf_cnpj_mask(val_doc)
            df.at[idx, colmap["cpf_cnpj"]] = new_doc or ""
        else:
            df.at[idx, colmap["cpf_cnpj"]] = ""

        # UF (prioridade: pelo código -> pelo nome -> default)
        val_uf = df.at[idx, colmap["uf"]]
        if pd.isna(val_uf) or str(val_uf).strip() == "":
            filled = False
            cur_code = _only_digits(df.at[idx, colmap["codigo_cidade"]])
            if cur_code and cur_code in code_to_uf:
                df.at[idx, colmap["uf"]] = code_to_uf[cur_code]
                filled = True
            if not filled:
                cidade_cell = df.at[idx, colmap["cidade"]]
                uf_by_name = name_to_uf.get(_normalize_city_for_match(cidade_cell), "")
                if uf_by_name:
                    df.at[idx, colmap["uf"]] = uf_by_name
                    filled = True
            if not filled and uf_default:
                df.at[idx, colmap["uf"]] = uf_default


        # CEP
        val_cep = df.at[idx, colmap["cep"]]
        if pd.isna(val_cep) or str(val_cep).strip() == "":
            df.at[idx, colmap["cep"]] = cep_default_masked
        else:
            df.at[idx, colmap["cep"]] = _smart_cep_mask(str(val_cep), cep_default_masked)

        # ----- CIDADE / IBGE conforme modo -----
        val_cid    = df.at[idx, colmap["cidade"]]
        val_codcid = df.at[idx, colmap["codigo_cidade"]]

        if modo == "codigo_para_cidade":
            # 1) Se cidade vazia, tenta pelo código IBGE
            if pd.isna(val_cid) or str(val_cid).strip() == "":
                cod_norm = _only_digits(val_codcid) if pd.notna(val_codcid) else ""
                if cod_norm and cod_norm in code_to_name:
                    df.at[idx, colmap["cidade"]] = code_to_name[cod_norm]
                elif cidade_default_norm:
                    df.at[idx, colmap["cidade"]] = cidade_default_norm.title()

            # 2) Garante código, se ainda vazio, a partir da cidade
            val_codcid = df.at[idx, colmap["codigo_cidade"]]
            if pd.isna(val_codcid) or str(val_codcid).strip() == "":
                cidade_cell = df.at[idx, colmap["cidade"]]
                key = _normalize_city_for_match(cidade_cell)
                ibge_code = name_to_code.get(key, "")
                if not ibge_code and ibge_default:
                    ibge_code = ibge_default
                if ibge_code:
                    df.at[idx, colmap["codigo_cidade"]] = ibge_code

        else:  # "cidade_para_codigo" (padrão)
            # 1) Se cidade vazia, usa default (se houver)
            if pd.isna(val_cid) or str(val_cid).strip() == "":
                if cidade_default_norm:
                    df.at[idx, colmap["cidade"]] = cidade_default_norm.title()
            else:
                # Corrige mojibake visível
                fixed_city = _fix_mojibake_pt(str(val_cid))
                df.at[idx, colmap["cidade"]] = fixed_city

            # 2) Código a partir da cidade
            val_codcid = df.at[idx, colmap["codigo_cidade"]]
            if pd.isna(val_codcid) or str(val_codcid).strip() == "":
                cidade_cell = df.at[idx, colmap["cidade"]]
                key = _normalize_city_for_match(cidade_cell)
                ibge_code = name_to_code.get(key, "")
                if not ibge_code and ibge_default:
                    ibge_code = ibge_default
                if ibge_code:
                    df.at[idx, colmap["codigo_cidade"]] = ibge_code

        # Número do endereço
        col_num = colmap["numero"]
        col_ender = colmap["endereco"]
        val_num = df.at[idx, col_num]
        if pd.isna(val_num) or str(val_num).strip() == "":
            val_end = df.at[idx, col_ender]
            num_extraido = _extract_num_from_endereco(val_end if pd.notna(val_end) else "")
            if num_extraido:
                df.at[idx, col_num] = num_extraido
                novo_end = re.sub(r"(\d+|S[\/\s]*N)\s*$", "", str(val_end).upper(), flags=re.IGNORECASE)
                df.at[idx, col_ender] = novo_end.strip()

        # Texto: nome, fantasia, endereco, bairro, contato
        for key in ["nome", "fantasia", "endereco", "bairro", "contato"]:
            col = colmap[key]
            val = df.at[idx, col]
            if pd.notna(val) and str(val).strip() != "":
                df.at[idx, col] = _sanitize_text(val)

        # Telefone
        col = colmap["fone"]
        val = df.at[idx, col]
        if pd.notna(val) and str(val).strip() != "":
            df.at[idx, col] = _smart_phone_mask(str(val), ddd_default)
        else:
            df.at[idx, col] = ""

    # Saída
    base_dir = os.path.dirname(os.path.abspath(fornecedores_path))
    out_xlsx = os.path.join(base_dir, "fornecedores_corrigido.xlsx")
    df.to_excel(out_xlsx, index=False, engine="openpyxl")

    # Log simples
    if created_cols:
        out_log = os.path.join(base_dir, "fornecedores_corrigido_log.csv")
        pd.DataFrame({"colunas_criadas": created_cols}).to_csv(out_log, index=False, encoding="utf-8")

    return out_xlsx

# ===================== API pública =====================

def processar_fornecedores_por_cidade(
    *,
    fornecedores_path: str,
    municipios_path: str,
    cidade_default: str,
    uf_default: str,
    ibge_default: str,
    cep_default: str,
    ddd_default: str,
) -> str:
    """
    Modo 'Cidade → IBGE' (comportamento original).
    """
    return _processar_fornecedores_impl(
        fornecedores_path=fornecedores_path,
        municipios_path=municipios_path,
        cidade_default=cidade_default,
        uf_default=uf_default,
        ibge_default=ibge_default,
        cep_default=cep_default,
        ddd_default=ddd_default,
        modo="cidade_para_codigo",
    )

def processar_fornecedores_por_codigo(
    *,
    fornecedores_path: str,
    municipios_path: str,
    cidade_default: str,
    uf_default: str,
    ibge_default: str,
    cep_default: str,
    ddd_default: str,
) -> str:
    """
    Modo 'Código IBGE → Cidade' (se cidade estiver vazia e houver código).
    """
    return _processar_fornecedores_impl(
        fornecedores_path=fornecedores_path,
        municipios_path=municipios_path,
        cidade_default=cidade_default,
        uf_default=uf_default,
        ibge_default=ibge_default,
        cep_default=cep_default,
        ddd_default=ddd_default,
        modo="codigo_para_cidade",
    )

# Compatibilidade com o main.py atual
def processar_fornecedores(
    *,
    fornecedores_path: str,
    municipios_path: str,
    cidade_default: str,
    uf_default: str,
    ibge_default: str,
    cep_default: str,
    ddd_default: str,
) -> str:
    """
    Alias para o modo padrão (Cidade → IBGE).
    """
    return processar_fornecedores_por_cidade(
        fornecedores_path=fornecedores_path,
        municipios_path=municipios_path,
        cidade_default=cidade_default,
        uf_default=uf_default,
        ibge_default=ibge_default,
        cep_default=cep_default,
        ddd_default=ddd_default,
    )
