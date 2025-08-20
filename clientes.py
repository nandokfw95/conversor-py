# clientes.py
import os
import re
import unicodedata
import warnings
from typing import Dict, List
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
            from decimal import Decimal, InvalidOperation
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

def _sanitize_text(s: str) -> str:
    # mesma tática do razao_nome: remove acentos e caracteres especiais
    s = _strip_accents(s)
    s = re.sub(r"[^A-Za-z0-9 ]+", " ", s)  # mantém apenas alfanum e espaço
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
        "razao_nome": find_any(["razao_nome", "razão_nome", "razao", "razão", "razao social", "razao_social", "nome", "nome_razao"]),
        # Campos texto extra
        "endereco": find_any(["endereco", "endereço", "logradouro"]),
        "bairro": find_any(["bairro"]),
        "complemento": find_any(["complemento", "compl"]),
        "contato": find_any(["contato", "responsavel", "responsável"]),
        "fantasia_apelido": find_any(["fantasia_apelido", "fantasia", "apelido", "nome_fantasia"]),
        # Telefones
        "tel_principal": find_any(["tel_principal", "telefone", "telefone1", "tel1", "fone", "fone1"]),
        "tel_comercial": find_any(["tel_comercial", "telefone2", "tel2", "fone2"]),
        # NOVO: coluna do número do endereço
        "numero": find_any(["numero", "número", "num"]),
    }

    # Fallbacks posicionais F/G (se fizer sentido para seu layout)
    if colmap["cidade"] is None and len(df.columns) > 5:
        colmap["cidade"] = df.columns[5]  # F
    if colmap["codigo_cidade"] is None and len(df.columns) > 6:
        colmap["codigo_cidade"] = df.columns[6]  # G

    # Garante colunas, criando se faltarem
    created = []
    for k in ["uf", "cpf_cnpj", "cep", "cidade", "codigo_cidade",
              "razao_nome", "endereco", "bairro", "complemento", "contato", "fantasia_apelido",
              "tel_principal", "tel_comercial", "numero"]:  # inclui 'numero'
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

def _load_municipios(muni_path: str) -> Dict[str, str]:
    """
    Carrega municípios de CSV ou Excel:
      - Coluna D (índice 3): nome do município
      - Coluna E (índice 4): código IBGE
    Aceita: .csv, .xls, .xlsx, .xlsm
    """
    if not os.path.isfile(muni_path):
        raise FileNotFoundError(f"Arquivo municipios não encontrado: {muni_path}")

    ext = os.path.splitext(muni_path.lower())[1]

    df = None
    if ext in [".xls", ".xlsx", ".xlsm"]:
        # Excel
        df = pd.read_excel(muni_path, header=None, dtype=str, engine="openpyxl")
    else:
        # CSV (tenta , e ;)
        for sep in (",", ";"):
            try:
                tmp = pd.read_csv(muni_path, header=None, sep=sep, dtype=str, encoding="utf-8", engine="python")
                if tmp.shape[1] >= 5:
                    df = tmp
                    break
            except Exception:
                continue

    if df is None or df.shape[1] < 5:
        raise ValueError("O arquivo de municípios deve ter pelo menos 5 colunas (D e E).")

    nome_col = df.iloc[:, 3].fillna("").astype(str)
    cod_col = df.iloc[:, 1].fillna("").astype(str)

    mapping = {}
    for nome, cod in zip(nome_col, cod_col):
        key = _strip_accents(nome).upper().strip()
        if key:
            mapping[key] = cod.strip()
    return mapping

def _normalize_city(s: str) -> str:
    return _strip_accents(s).upper().strip()

def _smart_cpf_cnpj_mask(raw: str) -> str:
    digits = _only_digits(raw)
    if len(digits) == 0:
        return ""

    # CPF exato
    if len(digits) == 11:
        return _mask_cpf(digits)

    # CNPJ exato
    if len(digits) == 14:
        return _mask_cnpj(digits)

    # Se tiver entre 9 e 10 → completa até 11 → CPF
    if 8 < len(digits) < 11:
        digits = digits.zfill(11)
        return _mask_cpf(digits)

    # Se tiver entre 12 e 13 → completa até 14 → CNPJ
    if 11 < len(digits) < 14:
        digits = digits.zfill(14)
        return _mask_cnpj(digits)

    # Caso não encaixe em nada → retorna apenas os dígitos
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
    """Formata com DDD informado e número local (8 ou 9 dígitos)."""
    ddd = _only_digits(ddd).zfill(2)[-2:]
    if len(local) == 8:
        return f"({ddd}){local[:4]}-{local[4:]}"
    elif len(local) == 9:
        return f"({ddd}){local[:5]}-{local[5:]}"
    else:
        return local  # fallback

def _smart_phone_mask(raw: str, ddd_default: str) -> str:
    """
    Regras:
    - 10 dígitos: (XX) XXXX-XXXX
    - 11 dígitos: (XX) XXXXX-XXXX
    - 8 dígitos : (YY) XXXX-XXXX   (usa DDD padrão)
    - 9 dígitos : (YY) XXXXX-XXXX  (usa DDD padrão)
    Observações:
    - Remove não dígitos.
    - Se >11 dígitos, usa os últimos 11 (descarta prefixos).
    - Se <8 dígitos, retorna como dígitos (sem máscara).
    """
    digits = _only_digits(raw)
    if not digits:
        return ""

    if len(digits) > 11:
        digits = digits[-11:]

    if len(digits) < 8:
        return digits

    if len(digits) in (8, 9):
        return _format_phone_with_ddd(ddd_default, digits)

    if len(digits) == 10:
        ddd, local = digits[:2], digits[2:]
        return f"({ddd}){local[:4]}-{local[4:]}"

    if len(digits) == 11:
        ddd, local = digits[:2], digits[2:]
        return f"({ddd}){local[:5]}-{local[5:]}"

    return digits

def _extract_num_from_endereco(endereco: str) -> str:
    """
    Extrai o número do FINAL do endereço ou 'SN' (sem número).
    Aceita variações: 'S N', 'SN', 'S/N', com ou sem pontuação/espaços finais.
    """
    if not endereco:
        return ""
    s = str(endereco).strip().upper()
    # Remove espaços e pontuações finais
    s = re.sub(r"[.,;:\s]+$", "", s)

    # 'SN' (sem número) no final
    if re.search(r"\bS[\/\s]*N\.?$", s):
        return "SN"

    # número no final (último bloco de dígitos no fim da string)
    m = re.search(r"(\d+)$", s)
    if m:
        return m.group(1)

    return ""

# ===================== API pública =====================

def processar_clientes(
    clientes_path: str,
    municipios_csv_path: str,  # mantém o nome para compatibilidade com main.py
    cidade_default: str,
    uf_default: str,
    ibge_default: str,
    cep_default: str,
    ddd_default: str,  # DDD padrão
) -> str:
    """
    Processa a planilha de clientes conforme regras especificadas.
    Aceita municipios em CSV ou Excel (.xls/.xlsx/.xlsm).
    Retorna o caminho do arquivo de saída (clientes_corrigido.xlsx)
    """
    df = _read_excel_any(clientes_path)
    colmap, created_cols = _best_guess_columns(df)

    # Agora aceita CSV ou XLS/XLSX:
    municipios_map = _load_municipios(municipios_csv_path)

    uf_default = (uf_default or "").strip().upper()
    cidade_default_norm = _normalize_city(cidade_default or "")
    cep_default_digits = _only_digits(cep_default or "")
    if len(cep_default_digits) < 8:
        cep_default_digits = cep_default_digits.zfill(8)
    cep_default_masked = _mask_cep(cep_default_digits)
    ddd_default = (_only_digits(ddd_default) or "00")[-2:]  # garante 2 dígitos

    # Itera linhas
    for idx in range(len(df)):
        # UF
        val_uf = df.at[idx, colmap["uf"]]
        if pd.isna(val_uf) or str(val_uf).strip() == "":
            if uf_default:
                df.at[idx, colmap["uf"]] = uf_default

        # CPF/CNPJ
        val_doc = df.at[idx, colmap["cpf_cnpj"]]
        new_doc = _smart_cpf_cnpj_mask(val_doc)
        if new_doc != (val_doc if pd.notna(val_doc) else ""):
            df.at[idx, colmap["cpf_cnpj"]] = new_doc

        # CEP
        val_cep = df.at[idx, colmap["cep"]]
        if pd.isna(val_cep) or str(val_cep).strip() == "":
            df.at[idx, colmap["cep"]] = cep_default_masked
        else:
            df.at[idx, colmap["cep"]] = _smart_cep_mask(str(val_cep), cep_default_masked)

        # Cidade
        val_cid = df.at[idx, colmap["cidade"]]
        if pd.isna(val_cid) or str(val_cid).strip() == "":
            if cidade_default_norm:
                df.at[idx, colmap["cidade"]] = cidade_default_norm.title()

        # Código IBGE (codigo_cidade)
        val_codcid = df.at[idx, colmap["codigo_cidade"]]
        if pd.isna(val_codcid) or str(val_codcid).strip() == "":
            cidade_cell = df.at[idx, colmap["cidade"]]
            cid_norm = _normalize_city(cidade_cell)
            ibge_code = municipios_map.get(cid_norm, "")
            if not ibge_code and ibge_default:
                ibge_code = ibge_default
            if ibge_code:
                df.at[idx, colmap["codigo_cidade"]] = ibge_code

                # NOVO: extrair número do endereço se 'numero' estiver vazio (antes da sanitização de textos)
        col_num = colmap["numero"]
        col_ender = colmap["endereco"]
        val_num = df.at[idx, col_num]
        if pd.isna(val_num) or str(val_num).strip() == "":
            val_end = df.at[idx, col_ender]
            num_extraido = _extract_num_from_endereco(val_end if pd.notna(val_end) else "")
            if num_extraido:
                # Preenche a Coluna M
                df.at[idx, col_num] = num_extraido
                # Remove o número ou SN do endereço (Coluna K)
                novo_end = re.sub(r"(\d+|S[\/\s]*N)\s*$", "", str(val_end).upper(), flags=re.IGNORECASE)
                df.at[idx, col_ender] = novo_end.strip()


        # Texto: razao_nome + extras
        for key in ["razao_nome", "endereco", "bairro", "complemento", "contato", "fantasia_apelido"]:
            col = colmap[key]
            val = df.at[idx, col]
            if pd.notna(val) and str(val).strip() != "":
                df.at[idx, col] = _sanitize_text(val)

        # Telefones
        for key in ["tel_principal", "tel_comercial"]:
            col = colmap[key]
            val = df.at[idx, col]
            if pd.notna(val) and str(val).strip() != "":
                df.at[idx, col] = _smart_phone_mask(str(val), ddd_default)

    # Saída
    base_dir = os.path.dirname(os.path.abspath(clientes_path))
    out_xlsx = os.path.join(base_dir, "clientes_corrigido.xlsx")
    df.to_excel(out_xlsx, index=False, engine="openpyxl")

    # Log das colunas criadas (se houver)
    if created_cols:
        out_log = os.path.join(base_dir, "clientes_corrigido_log.csv")
        pd.DataFrame({"colunas_criadas": created_cols}).to_csv(out_log, index=False, encoding="utf-8")

    return out_xlsx
