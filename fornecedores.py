# fornecedores.py
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

def _sanitize_text(s: str) -> str:
    # remove acentos e caracteres especiais; normaliza espaços
    s = _strip_accents(_fix_mojibake_pt(s))
    s = re.sub(r"[^A-Za-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _fix_mojibake_pt(s: str) -> str:
    """
    Tenta corrigir 'mojibake' comum em PT (ex.: 'VIT¢RIA', 'JUNDIA¡').
    Faz tentativas de reinterpretação latin1<->cp1252 e escolhe a melhor.
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
        # NOVO: coluna de número
        "numero": find_any(["numero", "número", "num"]),
    }

    # Fallbacks posicionais comuns: F (cidade) e G (codigo_cidade)
    if colmap["cidade"] is None and len(df.columns) > 5:
        colmap["cidade"] = df.columns[5]  # F
    if colmap["codigo_cidade"] is None and len(df.columns) > 6:
        colmap["codigo_cidade"] = df.columns[6]  # G

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

def _load_municipios_any(muni_path: str) -> Dict[str, str]:
    """
    Carrega municípios de CSV ou Excel.
      - Coluna D (índice 3): nome do município (para comparar)
      - Coluna B (índice 1): código IBGE (valor a retornar para fornecedores)
    Aceita: .csv, .xls, .xlsx, .xlsm
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
                if tmp.shape[1] >= 4:
                    df = tmp
                    break
            except Exception:
                continue

    if df is None or df.shape[1] < 4:
        raise ValueError("O arquivo de municípios precisa ter pelo menos colunas B e D.")

    nome_col = df.iloc[:, 3].fillna("").astype(str)  # D
    cod_col  = df.iloc[:, 1].fillna("").astype(str)  # B

    mapping = {}
    for nome, cod in zip(nome_col, cod_col):
        key = _normalize_city_for_match(nome)
        # IBGE: garantir somente dígitos e manter zero à esquerda
        cod_digits = re.sub(r"\D", "", str(cod))
        # (Opcional) padronizar com 7 dígitos
        if cod_digits:
            cod_digits = cod_digits.zfill(7)
        if key and cod_digits:
            mapping[key] = cod_digits
    return mapping

def _normalize_city_for_match(s: str) -> str:
    """
    Normaliza nome da cidade para comparação:
    - Corrige mojibake, remove acentos
    - Converte para UPPER
    - Remove sufixos de UF: " - MG", "(MG)", ", MG"
    - Remove caracteres não alfanuméricos (mantém espaço)
    - Normaliza espaços
    """
    if s is None:
        return ""
    t = _fix_mojibake_pt(str(s))
    t = _strip_accents(t).upper().strip()

    # Remove UF no final: " - MG", "(MG)", ", MG"
    t = re.sub(r"\s*[-,]?\s*\([A-Z]{2}\)$", "", t)         # "(MG)"
    t = re.sub(r"\s*[-,]?\s+[A-Z]{2}$", "", t)              # " - MG" ou ", MG"

    # Remove tudo que não for letra/número/espaço
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

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

    # Se não encaixar em nada → devolve os dígitos crus
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
    """
    - 10 dígitos: (XX) XXXX-XXXX
    - 11 dígitos: (XX) XXXXX-XXXX
    - 8 dígitos : (YY) XXXX-XXXX
    - 9 dígitos : (YY) XXXXX-XXXX
    - < 7 dígitos: retorna "" (vazio)
    Regras adicionais:
    - Remove não dígitos; se >11, usa últimos 11.
    """
    digits = _only_digits(raw)
    if not digits:
        return ""

    if len(digits) > 11:
        digits = digits[-11:]

    if len(digits) < 7:
        return ""  # conforme solicitado

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
    # remove pontuação/espaços do fim
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

def _normalize_city(s: str) -> str:
    return _strip_accents(s).upper().strip()

def processar_fornecedores(
    fornecedores_path: str,
    municipios_path: str,
    cidade_default: str,
    uf_default: str,
    ibge_default: str,  # usado como fallback se não achar no municípios
    cep_default: str,
    ddd_default: str,
) -> str:
    """
    Processa a planilha fornecedores.xls/.xlsx conforme regras especificadas.
    - codigo_cidade: compara cidade (planilha) com coluna D do municipios.* e retorna código da coluna B.
    - aceita municipios.* em CSV ou Excel.
    Retorna caminho do arquivo de saída (fornecedores_corrigido.xlsx).
    """
    df = _read_excel_any(fornecedores_path)
    colmap, created_cols = _best_guess_columns(df)

    mun_map = _load_municipios_any(municipios_path)

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
            if new_doc:
                df.at[idx, colmap["cpf_cnpj"]] = new_doc
        else:
            df.at[idx, colmap["cpf_cnpj"]] = ""  # garante string vazia

        # UF
        val_uf = df.at[idx, colmap["uf"]]
        if pd.isna(val_uf) or str(val_uf).strip() == "":
            if uf_default:
                df.at[idx, colmap["uf"]] = uf_default

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
        else:
            fixed_city = _fix_mojibake_pt(str(val_cid))
            df.at[idx, colmap["cidade"]] = fixed_city

        # Código da cidade (IBGE)
        val_codcid = df.at[idx, colmap["codigo_cidade"]]
        if pd.isna(val_codcid) or str(val_codcid).strip() == "":
            cidade_cell = df.at[idx, colmap["cidade"]]
            key = _normalize_city(str(cidade_cell))
            ibge_code = mun_map.get(key, "")
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
            df.at[idx, col] = ""  # garante string vazia

    # Saída
    base_dir = os.path.dirname(os.path.abspath(fornecedores_path))
    out_xlsx = os.path.join(base_dir, "fornecedores_corrigido.xlsx")
    df.to_excel(out_xlsx, index=False, engine="openpyxl")

    # Log simples
    if created_cols:
        out_log = os.path.join(base_dir, "fornecedores_corrigido_log.csv")
        pd.DataFrame({"colunas_criadas": created_cols}).to_csv(out_log, index=False, encoding="utf-8")

    return out_xlsx
