import pandas as pd
import re

START_DATA_INDEX = 0
_CTRL_RE = re.compile(r"[\x00-\x1F\x7F]")

# =========================
# 清洗工具
# =========================
def clean_str(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v)
    s = s.replace("\t", "").replace("\n", "").replace("\r", "")
    s = _CTRL_RE.sub("", s)
    return s.strip()

def clean_header_name(v):
    return clean_str(v).upper()

def normalize_key_value(col_name: str, value):
    s = clean_str(value)
    if s == "":
        return ""
    col = clean_header_name(col_name)
    if col == "VORNR":
        digits = re.sub(r"\D", "", s)
        return digits.zfill(4) if digits else ""
    return s

# =========================
# Key / 比對核心
# =========================
def make_key(df: pd.DataFrame, row_idx: int, key_cols):
    out = []
    for c in key_cols:
        col_name = str(df.columns[c])
        raw_val = df.iat[row_idx, c]
        out.append(normalize_key_value(col_name, raw_val))
    return tuple(out)

def build_key_map(df: pd.DataFrame, key_cols):
    m = {}
    for i in range(START_DATA_INDEX, len(df)):
        k = make_key(df, i, key_cols)
        if k not in m:
            m[k] = i
    return m

def count_duplicate_keys(df: pd.DataFrame, key_cols):
    keys = [make_key(df, i, key_cols) for i in range(START_DATA_INDEX, len(df))]
    s = pd.Series(keys)
    return int(s.duplicated(keep=False).sum())

def build_clean_col_map(df: pd.DataFrame):
    mp = {}
    for c in df.columns:
        cc = clean_header_name(c)
        mp.setdefault(cc, []).append(c)
    return mp

def diff_directional(df_base, df_other, map_base, map_other, key_cols_base, base_name, other_name):
    keys_base = set(map_base.keys())
    keys_other = set(map_other.keys())

    base_cols_map = build_clean_col_map(df_base)
    other_cols_map = build_clean_col_map(df_other)

    base_key_clean = {clean_header_name(df_base.columns[i]) for i in key_cols_base}
    comparable_clean_cols = sorted(
        (set(base_cols_map.keys()) & set(other_cols_map.keys())) - base_key_clean
    )

    rows = []
    diff_cells = 0
    missing_keys = 0

    for k in sorted(keys_base - keys_other):
        missing_keys += 1
        rows.append(list(k) + ["-", f"{base_name}=存在", f"{other_name}=不存在", "Key不存在"])

    for k in sorted(keys_base & keys_other):
        ib = map_base[k]
        io = map_other[k]
        for cc in comparable_clean_cols:
            col_b = base_cols_map[cc][0]
            col_o = other_cols_map[cc][0]
            vb = clean_str(df_base.at[ib, col_b])
            vo = clean_str(df_other.at[io, col_o])
            if vb != vo:
                diff_cells += 1
                rows.append(list(k) + [str(col_b), vb, vo, "值不同"])

    return rows, diff_cells, missing_keys, len(comparable_clean_cols)

def build_column_diff(df_a: pd.DataFrame, df_b: pd.DataFrame):
    a_map = build_clean_col_map(df_a)
    b_map = build_clean_col_map(df_b)

    rows = []

    for cc in sorted(set(a_map) - set(b_map)):
        rows.append([cc, a_map[cc][0], "", "A有B沒有"])

    for cc in sorted(set(b_map) - set(a_map)):
        rows.append([cc, "", b_map[cc][0], "B有A沒有"])

    for cc in sorted(set(a_map) & set(b_map)):
        a_real = a_map[cc][0]
        b_real = b_map[cc][0]
        status = "一致" if a_real == b_real else "清洗後一致（原欄名不同）"
        rows.append([cc, a_real, b_real, status])

    return pd.DataFrame(
        rows,
        columns=["欄位(清洗後)", "Excel_A欄名", "Excel_B欄名", "狀態"]
    )
