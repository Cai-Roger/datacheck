import pandas as pd

# =========================
# Strict mode helpers
# =========================

def normalize_raw_value(v):
    """
    嚴格模式（顯示用 & 比對用）：
    - None / pandas NaN → 空字串
    - 其他 → 原樣轉成字串
    """
    if v is None or pd.isna(v):
        return ""
    return str(v)


def values_equal_strict(a, b) -> bool:
    """
    嚴格比對（包含空白 / 換行 / tab）
    但 NaN / None 視為空字串
    """
    return normalize_raw_value(a) == normalize_raw_value(b)


# =========================
# Header helpers
# =========================

def clean_header_name(name) -> str:
    """
    僅用於 Key 預設推測，不影響資料內容
    """
    if name is None:
        return ""
    s = str(name).strip().upper()
    s = s.replace("\u3000", "").replace(" ", "").replace("\t", "")
    return s


# =========================
# Key helpers
# =========================

def normalize_key_value(v):
    """
    Key 欄位允許寬鬆處理（去前後空白）
    """
    if v is None or pd.isna(v):
        return ""
    return str(v).strip()


def make_key_tuple(row, key_cols):
    return tuple(normalize_key_value(row.iloc[i]) for i in key_cols)


def build_key_map(df: pd.DataFrame, key_cols: list[int]):
    """
    回傳 dict: key_tuple -> list[row_index]
    """
    key_map = {}
    for idx, row in df.iterrows():
        k = make_key_tuple(row, key_cols)
        key_map.setdefault(k, []).append(idx)
    return key_map


def count_duplicate_keys(df: pd.DataFrame, key_cols: list[int]) -> int:
    """
    回傳重複 key 的列數（不含第一筆）
    """
    m = build_key_map(df, key_cols)
    dup = 0
    for idxs in m.values():
        if len(idxs) > 1:
            dup += (len(idxs) - 1)
    return dup


# =========================
# Column diff
# =========================

def build_column_diff(df_a: pd.DataFrame, df_b: pd.DataFrame) -> pd.DataFrame:
    a_cols = list(df_a.columns)
    b_cols = list(df_b.columns)

    a_set = set(a_cols)
    b_set = set(b_cols)

    rows = []
    for c in a_cols:
        rows.append({
            "欄位": c,
            "A有": True,
            "B有": c in b_set,
            "狀態": "OK" if c in b_set else "B缺少"
        })
    for c in b_cols:
        if c not in a_set:
            rows.append({
                "欄位": c,
                "A有": False,
                "B有": True,
                "狀態": "A缺少"
            })

    return pd.DataFrame(rows)


# =========================
# Directional diff (Strict)
# =========================

def diff_directional(
    df_src: pd.DataFrame,
    df_tgt: pd.DataFrame,
    map_src: dict,
    map_tgt: dict,
    key_cols_src: list[int],
    src_label: str,  # "A" or "B"
    tgt_label: str   # "B" or "A"
):
    """
    從 src 角度比對到 tgt：
    - Key 不存在 → 一筆差異
    - Key 存在 → 逐欄位嚴格比對
    """

    common_cols = [c for c in df_src.columns if c in df_tgt.columns]
    key_names = [df_src.columns[i] for i in key_cols_src]
    compare_cols = [c for c in common_cols if c not in key_names]

    rows = []
    missing_keys = []
    matched_keys = 0
    diff_count = 0

    for _, row_src in df_src.iterrows():
        key_t = make_key_tuple(row_src, key_cols_src)
        key_out = list(key_t)

        # Key 不存在
        if key_t not in map_tgt:
            missing_keys.append(key_t)
            rows.append(
                key_out + [
                    "(Key不存在)",
                    f"存在於{src_label}",
                    f"不存在於{tgt_label}",
                    f"{src_label}→{tgt_label}"
                ]
            )
            diff_count += 1
            continue

        matched_keys += 1
        tgt_idx = map_tgt[key_t][0]
        row_tgt = df_tgt.loc[tgt_idx]

        for col in compare_cols:
            a_val = row_src[col]
            b_val = row_tgt[col]

            if not values_equal_strict(a_val, b_val):
                a_disp = normalize_raw_value(a_val)
                b_disp = normalize_raw_value(b_val)

                rows.append(
                    key_out + [
                        col,
                        a_disp if src_label == "A" else b_disp,
                        b_disp if src_label == "A" else a_disp,
                        f"{src_label}→{tgt_label}"
                    ]
                )
                diff_count += 1

    return rows, missing_keys, matched_keys, diff_count
