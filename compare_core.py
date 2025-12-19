import pandas as pd

# =========================
# Strict mode helpers
# =========================

NA_MARK = "<NaN>"


def normalize_raw_value(v):
    """
    嚴格模式：保留所有空白/換行/tab，不做 strip/replace
    只處理 NaN/None，避免 NaN vs "" 比對混亂
    """
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return NA_MARK
    return str(v)


def visualize_whitespace(s: str) -> str:
    """
    將不可見字元可視化，方便在結果表看到差異：
    空白 -> ␣
    Tab  -> ⇥
    換行 -> ↵\n  (保留換行，以便看出位置)
    CR   -> ␍
    """
    if s is None:
        return ""
    s = str(s)
    return (
        s.replace("\r", "␍")
         .replace("\t", "⇥")
         .replace(" ", "␣")
         .replace("\n", "↵\n")
    )


def values_equal_strict(a, b) -> bool:
    """
    嚴格比對：逐字元比對（包含空白/換行/tab）
    """
    return normalize_raw_value(a) == normalize_raw_value(b)


# =========================
# Header helpers
# =========================

def clean_header_name(name) -> str:
    """
    表頭清洗用（不影響資料值嚴格比對）
    用於預設 Key 欄位推測：PLNNR/VORNR 等
    """
    if name is None:
        return ""
    s = str(name).strip().upper()
    # 把常見空白符統一移除，避免表頭被奇怪空白影響
    s = s.replace("\u3000", "").replace(" ", "").replace("\t", "")
    return s


# =========================
# Key helpers (keys usually should be tolerant)
# =========================

def normalize_key_value(v):
    """
    Key 欄位建議做「寬鬆」處理：去前後空白
    （Key 主要用來對齊資料；嚴格模式主要針對內容欄位）
    """
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return NA_MARK
    return str(v).strip()


def make_key_tuple(row, key_cols):
    return tuple(normalize_key_value(row.iloc[i]) for i in key_cols)


def build_key_map(df: pd.DataFrame, key_cols: list[int]):
    """
    回傳 dict: key_tuple -> list[row_index]
    讓你能偵測重複 key，也能取第一筆做比對
    """
    key_map = {}
    for idx, row in df.iterrows():
        k = make_key_tuple(row, key_cols)
        key_map.setdefault(k, []).append(idx)
    return key_map


def count_duplicate_keys(df: pd.DataFrame, key_cols: list[int]) -> int:
    """
    回傳「重複 key 的列數（不含第一次）」：
    若某 key 出現 3 次，則重複列數 +2
    """
    m = build_key_map(df, key_cols)
    dup = 0
    for _, idxs in m.items():
        if len(idxs) > 1:
            dup += (len(idxs) - 1)
    return dup


# =========================
# Column diff
# =========================

def build_column_diff(df_a: pd.DataFrame, df_b: pd.DataFrame) -> pd.DataFrame:
    """
    比對兩份表的欄位差異（存在/缺少）
    """
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
    - 若 key 在 tgt 不存在：輸出一筆「Key 不存在」差異
    - 若 key 存在：逐欄位嚴格比對（含空白/換行/tab）
    回傳：
      rows: list[list] 給 app.py 組 DataFrame 用
      missing_keys: list[tuple]
      matched_keys: int
      diff_count: int
    """
    # 共同欄位（只比共同欄位）
    common_cols = [c for c in df_src.columns if c in df_tgt.columns]

    # key 欄位名稱（用來排除 key 本身不比）
    key_names = [df_src.columns[i] for i in key_cols_src]
    compare_cols = [c for c in common_cols if c not in key_names]

    rows = []
    missing_keys = []
    matched_keys = 0
    diff_count = 0

    for _, row_src in df_src.iterrows():
        key_t = make_key_tuple(row_src, key_cols_src)

        key_out = list(key_t)  # KEY_1, KEY_2...

        # tgt 無此 key
        if key_t not in map_tgt:
            missing_keys.append(key_t)
            rows.append(
                key_out
                + ["(Key不存在)", visualize_whitespace("存在於" + src_label), visualize_whitespace("不存在於" + tgt_label), f"{src_label}→{tgt_label}"]
            )
            diff_count += 1
            continue

        matched_keys += 1

        # tgt 有 key：取第一筆對應列（若 tgt 有重複 key，仍以第一筆比對）
        tgt_idx = map_tgt[key_t][0]
        row_tgt = df_tgt.loc[tgt_idx]

        for col in compare_cols:
            a_val = row_src[col]
            b_val = row_tgt[col]

            if not values_equal_strict(a_val, b_val):
                a_norm = normalize_raw_value(a_val)
                b_norm = normalize_raw_value(b_val)

                # 將不可見字元可視化，避免看不出差異
                a_disp = visualize_whitespace(a_norm)
                b_disp = visualize_whitespace(b_norm)

                rows.append(
                    key_out + [col, a_disp if src_label == "A" else b_disp, b_disp if src_label == "A" else a_disp, f"{src_label}→{tgt_label}"]
                )
                diff_count += 1

    return rows, missing_keys, matched_keys, diff_count
