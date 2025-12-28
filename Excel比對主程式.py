import streamlit as st
import pandas as pd
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

from config import APP_NAME, APP_VERSION, APP_FOOTER
from compare_core import (
    clean_header_name,
    build_key_map,
    count_duplicate_keys,
    diff_directional,
    build_column_diff,
)

# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title=f"{APP_NAME}",
    layout="wide",
)

# =========================================================
# 工具
# =========================================================
def now_tw():
    return datetime.now(ZoneInfo("Asia/Taipei"))

def gen_download_filename(base_name: str):
    ts = now_tw().strftime("%Y%m%d_%H%M%S")
    return f"{base_name}_{ts}.xlsx"

def clean_display_value(v):
    """移除可視化空白符號與 <NaN>"""
    if v is None:
        return ""
    s = str(v)
    if s == "<NaN>":
        return ""
    return (
        s.replace("␣", " ")
         .replace("⇥", "")
         .replace("␍", "")
         .replace("↵", "\n")
    )

# =========================================================
# 主畫面
# =========================================================
st.title(f"Excel 比對程式（Web {APP_VERSION}）")

st.markdown("""
### 使用說明
1. 上傳 Excel A、Excel B  
2. 選擇 Key 欄位  
3. 點擊「開始差異比對」  
""")

# =========================================================
# 上傳檔案
# =========================================================
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("📤 上傳 Excel A", type=["xlsx"])
with col2:
    file_b = st.file_uploader("📤 上傳 Excel B", type=["xlsx"])

if not file_a or not file_b:
    st.info("請先上傳兩份 Excel")
    st.stop()

df_a = pd.read_excel(file_a)
df_b = pd.read_excel(file_b)

st.success(f"A：{df_a.shape[0]} 筆 ｜ B：{df_b.shape[0]} 筆")

# =========================================================
# Key 設定
# =========================================================
st.subheader("🔑 Key 欄位設定")

cols = list(df_a.columns)
default_keys = [c for c in cols if clean_header_name(c) in {"PLNNR", "VORNR"}]
if not default_keys:
    default_keys = cols[:1]

selected_keys = st.multiselect(
    "選擇 Key 欄位（可多選）",
    options=cols,
    default=default_keys,
)

if not selected_keys:
    st.warning("請至少選擇一個 Key 欄位")
    st.stop()

st.markdown("---")

# =========================================================
# 開始比對
# =========================================================
if not st.button("🟢 開始差異比對", type="primary"):
    st.stop()

t0 = time.time()

# =========================================================
# Key map / 重複
# =========================================================
key_cols_a = [df_a.columns.get_loc(k) for k in selected_keys]
key_cols_b = [df_b.columns.get_loc(k) for k in selected_keys]

map_a = build_key_map(df_a, key_cols_a)
map_b = build_key_map(df_b, key_cols_b)

dup_a = count_duplicate_keys(df_a, key_cols_a)
dup_b = count_duplicate_keys(df_b, key_cols_b)

# =========================================================
# 欄位差異
# =========================================================
df_col_diff = build_column_diff(df_a, df_b)

# =========================================================
# 嚴格差異
# =========================================================
a_rows, _, _, _ = diff_directional(df_a, df_b, map_a, map_b, key_cols_a, "A", "B")
b_rows, _, _, _ = diff_directional(df_b, df_a, map_b, map_a, key_cols_b, "B", "A")

key_headers = [f"KEY_{i+1}" for i in range(len(selected_keys))]
headers = key_headers + ["差異欄位", "A值", "B值", "差異來源"]

df_a_to_b = pd.DataFrame(a_rows, columns=headers)
df_b_to_a = (
    pd.DataFrame(
        b_rows,
        columns=key_headers + ["差異欄位", "B值", "A值", "差異來源"]
    )[headers]
    if b_rows else pd.DataFrame(columns=headers)
)

# 👉 顯示清洗（只影響顯示）
for df in (df_a_to_b, df_b_to_a):
    for c in df.columns:
        df[c] = df[c].map(clean_display_value)

# =========================================================
# Summary
# =========================================================
df_summary = pd.DataFrame([
    ["Key 欄位", ", ".join(selected_keys), "", "", ""],
    ["A 重複 Key 列數", dup_a, "", "", ""],
    ["B 重複 Key 列數", dup_b, "", "", ""],
    ["A → B 差異列數", len(df_a_to_b), "", "", ""],
    ["B → A 差異列數", len(df_b_to_a), "", "", ""],
], columns=["項目", "值1", "值2", "值3", "值4"])

# =========================================================
# 匯出
# =========================================================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df_summary.to_excel(writer, "Summary", index=False)
    df_col_diff.to_excel(writer, "ColumnDiff", index=False)
    df_a_to_b.to_excel(writer, "A_to_B", index=False)
    df_b_to_a.to_excel(writer, "B_to_A", index=False)

duration = round(time.time() - t0, 2)
st.success(f"比對完成（耗時 {duration} 秒）")

st.download_button(
    "📥 下載差異比對結果",
    data=output.getvalue(),
    file_name=gen_download_filename("Excel差異比對結果"),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# =========================================================
# Footer
# =========================================================
st.markdown(
    f"""
    <div style="
        margin-top:40px;
        padding:12px 0;
        text-align:center;
        font-size:13px;
        color:#666;
        border-top:1px solid #e0e0e0;
    ">
        {APP_FOOTER} {APP_VERSION}
    </div>
    """,
    unsafe_allow_html=True
)
