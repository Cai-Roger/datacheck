import streamlit as st
import pandas as pd
from io import BytesIO
import time
import streamlit.components.v1 as components
from datetime import datetime
from zoneinfo import ZoneInfo

from compare_core import (
    clean_header_name,
    build_key_map,
    count_duplicate_keys,
    diff_directional,
    build_column_diff
)

# =========================
# Page config
# =========================
st.set_page_config(
    page_title="QQ資料製作小組｜Excel 比對程式",
    layout="wide"
)

st.title("Excel 比對程式（Web V2.1.2 正式版）")

st.markdown("""
***使用說明***
1. 上傳 Excel A、Excel B  
2. 勾選 Key 欄位（可多 Key）  
3. 執行比對後下載結果  

⚠️ 使用前請確認兩份 Excel 表頭名稱一致
""")

# =========================
# 檔名產生器（年月日＋時間＋流水）
# =========================
def gen_download_filename(base_name: str, suffix="compare", ext="xlsx"):
    """
    檔名時間使用台灣時間（Asia/Taipei）
    例：xxx_compare_20250109_162355_123.xlsx
    """
    tw_tz = ZoneInfo("Asia/Taipei")
    now_tw = datetime.now(tw_tz)

    ts = now_tw.strftime("%Y%m%d_%H%M%S")
    seq = int(time.time() * 1000) % 1000  # 流水號（避免同秒重複）
    return f"{base_name}_{suffix}_{ts}_{seq:03d}.{ext}"

# =========================
# 上傳檔案
# =========================
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

st.success(f"Excel A：{df_a.shape} ｜ Excel B：{df_b.shape}")

# =========================
# Key 勾選
# =========================
st.subheader("🔑 Key 欄位設定")

cols = list(df_a.columns)
default_keys = [c for c in cols if clean_header_name(c) in {"PLNNR", "VORNR"}]
if not default_keys:
    default_keys = cols[:2]

selected_keys = st.multiselect(
    "選擇 Key 欄位（可多選）",
    options=cols,
    default=default_keys
)

if not selected_keys:
    st.error("請至少選擇一個 Key 欄位")
    st.stop()

missing = [k for k in selected_keys if k not in df_b.columns]
if missing:
    st.error(f"Excel B 缺少 Key 欄位：{missing}")
    st.stop()

key_cols_a = [df_a.columns.get_loc(k) for k in selected_keys]
key_cols_b = [df_b.columns.get_loc(k) for k in selected_keys]

# =========================
# 比對執行
# =========================
output = None
download_filename = None

if st.button("🟢 開始差異比對 🟢 ", type="primary"):
    with st.spinner("資料比對中，請稍候..."):
        t0 = time.time()

        map_a = build_key_map(df_a, key_cols_a)
        map_b = build_key_map(df_b, key_cols_b)

        dup_a = count_duplicate_keys(df_a, key_cols_a)
        dup_b = count_duplicate_keys(df_b, key_cols_b)

        df_col_diff = build_column_diff(df_a, df_b)

        a_rows, _, _, _ = diff_directional(
            df_a, df_b, map_a, map_b, key_cols_a, "A", "B"
        )
        b_rows, _, _, _ = diff_directional(
            df_b, df_a, map_b, map_a, key_cols_b, "B", "A"
        )

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

        df_summary = pd.DataFrame([
            ["Key 欄位", ", ".join(selected_keys), "", "", ""],
            ["A 重複 Key 列數", dup_a, "", "", ""],
            ["B 重複 Key 列數", dup_b, "", "", ""],
            ["A → B 差異列數", len(df_a_to_b), "", "", ""],
            ["B → A 差異列數", len(df_b_to_a), "", "", ""],
        ], columns=["項目", "值1", "值2", "值3", "值4"])

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_summary.to_excel(writer, "Summary", index=False)
            df_col_diff.to_excel(writer, "ColumnDiff", index=False)
            df_a_to_b.to_excel(writer, "A_to_B", index=False)
            df_b_to_a.to_excel(writer, "B_to_A", index=False)

        duration = round(time.time() - t0, 2)

        key_part = "+".join(selected_keys)
        download_filename = gen_download_filename(
            base_name=f"Excel差異比對結果"
        )

    st.success(f"比對完成（耗時 {duration} 秒）")

# =========================
# 下載區（⚠️ 一定要在 button 區塊外）
# =========================
if output and download_filename:
    st.download_button(
        "📥 下載差異比對結果 Excel",
        data=output.getvalue(),
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================
# Footer（固定頁底）
# =========================
components.html(
    """
    <style>
      [data-testid="stMainBlockContainer"] {
          padding-bottom: 60px;
      }

      .app-footer {
          position: fixed;
          left: 0;
          bottom: 0;
          width: 100%;
          background-color: #f5f6f7;
          color: #555;
          text-align: center;
          padding: 10px 0;
          font-size: 13px;
          border-top: 1px solid #e0e0e0;
          z-index: 9999;
      }
    </style>

    <div class="app-footer">
        © 2025 Cai-Roger ｜ Excel 比對程式 ｜ V2.0
    </div>
    """,
    height=0
)
