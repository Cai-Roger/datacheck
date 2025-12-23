import streamlit as st
import pandas as pd
from pathlib import Path
from zoneinfo import ZoneInfo
from datetime import datetime

# =========================================================
# 基本設定
# =========================================================
APP_VERSION = "V3.2.1"
DATA_DIR = Path("data")
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

# =========================================================
# 🔐 登入檢查（沿用主程式 session）
# =========================================================
if "authenticated" not in st.session_state or not st.session_state.authenticated:
    st.warning("⚠️ 請先登入系統")
    st.stop()

# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title="管理者頁面｜回饋列表",
    layout="wide"
)

st.title("👤 管理者頁面｜回饋列表")
st.caption(f"版本：{APP_VERSION}")

# =========================================================
# 讀取資料
# =========================================================
if not FEEDBACK_XLSX.exists():
    st.info("目前尚無任何回饋資料")
    st.stop()

df = pd.read_excel(FEEDBACK_XLSX)

# =========================================================
# 僅顯示指定欄位（紅框）
# =========================================================
DISPLAY_COLS = [
    "time_tw",
    "name",
    "email",
    "message",
    "app_version",
]

DISPLAY_COLS = [c for c in DISPLAY_COLS if c in df.columns]
df_view = df[DISPLAY_COLS].copy()

# =========================================================
# 基本資訊
# =========================================================
st.success(f"📊 目前共 {len(df_view)} 筆回饋")

# =========================================================
# 🔍 搜尋 / 排序工具列
# =========================================================
with st.expander("🔍 搜尋 / 排序"):
    keyword = st.text_input("關鍵字（姓名 / Email / 內容）")
    sort_order = st.radio(
        "時間排序",
        ["最新在前", "最舊在前"],
        horizontal=True
    )

# =========================================================
# 搜尋處理
# =========================================================
if keyword:
    df_view = df_view[
        df_view.astype(str).apply(
            lambda r: r.str.contains(keyword, case=False, na=False).any(),
            axis=1
        )
    ]

# =========================================================
# 排序處理
# =========================================================
if "time_tw" in df_view.columns:
    df_view["__time"] = pd.to_datetime(df_view["time_tw"], errors="coerce")
    df_view = df_view.sort_values(
        "__time",
        ascending=(sort_order == "最舊在前")
    )
    df_view = df_view.drop(columns="__time")

# =========================================================
# 顯示表格
# =========================================================
st.dataframe(
    df_view,
    use_container_width=True,
    height=520
)

# =========================================================
# ⬇️ 下載回饋資料
# =========================================================
st.markdown("---")

def gen_admin_export_name():
    ts = datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y%m%d_%H%M%S")
    return f"feedback_admin_export_{ts}.xlsx"

output = None
with pd.ExcelWriter(
    output := pd.ExcelWriter,
    engine="xlsxwriter"
):
    pass  # just for editor hinting

export_buf = None
export_buf = pd.ExcelWriter

buf = None
buf = st.experimental_data_editor if False else None

export = None
from io import BytesIO
export = BytesIO()

with pd.ExcelWriter(export, engine="xlsxwriter") as writer:
    df_view.to_excel(writer, sheet_name="Feedback", index=False)

st.download_button(
    "⬇️ 下載回饋資料 Excel",
    data=export.getvalue(),
    file_name=gen_admin_export_name(),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
        © 2025 Roger＆Andy with GPT ｜ QQ資料製作小組 ｜ 管理者頁 {APP_VERSION}
    </div>
    """,
    unsafe_allow_html=True
)
