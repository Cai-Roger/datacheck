import streamlit as st
import pandas as pd
from pathlib import Path

APP_VERSION = "V3.2.0"
DATA_DIR = Path("data")
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

st.set_page_config(page_title="管理者回饋", layout="wide")

st.title("👤 管理者頁面｜回饋列表")
st.caption(f"版本：{APP_VERSION}")

# =========================================================
# 管理者登入（使用 secrets）
# =========================================================
admin_pwd = st.secrets.get("admin", {}).get("password")
if not admin_pwd:
    admin_pwd = st.secrets["auth"]["password"]  # 沒設 admin 就用一般密碼

if "admin_authed" not in st.session_state:
    st.session_state.admin_authed = False

if not st.session_state.admin_authed:
    st.info("請輸入管理者密碼")
    pwd = st.text_input("管理者密碼", type="password")
    if st.button("登入"):
        if pwd == admin_pwd:
            st.session_state.admin_authed = True
            st.rerun()
        else:
            st.error("密碼錯誤")
    st.stop()

# =========================================================
# 讀取回饋
# =========================================================
if not FEEDBACK_XLSX.exists():
    st.warning("目前尚無回饋資料（data/feedback.xlsx 不存在）")
    st.stop()

try:
    df = pd.read_excel(FEEDBACK_XLSX)
except Exception as e:
    st.error(f"讀取失敗：{e}")
    st.stop()

st.success(f"共 {len(df)} 筆回饋")
st.dataframe(df, use_container_width=True)

# 下載回饋檔
with open(FEEDBACK_XLSX, "rb") as f:
    st.download_button(
        "📥 下載 feedback.xlsx",
        data=f.read(),
        file_name="feedback.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.button("🔓 管理者登出"):
    st.session_state.admin_authed = False
    st.rerun()
