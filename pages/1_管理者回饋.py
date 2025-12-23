import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO

from config import APP_NAME, APP_VERSION, APP_FOOTER

# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title="管理者頁｜回饋列表",
    layout="wide"
)

# =========================================================
# 🔐 管理者登入（獨立權限）
# =========================================================
def check_admin():
    if "admin_authenticated" not in st.session_state:
        st.session_state.admin_authenticated = False

    if st.session_state.admin_authenticated:
        return True

    st.title("🔐 管理者登入")

    pwd = st.text_input(
        "請輸入管理者密碼",
        type="password"
    )

    if st.button("登入"):
        try:
            admin_pwd = st.secrets["admin"]["password"]
        except KeyError:
            st.error("❌ 系統未設定管理者密碼（admin.password）")
            return False

        if pwd == admin_pwd:
            st.session_state.admin_authenticated = True
            st.rerun()
        else:
            st.error("❌ 管理者密碼錯誤")

    return False


# ❗ 未通過管理者驗證，直接中止
if not check_admin():
    st.stop()

# =========================================================
# Sidebar（管理者）
# =========================================================
with st.sidebar:
    st.markdown("### 👤 管理者模式")
    st.caption(f"{APP_NAME}")
    st.caption(f"版本：{APP_VERSION}")

    if st.button("🔓 登出管理者"):
        st.session_state.admin_authenticated = False
        st.rerun()

# =========================================================
# 主畫面
# =========================================================
st.title("👤 管理者頁面｜回饋列表")
st.caption(f"系統版本：{APP_VERSION}")

DATA_DIR = Path("data")
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

# =========================================================
# 讀取回饋資料
# =========================================================
if not FEEDBACK_XLSX.exists():
    st.warning("目前尚無任何回饋資料")
    st.stop()

try:
    df = pd.read_excel(FEEDBACK_XLSX)
except Exception as e:
    st.error(f"讀取回饋資料失敗：{e}")
    st.stop()

st.success(f"共 {len(df)} 筆回饋")

# =========================================================
# ✅ 只顯示你指定的「紅框欄位」
# =========================================================
DISPLAY_COLS = [
    "time_tw",
    "name",
    "email",
    "message",
    "app_version",
]

DISPLAY_COLS = [c for c in DISPLAY_COLS if c in df.columns]
df_display = df[DISPLAY_COLS]

st.dataframe(
    df_display,
    use_container_width=True
)

# =========================================================
# 下載回饋 Excel（正確寫法）
# =========================================================
buf = BytesIO()
df_display.to_excel(buf, index=False, engine="openpyxl")
buf.seek(0)

st.download_button(
    label="📥 下載回饋 Excel",
    data=buf,
    file_name="feedback_export.xlsx",
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
        {APP_FOOTER} {APP_VERSION}
    </div>
    """,
    unsafe_allow_html=True
)
