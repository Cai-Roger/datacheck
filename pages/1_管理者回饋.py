import streamlit as st
import pandas as pd
from pathlib import Path
from config import APP_VERSION

# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title="管理者頁｜回饋列表",
    layout="wide"
)

# =========================================================
# 🔐 管理者登入檢查
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
        if pwd == st.secrets["admin"]["password"]:
            st.session_state.admin_authenticated = True
            st.rerun()
        else:
            st.error("❌ 管理者密碼錯誤")

    return False


# ❗ 未通過管理者驗證 → 停止
if not check_admin():
    st.stop()

# =========================================================
# Sidebar（管理者）
# =========================================================
with st.sidebar:
    st.markdown("### 👤 管理者模式")
    st.caption(f"Version: {APP_VERSION}")

    if st.button("🔓 登出管理者"):
        st.session_state.admin_authenticated = False
        st.rerun()

# =========================================================
# 主畫面
# =========================================================
st.title("👤 管理者頁面｜回饋列表")
st.caption(f"版本：{APP_VERSION}")

DATA_DIR = Path("data")
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

if not FEEDBACK_XLSX.exists():
    st.warning("目前尚無任何回饋資料")
else:
    df = pd.read_excel(FEEDBACK_XLSX)

    st.success(f"共 {len(df)} 筆回饋")

    # ✅ 依你要求：只保留紅框欄位
    display_cols = [
        "time_tw",
        "name",
        "email",
        "message",
        "app_version"
    ]

    display_cols = [c for c in display_cols if c in df.columns]

    st.dataframe(
        df[display_cols],
        use_container_width=True
    )

    # 下載
    st.download_button(
        "📥 下載回饋 Excel",
        data=df.to_excel(index=False, engine="openpyxl"),
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
        © 2025 Roger＆Andy with GPT ｜ 管理者頁 ｜ {APP_VERSION}
    </div>
    """,
    unsafe_allow_html=True
)
