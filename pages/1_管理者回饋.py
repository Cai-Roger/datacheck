import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
import time
from datetime import datetime

from config import APP_NAME, APP_VERSION, APP_FOOTER

# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title="管理者頁｜回饋管理",
    layout="wide"
)

# =========================================================
# 管理者逾時設定（10 分鐘）
# =========================================================
ADMIN_TIMEOUT_SECONDS = 10 * 60

# =========================================================
# 🔐 管理者登入（含逾時）
# =========================================================
def check_admin():
    now = time.time()

    if "admin_authenticated" not in st.session_state:
        st.session_state.admin_authenticated = False
    if "admin_last_active" not in st.session_state:
        st.session_state.admin_last_active = now

    # 已登入 → 檢查逾時
    if st.session_state.admin_authenticated:
        if now - st.session_state.admin_last_active > ADMIN_TIMEOUT_SECONDS:
            st.session_state.admin_authenticated = False
            st.warning("⏰ 管理者登入已逾時，請重新登入")
            return False

        st.session_state.admin_last_active = now
        return True

    # 尚未登入
    st.title("🔐 管理者登入")

    pwd = st.text_input("請輸入管理者密碼", type="password")

    if st.button("登入"):
        try:
            admin_pwd = st.secrets["admin"]["password"]
        except KeyError:
            st.error("❌ 系統未設定管理者密碼（admin.password）")
            return False

        if pwd == admin_pwd:
            st.session_state.admin_authenticated = True
            st.session_state.admin_last_active = now
            st.rerun()
        else:
            st.error("❌ 管理者密碼錯誤")

    return False


if not check_admin():
    st.stop()

# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.markdown("### 👤 管理者模式")
    st.caption(APP_NAME)
    st.caption(f"版本：{APP_VERSION}")

    if st.button("🔓 登出管理者"):
        st.session_state.admin_authenticated = False
        st.rerun()

# =========================================================
# 主畫面
# =========================================================
st.title("👤 管理者頁面｜回饋管理")
st.caption(f"系統版本：{APP_VERSION}")

DATA_DIR = Path("data")
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

if not FEEDBACK_XLSX.exists():
    st.warning("目前尚無任何回饋資料")
    st.stop()

df = pd.read_excel(FEEDBACK_XLSX)

# =========================================================
# 若無 status 欄位，自動補
# =========================================================
if "status" not in df.columns:
    df["status"] = "未處理"

# =========================================================
# Dashboard
# =========================================================
col1, col2, col3 = st.columns(3)
col1.metric("📨 總回饋數", len(df))
col2.metric("🟢 已處理", (df["status"] == "已處理").sum())
col3.metric("🔴 未處理", (df["status"] == "未處理").sum())

st.bar_chart(df["app_version"].value_counts())

st.markdown("---")

# =========================================================
# 篩選區
# =========================================================
with st.expander("🔍 搜尋 / 篩選"):
    keyword = st.text_input("關鍵字（姓名 / Email / 內容）")
    status_filter = st.selectbox("狀態", ["全部", "未處理", "已處理"])
    date_range = st.date_input(
        "日期區間",
        []
    )

df_view = df.copy()

if keyword:
    df_view = df_view[
        df_view["name"].astype(str).str.contains(keyword, na=False)
        | df_view["email"].astype(str).str.contains(keyword, na=False)
        | df_view["message"].astype(str).str.contains(keyword, na=False)
    ]

if status_filter != "全部":
    df_view = df_view[df_view["status"] == status_filter]

if len(date_range) == 2:
    start, end = date_range
    df_view["time_tw_dt"] = pd.to_datetime(df_view["time_tw"])
    df_view = df_view[
        (df_view["time_tw_dt"].dt.date >= start)
        & (df_view["time_tw_dt"].dt.date <= end)
    ]

# =========================================================
# 表格（可編輯 status）
# =========================================================
DISPLAY_COLS = [
    "time_tw",
    "name",
    "email",
    "message",
    "app_version",
    "status"
]

df_edit = st.data_editor(
    df_view[DISPLAY_COLS],
    use_container_width=True,
    num_rows="dynamic",
    key="editor"
)

# =========================================================
# 儲存狀態變更
# =========================================================
if st.button("💾 儲存狀態變更"):
    for idx, row in df_edit.iterrows():
        df.loc[df.index == idx, "status"] = row["status"]

    df.to_excel(FEEDBACK_XLSX, index=False, engine="openpyxl")
    st.success("✅ 狀態已更新")
    st.rerun()

# =========================================================
# 匯出（依目前篩選）
# =========================================================
buf = BytesIO()
df_edit.to_excel(buf, index=False, engine="openpyxl")
buf.seek(0)

st.download_button(
    "📥 匯出目前畫面資料",
    data=buf,
    file_name="feedback_filtered.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =========================================================
# Footer
# =========================================================
st.markdown(
    f"""
    <div style="margin-top:40px;padding:12px 0;text-align:center;
                font-size:13px;color:#666;border-top:1px solid #e0e0e0;">
        {APP_FOOTER} {APP_VERSION}
    </div>
    """,
    unsafe_allow_html=True
)
