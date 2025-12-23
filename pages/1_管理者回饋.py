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
# Sidebar（管理者）
# ✅ 你指定：拿掉 APP_NAME / 版本字樣
# =========================================================
with st.sidebar:
    st.markdown("### 👤 管理者模式")

    if st.button("🔓 登出管理者"):
        st.session_state.admin_authenticated = False
        st.rerun()

# =========================================================
# 主畫面
# =========================================================
st.title("👤 管理者頁面｜回饋管理")

DATA_DIR = Path("data")
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

if not FEEDBACK_XLSX.exists():
    st.warning("目前尚無任何回饋資料")
    st.stop()

try:
    df = pd.read_excel(FEEDBACK_XLSX)
except Exception as e:
    st.error(f"讀取回饋資料失敗：{e}")
    st.stop()

# =========================================================
# 若無 status 欄位，自動補：預設未處理
# =========================================================
if "status" not in df.columns:
    df["status"] = "未處理"

# =========================================================
# Dashboard
# =========================================================
col1, col2, col3 = st.columns(3)
col1.metric("📨 總回饋數", len(df))
col2.metric("🟢 已處理", int((df["status"] == "已處理").sum()))
col3.metric("🔴 未處理", int((df["status"] == "未處理").sum()))

if "app_version" in df.columns:
    st.subheader("📊 版本分布")
    st.bar_chart(df["app_version"].value_counts())

st.markdown("---")

# =========================================================
# 篩選區（搜尋 / 日期 / 狀態）
# =========================================================
with st.expander("🔍 搜尋 / 篩選", expanded=True):
    keyword = st.text_input("關鍵字（姓名 / Email / 內容）", placeholder="例如：王小明 / test@xxx.com / 無法下載")
    status_filter = st.selectbox("狀態", ["全部", "未處理", "已處理"])

    # 日期：可不選；選兩個才生效
    date_range = st.date_input("日期區間（選填）", value=[])

df_view = df.copy()

# 關鍵字篩選
if keyword:
    name_s = df_view["name"].astype(str) if "name" in df_view.columns else pd.Series([""] * len(df_view))
    email_s = df_view["email"].astype(str) if "email" in df_view.columns else pd.Series([""] * len(df_view))
    msg_s = df_view["message"].astype(str) if "message" in df_view.columns else pd.Series([""] * len(df_view))

    mask = (
        name_s.str.contains(keyword, na=False)
        | email_s.str.contains(keyword, na=False)
        | msg_s.str.contains(keyword, na=False)
    )
    df_view = df_view[mask]

# 狀態篩選
if status_filter != "全部":
    df_view = df_view[df_view["status"] == status_filter]

# 日期篩選：只有選兩個日期才做
if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    start, end = date_range
    if "time_tw" in df_view.columns:
        dt = pd.to_datetime(df_view["time_tw"], errors="coerce")
        df_view = df_view[(dt.dt.date >= start) & (dt.dt.date <= end)]

# =========================================================
# ✅ 只顯示紅框欄位 + status
# ✅ 只有 status 可改（下拉）
# =========================================================
DISPLAY_COLS = ["time_tw", "name", "email", "message", "app_version", "status"]
DISPLAY_COLS = [c for c in DISPLAY_COLS if c in df_view.columns]

# 建一個「顯示用 + 可回寫 index」的 DataFrame
df_table = df_view.copy()
df_table["_row_id"] = df_table.index  # 用來回寫原 df

# 欄位排序：row_id 放最前，但不顯示給使用者
table_cols = ["_row_id"] + DISPLAY_COLS

st.subheader("📋 回饋列表（僅 status 可調整）")

edited = st.data_editor(
    df_table[table_cols],
    use_container_width=True,
    hide_index=True,
    disabled=[c for c in table_cols if c not in ("status",)],  # ✅ 只有 status 可編輯
    column_config={
        "_row_id": st.column_config.NumberColumn("row_id", disabled=True, width="small"),
        "status": st.column_config.SelectboxColumn(
            "處理狀態",
            options=["未處理", "已處理"],
            required=True,
            help="僅此欄可修改"
        ),
    },
    key="admin_feedback_editor"
)

# 把 row_id 欄藏起來（更乾淨）
# Streamlit 沒有原生完全隱藏單欄的方法，這裡用 CSS 把第一欄（row_id）寬度壓到最小 + 透明
st.markdown(
    """
    <style>
      /* 盡量把 data_editor 第一欄縮到不可見（row_id） */
      div[data-testid="stDataFrame"] thead tr th:first-child,
      div[data-testid="stDataFrame"] tbody tr td:first-child {
        max-width: 0px !important;
        width: 0px !important;
        padding: 0 !important;
        opacity: 0 !important;
      }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# 儲存狀態變更（只寫回 status）
# =========================================================
if st.button("💾 儲存狀態變更"):
    st.session_state.admin_last_active = time.time()

    try:
        # edited 內有 _row_id 與 status
        for _, r in edited[["_row_id", "status"]].iterrows():
            rid = int(r["_row_id"])
            df.loc[rid, "status"] = r["status"]

        df.to_excel(FEEDBACK_XLSX, index=False, engine="openpyxl")
        st.success("✅ 狀態已更新並存檔")
        st.rerun()
    except Exception as e:
        st.error(f"❌ 儲存失敗：{e}")

# =========================================================
# 匯出（匯出目前篩選後資料，只含紅框欄位 + status）
# =========================================================
export_df = edited[DISPLAY_COLS].copy() if len(DISPLAY_COLS) else edited.copy()

buf = BytesIO()
export_df.to_excel(buf, index=False, engine="openpyxl")
buf.seek(0)

st.download_button(
    "📥 匯出目前畫面資料（Excel）",
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
