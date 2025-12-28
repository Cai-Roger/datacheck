import streamlit as st
import pandas as pd
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage
import smtplib
from pathlib import Path

from config import APP_NAME, APP_VERSION, APP_FOOTER
from compare_core import (
    clean_header_name,
    build_key_map,
    count_duplicate_keys,
    diff_directional,
    build_column_diff,
)

# =========================================================
# Page config（一定要第一個）
# =========================================================
st.set_page_config(
    page_title=f"{APP_NAME}",
    layout="wide",
)

# =========================================================
# 登入與逾時設定
# =========================================================
SESSION_TIMEOUT_SECONDS = 30 * 60
WARNING_SECONDS = 5 * 60

# =========================================================
# 資料路徑
# =========================================================
DATA_DIR = Path("data")
DATA_DIR.mkdir(parents=True, exist_ok=True)

FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"
USAGE_XLSX = DATA_DIR / "usage.xlsx"

# =========================================================
# 工具
# =========================================================
def now_tw():
    return datetime.now(ZoneInfo("Asia/Taipei"))

def gen_download_filename(base_name: str, suffix="compare", ext="xlsx"):
    ts = now_tw().strftime("%Y%m%d_%H%M%S")
    seq = int(time.time() * 1000) % 1000
    return f"{base_name}_{suffix}_{ts}_{seq:03d}.{ext}"

# =========================================================
# ✅【新增】寫入 Excel 前的清洗（只做兩件事）
# 1. NaN → ""
# 2. 移除 ␣（U+2423）
# =========================================================
def clean_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.fillna("")
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace("\u2423", "", regex=False)
            )
    return df

# =========================================================
# 系統累積比對次數（持久化）
# =========================================================
def get_total_compare_count() -> int:
    if not USAGE_XLSX.exists():
        return 0
    try:
        df = pd.read_excel(USAGE_XLSX)
        if "total_compare" not in df.columns or df.empty:
            return 0
        return int(df.loc[0, "total_compare"])
    except Exception:
        return 0

def set_total_compare_count(n: int):
    pd.DataFrame([{
        "total_compare": n,
        "updated_time_tw": now_tw().strftime("%Y-%m-%d %H:%M:%S"),
        "app_version": APP_VERSION,
    }]).to_excel(USAGE_XLSX, index=False, engine="openpyxl")

def bump_total_compare_count() -> int:
    n = get_total_compare_count() + 1
    set_total_compare_count(n)
    return n

# =========================================================
# 回饋相關
# =========================================================
def send_feedback_email(subject: str, body: str):
    cfg = st.secrets.get("mail")
    if not cfg:
        return
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f'{cfg.get("from_name","Feedback")} <{cfg["smtp_user"]}>'
    msg["To"] = cfg["to_addr"]
    msg.set_content(body)
    with smtplib.SMTP(cfg["smtp_host"], int(cfg["smtp_port"])) as server:
        server.starttls()
        server.login(cfg["smtp_user"], cfg["smtp_password"])
        server.send_message(msg)

def append_feedback_to_excel(row: dict):
    cols = ["time_tw","name","email","message","app_version","compare_count_session"]
    new_df = pd.DataFrame([[row.get(c,"") for c in cols]], columns=cols)
    if FEEDBACK_XLSX.exists():
        try:
            old = pd.read_excel(FEEDBACK_XLSX)
            out = pd.concat([old, new_df], ignore_index=True)
        except Exception:
            out = new_df
    else:
        out = new_df
    out.to_excel(FEEDBACK_XLSX, index=False, engine="openpyxl")

# =========================================================
# 登入檢查
# =========================================================
def check_password():
    now = time.time()
    st.session_state.setdefault("authenticated", False)
    st.session_state.setdefault("last_active_ts", now)
    st.session_state.setdefault("warned", False)
    st.session_state.setdefault("compare_count_session", 0)

    if st.session_state.authenticated:
        if now - st.session_state.last_active_ts >= SESSION_TIMEOUT_SECONDS:
            st.session_state.authenticated = False
            return False
        return True

    st.title("🔐 Excel 比對程式｜登入")
    pwd = st.text_input("請輸入系統密碼", type="password")
    if st.button("登入"):
        if pwd == st.secrets["auth"]["password"]:
            st.session_state.authenticated = True
            st.session_state.last_active_ts = now
            st.session_state.compare_count_session = 0
            st.stop()
        else:
            st.error("密碼錯誤")
    return False

if not check_password():
    st.stop()

# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.markdown("### 🟢 登入狀態")
    st.caption(f"版本：{APP_VERSION}")
    st.caption(f"📊 系統累積比對次數：{get_total_compare_count()}")
    st.caption(f"🔁 本次登入比對次數：{st.session_state.compare_count_session}")

    if st.button("🔁 延長登入"):
        st.session_state.last_active_ts = time.time()
        st.session_state.warned = False

    if st.button("🔓 登出"):
        st.session_state.authenticated = False
        st.stop()

    st.markdown("---")
    st.markdown("### ✉️ 意見箱")
    with st.form("feedback_form", clear_on_submit=True):
        fb_name = st.text_input("姓名 / 暱稱")
        fb_email = st.text_input("Email")
        fb_msg = st.text_area("意見內容")
        submitted = st.form_submit_button("送出")

    if submitted and fb_msg.strip():
        row = {
            "time_tw": now_tw().strftime("%Y-%m-%d %H:%M:%S"),
            "name": fb_name,
            "email": fb_email,
            "message": fb_msg,
            "app_version": APP_VERSION,
            "compare_count_session": st.session_state.compare_count_session,
        }
        append_feedback_to_excel(row)
        send_feedback_email(f"【{APP_NAME}】新回饋", fb_msg)
        st.success("已送出")

# =========================================================
# 主畫面
# =========================================================
st.title(f"Excel 比對程式（Web {APP_VERSION}）")

col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("上傳 Excel A", type=["xlsx"])
with col2:
    file_b = st.file_uploader("上傳 Excel B", type=["xlsx"])

if not file_a or not file_b:
    st.stop()

df_a = pd.read_excel(file_a)
df_b = pd.read_excel(file_b)
st.success(f"A：{len(df_a)} 筆 ｜ B：{len(df_b)} 筆")

cols = list(df_a.columns)
default_keys = [c for c in cols if clean_header_name(c) in {"PLNNR","VORNR"}] or cols[:2]

keys = st.multiselect("選擇 Key 欄位", cols, default=default_keys)
if not keys:
    st.stop()

if st.button("🟢 開始差異比對", type="primary"):
    st.session_state.compare_count_session += 1
    new_total = bump_total_compare_count()

    with st.spinner("比對中..."):
        key_cols_a = [df_a.columns.get_loc(k) for k in keys]
        key_cols_b = [df_b.columns.get_loc(k) for k in keys]

        map_a = build_key_map(df_a, key_cols_a)
        map_b = build_key_map(df_b, key_cols_b)

        df_col_diff = build_column_diff(df_a, df_b)
        a_rows, *_ = diff_directional(df_a, df_b, map_a, map_b, key_cols_a, "A", "B")
        b_rows, *_ = diff_directional(df_b, df_a, map_b, map_a, key_cols_b, "B", "A")

        headers = [f"KEY_{i+1}" for i in range(len(keys))] + ["差異欄位","A值","B值","差異來源"]
        df_a_to_b = pd.DataFrame(a_rows, columns=headers)
        df_b_to_a = pd.DataFrame(b_rows, columns=headers)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            clean_for_excel(df_col_diff).to_excel(writer, "ColumnDiff", index=False)
            clean_for_excel(df_a_to_b).to_excel(writer, "A_to_B", index=False)
            clean_for_excel(df_b_to_a).to_excel(writer, "B_to_A", index=False)

    st.download_button(
        "下載結果",
        output.getvalue(),
        gen_download_filename("Excel差異比對"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown(f"<hr>{APP_FOOTER} {APP_VERSION}", unsafe_allow_html=True)
