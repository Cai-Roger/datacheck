import streamlit as st
import pandas as pd
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
from email.message import EmailMessage
import smtplib

from config import APP_NAME, APP_VERSION, APP_FOOTER
from compare_core import (
    clean_header_name,
    build_key_map,
    diff_directional,
)

# =========================================================
# Page config（一定要第一個）
# =========================================================
st.set_page_config(
    page_title=f"{APP_NAME}｜Excel 比對程式",
    layout="wide"
)

# =========================================================
# 基本設定
# =========================================================
SESSION_TIMEOUT_SECONDS = 30 * 60
WARNING_SECONDS = 5 * 60

DATA_DIR = Path("data")
FEEDBACK_XLSX = DATA_DIR / "feedback.xlsx"

# =========================================================
# 工具：台灣時間
# =========================================================
def now_tw():
    return datetime.now(ZoneInfo("Asia/Taipei"))

def gen_download_filename(base_name: str, suffix="compare", ext="xlsx"):
    ts = now_tw().strftime("%Y%m%d_%H%M%S")
    seq = int(time.time() * 1000) % 1000
    return f"{base_name}_{suffix}_{ts}_{seq:03d}.{ext}"

# =========================================================
# 🔐 登入檢查
# =========================================================
def check_password():
    now = time.time()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "last_active_ts" not in st.session_state:
        st.session_state.last_active_ts = now
    if "warned" not in st.session_state:
        st.session_state.warned = False

    # Session 統計與事件鎖
    if "compare_count" not in st.session_state:
        st.session_state.compare_count = 0
    if "compare_clicked" not in st.session_state:
        st.session_state.compare_clicked = False

    # 儲存比對輸出（讓 rerun 後 download 還在）
    if "last_output_bytes" not in st.session_state:
        st.session_state.last_output_bytes = None
    if "last_output_name" not in st.session_state:
        st.session_state.last_output_name = None
    if "last_duration" not in st.session_state:
        st.session_state.last_duration = None

    if st.session_state.authenticated:
        if now - st.session_state.last_active_ts >= SESSION_TIMEOUT_SECONDS:
            st.session_state.authenticated = False
            return False
        return True

    st.title("🔐 系統登入")
    pwd = st.text_input("請輸入系統密碼", type="password")

    if st.button("登入"):
        if pwd == st.secrets["auth"]["password"]:
            st.session_state.authenticated = True
            st.session_state.last_active_ts = now
            st.session_state.warned = False
            st.session_state.compare_count = 0
            st.session_state.compare_clicked = False

            # 清掉上次結果
            st.session_state.last_output_bytes = None
            st.session_state.last_output_name = None
            st.session_state.last_duration = None

            st.rerun()
        else:
            st.error("密碼錯誤")

    return False

if not check_password():
    st.stop()

# =========================================================
# 回饋寫入 Excel
# =========================================================
def append_feedback_to_excel(row: dict):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    cols = ["time_tw", "name", "email", "message", "app_version", "compare_count_session"]
    new_df = pd.DataFrame([[row.get(c, "") for c in cols]], columns=cols)

    if FEEDBACK_XLSX.exists():
        old = pd.read_excel(FEEDBACK_XLSX)
        out = pd.concat([old, new_df], ignore_index=True)
    else:
        out = new_df

    out.to_excel(FEEDBACK_XLSX, index=False)

# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.markdown("### 🟢 登入狀態")
    st.caption(f"🔁 本次登入｜比對執行次數：{st.session_state.compare_count}")
    st.caption(f"版本：{APP_VERSION}")

    now = time.time()
    remaining = SESSION_TIMEOUT_SECONDS - (now - st.session_state.last_active_ts)

    if remaining <= WARNING_SECONDS and remaining > 0 and not st.session_state.warned:
        st.warning("⚠️ 登入即將逾時，請點擊延長登入")
        st.session_state.warned = True

    if remaining <= 0:
        st.session_state.authenticated = False
        st.rerun()

    if st.button("🔁 延長登入"):
        st.session_state.last_active_ts = time.time()
        st.session_state.warned = False
        st.rerun()

    if st.button("🔓 登出"):
        st.session_state.authenticated = False
        st.rerun()

    # 意見箱
    st.markdown("---")
    st.markdown("### ✉️ 意見箱")

    with st.form("feedback_form", clear_on_submit=True):
        fb_name = st.text_input("姓名 / 暱稱（選填）")
        fb_email = st.text_input("聯絡信箱（選填）")
        fb_msg = st.text_area("意見內容", height=120)
        submitted = st.form_submit_button("📩 送出")

    if submitted:
        if not fb_msg.strip():
            st.error("請先輸入意見內容")
        else:
            row = {
                "time_tw": now_tw().strftime("%Y-%m-%d %H:%M:%S"),
                "name": fb_name,
                "email": fb_email,
                "message": fb_msg,
                "app_version": APP_VERSION,
                "compare_count_session": st.session_state.compare_count,
            }
            append_feedback_to_excel(row)
            st.success("✅ 已收到回饋")

# =========================================================
# 主畫面
# =========================================================
st.title(f"{APP_NAME}（Web {APP_VERSION}）")

st.markdown("""
### 使用說明
1. 上傳 Excel A、Excel B  
2. 確認 Key 欄位  
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

# =========================================================
# 主流程
# =========================================================
df_a = None
df_b = None
selected_keys = []

if file_a and file_b:
    st.session_state.last_active_ts = time.time()

    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    st.success(f"📄 Excel A 筆數：{len(df_a)} ｜ Excel B 筆數：{len(df_b)}")

    st.subheader("🔑 Key 欄位設定")

    cols = df_a.columns.tolist()
    default_keys = [c for c in cols if clean_header_name(c) in {"PLNNR", "VORNR"}]
    if not default_keys:
        default_keys = cols[:1]

    selected_keys = st.multiselect(
        "選擇 Key 欄位（可多選）",
        options=cols,
        default=default_keys
    )

    if selected_keys:
        if st.button("🟢 開始差異比對 🟢", type="primary"):
            st.session_state.compare_clicked = True

# =========================================================
# 真正執行（只跑一次）— 完成後立即 rerun 刷新 sidebar
# =========================================================
if st.session_state.compare_clicked:
    st.session_state.compare_clicked = False

    # ✅ 計次：按下「開始比對」就算一次（與下載無關）
    st.session_state.compare_count += 1
    st.session_state.last_active_ts = time.time()
    st.session_state.warned = False

    t0 = time.time()

    with st.spinner("資料比對中，請稍候..."):
        # 這裡防呆：如果使用者在 rerun 過程中把檔案拿掉
        if df_a is None or df_b is None or not selected_keys:
            st.error("檔案或 Key 尚未準備好，請重新上傳並選擇 Key")
        else:
            key_cols_a = [df_a.columns.get_loc(k) for k in selected_keys]
            key_cols_b = [df_b.columns.get_loc(k) for k in selected_keys]

            map_a = build_key_map(df_a, key_cols_a)
            map_b = build_key_map(df_b, key_cols_b)

            a_rows, *_ = diff_directional(df_a, df_b, map_a, map_b, key_cols_a, "A", "B")
            b_rows, *_ = diff_directional(df_b, df_a, map_b, map_a, key_cols_b, "B", "A")

            headers = [f"KEY_{i+1}" for i in range(len(selected_keys))] + ["差異欄位", "A值", "B值", "差異來源"]
            df_out = pd.DataFrame(a_rows + b_rows, columns=headers)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_out.to_excel(writer, index=False)

            duration = round(time.time() - t0, 2)

            st.session_state.last_output_bytes = output.getvalue()
            st.session_state.last_output_name = gen_download_filename("Excel差異比對結果")
            st.session_state.last_duration = duration

    # ✅ 立刻 rerun：讓 sidebar 次數「不用下載」就立即更新
    st.rerun()

# =========================================================
# 顯示上次比對結果（下載區 & 耗時）
# =========================================================
if st.session_state.last_output_bytes:
    if st.session_state.last_duration is not None:
        st.success(f"✅ 比對完成，耗時 {st.session_state.last_duration} 秒")

    st.download_button(
        "📥 下載差異比對結果 Excel",
        data=st.session_state.last_output_bytes,
        file_name=st.session_state.last_output_name or "Excel差異比對結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# Footer
# =========================================================
st.markdown(
    f"""
    <div style="margin-top:40px;text-align:center;font-size:13px;color:#666;">
        {APP_FOOTER} {APP_VERSION}
    </div>
    """,
    unsafe_allow_html=True
)
