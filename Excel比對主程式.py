import streamlit as st
import pandas as pd
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

from config import APP_NAME, APP_VERSION, APP_FOOTER
from compare_core import (
    clean_header_name,
    build_key_map,
    count_duplicate_keys,
    diff_directional,
    build_column_diff
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
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

USAGE_XLSX = DATA_DIR / "usage_stats.xlsx"

SESSION_TIMEOUT_SECONDS = 30 * 60
WARNING_SECONDS = 5 * 60

# =========================================================
# 工具：台灣時間
# =========================================================
def now_tw():
    return datetime.now(ZoneInfo("Asia/Taipei"))

# =========================================================
# 🔢 系統總比對次數（永久）
# =========================================================
def load_total_compare_count():
    if not USAGE_XLSX.exists():
        df = pd.DataFrame([{
            "total_compare_count": 0,
            "last_update": now_tw().strftime("%Y-%m-%d %H:%M:%S")
        }])
        df.to_excel(USAGE_XLSX, index=False)
        return 0

    df = pd.read_excel(USAGE_XLSX)
    return int(df.loc[0, "total_compare_count"])


def increase_total_compare_count():
    total = load_total_compare_count() + 1
    df = pd.DataFrame([{
        "total_compare_count": total,
        "last_update": now_tw().strftime("%Y-%m-%d %H:%M:%S")
    }])
    df.to_excel(USAGE_XLSX, index=False)
    return total

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
    if "compare_count" not in st.session_state:
        st.session_state.compare_count = 0

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
            st.session_state.compare_count = 0
            st.session_state.warned = False
            st.rerun()
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

    total_compare = load_total_compare_count()

    st.caption(f"📊 系統累積比對次數：{total_compare}")
    st.caption(f"🔁 本次登入比對次數：{st.session_state.compare_count}")

    now = time.time()
    remaining = SESSION_TIMEOUT_SECONDS - (now - st.session_state.last_active_ts)

    if remaining <= WARNING_SECONDS and remaining > 0 and not st.session_state.warned:
        st.warning("⚠️ 登入即將逾時，請點擊「延長登入」")
        st.session_state.warned = True

    if remaining <= 0:
        st.session_state.authenticated = False
        st.rerun()

    if st.button("🔁 延長登入"):
        st.session_state.last_active_ts = time.time()
        st.session_state.warned = False
        st.success("已延長登入")
        st.rerun()

    if st.button("🔓 登出"):
        st.session_state.authenticated = False
        st.rerun()

# =========================================================
# 主畫面
# =========================================================
st.title(f"Excel 比對程式（{APP_VERSION}）")

st.markdown("""
### 使用說明
1. 上傳 Excel A、Excel B  
2. 勾選 Key 欄位（可多 Key）  
3. 點擊「開始差異比對」  

⚠️ 請確認兩份 Excel 表頭一致
""")

# =========================================================
# 檔名
# =========================================================
def gen_download_filename(base_name: str):
    ts = now_tw().strftime("%Y%m%d_%H%M%S")
    return f"{base_name}_{ts}.xlsx"

# =========================================================
# 上傳檔案
# =========================================================
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("📤 上傳 Excel A", type=["xlsx"])
with col2:
    file_b = st.file_uploader("📤 上傳 Excel B", type=["xlsx"])

output = None
download_filename = None
duration = None

# =========================================================
# 主流程
# =========================================================
if file_a and file_b:
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    st.success(f"Excel A：{df_a.shape[0]} 筆 ｜ Excel B：{df_b.shape[0]} 筆")

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

    if selected_keys:
        start_compare = st.button("🟢 開始差異比對 🟢", type="primary")
    else:
        start_compare = False
        st.info("請先選擇 Key 欄位")

    if start_compare:
        st.session_state.last_active_ts = time.time()

        # ✅ 計次「就在這裡」
        st.session_state.compare_count += 1
        increase_total_compare_count()

        with st.spinner("比對中..."):
            t0 = time.time()

            key_cols_a = [df_a.columns.get_loc(k) for k in selected_keys]
            key_cols_b = [df_b.columns.get_loc(k) for k in selected_keys]

            map_a = build_key_map(df_a, key_cols_a)
            map_b = build_key_map(df_b, key_cols_b)

            dup_a = count_duplicate_keys(df_a, key_cols_a)
            dup_b = count_duplicate_keys(df_b, key_cols_b)

            df_col_diff = build_column_diff(df_a, df_b)

            a_rows, *_ = diff_directional(df_a, df_b, map_a, map_b, key_cols_a, "A", "B")
            b_rows, *_ = diff_directional(df_b, df_a, map_b, map_a, key_cols_b, "B", "A")

            headers = [f"KEY_{i+1}" for i in range(len(selected_keys))] + ["差異欄位", "A值", "B值", "差異來源"]

            df_a_to_b = pd.DataFrame(a_rows, columns=headers)
            df_b_to_a = pd.DataFrame(b_rows, columns=headers) if b_rows else pd.DataFrame(columns=headers)

            df_summary = pd.DataFrame([
                ["Key 欄位", ", ".join(selected_keys), "", "", ""],
                ["A 重複 Key", dup_a, "", "", ""],
                ["B 重複 Key", dup_b, "", "", ""],
                ["A→B 差異列", len(df_a_to_b), "", "", ""],
                ["B→A 差異列", len(df_b_to_a), "", "", ""],
            ], columns=["項目", "值1", "值2", "值3", "值4"])

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_summary.to_excel(writer, "Summary", index=False)
                df_col_diff.to_excel(writer, "ColumnDiff", index=False)
                df_a_to_b.to_excel(writer, "A_to_B", index=False)
                df_b_to_a.to_excel(writer, "B_to_A", index=False)

            duration = round(time.time() - t0, 2)
            download_filename = gen_download_filename("Excel差異比對結果")

        st.success(f"比對完成（耗時 {duration} 秒）")

# =========================================================
# 下載
# =========================================================
if output:
    st.download_button(
        "📥 下載比對結果",
        data=output.getvalue(),
        file_name=download_filename,
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
