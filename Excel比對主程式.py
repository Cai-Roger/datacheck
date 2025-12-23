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
# Page config（一定要第一個 Streamlit 呼叫）
# =========================================================
st.set_page_config(
    page_title=f"{APP_NAME}｜Excel 比對程式",
    layout="wide"
)

# =========================================================
# 基本設定
# =========================================================
DATA_DIR = Path("data")
DATA_DIR.mkdir(parents=True, exist_ok=True)

USAGE_XLSX = DATA_DIR / "usage_stats.xlsx"

SESSION_TIMEOUT_SECONDS = 30 * 60
WARNING_SECONDS = 5 * 60

# =========================================================
# 工具：台灣時間
# =========================================================
def now_tw():
    return datetime.now(ZoneInfo("Asia/Taipei"))

def safe_read_excel(path: Path) -> pd.DataFrame:
    """避免因 engine 或檔案損壞導致整個程式掛掉"""
    try:
        return pd.read_excel(path, engine="openpyxl")
    except Exception:
        # 最後保底：讓你至少能繼續跑
        return pd.DataFrame()

def safe_write_excel(df: pd.DataFrame, path: Path, sheet_name: str = "Sheet1"):
    """用 temp 檔原子覆寫，避免寫到一半被讀/中斷造成檔案毀損"""
    tmp = path.with_suffix(".tmp.xlsx")
    df.to_excel(tmp, index=False, engine="openpyxl", sheet_name=sheet_name)
    tmp.replace(path)

# =========================================================
# 🔢 系統總比對次數（永久）
# =========================================================
def ensure_usage_file():
    """確保 usage_stats.xlsx 存在且欄位正確"""
    if not USAGE_XLSX.exists():
        df = pd.DataFrame([{
            "total_compare_count": 0,
            "last_update": now_tw().strftime("%Y-%m-%d %H:%M:%S")
        }])
        safe_write_excel(df, USAGE_XLSX, sheet_name="Usage")
        return

    df = safe_read_excel(USAGE_XLSX)
    if df.empty or "total_compare_count" not in df.columns:
        # 檔案壞掉或格式不對 -> 自修復
        df = pd.DataFrame([{
            "total_compare_count": 0,
            "last_update": now_tw().strftime("%Y-%m-%d %H:%M:%S")
        }])
        safe_write_excel(df, USAGE_XLSX, sheet_name="Usage")
        return

    # 若有資料但第一列缺值也自修復
    try:
        _ = int(df.loc[0, "total_compare_count"])
    except Exception:
        df.loc[0, "total_compare_count"] = 0
        df.loc[0, "last_update"] = now_tw().strftime("%Y-%m-%d %H:%M:%S")
        safe_write_excel(df, USAGE_XLSX, sheet_name="Usage")

def load_total_compare_count() -> int:
    ensure_usage_file()
    df = safe_read_excel(USAGE_XLSX)
    try:
        return int(df.loc[0, "total_compare_count"])
    except Exception:
        return 0

def increase_total_compare_count() -> int:
    ensure_usage_file()
    df = safe_read_excel(USAGE_XLSX)
    try:
        total = int(df.loc[0, "total_compare_count"]) + 1
    except Exception:
        total = 1
    df.loc[0, "total_compare_count"] = total
    df.loc[0, "last_update"] = now_tw().strftime("%Y-%m-%d %H:%M:%S")
    safe_write_excel(df, USAGE_XLSX, sheet_name="Usage")
    return total

# =========================================================
# 🔐 登入檢查
# =========================================================
def check_password() -> bool:
    now = time.time()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "last_active_ts" not in st.session_state:
        st.session_state.last_active_ts = now
    if "warned" not in st.session_state:
        st.session_state.warned = False
    if "compare_count_session" not in st.session_state:
        st.session_state.compare_count_session = 0

    if st.session_state.authenticated:
        if now - st.session_state.last_active_ts >= SESSION_TIMEOUT_SECONDS:
            st.session_state.authenticated = False
            st.warning("⏰ 登入逾時，請重新登入")
            return False
        return True

    st.title("🔐 系統登入")
    pwd = st.text_input("請輸入系統密碼", type="password")

    # 防呆：secrets 不存在時提示（避免直接 KeyError）
    try:
        system_pwd = st.secrets["auth"]["password"]
    except Exception:
        st.error("❌ 未設定 st.secrets['auth']['password']，請先設定 Secrets")
        return False

    if st.button("登入"):
        if pwd == system_pwd:
            st.session_state.authenticated = True
            st.session_state.last_active_ts = now
            st.session_state.warned = False
            st.session_state.compare_count_session = 0
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
    st.caption(f"版本：{APP_VERSION}")

    total_compare = load_total_compare_count()
    st.caption(f"📊 系統累積比對次數：{total_compare}")
    st.caption(f"🔁 本次登入比對次數：{st.session_state.compare_count_session}")

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
st.title(f"{APP_NAME}（Web {APP_VERSION}）")

st.markdown("""
### 使用說明
1. 上傳 Excel A、Excel B  
2. 勾選 Key 欄位（可多 Key）  
3. 點擊「開始差異比對」下載結果  

⚠️ 使用前請確認兩份 Excel 表頭名稱一致
""")

def gen_download_filename(base_name: str, suffix="compare", ext="xlsx"):
    ts = now_tw().strftime("%Y%m%d_%H%M%S")
    seq = int(time.time() * 1000) % 1000
    return f"{base_name}_{suffix}_{ts}_{seq:03d}.{ext}"

# =========================================================
# 上傳檔案
# =========================================================
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("📤 上傳 Excel A", type=["xlsx"])
with col2:
    file_b = st.file_uploader("📤 上傳 Excel B", type=["xlsx"])

output_bytes = None
download_filename = None

# =========================================================
# 主流程
# =========================================================
if not file_a or not file_b:
    st.info("請先上傳兩份 Excel")
else:
    # 讀檔防呆
    try:
        df_a = pd.read_excel(file_a, engine="openpyxl")
        df_b = pd.read_excel(file_b, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ 讀取 Excel 失敗：{e}")
        st.stop()

    st.session_state.last_active_ts = time.time()
    st.success(f"📄 Excel A：{df_a.shape[0]} 筆 ｜ Excel B：{df_b.shape[0]} 筆")

    st.subheader("🔑 Key 欄位設定")

    cols = list(df_a.columns)
    default_keys = [c for c in cols if clean_header_name(c) in {"PLNNR", "VORNR"}]
    if not default_keys:
        default_keys = cols[:2] if len(cols) >= 2 else cols

    selected_keys = st.multiselect(
        "選擇 Key 欄位（可多選）",
        options=cols,
        default=default_keys
    )

    if not selected_keys:
        st.info("請至少選擇一個 Key 欄位後，才能開始比對")
    else:
        # Key 是否存在於 B
        missing = [k for k in selected_keys if k not in df_b.columns]
        if missing:
            st.error(f"Excel B 缺少 Key 欄位：{missing}")
            st.stop()

        if st.button("🟢 開始差異比對 🟢", type="primary"):
            # ✅ 按下按鈕當下就計次（不等下載）
            st.session_state.compare_count_session += 1
            increase_total_compare_count()

            with st.spinner("資料比對中，請稍候..."):
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
                output_bytes = output.getvalue()
                download_filename = gen_download_filename("Excel差異比對結果")

            st.success(f"✅ 比對完成（耗時 {duration} 秒）")
            st.info(f"📊 系統累積比對次數：{load_total_compare_count()}（已更新）")

# =========================================================
# 下載區（下載不影響計次）
# =========================================================
if output_bytes and download_filename:
    st.download_button(
        "📥 下載差異比對結果 Excel",
        data=output_bytes,
        file_name=download_filename,
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
