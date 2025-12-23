import streamlit as st
import pandas as pd
import time
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

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
    page_title="QQ資料製作小組｜Excel 比對程式",
    layout="wide"
)

# =========================================================
# 登入與逾時設定
# =========================================================
SESSION_TIMEOUT_SECONDS = 30 * 60   # 30 分鐘
WARNING_SECONDS = 5 * 60            # 剩 5 分鐘警告一次

# =========================================================
# 🔐 登入檢查（含逾時）
# =========================================================
def check_password():
    now = time.time()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "last_active_ts" not in st.session_state:
        st.session_state.last_active_ts = now
    if "warned" not in st.session_state:
        st.session_state.warned = False

    # ===== 已登入 =====
    if st.session_state.authenticated:
        if now - st.session_state.last_active_ts >= SESSION_TIMEOUT_SECONDS:
            st.session_state.authenticated = False
            return False
        return True

    # ===== 尚未登入 =====
    st.title("🔐 系統登入")

    pwd = st.text_input("請輸入系統密碼", type="password")

    if st.button("登入"):
        if pwd == st.secrets["auth"]["password"]:
            st.session_state.authenticated = True
            st.session_state.last_active_ts = now
            st.session_state.warned = False
            st.rerun()
        else:
            st.error("密碼錯誤")

    return False


# ❗ 未登入或已逾時，直接停
if not check_password():
    st.stop()

# =========================================================
# Sidebar：登入狀態 / 警告 / 操作
# =========================================================
with st.sidebar:
    st.markdown("### 🟢 登入狀態")

    now = time.time()
    remaining = SESSION_TIMEOUT_SECONDS - (now - st.session_state.last_active_ts)

    # ⚠️ 剩 5 分鐘警告一次
    if remaining <= WARNING_SECONDS and remaining > 0 and not st.session_state.warned:
        st.warning("⚠️ 登入即將逾時，請點擊「延長登入」")
        st.session_state.warned = True

    # ⛔ 已逾時 → 強制回登入
    if remaining <= 0:
        st.session_state.authenticated = False
        st.rerun()

    if st.button("🔁 延長登入"):
        st.session_state.last_active_ts = time.time()
        st.session_state.warned = False
        st.success("已延長登入時間")
        st.rerun()

    if st.button("🔓 登出"):
        st.session_state.authenticated = False
        st.rerun()

# =========================================================
# 主畫面
# =========================================================
st.title("Excel 比對程式（Web V3.1.0 正式版）")

st.markdown("""
### 使用說明
1. 上傳 Excel A、Excel B  
2. 勾選 Key 欄位（可多 Key）  
3. Key 選完後，點擊「開始比對」下載結果  

⚠️ 使用前請確認兩份 Excel 表頭名稱一致
""")

# =========================================================
# 下載檔名（台灣時間）
# =========================================================
def gen_download_filename(base_name: str, suffix="compare", ext="xlsx"):
    tw_tz = ZoneInfo("Asia/Taipei")
    now_tw = datetime.now(tw_tz)
    ts = now_tw.strftime("%Y%m%d_%H%M%S")
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

output = None
download_filename = None

# =========================================================
# 主流程
# =========================================================
if file_a is None or file_b is None:
    st.info("請先上傳兩份 Excel")
else:
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    st.session_state.last_active_ts = time.time()

    st.success(f"Excel A：{df_a.shape} ｜ Excel B：{df_b.shape}")

    # =========================
    # Key 設定
    # =========================
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

    # =========================
    # Key 選完才顯示按鈕
    # =========================
    if selected_keys:
        st.success(f"已選擇 Key：{', '.join(selected_keys)}")
        st.markdown("---")
        start_compare = st.button("🟢 開始差異比對 🟢", type="primary")
    else:
        start_compare = False
        st.info("請至少選擇一個 Key 欄位後，才能開始比對")

    # =========================
    # 比對執行
    # =========================
    if start_compare:
        st.session_state.last_active_ts = time.time()

        with st.spinner("資料比對中，請稍候..."):
            t0 = time.time()

            key_cols_a = [df_a.columns.get_loc(k) for k in selected_keys]
            key_cols_b = [df_b.columns.get_loc(k) for k in selected_keys]

            map_a = build_key_map(df_a, key_cols_a)
            map_b = build_key_map(df_b, key_cols_b)

            dup_a = count_duplicate_keys(df_a, key_cols_a)
            dup_b = count_duplicate_keys(df_b, key_cols_b)

            df_col_diff = build_column_diff(df_a, df_b)

            a_rows, _, _, _ = diff_directional(
                df_a, df_b, map_a, map_b, key_cols_a, "A", "B"
            )
            b_rows, _, _, _ = diff_directional(
                df_b, df_a, map_b, map_a, key_cols_b, "B", "A"
            )

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
            download_filename = gen_download_filename("Excel差異比對結果")

        st.success(f"比對完成（耗時 {duration} 秒）")

# =========================================================
# 下載區
# =========================================================
if output and download_filename:
    st.download_button(
        "📥 下載差異比對結果 Excel",
        data=output.getvalue(),
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# Footer
# =========================================================
st.markdown(
    """
    <div style="
        margin-top:40px;
        padding:12px 0;
        text-align:center;
        font-size:13px;
        color:#666;
        border-top:1px solid #e0e0e0;
    ">
        © 2025 Roger＆Andy with GPT ｜ QQ資料製作小組 ｜ V3.1.0
    </div>
    """,
    unsafe_allow_html=True
)
