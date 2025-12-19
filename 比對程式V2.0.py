import pandas as pd
import os
import re
import customtkinter as ctk
from tkinter import filedialog, messagebox, Toplevel, Text, Scrollbar, RIGHT, Y
from tkinter.ttk import Progressbar
import traceback
import time

# =========================
# 設定：從第幾筆資料開始比對
# Excel 第2列 = pandas index 0（因為 header=0 會吃掉第1列欄名）
# =========================
START_DATA_INDEX = 0

# =========================
# 清洗工具（欄名/值統一）
# =========================
_CTRL_RE = re.compile(r"[\x00-\x1F\x7F]")

def clean_str(v):
    """去控制字元 + 去前後空白；None/NaN -> ''"""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass

    s = str(v)
    s = s.replace("\t", "").replace("\n", "").replace("\r", "")
    s = _CTRL_RE.sub("", s)
    return s.strip()

def clean_header_name(v):
    """欄名清洗後轉大寫，方便交集比對"""
    return clean_str(v).upper()

def normalize_key_value(col_name: str, value):
    """
    Key 專用正規化（你的案例最重要：VORNR）
    - VORNR：只留數字後補 4 碼（10 -> 0010）
    - 其他 key：一般清洗
    """
    s = clean_str(value)
    if s == "":
        return ""

    col = clean_header_name(col_name)

    if col == "VORNR":
        digits = re.sub(r"\D", "", s)   # 只保留數字
        if digits == "":
            return ""
        return digits.zfill(4)          # 補 4 碼

    return s

def build_clean_col_map(df: pd.DataFrame):
    """
    clean欄名 -> 原欄名 list（允許原始欄名重複）
    取第一個作為代表欄位
    """
    mp = {}
    for c in df.columns:
        cc = clean_header_name(c)
        mp.setdefault(cc, []).append(c)
    return mp

# =========================
# Log 視窗（彈窗，不輸出 txt）
# =========================
def show_log_window(log_text: str):
    w = Toplevel()
    w.title("執行 LOG")
    w.geometry("980x700")

    scrollbar = Scrollbar(w)
    scrollbar.pack(side=RIGHT, fill=Y)

    box = Text(w, wrap="word", font=("Consolas", 11))
    box.pack(expand=True, fill="both")

    box.insert("1.0", log_text)
    box.config(state="disabled")

    scrollbar.config(command=box.yview)
    box.config(yscrollcommand=scrollbar.set)

# =========================
# 執行中視窗
# =========================
def show_progress_window():
    win = Toplevel()
    win.title("處理中")
    win.geometry("420x160")
    win.resizable(False, False)

    lbl = ctk.CTkLabel(win, text="比對中，請稍候...", font=("Arial", 16, "bold"))
    lbl.pack(pady=(15, 8))

    pb = Progressbar(win, orient="horizontal", length=360, mode="determinate")
    pb.pack(pady=8)

    pct = ctk.CTkLabel(win, text="0%", font=("Arial", 14))
    pct.pack()

    info = ctk.CTkLabel(win, text="", font=("Arial", 12))
    info.pack(pady=(6, 0))

    win.update()
    return win, pb, pct, info

def update_progress(win, pb, pct, info, current, total, msg=""):
    if total <= 0:
        total = 1
    percent = int(min(100, (current / total) * 100))
    pb["value"] = percent
    pct.configure(text=f"{percent}%")
    info.configure(text=msg)
    win.update()

# =========================
# Key / 比對核心
# =========================
def make_key(df: pd.DataFrame, row_idx: int, key_cols):
    """
    key_cols 是欄位 index list
    key 回傳 tuple（固定順序）
    重要：用 normalize_key_value 做 Key 正規化
    """
    out = []
    for c in key_cols:
        col_name = str(df.columns[c])
        raw_val = df.iat[row_idx, c]
        out.append(normalize_key_value(col_name, raw_val))
    return tuple(out)

def build_key_map(df: pd.DataFrame, key_cols):
    """
    從 START_DATA_INDEX 開始
    Key 重複：保留第一次出現
    """
    m = {}
    for i in range(START_DATA_INDEX, len(df)):
        k = make_key(df, i, key_cols)
        if k not in m:
            m[k] = i
    return m

def count_duplicate_keys(df: pd.DataFrame, key_cols):
    keys = [make_key(df, i, key_cols) for i in range(START_DATA_INDEX, len(df))]
    s = pd.Series(keys)
    return int(s.duplicated(keep=False).sum())

def diff_directional(df_base, df_other, map_base, map_other, key_cols_base, base_name, other_name, log_lines):
    """
    以 base 為主：
    - base 有 key、other 沒有 -> Key不存在
    - base/other 都有 -> 比對交集欄位（清洗後欄名交集，排除 key 欄）
    產出 rows: [KEY..., 差異欄位, A值, B值, 差異來源]
    """
    keys_base = set(map_base.keys())
    keys_other = set(map_other.keys())

    base_cols_map = build_clean_col_map(df_base)
    other_cols_map = build_clean_col_map(df_other)

    # 排除 key 欄（用 base 的 key 欄清洗後名稱）
    base_key_clean = {clean_header_name(df_base.columns[i]) for i in key_cols_base}
    comparable_clean_cols = sorted((set(base_cols_map.keys()) & set(other_cols_map.keys())) - base_key_clean)

    rows = []
    diff_cells = 0
    missing_keys = 0

    # base 有、other 沒有
    for k in sorted(keys_base - keys_other):
        missing_keys += 1
        rows.append(list(k) + ["-", f"{base_name}=存在", f"{other_name}=不存在", "Key不存在"])

    # 兩邊都有 -> 比對值
    for k in sorted(keys_base & keys_other):
        ib = map_base[k]
        io = map_other[k]

        for cc in comparable_clean_cols:
            col_b = base_cols_map[cc][0]
            col_o = other_cols_map[cc][0]

            vb = clean_str(df_base.at[ib, col_b])
            vo = clean_str(df_other.at[io, col_o])

            if vb != vo:
                diff_cells += 1
                # 差異欄位顯示用 base 的原欄名
                rows.append(list(k) + [str(col_b), vb, vo, "值不同"])

    log_lines.append(f"[{base_name}→{other_name}] 可比對欄位數（排除Key後）={len(comparable_clean_cols)}")
    log_lines.append(f"[{base_name}→{other_name}] {other_name}缺少Key={missing_keys}；值不同(cell)={diff_cells}")
    return rows, diff_cells, missing_keys, len(comparable_clean_cols)

def build_column_diff(df_a: pd.DataFrame, df_b: pd.DataFrame):
    """
    產出欄位差異清單：
    - 清洗後欄名一致才視為同欄
    - 顯示 A/B 對應的原欄名（取第一個）
    """
    a_map = build_clean_col_map(df_a)
    b_map = build_clean_col_map(df_b)

    a_clean = set(a_map.keys())
    b_clean = set(b_map.keys())

    rows = []

    # A only
    for cc in sorted(a_clean - b_clean):
        rows.append([cc, str(a_map[cc][0]), "", "A有B沒有"])

    # B only
    for cc in sorted(b_clean - a_clean):
        rows.append([cc, "", str(b_map[cc][0]), "B有A沒有"])

    # both
    for cc in sorted(a_clean & b_clean):
        a_real = str(a_map[cc][0])
        b_real = str(b_map[cc][0])
        status = "一致"
        if a_real != b_real:
            status = "清洗後一致（原欄名不同）"
        rows.append([cc, a_real, b_real, status])

    return pd.DataFrame(rows, columns=["欄位(清洗後)", "Excel_A欄名", "Excel_B欄名", "狀態"])

# =========================
# GUI 變數
# =========================
file_a = None
file_b = None
output_folder = None

key_vars = {}
key_frame = None
label_a = None
label_b = None
label_out = None

# =========================
# GUI：選檔/選輸出
# =========================
def choose_file_a():
    global file_a
    file_a = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
    label_a.configure(text=os.path.basename(file_a) if file_a else "未選擇")
    if file_a:
        build_key_selector(file_a)

def choose_file_b():
    global file_b
    file_b = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
    label_b.configure(text=os.path.basename(file_b) if file_b else "未選擇")

def choose_output():
    global output_folder
    output_folder = filedialog.askdirectory()
    label_out.configure(text=output_folder if output_folder else "未選擇")

def build_key_selector(excel_path):
    """讀取 Excel A 欄名，動態建立 Key 勾選清單"""
    global key_vars, key_frame

    for w in key_frame.winfo_children():
        w.destroy()
    key_vars.clear()

    try:
        df_head = pd.read_excel(excel_path, nrows=0)
        cols = list(df_head.columns)
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取 Excel A 欄名失敗：{e}")
        return

    ctk.CTkLabel(
        key_frame,
        text="Key 欄位（勾選欄名，可多選）",
        font=("Arial", 14, "bold")
    ).pack(anchor="w", pady=(8, 4))

    ctk.CTkLabel(
        key_frame,
        text="提示：Routing 常見 Key = PLNNR + VORNR（本工具會自動把 VORNR 補成 4 碼，例如 10→0010）",
        font=("Arial", 12)
    ).pack(anchor="w", pady=(0, 10))

    # 預設：若有 PLNNR/VORNR 就勾，否則勾前兩欄
    default_checked = set()
    for c in cols:
        if clean_header_name(c) in {"PLNNR", "VORNR"}:
            default_checked.add(c)
    if not default_checked:
        default_checked = set(cols[:2])

    for c in cols:
        var = ctk.BooleanVar(value=(c in default_checked))
        key_vars[c] = var
        ctk.CTkCheckBox(key_frame, text=str(c), variable=var).pack(anchor="w", pady=2)

def get_selected_key_columns(df_a: pd.DataFrame, df_b: pd.DataFrame):
    selected = [k for k, v in key_vars.items() if v.get()]
    if not selected:
        raise ValueError("請至少勾選一個 Key 欄位")

    missing_in_b = [k for k in selected if k not in df_b.columns]
    if missing_in_b:
        raise ValueError(f"Excel B 缺少 Key 欄位：{missing_in_b}")

    key_cols_a = [df_a.columns.get_loc(k) for k in selected]
    key_cols_b = [df_b.columns.get_loc(k) for k in selected]
    return selected, key_cols_a, key_cols_b

# =========================
# 執行：比對 + 輸出
# =========================
def run_compare():
    log_lines = []
    t0 = time.time()
    win = None

    try:
        if not file_a or not file_b or not output_folder:
            messagebox.showerror("錯誤", "請選擇兩份 Excel（A、B）與匯出資料夾")
            return

        win, pb, pct, info = show_progress_window()
        step = 0
        total = 8

        def bump(msg):
            nonlocal step
            step += 1
            update_progress(win, pb, pct, info, step, total, msg)

        log_lines.append("=== Excel 差異比對 LOG（Key勾選 + VORNR補4碼）===")
        log_lines.append(f"[A] {file_a}")
        log_lines.append(f"[B] {file_b}")
        log_lines.append(f"[輸出資料夾] {output_folder}")
        log_lines.append(f"[資料起點] pandas index={START_DATA_INDEX}（Excel第2列）")
        log_lines.append("")

        bump("讀取 Excel A / B ...")
        df_a = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)

        log_lines.append(f"[A] Rows={len(df_a)} Cols={df_a.shape[1]}")
        log_lines.append(f"[B] Rows={len(df_b)} Cols={df_b.shape[1]}")
        log_lines.append("")

        bump("取得 Key 欄位（GUI 勾選）...")
        selected_keys, key_cols_a, key_cols_b = get_selected_key_columns(df_a, df_b)
        log_lines.append(f"[Key欄位] {selected_keys}")
        log_lines.append("（VORNR 會自動正規化補 4 碼）")
        log_lines.append("")

        bump("建立 Key Map ...")
        map_a = build_key_map(df_a, key_cols_a)
        map_b = build_key_map(df_b, key_cols_b)
        dup_a = count_duplicate_keys(df_a, key_cols_a)
        dup_b = count_duplicate_keys(df_b, key_cols_b)
        log_lines.append(f"[A] key筆數(去重)={len(map_a)}；重複Key列數={dup_a}")
        log_lines.append(f"[B] key筆數(去重)={len(map_b)}；重複Key列數={dup_b}")
        log_lines.append("")

        bump("欄位差異分析 ...")
        df_col_diff = build_column_diff(df_a, df_b)
        log_lines.append(f"[欄位差異] 共 {len(df_col_diff)} 筆（含一致/不一致/缺少）")
        log_lines.append("")

        bump("比對差異（A→B）...")
        a_to_b_rows, a_to_b_diffcells, a_to_b_missing, a_to_b_compcols = diff_directional(
            df_base=df_a, df_other=df_b,
            map_base=map_a, map_other=map_b,
            key_cols_base=key_cols_a,
            base_name="A", other_name="B",
            log_lines=log_lines
        )
        log_lines.append("")

        bump("比對差異（B→A）...")
        b_to_a_rows, b_to_a_diffcells, b_to_a_missing, b_to_a_compcols = diff_directional(
            df_base=df_b, df_other=df_a,
            map_base=map_b, map_other=map_a,
            key_cols_base=key_cols_b,
            base_name="B", other_name="A",
            log_lines=log_lines
        )
        log_lines.append("")

        bump("整理輸出資料 ...")
        key_headers = [f"KEY_{i+1}" for i in range(len(selected_keys))]
        headers = key_headers + ["差異欄位", "A值", "B值", "差異來源"]

        df_a_to_b = pd.DataFrame(a_to_b_rows, columns=headers)

        # B→A rows 目前是 [KEY..., 欄位, B值, A值, 來源]，轉成 A值/B值順序
        if b_to_a_rows:
            df_b_to_a_raw = pd.DataFrame(b_to_a_rows, columns=key_headers + ["差異欄位", "B值", "A值", "差異來源"])
            df_b_to_a = df_b_to_a_raw[key_headers + ["差異欄位", "A值", "B值", "差異來源"]]
        else:
            df_b_to_a = pd.DataFrame(columns=headers)

        # Summary
        summary_rows = [
            ["Key欄位", ", ".join([str(x) for x in selected_keys]), "", "", ""],
            ["A檔列/欄", f"{len(df_a)}/{df_a.shape[1]}", "", "", ""],
            ["B檔列/欄", f"{len(df_b)}/{df_b.shape[1]}", "", "", ""],
            ["A重複Key列數", dup_a, "", "", ""],
            ["B重複Key列數", dup_b, "", "", ""],
            ["A→B 可比對欄位數", a_to_b_compcols, "B缺少Key", a_to_b_missing, "值不同(cell)="+str(a_to_b_diffcells)],
            ["B→A 可比對欄位數", b_to_a_compcols, "A缺少Key", b_to_a_missing, "值不同(cell)="+str(b_to_a_diffcells)],
            ["A→B 差異列數", len(df_a_to_b), "", "", ""],
            ["B→A 差異列數", len(df_b_to_a), "", "", ""],
        ]
        df_summary = pd.DataFrame(summary_rows, columns=["項目", "值1", "值2", "值3", "值4"])

        bump("輸出 Excel ...")
        ts = time.strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(output_folder, f"Excel差異比對結果_{ts}.xlsx")

        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
            df_col_diff.to_excel(writer, sheet_name="ColumnDiff", index=False)
            df_a_to_b.to_excel(writer, sheet_name="A_to_B", index=False)
            df_b_to_a.to_excel(writer, sheet_name="B_to_A", index=False)

            # 基本可讀性：凍結列 + 篩選 + 欄寬
            for sh, df_tmp in [("Summary", df_summary), ("ColumnDiff", df_col_diff), ("A_to_B", df_a_to_b), ("B_to_A", df_b_to_a)]:
                ws = writer.sheets[sh]
                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, max(1, df_tmp.shape[0]), max(1, df_tmp.shape[1]-1))
                try:
                    for i, col in enumerate(df_tmp.columns):
                        max_len = max([len(str(col))] + [len(str(x)) for x in df_tmp[col].head(200).tolist()])
                        ws.set_column(i, i, min(60, max(12, max_len + 2)))
                except Exception:
                    pass

        bump("完成")
        duration = round(time.time() - t0, 2)
        log_lines.append(f"[輸出] {out_path}")
        log_lines.append(f"[耗時] {duration} 秒")
        log_lines.append("")
        log_lines.append(f"[A_to_B] 差異列數={len(df_a_to_b)}")
        log_lines.append(f"[B_to_A] 差異列數={len(df_b_to_a)}")
        log_lines.append(f"[ColumnDiff] 筆數={len(df_col_diff)}")

        if win:
            win.destroy()

        show_log_window("\n".join(log_lines))
        messagebox.showinfo("完成", f"差異比對完成 ✅\n\n{out_path}")

    except Exception as e:
        if win:
            try:
                win.destroy()
            except Exception:
                pass

        log_lines.append("")
        log_lines.append("!!! 發生錯誤 !!!")
        log_lines.append(str(e))
        log_lines.append(traceback.format_exc())

        show_log_window("\n".join(log_lines))
        messagebox.showerror("錯誤", f"{e}\n\n{traceback.format_exc()}")

# =========================
# GUI 主畫面
# =========================
def main_gui():
    global label_a, label_b, label_out, key_frame

    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Excel 差異比對工具（Key勾選 + VORNR補4碼）")
    root.geometry("820x760")

    ctk.CTkLabel(root, text="Excel 差異比對工具", font=("Arial", 22, "bold")).pack(pady=(12, 6))
    ctk.CTkLabel(
        root,
        text="1) 選 Excel A（載入欄名供 Key 勾選）  2) 選 Excel B  3) 選匯出資料夾  4) 開始比對",
        font=("Arial", 12)
    ).pack(pady=(0, 10))

    top = ctk.CTkFrame(root)
    top.pack(fill="x", padx=20, pady=8)

    btn_a = ctk.CTkButton(top, text="選擇 Excel A", command=choose_file_a, width=200)
    btn_a.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    label_a = ctk.CTkLabel(top, text="未選擇")
    label_a.grid(row=0, column=1, padx=10, pady=10, sticky="w")

    btn_b = ctk.CTkButton(top, text="選擇 Excel B", command=choose_file_b, width=200)
    btn_b.grid(row=1, column=0, padx=10, pady=10, sticky="w")
    label_b = ctk.CTkLabel(top, text="未選擇")
    label_b.grid(row=1, column=1, padx=10, pady=10, sticky="w")

    btn_out = ctk.CTkButton(top, text="選擇匯出資料夾", command=choose_output, width=200)
    btn_out.grid(row=2, column=0, padx=10, pady=10, sticky="w")
    label_out = ctk.CTkLabel(top, text="未選擇")
    label_out.grid(row=2, column=1, padx=10, pady=10, sticky="w")

    top.grid_columnconfigure(1, weight=1)

    ctk.CTkLabel(root, text="Key 欄位設定（勾選欄名）", font=("Arial", 16, "bold")).pack(pady=(10, 5))
    key_frame = ctk.CTkScrollableFrame(root, width=760, height=360)
    key_frame.pack(fill="both", expand=False, padx=20, pady=(0, 10))

    ctk.CTkLabel(
        key_frame,
        text="請先選擇 Excel A，這裡會載入欄名並提供 Key 勾選。",
        font=("Arial", 13)
    ).pack(anchor="w", pady=10)

    btn_run = ctk.CTkButton(
        root,
        text="開始差異比對並匯出 Excel",
        fg_color="#28a745",
        hover_color="#1e7e34",
        font=("Arial", 22, "bold"),
        height=60,
        command=run_compare
    )
    btn_run.pack(pady=18)

    root.mainloop()

if __name__ == "__main__":
    main_gui()