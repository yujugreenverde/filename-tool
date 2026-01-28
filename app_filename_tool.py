# app_filename_tool.py
# ------------------------------------------------------------
# Streamlit 檔案命名工具（雲端版 / 網頁版）
# vCloud-1.0（覆蓋版）
#
# ✅ 單張表格（24列）：可編輯 widget + 欄位7即時更新 + 重複檔名紅底 ⚠
# ✅ 產生「old → new」對照表：
#    - 支援「上傳檔案（只讀檔名）」或「貼上檔名清單」
#    - 依排序方式配對（自然排序 / 反向 / 原始順序）
#    - 下載對照表（Excel / CSV）
# ✅ 匯出目標清單 Excel（Use=✅）
#
# ⚠️ 雲端限制：
# - 無法直接改你本機資料夾的檔名（因此本版不提供直接改名功能）
#
# 執行：
#   streamlit run app_filename_tool.py
# ------------------------------------------------------------

import os
import re
from datetime import date
from io import BytesIO
from collections import Counter

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# ----------------------------
# Helpers
# ----------------------------
VIDEO_EXTS = (".mp4", ".avi", ".mov", ".mkv", ".m4v", ".wmv")
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp")

def yyyymmdd(d: date) -> str:
    return d.strftime("%Y%m%d")

def natural_key(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]

def clipboard_button(text: str, key: str):
    safe = text.replace("\\", "\\\\").replace("`", "\\`")
    components.html(
        f"""
        <button onclick="navigator.clipboard.writeText(`{safe}`)"
        style="padding:4px 8px;border-radius:6px;border:1px solid #ccc;background:#fff;cursor:pointer;font-size:12px;">
        Copy
        </button>
        """,
        height=32,
    )

def select_or_custom(prefix, options, default, i):
    sel_k = f"{prefix}_sel_{i}"
    cus_k = f"{prefix}_cus_{i}"

    if sel_k not in st.session_state:
        st.session_state[sel_k] = default
    if cus_k not in st.session_state:
        st.session_state[cus_k] = ""

    sel = st.selectbox(
        "",
        options + ["Custom"],
        index=(options + ["Custom"]).index(st.session_state[sel_k]),
        key=sel_k,
        label_visibility="collapsed",
    )

    if sel == "Custom":
        val = st.text_input("", key=cus_k, placeholder="Type", label_visibility="collapsed")
        return val.strip()
    return sel

def html_filename(text: str, is_dup: bool = False) -> str:
    safe = (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    if not safe:
        return ""
    if is_dup:
        return (
            f"<span style='font-family:ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;"
            f"background:#fff0f0;border:1px solid #ffb3b3;color:#7a0000;padding:2px 6px;border-radius:6px;'>"
            f"⚠ {safe}</span>"
        )
    return (
        f"<span style='font-family:ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;"
        f"background:#f6f8fa;border:1px solid #e5e7eb;color:#111827;padding:2px 6px;border-radius:6px;'>"
        f"{safe}</span>"
    )

def parse_pasted_list(text: str) -> list[str]:
    # split by newline / tab / comma; keep non-empty
    if not text:
        return []
    parts = re.split(r"[\r\n\t,]+", text.strip())
    return [p.strip() for p in parts if p.strip()]

def file_type_filter(names: list[str], file_type: str) -> list[str]:
    exts = VIDEO_EXTS if file_type == "Video" else IMAGE_EXTS
    out = []
    for n in names:
        low = n.lower()
        if low.endswith(exts):
            out.append(n)
    return out

def build_excel_bytes(df: pd.DataFrame, sheet_name: str, file_name_hint: str) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf.getvalue()


# ----------------------------
# App
# ----------------------------
st.set_page_config(page_title="Filename Tool (Cloud)", layout="wide")
st.title("📁 檔案命名工具（雲端版）")

st.info(
    "雲端版提供：檔名生成 / Copy / 匯出 Excel / 產生 old→new 對照表。\n"
    "雲端無法直接存取你的本機資料夾，因此不提供「直接改名」功能。"
)

# Top controls
c0, c1, c2 = st.columns([1.2, 1.6, 3.2])
with c0:
    file_type = st.selectbox("0. 檔案類型", ["Video", "Image"])
with c1:
    d = st.date_input("1. 日期", value=date.today())
date_str = yyyymmdd(d)
with c2:
    st.markdown(f"**日期字串：** `{date_str}`")

# Presets
STRAIN_OPTS = ["CS", "Y[1]"]
SEX_OPTS = ["M", "F", "Mix"]
TREAT_OPTS = ["Exp", "Ctrl"]
STAGE_OPTS = ["Te", "Tr1", "Tr2", "Tr3", "Tr4"]

# Controls
cA, cB, cC, cD = st.columns([1.2, 1.2, 2.0, 2.6])
with cA:
    if st.button("✅ 全選 Use"):
        for i in range(1, 25):
            st.session_state[f"use_{i}"] = True
with cB:
    if st.button("❌ 全部取消"):
        for i in range(1, 25):
            st.session_state[f"use_{i}"] = False
with cC:
    renumber = st.checkbox("順位依 Use=✅ 自動重排", value=True)
with cD:
    dup_scope = st.radio(
        "重複檔名檢查範圍",
        ["只看 Use=✅", "看全部（有檔名者）"],
        horizontal=True,
        index=0,
    )

st.divider()

# Ensure defaults
for i in range(1, 25):
    st.session_state.setdefault(f"use_{i}", (i == 1))
    st.session_state.setdefault(f"sex_{i}", "Mix")
    st.session_state.setdefault(f"strain_sel_{i}", "CS")
    st.session_state.setdefault(f"strain_cus_{i}", "")
    st.session_state.setdefault(f"treat_sel_{i}", "Exp")
    st.session_state.setdefault(f"treat_cus_{i}", "")
    st.session_state.setdefault(f"stage_sel_{i}", "Te")
    st.session_state.setdefault(f"stage_cus_{i}", "")

# Pre-pass for duplicate base detection (based on current state)
def current_row_values(i: int):
    use = bool(st.session_state.get(f"use_{i}", False))
    sex = st.session_state.get(f"sex_{i}", "Mix")

    strain_sel = st.session_state.get(f"strain_sel_{i}", "CS")
    strain_cus = (st.session_state.get(f"strain_cus_{i}", "") or "").strip()
    strain = strain_cus if strain_sel == "Custom" else strain_sel

    treat_sel = st.session_state.get(f"treat_sel_{i}", "Exp")
    treat_cus = (st.session_state.get(f"treat_cus_{i}", "") or "").strip()
    treat = treat_cus if treat_sel == "Custom" else treat_sel

    stage_sel = st.session_state.get(f"stage_sel_{i}", "Te")
    stage_cus = (st.session_state.get(f"stage_cus_{i}", "") or "").strip()
    stage = stage_cus if stage_sel == "Custom" else stage_sel

    return use, strain, sex, treat, stage

used_idx = 0
bases_for_dup = []
for i in range(1, 25):
    use, strain, sex, treat, stage = current_row_values(i)
    if renumber and use:
        used_idx += 1
        order = f"{used_idx:02d}"
    elif renumber and not use:
        order = ""
    else:
        order = f"{i:02d}"
    base = f"{date_str}_{order}_{strain}_{sex}_{treat}_{stage}" if order else ""
    if base:
        if dup_scope == "只看 Use=✅":
            if use:
                bases_for_dup.append(base)
        else:
            bases_for_dup.append(base)

dup_bases = {k for k, v in Counter(bases_for_dup).items() if v > 1}

# Single editable table
h = st.columns([0.6, 0.7, 1.0, 1.3, 1.0, 1.2, 1.2, 3.4, 0.8])
for col, name in zip(h, ["#", "Use", "順位", "品系", "性別", "處理", "階段", "欄位7：檔名", "Copy"]):
    col.markdown(f"**{name}**")

rows_out = []
used_idx = 0
for i in range(1, 25):
    c = st.columns([0.6, 0.7, 1.0, 1.3, 1.0, 1.2, 1.2, 3.4, 0.8])
    c[0].markdown(f"`{i:02d}`")

    with c[1]:
        use = st.checkbox("", key=f"use_{i}", label_visibility="collapsed")

    if renumber and use:
        used_idx += 1
        order = f"{used_idx:02d}"
    elif renumber and not use:
        order = ""
    else:
        order = f"{i:02d}"

    c[2].markdown(f"`{order}`" if order else "")

    with c[3]:
        strain = select_or_custom("strain", STRAIN_OPTS, "CS", i)
    with c[4]:
        sex = st.selectbox("", SEX_OPTS, key=f"sex_{i}", label_visibility="collapsed")
    with c[5]:
        treat = select_or_custom("treat", TREAT_OPTS, "Exp", i)
    with c[6]:
        stage = select_or_custom("stage", STAGE_OPTS, "Te", i)

    filename_base = f"{date_str}_{order}_{strain}_{sex}_{treat}_{stage}" if order else ""
    is_dup = filename_base in dup_bases and bool(filename_base)

    with c[7]:
        st.markdown(html_filename(filename_base, is_dup=is_dup), unsafe_allow_html=True)

    with c[8]:
        if filename_base:
            clipboard_button(filename_base, f"cp_{i}")

    rows_out.append(
        {
            "row": i,
            "use": bool(use),
            "order": order,
            "strain": strain,
            "sex": sex,
            "treatment": treat,
            "stage": stage,
            "filename_base": filename_base,
        }
    )

df_ui = pd.DataFrame(rows_out)

if dup_bases:
    st.warning("⚠️ 偵測到重複檔名（未含副檔名）。已在欄位7以紅底 ⚠ 高亮。")

st.divider()

# ----------------------------
# Export targets (Use=✅)
# ----------------------------
st.subheader("⬇️ 匯出目標清單 Excel（Use=✅）")
df_targets = df_ui[df_ui["use"]].copy()
df_targets = df_targets[df_targets["filename_base"].astype(str).str.len() > 0].copy()

df_export_targets = df_targets[["row", "order", "strain", "sex", "treatment", "stage", "filename_base"]].copy()

targets_xlsx = build_excel_bytes(df_export_targets, "targets", f"{date_str}_{file_type}")
st.download_button(
    "匯出 targets.xlsx",
    data=targets_xlsx,
    file_name=f"{date_str}_{file_type}_targets.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.divider()

# ----------------------------
# Old -> New mapping
# ----------------------------
st.subheader("🔁 生成 old → new 對照表（雲端可用）")
st.caption("你提供「舊檔名清單」→ 我用上方 Use=✅ 生成的檔名配對 → 下載對照表（Excel/CSV）。")

m1, m2 = st.columns([2.2, 1.8])
with m1:
    src_mode = st.radio("舊檔名來源", ["上傳檔案（只取檔名）", "貼上檔名清單"], horizontal=True)
with m2:
    pair_mode = st.selectbox("配對順序", ["自然排序（推薦）", "原始順序", "反向（倒序）"], index=0)

old_names: list[str] = []

if src_mode == "上傳檔案（只取檔名）":
    uploads = st.file_uploader(
        "拖曳/選取檔案（可多選）",
        accept_multiple_files=True,
        type=None,  # allow any; we will filter by file_type if user chooses
    )
    if uploads:
        old_names = [u.name for u in uploads]
else:
    pasted = st.text_area(
        "貼上檔名（每行一個；也可用逗號/Tab 分隔）",
        height=160,
        placeholder="例如：\nIMG_0001.JPG\nIMG_0002.JPG\n...\n",
    )
    old_names = parse_pasted_list(pasted)

# optional filter by file type
filter_by_type = st.checkbox(f"只保留符合 {file_type} 的副檔名", value=True)
if filter_by_type and old_names:
    before = len(old_names)
    old_names = file_type_filter(old_names, file_type)
    if len(old_names) != before:
        st.caption(f"已依檔案類型過濾：{before} → {len(old_names)}")

# ordering
if old_names:
    if pair_mode == "自然排序（推薦）":
        old_names_sorted = sorted(old_names, key=natural_key)
    elif pair_mode == "反向（倒序）":
        old_names_sorted = list(reversed(sorted(old_names, key=natural_key)))
    else:
        old_names_sorted = old_names[:]  # preserve input order
else:
    old_names_sorted = []

# build mapping
mapping_rows = []
n_old = len(old_names_sorted)
n_new = len(df_targets)

if n_old == 0:
    st.info("先提供舊檔名（上傳或貼上）才能生成對照表。")
elif n_new == 0:
    st.warning("上方 Use=✅ 目前沒有可用的目標檔名。請先勾選 Use 並確認欄位7生成。")
else:
    if n_old != n_new:
        st.warning(f"數量不一致：old={n_old} / new(Use=✅)={n_new}。將只配對前 min(old,new) 筆。")

    n = min(n_old, n_new)
    new_bases = df_targets["filename_base"].tolist()

    for i in range(n):
        old = old_names_sorted[i]
        ext = os.path.splitext(old)[1]  # keep original extension
        new = f"{new_bases[i]}{ext}"
        mapping_rows.append(
            {
                "idx": i + 1,
                "old_name": old,
                "new_name": new,
                "new_base": new_bases[i],
                "ext": ext,
            }
        )

df_map = pd.DataFrame(mapping_rows)

# validate duplicates in new_name
dup_new = set()
if not df_map.empty:
    dup_new = {k for k, v in Counter(df_map["new_name"].tolist()).items() if v > 1}
    if dup_new:
        st.error("❌ new_name（含副檔名）出現重複，請回上方調整條件或 Use 選擇。")

# preview table
if not df_map.empty:
    st.markdown("### 對照表預覽（old → new）")
    ph = st.columns([0.6, 3.2, 3.6])
    ph[0].markdown("**#**")
    ph[1].markdown("**old_name**")
    ph[2].markdown("**new_name**")

    for _, r in df_map.iterrows():
        pc = st.columns([0.6, 3.2, 3.6])
        pc[0].markdown(f"`{int(r['idx']):02d}`")
        pc[1].markdown(f"`{r['old_name']}`")
        with pc[2]:
            st.markdown(html_filename(r["new_name"], is_dup=(r["new_name"] in dup_new)), unsafe_allow_html=True)

    # downloads
    map_xlsx = build_excel_bytes(df_map, "mapping", f"{date_str}_{file_type}_old2new")
    st.download_button(
        "下載對照表 Excel（old2new.xlsx）",
        data=map_xlsx,
        file_name=f"{date_str}_{file_type}_old2new.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=bool(dup_new),
    )

    csv_bytes = df_map.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        "下載對照表 CSV（old2new.csv）",
        data=csv_bytes,
        file_name=f"{date_str}_{file_type}_old2new.csv",
        mime="text/csv",
        disabled=bool(dup_new),
    )

st.divider()
st.caption("小提醒：雲端版只產生對照表；真正批次改名可以用你本機的檔案管理器或另寫一個小腳本讀取 old2new.csv 來改名。")