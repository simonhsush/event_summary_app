# app.py
import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
import io
import re
from dateutil import parser as dateparser
import docx
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

st.set_page_config(page_title="文件日期篩選摘要器", layout="wide")
st.title("📄 Word 文件 — 找出含指定日期的欄位並產生摘要檔")

st.markdown(
    """
上傳 `.docx`（可含表格或段落），選擇：
- 使用「前一個工作日（前一個營業日）」或輸入**指定日期**，
- 選擇要比對的欄位名稱（若 Word 內為表格會列出欄位供選），
程式會找出在該欄位內含有目標日期（或日期字串）的列，並產生摘要檔下載。
"""
)

# -----------------------
# Helpers (保留你原來的日期處理邏輯)
# -----------------------
def prev_business_day(ref_date=None):
    if ref_date is None:
        ref = datetime.now()
    else:
        ref = ref_date
    one = timedelta(days=1)
    d = ref - one
    while d.weekday() >= 5:
        d -= one
    return d.date()

date_regex = re.compile(
    r'(?:(?:\d{4}[/-]\d{1,2}[/-]\d{1,2})|(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4})|(?:\d{1,2}[\u4e00-\u9fff]{1}\d{1,2}[\u4e00-\u9fff]{0,1}\d{0,2}))'
)

def extract_tables_to_dfs(doc):
    dfs = []
    for t in doc.tables:
        rows = []
        for r in t.rows:
            rows.append([c.text.strip() for c in r.cells])
        if len(rows) >= 2:
            header = rows[0]
            data = rows[1:]
            try:
                df = pd.DataFrame(data, columns=header)
            except Exception:
                df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(rows)
        dfs.append(df)
    return dfs

def find_date_like_in_text(text):
    found = []
    for m in date_regex.findall(text):
        s = m
        try:
            dt = dateparser.parse(s, dayfirst=False, fuzzy=True)
            if dt:
                found.append(dt.date())
        except Exception:
            continue
    return found

def filter_df_by_date_in_column(df, column, target_date):
    if column not in df.columns:
        return pd.DataFrame()
    matches = []
    for idx, cell in df[column].fillna("").items():
        text = str(cell)
        td_strs = [
            target_date.strftime("%Y-%m-%d"),
            target_date.strftime("%Y/%m/%d"),
            target_date.strftime("%Y%m%d"),
            target_date.strftime("%m/%d"),
            target_date.strftime("%#m/%#d/%Y") if hasattr(target_date, 'strftime') else ""
        ]
        td_chinese = f"{target_date.year}年{target_date.month}月{target_date.day}日"
        found = False
        if any(s in text for s in td_strs) or td_chinese in text:
            found = True
        else:
            parsed = find_date_like_in_text(text)
            if parsed and any(d == target_date for d in parsed):
                found = True
        if found:
            matches.append((idx, cell))
    if not matches:
        return pd.DataFrame()
    return df.loc[[idx for idx, _ in matches]]

# -----------------------
# 修改重點： export_to_word 與段落擷取邏輯
# 我只改動輸出內容：改為「日期字串出現位置開始，往後取 num_chars 個字」
# 並且在 export_to_word 中反白整段（日期 + 擷取內容）
# -----------------------
def export_to_word(data, target_date_str, num_chars, filename="摘要輸出.docx"):
    """
    data: DataFrame，需包含欄位 "text"（文字內容）或 若無則會把整列 join 成文字
    target_date_str: 用來在段落中找到日期位置的字串表示（例如 "2025-10-22"）
    num_chars: 從日期之後取的字數（int）
    """
    doc = Document()
    doc.add_heading("搜尋摘要結果", level=1)
    doc.add_paragraph(f"搜尋日期關鍵字：{target_date_str}")
    doc.add_paragraph(" ")

    # 確保有 text 欄位，若沒有則把整列 join 為文字
    if "text" not in data.columns:
        data = data.copy()
        data["text"] = data.apply(lambda r: " ".join([str(x) for x in r.values if pd.notna(x)]), axis=1)

    for idx, row in enumerate(data["text"].astype(str).tolist(), start=1):
        text = row
        # 嘗試用 target_date_str 找到位置；若找不到就直接取開頭 num_chars（但原則是找得到才處理）
        start_idx = text.find(target_date_str)
        # 若找不到完整 target_date_str（例如使用中文日期），嘗試其他形式（像 2025年10月22日）
        if start_idx == -1:
            # 嘗試中文年月日形式
            try:
                # 解析 yyyy-mm-dd
                parts = re.findall(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', target_date_str)
                if parts:
                    y, m, d = parts[0]
                    chinese = f"{int(y)}年{int(m)}月{int(d)}日"
                    start_idx = text.find(chinese)
                    target_for_slice = chinese
                else:
                    target_for_slice = target_date_str
            except Exception:
                target_for_slice = target_date_str
        else:
            target_for_slice = target_date_str

        if start_idx != -1:
            # 計算 slice 範圍：從日期開始 (包含日期本身) 往後取 num_chars 個字
            slice_start = start_idx
            slice_end = min(len(text), slice_start + len(target_for_slice) + num_chars)
            snippet = text[slice_start:slice_end]
        else:
            # 若完全找不到，則 snippet 設為空字串（之後會被過濾掉或顯示找不到）
            snippet = ""

        # 若 snippet 非空，新增段落並反白整段
        p = doc.add_paragraph(f"{idx}. ")
        if snippet:
            run = p.add_run(snippet)
            # 反白整段（使用 WD_COLOR_INDEX 黃）
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            p.add_run("(未找到可擷取的段落)")

    doc.save(filename)
    return filename

# -----------------------
# UI: 上傳與選項（保留原設定）
# -----------------------
uploaded_file = st.file_uploader("上傳 Word (.docx) 檔案", type=["docx"])
st.write("（檔案內容只在本次執行中處理，不會上傳到任何外部伺服器）")

# 這裡是你要讓使用者設定要取幾個字：我保留原本的 num_chars 控制
num_chars = st.number_input("請輸入要擷取的字數（從日期字串之後開始計算，預設20）", min_value=1, max_value=1000, value=20)

col1, col2 = st.columns([2, 1])
with col1:
    date_mode = st.radio("選擇目標日期：", ("前一個工作日", "輸入指定日期 (YYYY-MM-DD)"))
    if date_mode == "輸入指定日期 (YYYY-MM-DD)":
        user_date_str = st.text_input("指定日期 (例: 2025-10-22)", value="")
        try:
            user_date = dateparser.parse(user_date_str).date() if user_date_str.strip() else None
        except Exception:
            user_date = None
    else:
        user_date = None

with col2:
    st.write("進階選項")
    prefer_table = st.checkbox("優先從表格欄位篩選 (若檔案含表格則列出欄位)", value=True)
    download_format = st.selectbox("下載摘要格式", ["CSV", "純文字 (TXT)", "Word (.docx)"])

# -----------------------
# 處理檔案（主流程，保留你原本邏輯）
# -----------------------
if uploaded_file is not None:
    try:
        doc = docx.Document(uploaded_file)
    except Exception as e:
        st.error(f"無法讀取 Word 檔：{e}")
        st.stop()

    # 決定 target_date
    if date_mode == "前一個工作日":
        target_date = prev_business_day()
        # target_date_str 我們以 isoformat 傳遞給後續使用
        target_date_str = target_date.isoformat()
    else:
        if user_date is None:
            st.warning("請輸入有效的指定日期（YYYY-MM-DD）。")
            st.stop()
        target_date = user_date
        # target_date_str = target_date.isoformat()
        target_date_str = str(target_date)

    st.info(f"將比對的目標日期： {target_date_str}")

    # 嘗試從表格抓 dataframe
    dfs = extract_tables_to_dfs(doc)
    result_rows = []

    if dfs and prefer_table:
        st.write(f"偵測到 {len(dfs)} 個表格，正在掃描表格欄位...")
        all_cols = set()
        for df in dfs:
            all_cols.update(list(df.columns))
        all_cols = [c for c in all_cols if str(c).strip() != ""]
        if all_cols:
            chosen_cols = st.multiselect("選擇要比對的欄位（表格欄位）", options=all_cols, default=all_cols[:2])
        else:
            chosen_cols = []
        if chosen_cols:
            for i, df in enumerate(dfs):
                df = df.astype(str)
                for col in chosen_cols:
                    filtered = filter_df_by_date_in_column(df, col, target_date)
                    if not filtered.empty:
                        filtered = filtered.copy()
                        # 重要改動：在 table 的結果中，我們只取「從日期開始的 snippet」放到 text 欄位
                        snippets = []
                        for _, r in filtered.iterrows():
                            cell_text = str(r[col])
                            # 優先找多種日期字串（保留原邏輯格式）
                            # 先嘗試 yyyy-mm-dd
                            td_candidates = [
                                target_date.strftime("%Y-%m-%d"),
                                target_date.strftime("%Y/%m/%d"),
                                f"{target_date.year}年{target_date.month}月{target_date.day}日"
                            ]
                            start_idx = -1
                            chosen_td = None
                            for td in td_candidates:
                                if td in cell_text:
                                    start_idx = cell_text.find(td)
                                    chosen_td = td
                                    break
                            if start_idx != -1:
                                end_idx = min(len(cell_text), start_idx + len(chosen_td) + num_chars)
                                snippet = cell_text[start_idx:end_idx]
                            else:
                                snippet = "" # 若找不到，留空（會被後面過濾）
                            snippets.append(snippet)
                        # 建 DataFrame 保留來源欄位，但把 text 欄位改為 snippet（只含日期後 N 字）
                        filtered = filtered.reset_index(drop=True)
                        filtered["text"] = snippets
                        filtered["_source_table"] = f"table_{i+1}"
                        filtered["_matched_column"] = col
                        # 只保留有 snippet 的列
                        filtered = filtered[filtered["text"].str.strip() != ""]
                        if not filtered.empty:
                            result_rows.append(filtered)
        else:
            st.write("未選擇任何欄位 — 跳過表格欄位篩選。")

    # 若沒有結果或使用者不偏好表格，從段落中搜尋
    if not result_rows:
        st.write("從段落中搜尋含有目標日期的文字（包含 '日期:'、'Date:' 等）...")
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        para_matches = []

        for i, txt in enumerate(paragraphs):
            parsed = find_date_like_in_text(txt)
            if parsed and any(d == target_date for d in parsed):
                # 用字串比對位置，保留你原本的日期格式邏輯（嘗試多種形式）
                td_candidates = [
                    target_date.strftime("%Y-%m-%d"),
                    target_date.strftime("%Y/%m/%d"),
                    f"{target_date.year}年{target_date.month}月{target_date.day}日"
                ]
                start_idx = -1
                chosen_td = None
                for td in td_candidates:
                    if td in txt:
                        start_idx = txt.find(td)
                        chosen_td = td
                        break
                if start_idx != -1:
                    end_idx = min(len(txt), start_idx + len(chosen_td) + num_chars)
                    snippet = txt[start_idx:end_idx]
                    para_matches.append((i, snippet))

        if para_matches:
            dfp = pd.DataFrame(para_matches, columns=["para_index", "text"])
            dfp["_source_table"] = "paragraphs"
            dfp["_matched_column"] = "text"
            result_rows.append(dfp)

    # 合併並顯示結果（注意：現在 text 欄位皆為 snippet）
    if result_rows:
        final = pd.concat(result_rows, ignore_index=True, sort=False)
        # 避免空 snippet
        final = final[final["text"].astype(str).str.strip() != ""].reset_index(drop=True)

        st.subheader("找到的結果範例（前 50 列）")
        st.dataframe(final.head(50))
        st.write(f"共找到 {len(final)} 筆符合目標日期 ({target_date_str}) 的項目。")

        # -------- CSV 匯出 --------
        if download_format == "CSV":
            towrite = io.StringIO()
            # 我們只輸出有用的欄位（保留原有欄位並優先輸出 text 欄）
            final.to_csv(towrite, index=False, encoding="utf-8-sig")
            st.download_button(
                "下載 CSV 檔（UTF-8-SIG）",
                data=towrite.getvalue().encode("utf-8-sig"),
                file_name=f"summary_{target_date_str}.csv",
                mime="text/csv"
            )

        # -------- TXT 匯出 --------
        elif download_format == "純文字 (TXT)":
            txt_buf = io.StringIO()
            for i, row in final.iterrows():
                # 只輸出 snippet（text）與來源說明
                txt_buf.write(f"來源: {row.get('_source_table','')}, 欄位: {row.get('_matched_column','')}\n")
                txt_buf.write(f"{row.get('text','')}\n")
                txt_buf.write("----\n")
            st.download_button(
                "下載 TXT 檔",
                data=txt_buf.getvalue().encode("utf-8"),
                file_name=f"summary_{target_date_str}.txt",
                mime="text/plain"
            )

        # -------- Word 匯出 --------
        elif download_format == "Word (.docx)":
            # 匯出，反白整段 snippet（export_to_word 已經實作）
            export_to_word(final, target_date_str, num_chars, filename="日期段落摘要.docx")
            out_doc = docx.Document("日期段落摘要.docx")
            word_stream = io.BytesIO()
            out_doc.save(word_stream)
            word_stream.seek(0)
            st.download_button(
                "下載 Word 摘要檔",
                data=word_stream,
                file_name=f"summary_{target_date_str}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("沒有找到符合條件的項目。請確認：\n- Word 是否含有表格或段落中是否有日期字串。\n- 若日期格式特殊，可嘗試手動輸入精確日期字串作為比對條件。")

