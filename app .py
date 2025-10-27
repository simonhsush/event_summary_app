import streamlit as st
import pandas as pd
import re
from datetime import datetime
from docx import Document
from io import BytesIO

# ---------------------------
# 🔹 日期處理：支援民國年、全形、底線、中文日期
# ---------------------------
def parse_to_western_date(date_str):
    date_str = date_str.strip()
    date_str = date_str.replace("／", "/").replace("－", "-").replace("_", "/")

    # 民國年轉換
    match = re.match(r"^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})$", date_str)
    if match:
        year, month, day = map(int, match.groups())
        if year < 200: # 民國年
            year += 1911
        try:
            return datetime(year, month, day)
        except ValueError:
            return None

    # 西元或其他格式
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%m/%d", "%m-%d", "%m月%d日"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def find_date_like_in_text(text):
    """找出文字中所有可能日期"""
    text = text.replace("／", "/").replace("－", "-").replace("_", "/")
    patterns = [
        r"\d{3,4}[/-]\d{1,2}[/-]\d{1,2}",
        r"\d{1,2}[/-]\d{1,2}",
        r"\d{1,2}月\d{1,2}日"
    ]
    found_dates = []
    for pat in patterns:
        for m in re.finditer(pat, text):
            d = parse_to_western_date(m.group())
            if d:
                found_dates.append(d)
    return found_dates


def filter_df_by_date_in_column(df, column, target_date):
    """在特定欄位中搜尋日期"""
    if column not in df.columns:
        return pd.DataFrame()
    matches = []
    for idx, cell in df[column].fillna("").items():
        text = str(cell)
        normalized_text = text.replace("／", "/").replace("－", "-").replace("_", "/")
        found = False

        td_strs = [
            target_date.strftime("%Y-%m-%d"),
            target_date.strftime("%Y/%m/%d"),
            f"{target_date.month}/{target_date.day}",
            f"{target_date.month:02d}/{target_date.day:02d}",
            f"{target_date.month}-{target_date.day}",
        ]
        td_chinese = f"{target_date.month}月{target_date.day}日"

        if any(s in normalized_text for s in td_strs) or td_chinese in normalized_text:
            found = True
        else:
            parsed = find_date_like_in_text(normalized_text)
            for d in parsed:
                if d.year == target_date.year and d.month == target_date.month and d.day == target_date.day:
                    found = True
                    break

        if found:
            matches.append(idx)

    return df.loc[matches]


# ---------------------------
# 🔹 Streamlit 主介面
# ---------------------------
st.title("📄 日期段落擷取工具（支援民國、西元與中文格式）")

uploaded_file = st.file_uploader("請上傳 Word 檔（.docx）", type=["docx"])
if uploaded_file:
    doc = Document(uploaded_file)

    # 取出所有表格
    dfs = []
    for t in doc.tables:
        data = []
        for row in t.rows:
            data.append([cell.text.strip() for cell in row.cells])
        df = pd.DataFrame(data)
        df.columns = df.iloc[0]
        df = df[1:]
        dfs.append(df)
    if not dfs:
        st.error("❌ 未找到表格內容，請確認檔案格式。")
        st.stop()

    df_all = pd.concat(dfs, ignore_index=True)
    st.success(f"✅ 已載入表格，共 {len(df_all)} 列。")

    # 日期與設定
    date_str = st.text_input("輸入日期（例如：114/10/23、2025/10/23 或 10/23）")
    num_chars = st.number_input("擷取該日期後幾個字：", min_value=5, max_value=200, value=20, step=1)
    colname = st.selectbox("選擇比對欄位", df_all.columns)

    if st.button("執行擷取"):
        target_date = parse_to_western_date(date_str)
        if not target_date:
            st.error("❌ 無法辨識日期格式，請重新輸入。")
            st.stop()

        result_df = filter_df_by_date_in_column(df_all, colname, target_date)
        if result_df.empty:
            st.warning("⚠️ 找不到符合日期的資料。")
        else:
            # 擷取指定日期後的文字
            def extract_text(row):
                text = str(row[colname])
                normalized = text.replace("／", "/").replace("－", "-").replace("_", "/")
                for pat in [r"\d{3,4}[/-]\d{1,2}[/-]\d{1,2}", r"\d{1,2}[/-]\d{1,2}", r"\d{1,2}月\d{1,2}日"]:
                    for m in re.finditer(pat, normalized):
                        d = parse_to_western_date(m.group())
                        if d and d.month == target_date.month and d.day == target_date.day:
                            start = m.start()
                            return normalized[start:start + num_chars]
                return text[:num_chars]

            result_df["摘要內容"] = result_df.apply(extract_text, axis=1)

            # 高亮顯示
            def highlight(val):
                if isinstance(val, str) and any(str(x) in val for x in [date_str, date_str.replace("114", "2025")]):
                    return "background-color: yellow"
                return ""

            st.dataframe(result_df.style.applymap(highlight, subset=[colname]))

            # Word 匯出
            doc_out = Document()
            for _, r in result_df.iterrows():
                doc_out.add_paragraph(f"{r[colname]} → {r['摘要內容']}")
            buf = BytesIO()
            doc_out.save(buf)
            st.download_button("⬇️ 下載 Word 檔", buf.getvalue(), "date_extract_summary.docx")

            # Excel 匯出
            excel_buf = BytesIO()
            result_df.to_excel(excel_buf, index=False)
            st.download_button("⬇️ 下載 Excel 檔", excel_buf.getvalue(), "date_extract_summary.xlsx")

            # TXT 匯出
            txt_buf = BytesIO()
            txt_buf.write("\n".join(result_df["摘要內容"]).encode("utf-8"))
            st.download_button("⬇️ 下載 TXT 檔", txt_buf.getvalue(), "date_extract_summary.txt")

            st.success("✅ 已完成擷取並可下載。")
