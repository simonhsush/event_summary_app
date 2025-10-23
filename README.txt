# Word 摘要擷取與反白工具

這是一個可在雲端執行的 Streamlit 應用。
可從上傳的 Word 文件中，搜尋指定日期的段落內容，並將日期反白輸出成 Word 及 CSV。

## 功能特色
- 上傳 Word 檔 (.docx)
- 自動或自選日期
- 擷取包含日期的段落
- 日期反白輸出 Word 檔
- 匯出 CSV 摘要檔

## 雲端部署方式（Streamlit Cloud）
1. 建立 GitHub repository，例如 `event_summary_app`
2. 上傳 `app.py`, `requirements.txt`, `README.md`
3. 到 [https://streamlit.io/cloud](https://streamlit.io/cloud)
4. 登入 → 選擇你的 GitHub 專案 → Deploy

部署完成後，你會得到一個網址：