import datetime
import os
import xml.etree.ElementTree as ET
from jinja2 import Template
import pandas as pd
import requests

SEARCH_TERM = "Microbiome"
MAX_RESULTS = 10
# 🔑 已更新為最新的 2025 JCR 表格檔名
EXCEL_FILE_PATH = "JCR-ImapctFactor-2025.xlsx"


def get_full_text(element):
    """遞迴擷取 XML 節點內部的所有純文字 (包含 <i>, <b> 等子標籤內的文字)"""
    if element is None:
        return ""
    return "".join(element.itertext()).strip()


def load_impact_factors_from_excel(file_path):
    """讀取 Excel 檔案並轉為 Python 字典供快速查詢"""
    if not os.path.exists(file_path):
        print(f"⚠️ 警告: 找不到 Excel 檔案 ({file_path})，Impact Factor 將全顯示 N/A")
        return {}

    try:
        df = pd.read_excel(file_path)
        # 整理標頭欄位名稱去空白
        df.columns = [str(col).strip() for col in df.columns]

        if (
            "Journal Name" not in df.columns
            or "Impact Factor" not in df.columns
        ):
            print(
                "⚠️ 警告: Excel 欄位名稱需包含 'Journal Name' 與 'Impact Factor'"
            )
            return {}

        if_map = {}
        for _, row in df.iterrows():
            journal = str(row["Journal Name"]).strip().lower()
            if_value = str(row["Impact Factor"]).strip()
            if journal:
                if_map[journal] = if_value

        print(
            f"✅ 成功從 {file_path} 載入 {len(if_map)} 筆期刊 Impact Factor 資料！"
        )
        return if_map
    except Exception as e:
        print(f"❌ 讀取 Excel 檔案失敗: {e}")
        return {}


def get_impact_factor(journal_title, if_map):
    if not journal_title or not if_map:
        return "N/A"
    clean_title = journal_title.strip().lower()
    return if_map.get(clean_title, "N/A")


def fetch_latest_pubmed_articles(keyword, if_map, max_results=10):
    """使用 NCBI E-utilities API 抓取最新論文內容"""
    print(f"[{datetime.datetime.now()}] 開始搜尋 PubMed: {keyword}...")

    search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    search_params = {
        "db": "pubmed",
        "term": keyword,
        "retmax": max_results,
        "sort": "pub_date",
        "retmode": "json",
    }

    res = requests.get(search_url, params=search_params)
    res.raise_for_status()
    id_list = res.json()["esearchresult"]["idlist"]

    if not id_list:
        print("未找到相關論文。")
        return []

    fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    fetch_params = {
        "db": "pubmed",
        "id": ",".join(id_list),
        "retmode": "xml",
    }

    fetch_res = requests.get(fetch_url, params=fetch_params)
    fetch_res.raise_for_status()

    root = ET.fromstring(fetch_res.content)
    articles = []

    for article in root.findall(".//PubmedArticle"):
        pmid = article.findtext(".//PMID")

        # 完整擷取含子標籤的標題
        title_element = article.find(".//ArticleTitle")
        title = get_full_text(title_element) or "無標題"

        # 抓取期刊名稱
        journal_title = (
            article.findtext(".//Journal/Title")
            or article.findtext(".//Journal/ISOAbbreviation")
            or "未知期刊"
        )

        impact_factor = get_impact_factor(journal_title, if_map)

        # 完整擷取摘要
        abstract_texts = article.findall(".//AbstractText")
        if abstract_texts:
            abstract_parts = [get_full_text(a) for a in abstract_texts]
            abstract = " ".join([p for p in abstract_parts if p])
        else:
            abstract = "無提供摘要。"

        pub_date = article.find(".//Journal/JournalIssue/PubDate")
        year = (
            pub_date.findtext("Year") if pub_date is not None else ""
        ) or "Unknown"

        articles.append(
            {
                "pmid": pmid,
                "title": title,
                "journal": journal_title,
                "impact_factor": impact_factor,
                "abstract": abstract,
                "date": year,
                "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
            }
        )

    return articles


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PubMed 每日論文報告 - {{ keyword }}</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; }
        header { background: #0056b3; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        header h1 { margin: 0; font-size: 24px; }
        header p { margin: 5px 0 0 0; opacity: 0.8; font-size: 14px; }
        .card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
        .card h2 { margin-top: 0; font-size: 18px; }
        .card h2 a { color: #0056b3; text-decoration: none; }
        .card h2 a:hover { text-decoration: underline; }
        .meta { font-size: 13px; color: #555; margin-bottom: 12px; display: flex; gap: 12px; flex-wrap: wrap; align-items: center; }
        .badge-journal { background-color: #e9ecef; color: #495057; padding: 2px 8px; border-radius: 4px; font-weight: 500; }
        .badge-if { background-color: #d1e7dd; color: #0f5132; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
        .abstract { font-size: 14px; color: #444; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>PubMed 每日最新論文：{{ keyword }}</h1>
            <p>最後更新時間：{{ updated_at }} (UTC)</p>
        </header>

        {% for article in articles %}
        <div class="card">
            <h2><a href="{{ article.url }}" target="_blank">{{ article.title }}</a></h2>
            <div class="meta">
                <span class="badge-journal">📖 {{ article.journal }}</span>
                <span class="badge-if">IF: {{ article.impact_factor }}</span>
                <span>PMID: {{ article.pmid }}</span>
                <span>年份: {{ article.date }}</span>
            </div>
            <div class="abstract">{{ article.abstract }}</div>
        </div>
        {% endfor %}
    </div>
</body>
</html>
"""


def main():
    if_map = load_impact_factors_from_excel(EXCEL_FILE_PATH)
    articles = fetch_latest_pubmed_articles(
        SEARCH_TERM, if_map, max_results=MAX_RESULTS
    )

    if articles:
        template = Template(HTML_TEMPLATE)
        updated_at = datetime.datetime.now(datetime.timezone.utc).strftime(
            "%Y-%m-%d %H:%M:%S"
        )
        html_content = template.render(
            articles=articles, keyword=SEARCH_TERM, updated_at=updated_at
        )

        with open("index.html", "w", encoding="utf-8") as f:
            f.write(html_content)
        print("index.html 生成完成！")


if __name__ == "__main__":
    main()