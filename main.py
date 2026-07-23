import datetime
import xml.etree.ElementTree as ET
import requests
from jinja2 import Template

SEARCH_TERM = "Microbiome"
MAX_RESULTS = 10  # 抓取前 10 篇論文


def fetch_latest_pubmed_articles(keyword, max_results=10):
    """使用 NCBI E-utilities API 抓取最新論文內容"""
    print(f"[{datetime.datetime.now()}] 開始搜尋 PubMed: {keyword}...")

    # Step A: 搜尋符合條件的 PMID (PubMed ID)
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

    # Step B: 依 PMID 取得詳細摘要與資訊 (EFetch XML)
    fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    fetch_params = {
        "db": "pubmed",
        "id": ",".join(id_list),
        "retmode": "xml",
    }

    fetch_res = requests.get(fetch_url, params=fetch_params)
    fetch_res.raise_for_status()

    # Step C: 解析 XML
    root = ET.fromstring(fetch_res.content)
    articles = []

    for article in root.findall(".//PubmedArticle"):
        pmid = article.findtext(".//PMID")
        title = article.findtext(".//ArticleTitle") or "無標題"

        abstract_texts = article.findall(".//AbstractText")
        abstract = (
            " ".join([a.text for a in abstract_texts if a.text])
            if abstract_texts
            else "無提供摘要。"
        )

        pub_date = article.find(".//Journal/JournalIssue/PubDate")
        year = (
            pub_date.findtext("Year") if pub_date is not None else ""
        ) or "Unknown"

        articles.append(
            {
                "pmid": pmid,
                "title": title,
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
        .meta { font-size: 13px; color: #666; margin-bottom: 10px; }
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
            <div class="meta">PMID: {{ article.pmid }} | 年份: {{ article.date }}</div>
            <div class="abstract">{{ article.abstract }}</div>
        </div>
        {% endfor %}
    </div>
</body>
</html>
"""


def main():
    articles = fetch_latest_pubmed_articles(SEARCH_TERM, MAX_RESULTS)
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