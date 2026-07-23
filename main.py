import datetime
import os
import xml.etree.ElementTree as ET
from google import genai
from jinja2 import Template
import pandas as pd
import requests

# 搜尋關鍵字與設定
SEARCH_KEYWORDS = ["Microbiome", "metagenome", "metagenomic"]
SEARCH_TERM = " OR ".join(SEARCH_KEYWORDS)

MAX_RESULTS = 10
EXCEL_FILE_PATH = "JCR-ImapctFactor-2025.xlsx"

MONTH_MAP = {
    "jan": "01",
    "feb": "02",
    "mar": "03",
    "apr": "04",
    "may": "05",
    "jun": "06",
    "jul": "07",
    "aug": "08",
    "sep": "09",
    "oct": "10",
    "nov": "11",
    "dec": "12",
}

# 初始化 Gemini API 用戶端
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
ai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None


def summarize_with_llm(title, abstract, retries=3, delay=3):
    """使用 LLM 將論文標題與摘要總結為 250 字以內的繁體中文解述 (含 429 自動重試機制)"""
    if not GEMINI_API_KEY:
        print("❌ 錯誤: 未找到 GEMINI_API_KEY 環境變數！")
        return "⚠️ 未設定 GEMINI_API_KEY，無法產生中文摘要。"

    if not abstract or abstract == "無提供摘要。":
        return "這篇論文未提供原文摘要，無法進行摘要轉譯。"

    prompt = f"""
你是一位生物醫學與微生物學領域的專家。請根據以下論文標題與摘要，撰寫一份「250字以內」的繁體中文重點解述。

要求：
1. 語言為繁體中文，文筆通順、專業且精煉。
2. 說明該研究的核心目的、主要發現或臨床/科學意義。
3. 嚴格控制字數在 250 字以內，不要列點，直接以精煉的一到兩段文字呈現。

論文標題：{title}
原文摘要：{abstract}
"""

    client = genai.Client(api_key=GEMINI_API_KEY)

    # 🔑 自動重試機制 (Handling 429 Rate Limit)
    for attempt in range(1, retries + 1):
        try:
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt,
            )
            # 呼叫成功後強制暫停 2 秒，避免下一次迴圈瞬間衝高 RPM
            time.sleep(2)
            return response.text.strip()

        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                wait_time = delay * (2 ** (attempt - 1))  # 指數型退避: 3s, 6s, 12s...
                print(
                    f"⚠️ 觸發 API 頻率限制 (429)，等待 {wait_time} 秒後進行第 {attempt}/{retries} 次重試..."
                )
                time.sleep(wait_time)
            else:
                print(
                    f"❌ Gemini API 呼叫失敗，原因: {type(e).__name__} - {e}"
                )
                return "中文摘要生成失敗，請參考英文原文。"

    print("❌ 已達到最大重試次數，無法取得中文摘要。")
    return "中文摘要生成失敗 (請求頻率過高)，請參考英文原文。"
        

def get_full_text(element):
    """遞迴擷取 XML 節點內部的所有純文字 (包含 <i>, <b> 等子標籤內的文字)"""
    if element is None:
        return ""
    return "".join(element.itertext()).strip()


def parse_pub_date(pub_date_node):
    """從 PubMed XML 的 PubDate 節點中提取並格式化年月日 (YYYY-MM-DD)"""
    if pub_date_node is None:
        return "未知日期"

    year = pub_date_node.findtext("Year")
    month = pub_date_node.findtext("Month")
    day = pub_date_node.findtext("Day")

    if not year:
        medline_date = pub_date_node.findtext("MedlineDate")
        if medline_date:
            parts = medline_date.split()
            if parts and len(parts[0]) == 4 and parts[0].isdigit():
                return parts[0]
            return medline_date
        return "未知日期"

    if month:
        month_clean = month.strip().lower()[:3]
        if month_clean in MONTH_MAP:
            month = MONTH_MAP[month_clean]
        elif month.isdigit():
            month = f"{int(month):02d}"
    else:
        month = ""

    if day and day.isdigit():
        day = f"{int(day):02d}"
    else:
        day = ""

    if year and month and day:
        return f"{year}-{month}-{day}"
    elif year and month:
        return f"{year}-{month}"
    else:
        return year


def load_impact_factors_from_excel(file_path):
    """讀取 Excel 檔案並轉為 Python 字典供快速查詢"""
    if not os.path.exists(file_path):
        print(f"⚠️ 警告: 找不到 Excel 檔案 ({file_path})，Impact Factor 將全顯示 N/A")
        return {}

    try:
        df = pd.read_excel(file_path)
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
    print(f"[{datetime.datetime.now()}] 開始搜尋 PubMed: '{keyword}'...")

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

    for idx, article in enumerate(root.findall(".//PubmedArticle"), 1):
        pmid = article.findtext(".//PMID")

        title_element = article.find(".//ArticleTitle")
        title = get_full_text(title_element) or "無標題"

        journal_title = (
            article.findtext(".//Journal/Title")
            or article.findtext(".//Journal/ISOAbbreviation")
            or "未知期刊"
        )

        impact_factor = get_impact_factor(journal_title, if_map)

        abstract_texts = article.findall(".//AbstractText")
        if abstract_texts:
            abstract_parts = [get_full_text(a) for a in abstract_texts]
            abstract = " ".join([p for p in abstract_parts if p])
        else:
            abstract = "無提供摘要。"

        pub_date_node = article.find(".//Journal/JournalIssue/PubDate")
        pub_date_str = parse_pub_date(pub_date_node)

        # 🔑 呼叫 LLM 產生中文摘要
        print(f"[{idx}/{max_results}] 正在為 PMID: {pmid} 產生中文摘要...")
        zh_summary = summarize_with_llm(title, abstract)

        articles.append(
            {
                "pmid": pmid,
                "title": title,
                "journal": journal_title,
                "impact_factor": impact_factor,
                "abstract": abstract,
                "zh_summary": zh_summary,  # 新增 LLM 中文摘要欄位
                "date": pub_date_str,
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
    <title>PubMed 每日論文 AI 摘要快訊</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; }
        header { background: #0056b3; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        header h1 { margin: 0; font-size: 24px; }
        header p { margin: 5px 0 0 0; opacity: 0.8; font-size: 14px; }
        .keyword-tag { background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 4px; font-size: 13px; font-weight: bold; }
        .card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
        .card h2 { margin-top: 0; font-size: 18px; }
        .card h2 a { color: #0056b3; text-decoration: none; }
        .card h2 a:hover { text-decoration: underline; }
        .meta { font-size: 13px; color: #555; margin-bottom: 12px; display: flex; gap: 12px; flex-wrap: wrap; align-items: center; }
        .badge-journal { background-color: #e9ecef; color: #495057; padding: 2px 8px; border-radius: 4px; font-weight: 500; }
        .badge-if { background-color: #d1e7dd; color: #0f5132; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
        
        /* LLM 中文摘要區塊樣式 */
        .ai-summary { background-color: #f0f7ff; border-left: 4px solid #0056b3; padding: 12px 15px; border-radius: 0 6px 6px 0; margin-bottom: 15px; }
        .ai-summary-title { font-weight: bold; color: #0056b3; font-size: 14px; margin-bottom: 5px; display: flex; align-items: center; gap: 5px; }
        .ai-summary-content { font-size: 14px; color: #2c3e50; line-height: 1.6; }
        
        /* 英文原文摺疊選單 */
        details { font-size: 13px; color: #666; border-top: 1px solid #eee; padding-top: 8px; }
        summary { cursor: pointer; font-weight: 500; color: #666; }
        summary:hover { color: #0056b3; }
        .abstract-en { margin-top: 8px; line-height: 1.5; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>PubMed 每日最新論文 AI 快訊</h1>
            <p>搜尋主題：
                {% for kw in keywords %}
                <span class="keyword-tag">{{ kw }}</span>
                {% endfor %}
            </p>
            <p>最後更新時間：{{ updated_at }} (UTC)</p>
        </header>

        {% for article in articles %}
        <div class="card">
            <h2><a href="{{ article.url }}" target="_blank">{{ article.title }}</a></h2>
            <div class="meta">
                <span class="badge-journal">📖 {{ article.journal }}</span>
                <span class="badge-if">IF: {{ article.impact_factor }}</span>
                <span>📅 發表日期: {{ article.date }}</span>
                <span>PMID: {{ article.pmid }}</span>
            </div>
            
            <!-- AI 中文解述區塊 -->
            <div class="ai-summary">
                <div class="ai-summary-title">🤖 AI 核心解述 (250字內)</div>
                <div class="ai-summary-content">{{ article.zh_summary }}</div>
            </div>

            <!-- 英文原文摘要摺疊選單 -->
            <details>
                <summary>查看英文原文摘要 (Abstract)</summary>
                <div class="abstract-en">{{ article.abstract }}</div>
            </details>
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
            articles=articles, keywords=SEARCH_KEYWORDS, updated_at=updated_at
        )

        with open("index.html", "w", encoding="utf-8") as f:
            f.write(html_content)
        print("index.html 生成完成！")


if __name__ == "__main__":
    main()