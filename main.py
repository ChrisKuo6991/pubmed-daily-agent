import datetime
import os
import re
import time
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
from google import genai
from jinja2 import Template
import pandas as pd
import requests

# 搜尋關鍵字與設定
SEARCH_KEYWORDS = ["Microbiome", "metagenome", "metagenomic"]
SEARCH_TERM = " OR ".join(SEARCH_KEYWORDS)

MAX_RESULTS = 10
EXCEL_FILE_PATH = "JCR-ImapctFactor-2025.xlsx"

TAIPEI_TZ = ZoneInfo("Asia/Taipei")

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

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
ai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

PRIMARY_MODEL = "gemini-3.5-flash"
FALLBACK_MODEL = "gemini-3.1-flash-lite"


def summarize_with_llm(title, abstract, retries=3, delay=5):
    """使用 LLM 將論文標題與摘要總結為繁體中文解述，並自動判斷技術類型"""
    if not GEMINI_API_KEY:
        print("❌ 錯誤: 未找到 GEMINI_API_KEY 環境變數！")
        return "others", "⚠️ 未設定 GEMINI_API_KEY，無法產生中文摘要。"

    if not abstract or abstract == "無提供摘要。":
        return "others", "這篇論文未提供原文摘要，無法進行摘要轉譯。"

    # 🔑 在 Prompt 中加入技術類型判斷要求
    prompt = f"""
你是一位生物醫學與微生物學領域的專家。請根據以下論文標題與摘要，完成兩項任務：

任務一：判斷該研究主要使用的技術類型，必須且只能歸類為以下 5 個選項之一：
1. 16S (包含 16S rRNA, amplicon sequencing 等擴增子定序)
2. metagenomics (總體基因體學 / 宏基因組學 /  shotgun metagenomics)
3. metatranscriptomics (總體轉錄體學 / 宏轉錄組學)
4. small genome (小型基因體 / 菌株全基因體完成圖 / viral/bacterial genome assembly)
5. others (若不屬於上述四者，或無法明確判斷)

任務二：撰寫一份「250字以內」的繁體中文重點解述（說明核心目的、主要發現或意義）。

請嚴格按照以下格式輸出（不要增加額外開頭文字）：
[技術類型]: 選項名稱 (僅填寫 16S / metagenomics / metatranscriptomics / small genome / others 之一)
[中文摘要]: 摘要內文...

論文標題：{title}
原文摘要：{abstract}
"""

    client = genai.Client(api_key=GEMINI_API_KEY)

    def _call_api(model_name):
        for attempt in range(1, retries + 1):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                )
                time.sleep(2)
                return response.text.strip()
            except Exception as e:
                err_str = str(e).lower()
                is_rate_limit = any(
                    k in err_str
                    for k in ["429", "resource_exhausted", "quota", "limit"]
                )

                if is_rate_limit and attempt < retries:
                    wait_time = delay * attempt
                    print(
                        f"⚠️ [{model_name}] 觸發 Rate Limit/Quota (429)，等待 {wait_time} 秒後重試 ({attempt}/{retries})..."
                    )
                    time.sleep(wait_time)
                else:
                    raise e

    # 解析 LLM 傳回的字串，拆解出「技術類型」與「摘要內文」
    def _parse_llm_output(output_text):
        tech_type = "others"
        zh_summary = output_text

        # 匹配 [技術類型]: xxx
        tech_match = re.search(
            r"\[技術類型\]:\s*(16S|metagenomics|metatranscriptomics|small genome|others)",
            output_text,
            re.IGNORECASE,
        )
        if tech_match:
            tech_type = tech_match.group(1).strip()

        # 匹配 [中文摘要]: xxx
        summary_match = re.search(r"\[中文摘要\]:\s*(.*)", output_text, re.DOTALL)
        if summary_match:
            zh_summary = summary_match.group(1).strip()

        return tech_type, zh_summary

    raw_output = None
    try:
        raw_output = _call_api(PRIMARY_MODEL)
    except Exception as e:
        print(
            f"⚠️ 主要模型 [{PRIMARY_MODEL}] 呼叫失敗: {e}\n🔄 切換至備用模型 [{FALLBACK_MODEL}]..."
        )

    if not raw_output:
        try:
            raw_output = _call_api(FALLBACK_MODEL)
        except Exception as e:
            print(f"❌ 備用模型 [{FALLBACK_MODEL}] 亦呼叫失敗: {e}")
            return "others", "中文摘要生成失敗 (API 配額超限)，請參考英文原文。"

    return _parse_llm_output(raw_output)


def get_full_text(element):
    """遞迴擷取 XML 節點內部的所有純文字"""
    if element is None:
        return ""
    return "".join(element.itertext()).strip()


def parse_pub_date_from_article(article_node):
    """從 PubmedArticle XML 節點中多層級精準擷取發表日期 (YYYY-MM-DD)"""
    article_date = article_node.find(".//ArticleDate")
    if article_date is not None:
        year = article_date.findtext("Year")
        month = article_date.findtext("Month")
        day = article_date.findtext("Day")
        if year and month and day:
            return f"{year}-{int(month):02d}-{int(day):02d}"

    for status in ["pubmed", "entrez"]:
        pubmed_date = article_node.find(f".//PubMedPubDate[@PubStatus='{status}']")
        if pubmed_date is not None:
            year = pubmed_date.findtext("Year")
            month = pubmed_date.findtext("Month")
            day = pubmed_date.findtext("Day")
            if year and month and day:
                return f"{year}-{int(month):02d}-{int(day):02d}"

    pub_date_node = article_node.find(".//Journal/JournalIssue/PubDate")
    if pub_date_node is not None:
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

    return "未知日期"


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
    """使用 NCBI E-utilities API 抓取當日最新論文內容"""
    now_taipei = datetime.datetime.now(TAIPEI_TZ)
    today_str = now_taipei.strftime("%Y/%m/%d")

    print(f"[{now_taipei.strftime('%Y-%m-%d %H:%M:%S')}] 開始搜尋 PubMed (台灣當日 {today_str}): '{keyword}'...")

    search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"

    search_params = {
        "db": "pubmed",
        "term": keyword,
        "retmax": max_results,
        "sort": "pub_date",
        "retmode": "json",
        "datetype": "edat",
        "mindate": today_str,
        "maxdate": today_str,
    }

    res = requests.get(search_url, params=search_params)
    res.raise_for_status()
    id_list = res.json()["esearchresult"]["idlist"]

    if not id_list:
        print(f"⚠️ 當日 ({today_str}) 無新論文，自動切換至搜尋近 2 天內的論文...")
        search_params.pop("mindate", None)
        search_params.pop("maxdate", None)
        search_params["reldate"] = 2

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

        pub_date_str = parse_pub_date_from_article(article)

        print(f"[{idx}/{len(id_list)}] 正在為 PMID: {pmid} 分析技術類型與摘要...")
        tech_type, zh_summary = summarize_with_llm(title, abstract)

        articles.append(
            {
                "pmid": pmid,
                "title": title,
                "journal": journal_title,
                "impact_factor": impact_factor,
                "tech_type": tech_type,  # 🔑 新增技術類型欄位
                "abstract": abstract,
                "zh_summary": zh_summary,
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
    <title>PubMed 每日論文 AI 摘要快訊 by 克里斯</title>
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
        
        /* 🔑 新增技術類型標籤樣式 */
        .badge-tech { background-color: #fff3cd; color: #664d03; border: 1px solid #ffecb5; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
        
        .ai-summary { background-color: #f0f7ff; border-left: 4px solid #0056b3; padding: 12px 15px; border-radius: 0 6px 6px 0; margin-bottom: 15px; }
        .ai-summary-title { font-weight: bold; color: #0056b3; font-size: 14px; margin-bottom: 5px; display: flex; align-items: center; gap: 5px; }
        .ai-summary-content { font-size: 14px; color: #2c3e50; line-height: 1.6; }
        
        details { font-size: 13px; color: #666; border-top: 1px solid #eee; padding-top: 8px; }
        summary { cursor: pointer; font-weight: 500; color: #666; }
        summary:hover { color: #0056b3; }
        .abstract-en { margin-top: 8px; line-height: 1.5; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>PubMed 每日最新論文 AI 快訊 by 克里斯</h1>
            <p>搜尋主題：
                {% for kw in keywords %}
                <span class="keyword-tag">{{ kw }}</span>
                {% endfor %}
            </p>
            <p>最後更新時間：{{ updated_at }} (Asia/Taipei)</p>
        </header>

        {% for article in articles %}
        <div class="card">
            <h2><a href="{{ article.url }}" target="_blank">{{ article.title }}</a></h2>
            <div class="meta">
                <!-- 🔑 顯示技術類型標籤 -->
                <span class="badge-tech">🔬 {{ article.tech_type }}</span>
                <span class="badge-journal">📖 {{ article.journal }}</span>
                <span class="badge-if">IF: {{ article.impact_factor }}</span>
                <span>📅 發表日期: {{ article.date }}</span>
                <span>PMID: {{ article.pmid }}</span>
            </div>
            
            <div class="ai-summary">
                <div class="ai-summary-title">🤖 AI 核心解述 (250字內)</div>
                <div class="ai-summary-content">{{ article.zh_summary }}</div>
            </div>

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
        updated_at = datetime.datetime.now(TAIPEI_TZ).strftime(
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
