import datetime
import json
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
from google import genai
from jinja2 import Template
import pandas as pd
import requests

# 搜尋關鍵字與設定
SEARCH_KEYWORDS = ["Microbiom", "metagenome", "metagenomic"]
SEARCH_TERM = " OR ".join(SEARCH_KEYWORDS)

MAX_RESULTS = 30
EXCEL_IF_PATH = "JCR-ImapctFactor-2025.xlsx"
DB_EXCEL_PATH = "papers_database.xlsx"

TAIPEI_TZ = ZoneInfo("Asia/Taipei")

MONTH_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
    "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
ai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

PRIMARY_MODEL = "gemini-3.5-flash"
FALLBACK_MODEL = "gemini-3.1-flash-lite"


def fetch_open_access_fulltext(pmid):
    """嘗試透過 BioC API 抓取 PMC Open Access 全文"""
    url = f"https://www.ncbi.nlm.nih.gov/research/bionlp/RESTful/pmcoa.cgi/BioC_json/{pmid}/unicode"
    try:
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            data = res.json()
            full_passages = []
            for doc in data.get("documents", []):
                for passage in doc.get("passages", []):
                    text = passage.get("text", "")
                    if text:
                        full_passages.append(text)
            full_text = "\n".join(full_passages)
            if len(full_text) > 500:
                print(f"  [Full-Text Success] PMID: {pmid} 成功獲取 Open Access 全文 ({len(full_text)} 字)")
                return full_text[:25000]  # 控制長度在 25k 字內
    except Exception:
        pass
    return None


def summarize_with_llm(title, abstract, affiliation="", fulltext=None, retries=3, delay=5):
    """使用 LLM 解析技術類型、樣本數、通訊作者國家與中文摘要 (優先使用 Full Text)"""
    if not GEMINI_API_KEY:
        return ["others"], "未提及", "未知國家", "⚠️ 未設定 GEMINI_API_KEY"

    has_fulltext = bool(fulltext)
    content_type_str = "內文全文" if has_fulltext else "標題與摘要"
    content_body = fulltext if has_fulltext else f"論文標題：{title}\n原文摘要：{abstract}"

    prompt = f"""
你是一位生物醫學與微生物學領域的專家。請根據提供的論文{content_type_str}與作者機構資訊，完成四項任務：

任務一：判斷該研究主要使用的技術類型。請從以下 6 個標籤中選擇：
1. 16S (包含 16S rRNA, amplicon sequencing 等擴增子定序)
2. metagenomics (總體基因體學 / 宏基因組學 / shotgun metagenomics)
3. metatranscriptomics (總體轉錄體學 / 宏轉錄組學)
4. metabolomics (代謝組學 / 代謝體學 / LC-MS, GC-MS 等代謝物分析)
5. small genome (小型基因體 / 菌株全基因體完成圖 / viral/bacterial genome assembly)
6. others (若不屬於上述五者，或無法明確判斷)
⚠️ 若論文中同時使用了兩種以上的技術，請將使用到的技術全數列出，並以半形逗號「,」分隔（例如：16S, metabolomics）。

任務二：擷取該研究的「研究樣本數量」（例如：n=50、120 位受試者、45 個糞便檢體、12 個小鼠模型、1,200 個基因體等）。
⚠️ 請特別關注文章中的 Materials and Methods 或 Results 區塊。若文章完全未提及樣本數，請填寫「未提及」。

任務三：請根據提供的作者機構資訊（Affiliation），判斷通訊作者（或主要研究團隊）來自的「國家/地區名稱」（特別關注是否包含 Taiwan、ROC、Taiwan R.O.C. 等，若為台灣請務必精準輸出 Taiwan；其餘請輸出英文國家名稱如 USA, China, Germany, Japan 等）。若完全無法判斷，請填寫「未知國家」。

任務四：撰寫一份「250字以內」的繁體中文重點解述（說明核心目的、主要發現與臨床/科學意義）。

請嚴格按照以下格式輸出：
[技術類型]: 技術標籤1, 技術標籤2
[樣本數量]: 樣本數量描述
[研究國家]: 國家名稱
[中文摘要]: 摘要內文...

作者機構資訊：{affiliation}
論文內容：
{content_body}
"""

    client = genai.Client(api_key=GEMINI_API_KEY)

    def _call_api(model_name):
        for attempt in range(1, retries + 1):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                )
                time.sleep(1)
                return response.text.strip()
            except Exception as e:
                err_str = str(e).lower()
                if any(k in err_str for k in ["429", "quota", "limit"]) and attempt < retries:
                    time.sleep(delay * attempt)
                else:
                    raise e

    def _parse_llm_output(output_text):
        tech_types = ["others"]
        sample_size = "未提及"
        country = "未知國家"
        zh_summary = output_text

        tech_match = re.search(r"\[技術類型\]:\s*(.*)", output_text, re.IGNORECASE)
        if tech_match:
            raw_tech_line = tech_match.group(1).split("\n")[0].strip()
            parsed_techs = [t.strip() for t in re.split(r"[,;，]", raw_tech_line) if t.strip()]
            if parsed_techs:
                tech_types = parsed_techs

        sample_match = re.search(r"\[樣本數量\]:\s*(.*)", output_text, re.IGNORECASE)
        if sample_match:
            sample_size = sample_match.group(1).split("\n")[0].strip()

        country_match = re.search(r"\[研究國家\]:\s*(.*)", output_text, re.IGNORECASE)
        if country_match:
            country = country_match.group(1).split("\n")[0].strip()

        summary_match = re.search(r"\[中文摘要\]:\s*(.*)", output_text, re.DOTALL)
        if summary_match:
            zh_summary = summary_match.group(1).strip()

        return tech_types, sample_size, country, zh_summary

    raw_output = None
    try:
        raw_output = _call_api(PRIMARY_MODEL)
    except Exception:
        try:
            raw_output = _call_api(FALLBACK_MODEL)
        except Exception:
            return ["others"], "未提及", "未知國家", "AI 分析失敗，請參考英文原文。"

    return _parse_llm_output(raw_output)


def get_full_text(element):
    return "".join(element.itertext()).strip() if element is not None else ""


def parse_pub_date_from_article(article_node):
    article_date = article_node.find(".//ArticleDate")
    if article_date is not None:
        y, m, d = article_date.findtext("Year"), article_date.findtext("Month"), article_date.findtext("Day")
        if y and m and d:
            return f"{y}-{int(m):02d}-{int(d):02d}"

    for status in ["pubmed", "entrez"]:
        pubmed_date = article_node.find(f".//PubMedPubDate[@PubStatus='{status}']")
        if pubmed_date is not None:
            y, m, d = pubmed_date.findtext("Year"), pubmed_date.findtext("Month"), pubmed_date.findtext("Day")
            if y and m and d:
                return f"{y}-{int(m):02d}-{int(d):02d}"

    pub_date_node = article_node.find(".//Journal/JournalIssue/PubDate")
    if pub_date_node is not None:
        y = pub_date_node.findtext("Year")
        m = pub_date_node.findtext("Month")
        d = pub_date_node.findtext("Day")
        if not y:
            medline = pub_date_node.findtext("MedlineDate")
            return medline[:4] if medline and len(medline) >= 4 else "未知日期"
        m = MONTH_MAP.get(m.strip().lower()[:3], f"{int(m):02d}") if m and m.isdigit() else "01"
        d = f"{int(d):02d}" if d and d.isdigit() else "01"
        return f"{y}-{m}-{d}"
    return "未知日期"


def load_impact_factors_from_excel(file_path):
    if not os.path.exists(file_path):
        return {}
    try:
        df = pd.read_excel(file_path)
        df.columns = [str(col).strip() for col in df.columns]
        if "Journal Name" in df.columns and "Impact Factor" in df.columns:
            return {str(r["Journal Name"]).strip().lower(): str(r["Impact Factor"]).strip() for _, r in df.iterrows()}
    except Exception:
        pass
    return {}


def fetch_latest_pubmed_articles(keyword, if_map, max_results=15):
    now_taipei = datetime.datetime.now(TAIPEI_TZ)
    today_str = now_taipei.strftime("%Y/%m/%d")

    search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    search_params = {
        "db": "pubmed", "term": keyword, "retmax": max_results,
        "sort": "pub_date", "retmode": "json", "datetype": "edat",
        "mindate": today_str, "maxdate": today_str,
    }

    res = requests.get(search_url, params=search_params)
    id_list = res.json()["esearchresult"]["idlist"]

    if not id_list:
        search_params.pop("mindate", None)
        search_params.pop("maxdate", None)
        search_params["reldate"] = 2
        res = requests.get(search_url, params=search_params)
        id_list = res.json()["esearchresult"]["idlist"]

    if not id_list:
        return []

    fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    fetch_res = requests.get(fetch_url, params={"db": "pubmed", "id": ",".join(id_list), "retmode": "xml"})

    root = ET.fromstring(fetch_res.content)
    articles = []

    for idx, article in enumerate(root.findall(".//PubmedArticle"), 1):
        pmid = article.findtext(".//PMID")
        title = get_full_text(article.find(".//ArticleTitle")) or "無標題"
        journal_title = article.findtext(".//Journal/Title") or "未知期刊"
        impact_factor = if_map.get(journal_title.strip().lower(), "N/A")

        # 擷取機構資訊供國家判斷
        affiliations = []
        for aff in article.findall(".//AuthorList/Author/AffiliationInfo/Affiliation"):
            aff_text = get_full_text(aff)
            if aff_text:
                affiliations.append(aff_text)
        affiliation_str = " | ".join(affiliations[:3])

        abstract_texts = article.findall(".//AbstractText")
        abstract = " ".join([get_full_text(a) for a in abstract_texts]) if abstract_texts else "無提供摘要。"
        pub_date_str = parse_pub_date_from_article(article)

        # 嘗試抓取全文
        fulltext = fetch_open_access_fulltext(pmid)
        has_fulltext_flag = " (全文)" if fulltext else " (摘要)"

        print(f"[{idx}/{len(id_list)}] PMID: {pmid}{has_fulltext_flag} 分析中...")
        tech_types, sample_size, country, zh_summary = summarize_with_llm(title, abstract, affiliation_str, fulltext)

        articles.append({
            "pmid": pmid,
            "title": title,
            "journal": journal_title,
            "impact_factor": impact_factor,
            "tech_types": tech_types,
            "sample_size": sample_size,
            "country": country,
            "has_fulltext": bool(fulltext),
            "abstract": abstract,
            "zh_summary": zh_summary,
            "date": pub_date_str,
            "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
        })

    return articles


def sync_database_to_excel(new_articles, db_path):
    print("\n--- [開始執行 Excel 儲存作業] ---")
    print(f"📍 目標檔案路徑: {os.path.abspath(db_path)}")
    print(f"📦 收到待存入論文數量: {len(new_articles)} 筆")

    if not new_articles:
        if os.path.exists(db_path):
            print("ℹ️ 無新論文，直接載入既有資料庫...")
            return pd.read_excel(db_path)
        
        print("❌ 警告：傳入的論文陣列為空 (0 筆)，程式將強制建立一份測試 Excel 以確保檔案生成！")
        dummy_df = pd.DataFrame([{
            "pmid": "00000000",
            "title": "無新論文（系統初始化）",
            "journal": "N/A",
            "impact_factor": "N/A",
            "tech_types": "none",
            "sample_size": "未提及",
            "country": "Taiwan",
            "has_fulltext": False,
            "abstract": "無資料",
            "zh_summary": "今日未擷取到新論文",
            "date": "2026-01-01",
            "url": "#"
        }])
        dummy_df.to_excel(db_path, index=False, engine="openpyxl")
        print(f"✅ 已強制建立基礎 Excel 檔案：{db_path}")
        return dummy_df

    new_df = pd.DataFrame(new_articles)

    # 1. 轉化 list 欄位
    if "tech_types" in new_df.columns:
        new_df["tech_types"] = new_df["tech_types"].apply(
            lambda x: ", ".join(x) if isinstance(x, list) else str(x)
        )

    # 2. 強制確保 PMID 型態為純字串，避免型態比對錯誤
    new_df["pmid"] = new_df["pmid"].astype(str).str.strip()

    # 3. 讀取或合併既有 Excel
    if os.path.exists(db_path):
        print("ℹ️ 偵測到既有 Excel 檔案，進行資料合併與去重...")
        try:
            existing_df = pd.read_excel(db_path)
            if "pmid" in existing_df.columns:
                existing_df["pmid"] = existing_df["pmid"].astype(str).str.strip()
            combined_df = pd.concat([new_df, existing_df], ignore_index=True)
            combined_df.drop_duplicates(subset=["pmid"], keep="first", inplace=True)
        except Exception as e:
            print(f"⚠️ 讀取舊 Excel 失敗 ({e})，將直接覆寫新檔案。")
            combined_df = new_df
    else:
        print("ℹ️ 未發現既有檔案，準備建立全新 Excel 檔案...")
        combined_df = new_df

    # 4. 排序與實體寫入
    combined_df.sort_values(by="date", ascending=False, inplace=True)

    try:
        combined_df.to_excel(db_path, index=False, engine="openpyxl")
        print(f"🎉【成功】`{db_path}` 已成功寫入實體硬碟！總筆數：{len(combined_df)}")
    except PermissionError:
        print(f"❌【失敗】檔案 `{db_path}` 正在被 Excel 或其他軟體開啟中，請先關閉該檔案後再重新執行程式！")
        sys.exit(1)
    except Exception as e:
        print(f"❌【寫入例外錯誤】: {e}")
        sys.exit(1)

    return combined_df


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PubMed 每日論文 AI 快訊與資料庫</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 1000px; margin: 0 auto; }
        header { background: #0056b3; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        header h1 { margin: 0; font-size: 24px; }
        .stats-bar { font-size: 14px; margin-top: 8px; opacity: 0.95; }
        .highlight-count { color: #ffeb3b; font-weight: bold; }
        
        /* 篩選與搜尋面板 */
        .filter-panel { background: white; padding: 15px 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.08); margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 15px; align-items: center; }
        .filter-group { display: flex; align-items: center; gap: 8px; font-size: 14px; font-weight: 500; }
        .filter-group input[type="text"], .filter-group input[type="date"], .filter-group select { padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        
        /* Taiwan 勾選框特別樣式 */
        .taiwan-checkbox-label { background: #e7f3ff; color: #0056b3; padding: 6px 12px; border-radius: 20px; border: 1px solid #b6d4fe; font-weight: bold; cursor: pointer; display: flex; align-items: center; gap: 6px; user-select: none; }
        .taiwan-checkbox-label input { width: 16px; height: 16px; cursor: pointer; }

        .btn-reset { background: #6c757d; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 14px; }
        .btn-reset:hover { background: #5a6268; }

        .card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
        .card h2 { margin-top: 0; font-size: 18px; line-height: 1.4; }
        .card h2 a { color: #0056b3; text-decoration: none; }
        .card h2 a:hover { text-decoration: underline; }
        .meta { font-size: 13px; color: #555; margin-bottom: 12px; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
        
        .badge-journal { background-color: #e9ecef; color: #495057; padding: 2px 8px; border-radius: 4px; font-weight: 500; }
        .badge-if { background-color: #d1e7dd; color: #0f5132; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
        .badge-tech { background-color: #fff3cd; color: #664d03; border: 1px solid #ffecb5; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
        .badge-sample { background-color: #e2e3e5; color: #383d41; border: 1px solid #d6d8db; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
        .badge-country { background-color: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; padding: 2px 8px; border-radius: 4px; font-weight: bold; }
        .badge-fulltext { background-color: #d1e7dd; color: #0f5132; border: 1px solid #badbcc; padding: 2px 8px; border-radius: 4px; font-weight: bold; }

        .ai-summary { background-color: #f0f7ff; border-left: 4px solid #0056b3; padding: 12px 15px; border-radius: 0 6px 6px 0; margin-bottom: 15px; }
        .ai-summary-title { font-weight: bold; color: #0056b3; font-size: 14px; margin-bottom: 5px; }
        .ai-summary-content { font-size: 14px; color: #2c3e50; line-height: 1.6; }

        /* 分頁元件 */
        .pagination { display: flex; justify-content: center; align-items: center; gap: 10px; margin: 30px 0; }
        .pagination button { background: #0056b3; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 14px; }
        .pagination button:disabled { background: #ccc; cursor: not-allowed; }
        .page-info { font-size: 14px; font-weight: 500; }
        
        details { font-size: 13px; color: #666; border-top: 1px solid #eee; padding-top: 8px; }
        summary { cursor: pointer; font-weight: 500; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>PubMed 每日最新論文 AI 快訊與檢索庫 by 克里斯</h1>
            <p class="stats-bar">
                資料庫目前收錄：<strong id="total-db-count">0</strong> 筆論文 ｜ 
                🇹🇼 台灣團隊文章：<span class="highlight-count" id="taiwan-db-count">0</span> 筆 ｜ 
                最後更新：{{ updated_at }}
            </p>
        </header>

        <!-- 篩選控制列 -->
        <div class="filter-panel">
            <label class="taiwan-checkbox-label">
                <input type="checkbox" id="filter-taiwan-only"> 🇹🇼 只顯示 Taiwan 論文
            </label>

            <div class="filter-group">
                <label>📅 開始日期:</label>
                <input type="date" id="filter-start-date">
            </div>
            <div class="filter-group">
                <label>📅 結束日期:</label>
                <input type="date" id="filter-end-date">
            </div>
            <div class="filter-group">
                <label>🔬 技術類型:</label>
                <select id="filter-tech">
                    <option value="">全部技術</option>
                    <option value="16S">16S</option>
                    <option value="metagenomics">metagenomics</option>
                    <option value="metatranscriptomics">metatranscriptomics</option>
                    <option value="metabolomics">metabolomics</option>
                    <option value="small genome">small genome</option>
                    <option value="others">others</option>
                </select>
            </div>
            <div class="filter-group">
                <label>🔍 關鍵字:</label>
                <input type="text" id="filter-search" placeholder="搜尋標題/摘要/期刊...">
            </div>
            <button class="btn-reset" onclick="resetFilters()">重置篩選</button>
        </div>

        <!-- 文章容器 -->
        <div id="articles-list"></div>

        <!-- 分頁控制欄 -->
        <div class="pagination">
            <button id="btn-prev" onclick="changePage(-1)">上一頁</button>
            <span class="page-info" id="page-info">第 1 頁 / 共 1 頁</span>
            <button id="btn-next" onclick="changePage(1)">下一頁</button>
        </div>
    </div>

    <script>
        // 載入完整歷史資料庫 (JSON 格式注入)
        const rawArticlesData = {{ articles_json | safe }};
        
        let filteredArticles = [...rawArticlesData];
        let currentPage = 1;
        const itemsPerPage = 15;

        // 初始化
        document.addEventListener("DOMContentLoaded", () => {
            document.getElementById("total-db-count").textContent = rawArticlesData.length;
            
            // 計算目前收錄多少屬 Taiwan 的文章
            const taiwanCount = rawArticlesData.filter(art => {
                const countryStr = String(art.country || "").toLowerCase();
                return countryStr.includes("taiwan");
            }).length;
            document.getElementById("taiwan-db-count").textContent = taiwanCount;

            // 事件監聽
            document.getElementById("filter-taiwan-only").addEventListener("change", applyFilters);
            document.getElementById("filter-start-date").addEventListener("change", applyFilters);
            document.getElementById("filter-end-date").addEventListener("change", applyFilters);
            document.getElementById("filter-tech").addEventListener("change", applyFilters);
            document.getElementById("filter-search").addEventListener("input", applyFilters);

            applyFilters();
        });

        function applyFilters() {
            const taiwanOnly = document.getElementById("filter-taiwan-only").checked;
            const startDate = document.getElementById("filter-start-date").value;
            const endDate = document.getElementById("filter-end-date").value;
            const selectedTech = document.getElementById("filter-tech").value.toLowerCase();
            const searchText = document.getElementById("filter-search").value.toLowerCase().trim();

            filteredArticles = rawArticlesData.filter(art => {
                // Taiwan 專屬勾選篩選
                if (taiwanOnly) {
                    const countryStr = String(art.country || "").toLowerCase();
                    if (!countryStr.includes("taiwan")) return false;
                }

                // 日期篩選
                if (startDate && art.date < startDate) return false;
                if (endDate && art.date > endDate) return false;

                // 技術類型篩選
                if (selectedTech) {
                    const techStr = String(art.tech_types).toLowerCase();
                    if (!techStr.includes(selectedTech)) return false;
                }

                // 關鍵字搜尋
                if (searchText) {
                    const fullContent = (art.title + art.journal + art.zh_summary + art.abstract + (art.country || "")).toLowerCase();
                    if (!fullContent.includes(searchText)) return false;
                }

                return true;
            });

            // 預設由最新日期開始排序 (最新在最前)
            filteredArticles.sort((a, b) => (b.date > a.date ? 1 : -1));

            currentPage = 1;
            renderArticles();
        }

        function renderArticles() {
            const container = document.getElementById("articles-list");
            container.innerHTML = "";

            if (filteredArticles.length === 0) {
                container.innerHTML = `<div class="card"><p style="text-align:center; color:#888;">查無符合條件的論文。</p></div>`;
                updatePaginationControls(0);
                return;
            }

            const totalPages = Math.ceil(filteredArticles.length / itemsPerPage);
            const startIndex = (currentPage - 1) * itemsPerPage;
            const pageData = filteredArticles.slice(startIndex, startIndex + itemsPerPage);

            pageData.forEach(art => {
                // 處理多技術標籤顯示
                const techs = String(art.tech_types).split(",").map(t => t.trim());
                const techBadges = techs.map(t => `<span class="badge-tech">🔬 ${t}</span>`).join(" ");
                const fullTextBadge = art.has_fulltext ? `<span class="badge-fulltext">📄 全文分析</span>` : ``;
                const countryBadge = `<span class="badge-country">🌐 ${art.country || '未知國家'}</span>`;

                const cardHtml = `
                    <div class="card">
                        <h2><a href="${art.url}" target="_blank">${art.title}</a></h2>
                        <div class="meta">
                            ${countryBadge}
                            ${techBadges}
                            <span class="badge-sample">📊 樣本: ${art.sample_size}</span>
                            ${fullTextBadge}
                            <span class="badge-journal">📖 ${art.journal}</span>
                            <span class="badge-if">IF: ${art.impact_factor}</span>
                            <span>📅 日期: ${art.date}</span>
                            <span>PMID: ${art.pmid}</span>
                        </div>
                        <div class="ai-summary">
                            <div class="ai-summary-title">🤖 AI 核心解述</div>
                            <div class="ai-summary-content">${art.zh_summary}</div>
                        </div>
                        <details>
                            <summary>查看英文原文摘要 (Abstract)</summary>
                            <div class="abstract-en" style="margin-top:8px;">${art.abstract}</div>
                        </details>
                    </div>
                `;
                container.insertAdjacentHTML("beforeend", cardHtml);
            });

            updatePaginationControls(totalPages);
        }

        function updatePaginationControls(totalPages) {
            document.getElementById("page-info").textContent = `第 ${currentPage} 頁 / 共 ${totalPages || 1} 頁 (篩選出 ${filteredArticles.length} 筆)`;
            document.getElementById("btn-prev").disabled = (currentPage <= 1);
            document.getElementById("btn-next").disabled = (currentPage >= totalPages || totalPages === 0);
        }

        function changePage(delta) {
            currentPage += delta;
            renderArticles();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function resetFilters() {
            document.getElementById("filter-taiwan-only").checked = false;
            document.getElementById("filter-start-date").value = "";
            document.getElementById("filter-end-date").value = "";
            document.getElementById("filter-tech").value = "";
            document.getElementById("filter-search").value = "";
            applyFilters();
        }
    </script>
</body>
</html>
"""


def main():
    if_map = load_impact_factors_from_excel(EXCEL_IF_PATH)
    
    # 1. 抓取每日最新論文
    new_articles = fetch_latest_pubmed_articles(SEARCH_TERM, if_map, max_results=MAX_RESULTS)
    
    # 2. 儲存並同步至歷史 Excel 資料庫
    db_df = sync_database_to_excel(new_articles, DB_EXCEL_PATH)
    
    # 3. 將資料庫全數導出為 JSON 並嵌入 HTML 前端
    db_articles = db_df.to_dict(orient="records")
    
    # 確保 JSON 相容性 (處理布林值與欄位)
    for art in db_articles:
        if "has_fulltext" not in art or pd.isna(art["has_fulltext"]):
            art["has_fulltext"] = False
        else:
            art["has_fulltext"] = bool(art["has_fulltext"])

        if "country" not in art or pd.isna(art["country"]):
            art["country"] = "未知國家"

    template = Template(HTML_TEMPLATE)
    updated_at = datetime.datetime.now(TAIPEI_TZ).strftime("%Y-%m-%d %H:%M:%S")
    
    articles_json = json.dumps(db_articles, ensure_ascii=False)
    
    html_content = template.render(
        articles_json=articles_json,
        updated_at=updated_at
    )

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    print("🎉 index.html 與 Excel 資料庫更新完成！")


if __name__ == "__main__":
    main()
