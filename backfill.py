import datetime
import json
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
from google import genai
import pandas as pd
import requests

# ---------------------------------------------------------------------------
# 全域預設設定區 (若無環境變數傳入時使用)
# ---------------------------------------------------------------------------
DEFAULT_START_DATE = "2026/01/01"  # 總起始日期 (YYYY/MM/DD)
DEFAULT_END_DATE   = "2026/03/01"  # 總結束日期 (YYYY/MM/DD)
INTERVAL_DAYS     = 5             # 每次搜尋的天數跨度 (5 天)
MAX_FETCH         = 50            # 每次搜尋的最大論文筆數 (50 篇)
STATE_FILE        = "backfill_state.json"  # 儲存進度的紀錄檔
# ---------------------------------------------------------------------------

SEARCH_KEYWORDS = ["Microbio", "metagenome", "metagenomic"]
SEARCH_TERM = " OR ".join(SEARCH_KEYWORDS)

EXCEL_IF_PATH = "JCR-ImapctFactor-2025.xlsx"
DB_EXCEL_PATH = "papers_database.xlsx"

TAIPEI_TZ = ZoneInfo("Asia/Taipei")

MONTH_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
    "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
PRIMARY_MODEL = "gemini-3.5-flash"
FALLBACK_MODEL = "gemini-3.1-flash-lite"


def fetch_open_access_fulltext(pmid):
    url = f"https://www.ncbi.nlm.nih.gov/research/bionlp/RESTful/pmcoa.cgi/BioC_json/{pmid}/unicode"
    try:
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            data = res.json()
            full_passages = []
            for doc in data.get("documents", []):
                for passage in doc.get("passages", []):
                    text = passage.get("text", "")
                    if text: full_passages.append(text)
            full_text = "\n".join(full_passages)
            if len(full_text) > 500:
                print(f"  [Full-Text] PMID: {pmid} 獲取全文成功")
                return full_text[:13000]
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

⚠️【繁體中文與台灣生醫用語規範】（請務必嚴格遵循）：
必須完全使用「台灣繁體中文」的慣用語彙與用語習慣，嚴格禁止使用中國大陸的用語與譯名。請參考以下術語對照表進行翻譯：
- metagenomics/metagenome：請使用「總體基因體/總體基因體學」（嚴禁使用：宏基因組）
- metatranscriptomics：請使用「總體轉錄體/總體轉錄體學」（嚴禁使用：宏轉錄組）
- metabolomics：請使用「代謝體學」（嚴禁使用：代謝組學）
- genomics：請使用「基因體學」（嚴禁使用：基因組學）
- transcriptomics：請使用「轉錄體學」（嚴禁使用：轉錄組學）
- proteomics：請使用「蛋白質體學」（嚴禁使用：蛋白質組學）
- microbiome：請使用「微生物體/微生物群」（嚴禁使用：微生態）
- data：請使用「資料/數據」（優先使用：資料）
- pathway：請使用「路徑/傳導路徑」（嚴禁使用：通路）
- cohort：請使用「佇列/研究群體」（嚴禁使用：隊列）

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
                time.sleep(3)
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
        if y and m and d: return f"{y}-{int(m):02d}-{int(d):02d}"

    for status in ["pubmed", "entrez"]:
        pubmed_date = article_node.find(f".//PubMedPubDate[@PubStatus='{status}']")
        if pubmed_date is not None:
            y, m, d = pubmed_date.findtext("Year"), pubmed_date.findtext("Month"), pubmed_date.findtext("Day")
            if y and m and d: return f"{y}-{int(m):02d}-{int(d):02d}"

    pub_date_node = article_node.find(".//Journal/JournalIssue/PubDate")
    if pub_date_node is not None:
        y, m, d = pub_date_node.findtext("Year"), pub_date_node.findtext("Month"), pub_date_node.findtext("Day")
        if not y:
            medline = pub_date_node.findtext("MedlineDate")
            return medline[:4] if medline and len(medline) >= 4 else "未知日期"
        m = MONTH_MAP.get(m.strip().lower()[:3], f"{int(m):02d}") if m and m.isdigit() else "01"
        d = f"{int(d):02d}" if d and d.isdigit() else "01"
        return f"{y}-{m}-{d}"
    return "未知日期"


def load_impact_factors_from_excel(file_path):
    if not os.path.exists(file_path): return {}
    try:
        df = pd.read_excel(file_path)
        df.columns = [str(col).strip() for col in df.columns]
        if "Journal Name" in df.columns and "Impact Factor" in df.columns:
            return {str(r["Journal Name"]).strip().lower(): str(r["Impact Factor"]).strip() for _, r in df.iterrows()}
    except Exception: pass
    return {}


def fetch_range_pubmed_articles(keyword, start_date, end_date, if_map, max_results=50):
    print(f"🔍 開始搜尋 PubMed 區間：{start_date} ~ {end_date}，預計上限：{max_results} 筆...")

    search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    search_params = {
        "db": "pubmed", "term": keyword, "retmax": max_results,
        "sort": "pub_date", "retmode": "json", "datetype": "edat",
        "mindate": start_date, "maxdate": end_date,
    }

    res = requests.get(search_url, params=search_params)
    id_list = res.json()["esearchresult"]["idlist"]

    if not id_list:
        print("⚠️ 該區間內未搜尋到任何符合關鍵字的論文。")
        return []

    print(f"📦 搜尋到 {len(id_list)} 筆論文，開始進行詳細分析...")
    fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    fetch_res = requests.get(fetch_url, params={"db": "pubmed", "id": ",".join(id_list), "retmode": "xml"})

    root = ET.fromstring(fetch_res.content)
    articles = []

    for idx, article in enumerate(root.findall(".//PubmedArticle"), 1):
        pmid = article.findtext(".//PMID")
        title = get_full_text(article.find(".//ArticleTitle")) or "無標題"
        journal_title = article.findtext(".//Journal/Title") or "未知期刊"
        impact_factor = if_map.get(journal_title.strip().lower(), "N/A")

        affiliations = [get_full_text(aff) for aff in article.findall(".//AuthorList/Author/AffiliationInfo/Affiliation") if get_full_text(aff)]
        affiliation_str = " | ".join(affiliations[:3])

        abstract_texts = article.findall(".//AbstractText")
        abstract = " ".join([get_full_text(a) for a in abstract_texts]) if abstract_texts else "無提供摘要。"
        pub_date_str = parse_pub_date_from_article(article)

        fulltext = fetch_open_access_fulltext(pmid)

        print(f"[{idx}/{len(id_list)}] PMID: {pmid} 分析中...")
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
    if not new_articles:
        return pd.read_excel(db_path) if os.path.exists(db_path) else pd.DataFrame()

    new_df = pd.DataFrame(new_articles)
    if "tech_types" in new_df.columns:
        new_df["tech_types"] = new_df["tech_types"].apply(lambda x: ", ".join(x) if isinstance(x, list) else str(x))

    new_df["pmid"] = new_df["pmid"].astype(str).str.strip()

    if os.path.exists(db_path):
        existing_df = pd.read_excel(db_path)
        if "pmid" in existing_df.columns:
            existing_df["pmid"] = existing_df["pmid"].astype(str).str.strip()
        combined_df = pd.concat([new_df, existing_df], ignore_index=True)
        combined_df.drop_duplicates(subset=["pmid"], keep="first", inplace=True)
    else:
        combined_df = new_df

    combined_df.sort_values(by="date", ascending=False, inplace=True)
    combined_df.to_excel(db_path, index=False, engine="openpyxl")
    print(f"🎉【成功】歷史資料庫已更新！總計包含：{len(combined_df)} 筆論文。")
    return combined_df


def generate_date_batches(start_str, end_str, interval_days=5):
    """將總日期區間依照指定天數切分成多個小區間批次"""
    start_dt = datetime.datetime.strptime(start_str, "%Y/%m/%d")
    end_dt = datetime.datetime.strptime(end_str, "%Y/%m/%d")

    batches = []
    curr_dt = start_dt

    while curr_dt <= end_dt:
        batch_end_dt = curr_dt + datetime.timedelta(days=interval_days - 1)
        if batch_end_dt > end_dt:
            batch_end_dt = end_dt

        batch_start_str = curr_dt.strftime("%Y/%m/%d")
        batch_end_str = batch_end_dt.strftime("%Y/%m/%d")
        batches.append((batch_start_str, batch_end_str))

        curr_dt = batch_end_dt + datetime.timedelta(days=1)

    return batches


def load_or_init_state(start_date, end_date):
    """讀取進度檔案，若不存在則重新建立進度表"""
    batches = generate_date_batches(start_date, end_date, INTERVAL_DAYS)
    
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
                # 若全域範圍沒變，沿用目前的索引 index
                if state.get("start_date") == start_date and state.get("end_date") == end_date:
                    return state, batches
        except Exception as e:
            print(f"⚠️ 讀取狀態檔失敗 ({e})，重新建立狀態。")

    state = {
        "start_date": start_date,
        "end_date": end_date,
        "next_batch_index": 0,
        "total_batches": len(batches)
    }
    return state, batches


def save_state(state):
    """將進度儲存回 JSON 檔案"""
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    # 優先從環境變數讀取，若無則使用預設值
    global_start = os.environ.get("START_DATE") or DEFAULT_START_DATE
    global_end = os.environ.get("END_DATE") or DEFAULT_END_DATE

    state, batches = load_or_init_state(global_start, global_end)
    current_idx = state["next_batch_index"]

    if current_idx >= len(batches):
        print(f"🎉 所有人指定的日期區間（{global_start} ~ {global_end}）已全部執行完畢！無需再執行。")
        sys.exit(0)

    # 取得本次應執行的 5 天區間
    b_start, b_end = batches[current_idx]
    print(f"🚀 [進度 {current_idx + 1}/{len(batches)}] 今日自動執行區間：{b_start} ~ {b_end}")

    if_map = load_impact_factors_from_excel(EXCEL_IF_PATH)
    articles = fetch_range_pubmed_articles(
        keyword=SEARCH_TERM,
        start_date=b_start,
        end_date=b_end,
        if_map=if_map,
        max_results=MAX_FETCH
    )

    # 儲存 Excel 資料庫
    sync_database_to_excel(articles, DB_EXCEL_PATH)

    # 更新進度至下一個區間並寫回 backfill_state.json
    state["next_batch_index"] = current_idx + 1
    save_state(state)
    print(f"✅ 今日任務完成，進度已推進至第 {state['next_batch_index']}/{len(batches)} 個區間！")
