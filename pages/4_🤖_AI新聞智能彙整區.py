import streamlit as st
import pandas as pd
import gspread
import datetime
import io
import requests
import re
import time
import urllib3
from bs4 import BeautifulSoup
import google.generativeai as genai
from google.oauth2.credentials import Credentials
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import qn

# 🌟 關閉 SSL 憑證驗證警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# 🌟 全域設定 & 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！"); st.stop()

st.set_page_config(page_title="嘉大 AI 新聞智能彙整", page_icon="📰", layout="wide")

# 🌟 初始化 Gemini AI 大腦 (升級為 2.5 Flash Lite)
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
    # 🚀 使用最新最快的輕量級模型，省時省額度
    ai_model = genai.GenerativeModel('gemini-2.5-flash-lite') 
except Exception as e:
    st.error("⚠️ 無法載入 Gemini API Key，請確認 secrets.toml 設定！")
    st.stop()

st.markdown("""
<style>
    button[data-baseweb="tab"] { background-color: #E6F0F9 !important; border-radius: 8px 8px 0px 0px !important; margin-right: 4px !important; padding: 10px 20px !important; border: 1px solid #AED6F1 !important; border-bottom: none !important; transition: all 0.3s ease; }
    button[data-baseweb="tab"] p { font-size: 1.25em !important; font-weight: bold !important; color: #2C3E50 !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #154360 !important; border: 1px solid #0B2331 !important; border-bottom: 3px solid #0B2331 !important; }
    button[data-baseweb="tab"][aria-selected="true"] p { color: #FFFFFF !important; }
    .morandi-title { background-color: #738A96; color: white; padding: 15px 20px; border-radius: 8px; font-weight: bold; font-size: 1.5em; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .morandi-dark-title { background-color: #5C6B73; color: white; padding: 12px 20px; border-radius: 6px; font-weight: bold; font-size: 1.3em; margin-bottom: 10px; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    div.stButton > button[kind="primary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 30px !important; background-color: #D4A373 !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.stButton > button[kind="secondary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 30px !important; background-color: #8FAAB8 !important; color: white !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.stButton > button[kind="secondary"]:hover { background-color: #738A96 !important; }
    [data-testid="stDownloadButton"] button { background-color: #D4A373 !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 24px !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .raw-box { background-color: #FDF6E3; padding: 20px; border-left: 5px solid #E6C27A; border-radius: 6px; height: 300px; overflow-y: auto; line-height: 1.8; font-size: 1.1em; color: #333;}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 資料庫連線與精準讀寫引擎
# ==========================================
@st.cache_resource
def get_gcp_credentials():
    sk = st.secrets["gcp_oauth"].to_dict()
    return Credentials(token=None, refresh_token=sk["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=sk["client_id"], client_secret=sk["client_secret"])

def load_sheet(sheet_name):
    try: return pd.DataFrame(gspread.authorize(get_gcp_credentials()).open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet(sheet_name).get_all_records()).rename(columns=lambda x: str(x).strip())
    except: return pd.DataFrame()

def save_raw_to_db(record):
    try:
        ws = gspread.authorize(get_gcp_credentials()).open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet("AI新聞資料庫")
        row = [
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            record['題號'], record['中文標題'], record['新聞日期'],
            record['新聞標題'], record['原始內容'], "", 
            ",".join(record['照片清單']), record['新聞連結']
        ]
        ws.append_row(row)
        return True
    except Exception: return False

def delete_rows_from_db(row_indices):
    try:
        ws = gspread.authorize(get_gcp_credentials()).open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet("AI新聞資料庫")
        for r_idx in sorted(row_indices, reverse=True):
            ws.delete_rows(r_idx)
        return True
    except Exception: return False

def update_q_ids_in_db(updates):
    try:
        ws = gspread.authorize(get_gcp_credentials()).open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet("AI新聞資料庫")
        headers = ws.get_all_values()[0]
        q_id_col = headers.index('對應題號') + 1
        q_title_col = headers.index('中文標題') + 1
        cells = []
        for u in updates:
            cells.append(gspread.Cell(row=u['row'], col=q_id_col, value=str(u['new_id'])))
            cells.append(gspread.Cell(row=u['row'], col=q_title_col, value=str(u['new_title'])))
        ws.update_cells(cells)
        return True
    except Exception: return False

def update_ai_summary_by_row(row_idx, ai_summary):
    try:
        ws = gspread.authorize(get_gcp_credentials()).open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet("AI新聞資料庫")
        headers = ws.get_all_values()[0]
        ai_idx = headers.index('AI摘要') + 1
        ws.update_cell(row_idx, ai_idx, str(ai_summary))
        return True
    except Exception: return False

# ==========================================
# 🕷️ 爬蟲引擎
# ==========================================
def get_news_list(start_date, end_date):
    base_url = "https://www.ncyu.edu.tw"
    headers = {'User-Agent': 'Mozilla/5.0'}
    news_list = []
    seen_links = set()
    page = 1
    old_news_count = 0 
    
    while page <= 40: 
        list_url = f"{base_url}/ncyu/Subject?nodeId=835&page={page}"
        try:
            req = requests.get(list_url, headers=headers, timeout=10, verify=False)
            req.encoding = 'utf-8'
            soup = BeautifulSoup(req.text, 'html.parser')
            links = soup.find_all('a', href=True)
            has_news_in_page = False
            
            for a in links:
                href = a['href']
                if 'Subject/Detail/' in href and 'nodeId=835' in href:
                    has_news_in_page = True
                    full_link = href if href.startswith('http') else (base_url + href if href.startswith('/') else base_url + "/ncyu/" + href)
                    if full_link in seen_links: continue
                    seen_links.add(full_link)
                    
                    node = a
                    date_str = ""
                    for _ in range(4):
                        if not node.parent: break
                        text = node.parent.get_text(separator=' ')
                        date_match = re.search(r'(20\d{2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})', text)
                        if date_match:
                            y, m, d = date_match.groups()
                            date_str = f"{y}-{int(m):02d}-{int(d):02d}"
                            break
                        node = node.parent
                    
                    if date_str:
                        news_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                        if news_date < start_date: old_news_count += 1
                        else: old_news_count = 0 
                        
                        if start_date <= news_date <= end_date:
                            title = a.get('title', '').strip() or a.text.strip()
                            title = re.sub(r'\s+', ' ', title).strip() 
                            news_list.append({"新聞日期": date_str, "新聞標題": title, "新聞連結": full_link})
                            
            if not has_news_in_page: break 
            if old_news_count > 25: break 
        except Exception: break
        page += 1
        time.sleep(0.2)
    return news_list

# 🌟 升級版：純淨內文萃取器 (照片說明文字終極抹除)
def get_news_content(url):
    base_url = "https://www.ncyu.edu.tw"
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        d_req = requests.get(url, headers=headers, timeout=10, verify=False)
        d_req.encoding = 'utf-8'
        d_soup = BeautifulSoup(d_req.text, 'html.parser')
        
        # 1. 暴力清除所有無關的 HTML 標籤
        for tag in d_soup(["script", "style", "noscript", "header", "footer", "nav", "aside"]):
            tag.decompose()
            
        noise_patterns = re.compile(r'menu|nav|footer|breadcrumb|sidebar|top-bar|bottom-bar|search|language|header', re.I)
        for div in d_soup.find_all(['div', 'ul', 'ol', 'section']):
            class_str = ' '.join(div.get('class', [])) if isinstance(div.get('class'), list) else str(div.get('class', ''))
            id_str = str(div.get('id', ''))
            if noise_patterns.search(class_str) or noise_patterns.search(id_str):
                div.decompose()

        # 2. 尋找精準內文標籤
        specific_classes = ['m-edit', 'user_edit', 'article-content', 'news-content', 'CCms_Content', 'cg-desc', 'editor', 'app-article']
        content_div = None
        for cls in specific_classes:
            content_div = d_soup.find('div', class_=re.compile(cls, re.I))
            if content_div: break
            
        if not content_div:
            best_node = None
            max_score = 0
            for node in d_soup.find_all(['div', 'article', 'main']):
                text_len = len(node.get_text(strip=True))
                if text_len < 50: continue
                link_len = sum(len(a.get_text(strip=True)) for a in node.find_all('a'))
                if text_len > 0 and (link_len / text_len) > 0.4: continue 
                
                p_count = len(node.find_all('p', recursive=False))
                score = text_len + (p_count * 50)
                if score > max_score:
                    max_score = score
                    best_node = node
            content_div = best_node if best_node else d_soup.find('body')

        # 3. 智慧段落接合
        for br in content_div.find_all(['br', 'hr']):
            br.replace_with('\n')
        for block in content_div.find_all(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'tr']):
            block.append('\n')

        # 抽出文字
        full_text = content_div.get_text(separator=' ', strip=True)
        
        # 垃圾字元過濾黑名單
        garbage_words = [
            '國立嘉義大學', ':::', '回首頁', '網站導覽', '分眾導覽', '新生教務專欄', 
            '學生', '教師', '職員工', '校友', '民眾', '聯絡我們', 'ENGLISH', 
            'Select Language', '中文 (繁體)', '中文 (简体)', '日本語', 'Bahasa Indonesia', 
            'हिन्दी', 'Español', 'বাংলা', 'Français', 'العربية', 'English', 'ไทย', 
            'اردو', 'Bahasa Melayu', 'Filipino', 'Tiếng Việt', 'မြန်မာ', '한국어', 
            '請輸入關鍵字', '搜尋', '友善列印(開新視窗)', '分享至臉書(開新視窗)', '分享至Line(開新視窗)', '回上一頁'
        ]
        
        clean_lines = []
        for line in full_text.split('\n'):
            line_str = line.strip()
            line_str = re.sub(r' +', ' ', line_str) # 壓縮多餘空白
            
            # 🚀 新增：強制抹除照片說明與版權文字
            # 抹除形如 "(照片由XXX提供)" 或 "（由XXX攝影）"
            line_str = re.sub(r'[（\(][^）\)]*(照片|攝影|拍攝|提供)[^）\)]*[）\)]', '', line_str)
            # 抹除形如 "圖1：植物醫學系..." 這種從圖號開始一直到句號或結尾的說明文字
            line_str = re.sub(r'(?:^|\s)圖\s*\d+\s*[：:].*?(?:[）\)]|。|$)', '', line_str)
            
            line_str = line_str.strip()
            
            if not line_str: continue
            if line_str in garbage_words: continue 
            if len(line_str) <= 3 and '::' in line_str: continue
            
            clean_lines.append(line_str)
            
        final_text = '\n'.join(clean_lines)
        
        # 4. 抓取圖片
        img_urls = []
        for img in content_div.find_all('img'):
            src = img.get('src')
            if src and not src.startswith('data:'):
                img_link = src if src.startswith('http') else base_url + src
                if 'icon' not in img_link.lower() and 'logo' not in img_link.lower() and 'banner' not in img_link.lower():
                    img_urls.append(img_link)
                    
        return {"原始內容": final_text[:3000], "照片清單": img_urls}
    except Exception: 
        return {"原始內容": "無法擷取內文", "照片清單": []}

# ==========================================
# 🧠 Gemini AI 摘要引擎
# ==========================================
def process_news_with_ai(news_item, q_title):
    prompt = f"""
    你現在是大學永續發展(SDGs)評比專家。
    這篇新聞已經被確認可能符合評比題目『{q_title}』。請根據以下【新聞內容】進行濃縮。
    
    【新聞標題】：{news_item['新聞標題']}
    【新聞內容】：{news_item['原始內容']}
    
    任務要求：
    1. 將新聞濃縮精華為 150~300 字，重點放在與「{q_title}」相關的行動或成效。
    2. 濃縮內容若有列點，請使用「1. 」「2. 」的格式。
    3. 請在摘要的最下方，獨立一行，明確標註這篇新聞對應的 SDGs 指標 (例如：【對應SDGs】：SDG 4 優質教育)。
    """
    try:
        response = ai_model.generate_content(prompt)
        return response.text.strip()
    except Exception: return f"AI 處理失敗"

# ==========================================
# 🌟 Word 產製引擎
# ==========================================
def url_to_bytes(url):
    try:
        req = requests.get(url, timeout=5, verify=False)
        req.raise_for_status()
        return io.BytesIO(req.content)
    except: return None

def format_report_text_to_html(text):
    text = str(text)
    text = re.sub(r'(^|\n)(\d+[\.\)]|[-•*])\s*\n\s*', r'\1\2 ', text)
    html = ""
    for line in text.split('\n'):
        line = line.strip()
        if not line: html += "<div style='height: 10px;'></div>"; continue
        match = re.match(r'^([0-9a-zA-Z]+[\.、\)]|[\(（][0-9a-zA-Z一二三四五六七八九十]+[\)）]|[一二三四五六七八九十]+、|[-•*])\s*(.*)', line)
        if match: 
            marker = match.group(1); content = match.group(2)
            html += f"<div style='padding-left: 2.2em; text-indent: -2.2em; margin-bottom: 5px; line-height: 1.6;'>{marker} {content}</div>"
        else: html += f"<div style='margin-bottom: 5px; line-height: 1.6;'>{line}</div>"
    return html.replace('\n', ' ')

def set_run_font(run):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def generate_ai_word_report(q_id, q_title, news_records):
    doc = Document()
    doc.add_heading(f'{q_id} {q_title}', level=1)
    for idx, record in enumerate(news_records):
        if idx > 0: doc.add_paragraph("="*40)
        doc.add_heading(f"新聞標題：{record['新聞標題']}", level=2)
        p_date = doc.add_paragraph(); p_date.add_run(f"發布日期：{record['新聞日期']}").italic = True
        
        urls = str(record['照片清單']).split(',') if record['照片清單'] else []
        urls = [u for u in urls if u.strip()]
        if urls:
            table = doc.add_table(rows=0, cols=2); table.style = 'Table Grid'
            for i in range(0, len(urls), 2):
                row = table.add_row()
                c1 = row.cells[0]; c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_b1 = url_to_bytes(urls[i])
                if img_b1: 
                    try: p1.add_run().add_picture(img_b1, width=Cm(7.5))
                    except: p1.add_run("(圖檔格式不支援)")
                else: p1.add_run("(圖載入失敗)")
                
                c2 = row.cells[1]; c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p2 = c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if i + 1 < len(urls):
                    img_b2 = url_to_bytes(urls[i+1])
                    if img_b2: 
                        try: p2.add_run().add_picture(img_b2, width=Cm(7.5))
                        except: p2.add_run("(圖檔格式不支援)")
                    else: p2.add_run("(圖載入失敗)")
        doc.add_paragraph()
        for line in str(record['AI摘要']).replace('\r', '').split('\n'):
            line = line.strip()
            if not line: continue
            p = doc.add_paragraph()
            match = re.match(r'^([0-9a-zA-Z]+[\.、\)]|[\(（][0-9a-zA-Z一二三四五六七八九十]+[\)）]|[一二三四五六七八九十]+、|[-•*])\s*(.*)', line)
            if match:
                marker = match.group(1); content = match.group(2)
                p.paragraph_format.left_indent = Cm(0.8); p.paragraph_format.first_line_indent = Cm(-0.8)
                p.add_run(f"{marker} {content}")
            else: p.add_run(line)
        p_link = doc.add_paragraph(); p_link.add_run("原始新聞連結: ").bold = True; p_link.add_run(record['新聞連結'])
        
    for p in doc.paragraphs:
        if not p.style.name.startswith('Heading'): p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for r in p.runs: set_run_font(r)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs: set_run_font(r)
    out = io.BytesIO(); doc.save(out); out.seek(0)
    return out

# ==========================================
# 🚀 網頁主畫面與 UI 邏輯
# ==========================================
st.markdown("<div class='morandi-title'>🤖 綠色大學 AI 新聞智能彙整中心</div>", unsafe_allow_html=True)

tab_scrape, tab_ai, tab_view = st.tabs(["🕷️ 階段一：爬蟲抓取", "🔍 階段二：審核與 AI 改寫", "📥 階段三：檢視與下載"])

# ==========================================
# 🕷️ 階段一：爬蟲抓取與入庫
# ==========================================
with tab_scrape:
    st.markdown("<div class='morandi-dark-title'>⚙️ 第一階段：爬取原始新聞與入庫</div>", unsafe_allow_html=True)
    c_start, c_end, c_btn = st.columns([2, 2, 2])
    with c_start: start_date = st.date_input("新聞起日", datetime.date(2024, 1, 1))
    with c_end: end_date = st.date_input("新聞迄日", datetime.date.today())
    
    st.info("💡 **【免耗額度】** 系統將掃描此區間內的新聞標題，命中關鍵字後抓取內文並「陸續寫入」資料庫保存。")
    
    if st.button("🕷️ 開始啟動關鍵字爬取與入庫", type="secondary", use_container_width=True):
        with st.spinner("🤖 正在掃描並陸續寫入資料庫，請稍候..."):
            df_targets = load_sheet("原始新聞抓取")
            df_existing = load_sheet("AI新聞資料庫")
            
            existing_set = set()
            if not df_existing.empty and '對應題號' in df_existing.columns and '新聞連結' in df_existing.columns:
                for _, r in df_existing.iterrows():
                    existing_set.add(f"{str(r.get('對應題號','')).strip()}_{str(r.get('新聞連結','')).strip()}")
            
            if df_targets.empty:
                st.error("❌ 找不到《原始新聞抓取》工作表，或表內無資料！")
            else:
                st.toast("🕷️ 正在地毯式掃描新聞列表...", icon="👀")
                basic_news_list = get_news_list(start_date, end_date)
                
                if not basic_news_list:
                    st.warning(f"在 {start_date} 至 {end_date} 區間內未抓取到任何新聞。")
                else:
                    st.toast(f"✅ 成功獲取 {len(basic_news_list)} 篇新聞，準備進行比對與入庫...", icon="🎯")
                    
                    matched_tasks = []
                    kw_col = '搜尋關鍵字或判斷準則' if '搜尋關鍵字或判斷準則' in df_targets.columns else '關鍵字或判斷準則'
                    
                    for _, target in df_targets.iterrows():
                        q_id = str(target.get('題號', '')).strip()
                        q_title = target.get('中文標題', '')
                        raw_kw = str(target.get(kw_col, ''))
                        if not raw_kw.strip() or raw_kw.lower() == 'nan': raw_kw = q_title
                        keywords = [k.strip() for k in re.split(r'[、,，]', raw_kw) if k.strip()]
                        
                        if not q_id: continue
                        
                        for news in basic_news_list:
                            clean_title = re.sub(r'\s+', '', news['新聞標題']).lower()
                            for kw in keywords:
                                clean_kw = re.sub(r'\s+', '', kw).lower()
                                if clean_kw and clean_kw in clean_title:
                                    if f"{q_id}_{news['新聞連結']}" not in existing_set:
                                        matched_tasks.append({'q_id': q_id, 'q_title': q_title, 'news': news})
                                    break
                    
                    if not matched_tasks:
                        st.success("🎉 所有符合關鍵字的新聞都已經在資料庫中了，沒有需要新增的新聞！")
                    else:
                        success_count = 0
                        total_tasks = len(matched_tasks)
                        progress_text = st.empty()
                        my_bar = st.progress(0)
                        
                        for i, task in enumerate(matched_tasks):
                            q_id = task['q_id']
                            news_title = task['news']['新聞標題']
                            
                            progress_text.markdown(f"**📥 ({i+1}/{total_tasks}) 寫入中：** 題目 {q_id} v.s. 「{news_title}」")
                            my_bar.progress((i + 1) / total_tasks)
                            
                            detail = get_news_content(task['news']['新聞連結'])
                            full_news = {**task['news'], **detail}
                            
                            record = {
                                '題號': q_id, '中文標題': task['q_title'], '新聞日期': full_news['新聞日期'],
                                '新聞標題': full_news['新聞標題'], '原始內容': full_news['原始內容'],
                                '照片清單': full_news['照片清單'], '新聞連結': full_news['新聞連結']
                            }
                            
                            if save_raw_to_db(record):
                                success_count += 1
                            time.sleep(1.0) 
                                    
                        progress_text.empty()
                        my_bar.empty()
                        st.success(f"🎉 爬蟲作業完成！成功為資料庫新增了 {success_count} 筆原始新聞紀錄，請前往「階段二」進行人工審核與 AI 改寫。")

# ==========================================
# 🔍 階段二：人工審核與 Gemini 智慧改寫
# ==========================================
with tab_ai:
    st.markdown("<div class='morandi-dark-title'>🔍 第二階段：人工審核與 AI 智慧改寫</div>", unsafe_allow_html=True)
    
    df_ai_db = load_sheet("AI新聞資料庫")
    df_targets = load_sheet("原始新聞抓取")
    
    required_cols = ['對應題號', '中文標題', '新聞日期', '新聞標題', '原始內容', 'AI摘要', '照片清單', '新聞連結']
    
    if df_ai_db.empty:
        st.warning("目前資料庫為空，請先至「階段一」進行爬蟲抓取。")
    elif not all(col in df_ai_db.columns for col in required_cols):
        missing = [col for col in required_cols if col not in df_ai_db.columns]
        st.error(f"⚠️ 您的《AI新聞資料庫》缺少必要的欄位標題：`{', '.join(missing)}`")
        st.info("👉 解決方法：請至 Google Sheet 確保第一列有以下 9 個欄位：\n\n`抓取時間`, `對應題號`, `中文標題`, `新聞日期`, `新聞標題`, `原始內容`, `AI摘要`, `照片清單`, `新聞連結`")
    else:
        valid_q_list = []
        if not df_targets.empty:
            for _, r in df_targets.iterrows():
                if str(r.get('題號', '')).strip():
                    valid_q_list.append(f"{str(r['題號']).strip()} - {str(r.get('中文標題', '')).strip()}")
        
        df_ai_db['AI摘要'] = df_ai_db['AI摘要'].astype(str).str.strip()
        df_pending = df_ai_db[df_ai_db['AI摘要'] == ''].copy()
        
        if df_pending.empty:
            st.success("🎉 太棒了！資料庫中所有新聞都已經完成 AI 摘要改寫囉！")
        else:
            df_pending['_row_idx'] = df_pending.index + 2
            
            st.markdown("#### 👀 預覽新聞內文 (人工審核輔助)")
            with st.expander("點此展開選擇要預覽的新聞，確認其內容是否符合題意", expanded=False):
                prev_sel = st.selectbox("選擇預覽新聞", df_pending['新聞標題'].tolist(), key="preview_sel")
                if prev_sel:
                    prev_row = df_pending[df_pending['新聞標題'] == prev_sel].iloc[0]
                    st.markdown(f"**🔗 原文連結：** [{prev_row['新聞連結']}]({prev_row['新聞連結']})")
                    st.markdown(f"<div class='raw-box'>{prev_row['原始內容']}</div>", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown(f"#### 📝 待處理清單 (尚有 <span style='color:red;'>{len(df_pending)}</span> 筆)", unsafe_allow_html=True)
            st.info("💡 您可以在表格內修改「對應題號」，或勾選「刪除」剔除無關新聞。確認無誤後，勾選「跑 AI」讓 Gemini 接手處理！")
            
            df_pending.insert(0, "🤖 跑 AI", False)
            df_pending.insert(1, "🗑️ 刪除", False)
            df_pending['對應題號'] = df_pending['對應題號'].astype(str) + " - " + df_pending['中文標題'].astype(str)
            
            display_cols = ["🤖 跑 AI", "🗑️ 刪除", "對應題號", "新聞標題", "新聞日期", "新聞連結", "_row_idx"]
            
            edited_df = st.data_editor(
                df_pending[display_cols],
                hide_index=True,
                use_container_width=True,
                column_config={
                    "🤖 跑 AI": st.column_config.CheckboxColumn("🤖 跑 AI", default=False),
                    "🗑️ 刪除": st.column_config.CheckboxColumn("🗑️ 刪除", default=False),
                    "對應題號": st.column_config.SelectboxColumn("對應題號 (點擊可換題)", options=valid_q_list, required=True),
                    "新聞連結": st.column_config.LinkColumn("連結"),
                    "_row_idx": None
                },
                disabled=["新聞標題", "新聞日期", "新聞連結"]
            )
            
            st.write("")
            c_save, c_space, c_ai = st.columns([4, 1, 4])
            
            with c_save:
                if st.button("💾 僅儲存題號異動與刪除清單", type="secondary", use_container_width=True):
                    with st.spinner("正在處理資料庫異動..."):
                        del_rows = edited_df[edited_df['🗑️ 刪除']]['_row_idx'].tolist()
                        if del_rows: delete_rows_from_db(del_rows)
                            
                        updates = []
                        for idx, row in edited_df.iterrows():
                            if not row['🗑️ 刪除']:
                                orig_full_id = df_pending.loc[idx, '對應題號']
                                new_full_id = row['對應題號']
                                if orig_full_id != new_full_id:
                                    new_id = new_full_id.split(" - ")[0]
                                    new_title = new_full_id.split(" - ")[1]
                                    updates.append({'row': row['_row_idx'], 'new_id': new_id, 'new_title': new_title})
                        
                        if updates: update_q_ids_in_db(updates)
                            
                        st.success("✅ 異動與刪除已儲存！")
                        time.sleep(1)
                        st.rerun()

            with c_ai:
                if st.button("🚀 啟動 Gemini 智慧改寫已勾選項目", type="primary", use_container_width=True):
                    ai_rows = edited_df[(edited_df['🤖 跑 AI']) & (~edited_df['🗑️ 刪除'])]
                    
                    if ai_rows.empty:
                        st.warning("請先勾選至少一筆要「跑 AI」的新聞！")
                    else:
                        with st.spinner(f"正在呼叫 Gemini 處理 {len(ai_rows)} 筆新聞，請稍候..."):
                            success_ai = 0
                            progress_text2 = st.empty()
                            bar2 = st.progress(0)
                            
                            for i, (_, row) in enumerate(ai_rows.iterrows()):
                                real_row_idx = row['_row_idx']
                                new_full_id = row['對應題號']
                                target_title = new_full_id.split(" - ")[1]
                                
                                progress_text2.markdown(f"**✨ 處理中 ({i+1}/{len(ai_rows)})：** {row['新聞標題']}")
                                bar2.progress((i + 1) / len(ai_rows))
                                
                                full_content = df_pending[df_pending['_row_idx'] == real_row_idx].iloc[0]['原始內容']
                                news_item = {'新聞標題': row['新聞標題'], '原始內容': full_content}
                                
                                ai_summary = process_news_with_ai(news_item, target_title)
                                
                                if update_ai_summary_by_row(real_row_idx, ai_summary):
                                    success_ai += 1
                                    
                                time.sleep(1.5)
                                
                            progress_text2.empty()
                            bar2.empty()
                            st.success(f"🎉 完成！成功生成並寫入了 {success_ai} 筆新聞摘要！請前往「階段三」檢視與下載。")
                            time.sleep(1.5)
                            st.rerun()

# ==========================================
# 📥 階段三：檢視與下載彙整區
# ==========================================
with tab_view:
    st.markdown("<div class='morandi-dark-title'>📥 依題目檢視 AI 彙整成果</div>", unsafe_allow_html=True)
    
    df_ai_db = load_sheet("AI新聞資料庫")
    
    if df_ai_db.empty or not all(col in df_ai_db.columns for col in required_cols): 
        st.info("💡 目前資料庫中尚未有紀錄或欄位不齊全。")
    else:
        df_ai_db['AI摘要'] = df_ai_db['AI摘要'].astype(str).str.strip()
        df_completed = df_ai_db[df_ai_db['AI摘要'] != '']
        
        if df_completed.empty:
            st.warning("目前尚無完成 AI 摘要的資料，請先至「階段二」執行處理。")
        else:
            reported_q_ids = df_completed['對應題號'].astype(str).str.strip().unique().tolist()
            q_options = []
            for qid in reported_q_ids:
                title_match = df_completed[df_completed['對應題號'].astype(str).str.strip() == qid]
                title = title_match.iloc[0]['中文標題'] if not title_match.empty else ""
                q_options.append(f"{qid} - {title}")
                
            q_options.sort()
            sel_q = st.selectbox("選擇欲下載之題目", ["請選擇..."] + q_options, label_visibility="collapsed")
            
            if sel_q != "請選擇...":
                v_q_id = sel_q.split(" - ")[0]
                v_q_title = sel_q.split(" - ")[1]
                
                df_records = df_completed[df_completed['對應題號'].astype(str).str.strip() == v_q_id]
                
                st.markdown("---")
                st.markdown(f"<h2 style='color:#154360;'>📖 {v_q_id} {v_q_title}</h2>", unsafe_allow_html=True)
                
                records_to_print = []
                
                for idx, row in df_records.iterrows():
                    st.markdown(f"#### 📰 新聞標題：{row['新聞標題']}")
                    st.caption(f"📅 發布日期：{row['新聞日期']}")
                    
                    urls = str(row.get('照片清單', '')).split(',')
                    urls = [u for u in urls if u.strip()]
                    
                    if urls:
                        table_html = "<table style='width:100%; table-layout:fixed; border-collapse:collapse; border:1px solid #D9E0E3; background-color:white; margin-bottom: 20px;'>"
                        for i in range(0, len(urls), 2):
                            table_html += "<tr>"
                            table_html += f"<td style='border:1px solid #D9E0E3; padding:10px; text-align:center; width:50%; vertical-align:middle;'><img src='{urls[i]}' style='width:100%; max-height:300px; object-fit:contain; border-radius:8px;'></td>"
                            if i + 1 < len(urls):
                                table_html += f"<td style='border:1px solid #D9E0E3; padding:10px; text-align:center; width:50%; vertical-align:middle;'><img src='{urls[i+1]}' style='width:100%; max-height:300px; object-fit:contain; border-radius:8px;'></td>"
                            else:
                                table_html += "<td style='border:1px solid #D9E0E3; width:50%;'></td>"
                            table_html += "</tr>"
                        table_html += "</table>"
                        st.markdown(table_html, unsafe_allow_html=True)
                    
                    c_left, c_right = st.columns([1, 1])
                    with c_left:
                        st.markdown("**📄 原始新聞內容**")
                        st.markdown(f"<div style='background-color:#FDF6E3; padding:20px; border-radius:8px; border-left:5px solid #E6C27A; height:350px; overflow-y:auto; color:#555;'>{row['原始內容']}</div>", unsafe_allow_html=True)
                    with c_right:
                        st.markdown("**✨ Gemini 濃縮摘要 (含 SDGs)**")
                        fmt_ai = format_report_text_to_html(row['AI摘要'])
                        st.markdown(f"<div style='background-color:#E6F0F9; padding:20px; border-radius:8px; border-left:5px solid #8FAAB8; height:350px; overflow-y:auto; color:#154360; font-size:1.1em;'>{fmt_ai}</div>", unsafe_allow_html=True)
                    
                    st.markdown(f"<br>**🔗 原始新聞連結:** [{row['新聞連結']}]({row['新聞連結']})", unsafe_allow_html=True)
                    st.markdown("<hr style='border:1px dashed #ccc;'>", unsafe_allow_html=True)
                    
                    records_to_print.append({
                        '新聞標題': row['新聞標題'], '新聞日期': row['新聞日期'],
                        'AI摘要': row['AI摘要'], '照片清單': ",".join(urls),
                        '新聞連結': row['新聞連結']
                    })
                    
                st.markdown("<div style='text-align: center; margin-bottom: 10px;'><b>準備好下載這份 AI 智能彙整報告了嗎？</b></div>", unsafe_allow_html=True)
                with st.spinner("⏳ 產製 Word 檔中..."):
                    docx_bytes = generate_ai_word_report(v_q_id, v_q_title, records_to_print)
                    
                col1, col2, col3 = st.columns([3, 4, 3])
                with col2:
                    st.download_button(
                        label="📥 下載本題 AI 彙整報告 (Word檔)",
                        data=docx_bytes,
                        file_name=f"{v_q_id}_AI新聞彙整.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )