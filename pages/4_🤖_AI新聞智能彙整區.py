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

# 🌟 初始化 Gemini AI 大腦
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
    ai_model = genai.GenerativeModel('gemini-1.5-flash') 
except Exception as e:
    st.error("⚠️ 無法載入 Gemini API Key，請確認 secrets.toml 設定！")
    st.stop()

st.markdown("""
<style>
    button[data-baseweb="tab"] { background-color: #E6F0F9 !important; border-radius: 8px 8px 0px 0px !important; margin-right: 4px !important; padding: 10px 20px !important; border: 1px solid #AED6F1 !important; border-bottom: none !important; transition: all 0.3s ease; }
    button[data-baseweb="tab"] p { font-size: 1.4em !important; font-weight: bold !important; color: #2C3E50 !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #154360 !important; border: 1px solid #0B2331 !important; border-bottom: 3px solid #0B2331 !important; }
    button[data-baseweb="tab"][aria-selected="true"] p { color: #FFFFFF !important; }
    .morandi-title { background-color: #738A96; color: white; padding: 15px 20px; border-radius: 8px; font-weight: bold; font-size: 1.5em; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .morandi-dark-title { background-color: #5C6B73; color: white; padding: 12px 20px; border-radius: 6px; font-weight: bold; font-size: 1.3em; margin-bottom: 10px; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    div.stButton > button[kind="primary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 30px !important; background-color: #D4A373 !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    [data-testid="stDownloadButton"] button { background-color: #D4A373 !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 24px !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 資料庫連線
# ==========================================
@st.cache_resource
def get_gcp_credentials():
    sk = st.secrets["gcp_oauth"].to_dict()
    return Credentials(token=None, refresh_token=sk["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=sk["client_id"], client_secret=sk["client_secret"])

def load_sheet(sheet_name):
    try: return pd.DataFrame(gspread.authorize(get_gcp_credentials()).open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet(sheet_name).get_all_records()).rename(columns=lambda x: str(x).strip())
    except: return pd.DataFrame()

def save_to_ai_db(record):
    try:
        ws = gspread.authorize(get_gcp_credentials()).open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet("AI新聞資料庫")
        row = [
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            record['題號'], record['中文標題'], record['新聞日期'],
            record['新聞標題'], record['原始內容'], record['AI摘要'],
            ",".join(record['照片清單']), record['新聞連結']
        ]
        ws.append_row(row)
        return True
    except Exception as e:
        return False

# ==========================================
# 🕷️ 階段一：地毯式掃描清單 (日期終極校正版)
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
                    
                    # 🌟 向上溯源尋找日期，並加入「無差別空白容錯」的正則表達式
                    node = a
                    date_str = ""
                    for _ in range(4):
                        if not node.parent: break
                        text = node.parent.get_text(separator=' ')
                        # 完美相容 2025/10/28, 2025-10-28, 2025 / 10 / 28 等各種格式
                        date_match = re.search(r'(20\d{2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})', text)
                        if date_match:
                            y, m, d = date_match.groups()
                            # 強制校正並統一格式為 YYYY-MM-DD
                            date_str = f"{y}-{int(m):02d}-{int(d):02d}"
                            break
                        node = node.parent
                    
                    if date_str:
                        news_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                        
                        if news_date < start_date:
                            old_news_count += 1
                        else:
                            old_news_count = 0 
                        
                        if start_date <= news_date <= end_date:
                            title = a.get('title', '').strip() or a.text.strip()
                            title = re.sub(r'\s+', ' ', title).strip() 
                            news_list.append({"新聞日期": date_str, "新聞標題": title, "新聞連結": full_link})
                            
            if not has_news_in_page: break 
            if old_news_count > 25: break 
        except Exception:
            break
        page += 1
        time.sleep(0.2)
        
    return news_list

# ==========================================
# 🕷️ 階段二：精準抓取內文與照片
# ==========================================
def get_news_content(url):
    base_url = "https://www.ncyu.edu.tw"
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        d_req = requests.get(url, headers=headers, timeout=10, verify=False)
        d_req.encoding = 'utf-8'
        d_soup = BeautifulSoup(d_req.text, 'html.parser')
        
        content_div = d_soup.find('div', class_=re.compile(r'content|mcont|user_edit|article', re.I))
        if not content_div: content_div = d_soup.find('body')
        
        for script in content_div(["script", "style"]): script.decompose()
        full_text = content_div.text.strip()
        full_text = re.sub(r'\n+', '\n', full_text)
        
        img_urls = []
        for img in content_div.find_all('img'):
            src = img.get('src')
            if src and not src.startswith('data:'):
                img_link = src if src.startswith('http') else base_url + src
                if 'icon' not in img_link.lower() and 'logo' not in img_link.lower():
                    img_urls.append(img_link)
                    
        return {"原始內容": full_text[:2500], "照片清單": img_urls}
    except Exception:
        return {"原始內容": "無法擷取內文", "照片清單": []}

# ==========================================
# 🧠 Gemini AI 摘要引擎
# ==========================================
def process_news_with_ai(news_item, q_title):
    prompt = f"""
    你現在是大學永續發展(SDGs)評比專家。
    這篇新聞的標題已經被確認符合評比題目『{q_title}』。請根據以下【新聞內容】進行濃縮。
    
    【新聞標題】：{news_item['新聞標題']}
    【新聞內容】：{news_item['原始內容']}
    
    任務要求：
    1. 將新聞濃縮精華為 150~300 字，重點放在與題目相關的行動或成效。
    2. 濃縮內容若有列點，請使用「1. 」「2. 」的格式。
    3. 請在摘要的最下方，獨立一行，智慧標註這篇新聞對應的 SDGs 指標 (例如：【對應SDGs】：SDG 4 優質教育)。
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
        p_date = doc.add_paragraph()
        p_date.add_run(f"發布日期：{record['新聞日期']}").italic = True
        
        urls = str(record['照片清單']).split(',') if record['照片清單'] else []
        urls = [u for u in urls if u.strip()]
        
        if urls:
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
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
        
        p_link = doc.add_paragraph()
        p_link.add_run("原始新聞連結: ").bold = True
        p_link.add_run(record['新聞連結'])
        
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

tab_ai, tab_view = st.tabs(["🤖 AI 智慧搜尋與生成區", "📥 檢視與下載彙整區"])

# ==========================================
# 🏷️ TAB 1: AI 智慧搜尋與生成區
# ==========================================
with tab_ai:
    st.markdown("<div class='morandi-dark-title'>⚙️ 批次處理任務設定</div>", unsafe_allow_html=True)
    c_start, c_end, c_btn = st.columns([2, 2, 2])
    with c_start: start_date = st.date_input("新聞起日", datetime.date(2024, 1, 1))
    with c_end: end_date = st.date_input("新聞迄日", datetime.date.today())
    
    st.info("💡 系統會先比對「新聞標題」與「搜尋關鍵字」(以「、」分隔)，無差別忽略所有空白與特殊格式，命中後才會抓取內文並交由 Gemini 摘要！")
    
    if st.button("🚀 開始啟動關鍵字爬取與 AI 彙整", type="primary", use_container_width=True):
        with st.spinner("🤖 正在掃描新聞列表並比對關鍵字，請勿關閉網頁..."):
            df_targets = load_sheet("原始新聞抓取")
            if df_targets.empty:
                st.error("❌ 找不到《原始新聞抓取》工作表，或表內無資料！")
            else:
                st.toast("🕷️ 階段一：正在地毯式掃描新聞列表...", icon="👀")
                basic_news_list = get_news_list(start_date, end_date)
                
                if not basic_news_list:
                    st.warning(f"在 {start_date} 至 {end_date} 區間內未抓取到任何新聞。請確認網站連線狀態。")
                else:
                    st.toast(f"✅ 成功獲取 {len(basic_news_list)} 篇新聞，準備進行極速關鍵字比對...", icon="🎯")
                    
                    with st.expander(f"👀 點擊查看這 {len(basic_news_list)} 篇被納入比對範圍的新聞清單"):
                        for n in basic_news_list:
                            st.text(f"[{n['新聞日期']}] {n['新聞標題']}")
                    
                    matched_tasks = []
                    kw_col = '搜尋關鍵字或判斷準則' if '搜尋關鍵字或判斷準則' in df_targets.columns else '關鍵字或判斷準則'
                    
                    for _, target in df_targets.iterrows():
                        q_id = target.get('題號', '')
                        q_title = target.get('中文標題', '')
                        
                        raw_kw = str(target.get(kw_col, ''))
                        if not raw_kw.strip() or raw_kw.lower() == 'nan': raw_kw = q_title
                            
                        keywords = [k.strip() for k in re.split(r'[、,，]', raw_kw) if k.strip()]
                        if not q_id: continue
                        
                        for news in basic_news_list:
                            is_match = False
                            clean_title = re.sub(r'\s+', '', news['新聞標題']).lower()
                            
                            for kw in keywords:
                                clean_kw = re.sub(r'\s+', '', kw).lower()
                                if clean_kw and clean_kw in clean_title:
                                    is_match = True
                                    break
                                    
                            if is_match:
                                matched_tasks.append({'q_id': q_id, 'q_title': q_title, 'news': news})
                    
                    if not matched_tasks:
                        st.warning("⚠️ 新聞清單已抓取，但在比對時，沒有任何新聞的【標題】符合您設定的關鍵字。")
                    else:
                        success_count = 0
                        total_tasks = len(matched_tasks)
                        progress_text = st.empty()
                        my_bar = st.progress(0)
                        
                        for i, task in enumerate(matched_tasks):
                            q_id = task['q_id']
                            news_title = task['news']['新聞標題']
                            
                            progress_text.markdown(f"**🔍 ({i+1}/{total_tasks}) 命中！正在解析內文與 AI 摘要：** 題目 {q_id} v.s. 「{news_title}」")
                            my_bar.progress((i + 1) / total_tasks)
                            
                            detail = get_news_content(task['news']['新聞連結'])
                            full_news = {**task['news'], **detail}
                            
                            ai_summary = process_news_with_ai(full_news, task['q_title'])
                            
                            if ai_summary: 
                                record = {
                                    '題號': q_id, '中文標題': task['q_title'], '新聞日期': full_news['新聞日期'],
                                    '新聞標題': full_news['新聞標題'], '原始內容': full_news['原始內容'],
                                    'AI摘要': ai_summary, '照片清單': full_news['照片清單'], '新聞連結': full_news['新聞連結']
                                }
                                if save_to_ai_db(record):
                                    success_count += 1
                                    
                            time.sleep(1)
                                    
                        progress_text.empty()
                        my_bar.empty()
                        st.success(f"🎉 批次處理大功告成！共命中 {total_tasks} 篇，AI 成功寫入 {success_count} 筆新聞摘要至資料庫！請至【檢視與下載區】查看。")

# ==========================================
# 🏷️ TAB 2: 檢視與下載彙整區
# ==========================================
with tab_view:
    st.markdown("<div class='morandi-dark-title'>📥 依題目檢視 AI 彙整成果</div>", unsafe_allow_html=True)
    
    df_ai_db = load_sheet("AI新聞資料庫")
    
    if df_ai_db.empty: 
        st.info("💡 目前資料庫中尚未有任何真實的 AI 處理紀錄。請先至 Tab 1 執行抓取。")
    else:
        reported_q_ids = df_ai_db['對應題號'].astype(str).str.strip().unique().tolist()
        q_options = []
        for qid in reported_q_ids:
            title_match = df_ai_db[df_ai_db['對應題號'].astype(str).str.strip() == qid]
            title = title_match.iloc[0]['中文標題'] if not title_match.empty else ""
            q_options.append(f"{qid} - {title}")
            
        q_options.sort()
        sel_q = st.selectbox("選擇題目", ["請選擇..."] + q_options, label_visibility="collapsed")
        
        if sel_q != "請選擇...":
            v_q_id = sel_q.split(" - ")[0]
            v_q_title = sel_q.split(" - ")[1]
            
            df_records = df_ai_db[df_ai_db['對應題號'].astype(str).str.strip() == v_q_id]
            
            st.markdown("---")
            st.markdown(f"<h2 style='color:#154360;'>📖 {v_q_id} {v_q_title}</h2>", unsafe_allow_html=True)
            
            records_to_print = []
            
            for idx, row in df_records.iterrows():
                st.markdown(f"#### 📰 新聞標題：{row['新聞標題']}")
                st.caption(f"📅 發布日期：{row['新聞發布日期']}")
                
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
                    st.markdown(f"<div style='background-color:#FDF6E3; padding:20px; border-radius:8px; border-left:5px solid #E6C27A; height:350px; overflow-y:auto; color:#555;'>{row['新聞完整內容']}</div>", unsafe_allow_html=True)
                with c_right:
                    st.markdown("**✨ Gemini 濃縮摘要 (含 SDGs)**")
                    fmt_ai = format_report_text_to_html(row['AI摘要'])
                    st.markdown(f"<div style='background-color:#E6F0F9; padding:20px; border-radius:8px; border-left:5px solid #8FAAB8; height:350px; overflow-y:auto; color:#154360; font-size:1.1em;'>{fmt_ai}</div>", unsafe_allow_html=True)
                
                st.markdown(f"<br>**🔗 原始新聞連結:** [{row['新聞連結']}]({row['新聞連結']})", unsafe_allow_html=True)
                st.markdown("<hr style='border:1px dashed #ccc;'>", unsafe_allow_html=True)
                
                records_to_print.append({
                    '新聞標題': row['新聞標題'], '新聞日期': row['新聞發布日期'],
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