import streamlit as st
import pandas as pd
import gspread
import datetime
import io
import requests
import re
import base64
import time
from bs4 import BeautifulSoup
import google.generativeai as genai
from google.oauth2.credentials import Credentials
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import qn

# ==========================================
# 🌟 全域設定 & 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！"); st.stop()

st.set_page_config(page_title="嘉大 AI 新聞智能彙整", page_icon="🤖", layout="wide")

# 🌟 初始化 Gemini AI 大腦
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
    # 使用最新強大且快速的模型
    ai_model = genai.GenerativeModel('gemini-1.5-flash') 
except Exception as e:
    st.error("⚠️ 無法載入 Gemini API Key，請確認 secrets.toml 內是否已設定 `GEMINI_API_KEY`！")
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
        st.error(f"寫入資料庫失敗：{e}")
        return False

# ==========================================
# 🕷️ 真實版：嘉大新聞網頁爬蟲 (Web Scraper)
# ==========================================
def scrape_ncyu_news(start_date, end_date):
    st.toast("🕷️ 正在進入嘉義大學新聞櫥窗...", icon="⏳")
    base_url = "https://www.ncyu.edu.tw"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    news_items = []
    seen_links = set()
    
    # 支援翻頁，最多往前爬 3 頁以確保抓取區間
    for page in range(1, 4):
        list_url = f"{base_url}/ncyu/Subject?nodeId=835&page={page}"
        try:
            req = requests.get(list_url, headers=headers, timeout=10)
            req.encoding = 'utf-8'
            soup = BeautifulSoup(req.text, 'html.parser')
            
            # 尋找所有新聞連結
            links = soup.find_all('a', href=True)
            for a in links:
                href = a['href']
                # 確認是新聞內文連結
                if 'Subject?nodeId=835&id=' in href:
                    full_link = base_url + "/ncyu/" + href if not href.startswith('http') else href
                    
                    if full_link in seen_links: continue
                    seen_links.add(full_link)
                    
                    # 抓取日期 (通常在標題附近)
                    parent_text = a.parent.parent.text
                    date_match = re.search(r'20\d{2}[-/.]\d{2}[-/.]\d{2}', parent_text)
                    if date_match:
                        date_str = date_match.group().replace('/', '-').replace('.', '-')
                        news_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                        
                        # 判斷日期區間
                        if start_date <= news_date <= end_date:
                            title = a.get('title', '').strip() or a.text.strip()
                            title = re.sub(r'\s+', ' ', title) # 清除多餘空白
                            
                            # 進入內文頁面抓取完整內容與照片
                            d_req = requests.get(full_link, headers=headers, timeout=10)
                            d_req.encoding = 'utf-8'
                            d_soup = BeautifulSoup(d_req.text, 'html.parser')
                            
                            # 抓內文
                            content_div = d_soup.find('div', class_=re.compile(r'content|mcont|user_edit|article', re.I))
                            if not content_div: content_div = d_soup.find('body')
                            
                            # 移除網頁腳本，只保留純文字
                            for script in content_div(["script", "style"]): script.decompose()
                            full_text = content_div.text.strip()
                            full_text = re.sub(r'\n+', '\n', full_text)
                            
                            # 抓照片 (過濾掉 icon)
                            img_urls = []
                            for img in content_div.find_all('img'):
                                src = img.get('src')
                                if src and not src.startswith('data:'):
                                    img_link = src if src.startswith('http') else base_url + src
                                    if 'icon' not in img_link.lower() and 'logo' not in img_link.lower():
                                        img_urls.append(img_link)
                            
                            news_items.append({
                                "新聞日期": date_str, "新聞標題": title,
                                "原始內容": full_text[:2000], # 取前2000字，避免AI超載
                                "照片清單": img_urls, "新聞連結": full_link
                            })
                            time.sleep(0.3) # 友善爬蟲延遲
        except Exception as e:
            print(f"第 {page} 頁爬取失敗: {e}")
            
    return news_items

# ==========================================
# 🧠 真實版：Gemini AI 判斷與生成引擎
# ==========================================
def process_news_with_ai(news_item, q_id, q_title, q_desc):
    prompt = f"""
    你現在是大學永續發展(SDGs)評比的專家。
    請根據以下【新聞內容】，判斷是否符合評比題目『{q_title}』的範疇。
    題目的判斷準則為：{q_desc}
    
    【新聞標題】：{news_item['新聞標題']}
    【新聞內容】：{news_item['原始內容']}
    
    任務要求：
    1. 如果這篇新聞與題目「完全無關」，請你【僅回覆】五個字：「不符合此題」。
    2. 如果這篇新聞與題目「有關」，請你將新聞濃縮精華簡寫為 150~300 字。
    3. 濃縮內容若有列點，請使用「1. 」「2. 」的格式。
    4. 請在摘要的最下方，獨立一行，智慧判斷並明確標註這篇新聞對應的 SDGs 指標 (例如：【對應SDGs】：SDG 4 優質教育)。
    """
    try:
        response = ai_model.generate_content(prompt)
        ai_result = response.text.strip()
        
        if "不符合此題" in ai_result:
            return None # 代表 AI 認為無關，丟棄此新聞
        return ai_result
    except Exception as e:
        return f"AI 處理失敗：{e}"

# ==========================================
# 🌟 Word 產製引擎 (照片網格 + 完美縮排)
# ==========================================
def url_to_bytes(url):
    try:
        req = requests.get(url, timeout=5)
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
        
        # 1. 處理照片表格
        urls = str(record['照片清單']).split(',') if record['照片清單'] else []
        urls = [u for u in urls if u.strip()]
        
        if urls:
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            for i in range(0, len(urls), 2):
                row = table.add_row()
                # 第一張
                c1 = row.cells[0]; c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_b1 = url_to_bytes(urls[i])
                if img_b1: 
                    try: p1.add_run().add_picture(img_b1, width=Cm(7.5))
                    except: p1.add_run("(圖檔格式不支援)")
                else: p1.add_run("(圖載入失敗)")
                
                # 第二張
                c2 = row.cells[1]; c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p2 = c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if i + 1 < len(urls):
                    img_b2 = url_to_bytes(urls[i+1])
                    if img_b2: 
                        try: p2.add_run().add_picture(img_b2, width=Cm(7.5))
                        except: p2.add_run("(圖檔格式不支援)")
                    else: p2.add_run("(圖載入失敗)")
        
        # 2. 處理摘要與縮排
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
        
        # 3. 連結
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
    
    st.info("💡 系統將讀取《原始新聞抓取》總表，根據題意自動爬取此區間內的嘉大新聞，由 Gemini 判斷關聯、濃縮摘要並存入《AI新聞資料庫》。")
    
    if st.button("🚀 開始啟動 AI 批次抓取與彙整", type="primary", use_container_width=True):
        with st.spinner("🤖 爬蟲運作與 AI 摘要生成中，請切勿關閉網頁..."):
            
            df_targets = load_sheet("原始新聞抓取")
            if df_targets.empty:
                st.error("❌ 找不到《原始新聞抓取》工作表，或表內無資料！")
            else:
                # 執行真實爬蟲
                scraped_news = scrape_ncyu_news(start_date, end_date)
                
                if not scraped_news:
                    st.warning("此時間區間內未抓取到任何新聞，請放寬日期範圍。")
                else:
                    st.toast(f"✅ 成功抓取 {len(scraped_news)} 篇新聞，準備交給 AI 逐題判斷...", icon="🧠")
                    
                    success_count = 0
                    
                    # 顯示即時進度條
                    progress_text = st.empty()
                    my_bar = st.progress(0)
                    total_tasks = len(df_targets) * len(scraped_news)
                    current_task = 0
                    
                    # 巢狀迴圈：每一題去對比每一篇新聞
                    for _, target in df_targets.iterrows():
                        q_id = target.get('題號', '')
                        q_title = target.get('中文標題', '')
                        q_desc = target.get('關鍵字或判斷準則', q_title)
                        if not q_id: continue
                        
                        for news in scraped_news:
                            current_task += 1
                            progress_text.markdown(f"**🔍 正在分析：** 題目 {q_id} v.s. 新聞「{news['新聞標題']}」")
                            my_bar.progress(current_task / total_tasks)
                            
                            # 呼叫 Gemini 判斷
                            ai_summary = process_news_with_ai(news, q_id, q_title, q_desc)
                            
                            if ai_summary: # AI 判定符合
                                record = {
                                    '題號': q_id, '中文標題': q_title, '新聞日期': news['新聞日期'],
                                    '新聞標題': news['新聞標題'], '原始內容': news['原始內容'],
                                    'AI摘要': ai_summary, '照片清單': news['照片清單'], '新聞連結': news['新聞連結']
                                }
                                if save_to_ai_db(record):
                                    success_count += 1
                                    
                    progress_text.empty()
                    my_bar.empty()
                    st.success(f"🎉 批次處理大功告成！AI 成功篩選並寫入了 {success_count} 筆符合的新聞摘要至資料庫！請至【檢視與下載區】查看。")

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