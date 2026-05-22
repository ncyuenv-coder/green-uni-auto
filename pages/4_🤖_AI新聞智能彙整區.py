import streamlit as st
import pandas as pd
import gspread
import datetime
import io
import requests
import re
import time
import random
import urllib3
import base64
import zipfile
from bs4 import BeautifulSoup
from PIL import Image
import google.generativeai as genai
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
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
# 修正：set_page_config 必須是第一個被呼叫的 st 指令
st.set_page_config(page_title="嘉大 AI 新聞智能彙整", page_icon="📰", layout="wide")

# 常數集中管理
SHEET_ID = '1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8'
NEWS_IMG_FOLDER_ID = "1VNOna4gRdtTIiFPc4XqMJU2jP01LcyXM" 

if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！"); st.stop()

if st.session_state.get("username") != "admin_ui":
    st.error("🚫 權限不足！此頁面僅限系統管理員 (admin_ui) 存取。")
    st.stop()

if 'ready_zip' not in st.session_state:
    st.session_state['ready_zip'] = None

# 🌟 初始化 Gemini AI 大腦
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
    ai_model = genai.GenerativeModel('gemini-2.5-flash-lite') 
except Exception as e:
    st.error(f"⚠️ 無法載入 Gemini API Key，或設定有誤：{e}")
    st.stop()

st.markdown("""
<style>
    button[data-baseweb="tab"] { background-color: #E6F0F9 !important; border-radius: 8px 8px 0px 0px !important; margin-right: 4px !important; padding: 10px 20px !important; border: 1px solid #AED6F1 !important; border-bottom: none !important; transition: all 0.3s ease; }
    button[data-baseweb="tab"] p { font-size: 1.25em !important; font-weight: bold !important; color: #2C3E50 !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #154360 !important; border: 1px solid #0B2331 !important; border-bottom: 3px solid #0B2331 !important; }
    button[data-baseweb="tab"][aria-selected="true"] p { color: #FFFFFF !important; }
    .morandi-title { background-color: #738A96; color: white; padding: 15px 20px; border-radius: 8px; font-weight: bold; font-size: 1.5em; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .morandi-dark-title { background-color: #5C6B73; color: white; padding: 12px 20px; border-radius: 6px; font-weight: bold; font-size: 1.3em; margin-bottom: 10px; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    div.stButton > button[kind="primary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.2em !important; padding: 10px 20px !important; background-color: #D4A373 !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.stButton > button[kind="secondary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.2em !important; padding: 10px 20px !important; background-color: #8FAAB8 !important; color: white !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.stButton > button[kind="secondary"]:hover { background-color: #738A96 !important; }
    [data-testid="stDownloadButton"] button { background-color: #5C6B73 !important; color: #FFFFFF !important; border: none !important; border-radius: 8px !important; font-weight: bold !important; font-size: 1.2em !important; padding: 12px 24px !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important; }
    [data-testid="stDownloadButton"] button:hover { background-color: #45545E !important; }
    .raw-box { background-color: #FDF6E3; padding: 20px; border-left: 5px solid #E6C27A; border-radius: 6px; height: 500px; overflow-y: auto; line-height: 1.8; font-size: 1.1em; color: #333;}
    .action-box { background-color: #F4F6F6; padding: 20px; border-radius: 8px; border: 1px solid #D5DBDB; margin-top: 15px; }
    
    /* 🌟 階段零：視覺化卡片與深灰藍色 Expander 樣式 */
    [data-testid="stExpander"] details summary { background-color: #45545E !important; color: white !important; border-radius: 6px; }
    [data-testid="stExpander"] details summary p { color: white !important; font-size: 1.2em !important; font-weight: bold; }
    [data-testid="stExpander"] details summary svg { fill: white !important; }
    
    .topic-chip { display: inline-block; background-color: #E6F0F9; color: #154360; padding: 6px 12px; border-radius: 20px; font-size: 0.95em; font-weight: bold; margin: 4px 4px 4px 0; border: 1px solid #AED6F1; }
    
    /* 🌟 修改關鍵字按鈕：莫蘭迪橘色 */
    div.stButton > button:has(p:contains("修改關鍵字")) {
        background-color: #D4A373 !important; color: white !important; border: none !important; border-radius: 6px !important; font-size: 1.1em !important;
    }
    div.stButton > button:has(p:contains("修改關鍵字")):hover {
        background-color: #C39262 !important;
    }
    
    /* 🌟 刪除此題按鈕：莫蘭迪深灰藍色 */
    div.stButton > button:has(p:contains("刪除此題")) {
        background-color: #5C6B73 !important; color: white !important; border: none !important; border-radius: 6px !important; font-size: 1.1em !important;
    }
    div.stButton > button:has(p:contains("刪除此題")):hover {
        background-color: #34495E !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 資料庫連線與精準讀寫引擎 (重構集中化)
# ==========================================
@st.cache_resource
def init_google_apis():
    sk = st.secrets["gcp_oauth"].to_dict()
    creds = Credentials(token=None, refresh_token=sk["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=sk["client_id"], client_secret=sk["client_secret"])
    gc = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return gc, drive_service

@st.cache_data(ttl=600)
def load_sheet(sheet_name, max_retries=3):
    gc, _ = init_google_apis()
    for attempt in range(max_retries):
        try:
            ws = gc.open_by_key(SHEET_ID).worksheet(sheet_name)
            all_data = ws.get_all_values()
            if not all_data or len(all_data) < 2:
                return pd.DataFrame(columns=[str(x).strip() for x in (all_data[0] if all_data else [])])
            headers = [str(x).strip() for x in all_data[0]]
            
            # 建立 DataFrame
            df = pd.DataFrame(all_data[1:], columns=headers)
            
            # [穩定度優化] 1. 容錯處理：若表頭誤寫為「題號」，自動對齊為「對應題號」
            if '題號' in df.columns and '對應題號' not in df.columns:
                df = df.rename(columns={'題號': '對應題號'})
            
            # [穩定度優化] 2. 過濾幽靈空白列：移除整列都是空白的無效資料，避免後續 DataFrame 判斷異常
            df = df.replace('', pd.NA).dropna(how='all').fillna('')
            
            return df
        except Exception:
            if attempt < max_retries - 1: time.sleep(2)
            else: return pd.DataFrame()

def save_raw_to_db(ws, record, max_retries=5):
    for attempt in range(max_retries):
        try:
            row = [
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                str(record['題號']).strip(), str(record['中文標題']).strip(), record['新聞日期'],
                record['新聞標題'], record['原始內容'], "", 
                ",".join(record['照片清單']), record['新聞連結']
            ]
            ws.append_row(row)
            return True
        except Exception:
            if attempt < max_retries - 1: time.sleep(random.uniform(2.0, 5.0))
            else: return False

def delete_rows_from_db(ws, row_indices, max_retries=3):
    for attempt in range(max_retries):
        try:
            for r_idx in sorted(row_indices, reverse=True):
                ws.delete_rows(int(r_idx))
            return True
        except Exception as e:
            if attempt < max_retries - 1: time.sleep(2)
            else: 
                st.error(f"Sheet 刪除失敗: {e}")
                return False

def update_q_ids_in_db(ws, updates, max_retries=3):
    for attempt in range(max_retries):
        try:
            headers = [str(h).strip() for h in ws.get_all_values()[0]]
            q_id_col = headers.index('對應題號') + 1
            q_title_col = headers.index('中文標題') + 1
            cells = []
            for u in updates:
                cells.append(gspread.Cell(row=int(u['row']), col=int(q_id_col), value=str(u['new_id'])))
                cells.append(gspread.Cell(row=int(u['row']), col=int(q_title_col), value=str(u['new_title'])))
            ws.update_cells(cells)
            return True
        except Exception:
            if attempt < max_retries - 1: time.sleep(2)
            else: return False

def update_ai_summary_by_row(ws, row_idx, ai_col_idx, ai_summary, max_retries=5):
    for attempt in range(max_retries):
        try:
            ws.update_cell(int(row_idx), int(ai_col_idx), str(ai_summary))
            return True
        except Exception:
            if attempt < max_retries - 1: time.sleep(random.uniform(2.0, 5.0))
            else: return False

def update_photo_list_by_row(ws, row_idx, new_photo_str, max_retries=3):
    for attempt in range(max_retries):
        try:
            headers = [str(h).strip() for h in ws.get_all_values()[0]]
            col_idx = headers.index('照片清單') + 1
            ws.update_cell(int(row_idx), col_idx, new_photo_str)
            return True
        except Exception:
            if attempt < max_retries - 1: time.sleep(random.uniform(1.0, 3.0))
            else: return False

# ==========================================
# 🌟 Drive 讀取、寫入、命名與刪除引擎
# ==========================================
def delete_drive_files(file_ids, max_retries=3):
    if not file_ids: return
    _, drive_service = init_google_apis()
    try:
        for file_id in file_ids:
            fid = file_id.strip()
            if not fid: continue
            for attempt in range(max_retries):
                try:
                    drive_service.files().delete(fileId=fid).execute()
                    break
                except Exception:
                    if attempt < max_retries - 1: time.sleep(1)
    except Exception: pass 

def rename_drive_files(file_ids, new_q_id, news_title, max_retries=3):
    if not file_ids: return
    _, drive_service = init_google_apis()
    try:
        safe_title = re.sub(r'[/\\:*?"<>|]', '', news_title)[:15]
        for idx, file_id in enumerate(file_ids):
            fid = file_id.strip()
            if not fid: continue
            new_name = f"News_{new_q_id}_{safe_title}_{idx+1}.jpg"
            file_metadata = {'name': new_name}
            for attempt in range(max_retries):
                try:
                    drive_service.files().update(fileId=fid, body=file_metadata).execute()
                    break
                except Exception:
                    if attempt < max_retries - 1: time.sleep(1)
    except Exception: pass 

def drive_id_to_bytes(file_id, max_retries=3):
    _, drive_service = init_google_apis()
    for attempt in range(max_retries):
        try:
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: status, done = downloader.next_chunk()
            fh.seek(0)
            return fh
        except Exception:
            if attempt < max_retries - 1: time.sleep(random.uniform(1.0, 2.0))
    return None

@st.cache_data(ttl=86400, show_spinner=False)
def get_drive_image_info(file_id):
    fh = drive_id_to_bytes(file_id)
    if not fh: return None
    
    file_bytes = fh.read()
    is_landscape = True
    is_panorama = False
    try:
        img = Image.open(io.BytesIO(file_bytes))
        ratio = img.width / img.height
        if ratio < 1.0: is_landscape = False
        if ratio >= 1.9: is_panorama = True 
    except Exception: pass
    
    return {
        'b64': base64.b64encode(file_bytes).decode('utf-8'),
        'is_landscape': is_landscape,
        'is_panorama': is_panorama
    }

# ==========================================
# 🕸️ 爬蟲引擎與圖片雙重瘦身上傳模組 
# ==========================================
def download_compress_upload_image(img_url, drive_service, file_name_prefix, max_retries=3):
    for attempt in range(max_retries):
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            req = requests.get(img_url, headers=headers, timeout=15, verify=False)
            req.raise_for_status()
            img_bytes = req.content

            img = Image.open(io.BytesIO(img_bytes))
            if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                
            MAX_SIZE = (1280, 1280)
            img.thumbnail(MAX_SIZE, Image.Resampling.LANCZOS)

            out_io = io.BytesIO()
            img.save(out_io, format='JPEG', optimize=True, quality=85)
            out_io.seek(0)

            if len(out_io.getvalue()) > len(img_bytes):
                out_io = io.BytesIO(img_bytes)
                mime_type = req.headers.get('Content-Type', 'image/jpeg')
            else:
                mime_type = 'image/jpeg'

            file_meta = {'name': f"{file_name_prefix}.jpg", 'parents': [NEWS_IMG_FOLDER_ID]}
            media = MediaIoBaseUpload(out_io, mimetype=mime_type, resumable=True)
            file = drive_service.files().create(body=file_meta, media_body=media, fields='id').execute()
            return file.get('id')
        except Exception:
            if attempt < max_retries - 1: time.sleep(random.uniform(2.0, 4.0))
            else: return None

def get_news_list(start_date, end_date):
    base_url = "https://www.ncyu.edu.tw"
    headers = {'User-Agent': 'Mozilla/5.0'}
    news_list = []
    seen_links = set()
    page = 1
    old_news_count = 0 
    
    while page <= 40: 
        list_url = f"{base_url}/ncyu/Subject?nodeId=835&page={page}"
        success = False
        
        for attempt in range(3):
            try:
                req = requests.get(list_url, headers=headers, timeout=15, verify=False)
                req.encoding = 'utf-8'
                success = True
                break
            except Exception:
                time.sleep(2)
                
        if not success: break
        
        try:
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
        time.sleep(0.5)
    return news_list

def get_news_content(url, max_retries=3):
    base_url = "https://www.ncyu.edu.tw"
    headers = {'User-Agent': 'Mozilla/5.0'}
    
    for attempt in range(max_retries):
        try:
            d_req = requests.get(url, headers=headers, timeout=15, verify=False)
            d_req.encoding = 'utf-8'
            d_soup = BeautifulSoup(d_req.text, 'html.parser')
            
            for tag in d_soup(["script", "style", "noscript", "header", "footer", "nav", "aside"]):
                tag.decompose()
                
            noise_patterns = re.compile(r'menu|nav|footer|breadcrumb|sidebar|top-bar|bottom-bar|search|language|header', re.I)
            for div in d_soup.find_all(['div', 'ul', 'ol', 'section']):
                class_str = ' '.join(div.get('class', [])) if isinstance(div.get('class'), list) else str(div.get('class', ''))
                id_str = str(div.get('id', ''))
                if noise_patterns.search(class_str) or noise_patterns.search(id_str):
                    div.decompose()

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

            for br in content_div.find_all(['br', 'hr']):
                br.replace_with('\n')
            for block in content_div.find_all(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'tr']):
                block.append('\n')

            full_text = content_div.get_text(separator=' ', strip=True)
            
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
                line_str = re.sub(r' +', ' ', line_str) 
                line_str = re.sub(r'[（\(][^）\)]*(照片|攝影|拍攝|提供)[^）\)]*[）\)]', '', line_str)
                line_str = re.sub(r'(?:^|\s)圖\s*\d+\s*[：:].*?(?:[）\)]|。|$)', '', line_str)
                line_str = line_str.strip()
                
                if not line_str: continue
                if line_str in garbage_words: continue 
                if len(line_str) <= 3 and '::' in line_str: continue
                clean_lines.append(line_str)
                
            final_text = '\n'.join(clean_lines)
            
            img_urls = []
            for img in content_div.find_all('img'):
                src = img.get('src')
                if src and not src.startswith('data:'):
                    img_link = src if src.startswith('http') else base_url + src
                    if 'icon' not in img_link.lower() and 'logo' not in img_link.lower() and 'banner' not in img_link.lower():
                        img_urls.append(img_link)
                        
            return {"原始內容": final_text[:3000], "照片清單": img_urls}
        except Exception:
            if attempt < max_retries - 1: time.sleep(2)
            
    return {"原始內容": "無法擷取內文", "照片清單": []}

# ==========================================
# 🧠 Gemini AI 摘要引擎 
# ==========================================
def process_news_with_ai(news_item, q_title):
    prompt = f"""
    你現在是大學永續發展(SDGs)評比專家。
    這篇新聞已經被確認可能符合評比題目『{q_title}』。請根據以下【新聞內容】進行濃縮與分析。
    
    【新聞內容】：{news_item['原始內容']}
    
    任務要求：
    1. 將新聞摘要濃縮為 150~250 字，重點放在與「{q_title}」相關的具體行動或成效。
    2. 濃縮內容若有列點，請使用「1. 」「2. 」的格式。
    3. 內容排版請使用純文字，【絕對不要】使用 Markdown 的粗體符號（例如：不要寫成 **專業人才培育**，請直接寫 專業人才培育）。
    4. 獨立於摘要之外，請在最下方獨立段落，明確標註這篇新聞所對應的「所有」SDGs 細項指標。若新聞內容有對應多個，請【務必全部列上】，不限於一個。請參照聯合國永續發展指標網站 (globalgoals.tw) 的標準，給出精準的「細項指標代碼」與「中英文完整說明」。
       格式範例：
       【對應SDGs】：
       SDG 8.8：Protect labour rights and promote safe and secure working environments for all workers, including migrant workers, in particular women migrants, and those in precarious employment.(保護勞工權利和促進安全的工作環境)
       SDG 11.4：Strengthen efforts to protect and safeguard the world’s cultural and natural heritage.(進一步努力保護和捍衛世界文化與自然遺產)
    5. 摘要內容請直接開始，絕對不要在開頭重複寫出「新聞標題」。
    """
    try:
        response = ai_model.generate_content(prompt)
        try:
            return response.text.strip()
        except ValueError:
            return f"⚠️ 內容可能因觸發 Google 安全機制（如特定敏感詞彙）而被拒絕摘要。"
    except Exception as e: 
        return f"❌ 系統錯誤: {str(e)}"

# ==========================================
# 🌟 Word 報表產製與 ZIP 壓縮模組
# ==========================================
def format_report_text_to_html(text):
    text = str(text)
    text = re.sub(r'(^|\n)(\d+[\.\)]|[-•*])\s*\n\s*', r'\1\2 ', text)
    html = ""
    for line in text.split('\n'):
        line = line.strip()
        line = line.replace('**', '') 
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
        
        raw_file_ids = str(record['照片清單']).split(',')
        file_ids = [fid.strip() for fid in raw_file_ids if fid.strip()]
        
        if file_ids:
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            
            img_infos = []
            for fid in file_ids:
                info = get_drive_image_info(fid)
                if info: img_infos.append(info)
                
            chunks = []
            current_row = []
            slots_filled = 0
            
            for info in img_infos:
                slots_needed = 2 if info['is_panorama'] else 1
                if slots_filled + slots_needed > 2:
                    chunks.append((current_row.copy(), slots_filled))
                    current_row.clear()
                    slots_filled = 0
                
                current_row.append(info)
                slots_filled += slots_needed
                
                if slots_filled == 2:
                    chunks.append((current_row.copy(), slots_filled))
                    current_row.clear()
                    slots_filled = 0
                    
            if current_row:
                chunks.append((current_row.copy(), slots_filled))
                
            for chunk_items, filled in chunks:
                row = table.add_row()
                if len(chunk_items) == 1 and chunk_items[0]['is_panorama']:
                    info = chunk_items[0]
                    cell = row.cells[0]
                    cell.merge(row.cells[1])
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        img_bytes = base64.b64decode(info['b64'])
                        inline = p.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(15.5))
                        try:
                            spPr = inline._inline.xpath('.//pic:spPr')[0]
                            spPr.append(parse_xml('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:softEdge rad="31750"/></a:effectLst>'))
                        except: pass
                    except: p.add_run("(圖檔格式不支援)")
                else:
                    for j in range(2):
                        cell = row.cells[j]
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        if j < len(chunk_items):
                            info = chunk_items[j]
                            p = cell.paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            try:
                                img_bytes = base64.b64decode(info['b64'])
                                if info['is_landscape']:
                                    inline = p.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(7.5), height=Cm(5.5))
                                else:
                                    inline = p.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(7.0), height=Cm(9.0))
                                try:
                                    spPr = inline._inline.xpath('.//pic:spPr')[0]
                                    spPr.append(parse_xml('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:softEdge rad="31750"/></a:effectLst>'))
                                except: pass
                            except: p.add_run("(圖檔格式不支援)")
                            
            row_title = table.add_row()
            cell_title = row_title.cells[0]
            cell_title.merge(row_title.cells[1])
            cell_title.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p_title = cell_title.paragraphs[0]
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(record['新聞標題'])
            r_title.bold = True
                    
        doc.add_paragraph()
        for line in str(record['AI摘要']).replace('\r', '').split('\n'):
            line = line.strip()
            line = line.replace('**', '')
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

def create_all_reports_zip(df_completed):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        reported_q_ids = df_completed['對應題號'].astype(str).str.strip().unique().tolist()
        for qid in reported_q_ids:
            df_records = df_completed[df_completed['對應題號'].astype(str).str.strip() == qid]
            if df_records.empty: continue
            
            q_title = df_records.iloc[0]['中文標題']
            records_to_print = []
            
            for _, row in df_records.iterrows():
                file_ids = str(row.get('照片清單', '')).split(',')
                file_ids = [fid.strip() for fid in file_ids if fid.strip()]
                records_to_print.append({
                    '新聞標題': row['新聞標題'], '新聞日期': row['新聞日期'],
                    'AI摘要': row['AI摘要'], '照片清單': ",".join(file_ids),
                    '新聞連結': row['新聞連結']
                })
                
            docx_bytes = generate_ai_word_report(qid, q_title, records_to_print)
            safe_title = re.sub(r'[/\\:*?"<>|]', '', q_title)[:30]
            zip_file.writestr(f"{qid}_{safe_title}_AI新聞彙整.docx", docx_bytes.getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

# ==========================================
# 分頁模組
# ==========================================
@st.fragment
def render_tab_manage():
    st.markdown("<div class='morandi-dark-title'>📝 初始階段：題目與關鍵字管理</div>", unsafe_allow_html=True)
    
    if st.button("🔄 同步最新題庫", type="secondary"):
        load_sheet.clear()
        st.rerun()

    df_targets = load_sheet("原始新聞抓取")

    with st.expander("➕ 點此新增「題目與關鍵字」", expanded=False):
        with st.form("add_topic_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            new_id = c1.text_input("對應題號 (如 1.1)")
            new_title = c2.text_input("中文名稱")
            new_kw = st.text_input("搜尋關鍵字 (請用逗號 , 分隔，如：綠能,太陽能,節能)")
            if st.form_submit_button("💾 新增題目入庫"):
                if new_id and new_title and new_kw:
                    gc, _ = init_google_apis()
                    ws = gc.open_by_key(SHEET_ID).worksheet("原始新聞抓取")
                    ws.append_row([new_id, new_title, new_kw])
                    st.success("✅ 新增成功！")
                    load_sheet.clear()
                    time.sleep(1)
                    st.rerun()
                else: st.error("⚠️ 請完整填寫所有欄位。")

    st.markdown("#### 📂 題目與搜尋關鍵字設定")
    if df_targets.empty: 
        st.info("目前無任何題目，請由上方新增。")
    else:
        for idx, row in df_targets.iterrows():
            real_row_idx = idx + 2
            q_id = str(row.get('對應題號', row.get('題號', ''))).strip()
            
            # 🌟 終極防護：優先抓取「中文標題/名稱」，若無則「強制抓取 B 欄 (iloc[1])」
            q_title = str(row.get('中文名稱', row.get('中文標題', row.iloc[1] if len(row) > 1 else '')))).strip()
            if not q_title or q_title.lower() == 'nan':
                q_title = "(未命名)"
                
            kw_col = '搜尋關鍵字或判斷準則' if '搜尋關鍵字或判斷準則' in df_targets.columns else '關鍵字或判斷準則'
            raw_kw = str(row.get(kw_col, '')).strip()

            if not q_id: continue

            chips_html = "".join([f"<span class='topic-chip'>#{k.strip()}</span>" for k in raw_kw.split(',') if k.strip()])
            
            c_left, c_right = st.columns([6, 4])
            
            with c_left:
                st.markdown(f"""
                <div style='background-color: #8FAAB8; padding: 25px; border-radius: 10px; margin-bottom: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 100%; display: flex; flex-direction: column; justify-content: center;'>
                    <div style='font-size: 2.2em; font-weight: 900; color: white; line-height: 1.1;'>{q_id}</div>
                    <div style='font-size: 1.4em; font-weight: 600; color: #E8EEF2; margin-top: 8px;'>{q_title}</div>
                </div>
                """, unsafe_allow_html=True)
                
            with c_right:
                # 🌟 將底色改為左側 #8FAAB8 的刷淡同色系 (#E4ECF0)
                st.markdown(f"""
                <div style='background-color: #E4ECF0; padding: 15px; border-radius: 10px; border: 1px solid #C4D3DD; margin-bottom: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); min-height: 90px;'>
                    <div style='color: #4A606D; font-weight: bold; margin-bottom: 10px;'>🔑 搜尋關鍵字：</div>
                    {chips_html}
                </div>
                """, unsafe_allow_html=True)
                
                c_b1, c_b2 = st.columns(2)
                with c_b1:
                    with st.popover("✏️ 修改關鍵字", use_container_width=True):
                        edit_kw = st.text_input("編輯關鍵字 (逗號分隔)", value=raw_kw, key=f"edit_{real_row_idx}")
                        if st.button("💾 儲存修改", key=f"save_{real_row_idx}", use_container_width=True):
                            gc, _ = init_google_apis()
                            ws = gc.open_by_key(SHEET_ID).worksheet("原始新聞抓取")
                            headers = [str(h).strip() for h in ws.get_all_values()[0]]
                            k_idx = headers.index(kw_col) + 1
                            ws.update_cell(real_row_idx, k_idx, edit_kw)
                            st.success("✅ 更新成功！")
                            load_sheet.clear()
                            time.sleep(1)
                            st.rerun()
                with c_b2:
                    if st.button("🗑️ 刪除此題", key=f"del_{real_row_idx}", use_container_width=True):
                        gc, _ = init_google_apis()
                        ws = gc.open_by_key(SHEET_ID).worksheet("原始新聞抓取")
                        delete_rows_from_db(ws, [real_row_idx])
                        st.success("✅ 刪除成功！")
                        load_sheet.clear()
                        time.sleep(1)
                        st.rerun()
            
            st.markdown("<hr style='margin: 5px 0 15px 0; border: 0; border-top: 2px solid #E6F0F9;'>", unsafe_allow_html=True)
# ==========================================
# 分頁模組
# ==========================================
@st.fragment
def render_tab_scrape():
    st.markdown("<div class='morandi-dark-title'>⚙️ 第一階段：爬取原始新聞與入庫</div>", unsafe_allow_html=True)
    
    df_targets = load_sheet("原始新聞抓取")
    if df_targets.empty:
        st.error("❌ 題庫為空，請先至「初始階段」建立題目！")
        return

    # 🌟 增強下拉選單資料抓取：強制鎖定 B 欄 (iloc[1])
    target_options = []
    for _, r in df_targets.iterrows():
        q_id = str(r.get('對應題號', r.get('題號', ''))).strip()
        q_title = str(r.get('中文名稱', r.get('中文標題', r.iloc[1] if len(r) > 1 else '')))).strip()
        if not q_title or q_title.lower() == 'nan': q_title = "(未命名)"
        if q_id: target_options.append(f"{q_id} - {q_title}")

    if 'scan_done' not in st.session_state: st.session_state.scan_done = False
    if 'green_zone_saved' not in st.session_state: st.session_state.green_zone_saved = []
    if 'yellow_zone' not in st.session_state: st.session_state.yellow_zone = []
    if 'skipped_count' not in st.session_state: st.session_state.skipped_count = 0

    c_opt1, c_opt2 = st.columns([3, 1])
    with c_opt1:
        selected_targets = st.multiselect("🎯 選擇爬取範圍 (若留空則代表掃描全校題目)", target_options, placeholder="請選擇一題或多題...")
    with c_opt2:
        st.write("") 
        if st.button("🧹 重置掃描狀態", use_container_width=True):
            st.session_state.scan_done = False
            st.session_state.green_zone_saved = []
            st.session_state.yellow_zone = []
            st.session_state.skipped_count = 0
            st.rerun()

    c_start, c_end = st.columns(2)
    with c_start: start_date = st.date_input("新聞起日", datetime.date(2024, 1, 1))
    with c_end: end_date = st.date_input("新聞迄日", datetime.date.today())

    # ==================================
    # 動作一：外網掃描 + 綠燈直接寫入 + 黃燈保留
    # ==================================
    if not st.session_state.scan_done:
        if st.button("🕷️ 動作一：外網掃描與分析比對", type="secondary", use_container_width=True):
            with st.spinner("🤖 正在掃描並分析新聞命中狀況，請稍候..."):
                df_existing = load_sheet("AI新聞資料庫")
                existing_urls = set(df_existing['新聞連結'].astype(str).str.strip().tolist()) if not df_existing.empty and '新聞連結' in df_existing.columns else set()

                basic_news_list = get_news_list(start_date, end_date)
                if not basic_news_list:
                    st.warning("⚠️ 此區間內未抓取到任何新聞！")
                else:
                    news_matches = {}
                    for news in basic_news_list:
                        news_matches[news['新聞連結']] = {'news': news, 'matches': []}

                    targets_to_scan = df_targets
                    if selected_targets:
                        sel_ids = [s.split(" - ")[0] for s in selected_targets]
                        col_name = '對應題號' if '對應題號' in df_targets.columns else '題號'
                        targets_to_scan = df_targets[df_targets[col_name].astype(str).str.strip().isin(sel_ids)]

                    kw_col = '搜尋關鍵字或判斷準則' if '搜尋關鍵字或判斷準則' in df_targets.columns else '關鍵字或判斷準則'

                    for _, target in targets_to_scan.iterrows():
                        q_id = str(target.get('對應題號', target.get('題號', ''))).strip()
                        
                        # 🌟 同步增強迴圈內的抓取邏輯
                        q_title = str(target.get('中文名稱', target.get('中文標題', target.iloc[1] if len(target) > 1 else '未命名')))).strip()
                        if not q_title or q_title.lower() == 'nan': q_title = "未命名"
                        
                        raw_kw = str(target.get(kw_col, ''))
                        if not raw_kw.strip() or raw_kw.lower() == 'nan': raw_kw = q_title
                        keywords = [k.strip() for k in re.split(r'[、,，]', raw_kw) if k.strip()]

                        if not q_id: continue

                        for news in basic_news_list:
                            url = news['新聞連結']
                            clean_title = re.sub(r'\s+', '', news['新聞標題']).lower()
                            for kw in keywords:
                                clean_kw = re.sub(r'\s+', '', kw).lower()
                                if clean_kw and clean_kw in clean_title:
                                    news_matches[url]['matches'].append({'q_id': q_id, 'q_title': q_title})
                                    break

                    green_zone_tasks = []
                    yellow_zone = []
                    skipped_count = 0

                    for url, data in news_matches.items():
                        if len(data['matches']) == 0: continue
                        if url in existing_urls: 
                            skipped_count += 1
                            continue
                        if len(data['matches']) == 1: green_zone_tasks.append(data)
                        else: yellow_zone.append(data)

                green_saved_list = []
                if green_zone_tasks:
                    st.info(f"🟢 發現 {len(green_zone_tasks)} 篇單一命中新聞，正在自動為您擷取與寫入資料庫...")
                    progress_text = st.empty()
                    my_bar = st.progress(0)
                    total_g = len(green_zone_tasks)
                    
                    gc, drive_service = init_google_apis()
                    ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")
                    
                    for i, item in enumerate(green_zone_tasks):
                        q_id = item['matches'][0]['q_id']
                        q_title = item['matches'][0]['q_title']
                        news_title = item['news']['新聞標題']
                        
                        progress_text.markdown(f"**📥 綠燈區 ({i+1}/{total_g}) 擷取與上傳中：** {q_id} v.s. 「{news_title}」")
                        my_bar.progress((i + 1) / total_g)
                        
                        detail = get_news_content(item['news']['新聞連結'])
                        full_news = {**item['news'], **detail}
                        
                        uploaded_file_ids = []
                        for img_idx, img_url in enumerate(full_news['照片清單']):
                            clean_name_prefix = f"News_{q_id}_{re.sub(r'[/\\:*?\"<>|]', '', news_title)[:15]}_{img_idx+1}"
                            fid = download_compress_upload_image(img_url, drive_service, clean_name_prefix)
                            if fid: uploaded_file_ids.append(fid)
                            
                        record = {
                            '題號': q_id, '中文標題': q_title, '新聞日期': full_news['新聞日期'],
                            '新聞標題': full_news['新聞標題'], '原始內容': full_news['原始內容'],
                            '照片清單': uploaded_file_ids, '新聞連結': full_news['新聞連結']
                        }
                        if save_raw_to_db(ws_ai_db, record): 
                            green_saved_list.append(f"【{q_id}】 {news_title}")
                        time.sleep(1.0)
                        
                    progress_text.empty()
                    my_bar.empty()
                    st.success(f"✅ {len(green_saved_list)} 篇單一命中新聞已成功自動入庫！")

                st.session_state.green_zone_saved = green_saved_list
                st.session_state.yellow_zone = yellow_zone
                st.session_state.skipped_count = skipped_count
                st.session_state.scan_done = True
                st.rerun()

    # ==================================
    # 動作二：人工分流黃燈區
    # ==================================
    else:
        st.success(f"✅ 掃描與綠燈作業完畢！共過濾掉 {st.session_state.skipped_count} 篇資料庫已存在的重複新聞。")

        st.markdown(f"#### 🟢 綠燈區：已自動入庫名單 ({len(st.session_state.green_zone_saved)} 篇)")
        if not st.session_state.green_zone_saved: 
            st.info("無單一命中新聞。")
        else:
            with st.expander("點擊檢視已成功入庫新聞", expanded=False):
                for title in st.session_state.green_zone_saved:
                    st.markdown(f"✔️ {title}")

        st.markdown(f"#### ⚠️ 黃燈區：多題衝突需人工分流 ({len(st.session_state.yellow_zone)} 篇)")
        if not st.session_state.yellow_zone: 
            st.info("無衝突新聞，太完美了！所有新聞已處理完畢。")
        else:
            for i, item in enumerate(st.session_state.yellow_zone):
                st.markdown(f"""
                <div style='background-color: #FFF9E6; padding: 15px; border-left: 5px solid #F1C40F; border-radius: 6px; margin-bottom: 5px;'>
                    <b style='font-size: 1.1em; color: #2C3E50;'>📰 {item['news']['新聞標題']}</b><br>
                    <span style='color:#7F8C8D; font-size:0.95em;'>📅 發布日期：{item['news']['新聞日期']} &nbsp;|&nbsp; 🔗 <a href='{item['news']['新聞連結']}' target='_blank' style='color:#2980B9; font-weight:bold;'>點擊檢視原始新聞</a></span>
                </div>
                """, unsafe_allow_html=True)
                
                # 🌟 選單現在會帶有透過強制抓取 B 欄獲得的完美中文名稱
                options = [f"{m['q_id']} - {m['q_title']}" for m in item['matches']]
                st.selectbox("🎯 請為這篇新聞選擇最終歸屬：", options, key=f"yz_sel_{i}")
                st.markdown("<hr style='margin: 15px 0; border-top: 1px dashed #D5DBDB;'>", unsafe_allow_html=True)

            if st.button("🚀 動作二：確認分發並正式寫入黃燈區新聞", type="primary", use_container_width=True):
                final_tasks = []
                for i, item in enumerate(st.session_state.yellow_zone):
                    chosen = st.session_state[f"yz_sel_{i}"]
                    q_id = chosen.split(" - ")[0].strip()
                    q_title = chosen.split(" - ")[1].strip() if " - " in chosen else ""
                    final_tasks.append({'q_id': q_id, 'q_title': q_title, 'news': item['news']})

                with st.spinner(f"正在為您擷取與處理黃燈區 {len(final_tasks)} 篇新聞，請稍候..."):
                    success_count = 0
                    gc, drive_service = init_google_apis()
                    ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")

                    progress_text = st.empty()
                    my_bar = st.progress(0)
                    total_t = len(final_tasks)

                    for i, task in enumerate(final_tasks):
                        q_id = task['q_id']
                        news_title = task['news']['新聞標題']
                        progress_text.markdown(f"**📥 ({i+1}/{total_t}) 擷取與上傳中：** 題目 {q_id} v.s. 「{news_title}」")
                        my_bar.progress((i + 1) / total_t)

                        detail = get_news_content(task['news']['新聞連結'])
                        full_news = {**task['news'], **detail}

                        uploaded_file_ids = []
                        for img_idx, img_url in enumerate(full_news['照片清單']):
                            clean_name_prefix = f"News_{q_id}_{re.sub(r'[/\\:*?\"<>|]', '', news_title)[:15]}_{img_idx+1}"
                            fid = download_compress_upload_image(img_url, drive_service, clean_name_prefix)
                            if fid: uploaded_file_ids.append(fid)

                        record = {
                            '題號': q_id, '中文標題': task['q_title'], '新聞日期': full_news['新聞日期'],
                            '新聞標題': full_news['新聞標題'], '原始內容': full_news['原始內容'],
                            '照片清單': uploaded_file_ids, '新聞連結': full_news['新聞連結']
                        }

                        if save_raw_to_db(ws_ai_db, record): success_count += 1
                        time.sleep(1.0)

                    progress_text.empty()
                    my_bar.empty()
                    load_sheet.clear()
                    
                    st.session_state.scan_done = False
                    st.session_state.green_zone_saved = []
                    st.session_state.yellow_zone = []
                    st.session_state.skipped_count = 0
                    
                    st.success(f"🎉 作業完成！成功將 {success_count} 筆分流新聞入庫，請前往「階段二」進行改寫。")

@st.fragment
def render_tab_ai():
    st.markdown("<div class='morandi-dark-title'>🔍 第二階段：人工審核與 AI 智慧改寫</div>", unsafe_allow_html=True)
    
    # [新增] 手動更新資料庫按鈕
    if st.button("🔄 點我重新讀取/同步最新資料庫", key="refresh_tab2", type="secondary", use_container_width=True):
        load_sheet.clear() # 清除舊快取
        st.rerun()         # 重新渲染畫面以抓取新資料
        
    df_ai_db = load_sheet("AI新聞資料庫")
    df_targets = load_sheet("原始新聞抓取")
    required_cols = ['對應題號', '中文標題', '新聞日期', '新聞標題', '原始內容', 'AI摘要', '照片清單', '新聞連結']
    
    if df_ai_db.empty: st.warning("目前資料庫為空，請先至「階段一」進行爬蟲抓取。")
    elif not all(col in df_ai_db.columns for col in required_cols):
        missing = [col for col in required_cols if col not in df_ai_db.columns]
        st.error(f"⚠️ 您的《AI新聞資料庫》缺少必要的欄位標題：`{', '.join(missing)}`")
    else:
        valid_q_list = []
        if not df_targets.empty:
            for _, r in df_targets.iterrows():
                # [精準修正] 增加容錯：優先讀取 '對應題號'，若無則讀取 '題號'
                qid = str(r.get('對應題號', r.get('題號', ''))).strip()
                qtitle = str(r.get('中文標題', '')).strip()
                if qid: valid_q_list.append(f"{qid} - {qtitle}")
        
        df_ai_db['AI摘要'] = df_ai_db['AI摘要'].astype(str).str.strip()
        df_pending = df_ai_db[df_ai_db['AI摘要'] == ''].copy()
        
        if df_pending.empty: st.success("🎉 太棒了！資料庫中所有新聞都已經完成 AI 摘要改寫囉！")
        else:
            df_pending['_row_idx'] = df_pending.index + 2
            df_pending['對應題號'] = df_pending['對應題號'].astype(str).str.strip() + " - " + df_pending['中文標題'].astype(str).str.strip()
            df_pending['下拉選項'] = "【" + df_pending['對應題號'].astype(str) + "】 " + df_pending['新聞標題'].astype(str)
            
            st.markdown(f"#### 📝 尚有 <span style='color:red;'>{len(df_pending)}</span> 筆待處理新聞", unsafe_allow_html=True)
            selected_news_str = st.selectbox("📌 請選擇要處理的新聞：", df_pending['下拉選項'].tolist())
            
            if selected_news_str:
                target_row = df_pending[df_pending['下拉選項'] == selected_news_str].iloc[0]
                real_row_idx = target_row['_row_idx']
                current_full_id = target_row['對應題號']
                news_title = target_row['新聞標題']
                news_content = target_row['原始內容']
                news_link = target_row['新聞連結']
                news_date = target_row['新聞日期']
                
                st.markdown("---")
                st.markdown(f"### 📰 {news_title}")
                st.caption(f"📅 發布日期：{news_date} | 🎯 目前對應題目：**{current_full_id}**")
                st.markdown(f"**🔗 原文連結：** [{news_link}]({news_link})")
                st.markdown(f"<div class='raw-box'>{news_content}</div>", unsafe_allow_html=True)
                
                st.markdown("<div class='action-box'>", unsafe_allow_html=True)
                st.markdown("### 🛠️ 請選擇處理模式")
                
                col_act1, col_act2 = st.columns([1, 1], gap="large")
                
                with col_act1:
                    st.markdown("#### 🚀 模式一：AI 處理")
                    st.write("確認新聞與題目相符，直接交由 Gemini 濃縮摘要。")
                    if st.button("✨ 啟提 Gemini 智慧改寫", type="primary", use_container_width=True):
                        with st.spinner("呼叫 Gemini 處理中..."):
                            gc, _ = init_google_apis()
                            ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")
                            headers = [str(h).strip() for h in ws_ai_db.get_all_values()[0]]
                            ai_col_idx = headers.index('AI摘要') + 1
                            
                            target_title = current_full_id.split(" - ", 1)[1].strip() if " - " in current_full_id else ""
                            ai_summary = process_news_with_ai({'新聞標題': news_title, '原始內容': news_content}, target_title)
                            
                            if "❌ 系統錯誤" in ai_summary or "⚠️" in ai_summary:
                                st.error(f"AI 處理失敗，請稍後再試：\n{ai_summary}")
                            else:
                                if update_ai_summary_by_row(ws_ai_db, real_row_idx, ai_col_idx, ai_summary):
                                    st.success("🎉 完成！成功生成新聞摘要並寫入資料庫！")
                                    load_sheet.clear()
                                    time.sleep(1.0)
                                    st.rerun()
                                else: st.error("❌ 寫入資料庫失敗，請檢查網路連線。")
                                
                    st.markdown("---")
                    st.markdown("#### 🗑️ 模式二：刪除廢棄")
                    st.write("若此新聞無參考價值，將徹底從資料庫與雲端硬碟移除。")
                    if st.button("🗑️ 徹底刪除此新聞 (含照片)", type="secondary", use_container_width=True):
                        with st.spinner("正在安全刪除照片與紀錄..."):
                            gc, _ = init_google_apis()
                            ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")
                            
                            photos_str = str(target_row.get('照片清單', ''))
                            if photos_str:
                                fids = [f.strip() for f in photos_str.split(',') if f.strip()]
                                if fids: delete_drive_files(fids)
                                
                            if delete_rows_from_db(ws_ai_db, [real_row_idx]):
                                st.success("✅ 該篇新聞與附屬的雲端照片皆已徹底刪除！")
                                load_sheet.clear()
                                time.sleep(1.5)
                                st.rerun()
                            else: st.error("❌ Google Sheet 資料刪除失敗，請再試一次。")
                            
                with col_act2:
                    st.markdown("#### ✏️ 模式三：修改題號")
                    st.write("若此新聞判斷的題目不正確，請從下方選擇正確題號並儲存。")
                    
                    idx_default = valid_q_list.index(current_full_id) if current_full_id in valid_q_list else 0
                    new_q_selection = st.selectbox("選擇新對應題號：", valid_q_list, index=idx_default, label_visibility="collapsed")
                    
                    if st.button("💾 儲存新題號", use_container_width=True):
                        if new_q_selection != current_full_id:
                            with st.spinner("正在為您更新對應題號與照片檔名..."):
                                gc, _ = init_google_apis()
                                ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")
                                
                                new_id = new_q_selection.split(" - ", 1)[0].strip() if " - " in new_q_selection else new_q_selection.strip()
                                new_title = new_q_selection.split(" - ", 1)[1].strip() if " - " in new_q_selection else ""
                                
                                photos_str = str(target_row.get('照片清單', ''))
                                if photos_str:
                                    fids = [f.strip() for f in photos_str.split(',') if f.strip()]
                                    if fids: rename_drive_files(fids, new_id, news_title)
                                
                                if update_q_ids_in_db(ws_ai_db, [{'row': real_row_idx, 'new_id': new_id, 'new_title': new_title}]):
                                    st.success("✅ 題號與雲端照片檔名修改成功！")
                                    load_sheet.clear()
                                    time.sleep(1.5)
                                    st.rerun()
                                else: st.error("❌ 題號修改寫入失敗。")
                        else: st.warning("您選擇的題號與目前相同，無須修改喔！")
                
                st.markdown("</div>", unsafe_allow_html=True)

@st.fragment
def render_tab_view():
    st.markdown("<div class='morandi-dark-title'>📥 依題目檢視與打包下載</div>", unsafe_allow_html=True)
    
    # [新增] 手動更新資料庫按鈕
    if st.button("🔄 點我重新讀取/同步最新資料庫", key="refresh_tab3", type="secondary", use_container_width=True):
        load_sheet.clear() # 清除舊快取
        st.rerun()         # 重新渲染畫面以抓取新資料
        
    df_ai_db = load_sheet("AI新聞資料庫")
    required_cols = ['對應題號', '中文標題', '新聞日期', '新聞標題', '原始內容', 'AI摘要', '照片清單', '新聞連結']
    
    if df_ai_db.empty or not all(col in df_ai_db.columns for col in required_cols): 
        st.info("💡 目前資料庫中尚未有紀錄或欄位不齊全。")
    else:
        df_ai_db['AI摘要'] = df_ai_db['AI摘要'].astype(str).str.strip()
        df_completed = df_ai_db[df_ai_db['AI摘要'] != ''].copy()
        
        if df_completed.empty: st.warning("目前尚無完成 AI 摘要的資料，請先至「階段二」執行處理。")
        else:
            df_completed['_row_idx'] = df_completed.index + 2
            reported_q_ids = df_completed['對應題號'].astype(str).str.strip().unique().tolist()
            q_options = []
            for qid in reported_q_ids:
                title_match = df_completed[df_completed['對應題號'].astype(str).str.strip() == qid]
                title = title_match.iloc[0]['中文標題'] if not title_match.empty else ""
                q_options.append(f"{qid} - {title}")
                
            q_options.sort()
            
            st.markdown("#### 📦 全校成果一次打包")
            st.info("💡 貼心提醒：打包全校數十篇新聞需要較多運算資源。請點擊下方按鈕開始產生壓縮檔。")
            
            if st.button("⚙️ 產生全校彙整 ZIP 壓縮檔", type="secondary", use_container_width=True):
                with st.spinner("後台正為您將所有題目的 Word 檔打包中，請稍候..."):
                    st.session_state['ready_zip'] = create_all_reports_zip(df_completed)
                    st.success("✅ 打包完成！請點擊下方深色按鈕下載。")
                    
            if st.session_state.get('ready_zip'):
                st.download_button(
                    label="📥 下載全部題目彙整 (ZIP 壓縮檔)",
                    data=st.session_state['ready_zip'],
                    file_name=f"全校綠色大學AI彙整成果_{datetime.date.today()}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            
            st.markdown("---")
            st.markdown("#### 🔍 單一題目檢視與下載")
            sel_q = st.selectbox("選擇欲檢視之題目", ["請選擇..."] + q_options, label_visibility="collapsed")
            
            if sel_q != "請選擇...":
                v_q_id = sel_q.split(" - ")[0]
                v_q_title = sel_q.split(" - ")[1]
                df_records = df_completed[df_completed['對應題號'].astype(str).str.strip() == v_q_id]
                
                st.markdown(f"<h2 style='color:#154360;'>📖 {v_q_id} {v_q_title}</h2>", unsafe_allow_html=True)
                records_to_print = []
                
                for idx, row in df_records.iterrows():
                    real_row_idx = row['_row_idx']
                    c_title, c_del = st.columns([8, 2])
                    with c_title:
                        st.markdown(f"#### 📰 【新聞】{row['新聞標題']}")
                        st.caption(f"📅 發布日期：{row['新聞日期']} | 🔗 [點擊查看原始新聞]({row['新聞連結']})")
                    with c_del:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("🗑️ 刪除此篇新聞", key=f"del_news_tab3_{real_row_idx}", use_container_width=True):
                            with st.spinner("刪除新聞與雲端照片中..."):
                                gc, _ = init_google_apis()
                                ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")
                                
                                photos_str = str(row.get('照片清單', ''))
                                if photos_str:
                                    fids = [f.strip() for f in photos_str.split(',') if f.strip()]
                                    if fids: delete_drive_files(fids)
                                    
                                delete_rows_from_db(ws_ai_db, [real_row_idx])
                                load_sheet.clear()
                                st.success("✅ 該篇新聞與附屬照片已徹底刪除！")
                                st.rerun()
                    
                    raw_file_ids = str(row.get('照片清單', '')).split(',')
                    raw_file_ids = [fid.strip() for fid in raw_file_ids if fid.strip()]
                    valid_infos = []
                    failed_fids = []
                    
                    if raw_file_ids:
                        for fid in raw_file_ids:
                            info = get_drive_image_info(fid)
                            if info:
                                info['id'] = fid; valid_infos.append(info)
                            else: failed_fids.append(fid)
                                
                        if failed_fids:
                            valid_fids = [info['id'] for info in valid_infos]
                            gc, _ = init_google_apis()
                            ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")
                            update_photo_list_by_row(ws_ai_db, real_row_idx, ",".join(valid_fids))
                            delete_drive_files(failed_fids)
                            load_sheet.clear()
                        
                        if valid_infos:
                            st.markdown("**📸 新聞照片 (點擊下方按鈕可刪除不需要的圖片)**")
                            ui_rows = []
                            current_ui_row = []
                            current_weight = 0
                            
                            for info in valid_infos:
                                weight = 2 if info['is_panorama'] else 1
                                if current_weight + weight > 4:
                                    ui_rows.append(current_ui_row.copy())
                                    current_ui_row = [info]
                                    current_weight = weight
                                else:
                                    current_ui_row.append(info)
                                    current_weight += weight
                            if current_ui_row: ui_rows.append(current_ui_row.copy())
                                
                            for r_items in ui_rows:
                                weights = [2 if item['is_panorama'] else 1 for item in r_items]
                                total_w = sum(weights)
                                cols = st.columns(weights + [1] * (4 - total_w))
                                
                                for i, info in enumerate(r_items):
                                    with cols[i]:
                                        fid = info['id']
                                        img_height = "auto" if info['is_panorama'] else "200px"
                                        st.markdown(f"<img src='data:image/jpeg;base64,{info['b64']}' style='width:100%; height:{img_height}; object-fit:contain; border-radius:8px; border:1px solid #ddd; margin-bottom:5px;'>", unsafe_allow_html=True)
                                        if st.button("🗑️ 刪除此圖", key=f"del_img_{real_row_idx}_{fid}", use_container_width=True):
                                            with st.spinner("刪除雲端圖片中..."):
                                                delete_drive_files([fid])
                                                new_fids = [f['id'] for f in valid_infos if f['id'] != fid]
                                                gc, _ = init_google_apis()
                                                ws_ai_db = gc.open_by_key(SHEET_ID).worksheet("AI新聞資料庫")
                                                update_photo_list_by_row(ws_ai_db, real_row_idx, ",".join(new_fids))
                                                load_sheet.clear()
                                                st.success("✅ 照片已刪除")
                                                st.rerun()
                    
                    st.markdown("**✨ Gemini 濃縮摘要 (含 SDGs)**")
                    clean_ai_summary = str(row['AI摘要']).replace('**', '')
                    fmt_ai = format_report_text_to_html(clean_ai_summary)
                    st.markdown(f"<div style='background-color:#E6F0F9; padding:20px; border-radius:8px; border-left:5px solid #8FAAB8; color:#154360; font-size:1.1em; line-height: 1.6;'>{fmt_ai}</div>", unsafe_allow_html=True)
                    st.markdown("<hr style='border:1px dashed #ccc; margin-top:20px; margin-bottom:20px;'>", unsafe_allow_html=True)
                    
                    records_to_print.append({
                        '新聞標題': row['新聞標題'], '新聞日期': row['新聞日期'],
                        'AI摘要': clean_ai_summary, '照片清單': ",".join([f['id'] for f in valid_infos]),
                        '新聞連結': row['新聞連結']
                    })
                    
                with st.spinner("⏳ 產製本題 Word 檔中..."):
                    docx_bytes = generate_ai_word_report(v_q_id, v_q_title, records_to_print)
                    
                st.download_button(
                    label=f"📥 單題彙整下載 ({v_q_id} Word檔)",
                    data=docx_bytes,
                    file_name=f"{v_q_id}_AI新聞彙整.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

# ==========================================
# 🚀 執行介面佈署
# ==========================================
st.markdown("<div class='morandi-title'>🤖 綠色大學 AI 新聞智能彙整中心</div>", unsafe_allow_html=True)
tab_manage, tab_scrape, tab_ai, tab_view = st.tabs(["📝 初始階段", "🕷️ 階段一：爬蟲", "🔍 階段二：AI 改寫", "📥 階段三：檢視下載"])

with tab_manage: render_tab_manage()
with tab_scrape: render_tab_scrape()
with tab_ai: render_tab_ai()
with tab_view: render_tab_view()