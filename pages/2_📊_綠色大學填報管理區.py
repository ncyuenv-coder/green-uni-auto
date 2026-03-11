import streamlit as st
import pandas as pd
import gspread
import io
import re
import datetime
import base64
import zipfile
from PIL import Image
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import qn

# ==========================================
# 🌟 全域參數設定 & 資安防護罩 (嚴格權限控管)
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

# 僅限 admin_ui 檢視與管理
if st.session_state.get("username") != "admin_ui":
    st.error("🚫 權限不足！此頁面僅限系統管理員 (admin_ui) 存取。")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報管理區", page_icon="📊", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定
# ==========================================
st.markdown("""
<style>
    /* 標籤頁 (Tabs) 樣式 */
    button[data-baseweb="tab"] { 
        background-color: #E6F0F9 !important; border-radius: 8px 8px 0px 0px !important; 
        margin-right: 4px !important; padding: 10px 20px !important; 
        border: 1px solid #AED6F1 !important; border-bottom: none !important; transition: all 0.3s ease; 
    }
    button[data-baseweb="tab"] p { font-size: 1.4em !important; font-weight: bold !important; color: #2C3E50 !important; }
    button[data-baseweb="tab"][aria-selected="true"] { 
        background-color: #154360 !important; border: 1px solid #0B2331 !important; border-bottom: 3px solid #0B2331 !important; 
    }
    button[data-baseweb="tab"][aria-selected="true"] p { color: #FFFFFF !important; }

    /* [新增] 選項標籤設計 (底色區塊按鈕) */
    .stRadio div[role="radiogroup"] {
        flex-direction: row !important;
        gap: 15px !important;
        margin-bottom: 20px !important;
    }
    .stRadio div[role="radiogroup"] label {
        background-color: #F0F3F4 !important; 
        border: 2px solid #BDC3C7 !important;
        border-radius: 10px !important; 
        padding: 12px 25px !important; 
        cursor: pointer;
        transition: all 0.2s;
    }
    .stRadio div[role="radiogroup"] label:hover { background-color: #E5E8E8 !important; }
    .stRadio div[role="radiogroup"] label p { 
        font-size: 1.2rem !important; 
        font-weight: 800 !important; 
        color: #566573 !important; 
        margin: 0 !important;
    }
    .stRadio div[role="radiogroup"] label[data-checked="true"] { 
        background-color: #34495E !important; 
        border-color: #34495E !important; 
    }
    .stRadio div[role="radiogroup"] label[data-checked="true"] p { color: #FFFFFF !important; }

    /* 各區塊標題 */
    .morandi-select-title { background-color: #9DAB86; color: white; padding: 12px 15px; border-radius: 6px; font-weight: bold; font-size: 1.2em; margin-bottom: 5px; margin-top: 15px; }
    .morandi-question-title { background-color: #948B89; color: white; padding: 15px 18px; border-radius: 6px; font-weight: bold; font-size: 1.4em; margin-bottom: 15px; margin-top: 10px; }
    .morandi-dark-title { background-color: #5C6B73; color: white; padding: 12px 20px; border-radius: 6px; font-weight: bold; font-size: 1.3em; margin-bottom: 10px; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .light-brown-box { background-color: #F5EBE0; color: #2C3E50; padding: 15px 20px; border-left: 6px solid #D4A373; border-radius: 6px; margin-bottom: 5px; line-height: 1.6; }
    
    /* 下載按鈕 */
    div.stButton > button[kind="primary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.4em !important; padding: 12px 30px !important; }
    [data-testid="stDownloadButton"] button { background-color: #D4A373 !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 24px !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    [data-testid="stDownloadButton"] button:hover { background-color: #C39262 !important; }
    [data-testid="stDownloadButton"] button * { color: white !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 初始化及資料庫連線功能
# ==========================================
@st.cache_resource
def get_gcp_credentials():
    skey = st.secrets["gcp_oauth"].to_dict()
    return Credentials(
        token=None, refresh_token=skey.get("refresh_token"), 
        token_uri="https://oauth2.googleapis.com/token", 
        client_id=skey.get("client_id"), client_secret=skey.get("client_secret")
    )

@st.cache_data(ttl=60)
def load_gsheet_data(sheet_name):
    try:
        creds = get_gcp_credentials()
        gc = gspread.authorize(creds)
        sh = gc.open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8')
        df = pd.DataFrame(sh.worksheet(sheet_name).get_all_records())
        df.columns = df.columns.str.strip()
        return df
    except Exception: 
        return pd.DataFrame()

@st.cache_data(ttl=3600, show_spinner=False)
def get_file_info(file_id, desc=""):
    try:
        creds = get_gcp_credentials()
        drive_service = build('drive', 'v3', credentials=creds)
        meta = drive_service.files().get(fileId=file_id, fields='mimeType').execute()
        mime_type = meta.get('mimeType', '')
        
        is_pdf = 'pdf' in mime_type.lower()
        is_image = mime_type.startswith('image/')
        
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: 
            status, done = downloader.next_chunk()
        fh.seek(0)
        file_bytes = fh.read()
        
        is_landscape = True
        is_panorama = False
        if is_image:
            try:
                img = Image.open(io.BytesIO(file_bytes))
                ratio = img.width / img.height
                if ratio < 1.0: is_landscape = False
                if ratio >= 1.9: is_panorama = True 
            except Exception: pass
            
        b64_str = base64.b64encode(file_bytes).decode('utf-8') if is_image else ""
        return {
            'id': file_id, 'desc': desc, 'is_pdf': is_pdf, 'is_image': is_image, 
            'is_landscape': is_landscape, 'is_panorama': is_panorama, 
            'b64': b64_str, 'mime_type': mime_type
        }
    except Exception: 
        return None

# ==========================================
# 🌟 排版模組 (HTML & Word 表格共用引擎)
# ==========================================
def format_report_text_to_html(text):
    text = str(text)
    text = re.sub(r'(^|\n)(\d+[\.\)]|[-•*])\s*\n\s*', r'\1\2 ', text)
    html = ""
    for line in text.split('\n'):
        line = line.strip()
        if not line: 
            html += "<div style='height: 10px;'></div>"
            continue
        if re.match(r'^【.*】$', line):
            html += f"<div style='font-size: 1.15em; font-weight: bold; color: #154360; margin-top: 15px; margin-bottom: 5px; border-bottom: 2px solid #AED6F1; padding-bottom: 3px;'>{line}</div>"
            continue
        match = re.match(r'^([0-9a-zA-Z]+[\.、\)]|[\(（][0-9a-zA-Z一二三四五六七八九十]+[\)）]|[一二三四五六七八九十]+、|[-•*])\s*(.*)', line)
        if match: 
            marker = match.group(1); content = match.group(2)
            html += f"<div style='padding-left: 2.2em; text-indent: -2.2em; margin-bottom: 5px; line-height: 1.6;'>{marker} {content}</div>"
        else: 
            html += f"<div style='margin-bottom: 5px; line-height: 1.6;'>{line}</div>"
    return html.replace('\n', ' ')

def generate_html_image_table(enriched_files_dict):
    has_image = any(f.get('is_image') for files in enriched_files_dict.values() for f in files)
    if not has_image: return ""
    
    table_html = "<table style='width:100%; table-layout:fixed; border-collapse:collapse; border:1px solid #D9E0E3; background-color:white; margin-bottom: 25px;'>"
    for section_title, files in enriched_files_dict.items():
        images = [f for f in files if f.get('is_image')]
        if not images: continue
        
        if len(enriched_files_dict) > 1 or "新聞" in section_title:
            table_html += f"<tr><td colspan='2' style='border:1px solid #D9E0E3; padding:10px; text-align:center; background-color:#E6F0F9;'><strong style='font-size:1.2em; color:#154360;'>{section_title}</strong></td></tr>"
        
        panoramas = [f for f in images if f.get('is_panorama')]
        landscapes = [f for f in images if f.get('is_landscape') and not f.get('is_panorama')]
        portraits = [f for f in images if not f.get('is_landscape')]
        
        for pano in panoramas:
            table_html += f"<tr><td colspan='2' style='border:1px solid #D9E0E3; padding:15px; text-align:center; vertical-align:middle;'><img src='data:{pano['mime_type']};base64,{pano['b64']}' style='width:100%; height:auto; max-height:400px; object-fit:contain; border-radius:8px; margin-bottom:10px;'><br><b>{pano['desc']}</b></td></tr>"
        
        pairs = []
        for i in range(0, len(landscapes) - 1, 2): pairs.append((landscapes[i], landscapes[i+1]))
        rem_l = landscapes[-1] if len(landscapes) % 2 != 0 else None
        for i in range(0, len(portraits) - 1, 2): pairs.append((portraits[i], portraits[i+1]))
        rem_p = portraits[-1] if len(portraits) % 2 != 0 else None
        
        if rem_l and rem_p: pairs.append((rem_l, rem_p))
        elif rem_l: pairs.append((rem_l, None))
        elif rem_p: pairs.append((rem_p, None))
        
        for p1, p2 in pairs:
            table_html += "<tr>"
            for f in [p1, p2]:
                if f:
                    ratio = "7.5/5.5" if f['is_landscape'] else "7/9"
                    table_html += f"<td style='border:1px solid #D9E0E3; padding:15px; text-align:center; width:50%; vertical-align:middle;'><img src='data:{f['mime_type']};base64,{f['b64']}' style='aspect-ratio:{ratio}; width:100%; object-fit:contain; background-color:#f1f1f1; border-radius:8px; margin-bottom:10px;'><br><b>{f['desc']}</b></td>"
                else:
                    table_html += "<td style='border:1px solid #D9E0E3; width:50%;'></td>"
            table_html += "</tr>"
    table_html += "</table>"
    return table_html

def add_images_to_word_table(doc, enriched_files_dict):
    has_image = any(f.get('is_image') for files in enriched_files_dict.values() for f in files)
    if not has_image: return
    
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    for section_title, files in enriched_files_dict.items():
        images = [f for f in files if f.get('is_image')]
        if not images: continue
        
        if len(enriched_files_dict) > 1 or "新聞" in section_title:
            row = table.add_row()
            cell = row.cells[0]
            cell.merge(row.cells[1])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(section_title).bold = True
        
        panoramas = [f for f in images if f.get('is_panorama')]
        landscapes = [f for f in images if f.get('is_landscape') and not f.get('is_panorama')]
        portraits = [f for f in images if not f.get('is_landscape')]
        
        for pano in panoramas:
            row = table.add_row()
            cell = row.cells[0]
            cell.merge(row.cells[1])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                img_bytes = base64.b64decode(pano['b64'])
                inline = p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(15.5))
                try:
                    spPr = inline._inline.xpath('.//pic:spPr')[0]
                    spPr.append(parse_xml('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:softEdge rad="31750"/></a:effectLst>'))
                except: pass
            except: p_img.add_run("(圖片無法插入)")
            cell.add_paragraph(f"{pano['desc']}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        pairs = []
        for i in range(0, len(landscapes) - 1, 2): pairs.append((landscapes[i], landscapes[i+1]))
        rem_l = landscapes[-1] if len(landscapes) % 2 != 0 else None
        for i in range(0, len(portraits) - 1, 2): pairs.append((portraits[i], portraits[i+1]))
        rem_p = portraits[-1] if len(portraits) % 2 != 0 else None
        if rem_l and rem_p: pairs.append((rem_l, rem_p)) 
        elif rem_l: pairs.append((rem_l, None))
        elif rem_p: pairs.append((rem_p, None))
        
        for p1, p2 in pairs:
            row = table.add_row()
            for idx, f in enumerate([p1, p2]):
                cell = row.cells[idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                if f:
                    p_img = cell.paragraphs[0]
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        img_bytes = base64.b64decode(f['b64'])
                        inline = p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(7.5) if f['is_landscape'] else Cm(7.0), height=Cm(5.5) if f['is_landscape'] else Cm(9.0))
                        try:
                            spPr = inline._inline.xpath('.//pic:spPr')[0]
                            spPr.append(parse_xml('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:softEdge rad="31750"/></a:effectLst>'))
                        except: pass
                    except: p_img.add_run("(圖片無法插入)")
                    cell.add_paragraph(f"{f['desc']}").alignment = WD_ALIGN_PARAGRAPH.CENTER

def set_run_font(run):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# ==========================================
# 🌟 核心資料處理與 Word 產製流程 (包含精準新聞抓取)
# ==========================================
def extract_single_question_data(v_q_id, df_q, df_r, df_news):
    q_match_base = df_q[df_q['當年度題目'].astype(str).str.strip() == v_q_id]
    v_q_title = str(q_match_base.iloc[0]['中文標題']).strip() if not q_match_base.empty else "未知標題"
    v_desc_text = str(q_match_base.iloc[0].get('中文說明', '無')).replace('中文說明：', '').strip() if not q_match_base.empty else "無"
    
    # 1. 處理單位填報紀錄
    df_q_records = df_r[df_r['題號'].astype(str).str.strip() == v_q_id].copy()
    if not df_q_records.empty:
        df_q_records['填報時間'] = pd.to_datetime(df_q_records['填報時間'])
        df_q_records = df_q_records.sort_values('填報時間').drop_duplicates(subset=['權責單位'], keep='last')
    
    info_strings = []; combined_reqs = ""; combined_texts = ""; unit_files_dict = {}
    unit_count = len(df_q_records)
    
    for idx, row in df_q_records.iterrows():
        u = str(row['權責單位']).strip()
        rep = str(row.get('填報人', '無')).strip()
        ext = str(row.get('填報人分機', '無')).strip()
        em = str(row.get('填報人電子郵件', '無')).strip()
        info_strings.append(f"填報單位：{u}  |  👤 填報人：{rep}  |  📞 分機：{ext}  |  📧 Email：{em}")
        
        q_match_unit = df_q[(df_q['當年度題目'].astype(str).str.strip() == v_q_id) & (df_q['權責單位'].astype(str).str.strip() == u)]
        req_text = str(q_match_unit.iloc[0].get('資料需求', '無特別說明')) if not q_match_unit.empty else (str(q_match_base.iloc[0].get('資料需求', '無')) if not q_match_base.empty else "無")
        rep_text = str(row.get('填報內容', '無')).strip()
        
        if unit_count > 1:
            combined_reqs += f"【{u}】\n{req_text}\n\n"
            combined_texts += f"【{u}】\n{rep_text}\n\n"
        else:
            combined_reqs = req_text; combined_texts = rep_text
            
        files_for_u = []
        for i in range(1, 11):
            fid = str(row.get(f'檔案{i}_ID', '')).strip()
            desc = str(row.get(f'檔案{i}_說明', '')).strip()
            if fid: files_for_u.append({'id': fid, 'desc': desc})
        unit_files_dict[u] = files_for_u

    unit_enriched_files = {}
    for u, files in unit_files_dict.items():
        enriched = [get_file_info(f['id'], f['desc']) for f in files]
        unit_enriched_files[f"【{u}】"] = [e for e in enriched if e]
        
    # 2. 處理對應新聞資料 (精準抓取 E, G, H, I 欄位)
    news_data = []
    if not df_news.empty:
        # 尋找可能是「題號」的欄位 (通常在前三欄 A, B, C)
        target_col_idx = 0
        for i, col in enumerate(df_news.columns[:5]):
            if '題' in str(col) or '指標' in str(col):
                target_col_idx = i
                break
        
        for _, row in df_news.iterrows():
            row_id = str(row.iloc[target_col_idx]).strip()
            # 支援一個新聞對應多個題號 (例如: "1.1, 1.2")
            ids = [x.strip() for x in re.split(r'[,、/，\s]+', row_id) if x.strip()]
            
            if v_q_id in ids:
                # 依需求指定抓取欄位索引：E欄=4, G欄=6, H欄=7, I欄=8
                n_title = str(row.iloc[4]).strip() if len(row.index) > 4 else ""
                n_summary = str(row.iloc[6]).strip() if len(row.index) > 6 else ""
                n_photos_str = str(row.iloc[7]).strip() if len(row.index) > 7 else ""
                n_link = str(row.iloc[8]).strip() if len(row.index) > 8 else ""
                
                # 防呆機制：若無標題且無摘要，則不顯示該新聞
                if not n_title and not n_summary: continue
                
                # 從字串中萃取 Google Drive 的 檔案 ID
                file_ids = re.findall(r'[-\w]{25,}', n_photos_str)
                n_enriched = [get_file_info(fid, f"新聞配圖 {i+1}") for i, fid in enumerate(file_ids)]
                
                news_data.append({
                    'title': n_title,
                    'summary': n_summary,
                    'link': n_link,
                    'enriched_files': [e for e in n_enriched if e]
                })
                
    return v_q_title, v_desc_text, info_strings, combined_reqs.strip(), combined_texts.strip(), unit_enriched_files, news_data

def build_word_document(v_q_id, v_q_title, info_strings, desc_text, req_text, report_text, unit_enriched_files, news_data):
    doc = Document()
    doc.add_heading(f'嘉大綠色大學填報成果彙整', 0)
    for info in info_strings: doc.add_paragraph(info).alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_heading(f'{v_q_id} {v_q_title}', level=1)
    
    doc.add_heading('題目說明：', level=2)
    for line in str(desc_text).replace('\r', '').split('\n'):
        if line.strip(): doc.add_paragraph(line.strip())
        
    doc.add_heading('資料需求：', level=2)
    for line in str(req_text).replace('<br>', '\n').replace('<b>', '').replace('</b>', '').replace('\r', '').split('\n'):
        if line.strip(): doc.add_paragraph(line.strip())
        
    doc.add_heading('填報資訊 / 年度執行亮點成果：', level=2)
    report_text_clean = re.sub(r'(^|\n)(\d+[\.\)]|[-•*])\s*\n\s*', r'\1\2 ', str(report_text))
    for line in report_text_clean.replace('\r', '').split('\n'):
        line = line.strip()
        if not line: continue
        p = doc.add_paragraph()
        if re.match(r'^【.*】$', line): p.add_run(line).bold = True
        else:
            match = re.match(r'^([0-9a-zA-Z]+[\.、\)]|[\(（][0-9a-zA-Z一二三四五六七八九十]+[\)）]|[一二三四五六七八九十]+、|[-•*])\s*(.*)', line)
            if match:
                p.paragraph_format.left_indent = Cm(0.8)
                p.paragraph_format.first_line_indent = Cm(-0.8)
                p.add_run(f"{match.group(1)} {match.group(2)}")
            else: p.add_run(line)
            
    doc.add_heading('佐證照片/檔案：', level=2)
    add_images_to_word_table(doc, unit_enriched_files)
    
    has_any_doc = any(not f.get('is_image') for files in unit_enriched_files.values() for f in files)
    if has_any_doc:
        doc.add_heading('其他參考資料：', level=2)
        for u, files in unit_enriched_files.items():
            docs = [f for f in files if not f.get('is_image')]
            if docs:
                if len(unit_enriched_files) > 1: doc.add_paragraph(f"{u}").runs[0].bold = True
                for f in docs:
                    doc.add_paragraph().add_run(f"{f['desc']} (此為 {'PDF 檔案' if f.get('is_pdf') else '其他文件'})").bold = True
                    doc.add_paragraph(f"請至雲端檢視：https://drive.google.com/file/d/{f['id']}/view")
    
    if not any(f.get('is_image') for files in unit_enriched_files.values() for f in files) and not has_any_doc:
        doc.add_paragraph("此紀錄無上傳任何單位附件。")
        
    # [新增] 寫入新聞亮點與照片至 Word 文件最下方
    if news_data:
        doc.add_heading('相關新聞亮點與佐證：', level=2)
        for idx, news in enumerate(news_data):
            p_title = doc.add_paragraph()
            p_title.add_run(f"【新聞 {idx+1}】 {news['title']}").bold = True
            
            p_sum = doc.add_paragraph()
            p_sum.add_run("亮點摘要：\n").bold = True
            p_sum.add_run(news['summary'])
            
            if news['link'] and news['link'] != 'nan':
                p_link = doc.add_paragraph()
                p_link.add_run("新聞連結：").bold = True
                p_link.add_run(news['link'])
                
            if news['enriched_files']:
                add_images_to_word_table(doc, {f"【新聞 {idx+1} 圖片】": news['enriched_files']})
    
    # 字體美化統一
    for p in doc.paragraphs:
        if not p.style.name.startswith('Heading'): p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for r in p.runs: set_run_font(r)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs: set_run_font(r)
                    
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# ==========================================
# 🚀 網頁主畫面與 UI 邏輯
# ==========================================
col_title, col_btn = st.columns([3, 1])
with col_title: st.title("📊 嘉大綠色大學填報管理區")
with col_btn:
    st.write("") 
    if st.button("🔄 同步最新雲端資料", use_container_width=True):
        load_gsheet_data.clear()
        get_file_info.clear() 
        st.rerun()

tab_missing, tab_download = st.tabs(["🎯 追蹤未申報單位", "📥 下載彙整資料"])

# ==========================================
# 🏷️ TAB 1: 追蹤未申報單位
# ==========================================
with tab_missing:
    with st.spinner("⏳ 正在比對資料庫，找出未申報名單..."):
        df_req = load_gsheet_data("評比題目表")
        df_sub = load_gsheet_data("填報資料庫")
        
        if df_req.empty:
            st.error("無法讀取評比題目表，請確認連線。")
        else:
            df_req_clean = df_req[['權責單位', '當年度題目', '中文標題']].copy()
            df_req_clean['權責單位'] = df_req_clean['權責單位'].astype(str).str.strip()
            df_req_clean['當年度題目'] = df_req_clean['當年度題目'].astype(str).str.strip()
            df_req_clean = df_req_clean[df_req_clean['權責單位'] != ''].drop_duplicates()
            
            df_sub_clean = pd.DataFrame(columns=['權責單位', '當年度題目'])
            if not df_sub.empty:
                df_sub_clean = df_sub[['權責單位', '題號']].copy()
                df_sub_clean.rename(columns={'題號': '當年度題目'}, inplace=True)
                df_sub_clean['權責單位'] = df_sub_clean['權責單位'].astype(str).str.strip()
                df_sub_clean['當年度題目'] = df_sub_clean['當年度題目'].astype(str).str.strip()
                df_sub_clean = df_sub_clean.drop_duplicates()
            
            merged = df_req_clean.merge(df_sub_clean, on=['權責單位', '當年度題目'], how='left', indicator=True)
            missing = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])
            
            st.markdown("<div class='morandi-dark-title'>🔍 尚未填報之單位與題目清單</div>", unsafe_allow_html=True)
            
            if missing.empty:
                st.success("🎉 太棒了！所有單位皆已完成所有題目的填報！")
            else:
                st.info(f"💡 目前共有 **{len(missing)}** 筆待填報項目。")
                grouped_missing = missing.groupby('權責單位')
                
                html_str = ""
                count = 0
                for unit, group in grouped_missing:
                    bg_summary = "#9CB4C4" if count % 2 == 0 else "#D0DCE3"
                    bg_content = "#E8EEF2" if count % 2 == 0 else "#F7F9FA"
                    border_color = bg_summary
                        
                    html_str += f'''
                    <details style="margin-bottom: 12px; border: 1px solid {border_color}; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                        <summary style="background-color: {bg_summary}; color: black; padding: 15px 20px; font-size: 1.25em; font-weight: bold; border-radius: 8px; cursor: pointer;">
                            🏢 {unit} <span style="color: #C0392B; font-size: 0.9em;">(尚缺 {len(group)} 題)</span>
                        </summary>
                        <div style="background-color: {bg_content}; padding: 20px; color: black; border-radius: 0 0 8px 8px; border-top: 1px solid {border_color};">
                            <ul style="margin: 0; padding-left: 25px; line-height: 1.8; font-size: 1.15em;">
                    '''
                    for idx, row in group.iterrows():
                        html_str += f"<li style='margin-bottom: 8px;'><b>{row['當年度題目']}</b>：{row['中文標題']}</li>"
                    html_str += "</ul></div></details>"
                    count += 1
                st.markdown(html_str, unsafe_allow_html=True)

# ==========================================
# 🏷️ TAB 2: 下載彙整資料 (單題 & 批次 ZIP)
# ==========================================
with tab_download:
    df_r = load_gsheet_data("填報資料庫")
    df_q = load_gsheet_data("評比題目表")
    df_news = load_gsheet_data("AI新聞資料庫")
    
    if df_r.empty: 
        st.info("💡 目前資料庫中尚未有任何填報紀錄可供下載。")
    else:
        # 底色區塊按鈕呈現
        st.markdown("<br>", unsafe_allow_html=True)
        dl_mode = st.radio("選擇下載模式", ["單題彙整下載", "全部題目彙整下載"], horizontal=True, label_visibility="collapsed")
        
        reported_q_ids = df_r['題號'].astype(str).str.strip().unique().tolist()
        reported_q_ids.sort()
        
        if dl_mode == "單題彙整下載":
            st.markdown("<div class='morandi-select-title'>📌 請選擇欲彙整下載之題目</div>", unsafe_allow_html=True)
            q_options = []
            for qid in reported_q_ids:
                title_match = df_q[df_q['當年度題目'].astype(str).str.strip() == qid]
                title = title_match.iloc[0]['中文標題'] if not title_match.empty else "未知標題"
                q_options.append(f"{qid} - {title}")
                
            q_options.sort()
            view_item = st.selectbox("", ["請選擇..."] + q_options, label_visibility="collapsed", key="m_item")
            
            if view_item != "請選擇...":
                v_q_id = view_item.split(" - ")[0]
                
                with st.spinner("⏳ 正在拉取資料與雲端圖片，請稍候..."):
                    v_q_title, v_desc_text, info_strings, combined_reqs, combined_texts, unit_enriched_files, news_data = extract_single_question_data(v_q_id, df_q, df_r, df_news)
                
                st.markdown("---")
                for info in info_strings:
                    st.markdown(f"<div style='color: #555555; font-size: 1.15em; margin-bottom: 5px; font-weight: bold;'>{info}</div>", unsafe_allow_html=True)
                
                st.markdown(f"<div class='morandi-question-title'>📖 {v_q_id}：{v_q_title}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: #B85042; font-weight: bold; font-size: 1.3em; padding-left: 5px; margin-bottom: 25px;'>{v_desc_text}</div>", unsafe_allow_html=True)
                
                display_reqs = combined_reqs.replace('\n', '<br>')
                display_reqs = re.sub(r'【(.*?)】', r'<span style="color:#154360; font-weight:bold;">【\1】</span>', display_reqs)
                st.markdown(f"""
                <div class='light-brown-box'>
                    <strong style='font-size: 1.4em; color: #B85042;'>🔍 資料需求：</strong><br>
                    <div style='font-size: 1.2em; margin-top: 10px; padding-left: 10px; color: #2C3E50;'>{display_reqs}</div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<div class='morandi-dark-title'>✍️ 填報資訊 / 年度執行亮點成果</div>", unsafe_allow_html=True)
                formatted_report = format_report_text_to_html(combined_texts)
                st.markdown(f"<div style='background-color: #F8FAFB; padding: 20px; border-radius: 8px; border: 1px solid #E2E7E3; font-size: 1.1em; color: #2C3E50; margin-bottom: 25px;'>{formatted_report}</div>", unsafe_allow_html=True)
                
                # 網頁版照片呈現 (單位)
                st.markdown("<div class='morandi-dark-title'>📎 佐證照片或檔案</div>", unsafe_allow_html=True)
                unit_html_table = generate_html_image_table(unit_enriched_files)
                if unit_html_table: st.markdown(unit_html_table, unsafe_allow_html=True)
                else: st.info("此紀錄無上傳任何單位附件。")
                
                # 網頁版新聞呈現
                if news_data:
                    st.markdown("<div class='morandi-dark-title'>📰 相關新聞亮點與佐證</div>", unsafe_allow_html=True)
                    for idx, news in enumerate(news_data):
                        st.markdown(f"<div style='font-size: 1.2em; font-weight: bold; color: #154360; margin-top:15px;'>【新聞 {idx+1}】 {news['title']}</div>", unsafe_allow_html=True)
                        st.markdown(f"**亮點摘要：** {news['summary']}")
                        if news['link'] and news['link'] != 'nan':
                            st.markdown(f"**新聞連結：** [{news['link']}]({news['link']})")
                        
                        if news['enriched_files']:
                            news_html_table = generate_html_image_table({f"【新聞 {idx+1} 圖片】": news['enriched_files']})
                            st.markdown(news_html_table, unsafe_allow_html=True)
                        st.markdown("<hr style='margin: 10px 0; border-top: 1px dashed #BDC3C7;'>", unsafe_allow_html=True)
                
                st.markdown("<div style='text-align: center; margin-top: 30px; margin-bottom: 10px;'><b>準備好彙整這份報告了嗎？</b></div>", unsafe_allow_html=True)
                
                with st.spinner("⏳ 正在為您產生 Word 報告檔案..."):
                    docx_bytes = build_word_document(v_q_id, v_q_title, info_strings, v_desc_text, combined_reqs, combined_texts, unit_enriched_files, news_data)
                
                col_dl1, col_dl2, col_dl3 = st.columns([3, 4, 3])
                with col_dl2:
                    st.download_button(
                        label="📥 下載彙整成果 (Word檔)",
                        data=docx_bytes,
                        file_name=f"{v_q_id}_填報彙整成果.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
        elif dl_mode == "全部題目彙整下載":
            st.markdown("<div class='morandi-select-title'>📦 批量打包下載所有題目報告</div>", unsafe_allow_html=True)
            st.info("💡 此功能將自動為「所有已填報之題目」分別產生獨立的 Word 報告（含對應新聞與照片佐證），並打包成一個 ZIP 壓縮檔，方便您一次性下載留存。")
            
            if st.button("🚀 一鍵打包產生 ZIP 壓縮檔", type="primary"):
                with st.spinner("⏳ 正在背景逐題解析資料、抓取圖片並產生 Word 報告，這需要幾分鐘的時間，請耐心等候..."):
                    
                    zip_buffer = io.BytesIO()
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        total_q = len(reported_q_ids)
                        for i, qid in enumerate(reported_q_ids):
                            status_text.text(f"正在處理第 {i+1}/{total_q} 題：{qid} ...")
                            
                            # 取資料並產製 Word (自動包含該題的新聞資訊)
                            v_q_title, v_desc_text, info_strings, combined_reqs, combined_texts, unit_enriched_files, news_data = extract_single_question_data(qid, df_q, df_r, df_news)
                            docx_bytes = build_word_document(qid, v_q_title, info_strings, v_desc_text, combined_reqs, combined_texts, unit_enriched_files, news_data)
                            
                            # 寫入 ZIP
                            safe_title = re.sub(r'[\\/*?:"<>|]', "", v_q_title) 
                            file_name = f"{qid}_{safe_title}_填報彙整.docx"
                            zip_file.writestr(file_name, docx_bytes.getvalue())
                            
                            progress_bar.progress((i + 1) / total_q)
                            
                    zip_buffer.seek(0)
                    status_text.empty()
                    st.success("✅ 全部打包完成！請點擊下方按鈕下載。")
                    
                    st.download_button(
                        label="📥 下載全題彙整成果 (ZIP壓縮檔)",
                        data=zip_buffer,
                        file_name=f"嘉大綠色大學_全題彙整_{datetime.datetime.now().strftime('%Y%m%d')}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )