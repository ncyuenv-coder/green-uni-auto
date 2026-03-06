import streamlit as st
import pandas as pd
import gspread
import io
import re
import datetime
import base64
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
# 🌟 全域參數設定 & 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報管理區", page_icon="📊", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定 (包含斑馬紋交錯底色與黑字設定)
# ==========================================
st.markdown("""
<style>
    /* 標籤頁 (Tabs) 樣式 */
    button[data-baseweb="tab"] { 
        background-color: #E6F0F9 !important; 
        border-radius: 8px 8px 0px 0px !important; 
        margin-right: 4px !important; 
        padding: 10px 20px !important; 
        border: 1px solid #AED6F1 !important; 
        border-bottom: none !important; 
        transition: all 0.3s ease; 
    }
    button[data-baseweb="tab"] p { 
        font-size: 1.4em !important; 
        font-weight: bold !important; 
        color: #2C3E50 !important; 
    }
    button[data-baseweb="tab"][aria-selected="true"] { 
        background-color: #154360 !important; 
        border: 1px solid #0B2331 !important; 
        border-bottom: 3px solid #0B2331 !important; 
    }
    button[data-baseweb="tab"][aria-selected="true"] p { 
        color: #FFFFFF !important; 
    }

    /* 各區塊標題 */
    .morandi-select-title { background-color: #9DAB86; color: white; padding: 12px 15px; border-radius: 6px; font-weight: bold; font-size: 1.2em; margin-bottom: 5px; margin-top: 15px; }
    .morandi-question-title { background-color: #948B89; color: white; padding: 15px 18px; border-radius: 6px; font-weight: bold; font-size: 1.4em; margin-bottom: 15px; margin-top: 10px; }
    .morandi-dark-title { background-color: #5C6B73; color: white; padding: 12px 20px; border-radius: 6px; font-weight: bold; font-size: 1.3em; margin-bottom: 10px; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .light-brown-box { background-color: #F5EBE0; color: #2C3E50; padding: 15px 20px; border-left: 6px solid #D4A373; border-radius: 6px; margin-bottom: 5px; line-height: 1.6; }
    
    /* 🌟 TAB 1：折疊面板標題區 (交錯底色) */
    div[data-testid="stExpander"]:nth-of-type(odd) details summary { background-color: #738A96 !important; color: white !important; border-radius: 6px; }
    div[data-testid="stExpander"]:nth-of-type(even) details summary { background-color: #8FAAB8 !important; color: white !important; border-radius: 6px; }
    div[data-testid="stExpander"] details summary p { font-size: 1.2em !important; font-weight: bold; color: white !important; }
    div[data-testid="stExpander"] details summary svg { fill: white !important; }
    
    /* 🌟 TAB 1：展開內容區 (交錯底色 + 純黑字體) */
    div[data-testid="stExpander"]:nth-of-type(odd) details[open] > div:nth-child(2) { 
        background-color: #E4EBEE !important; /* 較深標題的淡化版 */
        border: 2px solid #738A96; border-top: none; border-radius: 0 0 6px 6px; padding: 15px; 
    }
    div[data-testid="stExpander"]:nth-of-type(even) details[open] > div:nth-child(2) { 
        background-color: #F2F6F8 !important; /* 較淺標題的淡化版 */
        border: 2px solid #8FAAB8; border-top: none; border-radius: 0 0 6px 6px; padding: 15px; 
    }
    div[data-testid="stExpander"] details[open] > div:nth-child(2) p,
    div[data-testid="stExpander"] details[open] > div:nth-child(2) li { 
        font-size: 1.15em !important; line-height: 1.6; 
        color: #000000 !important; /* 強制設定為黑色 */
        font-weight: 500;
    }
    
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

# ==========================================
# 🌟 智慧判斷檔案格式、照片方向與「全景圖偵測」
# ==========================================
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
# 🌟 公文級文字解析引擎：網頁完美縮排對齊 (支援跨單位綜整)
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
            marker = match.group(1)
            content = match.group(2)
            html += f"<div style='padding-left: 2.2em; text-indent: -2.2em; margin-bottom: 5px; line-height: 1.6;'>{marker} {content}</div>"
        else: 
            html += f"<div style='margin-bottom: 5px; line-height: 1.6;'>{line}</div>"
    return html.replace('\n', ' ')

# ==========================================
# 🌟 Word 報告產製引擎 (垂直置中 + 公文排版 + 跨單位合併表格)
# ==========================================
def set_run_font(run):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def generate_word_report(q_id, q_title, info_strings, desc_text, req_text, report_text, unit_enriched_files):
    doc = Document()
    doc.add_heading(f'嘉大綠色大學填報成果彙整', 0)
    
    for info in info_strings:
        p_info = doc.add_paragraph(info)
        p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
    doc.add_heading(f'{q_id} {q_title}', level=1)
    
    doc.add_heading('題目說明：', level=2)
    for line in str(desc_text).replace('\r', '').split('\n'):
        if line.strip(): doc.add_paragraph(line.strip())
    
    doc.add_heading('資料需求：', level=2)
    clean_req = str(req_text).replace('<br>', '\n').replace('<b>', '').replace('</b>', '')
    for line in clean_req.replace('\r', '').split('\n'):
        if line.strip(): doc.add_paragraph(line.strip())
    
    doc.add_heading('填報資訊 / 年度執行亮點成果：', level=2)
    report_text_clean = re.sub(r'(^|\n)(\d+[\.\)]|[-•*])\s*\n\s*', r'\1\2 ', str(report_text))
    
    for line in report_text_clean.replace('\r', '').split('\n'):
        line = line.strip()
        if not line: continue
            
        p = doc.add_paragraph()
        if re.match(r'^【.*】$', line):
            p.add_run(line).bold = True
            continue
            
        match = re.match(r'^([0-9a-zA-Z]+[\.、\)]|[\(（][0-9a-zA-Z一二三四五六七八九十]+[\)）]|[一二三四五六七八九十]+、|[-•*])\s*(.*)', line)
        if match:
            marker = match.group(1)
            content = match.group(2)
            p.paragraph_format.left_indent = Cm(0.8)
            p.paragraph_format.first_line_indent = Cm(-0.8)
            p.add_run(f"{marker} {content}")
        else:
            p.add_run(line)
    
    doc.add_heading('佐證照片/檔案：', level=2)
    
    has_any_image = any(f.get('is_image') for files in unit_enriched_files.values() for f in files)
    
    if has_any_image:
        table = doc.add_table(rows=0, cols=2)
        table.style = 'Table Grid'
        
        for u, files in unit_enriched_files.items():
            images = [f for f in files if f.get('is_image')]
            if not images: continue
            
            if len(unit_enriched_files) > 1:
                row = table.add_row()
                cell = row.cells[0]
                cell.merge(row.cells[1])
                # 🌟 加入垂直置中設定
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(f"【{u}】").bold = True
            
            panoramas = [f for f in images if f.get('is_panorama')]
            landscapes = [f for f in images if f.get('is_landscape') and not f.get('is_panorama')]
            portraits = [f for f in images if not f.get('is_landscape')]
            
            # 1. 處理全景圖
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
            
            # 2. 處理配對照片
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
                    # 🌟 將每一格都設定為垂直置中
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
    
    # 處理文件
    has_any_doc = any(not f.get('is_image') for files in unit_enriched_files.values() for f in files)
    
    if has_any_doc:
        doc.add_heading('其他參考資料：', level=2)
        for u, files in unit_enriched_files.items():
            docs = [f for f in files if not f.get('is_image')]
            if docs:
                if len(unit_enriched_files) > 1:
                    doc.add_paragraph(f"【{u}】").runs[0].bold = True
                for f in docs:
                    p1 = doc.add_paragraph()
                    ftype = "PDF 檔案" if f.get('is_pdf') else "其他文件"
                    p1.add_run(f"{f['desc']} (此為 {ftype})").bold = True
                    p2 = doc.add_paragraph()
                    p2.add_run(f"請至雲端檢視：https://drive.google.com/file/d/{f['id']}/view")
                    
    if not has_any_image and not has_any_doc:
        doc.add_paragraph("此紀錄無上傳任何附件。")
    
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
                for unit, group in grouped_missing:
                    with st.expander(f"🏢 {unit} (尚缺 {len(group)} 題)", expanded=False):
                        for idx, row in group.iterrows():
                            st.markdown(f"- **{row['當年度題目']}**：{row['中文標題']}")

# ==========================================
# 🏷️ TAB 2: 下載彙整資料 (跨單位完美合併版)
# ==========================================
with tab_download:
    df_r = load_gsheet_data("填報資料庫")
    df_q = load_gsheet_data("評比題目表")
    
    if df_r.empty: 
        st.info("💡 目前資料庫中尚未有任何填報紀錄可供下載。")
    else:
        st.markdown("<div class='morandi-select-title'>📌 請選擇欲彙整下載之題目</div>", unsafe_allow_html=True)
        
        reported_q_ids = df_r['題號'].astype(str).str.strip().unique().tolist()
        
        q_options = []
        for qid in reported_q_ids:
            title_match = df_q[df_q['當年度題目'].astype(str).str.strip() == qid]
            title = title_match.iloc[0]['中文標題'] if not title_match.empty else "未知標題"
            q_options.append(f"{qid} - {title}")
            
        q_options.sort()
        view_item = st.selectbox("", ["請選擇..."] + q_options, label_visibility="collapsed", key="m_item")
        
        if view_item != "請選擇...":
            v_q_id = view_item.split(" - ")[0]
            v_q_title = view_item.split(" - ")[1]
            
            df_q_records = df_r[df_r['題號'].astype(str).str.strip() == v_q_id].copy()
            df_q_records['填報時間'] = pd.to_datetime(df_q_records['填報時間'])
            df_q_records = df_q_records.sort_values('填報時間').drop_duplicates(subset=['權責單位'], keep='last')
            
            unit_count = len(df_q_records)
            
            v_desc_text = "無"
            df_q_clean = df_q.copy()
            df_q_clean['當年度題目'] = df_q_clean['當年度題目'].astype(str).str.strip()
            df_q_clean['權責單位'] = df_q_clean['權責單位'].astype(str).str.strip()
            
            q_match_base = df_q_clean[df_q_clean['當年度題目'] == v_q_id]
            if not q_match_base.empty:
                v_desc_text = str(q_match_base.iloc[0].get('中文說明', '無')).replace('中文說明：', '').strip()
            
            info_strings = []
            combined_reqs = ""
            combined_texts = ""
            unit_files_dict = {}
            
            for idx, row in df_q_records.iterrows():
                u = str(row['權責單位']).strip()
                rep = str(row.get('填報人', '無')).strip()
                ext = str(row.get('填報人分機', '無')).strip()
                em = str(row.get('填報人電子郵件', '無')).strip()
                
                info_strings.append(f"填報單位：{u}  |  👤 填報人：{rep}  |  📞 分機：{ext}  |  📧 Email：{em}")
                
                q_match_unit = df_q_clean[(df_q_clean['當年度題目'] == v_q_id) & (df_q_clean['權責單位'] == u)]
                if q_match_unit.empty:
                    q_match_unit = q_match_base
                req_text = str(q_match_unit.iloc[0].get('資料需求', '無特別說明')) if not q_match_unit.empty else "無"
                rep_text = str(row.get('填報內容', '無')).strip()
                
                if unit_count > 1:
                    combined_reqs += f"【{u}】\n{req_text}\n\n"
                    combined_texts += f"【{u}】\n{rep_text}\n\n"
                else:
                    combined_reqs = req_text
                    combined_texts = rep_text
                
                files_for_u = []
                for i in range(1, 11):
                    fid = str(row.get(f'檔案{i}_ID', '')).strip()
                    desc = str(row.get(f'檔案{i}_說明', '')).strip()
                    if fid: files_for_u.append({'id': fid, 'desc': desc})
                unit_files_dict[u] = files_for_u
            
            combined_reqs = combined_reqs.strip()
            combined_texts = combined_texts.strip()

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
            
            # ==========================================
            # 🌟 網頁照片/檔案區 (垂直置中排版)
            # ==========================================
            st.markdown("<div class='morandi-dark-title'>📎 佐證照片或檔案</div>", unsafe_allow_html=True)
            
            unit_enriched_files = {}
            has_any_files = False
            for u, files in unit_files_dict.items():
                enriched = []
                for f in files:
                    info = get_file_info(f['id'], f['desc'])
                    if info: enriched.append(info)
                unit_enriched_files[u] = enriched
                if enriched: has_any_files = True
            
            if not has_any_files:
                st.info("此紀錄無上傳任何附件。")
            else:
                table_html = "<table style='width:100%; table-layout:fixed; border-collapse:collapse; border:1px solid #D9E0E3; background-color:white; margin-bottom: 25px;'>"
                has_image_table = False
                
                for u, files in unit_enriched_files.items():
                    images = [f for f in files if f.get('is_image')]
                    if not images: continue
                    has_image_table = True
                    
                    if unit_count > 1:
                        table_html += f"<tr><td colspan='2' style='border:1px solid #D9E0E3; padding:10px; text-align:center; background-color:#E6F0F9;'><strong style='font-size:1.2em; color:#154360;'>【{u}】</strong></td></tr>"
                    
                    panoramas = [f for f in images if f.get('is_panorama')]
                    landscapes = [f for f in images if f.get('is_landscape') and not f.get('is_panorama')]
                    portraits = [f for f in images if not f.get('is_landscape')]
                    
                    for pano in panoramas:
                        # 🌟 全景圖垂直置中
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
                                # 🌟 配對圖片垂直置中 (vertical-align: middle)
                                table_html += f"<td style='border:1px solid #D9E0E3; padding:15px; text-align:center; width:50%; vertical-align:middle;'><img src='data:{f['mime_type']};base64,{f['b64']}' style='aspect-ratio:{ratio}; width:100%; object-fit:contain; background-color:#f1f1f1; border-radius:8px; margin-bottom:10px;'><br><b>{f['desc']}</b></td>"
                            else:
                                table_html += "<td style='border:1px solid #D9E0E3; width:50%;'></td>"
                        table_html += "</tr>"
                
                table_html += "</table>"
                
                if has_image_table:
                    st.markdown(table_html, unsafe_allow_html=True)
                
                for u, files in unit_enriched_files.items():
                    docs = [f for f in files if not f.get('is_image')]
                    if docs:
                        if unit_count > 1:
                            st.markdown(f"<div style='font-size: 1.15em; font-weight: bold; color: #154360; margin-bottom: 10px;'>【{u}】的參考資料：</div>", unsafe_allow_html=True)
                        for f in docs:
                            drive_link = f"https://drive.google.com/file/d/{f['id']}/view"
                            icon, ftype = ("📄", "PDF 檔案") if f.get('is_pdf') else ("📁", "其他文件")
                            st.markdown(f"""
                            <div style='background-color: #f1f1f1; border-radius: 8px; border: 1px solid #ccc; padding: 25px; text-align: center; margin-bottom: 20px;'>
                                <span style='font-size:1.5em;'>{icon}</span> <span style='font-size:1.1em; font-weight:bold;'>📌 {f['desc']} ({ftype})</span><br><br>
                                <a href='{drive_link}' target='_blank' style='background-color: #5C6B73; color: white; padding: 10px 15px; border-radius: 6px; text-decoration: none; font-weight: bold;'>🔗 點擊前往雲端檢視</a>
                            </div>
                            """, unsafe_allow_html=True)
            
            st.markdown("---")
            st.markdown("<div style='text-align: center; margin-bottom: 10px;'><b>準備好彙整這份跨單位總報告了嗎？</b></div>", unsafe_allow_html=True)
            
            with st.spinner("⏳ 正在為您產生 Word 報告檔案..."):
                docx_bytes = generate_word_report(
                    v_q_id, v_q_title, info_strings, 
                    v_desc_text, combined_reqs, combined_texts, unit_enriched_files
                )
            
            col_dl1, col_dl2, col_dl3 = st.columns([3, 4, 3])
            with col_dl2:
                st.download_button(
                    label="📥 下載彙整成果 (Word檔)",
                    data=docx_bytes,
                    file_name=f"{v_q_id}_填報彙整成果.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )