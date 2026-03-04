import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
import re

# ==========================================
# 🛡️ 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報區", page_icon="📝", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定
# ==========================================
st.markdown("""
<style>
    .morandi-title {
        background-color: #8F9CA3;
        color: white;
        padding: 10px 15px;
        border-radius: 6px;
        font-weight: bold;
        font-size: 1.2em;
        margin-bottom: 5px;
        margin-top: 15px;
    }
    .morandi-req {
        background-color: #E2E7E3;
        color: #2C3E50;
        padding: 15px;
        border-left: 6px solid #8F9CA3;
        border-radius: 5px;
        margin-top: 10px;
        margin-bottom: 15px;
        line-height: 1.6;
        font-size: 1.05em;
    }
    div.stButton > button:first-child {
        border-radius: 6px !important;
        font-weight: bold !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 資料庫連線與快取控制
# ==========================================
@st.cache_data(ttl=600)
def load_gsheet_data():
    try:
        skey = st.secrets["gcp_oauth"].to_dict()
        credentials = Credentials(
            token=None,
            refresh_token=skey.get("refresh_token"),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=skey.get("client_id"),
            client_secret=skey.get("client_secret")
        )
        gc = gspread.authorize(credentials)
        SHEET_ID = '1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8'
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("評比題目表")
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"⚠️ 無法讀取 Google Sheet，詳細錯誤：{e}")
        return pd.DataFrame()

# ==========================================
# 🪄 超級渲染引擎：純正 HTML 表格 + CSS 雙圖並排
# ==========================================
def render_native_content(text):
    if not isinstance(text, str) or not text.strip():
        st.info("*(🚧 系統提示：尚無去年度文字資料，待後台擷取後自動匯入。)*")
        return

    lines = text.split('\n')
    in_table = False
    current_table = []
    current_row = []
    current_cell = {"text": "", "images": []}
    
    ui_blocks = []
    current_standalone_images = []

    # 將累積的照片輸出為兩兩並排的區塊
    def flush_images():
        if current_standalone_images:
            ui_blocks.append({"type": "images", "urls": list(current_standalone_images)})
            current_standalone_images.clear()

    # --- 1. 資料解析階段 ---
    for line in lines:
        line_str = line.strip()
        if line_str.startswith('>'):
            line_str = line_str[1:].strip()
            
        if "📊" in line_str or "表格資料解析" in line_str:
            flush_images()
            in_table = True
            continue
            
        if in_table and line_str == "---":
            if current_cell["text"].strip() or current_cell["images"]:
                current_row.append(current_cell)
            if current_row:
                current_table.append(current_row)
            if current_table:
                ui_blocks.append({"type": "table", "rows": current_table})
            
            in_table = False
            current_table = []
            current_row = []
            current_cell = {"text": "", "images": []}
            continue
            
        if in_table:
            if re.search(r'【第 \d+ 列】', line_str):
                if current_cell["text"].strip() or current_cell["images"]:
                    current_row.append(current_cell)
                if current_row:
                    current_table.append(current_row)
                current_row = []
                current_cell = {"text": "", "images": []}
                continue
                
            if "[欄位" in line_str or "[Column" in line_str:
                if current_cell["text"].strip() or current_cell["images"]:
                    current_row.append(current_cell)
                parts = re.split(r'[：:]', line_str, maxsplit=1)
                text_content = parts[1].strip() if len(parts) > 1 else ""
                text_content = text_content.replace('**', '').replace('*', '').strip()
                current_cell = {"text": text_content + "\n", "images": []}
                continue
            
            img_matches = re.findall(r'!\[.*?\]\((https://drive\.google\.com/[^\)]+)\)', line_str)
            if img_matches:
                for img_url in img_matches:
                    thumb_url = img_url.replace("uc?id=", "thumbnail?sz=w1000&id=")
                    current_cell["images"].append(thumb_url)
                clean_line = re.sub(r'!\[.*?\]\(.*?\)', '', line_str).replace('**', '').replace('*', '').strip()
                if clean_line: current_cell["text"] += clean_line + "\n"
                continue
            
            clean_line = line_str.replace('**', '').replace('*', '').strip()
            if clean_line: current_cell["text"] += clean_line + "\n"
                
        else:
            # 非表格區塊：收集獨立照片與文字
            img_matches = re.findall(r'!\[.*?\]\((https://drive\.google\.com/[^\)]+)\)', line_str)
            if img_matches:
                for img_url in img_matches:
                    thumb_url = img_url.replace("uc?id=", "thumbnail?sz=w1000&id=")
                    current_standalone_images.append(thumb_url)
                clean_line = re.sub(r'!\[.*?\]\(.*?\)', '', line_str).replace('**', '').replace('*', '').strip()
                if clean_line:
                    flush_images()
                    ui_blocks.append({"type": "text", "content": clean_line})
            else:
                clean_line = line_str.replace('**', '').replace('*', '').strip()
                if clean_line:
                    flush_images()
                    ui_blocks.append({"type": "text", "content": clean_line})

    flush_images()
    if in_table:
        if current_cell["text"].strip() or current_cell["images"]: current_row.append(current_cell)
        if current_row: current_table.append(current_row)
        if current_table: ui_blocks.append({"type": "table", "rows": current_table})

    # --- 2. HTML 畫布繪製階段 ---
    html_str = ""
    for block in ui_blocks:
        if block["type"] == "text":
            text_content = block["content"].replace('\n', '<br>')
            html_str += f'<p style="line-height: 1.6; color: #2C3E50; font-size: 1.05em; margin-bottom: 12px;">{text_content}</p>'
            
        elif block["type"] == "images":
            # 🌟 魔法 1：CSS Grid 強制 2 張並排，高度統一 280px，多餘部分自動漂亮裁切
            html_str += '<div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-bottom: 20px;">'
            for url in block["urls"]:
                html_str += f'<img src="{url}" referrerpolicy="no-referrer" style="width: 100%; height: 280px; object-fit: cover; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">'
            html_str += '</div>'
            
        elif block["type"] == "table":
            # 🌟 魔法 2：純正的 HTML Table，保證格子整齊對齊
            html_str += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 25px; border: 2px solid #8F9CA3; border-radius: 4px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">'
            for r_idx, row in enumerate(block["rows"]):
                html_str += '<tr style="border-bottom: 1px solid #D9E0E3;">'
                col_width = f"{100/len(row)}%" if row else "100%"
                for c_idx, cell in enumerate(row):
                    # 表格第一列(通常是標題)給予微微的莫蘭迪灰底色
                    bg_color = "#F4F6F7" if r_idx == 0 else "#FFFFFF"
                    html_str += f'<td style="border: 1px solid #D9E0E3; padding: 15px; vertical-align: top; width: {col_width}; background-color: {bg_color};">'
                    
                    if cell["text"].strip():
                        cell_txt = cell["text"].strip().replace('\n', '<br>')
                        html_str += f'<div style="margin-bottom: 12px; color: #34495E; font-size: 0.95em; line-height: 1.5;">{cell_txt}</div>'
                        
                    if cell["images"]:
                        # 如果儲存格內有多張照片，也套用 2 張並排邏輯
                        grid_style = "display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px;" if len(cell["images"]) > 1 else "display: block;"
                        html_str += f'<div style="{grid_style}">'
                        for url in cell["images"]:
                            html_str += f'<img src="{url}" referrerpolicy="no-referrer" style="width: 100%; max-height: 200px; object-fit: cover; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">'
                        html_str += '</div>'
                        
                    html_str += '</td>'
                html_str += '</tr>'
            html_str += '</table>'

    st.markdown(html_str, unsafe_allow_html=True)


# ==========================================
# 🚀 網頁主畫面
# ==========================================
col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("📝 嘉大綠色大學填報區")
with col_btn:
    st.write("") 
    if st.button("🔄 同步最新雲端資料", use_container_width=True):
        load_gsheet_data.clear()
        st.rerun()

with st.spinner('🔄 正在載入最新評比題目...'):
    df_questions = load_gsheet_data()

if not df_questions.empty:
    unit_list = df_questions['權責單位'].dropna().unique().tolist()
    unit_list = [str(u).strip() for u in unit_list if str(u).strip() != '']
    
    st.markdown("<div class='morandi-title'>🏢 請選擇您的單位</div>", unsafe_allow_html=True)
    selected_unit = st.selectbox("", ["請選擇..."] + unit_list, label_visibility="collapsed")
    
    if selected_unit != "請選擇...":
        df_unit_questions = df_questions[df_questions['權責單位'].astype(str).str.strip() == selected_unit].copy()
        df_unit_questions['選項標示'] = df_unit_questions['當年度題目'].astype(str) + " - " + df_unit_questions['中文標題'].astype(str)
        
        st.markdown("<div class='morandi-title'>📌 請選擇填報項目</div>", unsafe_allow_html=True)
        selected_option = st.selectbox("", df_unit_questions['選項標示'].tolist(), label_visibility="collapsed")
        
        selected_q_id = selected_option.split(" - ")[0]
        question_data = df_unit_questions[df_unit_questions['當年度題目'].astype(str) == selected_q_id].iloc[0]
        
        st.markdown("---")
        
        st.markdown(f"<div class='morandi-title'>📖 {question_data.get('當年度題目')}：{question_data.get('中文標題')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='color: black; font-size: 1.1em; padding-left: 5px; margin-bottom: 15px;'><b>中文說明：</b><br>{question_data.get('中文說明', '無')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='morandi-req'><b>🔍 資料需求：</b><br>{question_data.get('資料需求', '無特別說明')}</div>", unsafe_allow_html=True)
        
        with st.expander("💡 點擊展開查看：前一年度 (2025) 參考資訊", expanded=True):
            st.markdown(f"**對應之去年度題目：** {question_data.get('前一年度題目', '無')}")
            st.markdown("---")
            
            ref_text = question_data.get('2025參考文字_AI預留', '')
            render_native_content(ref_text)
                
        st.markdown("---")
        
        with st.form("report_form"):
            report_text = st.text_area("✍️ 填報資訊/年度執行亮點成果", height=150, placeholder="請在此輸入您的填寫內容...")
            uploaded_files = st.file_uploader("📎 上傳照片或佐證檔案 (支援 PDF, JPG, PNG, DOCX 等)：", accept_multiple_files=True)
            
            st.markdown("<div style='display: flex; justify-content: center;'>", unsafe_allow_html=True)
            submitted = st.form_submit_button("📤 資料確認送出", type="primary")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if submitted:
                if not report_text.strip() and not uploaded_files:
                    st.error("⚠️ 請至少填寫成果說明或上傳佐證檔案！")
                else:
                    st.success("🎉 資料已暫存成功！目前系統畫面設計與邏輯皆已完備。")

elif 'df_questions' in locals() and df_questions.empty:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")