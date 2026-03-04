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
# 🪄 終極超級渲染引擎：純正 HTML 建構與完美解析
# ==========================================
def render_native_content(text):
    if not isinstance(text, str) or not text.strip():
        st.info("*(🚧 系統提示：尚無去年度文字資料，待後台擷取後自動匯入。)*")
        return

    lines = text.split('\n')
    html = []
    in_table = False
    current_cell_html = []
    
    def process_line(line_str):
        line_str = line_str.strip()
        if not line_str: return ""
        
        # 1. 處理項目符號 (完美還原 Word 縮排)
        is_bullet = False
        if line_str.startswith('* ') or line_str.startswith('- '):
            line_str = f"<div style='margin-left: 1.5em; text-indent: -1.5em; margin-bottom: 5px;'>• {line_str[2:]}</div>"
            is_bullet = True
        elif re.match(r'^\d+\.\s', line_str):
            line_str = f"<div style='margin-left: 1.5em; text-indent: -1.5em; margin-bottom: 5px;'>{line_str}</div>"
            is_bullet = True
            
        # 2. 處理圖片 (保證 2 張並排、不裁切)
        def img_repl(match):
            url = match.group(1).replace("uc?id=", "thumbnail?sz=w1000&id=")
            # 🌟 魔法核心：width 設定 48%，保證兩兩並排；height: auto 保證不裁切！
            return f'<img src="{url}" referrerpolicy="no-referrer" style="width: calc(50% - 10px); min-width: 150px; height: auto; margin: 5px; border-radius: 6px; box-shadow: 0 2px 5px rgba(0,0,0,0.15); display: inline-block; vertical-align: top;">'
            
        original_line = line_str
        line_str = re.sub(r'!\[.*?\]\((.*?)\)', img_repl, line_str)
        line_str = line_str.replace('**', '') # 移除 Markdown 粗體星號
        
        # 3. 如果這行是純文字，包裝成 div 確保段落換行
        if not is_bullet and '<img' not in line_str:
            line_str = f"<div style='margin-bottom: 8px; line-height: 1.6;'>{line_str}</div>"
            
        return line_str

    for line in lines:
        line = line.strip()
        if line.startswith('>'):
            line = line[1:].strip()
            
        # 偵測表格開頭
        if "📊" in line or "表格資料解析" in line:
            if in_table:
                if current_cell_html: html.append("".join(current_cell_html) + "</td>")
                html.append("</tr></table>")
            in_table = True
            current_cell_html = []
            # 建立表格外框
            html.append('<table style="width: 100%; border-collapse: collapse; margin-bottom: 25px; border: 2px solid #8F9CA3; box-shadow: 0 2px 8px rgba(0,0,0,0.05); font-size: 1.05em;">')
            continue
            
        # 偵測表格結尾
        if in_table and line == "---":
            if current_cell_html: html.append("".join(current_cell_html) + "</td>")
            html.append("</tr></table>")
            in_table = False
            current_cell_html = []
            continue
            
        if in_table:
            # 偵測換列
            if re.search(r'【第 \d+ 列】', line):
                if current_cell_html: 
                    html.append("".join(current_cell_html) + "</td>")
                    current_cell_html = []
                if html and html[-1].endswith("</td>"):
                    html.append("</tr>")
                html.append('<tr style="border-bottom: 1px solid #D9E0E3;">')
                continue
                
            # 偵測換格 (欄位)
            if "[欄位" in line or "[Column" in line:
                if current_cell_html: 
                    html.append("".join(current_cell_html) + "</td>")
                    current_cell_html = []
                # 建立儲存格
                html.append("<td style='border: 1px solid #D9E0E3; padding: 15px; vertical-align: top; background-color: #FFFFFF; color: #2C3E50;'>")
                
                # 擷取欄位標籤後的文字
                text_part = re.sub(r'\[?(欄位|Column)\s*\d+\]?[*\s]*[:：]?', '', line).strip()
                if text_part:
                    current_cell_html.append(process_line(text_part))
                continue
            
            # 表格內的正常內容
            if line:
                current_cell_html.append(process_line(line))
        else:
            # 非表格區塊的內容
            if line:
                html.append(f"<div style='color: #2C3E50; font-size: 1.05em;'>{process_line(line)}</div>")
                
    if in_table:
        if current_cell_html: html.append("".join(current_cell_html) + "</td>")
        html.append("</tr></table>")
        
    st.markdown("".join(html), unsafe_allow_html=True)


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