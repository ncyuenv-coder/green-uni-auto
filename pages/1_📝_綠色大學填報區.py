import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
import re

# ==========================================
# 🪄 魔法轉換器：破解防盜鏈 + 統一照片排版 (兩張一列、固定大小)
# ==========================================
def fix_drive_images(text):
    if not isinstance(text, str):
        return ""
    
    # 1. 將傳統的 uc?id 換成隱藏版的 thumbnail API (設定寬度 1000px 確保高畫質)
    text = text.replace("https://drive.google.com/uc?id=", "https://drive.google.com/thumbnail?sz=w1000&id=")
    
    # 2. 移除相鄰圖片之間的「換行符號」，讓它們能在同一行並排
    text = re.sub(r'(\!\[.*?\]\(.*?\))\s*\n+\s*(?=\!\[.*?\]\(.*?\))', r'\1 ', text)
    
    # 3. 轉為 HTML，加入強大的 CSS 樣式
    # width: 48% (讓兩張圖剛好佔滿一列)
    # height: 250px (固定高度) 
    # object-fit: cover (保證圖片比例不變形，多餘部分自動裁切)
    # display: inline-block (讓圖片並排顯示)
    pattern = r'!\[.*?\]\((https://drive\.google\.com/thumbnail\?sz=w1000&id=[a-zA-Z0-9_-]+)\)'
    replacement = r'<img src="\1" referrerpolicy="no-referrer" style="width: 48%; height: 250px; object-fit: cover; border-radius: 8px; margin: 5px 1%; box-shadow: 0 4px 6px rgba(0,0,0,0.1); display: inline-block;">'
    html_text = re.sub(pattern, replacement, text)
    
    return html_text

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
    div.stButton {
        display: flex;
        justify-content: center;
        margin-top: 20px;
        margin-bottom: 40px;
    }
    div.stButton > button:first-child {
        background-color: #D4A373 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 30px !important;
        font-size: 18px !important;
        font-weight: bold !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #C08D5D !important;
        box-shadow: 0 6px 8px rgba(0,0,0,0.15) !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 資料庫連線設定
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
# 🚀 網頁主畫面
# ==========================================
st.title("📝 嘉大綠色大學填報區")

with st.spinner('🔄 正在載入最新評比題目...'):
    df_questions = load_gsheet_data()

if not df_questions.empty:
    required_columns = ['當年度題目', '英文標題', '中文標題', '中文說明', '資料需求', '權責單位', '前一年度題目']
    missing_columns = [col for col in required_columns if col not in df_questions.columns]
    
    if missing_columns:
        st.error(f"⚠️ Google Sheet 找不到以下欄位：{', '.join(missing_columns)}")
        st.stop()
        
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
        
        # ✨ 優化後的參考資訊區塊
        with st.expander("💡 點擊展開查看：前一年度 (2025) 參考資訊"):
            st.write(f"**對應之去年度題目：** {question_data.get('前一年度題目', '無')}")
            
            ref_text = question_data.get('2025參考文字_AI預留', '')
            st.markdown("---")
            
            if pd.notna(ref_text) and str(ref_text).strip() != "":
                st.markdown("**📝 去年度填報內容 (圖文整合)：**")
                fixed_content = fix_drive_images(str(ref_text))
                st.markdown(fixed_content, unsafe_allow_html=True)
            else:
                st.info("*(🚧 系統提示：尚無去年度文字資料，待後台擷取後自動匯入。)*")
            
            # (舊版的獨立照片顯示區塊已徹底刪除！)
            
        st.markdown("---")
        
        # 🎯 年度成果填報與資料上傳
        with st.form("report_form"):
            report_text = st.text_area("✍️ 填報資訊/年度執行亮點成果", height=150, placeholder="請在此輸入您的填寫內容...")
            uploaded_files = st.file_uploader("📎 上傳照片或佐證檔案 (支援 PDF, JPG, PNG, DOCX 等)：", accept_multiple_files=True)
            
            submitted = st.form_submit_button("📤 資料確認送出")
            
            if submitted:
                if not report_text.strip() and not uploaded_files:
                    st.error("⚠️ 請至少填寫成果說明或上傳佐證檔案！")
                else:
                    st.success("🎉 資料已暫存成功！目前系統畫面設計與邏輯皆已完備。")
                    st.info("*(準備進入下一階段：將填寫內容寫入資料庫！)*")

elif 'df_questions' in locals() and df_questions.empty:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")