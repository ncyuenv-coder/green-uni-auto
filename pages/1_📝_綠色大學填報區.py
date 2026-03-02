import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials

# ==========================================
# 🛡️ 資安防護罩：檢查是否已登入
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報區", page_icon="📝", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定 (莫蘭迪色系與排版)
# ==========================================
st.markdown("""
<style>
    /* 莫蘭迪灰藍色標題區塊 */
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
    
    /* 莫蘭迪淺灰綠色 - 資料需求區塊 */
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

    /* 莫蘭迪橘色 - 置中送出按鈕 */
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
# 🚀 網頁主畫面：填報介面設計
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
        
    # ==========================================
    # 🎯 選擇單位與題目
    # ==========================================
    unit_list = df_questions['權責單位'].dropna().unique().tolist()
    unit_list = [str(u).strip() for u in unit_list if str(u).strip() != '']
    
    # 步驟一：選擇單位
    st.markdown("<div class='morandi-title'>🏢 請選擇您的單位</div>", unsafe_allow_html=True)
    selected_unit = st.selectbox("", ["請選擇..."] + unit_list, label_visibility="collapsed")
    
    if selected_unit != "請選擇...":
        df_unit_questions = df_questions[df_questions['權責單位'].astype(str).str.strip() == selected_unit].copy()
        
        # 步驟二：選擇填報項目 (當年度題目 - 中文標題)
        df_unit_questions['選項標示'] = df_unit_questions['當年度題目'].astype(str) + " - " + df_unit_questions['中文標題'].astype(str)
        
        st.markdown("<div class='morandi-title'>📌 請選擇填報項目</div>", unsafe_allow_html=True)
        selected_option = st.selectbox("", df_unit_questions['選項標示'].tolist(), label_visibility="collapsed")
        
        selected_q_id = selected_option.split(" - ")[0]
        question_data = df_unit_questions[df_unit_questions['當年度題目'].astype(str) == selected_q_id].iloc[0]
        
        st.markdown("---")
        
        # ==========================================
        # 🎯 顯示題目詳細資訊
        # ==========================================
        # 1. 填報項目顯示標題 (當年度題目：中文標題)
        st.markdown(f"<div class='morandi-title'>📖 {question_data.get('當年度題目')}：{question_data.get('中文標題')}</div>", unsafe_allow_html=True)
        
        # 2. 中文說明 (黑色字體)
        st.markdown(f"<div style='color: black; font-size: 1.1em; padding-left: 5px; margin-bottom: 15px;'><b>中文說明：</b><br>{question_data.get('中文說明', '無')}</div>", unsafe_allow_html=True)
        
        # 3. 資料需求 (莫蘭迪區塊)
        st.markdown(f"<div class='morandi-req'><b>🔍 資料需求：</b><br>{question_data.get('資料需求', '無特別說明')}</div>", unsafe_allow_html=True)
        
        # 4. 前一年度參考資訊 (摺疊展開) - ✨ 本次更新重點區域
        with st.expander("💡 點擊展開查看：前一年度 (2025) 參考資訊"):
            st.write(f"**對應之去年度題目：** {question_data.get('前一年度題目', '無')}")
            
            # 安全讀取預留欄位 (使用 pd.notna 避免讀到空值報錯)
            ref_text = question_data.get('2025參考文字_AI預留', '')
            ref_img_urls = question_data.get('2025參考圖片連結_AI預留', '')
            
            st.markdown("---")
            
            # 顯示文字翻譯
            if pd.notna(ref_text) and str(ref_text).strip() != "":
                st.markdown("**📝 去年度填報內容 (中文翻譯)：**")
                st.write(str(ref_text))
            else:
                st.info("*(🚧 系統提示：尚無去年度文字資料，待後台擷取後自動匯入。)*")
                
            # 顯示照片 (將逗號隔開的網址切開，並排顯示)
            if pd.notna(ref_img_urls) and str(ref_img_urls).strip() != "":
                st.markdown("**📸 去年度佐證照片：**")
                # 把網址用逗號切開，變成一個列表
                urls = [url.strip() for url in str(ref_img_urls).split(',') if url.strip()]
                
                if urls:
                    # 根據照片數量建立對應數量的欄位 (columns) 讓照片並排
                    cols = st.columns(len(urls))
                    for idx, col in enumerate(cols):
                        with col:
                            try:
                                # use_container_width=True 讓照片自動縮放適應欄位寬度
                                st.image(urls[idx], use_container_width=True)
                            except Exception as e:
                                st.error("⚠️ 圖片載入失敗，請確認網址是否正確。")
            
        st.markdown("---")
        
        # ==========================================
        # 🎯 年度成果填報與資料上傳
        # ==========================================
        with st.form("report_form"):
            # A. 填報資訊/年度執行亮點成果
            report_text = st.text_area(
                "✍️ 填報資訊/年度執行亮點成果", 
                height=150, 
                placeholder="請在此輸入您的填寫內容..."
            )
            
            # B. 上傳照片或佐證檔案
            uploaded_files = st.file_uploader(
                "📎 上傳照片或佐證檔案 (支援 PDF, JPG, PNG, DOCX 等)：", 
                accept_multiple_files=True
            )
            
            # 5. 資料確認送出按鈕
            submitted = st.form_submit_button("📤 資料確認送出")
            
            if submitted:
                if not report_text.strip() and not uploaded_files:
                    st.error("⚠️ 請至少填寫成果說明或上傳佐證檔案！")
                else:
                    st.success("🎉 資料已暫存成功！目前系統畫面設計與邏輯皆已完備。")
                    st.info("*(準備進入下一階段：我們將把這些填寫的內容與檔案，自動寫入 Google Sheet 資料庫與 Drive 中！)*")

elif 'df_questions' in locals() and df_questions.empty:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")