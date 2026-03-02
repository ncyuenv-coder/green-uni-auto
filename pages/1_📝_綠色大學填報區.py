import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 🛡️ 資安防護罩：檢查是否已登入
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報區", page_icon="📝", layout="wide")

# ==========================================
# 🔌 資料庫連線設定 (使用快取避免超過 API 限制)
# ==========================================
@st.cache_data(ttl=600) # 快取 10 分鐘，不用每次點擊都重新讀取 Google Sheet
def load_gsheet_data():
    try:
        # 1. 設定 Google API 權限範圍
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        # 2. 從保險箱 (secrets) 讀取金鑰
        skey = dict(st.secrets["gcp_oauth"])
        credentials = Credentials.from_service_account_info(skey, scopes=scopes)
        
        # 3. 授權並連線
        gc = gspread.authorize(credentials)
        
        # 4. 打開你的 Google Sheet
        SHEET_ID = '1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8'
        sh = gc.open_by_key(SHEET_ID)
        
        # 5. 選擇名為「評比題目表」的分頁
        worksheet = sh.worksheet("評比題目表")
        
        # 6. 把資料抓下來，轉成 Pandas 格式方便我們操作
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        
        # 把所有欄位名稱的頭尾空白去掉，避免對應不到
        df.columns = df.columns.str.strip()
        return df
        
    except Exception as e:
        st.error(f"⚠️ 無法讀取 Google Sheet，請確認權限是否設定正確或分頁名稱是否為「評比題目表」。\n詳細錯誤：{e}")
        return pd.DataFrame()

# ==========================================
# 🚀 網頁主畫面：填報介面設計
# ==========================================
st.title("📝 嘉大綠色大學填報區")
st.info(f"👤 目前登入單位：{st.session_state.get('username', '未知')}")

# 顯示載入動畫，呼叫我們上面寫好的讀取函數
with st.spinner('🔄 正在從雲端載入最新評比題目，請稍候...'):
    df_questions = load_gsheet_data()

# 如果有成功抓到資料，就顯示在畫面上
if not df_questions.empty:
    st.success("✅ 題目資料庫連線成功！")
    st.markdown("---")
    
    # 【第一區：選擇題目】
    # 把「題號」跟「題目名稱」合併起來，放在下拉選單讓使用者選
    df_questions['選項標示'] = df_questions['題號'].astype(str) + " - " + df_questions['題目名稱'].astype(str)
    
    selected_option = st.selectbox(
        "📌 請選擇您要填報的評比項目：", 
        df_questions['選項標示'].tolist()
    )
    
    # 根據使用者選的項目，把那一題的詳細資料抓出來
    selected_q_id = selected_option.split(" - ")[0]
    # 找出題號符合的那一橫列資料
    question_data = df_questions[df_questions['題號'].astype(str) == selected_q_id].iloc[0]
    
    # 【第二區：顯示題目詳細資訊】
    st.subheader(f"📖 {selected_option}")
    
    # 用兩個欄位並排顯示，左邊是需求，右邊是參考資料
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**🔑 關鍵字：**")
        st.write(question_data.get('關鍵字', '無'))
        
        st.markdown("**📄 內容需求：**")
        st.write(question_data.get('內容需求', '無'))
        
    with col2:
        st.markdown("**💡 前一年度參考資料：**")
        st.info(question_data.get('前一年度參考資料', '尚無參考資料（待 AI 翻譯後自動帶入）'))
        
    st.markdown("---")
    
    # 【第三區：實際填寫表單 (預留位置)】
    st.subheader("📤 本年度資料填報與上傳")
    st.write("*(🚧 這裡將是我們下一步要開發的：文字輸入框與檔案上傳按鈕)*")
    
else:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")