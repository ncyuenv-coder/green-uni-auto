import streamlit as st
import streamlit_authenticator as stauth

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="嘉大綠色大學填報及彙整系統", page_icon="🌱", layout="wide")

# --- 2. 讀取安全保險箱 (Secrets) ---
try:
    credentials = dict(st.secrets["credentials"])
    cookie = dict(st.secrets["cookie"])
except KeyError:
    st.error("⚠️ 系統找不到密碼設定，請確認 secrets.toml 檔案結構。")
    st.stop()

# --- 3. 初始化登入驗證器 ---
authenticator = stauth.Authenticate(
    credentials,
    cookie["name"],
    cookie["key"],
    cookie["expiry_days"]
)

# --- 4. 主程式邏輯分流與登入介面 ---
if st.session_state.get("authentication_status") is None or st.session_state.get("authentication_status") is False:
    # --- 未登入或登入失敗的畫面 (置中登入框) ---
    st.markdown("<br><br><br>", unsafe_allow_html=True) # 增加上方空白，讓畫面更置中
    
    # 建立三個欄位：左(1)、中(1.5)、右(1)，把內容擠在中間
    col1, col2, col3 = st.columns([1, 1.5, 1])
    
    with col2:
        st.title("🌱 嘉大綠色大學系統")
        st.markdown("### 填報及彙整平台")
        
        # 如果密碼錯誤，顯示提示
        if st.session_state.get("authentication_status") is False:
            st.error("❌ 帳號或密碼錯誤，請重試。")
            
        # 🌟 修正點：將 location 設為 'main'，它就會顯示在這裡而不是側邊欄
        authenticator.login(location='main')

elif st.session_state.get("authentication_status") is True:
    # --- 登入成功的畫面 ---
    st.sidebar.title(f"👤 歡迎, {st.session_state['name']}")
    authenticator.logout("登出", "sidebar")
    
    st.title("🌱 嘉大綠色大學填報及彙整系統")
    username = st.session_state["username"]
    
    # 根據帳號判斷專屬權限
    if username == "admin_ui":
        st.success("👑 您目前的身分是：系統管理者")
        admin_action = st.radio("請選擇管理員功能：", 
                                ["📊 填報狀況總覽", "📰 新聞爬蟲與 AI 摘要", "📄 產製最終 Word 報表"], 
                                horizontal=True)
        st.markdown("---")
        if admin_action == "📊 填報狀況總覽":
            st.write("這裡未來會顯示各單位的填報進度。")
        elif admin_action == "📰 新聞爬蟲與 AI 摘要":
            st.write("這裡未來會放置爬蟲工具，自動抓取新聞並呼叫 Gemini 進行 SDGs 分類。")
        elif admin_action == "📄 產製最終 Word 報表":
            st.write("這裡未來會有一鍵下載按鈕，將所有資料打包成 Word 檔。")
            
    elif username == "ncyu_ui":
        st.success("✅ 您目前的身分是：一般填報單位")
        st.subheader("📝 年度資料填報區")
        st.write("這裡未來會自動帶入前一年度的 Word 翻譯參照，並讓您上傳佐證資料。")