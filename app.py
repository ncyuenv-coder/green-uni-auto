import streamlit as st
import streamlit_authenticator as stauth

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="嘉大綠色大學填報及彙整系統", page_icon="🌱", layout="wide")

# --- 2. 讀取安全保險箱 (Secrets) ---
try:
    credentials = dict(st.secrets["credentials"])
    cookie = st.secrets["cookie"]
except KeyError:
    st.error("⚠️ 系統找不到密碼設定，請確認 secrets.toml 檔案結構。")
    st.stop()

# --- 3. 初始化登入驗證器 ---
authenticator = stauth.Authenticate(
    credentials["usernames"],
    cookie["name"],
    cookie["key"],
    cookie["expiry_days"]
)

# --- 4. 側邊欄：渲染登入介面 ---
# authenticator.login 會自動比對我們輸入的密碼與 secrets 裡面的亂碼是否相符
name, authentication_status, username = authenticator.login("main") 

# --- 5. 主程式邏輯分流 ---
if authentication_status == False:
    st.error("❌ 帳號或密碼錯誤，請重試。")
    st.title("🌱 嘉大綠色大學填報及彙整系統")
    st.info("👈 請先從左側側邊欄輸入帳號密碼登入。")

elif authentication_status == None:
    st.title("🌱 嘉大綠色大學填報及彙整系統")
    st.info("👈 請先從左側側邊欄輸入帳號密碼登入。")

elif authentication_status == True:
    # 登入成功！
    st.sidebar.title(f"👤 歡迎, {name}")
    authenticator.logout("登出", "sidebar")
    
    st.title("🌱 嘉大綠色大學填報及彙整系統")
    
    # 根據帳號 (username) 判斷權限
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