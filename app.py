import streamlit as st
import streamlit_authenticator as stauth

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="嘉大綠色大學評比資料填報及彙整平台", page_icon="🌱", layout="wide")

# --- 2. 讀取安全保險箱 (Secrets) ---
try:
    credentials = st.secrets["credentials"].to_dict()
    cookie = st.secrets["cookie"].to_dict()
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
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # 調整欄寬比例，讓標題有足夠的空間展開不折行
    col1, col2, col3 = st.columns([1, 4, 1])
    
    with col2:
        # [修改 1] 登入頁面標題放大並加入英文
        st.markdown("""
            <h1 style='text-align: center; color: #2C3E50;'>🌱 國立嘉義大學綠色大學評比資料填報及彙整平台</h1>
            <h3 style='text-align: center; color: #566573; margin-top: -10px; margin-bottom: 30px;'>National Chiayi University UI GreenMetric Data Collection and Integration Platform</h3>
        """, unsafe_allow_html=True)
        
        # 如果密碼錯誤，顯示提示
        if st.session_state.get("authentication_status") is False:
            st.error("❌ 帳號或密碼錯誤，請重試。")
            
        authenticator.login(location='main')

elif st.session_state.get("authentication_status") is True:
    # --- 登入成功的畫面 ---
    username = st.session_state["username"]
    
    # 側邊欄帳號資訊
    st.sidebar.title(f"👤 歡迎, {st.session_state['name']}")
    authenticator.logout("登出", "sidebar")
    st.sidebar.markdown("---")
    
    # [修改 4] 側邊欄動態選單 (區分頁面檢視權限)
    st.sidebar.markdown("### 📌 系統導覽選單")
    if username == "admin_ui":
        # 管理員解鎖全部功能
        menu_options = [
            "🏠 平台首頁", 
            "📝 綠色大學填報區", 
            "⚙️ 綠色大學填報管理區", 
            "🌍 翻譯校正區", 
            "📰 AI新聞智能彙整區"
        ]
    else:
        # 一般使用者僅顯示首頁與填報區
        menu_options = [
            "🏠 平台首頁", 
            "📝 綠色大學填報區"
        ]
        
    selected_page = st.sidebar.radio("請選擇前往區域", menu_options, label_visibility="collapsed")
    
    # ==========================================
    # 根據選單顯示對應頁面內容
    # ==========================================
    if selected_page == "🏠 平台首頁":
        # [修改 2] 登入後畫面標題置中放大並加入英文
        st.markdown("""
            <h2 style='text-align: center; color: #2C3E50;'>🌱 國立嘉義大學綠色大學評比資料填報及彙整平台</h2>
            <h4 style='text-align: center; color: #566573; margin-top: -10px; margin-bottom: 30px;'>National Chiayi University UI GreenMetric Data Collection and Integration Platform</h4>
        """, unsafe_allow_html=True)
        
        # [修改 3] 歡迎語及說明文字卡片
        st.markdown("""
        <div style='background-color: #EBF5FB; padding: 25px; border-radius: 12px; border-left: 8px solid #3498DB; box-shadow: 0 4px 6px rgba(0,0,0,0.05);'>
            <h3 style='color: #2C3E50; margin-top: 0;'>👋 您好！歡迎來到本校「綠色大學評比資料填報及彙整平台」</h3>
            <p style='font-size: 1.15rem; color: #34495E; line-height: 1.8;'>
                UI GreenMetric世界綠色大學評比由印尼大學（Universitas Indonesia）主辦，為目前全球規模最大的永續校園評比之一。本校持續深化永續治理架構，強化跨單位協作與資料品質管理，朝向更具韌性與責任感的永續大學邁進。
            </p>
            <p style='font-size: 1.25rem; font-weight: 900; color: #C0392B; margin-bottom: 0;'>
                👉 請查看 👈 左側側邊欄 (Sidebar) 的「📝 綠色大學填報區」，進入填報資料，謝謝！
            </p>
        </div>
        """, unsafe_allow_html=True)

    elif selected_page == "📝 綠色大學填報區":
        st.title("📝 綠色大學填報區")
        st.info("💡 此區開放給「外部填報者」與「後台管理員」填報與檢視。")
        st.write("（請在此處放入您的填報表單程式碼或模組呼叫）")
        
    elif selected_page == "⚙️ 綠色大學填報管理區":
        st.title("⚙️ 綠色大學填報管理區")
        st.success("👑 此區僅供後台管理員檢視與管理。")
        st.write("（請在此處放入資料匯出、稽核或管理面板的程式碼）")
        
    elif selected_page == "🌍 翻譯校正區":
        st.title("🌍 翻譯校正區")
        st.success("👑 此區僅供後台管理員檢視與管理。")
        st.write("（請在此處放入翻譯校正工具的程式碼）")
        
    elif selected_page == "📰 AI新聞智能彙整區":
        st.title("📰 AI新聞智能彙整區")
        st.success("👑 此區僅供後台管理員檢視與管理。")
        st.write("（請在此處放入爬蟲與 AI 摘要模組的程式碼）")