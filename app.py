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

# ==========================================
# 共用元件：滿版橫幅資訊卡模板 (HTML/CSS)
# ==========================================
banner_html = """
<div style='background-color: #2C3E50; padding: 40px 20px; border-radius: 15px; text-align: center; box-shadow: 0 8px 16px rgba(0,0,0,0.15); margin-bottom: 40px;'>
    <h1 style='color: #FFFFFF; margin-bottom: 10px; font-weight: 900; letter-spacing: 1px;'>🌱 國立嘉義大學綠色大學評比資料填報及彙整平台</h1>
    <h3 style='color: #AED6F1; margin-top: 0; font-weight: 500;'>National Chiayi University UI GreenMetric Data Collection and Integration Platform</h3>
</div>
"""

# --- 4. 主程式邏輯分流與登入介面 ---
if st.session_state.get("authentication_status") is None or st.session_state.get("authentication_status") is False:
    # --- 未登入或登入失敗的畫面 ---
    st.markdown("<br>", unsafe_allow_html=True)
    
    # [修改 1] 登入頁面：套用滿版橫幅資訊卡
    st.markdown(banner_html, unsafe_allow_html=True)
    
    # 將登入框維持在畫面正中央 (調整比例)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        # 如果密碼錯誤，顯示提示
        if st.session_state.get("authentication_status") is False:
            st.error("❌ 帳號或密碼錯誤，請重試。")
            
        authenticator.login(location='main')

elif st.session_state.get("authentication_status") is True:
    # --- 登入成功的畫面 ---
    username = st.session_state["username"]
    
    # 側邊欄：僅保留帳號資訊與登出按鈕 (已刪除手寫的系統導覽選單)
    st.sidebar.title(f"👤 歡迎, {st.session_state['name']}")
    authenticator.logout("登出", "sidebar")
    st.sidebar.markdown("---")
    
    # [修改 2] 首頁：套用滿版橫幅資訊卡
    st.markdown(banner_html, unsafe_allow_html=True)
    
    # [修改 3] 歡迎語及說明文字卡片
    st.markdown("""
    <div style='background-color: #EBF5FB; padding: 30px; border-radius: 12px; border-left: 8px solid #3498DB; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px;'>
        <h3 style='color: #2C3E50; margin-top: 0;'>👋 您好！歡迎來到本校「綠色大學評比資料填報及彙整平台」</h3>
        <p style='font-size: 1.15rem; color: #34495E; line-height: 1.8;'>
            UI GreenMetric世界綠色大學評比由印尼大學（Universitas Indonesia）主辦，為目前全球規模最大的永續校園評比之一。本校持續深化永續治理架構，強化跨單位協作與資料品質管理，朝向更具韌性與責任感的永續大學邁進。
        </p>
        <p style='font-size: 1.25rem; font-weight: 900; color: #C0392B; margin-bottom: 0;'>
            👉 請查看 👈 左側側邊欄 (Sidebar) 的「📝 綠色大學填報區」，進入填報資料，謝謝！
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # 若有針對不同管理員的客製化首頁提示，可寫在這裡
    if username == "admin_ui":
        st.info("👑 系統管理員您好，您可透過左側選單進入各項管理與校正工具。")