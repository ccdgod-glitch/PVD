import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ---------------------------------------------------------
# 1. PAGE CONFIGURATION & CUSTOM CSS
# ---------------------------------------------------------
st.set_page_config(
    page_title="PVD Settlement Analytics",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS เพิ่มความสวยงามสไตล์ Modern Dashboard
st.markdown("""
<style>
    /* ปรับแต่งส่วน Header */
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #6B7280;
        margin-bottom: 1.5rem;
    }
    /* แต่งการ์ด Metric */
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: bold;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding-left: 16px;
        padding-right: 16px;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. HEADER
# ---------------------------------------------------------
st.markdown('<div class="main-header">🏗️ PVD Consolidation Analytics</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">ระบบจำลองการทรุดตัวและคำนวณเวลาการอัดตัวคายน้ำชั้นดินอ่อน (Barron & Terzaghi Theory)</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# 3. SIDEBAR - INPUT PARAMETERS (จัดกลุ่มด้วย Expander)
# ---------------------------------------------------------
st.sidebar.image("https://img.icons8.com/color/96/000000/crane.png", width=60)
st.sidebar.title("⚙️ ตั้งค่าพารามิเตอร์")

with st.sidebar.expander("📐 1. รูปแบบ PVD & ระยะติดตั้ง", expanded=True):
    pattern = st.selectbox("รูปแบบการจัดวาง", ["สามเหลี่ยม (Triangular)", "สี่เหลี่ยม (Square)"])
    S = st.number_input("ระยะห่างการติดตั้ง S (m)", value=1.0, step=0.1)
    H_pvd = st.number_input("ความลึกแผ่น PVD (m)", value=10.0, step=0.5)
    a = st.number_input("ความกว้าง PVD: a (mm)", value=100.0, step=5.0) / 1000.0
    b = st.number_input("ความหนา PVD: b (mm)", value=4.0, step=0.5) / 1000.0

with st.sidebar.expander("🧪 2. คุณสมบัติชั้นดิน (Soil Profile)", expanded=False):
    H_soil = st.number_input("ความหนาชั้นดินอ่อน (m)", value=10.0, step=0.5)
    Cv = st.number_input("ค่า Cv (m²/year)", value=2.0, step=0.1)
    ratio_Cr_Cv = st.number_input("อัตราส่วน Cr / Cv", value=3.0, step=0.5)
    Cr = Cv * ratio_Cr_Cv
    Cc = st.number_input("Compression Index (Cc)", value=0.8, step=0.05)
    e0 = st.number_input("Initial Void Ratio (e0)", value=2.0, step=0.1)
    sigma_0 = st.number_input("Effective Stress เดิม: σ0' (kPa)", value=50.0, step=5.0)
    delta_sigma = st.number_input("น้ำหนักถมเพิ่ม: Δσ (kPa)", value=80.0, step=5.0)

with st.sidebar.expander("🚧 3. ผลกระทบ Smear Effect", expanded=False):
    include_smear = st.checkbox("คิดผลกระทบ Smear Effect", value=True)
    if include_smear:
        d_s_ratio = st.number_input("อัตราส่วน ds / dw", value=2.5, step=0.1)
        kh_ks_ratio = st.number_input("อัตราส่วน kh / ks", value=3.0, step=0.5)
    else:
        d_s_ratio, kh_ks_ratio = 1.0, 1.0

# ---------------------------------------------------------
# 4. CALCULATION ENGINE
# ---------------------------------------------------------
d_w = (2 * (a + b)) / np.pi
d_e = 1.05 * S if "สามเหลี่ยม" in pattern else 1.13 * S
n = d_e / d_w

if include_smear:
    s = d_s_ratio
    Fn = np.log(n / s) + (kh_ks_ratio * np.log(s)) - 0.75
else:
    Fn = (n**2 / (n**2 - 1)) * np.log(n) - (3 * n**2 - 1) / (4 * n**2)

S_final = H_soil * (Cc / (1 + e0)) * np.log10((sigma_0 + delta_sigma) / sigma_0)

days = np.arange(1, 366)
times_years = days / 365.25
H_dr = H_soil / 2.0

U_v_list, U_r_list, U_av_list, S_t_list = [], [], [], []

for t in times_years:
    Tv = (Cv * t) / (H_dr**2)
    U_v = np.sqrt((4 * Tv) / np.pi) if Tv <= 0.286 else 1 - (10**(-0.085 - 0.933 * Tv))
    U_v = min(U_v, 1.0)
    
    Tr = (Cr * t) / (d_e**2)
    U_r = 1 - np.exp((-8 * Tr) / Fn)
    U_r = min(U_r, 1.0)
    
    U_av = 1 - (1 - U_r) * (1 - U_v)
    S_t = U_av * S_final
    
    U_v_list.append(U_v * 100)
    U_r_list.append(U_r * 100)
    U_av_list.append(U_av * 100)
    S_t_list.append(S_t)

df = pd.DataFrame({
    "Day": days, "U_v": U_v_list, "U_r": U_r_list, "U_av": U_av_list, "Settlement": S_t_list
})

# ---------------------------------------------------------
# 5. DASHBOARD DISPLAY
# ---------------------------------------------------------
# Key Metrics
u90_idx = df[df["U_av"] >= 90].first_valid_index()
days_90 = df.loc[u90_idx, "Day"] if u90_idx is not None else "> 365"

m1, m2, m3, m4 = st.columns(4)
m1.metric("เส้นผ่านศูนย์กลางเทียบเท่า (dw)", f"{d_w*1000:.1f} mm")
m2.metric("ระยะอิทธิพลการระบาย (de)", f"{d_e:.2f} m")
m3.metric("การทรุดตัวสูงสุด (S final)", f"{S_final:.3f} m")
m4.metric("เวลาบรรลุ U = 90%", f"{days_90} วัน", 
          delta="ตามเป้าหมาย" if isinstance(days_90, (int, np.integer)) else "ช้าเกินไป", 
          delta_color="normal" if isinstance(days_90, (int, np.integer)) else "inverse")

st.markdown("<br>", unsafe_allow_html=True)

# Main Content Tabs
tab_charts, tab_data, tab_summary = st.tabs(["📊 กราฟวิเคราะห์ (Interactive Charts)", "📋 ตารางข้อมูล (Data Table)", "💡 สรุปผลการออกแบบ"])

with tab_charts:
    col_left, col_right = st.columns(2)
    
    with col_left:
        # Plotly: กราฟการทรุดตัว
        fig_settle = go.Figure()
        fig_settle.add_trace(go.Scatter(
            x=df["Day"], y=df["Settlement"],
            mode='lines', name='Settlement (m)',
            line=dict(color='#EF4444', width=3),
            hovertemplate='วันที %{x}<br>ทรุดตัว: %{y:.3f} m<extra></extra>'
        ))
        fig_settle.update_layout(
            title="📉 กราฟพัฒนาการการทรุดตัวตามเวลา",
            xaxis_title="เวลา (วัน)", yaxis_title="การทรุดตัว (เมตร)",
            yaxis=dict(autorange="reversed"), # กลับแกน Y ให้เห็นการยุบตัว
            template="plotly_white", height=380
        )
        st.plotly_chart(fig_settle, use_container_width=True)

    with col_right:
        # Plotly: กราฟ Degree of Consolidation
        fig_u = go.Figure()
        fig_u.add_trace(go.Scatter(x=df["Day"], y=df["U_av"], name='รวม (U_av)', line=dict(color='#10B981', width=3)))
        fig_u.add_trace(go.Scatter(x=df["Day"], y=df["U_r"], name='แนวรัศมี PVD (U_r)', line=dict(color='#3B82F6', dash='dash')))
        fig_u.add_trace(go.Scatter(x=df["Day"], y=df["U_v"], name='แนวดิ่ง ดิน (U_v)', line=dict(color='#9CA3AF', dash='dot')))
        fig_u.add_hline(y=90, line_dash="dash", line_color="#F59E0B", annotation_text="Target 90%")
        
        fig_u.update_layout(
            title="📊 อัตราการอัดตัวคายน้ำ (Degree of Consolidation)",
            xaxis_title="เวลา (วัน)", yaxis_title="Consolidation (%)",
            yaxis=dict(range=[0, 105]), template="plotly_white", height=380
        )
        st.plotly_chart(fig_u, use_container_width=True)

with tab_data:
    st.subheader("ตารางแสดงค่าการคำนวณรายวัน")
    # ปุ่ม Download CSV
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("📥 ดาวน์โหลดข้อมูลเป็น CSV", data=csv, file_name="pvd_consolidation_result.csv", mime="text/csv")
    
    st.dataframe(
        df.style.format({"U_v": "{:.2f}%", "U_r": "{:.2f}%", "U_av": "{:.2f}%", "Settlement": "{:.4f} m"}),
        use_container_width=True, height=300
    )

with tab_summary:
    st.success(f"**สรุปผลการประเมิน:** ในระยะติดตั้ง PVD ที่ **{S:.2f} เมตร** รูปแบบ **{pattern}**")
    st.write(f"- การทรุดตัวทั้งหมดเมื่อสิ้นสุดกระบวนการ ($S_{{final}}$): **{S_final:.3f} เมตร**")
    if isinstance(days_90, (int, np.integer)):
        st.write(f"- ดินจะทรุดตัวถึง 90% ($S_{{90}} = {S_final*0.9:.3f}$ เมตร) ภายในเวลา **{days_90} วัน**")
    else:
        st.warning("- การติดตั้ง PVD ระยะนี้ยังไม่สามารถทำให้ดินทรุดตัวถึง 90% ได้ภายใน 1 ปี แนะนำให้ **ลดระยะห่าง S** ลง")