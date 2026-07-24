import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ตั้งค่าหน้าเว็บ Streamlit
st.set_page_config(
    page_title="PVD Consolidation Calculator",
    page_icon="🏗️",
    layout="wide"
)

st.title("🏗️ โปรแกรมคำนวณการทรุดตัวและเวลาสำหรับงาน PVD")
st.caption("พัฒนาตามหลักทฤษฎีการอัดตัวคายน้ำ (Barron, Terzaghi & Carillo) พร้อม Smear Effect")

st.markdown("---")

# ==========================================
# SIDEBAR: รับค่า INPUT จากผู้ใช้งาน
# ==========================================
st.sidebar.header("⚙️ 1. พารามิเตอร์ PVD & รูปแบบการติดตั้ง")
pattern = st.sidebar.selectbox("รูปแบบการวาง PVD", ["สามเหลี่ยม (Triangular)", "สี่เหลี่ยม (Square)"])
S = st.sidebar.number_input("ระยะห่างการติดตั้ง PVD: S (m)", value=1.0, step=0.1)
H_pvd = st.sidebar.number_input("ความลึกแผ่น PVD: H (m)", value=10.0, step=0.5)
a = st.sidebar.number_input("ความกว้าง PVD: a (mm)", value=100.0, step=5.0) / 1000.0 # แปลงเป็น m
b = st.sidebar.number_input("ความหนา PVD: b (mm)", value=4.0, step=0.5) / 1000.0   # แปลงเป็น m

st.sidebar.header("🧪 2. คุณสมบัติของดิน (Soil Properties)")
H_soil = st.sidebar.number_input("ความหนาชั้นดินอ่อน: H_soil (m)", value=10.0, step=0.5)
Cv = st.sidebar.number_input("ค่า Cv (m²/year)", value=2.0, step=0.1)
ratio_Cr_Cv = st.sidebar.number_input("อัตราส่วน Cr / Cv", value=3.0, step=0.5)
Cr = Cv * ratio_Cr_Cv
Cc = st.sidebar.number_input("Compression Index: Cc", value=0.8, step=0.05)
e0 = st.sidebar.number_input("Initial Void Ratio: e0", value=2.0, step=0.1)
sigma_0 = st.sidebar.number_input("Effective Stress เดิม: σ0' (kPa)", value=50.0, step=5.0)
delta_sigma = st.sidebar.number_input("น้ำหนักถมเพิ่ม: Δσ (kPa)", value=80.0, step=5.0)

st.sidebar.header("🚧 3. ปัจจัย Smear Effect & Sand Mat")
include_smear = st.sidebar.checkbox("คิดผลกระทบ Smear Effect", value=True)
if include_smear:
    d_s_ratio = st.sidebar.number_input("อัตราส่วน ds / dw", value=2.5, step=0.1)
    kh_ks_ratio = st.sidebar.number_input("อัตราส่วน kh / ks", value=3.0, step=0.5)
else:
    d_s_ratio = 1.0
    kh_ks_ratio = 1.0

# ==========================================
# CALCULATION CORE (ส่วนการคำนวณ)
# ==========================================
# 1. เรขาคณิต PVD
d_w = (2 * (a + b)) / np.pi  # Equivalent diameter ของ PVD
d_e = 1.05 * S if "สามเหลี่ยม" in pattern else 1.13 * S # Influence zone
n = d_e / d_w

# 2. คำนวณ F(n) รวม Smear Effect
if include_smear:
    s = d_s_ratio
    Fn = np.log(n / s) + (kh_ks_ratio * np.log(s)) - 0.75
else:
    Fn = (n**2 / (n**2 - 1)) * np.log(n) - (3 * n**2 - 1) / (4 * n**2)

# 3. คำนwณการทรุดตัวสูงสุด (Final Settlement)
S_final = H_soil * (Cc / (1 + e0)) * np.log10((sigma_0 + delta_sigma) / sigma_0)

# 4. คำนวณการทรุดตัวตามเวลา (0 ถึง 365 วัน)
days = np.arange(1, 366)
times_years = days / 365.25

# Drainage path (สมมติระบายน้ำด้านบนและล่าง = H_soil / 2)
H_dr = H_soil / 2.0

U_v_list, U_r_list, U_av_list, S_t_list = [], [], [], []

for t in times_years:
    # Vertical Consolidation (Terzaghi)
    Tv = (Cv * t) / (H_dr**2)
    if Tv <= 0.286:
        U_v = np.sqrt((4 * Tv) / np.pi)
    else:
        U_v = 1 - (10**(-0.085 - 0.933 * Tv))
    U_v = min(U_v, 1.0)
    
    # Radial Consolidation (Barron)
    Tr = (Cr * t) / (d_e**2)
    U_r = 1 - np.exp((-8 * Tr) / Fn)
    U_r = min(U_r, 1.0)
    
    # Combined Consolidation (Carillo 1942)
    U_av = 1 - (1 - U_r) * (1 - U_v)
    
    # Settlement at time t
    S_t = U_av * S_final
    
    U_v_list.append(U_v * 100)
    U_r_list.append(U_r * 100)
    U_av_list.append(U_av * 100)
    S_t_list.append(S_t)

df_results = pd.DataFrame({
    "Day": days,
    "U_v (%)": U_v_list,
    "U_r (%)": U_r_list,
    "U_av (%)": U_av_list,
    "Settlement (m)": S_t_list
})

# ==========================================
# DISPLAY RESULTS (ส่วนแสดงผลลัพธ์)
# ==========================================
col1, col2, col3, col4 = st.columns(4)
col1.metric("ขนาดเส้นผ่านศูนย์กลางสมมูล (dw)", f"{d_w*1000:.1f} mm")
col2.metric("ระยะอิทธิพล (de)", f"{d_e:.2f} m")
col3.metric("การทรุดตัวสูงสุด (S_final)", f"{S_final:.3f} m")

# หาเวลาที่บรรลุ U_av = 90%
u90_row = df_results[df_results["U_av (%)"] >= 90].first_valid_index()
if u90_row is not None:
    days_to_90 = df_results.loc[u90_row, "Day"]
    col4.metric("เวลาบรรลุ U = 90%", f"{days_to_90} วัน", delta="เป้าหมายสำเร็จ", delta_color="normal")
else:
    col4.metric("เวลาบรรลุ U = 90%", "> 365 วัน", delta="ต้องเพิ่มระยะ PVD", delta_color="inverse")

st.markdown("---")

# กราฟแสดงผล
tab1, tab2 = st.tabs(["📉 กราฟการทรุดตัว (Settlement Curve)", "📊 เปอร์เซ็นต์การอัดตัวคายน้ำ (Degree of Consolidation)"])

with tab1:
    fig1, ax1 = plt.subplots(figsize=(10, 4))
    ax1.plot(df_results["Day"], df_results["Settlement (m)"], color="red", linewidth=2, label="Settlement (m)")
    ax1.set_xlabel("เวลา (วัน)")
    ax1.set_ylabel("การทรุดตัว (m)")
    ax1.invert_yaxis()  # กลับแกน Y เพื่อให้เห็นการยุบตัวลงล่าง
    ax1.grid(True, linestyle="--", alpha=0.6)
    ax1.legend()
    st.pyplot(fig1)

with tab2:
    fig2, ax2 = plt.subplots(figsize=(10, 4))
    ax2.plot(df_results["Day"], df_results["U_r (%)"], label="Radial (Ur) - PVD", linestyle="--")
    ax2.plot(df_results["Day"], df_results["U_v (%)"], label="Vertical (Uv) - Soil", linestyle="--")
    ax2.plot(df_results["Day"], df_results["U_av (%)"], label="Combined (Uav)", color="green", linewidth=2)
    ax2.axhline(y=90, color="orange", linestyle=":", label="Target 90%")
    ax2.set_xlabel("เวลา (วัน)")
    ax2.set_ylabel("Consolidation Degree (%)")
    ax2.set_ylim(0, 100)
    ax2.grid(True, linestyle="--", alpha=0.6)
    ax2.legend()
    st.pyplot(fig2)

# ตารางข้อมูล
with st.expander("📋 ดูตารางข้อมูลผลการคำนวณรายวัน"):
    st.dataframe(df_results.style.format({
        "U_v (%)": "{:.2f}",
        "U_r (%)": "{:.2f}",
        "U_av (%)": "{:.2f}",
        "Settlement (m)": "{:.4f}"
    }))