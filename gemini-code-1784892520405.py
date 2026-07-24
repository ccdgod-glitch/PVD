import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------------------------------------------------------
# 1. PAGE CONFIGURATION & CUSTOM CSS
# ---------------------------------------------------------
st.set_page_config(
    page_title="PVD Settlement Analytics",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.2rem; }
    .sub-header { font-size: 1rem; color: #6B7280; margin-bottom: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem; font-weight: bold; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { padding-left: 16px; padding-right: 16px; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. HELPER FUNCTION: สร้างไฟล์ WORD สรุปรายงานเชิงวิศวกรรม
# ---------------------------------------------------------
def generate_word_report(pattern, S, H_soil, H_pvd, Cv, Cr, target_day, Uav_pct, Ur_pct, Uv_pct, S_final, S_target, days_90):
    doc = Document()
    
    # 1. หัวข้อรายงาน
    title = doc.add_heading('รายงานสรุปผลการวิเคราะห์การเร่งการทรุดตัวด้วย PVD', level=1)
    title.runs[0].font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)  # <--- แก้เป็นแบบนี้
    
    doc.add_paragraph(f"วันที่ออกรายงาน: {pd.Timestamp.now().strftime('%d/%m/%Y')}")
    doc.add_paragraph("-" * 55)
    
    # 2. หมวดพารามิเตอร์
    doc.add_heading('1. ข้อมูลและพารามิเตอร์การออกแบบ (Input Parameters)', level=2)
    p1 = doc.add_paragraph()
    p1.add_run("จากการประเมินพื้นที่โครงการ ซึ่งมีชั้นดินอ่อนความหนา ")
    p1.add_run(f"{H_soil:.1f} เมตร ").bold = True
    p1.add_run(f"ทำการติดตั้งแผ่นระบายน้ำแนวดิ่ง (PVD) ความลึก {H_pvd:.1f} เมตร โดยจัดวางในรูปแบบ ")
    p1.add_run(f"{pattern} ").bold = True
    p1.add_run("ที่ระยะห่างการติดตั้ง (Spacing) เท่ากับ ")
    p1.add_run(f"{S:.2f} เมตร ").bold = True
    p1.add_run(f"กำหนดค่าสัมประสิทธิ์การอัดตัวคายน้ำ Cv = {Cv:.1f} cm²/day และ Cr = {Cr:.1f} cm²/day")

    # 3. หมวดผลการคำนวณ
    doc.add_heading('2. ผลการวิเคราะห์และคำนวณ ณ เวลาเป้าหมาย', level=2)
    status_text = "บรรลุตามข้อกำหนดการออกแบบ (>= 90%)" if Uav_pct >= 90 else "ยังไม่บรรลุตามเป้าหมาย (< 90%)"
    
    p2 = doc.add_paragraph()
    p2.add_run(f"ณ ระยะเวลาเป้าหมายที่ ")
    p2.add_run(f"{target_day} วัน ").bold = True
    p2.add_run("ภายหลังการติดตั้ง PVD และถมดินเพิ่มน้ำหนัก ผลการคำนวณพารามิเตอร์สำคัญมีดังนี้:\n\n")
    p2.add_run(f"  • อัตราการอัดตัวคายน้ำแนวดิ่งของชั้นดิน (Uv): {Uv_pct:.2f}%\n")
    p2.add_run(f"  • อัตราการอัดตัวคายน้ำแนวรัศมีผ่าน PVD (Ur): {Ur_pct:.2f}%\n")
    p2.add_run(f"  • อัตราการอัดตัวคายน้ำเฉลี่ยรวม (Uav): {Uav_pct:.2f}% [{status_text}]\n")
    p2.add_run(f"  • ปริมาณการทรุดตัวที่เกิดขึ้นแล้ว (St): {S_target:.3f} เมตร (จากค่าการทรุดตัวสูงสุด Sfinal = {S_final:.3f} เมตร)")

    # 4. บทสรุปและข้อเสนอแนะเชิงวิศวกรรม (ข้อความปรับเปลี่ยนตามผลลัพธ์อัตโนมัติ)
    doc.add_heading('3. สรุปผลและข้อเสนอแนะเชิงวิศวกรรม', level=2)
    p3 = doc.add_paragraph()
    
    if Uav_pct >= 90.0:
        p3.add_run(
            f"สรุปการประเมิน: การติดตั้ง PVD ระยะห่าง {S:.2f} เมตร มีความเหมาะสมและเพียงพอทางการวิศวกรรม "
            f"โดยชั้นดินสามารถคายน้ำและเกิดการทรุดตัวได้ถึง {Uav_pct:.2f}% ภายในระยะเวลา {target_day} วัน "
            f"(ใช้เวลาเพียง {days_90} วันในการบรรลุ U=90%) ซึ่งเป็นไปตามมาตรฐานการควบคุมการทรุดตัวก่อนการเปิดใช้งานพื้นที่"
        )
    else:
        p3.add_run(
            f"ข้อสังเกตทางการวิศวกรรม: การติดตั้ง PVD ที่ระยะห่าง {S:.2f} เมตร ณ เวลา {target_day} วัน "
            f"ทำให้เกิดการอัดตัวคายน้ำได้เพียง {Uav_pct:.2f}% ซึ่งยังไม่ถึงเกณฑ์ 90% "
            f"จึงมีข้อเสนอแนะให้ปรับลดระยะห่างการติดตั้ง PVD (S) ให้แคบลง เช่น ปรับลดเหลือ {max(0.6, S-0.2):.2f} เมตร "
            f"หรือขยายระยะเวลาในการรอคอยการทรุดตัว (Preloading Period) เพิ่มเติม"
        )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# 3. HEADER & SIDEBAR INPUTS
# ---------------------------------------------------------
st.markdown('<div class="main-header">🏗️ PVD Consolidation Analytics</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">ระบบจำลองการทรุดตัวและคำนวณเวลาการอัดตัวคายน้ำชั้นดินอ่อน (พร้อมระบบออกรายงาน Word)</div>', unsafe_allow_html=True)

st.sidebar.title("⚙️ ตั้งค่าพารามิเตอร์การออกแบบ")

with st.sidebar.expander("📐 1. รูปแบบ PVD & ระยะติดตั้ง", expanded=True):
    pattern = st.selectbox("รูปแบบการจัดวาง", ["สามเหลี่ยม (Triangular)", "สี่เหลี่ยม (Square)"])
    S = st.number_input("ระยะห่างการติดตั้ง S (m)", value=1.00, step=0.05, format="%.2f")
    H_pvd = st.number_input("ความลึกแผ่น PVD (m)", value=10.0, step=0.5)
    a_mm = st.number_input("ความกว้าง PVD: a (mm)", value=100.0, step=5.0)
    b_mm = st.number_input("ความหนา PVD: b (mm)", value=5.0, step=0.5)

with st.sidebar.expander("🧪 2. คุณสมบัติชั้นดิน & ทางระบายน้ำ", expanded=True):
    H_soil = st.number_input("ความหนาชั้นดินอ่อน: H_soil (m)", value=30.0, step=1.0)
    drainage_type = st.radio("ลักษณะการระบายน้ำแนวดิ่ง (Drainage Path)", ["ระบายน้ำ 2 ทาง (บน-ล่าง: Hd = H/2)", "ระบายน้ำทางเดียว (Hd = H)"])
    
    st.markdown("---")
    unit_choice = st.selectbox("เลือกหน่วยของ Cv และ Cr", ["cm²/day (ตรงตามสไลด์)", "m²/year"])
    
    if "cm²/day" in unit_choice:
        Cv_input = st.number_input("ค่า Cv (cm²/day)", value=20.0, step=1.0)
        Cr_input = st.number_input("ค่า Cr (cm²/day)", value=140.0, step=5.0)
        Cv_cm2_day, Cr_cm2_day = Cv_input, Cr_input
    else:
        Cv_m2_yr = st.number_input("ค่า Cv (m²/year)", value=2.0, step=0.1)
        Cr_m2_yr = st.number_input("ค่า Cr (m²/year)", value=6.0, step=0.1)
        Cv_cm2_day = (Cv_m2_yr / 365.25) * 10000.0
        Cr_cm2_day = (Cr_m2_yr / 365.25) * 10000.0

    st.markdown("---")
    Cc = st.number_input("Compression Index (Cc)", value=0.80, step=0.05)
    e0 = st.number_input("Initial Void Ratio (e0)", value=2.00, step=0.1)
    sigma_0 = st.number_input("Effective Stress เดิม: σ0' (kPa)", value=50.0, step=5.0)
    delta_sigma = st.number_input("น้ำหนักถมเพิ่ม: Δσ (kPa)", value=80.0, step=5.0)

with st.sidebar.expander("🎯 3. กำหนดวันเป้าหมาย & Smear Effect", expanded=True):
    target_day = st.number_input("วันเป้าหมายหลักที่ต้องการตรวจสอบ (วัน)", value=90, step=10, min_value=1)
    include_smear = st.checkbox("คิดผลกระทบ Smear Effect", value=False)
    if include_smear:
        d_s_ratio = st.number_input("อัตราส่วน ds / dw", value=2.5, step=0.1)
        kh_ks_ratio = st.number_input("อัตราส่วน kh / ks", value=3.0, step=0.5)
    else:
        d_s_ratio, kh_ks_ratio = 1.0, 1.0

# ---------------------------------------------------------
# 4. CALCULATION ENGINE
# ---------------------------------------------------------
a_m, b_m = a_mm / 1000.0, b_mm / 1000.0
d_w_m = (a_m + b_m) / 2.0
d_w_cm, d_w_mm = d_w_m * 100.0, d_w_m * 1000.0

d_e_m = 1.05 * S if "สามเหลี่ยม" in pattern else 1.13 * S
d_e_cm = d_e_m * 100.0
n = d_e_m / d_w_m

if include_smear:
    s = d_s_ratio
    Fn = np.log(n / s) + (kh_ks_ratio * np.log(s)) - 0.75
else:
    Fn = (n**2 / (n**2 - 1)) * np.log(n) - (3 * n**2 - 1) / (4 * n**2)

H_d_m = H_soil / 2.0 if "2 ทาง" in drainage_type else H_soil
H_d_cm = H_d_m * 100.0

S_final = H_soil * (Cc / (1 + e0)) * np.log10((sigma_0 + delta_sigma) / sigma_0)

days = np.arange(1, 366)
times_years = days / 365.25
U_v_list, U_r_list, U_av_list, S_t_list = [], [], [], []

for t_yr, d_day in zip(times_years, days):
    Tv = (Cv_cm2_day * d_day) / (H_d_cm**2)
    U_v = np.sqrt(4 * Tv) / np.pi if Tv <= 0.286 else 1 - (10**(-0.085 - 0.933 * Tv))
    U_v = min(U_v, 1.0)
    
    Tr = (Cr_cm2_day * d_day) / (d_e_cm**2)
    U_r = 1 - np.exp((-8 * Tr) / Fn)
    U_r = min(U_r, 1.0)
    
    U_av = 1 - (1 - U_r) * (1 - U_v)
    S_t = U_av * S_final
    
    U_v_list.append(U_v * 100)
    U_r_list.append(U_r * 100)
    U_av_list.append(U_av * 100)
    S_t_list.append(S_t)

df = pd.DataFrame({"Day": days, "U_v": U_v_list, "U_r": U_r_list, "U_av": U_av_list, "Settlement": S_t_list})

u90_idx = df[df["U_av"] >= 90].first_valid_index()
days_90 = df.loc[u90_idx, "Day"] if u90_idx is not None else "> 365"

# ---------------------------------------------------------
# 5. DASHBOARD DISPLAY & TABS
# ---------------------------------------------------------
m1, m2, m3, m4 = st.columns(4)
m1.metric("เส้นผ่านศูนย์กลางเทียบเท่า (dw)", f"{d_w_mm:.1f} mm ({d_w_cm:.2f} cm)")
m2.metric("ระยะอิทธิพลการระบาย (de)", f"{d_e_m:.2f} m ({d_e_cm:.0f} cm)")
m3.metric("การทรุดตัวสูงสุด (S final)", f"{S_final:.3f} m")
m4.metric("เวลาบรรลุ U = 90%", f"{days_90} วัน", 
          delta="ตามเป้าหมาย" if isinstance(days_90, (int, np.integer)) else "ช้าเกินไป", 
          delta_color="normal" if isinstance(days_90, (int, np.integer)) else "inverse")

st.markdown("<br>", unsafe_allow_html=True)

tab_steps, tab_charts, tab_data, tab_summary = st.tabs([
    "📑 วิธีทำแบบตาราง (Step-by-Step)", 
    "📊 กราฟวิเคราะห์ (Interactive Charts)", 
    "📋 ตารางข้อมูล (Data Table)", 
    "💡 สรุปผลการออกแบบ & Export Report"
])

# TAB 1: Step-by-Step Table
with tab_steps:
    st.markdown(f"### 📑 ตารางคำนวณและตรวจสอบที่ระยะ $S = {S:.2f}$ m ณ เวลา $t = {target_day}$ วัน")
    
    # STEP 5
    st.markdown("---")
    st.markdown("#### ⑤ ตารางคำนวณเพื่อหาค่า Drain Spacing Factor: $F(n)$")
    s_test_list = [round(max(0.5, S - 0.1), 2), round(S, 2), round(S + 0.1, 2)]
    fn_rows = []
    for s_val in s_test_list:
        de_val_cm = (1.05 if "สามเหลี่ยม" in pattern else 1.13) * s_val * 100.0
        n_val = de_val_cm / d_w_cm
        n2 = n_val**2
        t1, t2, t3 = n2 / (n2 - 1), np.log(n_val), (3 * n2 - 1) / (4 * n2)
        fn_val = (t1 * t2) - t3
        fn_rows.append({
            "(1) S (m)": f"{s_val:.2f}", "(2) de (cm)": f"{de_val_cm:.1f}", "(3) n = de/dw": f"{n_val:.2f}",
            "(4) n²": f"{n2:.2f}", "(5) n²/(n²-1)": f"{t1:.3f}", "(6) ln(n)": f"{t2:.3f}",
            "(7) (3n²-1)/(4n²)": f"{t3:.3f}", "(8) F(n)": f"★ {fn_val:.3f}" if s_val == S else f"{fn_val:.3f}"
        })
    st.table(pd.DataFrame(fn_rows))

    # STEP 6
    st.markdown("---")
    st.markdown("#### ⑥ ตารางคำนวณเพื่อหาค่า Degree of Consolidation ในแนวรัศมี ($U_r$)")
    de2_cm2 = d_e_cm**2
    ur_rows = []
    for d_val in sorted(list(set([30, 60, target_day, 180]))):
        cr_t = Cr_cm2_day * d_val
        Tr_val = cr_t / de2_cm2
        exp_term = np.exp((-8 * Tr_val) / Fn)
        ur_val = (1 - exp_term) * 100.0
        ur_rows.append({
            "(1) S (m)": f"{S:.2f}", "(2) Cr (cm²/day)": f"{Cr_cm2_day:.1f}", "(3) t (day)": f"{d_val}",
            "(4) de² (cm²)": f"{de2_cm2:.0f}", "(5) Cr × t (cm²)": f"{cr_t:.0f}", "(6) Tr": f"{Tr_val:.4f}",
            "(7) F(n)": f"{Fn:.3f}", "(8) exp(-8Tr/Fn)": f"{exp_term:.4f}", "(9) Ur (%)": f"★ {ur_val:.2f}%" if d_val == target_day else f"{ur_val:.2f}%"
        })
    st.table(pd.DataFrame(ur_rows))
    
    Ur_target = df.loc[df["Day"] == target_day, "U_r"].values[0] if target_day <= 365 else (1 - np.exp((-8 * (Cr_cm2_day * target_day) / de2_cm2) / Fn)) * 100

    # STEP 7
    st.markdown("---")
    st.markdown(f"#### ⑦ คำนวณระดับการอัดตัวคายน้ำในแนวดิ่ง ($U_v$) ที่เวลา $t = {target_day}$ วัน")
    col_v1, col_v2 = st.columns([1, 1.5])
    with col_v1:
        st.latex(r"T_v = \frac{C_v \times t}{(H_d)^2}")
        st.latex(r"U_v = \frac{\sqrt{4 \times T_v}}{\pi} \quad (\text{เมื่อ } U_v \le 60\%)")
    with col_v2:
        Tv_target = (Cv_cm2_day * target_day) / (H_d_cm**2)
        Uv_target = np.sqrt(4 * Tv_target) / np.pi if Tv_target <= 0.286 else 1 - (10**(-0.085 - 0.933 * Tv_target))
        Uv_target_pct = min(Uv_target * 100.0, 100.0)
        st.write(f"- ค่าสัมประสิทธิ์การอัดตัว $C_v = \\mathbf{{{Cv_cm2_day:.1f}}}$ cm²/day")
        st.write(f"- ระยะระบายน้ำ $H_d = \\mathbf{{{H_d_cm:.0f}}}$ cm ({drainage_type.split(' ')[0]})")
        st.write(f"- ค่า Time Factor: $T_v = \\mathbf{{{Tv_target:.4f}}}$")
        st.write(f"- **ผลลัพธ์ $U_v$:** $\\mathbf{{{Uv_target_pct:.2f}\\%}}$")

    # STEP 8 & 9
    st.markdown("---")
    st.markdown("#### ⑧-⑨ คำนวณระดับการอัดตัวคายน้ำเฉลี่ยรวม ($U_{av}$)")
    Uav_target_pct = 100.0 * (1.0 - (1.0 - (Ur_target/100.0)) * (1.0 - (Uv_target_pct/100.0)))
    st.latex(f"U_{{av}} = 1 - (1 - {Ur_target/100:.4f})(1 - {Uv_target_pct/100:.4f}) = \\mathbf{{{Uav_target_pct:.2f}\\%}}")

# TAB 2 & 3: Charts & Data
with tab_charts:
    c1, c2 = st.columns(2)
    with c1:
        fig1 = go.Figure(go.Scatter(x=df["Day"], y=df["Settlement"], line=dict(color='#EF4444', width=3)))
        fig1.update_layout(title="📉 พัฒนาการการทรุดตัวตามเวลา", xaxis_title="วัน", yaxis_title="การทรุดตัว (m)", yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig1, use_container_width=True)
    with c2:
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=df["Day"], y=df["U_av"], name='รวม (U_av)', line=dict(color='#10B981', width=3)))
        fig2.add_trace(go.Scatter(x=df["Day"], y=df["U_r"], name='แนวรัศมี (U_r)', line=dict(color='#3B82F6', dash='dash')))
        fig2.add_trace(go.Scatter(x=df["Day"], y=df["U_v"], name='แนวดิ่ง (U_v)', line=dict(color='#9CA3AF', dash='dot')))
        fig2.add_hline(y=90, line_dash="dash", line_color="#F59E0B")
        fig2.update_layout(title="📊 อัตราการอัดตัวคายน้ำ (%)", xaxis_title="วัน", yaxis_title="%", yaxis=dict(range=[0, 105]))
        st.plotly_chart(fig2, use_container_width=True)

with tab_data:
    st.dataframe(df.style.format({"U_v": "{:.2f}%", "U_r": "{:.2f}%", "U_av": "{:.2f}%", "Settlement": "{:.4f} m"}), use_container_width=True)

# TAB 4: Summary & Export Word Report
with tab_summary:
    st.subheader("💡 สรุปผลการประเมินและออกรายงานเชิงวิศวกรรม")
    
    S_target = (Uav_target_pct / 100.0) * S_final
    
    if Uav_target_pct >= 90.0:
        st.success(
            f"**สรุปผลการประเมิน:** การติดตั้ง PVD ระยะห่าง **{S:.2f} เมตร** ({pattern}) "
            f"สามารถเร่งการอัดตัวคายน้ำได้ **{Uav_target_pct:.2f}%** ณ เวลา **{target_day} วัน** "
            f"คิดเป็นระยะทรุดตัว **{S_target:.3f} เมตร** (จากทั้งหมด {S_final:.3f} m) **ซึ่งผ่านเกณฑ์การออกแบบ (> 90%)**"
        )
    else:
        st.warning(
            f"**สรุปผลการประเมิน:** ณ เวลา **{target_day} วัน** การอัดตัวคายน้ำคืบหน้าไปได้ **{Uav_target_pct:.2f}%** "
            f"(ยังไม่ถึงเกณฑ์ 90%) คิดเป็นระยะทรุดตัว **{S_target:.3f} เมตร** "
            f"แนะนำให้ลดระยะห่าง S หรือเพิ่มเวลา Preloading"
        )
    
    st.markdown("---")
    st.markdown("#### 📄 ส่งออกรายงานสรุปผลสรุปเป็นคำพูด (Word Document)")
    st.write("กดปุ่มด้านล่างเพื่อดาวน์โหลดไฟล์ `.docx` ซึ่งถูกเรียบเรียงสรุปผลเป็นข้อความเชิงวิศวกรรมเรียบร้อยแล้ว")
    
    # ดึงฟังก์ชันสร้างไฟล์ Word
    doc_file = generate_word_report(
        pattern=pattern,
        S=S,
        H_soil=H_soil,
        H_pvd=H_pvd,
        Cv=Cv_cm2_day,
        Cr=Cr_cm2_day,
        target_day=target_day,
        Uav_pct=Uav_target_pct,
        Ur_pct=Ur_target,
        Uv_pct=Uv_target_pct,
        S_final=S_final,
        S_target=S_target,
        days_90=days_90
    )
    
    st.download_button(
        label="📥 ดาวน์โหลดรายงานสรุปผล (Word File .docx)",
        data=doc_file,
        file_name=f"PVD_Engineering_Report_S{S}m_{target_day}days.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
