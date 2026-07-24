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
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. HEADER
# ---------------------------------------------------------
st.markdown('<div class="main-header">🏗️ PVD Consolidation Analytics</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">ระบบจำลองการทรุดตัวและคำนวณเวลาการอัดตัวคายน้ำชั้นดินอ่อน (Barron & Terzaghi Theory)</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# 3. SIDEBAR - INPUT PARAMETERS
# ---------------------------------------------------------
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

u90_idx = df[df["U_av"] >= 90].first_valid_index()
days_90 = df.loc[u90_idx, "Day"] if u90_idx is not None else "> 365"

# ---------------------------------------------------------
# FUNCTION: สแกนสร้างรายงาน WORD (.DOCX)
# ---------------------------------------------------------
def generate_word_report():
    doc = Document()
    
    # หัวข้อเอกสาร
    title = doc.add_heading('รายงานสรุปผลการวิเคราะห์การติดตั้ง PVD', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph('รายงานฉบับนี้สร้างขึ้นโดยอัตโนมัติจากระบบ PVD Consolidation Analytics เพื่อสรุปผลการคำนวณการประเมินการทรุดตัวและระยะเวลาการอัดตัวคายน้ำของชั้นดินอ่อน')
    
    # ส่วนที่ 1: พารามิเตอร์นำเข้า
    doc.add_heading('1. ข้อมูลพารามิเตอร์การออกแบบ (Design Parameters)', level=1)
    p_param = doc.add_paragraph()
    p_param.add_run(f'• รูปแบบการติดตั้ง: {pattern}\n')
    p_param.add_run(f'• ระยะห่างการติดตั้ง PVD (S): {S:.2f} เมตร\n')
    p_param.add_run(f'• ความลึกแผ่น PVD (H): {H_pvd:.2f} เมตร (ขนาดแผ่น {a*1000:.0f}x{b*1000:.0f} mm)\n')
    p_param.add_run(f'• ความหนาชั้นดินอ่อน (H_soil): {H_soil:.2f} เมตร\n')
    p_param.add_run(f'• สัมประสิทธิ์การอัดตัวคายน้ำแนวดิ่ง (Cv): {Cv:.2f} m²/ปี, แนวรัศมี (Cr): {Cr:.2f} m²/ปี\n')
    p_param.add_run(f'• การคิดผลกระทบ Smear Effect: {"คิดคำนวณ" if include_smear else "ไม่นำมาคิด"}')
    
    # ส่วนที่ 2: สรุปผลวิเคราะห์คำพูด
    doc.add_heading('2. สรุปผลการคำนวณและการวิเคราะห์ทางวิศวกรรม', level=1)
    
    # คำพูดวิเคราะห์แบบร้อยแก้ว
    analysis_text = (
        f"จากการวิเคราะห์ทางปฐพีวิศวกรรมโดยใช้ทฤษฎีของ Barron, Terzaghi และ Carillo พบว่า "
        f"ขนาดเส้นผ่านศูนย์กลางเทียบเท่าของ PVD (dw) มีค่าเท่ากับ {d_w*1000:.1f} มิลลิเมตร "
        f"และมีระยะอิทธิพลการระบายน้ำ (de) เท่ากับ {d_e:.2f} เมตร "
        f"เมื่อได้รับน้ำหนักถมเพิ่มจำนวน {delta_sigma:.2f} kPa ชั้นดินเหนียวอ่อนจะเกิดการทรุดตัวสูงสุดขั้นปฐมภูมิ "
        f"(Primary Consolidation Settlement, S_final) รวมทั้งสิ้น {S_final:.3f} เมตร (ประมาณ {S_final*100:.1f} เซนติเมตร)\n\n"
    )
    
    if isinstance(days_90, (int, np.integer)):
        analysis_text += (
            f"ในส่วนของอัตราเร่งการระบายน้ำ การติดตั้ง PVD ที่ระยะห่าง {S:.2f} เมตร สามารถเร่งการอัดตัวคายน้ำ "
            f"ให้บรรลุเป้าหมาย 90% (U_av = 90%) ได้ภายในระยะเวลา {days_90} วัน "
            f"โดย ณ วันดังกล่าว ชั้นดินจะเกิดการทรุดตัวไปแล้วประมาณ {S_final*0.9:.3f} เมตร "
            f"ซึ่งถือว่าอยู่ในเกณฑ์ที่มีประสิทธิภาพสำหรับการก่อสร้าง"
        )
    else:
        analysis_text += (
            f"อย่างไรก็ตาม เมื่อพิจารณาในระยะเวลา 1 ปี (365 วัน) พบว่า อัตราการอัดตัวคายน้ำรวมยังไม่สามารถบรรลุเป้าหมาย 90% ได้ "
            f"(ทำได้เพียง {df.iloc[-1]['U_av']:.1f}% ณ วันที่ 365) "
            f"ดังนั้น ทางวิศวกรผู้ออกแบบควรพิจารณาปรับลดระยะห่างการติดตั้ง PVD (S) ให้ถี่ขึ้น หรือเพิ่มน้ำหนักกดทับชั่วคราว (Preloading/Surcharge)"
        )
        
    doc.add_paragraph(analysis_text)
    
    # ส่วนที่ 3: ตารางสรุปเวลาสำคัญ
    doc.add_heading('3. ตารางสรุปพัฒนาการการทรุดตัวตามช่วงเวลา', level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'เวลา (วัน)'
    hdr_cells[1].text = 'U_v (แนวดิ่ง %)'
    hdr_cells[2].text = 'U_r (แนวรัศมี %)'
    hdr_cells[3].text = 'U_av (รวม %)'
    hdr_cells[4].text = 'การทรุดตัว (เมตร)'
    
    # ดึงข้อมูลวันที่ 30, 60, 90, 180, 270, 365 มาแสดง
    target_days = [30, 60, 90, 180, 270, 365]
    for d in target_days:
        row_data = df[df["Day"] == d]
        if not row_data.empty:
            r = row_data.iloc[0]
            row_cells = table.add_row().cells
            row_cells[0].text = str(int(r['Day']))
            row_cells[1].text = f"{r['U_v']:.1f}%"
            row_cells[2].text = f"{r['U_r']:.1f}%"
            row_cells[3].text = f"{r['U_av']:.1f}%"
            row_cells[4].text = f"{r['Settlement']:.3f}"
            
    # Save ลง memory buffer
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ---------------------------------------------------------
# 5. DASHBOARD DISPLAY & DOWNLOAD BUTTON
# ---------------------------------------------------------
m1, m2, m3, m4 = st.columns(4)
m1.metric("เส้นผ่านศูนย์กลางเทียบเท่า (dw)", f"{d_w*1000:.1f} mm")
m2.metric("ระยะอิทธิพลการระบาย (de)", f"{d_e:.2f} m")
m3.metric("การทรุดตัวสูงสุด (S final)", f"{S_final:.3f} m")
m4.metric("เวลาบรรลุ U = 90%", f"{days_90} วัน", 
          delta="ตามเป้าหมาย" if isinstance(days_90, (int, np.integer)) else "ช้าเกินไป", 
          delta_color="normal" if isinstance(days_90, (int, np.integer)) else "inverse")

st.markdown("<br>", unsafe_allow_html=True)

tab_charts, tab_data, tab_summary = st.tabs(["📊 กราฟวิเคราะห์ (Interactive Charts)", "📋 ตารางข้อมูล (Data Table)", "💡 สรุปผลการออกแบบ & Export Report"])

with tab_charts:
    col_left, col_right = st.columns(2)
    with col_left:
        fig_settle = go.Figure()
        fig_settle.add_trace(go.Scatter(x=df["Day"], y=df["Settlement"], mode='lines', name='Settlement (m)', line=dict(color='#EF4444', width=3)))
        fig_settle.update_layout(title="📉 กราฟพัฒนาการการทรุดตัวตามเวลา", xaxis_title="เวลา (วัน)", yaxis_title="การทรุดตัว (เมตร)", yaxis=dict(autorange="reversed"), template="plotly_white", height=380)
        st.plotly_chart(fig_settle, use_container_width=True)

    with col_right:
        fig_u = go.Figure()
        fig_u.add_trace(go.Scatter(x=df["Day"], y=df["U_av"], name='รวม (U_av)', line=dict(color='#10B981', width=3)))
        fig_u.add_trace(go.Scatter(x=df["Day"], y=df["U_r"], name='แนวรัศมี PVD (U_r)', line=dict(color='#3B82F6', dash='dash')))
        fig_u.add_trace(go.Scatter(x=df["Day"], y=df["U_v"], name='แนวดิ่ง ดิน (U_v)', line=dict(color='#9CA3AF', dash='dot')))
        fig_u.add_hline(y=90, line_dash="dash", line_color="#F59E0B", annotation_text="Target 90%")
        fig_u.update_layout(title="📊 อัตราการอัดตัวคายน้ำ (Degree of Consolidation)", xaxis_title="เวลา (วัน)", yaxis_title="Consolidation (%)", yaxis=dict(range=[0, 105]), template="plotly_white", height=380)
        st.plotly_chart(fig_u, use_container_width=True)

with tab_data:
    st.subheader("ตารางแสดงค่าการคำนวณรายวัน")
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("📥 ดาวน์โหลดข้อมูลเป็น CSV", data=csv, file_name="pvd_consolidation_result.csv", mime="text/csv")
    st.dataframe(df.style.format({"U_v": "{:.2f}%", "U_r": "{:.2f}%", "U_av": "{:.2f}%", "Settlement": "{:.4f} m"}), use_container_width=True, height=300)

with tab_summary:
    st.success(f"**สรุปผลการประเมิน:** ในระยะติดตั้ง PVD ที่ **{S:.2f} เมตร** รูปแบบ **{pattern}**")
    st.write(f"- การทรุดตัวทั้งหมดเมื่อสิ้นสุดกระบวนการ ($S_{{final}}$): **{S_final:.3f} เมตร**")
    if isinstance(days_90, (int, np.integer)):
        st.write(f"- ดินจะทรุดตัวถึง 90% ภายในเวลา **{days_90} วัน**")
    else:
        st.warning("- การติดตั้ง PVD ระยะนี้ยังไม่สามารถทำให้ดินทรุดตัวถึง 90% ได้ภายใน 1 ปี แนะนำให้ **ลดระยะห่าง S** ลง")
    
    st.markdown("---")
    st.subheader("📄 ส่งออกรายงานสรุปสำหรับใช้งาน (Export Word Report)")
    st.write("คลิกปุ่มด้านล่างเพื่อสร้างไฟล์รายงานสรุปผลการวิเคราะห์เป็นข้อความร้อยแก้วพร้อมตารางสำเร็จรูป นำไปใช้ทำรายงานต่อใน Microsoft Word ได้ทันที")
    
    # ปุ่มสำหรับดาวน์โหลดรายงาน WORD
    word_file = generate_word_report()
    st.download_button(
        label="📝 ดาวน์โหลดรายงานสรุป (.docx)",
        data=word_file,
        file_name=f"PVD_Report_S_{S:.1f}m.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
