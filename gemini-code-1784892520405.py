with tab_steps:
    st.markdown("### 📑 ตารางแสดงวิธีทำและขั้นตอนการคำนวณ (Step-by-Step Table Calculation)")
    st.caption("จำลองตารางการคำนวณทีละพจน์ตามวิธีของ Barron (1948), Terzaghi และ Carillo (1942)")

    # =========================================================
    # STEP 5: ตารางคำนวณหาค่า Drain Spacing Factor, F(n)
    # =========================================================
    st.markdown("---")
    st.markdown("#### ⑤ ตารางคำนวณเพื่อหาค่า Drain Spacing Factor: $F(n)$")
    
    # สร้างอัตราส่วน S หลายๆ ระยะเพื่อแสดงเปรียบเทียบในตาราง (เช่น S-0.1, S, S+0.1)
    s_test_list = [max(0.5, S - 0.1), S, S + 0.1]
    fn_rows = []
    
    for s_val in s_test_list:
        de_val = (1.05 * s_val if "สามเหลี่ยม" in pattern else 1.13 * s_val) * 100 # แปลงเป็น cm
        dw_cm = d_w * 100
        n_val = de_val / dw_cm
        n2 = n_val**2
        term1 = n2 / (n2 - 1)
        term2 = np.log(n_val)
        term3 = (3 * n2 - 1) / (4 * n2)
        fn_val = (term1 * term2) - term3
        
        fn_rows.append({
            "(1) S (m)": f"{s_val:.2f}",
            "(2) de (cm)": f"{de_val:.1f}",
            "(3) n": f"{n_val:.2f}",
            "(4) n²": f"{n2:.2f}",
            "(5) n²/(n²-1)": f"{term1:.3f}",
            "(6) ln(n)": f"{term2:.3f}",
            "(7) (3n²-1)/(4n²)": f"{term3:.3f}",
            "(8) F(n) = (5)×(6)-(7)": f"★ {fn_val:.3f}" if s_val == S else f"{fn_val:.3f}"
        })
        
    st.table(pd.DataFrame(fn_rows))
    st.info(f"💡 **สำหรับงานออกแบบของคุณ (ที่ระยะ S = {S:.2f} m):** คำนวณค่า **$F(n) = {Fn:.3f}$**")

    # =========================================================
    # STEP 6: ตารางคำนวณหาค่า Degree of Consolidation ในแนวรัศมี (Ur)
    # =========================================================
    st.markdown("---")
    st.markdown("#### ⑥ ตารางคำนวณเพื่อหาค่า Degree of Consolidation ในแนวรัศมี ($U_r$)")
    
    # แปลงหน่วยให้ตรงกับในสไลด์ (cm, day)
    de_cm = d_e * 100
    de2_cm2 = de_cm**2
    Cr_day = Cr / 365.25          # m²/day
    Cr_cm2_day = Cr_day * 10000   # cm²/day
    
    # ดึงเวลาเป้าหมายมาแสดงในตาราง (30, 60, 90, 180 วัน)
    test_days = [30, 60, 90, 180]
    ur_rows = []
    
    for d_val in test_days:
        cr_t = Cr_cm2_day * d_val
        Tr_val = cr_t / de2_cm2
        tr_fn = Tr_val / Fn
        exp_term = np.exp((-8 * Tr_val) / Fn)
        ur_val = 1 - exp_term
        
        ur_rows.append({
            "(1) S (m)": f"{S:.2f}",
            "(2) Cr (cm²/day)": f"{Cr_cm2_day:.1f}",
            "(3) t (day)": str(d_val),
            "(4) de² (cm²)": f"{de2_cm2:.0f}",
            "(5) Cr × t": f"{cr_t:.0f}",
            "(6) Tr = (5)/(4)": f"{Tr_val:.3f}",
            "(7) F(n)": f"{Fn:.2f}",
            "(8) exp(-8Tr/Fn)": f"{exp_term:.4f}",
            "(9) Ur (%) = 1 - (8)": f"★ {ur_val*100:.2f}%" if d_val == 90 else f"{ur_val*100:.2f}%"
        })
        
    st.table(pd.DataFrame(ur_rows))

    # =========================================================
    # STEP 7: คำนวณระดับการอัดตัวคายน้ำในแนวดิ่ง (Uv)
    # =========================================================
    st.markdown("---")
    st.markdown("#### ⑦ คำนวณระดับการอัดตัวคายน้ำในแนวดิ่ง ($U_v$) - Theory of Terzaghi")
    
    col_v1, col_v2 = st.columns([1, 2])
    with col_v1:
        st.latex(r"T_v = \frac{C_v \times t}{(H_d)^2}")
        st.latex(r"U_v = \frac{\sqrt{4 \times T_v}}{\pi} \quad (\text{เมื่อ } U_v \le 60\%)")
        
    with col_v2:
        Cv_cm2_day = (Cv / 365.25) * 10000
        Hd_cm = (H_soil / 2.0) * 100 # ระบายน้ำ 2 ทาง
        t_sample = 90
        Tv_sample = (Cv_cm2_day * t_sample) / (Hd_cm**2)
        Uv_sample = np.sqrt((4 * Tv_sample) / np.pi) if Tv_sample <= 0.286 else 1 - (10**(-0.085 - 0.933 * Tv_sample))
        
        st.write(f"**ตัวอย่างการแทนค่าที่เวลา $t = {t_sample}$ วัน:**")
        st.write(f"- $C_v = {Cv_cm2_day:.1f}\text{{ cm}}^2/\text{{day}}$, ระยะระบายน้ำ $H_d = {Hd_cm:.0f}\text{{ cm}}$")
        st.write(f"- ค่า Time Factor: $T_v = \\frac{{{Cv_cm2_day:.1f} \\times {t_sample}}}{{({Hd_cm:.0f})^2}} = \\mathbf{{{Tv_sample:.4f}}}$")
        st.write(f"- **ดังนั้นได้ค่า $U_v$:** $\\mathbf{{{Uv_sample*100:.2f}\%}}$")

    # =========================================================
    # STEP 8 & 9: ทฤษฎีของ Carillo และตรวจระดับการอัดตัวคายน้ำรวม (Uav)
    # =========================================================
    st.markdown("---")
    st.markdown("#### ⑧-⑨ คำนวณระดับการอัดตัวคายน้ำเฉลี่ยรวม ($U_{av}$) - Theory of Carillo (1942)")
    st.latex(r"U_{av} = 1 - (1 - U_r)(1 - U_v)")
    
    # สร้างตารางสรุปผลรวมตามรูปแบบในสไลด์
    summary_rows = []
    for d_val in [30, 60, 90, 180, 270, 365]:
        r = df[df["Day"] == d_val].iloc[0]
        ur_pct = r['U_r']
        uv_pct = r['U_v']
        uav_pct = r['U_av']
        
        # เช็คเงื่อนไข Step 9 (> 90%)
        status = "✅ ผ่านเกณฑ์ (> 90%)" if uav_pct >= 90 else "⏳ ยังไม่ถึงเกณฑ์"
        
        summary_rows.append({
            "เวลา t (วัน)": f"{d_val} วัน",
            "U_r แนวรัศมี (จากข้อ ⑥)": f"{ur_pct:.2f}%",
            "U_v แนวดิ่ง (จากข้อ ⑦)": f"{uv_pct:.2f}%",
            "⑧ U_av รวม = 1 - (1-Ur)(1-Uv)": f"★ {uav_pct:.2f}%" if d_val == 90 else f"{uav_pct:.2f}%",
            "⑨ ตรวจระดับ U_av > 90%": status
        })
        
    st.table(pd.DataFrame(summary_rows))
    
    # กล่องข้อความสรุปผลลัพธ์ Step 9
    if isinstance(days_90, (int, np.integer)):
        st.success(f"🎯 **บทสรุปการตรวจสอบ (Step ⑨):** ที่ระยะการติดตั้ง **S = {S:.2f} เมตร** ระบบสามารถอัดตัวคายน้ำบรรลุเป้าหมาย **$U_{{av}} > 90\%$ ได้ในวันที่ {days_90}**")
    else:
        st.error(f"⚠️ **บทสรุปการตรวจสอบ (Step ⑨):** ที่ระยะการติดตั้ง **S = {S:.2f} เมตร** ระบายน้ำได้สูงสุดเพียง **{df.iloc[-1]['U_av']:.2f}% ในเวลา 1 ปี** (ยังไม่ผ่านเกณฑ์ > 90%) แนะนำให้ลดระยะห่าง S ลงครับ")
