import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

st.set_page_config(page_title="Phiếu đánh giá giảng viên 2026", page_icon="📋", layout="wide")
st.markdown("""
<style>
.block-container{max-width:1120px;padding-top:1.5rem}
.section{background:#eaf2f8;padding:10px 14px;border-left:5px solid #2c6e9b;border-radius:4px;margin:18px 0 10px;font-weight:700}
.input-note{background:#fff8df;padding:9px 12px;border-radius:5px;border:1px solid #f0dfa3}
.result{background:#e8f5e9;border:1px solid #b7dfba;border-radius:8px;padding:14px}
.metric{background:#f1f7fb;border-radius:8px;padding:12px;text-align:center}
.small{font-size:0.85rem;color:#5f6b72}
</style>
""", unsafe_allow_html=True)

DEPARTMENTS = [
"Khoa Toán - Tin","Khoa Lý - Hóa","Khoa Sinh - Nông nghiệp - Môi trường",
"Khoa Ngữ văn - Truyền thông","Khoa Sử - Địa - Chính trị",
"Khoa Tâm lý - Giáo dục - CTXH","Khoa GD Tiểu học - Mầm non",
"Khoa GD Nghệ thuật - Thể chất"
]
GROUPS=["GV thường","GV QL/kiêm nhiệm"]
RANKS=["GS/PGS","TS","ĐH/ThS"]
STATUSES=["Bình thường","Năm đầu sau tập sự/chuyển về","Tập sự"]
YESNO=["Không","Có"]
PASS=["Đạt","Không đạt"]

def pct(a,b,zero_value=0):
    return zero_value if b == 0 else a/b*100

def evaluate(x):
    # Bám theo công thức trong DGXL_GV của workbook được cung cấp.
    gd_target = 0 if x["status"]=="Tập sự" else x["gd_base"]*(0.5 if x["status"]=="Năm đầu sau tập sự/chuyển về" else 1)
    gd_pct = 100 if gd_target==0 else pct(x["gd_done"],gd_target)
    if x["group"]=="GV thường":
        research_standard={"GS/PGS":2.0,"TS":1.5,"ĐH/ThS":1.0}[x["rank"]]
    else:
        g=x["research_share"]
        if g<35: factors={"GS/PGS":0.7,"TS":0.5,"ĐH/ThS":0.35}
        elif g<=70: factors={"GS/PGS":1.0,"TS":0.75,"ĐH/ThS":0.5}
        else: factors={"GS/PGS":2.0,"TS":1.5,"ĐH/ThS":1.0}
        research_standard=factors[x["rank"]]
    research_target=0 if x["status"]=="Tập sự" else research_standard*(0.5 if x["status"]=="Năm đầu sau tập sự/chuyển về" else 1)*min(max(x["research_months"],0),12)/12
    research_pct=100 if research_target==0 else pct(x["research_done"],research_target)
    nvk_target=350*(0.5 if x["status"]=="Năm đầu sau tập sự/chuyển về" else 1)
    nvk_pct=0 if nvk_target<=0 else min(pct(x["nvk_done"],nvk_target),100)
    total=gd_pct*.4+research_pct*.4+nvk_pct*.2

    research_threshold_met=(research_target==0 or x["research_done"]>=research_target)
    if x["group"]=="GV QL/kiêm nhiệm" or x["rank"]=="ĐH/ThS":
        extra=bool(x["criteria_item_1_6"] or x["student_award"])
    else:
        extra=any([x["q1q2"],x["supervision_q1q2"],x["student_award"],x["patent"],x["criteria_item_3_5"]])

    eligible = x["work_months"]>=6 or x["maternity"]=="Có"
    if not eligible:
        base="KHÔNG ĐÁNH GIÁ"
    elif (x["teaching_quality"]!="Đạt" or x["common_criteria"]!="Đạt" or total<50
          or (gd_target>0 and x["gd_done"]==0)
          or (research_target>0 and x["research_done"]==0)
          or (nvk_target>0 and x["nvk_done"]==0)):
        base="KHTNV"
    elif (gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and total>=150
          and research_threshold_met and extra):
        base="HTXSNV"
    elif (gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and total>=100 and research_threshold_met):
        base="HTTNV"
    elif total>=50 and nvk_pct>=50:
        base="HTNV"
    else:
        base="KHTNV"

    if base=="KHÔNG ĐÁNH GIÁ":
        final=base
    elif x["discipline"]=="Có":
        final="KHTNV"
    else:
        level={"HTXSNV":1,"HTTNV":2,"HTNV":3,"KHTNV":4}[base]
        level=min(4,level+(3 if x["violation_points"]>=30 else 2 if x["violation_points"]>=20 else 1 if x["violation_points"]>=10 else 0))
        if 3<=x["leave_months"]<6 and x["maternity"]!="Có":
            level=max(level,3)
        final={1:"HTXSNV",2:"HTTNV",3:"HTNV",4:"KHTNV"}[level]

    alerts=[]
    if not eligible: alerts.append("Không đủ 06 tháng công tác (trừ thai sản).")
    if x["teaching_quality"]!="Đạt": alerts.append("Chất lượng GD không đạt.")
    if x["common_criteria"]!="Đạt": alerts.append("Vi phạm tiêu chí chung.")
    if research_target>0 and x["research_done"]<research_target: alerts.append("NCKH chưa đạt ngưỡng HTT theo QĐ 4882.")
    if total>=150 and not extra: alerts.append("Thiếu điều kiện bổ sung NCKH để HTXSNV.")
    if x["violation_points"]>=10: alerts.append("Bị khống chế do điểm vi phạm.")
    if x["discipline"]=="Có": alerts.append("Có quyết định kỷ luật trong năm.")
    if 3<=x["leave_months"]<6 and x["maternity"]!="Có": alerts.append("Nghỉ 03-<06 tháng: tối đa HTNV.")

    tntt=gd_pct*.5+research_pct*.5
    if (x["teaching_quality"]!="Đạt" or x["common_criteria"]!="Đạt"
        or (gd_target>0 and x["gd_done"]==0)
        or (research_target>0 and x["research_done"]==0)
        or (nvk_target>0 and x["nvk_done"]==0) or tntt<50):
        tntt_grade="D"
    elif (gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and tntt>=400 and final=="HTXSNV" and x["nvk_done"]>=nvk_target):
        tntt_grade="A"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and 300<=tntt<400 and x["nvk_done"]>=nvk_target:
        tntt_grade="B1"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and 200<=tntt<300 and x["nvk_done"]>=nvk_target:
        tntt_grade="B2"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and 100<=tntt<200 and x["nvk_done"]>=nvk_target:
        tntt_grade="B3"
    elif gd_pct>0 and research_pct>0 and 50<=tntt<100 and x["teaching_quality"]=="Đạt" and x["common_criteria"]=="Đạt" and x["nvk_done"]>=nvk_target*.5:
        tntt_grade="C"
    else:
        tntt_grade="CẦN XEM XÉT"
    tntt_alerts=[]
    if tntt_grade=="CẦN XEM XÉT":
        tntt_alerts.append("Không thuộc trọn vẹn tập điều kiện A/B1/B2/B3/C/D theo Điều 18; cần đơn vị/Hội đồng xem xét.")
    if tntt>=400 and final!="HTXSNV":
        tntt_alerts.append("Tổng TNTT >=400% nhưng xếp loại VC cuối không phải HTXSNV nên không đủ A.")
    if tntt>=100 and 0<gd_pct<100: tntt_alerts.append("Giảng dạy chưa đạt 100% định mức.")
    if tntt>=100 and 0<research_pct<100: tntt_alerts.append("NCKH chưa đạt 100% định mức/ngưỡng áp dụng.")
    if tntt>=100 and 0<nvk_pct<100: tntt_alerts.append("NVK chưa đạt 100% định mức.")
    if 50<=tntt<100 and 0<x["nvk_done"]<nvk_target*.5: tntt_alerts.append("NVK chưa đạt 50% định mức để xếp C.")

    return dict(gd_target=gd_target,gd_pct=gd_pct,research_standard=research_standard,
        research_target=research_target,research_pct=research_pct,nvk_target=nvk_target,
        nvk_pct=nvk_pct,total=total,research_threshold_met=research_threshold_met,
        extra=extra,eligible=eligible,base=base,final=final,alerts=alerts,
        tntt=tntt,tntt_grade=tntt_grade,tntt_alerts=tntt_alerts)

st.title("📋 PHIẾU ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG GIẢNG VIÊN NĂM 2026")
st.markdown("<div style='text-align:center;color:#566573'>TRƯỜNG ĐẠI HỌC SƯ PHẠM – ĐẠI HỌC ĐÀ NẴNG<br/>Dành cho viên chức ngạch giảng viên</div>",unsafe_allow_html=True)
st.info("Nhập thông tin vào các ô màu trắng. Các chỉ tiêu nền xanh được ứng dụng tự động tính lại ngay khi bạn thay đổi dữ liệu. Mỗi lần mở ứng dụng là một phiếu cá nhân riêng.")

st.markdown('<div class="section">I. THÔNG TIN CÁ NHÂN</div>',unsafe_allow_html=True)
c1,c2,c3=st.columns([2,1,2])
with c1: name=st.text_input("Họ và tên *")
with c2: staff_id=st.text_input("Mã cán bộ")
with c3: department=st.selectbox("Đơn vị",DEPARTMENTS)
c1,c2,c3=st.columns(3)
with c1: rank=st.selectbox("Chức danh / trình độ",RANKS)
with c2: position=st.text_input("Chức vụ (nếu có)")
with c3: group=st.selectbox("Nhóm NCKH",GROUPS)
c1,c2,c3=st.columns(3)
with c1: status=st.selectbox("Trạng thái định mức",STATUSES)
with c2: work_months=st.number_input("Số tháng công tác trong năm",0,12,12)
with c3: research_months=st.number_input("Số tháng tính NCKH",0,12,12)
c1,c2=st.columns(2)
with c1: maternity=st.selectbox("Nghỉ thai sản trong năm?",YESNO)
with c2: leave_months=st.number_input("Số tháng nghỉ không tham gia công tác",0.0,12.0,0.0,0.5)

st.markdown('<div class="section">II. KẾT QUẢ THỰC HIỆN CHỨC TRÁCH, NHIỆM VỤ</div>',unsafe_allow_html=True)
st.markdown("**1. Giảng dạy (40%)**")
c1,c2,c3=st.columns(3)
with c1: gd_base=st.number_input("Định mức giảng dạy gốc",0.0,10000.0,270.0,10.0)
with c2: 
    gd_done=st.number_input("Khối lượng giảng dạy thực hiện",0.0,100000.0,0.0,1.0)
with c3: st.metric("Định mức giảng dạy áp dụng",f"{(0 if status=='Tập sự' else gd_base*(.5 if status=='Năm đầu sau tập sự/chuyển về' else 1)):,.2f}")

st.markdown("**2. Nghiên cứu khoa học (40%)**")
c1,c2,c3=st.columns(3)
with c1: research_share=st.number_input("Tỷ lệ định mức NCKH QL/kiêm nhiệm (%)",0.0,100.0,100.0,1.0,disabled=(group=="GV thường"))
with c2: research_done=st.number_input("Điểm NCKH thực hiện (quy đổi Bảng 1 PL02)",0.0,1000.0,0.0,0.1)
with c3: st.caption("Ngưỡng NCKH được xác định tự động theo nhóm NCKH, chức danh và số tháng tính NCKH.")

st.markdown("**3. Nhiệm vụ khác (20%)**")
c1,c2,c3=st.columns(3)
with c1: nvk_done=st.number_input("Số giờ nhiệm vụ khác thực hiện",0.0,100000.0,0.0,1.0)
with c2: st.metric("Định mức NVK áp dụng",f"{350*(.5 if status=='Năm đầu sau tập sự/chuyển về' else 1):,.0f} giờ")
with c3: st.caption("Tỷ lệ NVK dùng tính tổng được giới hạn tối đa 100%.")

st.markdown('<div class="section">III. CHẤT LƯỢNG, TIÊU CHÍ BỔ SUNG VÀ VI PHẠM</div>',unsafe_allow_html=True)
c1,c2=st.columns(2)
with c1: teaching_quality=st.selectbox("Chất lượng giảng dạy",PASS)
with c2: common_criteria=st.selectbox("Tiêu chí chung",PASS)
st.markdown("**Điều kiện bổ sung NCKH để xét HTXSNV**")
c1,c2,c3=st.columns(3)
with c1: q1q2=st.checkbox("Có bài WoS/Scopus Q1/Q2")
with c2: supervision_q1q2=st.checkbox("Hướng dẫn SV/HV có bài Q1/Q2")
with c3: student_award=st.checkbox("Hướng dẫn SV NCKH đạt giải 1–3 cấp Bộ/tỉnh/thành phố trở lên")
c1,c2=st.columns(2)
with c1: patent=st.checkbox("Có bằng độc quyền sáng chế")
with c2: criteria_item_3_5=st.checkbox("Có công trình mục 3 hoặc 5 Bảng 1")
criteria_item_1_6=st.checkbox("Có công trình thuộc mục 1–6 Bảng 1")
st.markdown("**Thông tin vi phạm**")
violation_points=st.number_input("Tổng điểm vi phạm",0.0,1000.0,0.0,1.0)
discipline=st.selectbox("Có quyết định kỷ luật trong năm?",YESNO)
with st.expander("Nhập chi tiết các nội dung vi phạm (không bắt buộc)"):
    violations=st.data_editor(pd.DataFrame([{"Nội dung vi phạm":"","Điểm":0.0,"Ghi chú":""} for _ in range(3)]),
        num_rows="dynamic",use_container_width=True,key="violations")
    if not violations.empty and st.button("Tính tổng điểm từ bảng vi phạm"):
        st.session_state["violation_points_from_table"]=float(pd.to_numeric(violations["Điểm"],errors="coerce").fillna(0).sum())
        st.rerun()
    if "violation_points_from_table" in st.session_state:
        st.caption(f"Tổng điểm theo bảng chi tiết: {st.session_state['violation_points_from_table']:g} (hãy cập nhật vào ô Tổng điểm vi phạm phía trên).")

x=dict(status=status,gd_base=gd_base,gd_done=gd_done,group=group,rank=rank,research_share=research_share,
       research_months=research_months,research_done=research_done,nvk_done=nvk_done,
       work_months=work_months,maternity=maternity,leave_months=leave_months,
       teaching_quality=teaching_quality,common_criteria=common_criteria,
       q1q2=q1q2,supervision_q1q2=supervision_q1q2,student_award=student_award,
       patent=patent,criteria_item_3_5=criteria_item_3_5,criteria_item_1_6=criteria_item_1_6,
       violation_points=violation_points,discipline=discipline)
r=evaluate(x)

st.markdown('<div class="section">IV. KẾT QUẢ TÍNH TOÁN TỰ ĐỘNG</div>',unsafe_allow_html=True)
if not name.strip():
    st.warning("Vui lòng nhập họ và tên để hoàn thiện phiếu.")
cols=st.columns(4)
metrics=[("Định mức GD áp dụng",f"{r['gd_target']:,.2f}"),("% Giảng dạy",f"{r['gd_pct']:.2f}%"),
         ("Ngưỡng NCKH áp dụng",f"{r['research_target']:.3f}"),("% NCKH",f"{r['research_pct']:.2f}%")]
for col,(label,value) in zip(cols,metrics):
    with col: st.metric(label,value)
cols=st.columns(4)
metrics=[("Định mức NVK",f"{r['nvk_target']:.0f} giờ"),("% NVK",f"{r['nvk_pct']:.2f}%"),
         ("Tổng đánh giá 40-40-20",f"{r['total']:.2f}%"),("Tổng TNTT 50-50",f"{r['tntt']:.2f}%")]
for col,(label,value) in zip(cols,metrics):
    with col: st.metric(label,value)
st.markdown("#### Kết quả xếp loại")
a,b=st.columns(2)
with a:
    st.success(f"**Xếp loại viên chức: {r['final']}**")
    st.caption(f"Xếp loại cơ sở trước khi xét khống chế: {r['base']}")
with b:
    if r["tntt_grade"]=="CẦN XEM XÉT": st.warning(f"**Xếp thu nhập tăng thêm: {r['tntt_grade']}**")
    else: st.success(f"**Xếp thu nhập tăng thêm: {r['tntt_grade']}**")
st.markdown("#### Kiểm tra điều kiện")
c1,c2,c3=st.columns(3)
c1.write("**Đủ ngưỡng NCKH HTT:** "+("Có" if r["research_threshold_met"] else "Không"))
c2.write("**Đủ điều kiện bổ sung HTX:** "+("Có" if r["extra"] else "Không"))
c3.write("**Đủ điều kiện đánh giá:** "+("Có" if r["eligible"] else "Không"))
if r["alerts"]:
    st.warning("**Cảnh báo xếp loại viên chức**\n\n"+"\n\n".join("• "+a for a in r["alerts"]))
if r["tntt_alerts"]:
    st.warning("**Cảnh báo thu nhập tăng thêm**\n\n"+"\n\n".join("• "+a for a in r["tntt_alerts"]))
if not r["alerts"] and not r["tntt_alerts"]:
    st.success("Chưa phát hiện cảnh báo theo các điều kiện được mã hóa từ workbook.")

st.markdown('<div class="section">V. KẾT QUẢ TỰ ĐÁNH GIÁ VÀ XÁC NHẬN</div>',unsafe_allow_html=True)
self_comment=st.text_area("Tự nhận xét, đánh giá",height=100)
date_value=st.date_input("Ngày lập phiếu",value=date.today())
st.caption("Sau khi hoàn tất, bạn có thể tải phiếu Excel để lưu, hoặc in trang này thành PDF bằng chức năng Print của trình duyệt (Ctrl+P).")

record={
"Mã CB":staff_id,"Họ và tên":name,"Đơn vị":department,"Chức vụ":position,
"Nhóm NCKH":group,"Chức danh/trình độ":rank,"Trạng thái định mức":status,
"Số tháng công tác":work_months,"Số tháng tính NCKH":research_months,
"Nghỉ thai sản":maternity,"Số tháng nghỉ không tham gia":leave_months,
"ĐM GD gốc":gd_base,"ĐM GD áp dụng":r["gd_target"],"TH GD":gd_done,"%GD":r["gd_pct"],
"Tỷ lệ ĐM NCKH QL/KN (%)":research_share,"Ngưỡng NCKH chuẩn":r["research_standard"],
"Ngưỡng NCKH áp dụng":r["research_target"],"Điểm NCKH TH":research_done,"%NCKH":r["research_pct"],
"ĐM NVK áp dụng":r["nvk_target"],"TH NVK":nvk_done,"%NVK":r["nvk_pct"],
"GD x 0,4":r["gd_pct"]*.4,"NCKH x 0,4":r["research_pct"]*.4,"NVK x 0,2":r["nvk_pct"]*.2,
"Tổng %":r["total"],"Chất lượng GD":teaching_quality,"Tiêu chí chung":common_criteria,
"Có bài Q1/Q2":int(q1q2),"HD SV/HV có Q1/Q2":int(supervision_q1q2),
"HD SV NCKH đạt giải":int(student_award),"Bằng sáng chế":int(patent),
"CT mục 3/5 B1":int(criteria_item_3_5),"CT mục 1-6 B1":int(criteria_item_1_6),
"Đủ ĐK NCKH HTT":"Có" if r["research_threshold_met"] else "Không",
"Đủ ĐK bổ sung HTX":"Có" if r["extra"] else "Không",
"Điểm vi phạm":violation_points,"Có QĐ kỷ luật":discipline,
"Đủ ĐK đánh giá":"Có" if r["eligible"] else "Không","Xếp loại cơ sở":r["base"],
"Xếp loại cuối":r["final"],"Cảnh báo / lý do":" ".join(r["alerts"]),
"Tổng % TNTT (GD 50% + NCKH 50%)":r["tntt"],"Xếp TNTT":r["tntt_grade"],
"Cảnh báo TNTT":" ".join(r["tntt_alerts"]),"Tự nhận xét":self_comment,"Ngày lập phiếu":date_value.isoformat()
}
df=pd.DataFrame([record])
out=BytesIO()
with pd.ExcelWriter(out,engine="openpyxl") as writer:
    df.to_excel(writer,index=False,sheet_name="PHIEU_CA_NHAN")
    pd.DataFrame([{"Mục": "Xếp loại viên chức","Kết quả":r["final"]},
                  {"Mục":"Xếp loại TNTT","Kết quả":r["tntt_grade"]},
                  {"Mục":"Tổng đánh giá 40-40-20 (%)","Kết quả":r["total"]},
                  {"Mục":"Tổng TNTT 50-50 (%)","Kết quả":r["tntt"]},
                  {"Mục":"Cảnh báo","Kết quả":" ".join(r["alerts"]+r["tntt_alerts"])}]).to_excel(writer,index=False,sheet_name="KET_QUA")
    for sheet in writer.book.worksheets:
        sheet.freeze_panes="A2"
        sheet.auto_filter.ref=sheet.dimensions
        for col in sheet.columns:
            maxlen=max((len(str(c.value)) if c.value is not None else 0) for c in col)
            sheet.column_dimensions[col[0].column_letter].width=min(max(maxlen+2,12),42)
st.download_button("⬇️ Tải phiếu đánh giá cá nhân (Excel)",out.getvalue(),
    file_name=f"Phieu_danh_gia_{staff_id or 'giang_vien'}_2026.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",disabled=not name.strip(),use_container_width=True)

with st.expander("Căn cứ và cách tính"):
    st.markdown("""
    - Tổng đánh giá viên chức = %GD × 0,4 + %NCKH × 0,4 + %NVK × 0,2; %NVK tối đa 100%.
    - Tổng thu nhập tăng thêm = %GD × 0,5 + %NCKH × 0,5; NVK là điều kiện khống chế theo từng mức.
    - Ngưỡng NCKH GV thường: GS/PGS 2,0; TS 1,5; ĐH/ThS 1,0.
    - GV QL/kiêm nhiệm áp dụng hệ số theo tỷ lệ định mức NCKH: dưới 35%; từ 35% đến 70%; trên 70%.
    - Các trường hợp TNTT không thuộc trọn vẹn điều kiện A/B1/B2/B3/C/D được hiển thị “CẦN XEM XÉT”.

    Các quy tắc trên được lấy từ sheet HUONG_DAN, DM_THAM_CHIEU, DGXL_GV và TNTT_264 của workbook gốc.
    """)
st.caption("Lưu ý: đây là phiếu tự đánh giá. Kết quả do ứng dụng tính hỗ trợ; đơn vị/Hội đồng có thẩm quyền xác nhận kết quả chính thức.")
