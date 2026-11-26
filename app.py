import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Đánh giá, xếp loại giảng viên 2026", layout="wide")
st.title("Ứng dụng đánh giá, xếp loại chất lượng giảng viên năm 2026")
st.caption("Chuyển đổi từ mẫu Excel người dùng cung cấp. Các kết quả được tính tự động theo các quy tắc và công thức trong workbook.")

DEPARTMENTS = [
    "Khoa Toán - Tin", "Khoa Lý - Hóa",
    "Khoa Sinh - Nông nghiệp - Môi trường",
    "Khoa Ngữ văn - Truyền thông", "Khoa Sử - Địa - Chính trị",
    "Khoa Tâm lý - Giáo dục - CTXH", "Khoa GD Tiểu học - Mầm non",
    "Khoa GD Nghệ thuật - Thể chất"
]
GROUPS = ["GV thường", "GV QL/kiêm nhiệm"]
RANKS = ["GS/PGS", "TS", "ĐH/ThS"]
STATUS = ["Bình thường", "Năm đầu sau tập sự/chuyển về", "Tập sự"]
YESNO = ["Có", "Không"]
PASS = ["Đạt", "Không đạt"]

INPUT_COLS = [
    "Mã CB","Họ và tên","Đơn vị","Nhóm NCKH","Chức danh/trình độ",
    "Tỷ lệ ĐM NCKH QL/KN (%)","Trạng thái định mức","Số tháng công tác",
    "Số tháng tính NCKH","Nghỉ thai sản","Số tháng nghỉ không tham gia",
    "ĐM GD gốc","TH GD","Điểm NCKH TH","TH NVK","Chất lượng GD",
    "Tiêu chí chung","Có bài Q1/Q2","HD SV/HV có Q1/Q2",
    "HD SV NCKH đạt giải","Bằng sáng chế","CT mục 3/5 B1","CT mục 1-6 B1",
    "Điểm vi phạm","Có QĐ kỷ luật"
]
OUTPUT_COLS = [
    "ĐM GD áp dụng","%GD","Ngưỡng NCKH chuẩn","Ngưỡng NCKH áp dụng","%NCKH",
    "ĐM NVK áp dụng","%NVK","GD x 0,4","NCKH x 0,4","NVK x 0,2",
    "Tổng %","Đủ ĐK NCKH HTT","Đủ ĐK bổ sung HTX","Đủ ĐK đánh giá",
    "Xếp loại cơ sở","Xếp loại cuối","Cảnh báo / lý do",
    "Tổng % TNTT (GD 50% + NCKH 50%)","Xếp TNTT","Cảnh báo TNTT"
]
def blank_record():
    return {c: None for c in INPUT_COLS}

if "records" not in st.session_state:
    st.session_state.records = pd.DataFrame([blank_record()])

def num(v, default=0.0):
    try:
        if pd.isna(v) or v == "": return default
        return float(v)
    except Exception: return default

def calculate(row):
    name = str(row.get("Họ và tên") or "").strip()
    if not name:
        return {}
    group = row.get("Nhóm NCKH") or "GV thường"
    rank = row.get("Chức danh/trình độ") or "ĐH/ThS"
    status = row.get("Trạng thái định mức") or "Bình thường"
    g = num(row.get("Tỷ lệ ĐM NCKH QL/KN (%)"),100)
    months_work = num(row.get("Số tháng công tác"),12)
    months_research = num(row.get("Số tháng tính NCKH"),12)
    leave = num(row.get("Số tháng nghỉ không tham gia"),0)
    gd_base = num(row.get("ĐM GD gốc"),100)
    gd_done = num(row.get("TH GD"),0)
    research_done = num(row.get("Điểm NCKH TH"),0)
    nvk_done = num(row.get("TH NVK"),0)
    if status == "Tập sự":
        gd_target = 0
    else:
        gd_target = gd_base * (0.5 if status == "Năm đầu sau tập sự/chuyển về" else 1)
    gd_pct = 100 if gd_target == 0 else gd_done/gd_target*100
    base_by_rank = {"GS/PGS":2,"TS":1.5,"ĐH/ThS":1}
    if group == "GV thường":
        research_standard = base_by_rank.get(rank,1)
    elif group == "GV QL/kiêm nhiệm":
        if g < 35: factors={"GS/PGS":0.7,"TS":0.5,"ĐH/ThS":0.35}
        elif g <= 70: factors={"GS/PGS":1,"TS":0.75,"ĐH/ThS":0.5}
        else: factors=base_by_rank
        research_standard=factors.get(rank,1)
    else:
        research_standard=1
    research_target = 0 if status=="Tập sự" else research_standard*(0.5 if status=="Năm đầu sau tập sự/chuyển về" else 1)*min(max(months_research,0),12)/12
    research_pct = 100 if research_target==0 else research_done/research_target*100
    nvk_target=350*(0.5 if status=="Năm đầu sau tập sự/chuyển về" else 1)
    nvk_pct=0 if nvk_target<=0 else min(nvk_done/nvk_target*100,100)
    total=gd_pct*.4+research_pct*.4+nvk_pct*.2
    tntt=gd_pct*.5+research_pct*.5
    teaching_quality=row.get("Chất lượng GD") or "Đạt"
    common=row.get("Tiêu chí chung") or "Đạt"
    discipline=row.get("Có QĐ kỷ luật") or "Không"
    violation=num(row.get("Điểm vi phạm"),0)
    enough_eval = "Có" if months_work>=6 or row.get("Nghỉ thai sản")=="Có" else "Không"
    special_nkh = research_target==0 or research_done>=research_target
    extra = (num(row.get("CT mục 1-6 B1"),0)==1 or num(row.get("HD SV NCKH đạt giải"),0)==1) if (group=="GV QL/kiêm nhiệm" or rank=="ĐH/ThS") else any(num(row.get(c),0)==1 for c in ["Có bài Q1/Q2","HD SV/HV có Q1/Q2","HD SV NCKH đạt giải","Bằng sáng chế","CT mục 3/5 B1"])
    if enough_eval!="Có":
        base_result="KHÔNG ĐÁNH GIÁ"
    elif (teaching_quality!="Đạt" or common!="Đạt" or total<50 or (gd_target>0 and gd_done==0) or (research_target>0 and research_done==0) or (nvk_target>0 and nvk_done==0)):
        base_result="KHTNV"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and total>=150 and special_nkh and extra:
        base_result="HTXSNV"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and total>=100 and special_nkh:
        base_result="HTTNV"
    elif total>=50 and nvk_pct>=50:
        base_result="HTNV"
    else:
        base_result="KHTNV"
    if base_result=="KHÔNG ĐÁNH GIÁ":
        final_result=base_result
    elif discipline=="Có":
        final_result="KHTNV"
    else:
        # Hạ mức theo điểm vi phạm; nghỉ 3 đến dưới 6 tháng tối đa HTNV
        level={"HTXSNV":1,"HTTNV":2,"HTNV":3,"KHTNV":4}.get(base_result,4)
        level=min(4,level+(3 if violation>=30 else 2 if violation>=20 else 1 if violation>=10 else 0))
        if 3<=leave<6 and row.get("Nghỉ thai sản")!="Có":
            level=max(level,3)
        final_result={1:"HTXSNV",2:"HTTNV",3:"HTNV",4:"KHTNV"}[level]
    reasons=[]
    if enough_eval!="Có": reasons.append("Không đủ 06 tháng công tác (trừ thai sản)")
    if teaching_quality!="Đạt": reasons.append("Chất lượng GD không đạt")
    if common!="Đạt": reasons.append("Vi phạm tiêu chí chung")
    if research_target>0 and research_done<research_target: reasons.append("NCKH chưa đạt ngưỡng HTT theo QĐ 4882")
    if total>=150 and not extra: reasons.append("Thiếu điều kiện bổ sung NCKH để HTXSNV")
    if violation>=10: reasons.append("Bị khống chế do điểm vi phạm")
    if discipline=="Có": reasons.append("Có quyết định kỷ luật trong năm")
    if 3<=leave<6 and row.get("Nghỉ thai sản")!="Có": reasons.append("Nghỉ 03-<06 tháng: tối đa HTNV")
    if nvk_pct>=100: nvk_condition=True
    else: nvk_condition=False
    if teaching_quality!="Đạt" or common!="Đạt" or (gd_target>0 and gd_done==0) or (research_target>0 and research_done==0) or (nvk_target>0 and nvk_done==0) or tntt<50:
        tntt_grade="D"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and tntt>=400 and final_result=="HTXSNV" and nvk_done>=nvk_target:
        tntt_grade="A"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and 300<=tntt<400 and nvk_done>=nvk_target:
        tntt_grade="B1"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and 200<=tntt<300 and nvk_done>=nvk_target:
        tntt_grade="B2"
    elif gd_pct>=100 and research_pct>=100 and nvk_pct>=100 and 100<=tntt<200 and nvk_done>=nvk_target:
        tntt_grade="B3"
    elif gd_pct>0 and research_pct>0 and 50<=tntt<100 and teaching_quality=="Đạt" and common=="Đạt" and nvk_done>=nvk_target*.5:
        tntt_grade="C"
    else:
        tntt_grade="CẦN XEM XÉT"
    tntt_reasons=[]
    if tntt_grade=="CẦN XEM XÉT": tntt_reasons.append("Không thuộc trọn vẹn tập điều kiện A/B1/B2/B3/C/D theo Điều 18; cần đơn vị/Hội đồng xem xét")
    if tntt>=400 and final_result!="HTXSNV": tntt_reasons.append("Tổng TNTT >=400% nhưng xếp loại VC cuối không phải HTXSNV nên không đủ A")
    if tntt>=100 and 0<gd_pct<100: tntt_reasons.append("Giảng dạy chưa đạt 100% định mức")
    if tntt>=100 and 0<research_pct<100: tntt_reasons.append("NCKH chưa đạt 100% định mức/ngưỡng áp dụng")
    if tntt>=100 and 0<nvk_pct<100: tntt_reasons.append("NVK chưa đạt 100% định mức")
    if 50<=tntt<100 and 0<nvk_done<nvk_target*.5: tntt_reasons.append("NVK chưa đạt 50% định mức để xếp C")
    return {
      "ĐM GD áp dụng":gd_target,"%GD":gd_pct,"Ngưỡng NCKH chuẩn":research_standard,
      "Ngưỡng NCKH áp dụng":research_target,"%NCKH":research_pct,
      "ĐM NVK áp dụng":nvk_target,"%NVK":nvk_pct,"GD x 0,4":gd_pct*.4,
      "NCKH x 0,4":research_pct*.4,"NVK x 0,2":nvk_pct*.2,"Tổng %":total,
      "Đủ ĐK NCKH HTT":"Có" if special_nkh else "Không",
      "Đủ ĐK bổ sung HTX":"Có" if extra else "Không",
      "Đủ ĐK đánh giá":enough_eval,"Xếp loại cơ sở":base_result,
      "Xếp loại cuối":final_result,"Cảnh báo / lý do":". ".join(reasons),
      "Tổng % TNTT (GD 50% + NCKH 50%)":tntt,"Xếp TNTT":tntt_grade,
      "Cảnh báo TNTT":". ".join(tntt_reasons)
    }

def make_results(df):
    rows=[]
    for i,(_,r) in enumerate(df.iterrows(),1):
        if str(r.get("Họ và tên") or "").strip():
            d=r.to_dict(); d["STT"]=i; d.update(calculate(d)); rows.append(d)
    return pd.DataFrame(rows)

with st.sidebar:
    st.header("Quản lý dữ liệu")
    st.write("Nhập trực tiếp, hoặc tải lên bảng dữ liệu đã chuẩn bị.")
    uploaded=st.file_uploader("Nhập dữ liệu Excel/CSV",type=["xlsx","csv"])
    if uploaded is not None:
        try:
            if uploaded.name.lower().endswith(".csv"):
                incoming=pd.read_csv(uploaded).reindex(columns=INPUT_COLS)
            else:
                incoming=pd.read_excel(uploaded).reindex(columns=INPUT_COLS)
            st.session_state.records=incoming.fillna("")
            st.success(f"Đã nạp {len(incoming)} dòng.")
        except Exception as e: st.error(f"Không đọc được tệp: {e}")
    if st.button("Thêm 10 dòng trống"):
        st.session_state.records=pd.concat([st.session_state.records,pd.DataFrame([blank_record() for _ in range(10)])],ignore_index=True)
    if st.button("Xóa toàn bộ dữ liệu"):
        st.session_state.records=pd.DataFrame([blank_record() for _ in range(10)])
    st.divider()
    st.caption("Lưu ý: hãy tải xuống bản Excel sau khi nhập liệu để lưu trữ.")

tab1,tab2,tab3=st.tabs(["Nhập liệu","Kết quả cá nhân","Tổng hợp"])
with tab1:
    st.subheader("Danh sách giảng viên – vùng nhập liệu")
    st.info("Chỉ nhập các cột hiển thị. Các cột kết quả được ứng dụng tự tính. Đối với các tiêu chí dạng Có/Không, nhập 1 nếu có và 0 nếu không.")
    edited=st.data_editor(
        st.session_state.records,
        num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
          "Đơn vị":st.column_config.SelectboxColumn(options=DEPARTMENTS,required=False),
          "Nhóm NCKH":st.column_config.SelectboxColumn(options=GROUPS),
          "Chức danh/trình độ":st.column_config.SelectboxColumn(options=RANKS),
          "Trạng thái định mức":st.column_config.SelectboxColumn(options=STATUS),
          "Nghỉ thai sản":st.column_config.SelectboxColumn(options=YESNO),
          "Chất lượng GD":st.column_config.SelectboxColumn(options=PASS),
          "Tiêu chí chung":st.column_config.SelectboxColumn(options=PASS),
          "Có QĐ kỷ luật":st.column_config.SelectboxColumn(options=YESNO),
          "Tỷ lệ ĐM NCKH QL/KN (%)":st.column_config.NumberColumn(min_value=0,max_value=100,step=1),
          "Số tháng công tác":st.column_config.NumberColumn(min_value=0,max_value=12,step=1),
          "Số tháng tính NCKH":st.column_config.NumberColumn(min_value=0,max_value=12,step=1),
        }, key="editor"
    )
    st.session_state.records=edited
    st.caption(f"Số giảng viên đã nhập: {edited['Họ và tên'].fillna('').astype(str).str.strip().ne('').sum()}")

results=make_results(st.session_state.records)
with tab2:
    st.subheader("Kết quả đánh giá từng giảng viên")
    if results.empty: st.warning("Chưa có dữ liệu. Vui lòng nhập họ tên ở tab Nhập liệu.")
    else:
        query=st.text_input("Tìm theo họ tên hoặc mã cán bộ")
        view=results.copy()
        if query:
            mask=view["Họ và tên"].astype(str).str.contains(query,case=False,na=False) | view["Mã CB"].astype(str).str.contains(query,case=False,na=False)
            view=view[mask]
        show=["STT","Mã CB","Họ và tên","Đơn vị","%GD","%NCKH","%NVK","Tổng %","Xếp loại cuối","Tổng % TNTT (GD 50% + NCKH 50%)","Xếp TNTT","Cảnh báo / lý do","Cảnh báo TNTT"]
        st.dataframe(view[show],use_container_width=True,hide_index=True)
with tab3:
    st.subheader("Tổng hợp theo đơn vị")
    if results.empty: st.warning("Chưa có dữ liệu.")
    else:
        categories=["HTXSNV","HTTNV","HTNV","KHTNV","KHÔNG ĐÁNH GIÁ"]
        tntt_cats=["A","B1","B2","B3","C","D","CẦN XEM XÉT"]
        summary=[]
        for dep in DEPARTMENTS:
            part=results[results["Đơn vị"]==dep]
            item={"Đơn vị":dep,"Số GV có dữ liệu":len(part)}
            item.update({x:int((part["Xếp loại cuối"]==x).sum()) for x in categories})
            item.update({x:int((part["Xếp TNTT"]==x).sum()) for x in tntt_cats})
            summary.append(item)
        summ=pd.DataFrame(summary)
        total={"Đơn vị":"TOÀN TRƯỜNG"}
        for c in summ.columns[1:]: total[c]=int(summ[c].sum())
        summ=pd.concat([summ,pd.DataFrame([total])],ignore_index=True)
        st.dataframe(summ,use_container_width=True,hide_index=True)
        st.bar_chart(summ[summ["Đơn vị"]!="TOÀN TRƯỜNG"].set_index("Đơn vị")[categories])

if not results.empty:
    st.divider()
    st.subheader("Xuất dữ liệu")
    output=BytesIO()
    with pd.ExcelWriter(output,engine="openpyxl") as writer:
        results.to_excel(writer,index=False,sheet_name="KET_QUA")
        if "summ" in locals(): summ.to_excel(writer,index=False,sheet_name="TONG_HOP")
        st.session_state.records.to_excel(writer,index=False,sheet_name="DU_LIEU_NHAP")
    st.download_button("Tải kết quả Excel",data=output.getvalue(),file_name="ket_qua_danh_gia_giang_vien.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
