import io
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Đánh giá kết quả thực hiện nhiệm vụ",
    page_icon="📊",
    layout="wide",
)

st.title("ĐÁNH GIÁ KẾT QUẢ THỰC HIỆN CHỨC TRÁCH, NHIỆM VỤ")
st.caption("Ứng dụng chuyển đổi từ biểu mẫu Excel 1.MS01.xlsx")

st.info(
    "Nhập số liệu thực hiện và các tiêu chí NCKH. Ứng dụng tự động tính điểm, "
    "đánh giá điều kiện và xếp loại theo công thức trong file Excel gốc."
)

# -------------------- Input --------------------
st.subheader("1. Kết quả thực hiện nhiệm vụ")

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown("**Giảng dạy (40%)**")
    teaching_target = st.number_input("Định mức giảng dạy", min_value=0.0, value=270.0, step=1.0)
    teaching_done = st.number_input("Thực hiện giảng dạy", min_value=0.0, value=0.0, step=1.0)
with c2:
    st.markdown("**Nghiên cứu khoa học (40%)**")
    research_target = st.number_input("Ngưỡng NCKH (điểm PL02)", min_value=0.0, value=600.0, step=1.0)
    research_done = st.number_input("Điểm NCKH thực hiện", min_value=0.0, value=0.0, step=1.0)
with c3:
    st.markdown("**Nhiệm vụ khác (20%)**")
    other_target = st.number_input("Định mức nhiệm vụ khác", min_value=0.0, value=350.0, step=1.0)
    other_done = st.number_input("Thực hiện nhiệm vụ khác", min_value=0.0, value=0.0, step=1.0)

st.divider()
st.subheader("2. Điều kiện và điểm quy đổi nghiên cứu khoa học")

rank_labels = {
    1: "1 – Giáo sư / Phó Giáo sư",
    2: "2 – Tiến sĩ",
    3: "3 – Đại học / Thạc sĩ",
}
rank_label = st.selectbox("Chức danh", list(rank_labels.values()), index=0)
rank = {v: k for k, v in rank_labels.items()}[rank_label]

binary_labels = [
    "Có bài WoS/Scopus Q1/Q2",
    "Hướng dẫn SV/HV có bài WoS/Scopus Q1/Q2",
    "Hướng dẫn SV nghiên cứu khoa học đạt giải 1–3",
    "Có bằng sáng chế",
    "Có công trình mục 3 hoặc 5 B1",
    "Có công trình mục 1–6 B1",
]
binary_keys = ["C", "D", "E", "F", "G", "H"]
binary_values = {}
cols = st.columns(3)
for i, (label, key) in enumerate(zip(binary_labels, binary_keys)):
    with cols[i % 3]:
        binary_values[key] = 1 if st.checkbox(label, value=False, key=f"binary_{key}") else 0

research_score = st.number_input(
    "Điểm NCKH quy đổi theo Bảng 1, Phụ lục 02 (điểm)",
    min_value=0.0,
    value=0.0,
    step=1.0,
    help="Nhập điểm quy đổi NCKH dùng để so sánh với ngưỡng áp dụng theo chức danh.",
)

# -------------------- Calculations (mirror workbook formulas) --------------------
teaching_pct = (teaching_done / teaching_target * 100 * 0.4) if teaching_target > 0 else 0.0
research_pct = (research_done / research_target * 100 * 0.4) if research_target > 0 else 0.0
other_pct = min(other_done / other_target * 100 * 0.2, 20.0) if other_target > 0 else 0.0
total_pct = teaching_pct + research_pct + other_pct

# Excel I8 = IF(OR(C8=1,D8=1,E8=1,G8=1,H8=1),"Đạt","Không đạt")
additional_criteria = any(binary_values[k] == 1 for k in ["C", "D", "E", "G", "H"])
additional_status = "Đạt" if additional_criteria else "Không đạt"

# Excel J8 = IF(B8=3,1,IF(B8=2,1.5,IF(B8=1,2,"")))
research_threshold = {1: 2.0, 2: 1.5, 3: 1.0}[rank]
research_status = "Đạt" if research_score >= research_threshold else "Không đạt"

# Excel L4 classification formula, preserving its original condition order.
if total_pct >= 150 and additional_status == "Đạt" and research_status == "Đạt" and rank != 3:
    performance_rating = "HTXSNV"
elif total_pct >= 100:
    performance_rating = "HTTNV"
elif rank == 3 and research_status == "Đạt" and total_pct >= 100:
    performance_rating = "HTNV"
else:
    performance_rating = "KHTNV"

# Excel K9 and L9: income-increase score and classification.
teaching_ratio = teaching_done / teaching_target if teaching_target > 0 else 0.0
research_ratio = research_done / research_target if research_target > 0 else 0.0
income_score = teaching_ratio * 100 * 0.5 + research_ratio * 100 * 0.5

if (
    income_score >= 400 and teaching_ratio >= 1 and research_ratio >= 1
    and other_target >= 350 and performance_rating == "HTXSNV"
):
    income_rating = "A"
elif income_score >= 300 and teaching_ratio >= 1 and research_ratio >= 1 and other_target >= 350:
    income_rating = "B1"
elif income_score >= 200 and teaching_ratio >= 1 and research_ratio >= 1 and other_target >= 350:
    income_rating = "B2"
elif income_score >= 100 and teaching_ratio >= 1 and research_ratio >= 1 and other_target >= 350:
    income_rating = "B3"
elif income_score >= 50 and teaching_ratio > 0 and research_ratio > 0 and other_target >= 175:
    income_rating = "C"
else:
    income_rating = "D"

# -------------------- Results --------------------
st.divider()
st.subheader("3. Kết quả đánh giá")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Giảng dạy (đóng góp)", f"{teaching_pct:.2f}%")
m2.metric("NCKH (đóng góp)", f"{research_pct:.2f}%")
m3.metric("Nhiệm vụ khác (đóng góp)", f"{other_pct:.2f}%")
m4.metric("Tổng tỷ lệ đạt được", f"{total_pct:.2f}%")

r1, r2, r3 = st.columns(3)
r1.metric("Điều kiện bổ sung", additional_status)
r2.metric("Định mức NCKH", f"{research_threshold:g} điểm")
r3.metric("Đủ định mức NCKH", research_status)

st.markdown("#### Xếp loại")
st.success(f"**Kết quả xếp loại: {performance_rating}**")
st.markdown("#### Thu nhập tăng thêm")
i1, i2 = st.columns(2)
i1.metric("Điểm thu nhập tăng thêm", f"{income_score:.2f}")
i2.metric("Mức xếp loại thu nhập tăng thêm", income_rating)

with st.expander("Xem bảng tính chi tiết"):
    detail = pd.DataFrame(
        [
            ["Giảng dạy", teaching_target, teaching_done, teaching_pct],
            ["Nghiên cứu khoa học", research_target, research_done, research_pct],
            ["Nhiệm vụ khác", other_target, other_done, other_pct],
        ],
        columns=["Nhóm tiêu chí", "Định mức/ngưỡng", "Thực hiện", "Tỷ lệ đóng góp (%)"],
    )
    st.dataframe(detail, use_container_width=True, hide_index=True)
    st.write("**Các tiêu chí NCKH:**")
    st.dataframe(
        pd.DataFrame(
            {
                "Tiêu chí": binary_labels,
                "Giá trị": [binary_values[k] for k in binary_keys],
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

# -------------------- Excel export --------------------
def make_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Ket qua danh gia"
    ws.merge_cells("A1:D1")
    ws["A1"] = "KẾT QUẢ ĐÁNH GIÁ THỰC HIỆN CHỨC TRÁCH, NHIỆM VỤ"
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center")
    rows = [
        ["Thông tin", "Định mức/ngưỡng", "Thực hiện/giá trị", "Kết quả"],
        ["Giảng dạy (40%)", teaching_target, teaching_done, teaching_pct],
        ["NCKH (40%)", research_target, research_done, research_pct],
        ["Nhiệm vụ khác (20%)", other_target, other_done, other_pct],
        ["Tổng tỷ lệ đạt được", "", "", total_pct],
        ["Chức danh", "", rank, rank_labels[rank]],
        ["Ngưỡng NCKH áp dụng", research_threshold, research_score, research_status],
        ["Điều kiện bổ sung", "", "", additional_status],
        ["Xếp loại nhiệm vụ", "", "", performance_rating],
        ["Điểm thu nhập tăng thêm", "", "", income_score],
        ["Xếp loại thu nhập tăng thêm", "", "", income_rating],
    ]
    for row in rows:
        ws.append(row)
    ws.append([])
    ws.append(["Tiêu chí NCKH", "Giá trị (1=có, 0=không)"])
    for label, key in zip(binary_labels, binary_keys):
        ws.append([label, binary_values[key]])
    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for cell in ws[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for col, width in {"A": 42, "B": 20, "C": 20, "D": 28}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

st.download_button(
    "⬇️ Tải kết quả đánh giá (Excel)",
    data=make_excel(),
    file_name=f"Ket_qua_danh_gia_{datetime.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Lưu ý về công thức và phạm vi chuyển đổi"):
    st.markdown(
        """
        - Các định mức mặc định 270, 600 và 350 được lấy từ biểu mẫu Excel.
        - Ngưỡng NCKH theo chức danh được giữ nguyên: GS/PGS = 2; TS = 1,5; ĐH/ThS = 1.
        - Điều kiện bổ sung được xác định từ các cột C, D, E, G, H của biểu mẫu; cột F không tham gia công thức I8.
        - Công thức xếp loại và xếp loại thu nhập tăng thêm được chuyển theo thứ tự điều kiện trong Excel gốc.
        - Ứng dụng hiện sử dụng các định mức mặc định của file; nếu đơn vị có định mức khác, cần điều chỉnh trong mã nguồn hoặc yêu cầu bổ sung chức năng cấu hình.
        """
    )
