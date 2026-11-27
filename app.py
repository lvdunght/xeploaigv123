from io import BytesIO
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Phiếu Đánh Giá Xếp Loại Chất Lượng Viên Chức (Giảng Viên)",
    page_icon="📝",
    layout="wide",
)

TEMPLATE_PATH = "Phieu danh gia GV.xlsx"
SHEET_NAME = "PHIEU ĐG giang vien"


def tinh_ket_qua(
    dm_gd: float,
    th_gd: float | None,
    nguong_nckh: float,
    diem_nckh: float | None,
    dm_khac: float,
    th_khac: float | None,
    chuc_danh: int,
    wos_q12: int,
    hd_wos_q12: int,
    hd_sv_giai: int,
    bang_sang_che: int,
    ct_3_5_b1: int,
    ct_1_6_b1: int,
    diem_quy_doi_nckh: float,
    tong_diem_vi_pham: float,
) -> dict:
    # Hàng 12: Điểm thành phần & Tổng %
    d12 = None if (th_gd is None or dm_gd <= 0) else (th_gd / dm_gd) * 100.0 * 0.4
    g12 = None if (diem_nckh is None or nguong_nckh <= 0) else (diem_nckh / nguong_nckh) * 100.0 * 0.4
    j12 = (
        None
        if (th_khac is None or dm_khac <= 0)
        else min((th_khac / dm_khac) * 100.0 * 0.2, 20.0)
    )

    k12 = 0.0 if (d12 is None or g12 is None or j12 is None) else (d12 + g12 + j12)

    # Hàng 16: Tiêu chí bổ sung & Ngưỡng NCKH theo chức danh
    i16 = (
        "Đạt"
        if any(x == 1 for x in [wos_q12, hd_wos_q12, hd_sv_giai, ct_3_5_b1, ct_1_6_b1])
        else "Không đạt"
    )
    j16 = {1: 2.0, 2: 1.5, 3: 1.0}.get(chuc_danh, 0.0)
    l16 = "Đạt" if diem_quy_doi_nckh >= j16 else "Không đạt"

    # L12: Xếp loại theo công thức Excel
    if k12 >= 150 and i16 == "Đạt" and l16 == "Đạt" and chuc_danh != 3:
        l12 = "HTXSNV"
    elif k12 >= 100 and i16 == "Đạt" and l16 == "Đạt":
        l12 = "HTTNV"
    elif k12 >= 100:
        l12 = "HTNV"
    else:
        l12 = "KHTNV"

    # Điều chỉnh xếp loại theo lưu ý (*) tổng điểm vi phạm
    thu_tu = ["HTXSNV", "HTTNV", "HTNV", "KHTNV"]
    idx = thu_tu.index(l12)
    if tong_diem_vi_pham >= 30:
        l12_sau_vp = "KHTNV"
    elif tong_diem_vi_pham >= 20:
        l12_sau_vp = thu_tu[min(idx + 2, len(thu_tu) - 1)]
    elif tong_diem_vi_pham >= 10:
        l12_sau_vp = thu_tu[min(idx + 1, len(thu_tu) - 1)]
    else:
        l12_sau_vp = l12

    # Hàng 17: Thu nhập tăng thêm
    ty_le_gd = 0.0 if (th_gd is None or dm_gd <= 0) else (th_gd / dm_gd)
    ty_le_nckh = 0.0 if (diem_nckh is None or nguong_nckh <= 0) else (diem_nckh / nguong_nckh)
    k17 = ty_le_gd * 100.0 * 0.5 + ty_le_nckh * 100.0 * 0.5

    if k17 >= 400 and ty_le_gd >= 1 and ty_le_nckh >= 1 and dm_khac >= 350 and l12 == "HTXSNV":
        l17 = "A"
    elif k17 >= 300 and ty_le_gd >= 1 and ty_le_nckh >= 1 and dm_khac >= 350:
        l17 = "B1"
    elif k17 >= 200 and ty_le_gd >= 1 and ty_le_nckh >= 1 and dm_khac >= 350:
        l17 = "B2"
    elif k17 >= 100 and ty_le_gd >= 1 and ty_le_nckh >= 1 and dm_khac >= 350:
        l17 = "B3"
    elif k17 >= 50 and ty_le_gd > 0 and ty_le_nckh > 0 and dm_khac >= 175:
        l17 = "C"
    else:
        l17 = "D"

    return {
        "D12": d12,
        "G12": g12,
        "J12": j12,
        "K12": k12,
        "L12": l12,
        "L12_sau_vp": l12_sau_vp,
        "I16": i16,
        "J16": j16,
        "L16": l16,
        "K17": k17,
        "L17": l17,
    }


def tao_file_excel(info: dict, inputs: dict, kq: dict, df_vp: pd.DataFrame) -> bytes:
    if os.path.exists(TEMPLATE_PATH):
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

        ws["B3"] = (
            f"PHIẾU ĐÁNH GIÁ XẾP, LOẠI CHẤT LƯỢNG VIÊN CHỨC NĂM {info['nam']}\n"
            "(Dành cho viên chức ngạch giảng viên)"
        )
        ws["B4"] = f"Họ và tên: {info['ho_ten']}"
        ws["B5"] = f"Đơn vị: {info['don_vi']}"
        ws["B6"] = f"Chức danh: {info['chuc_danh_text']}"
        ws["B7"] = f"Chức vụ: {info['chuc_vu']}"
        ws["B8"] = f"1. Tiêu chí chung: về chính trị tư tưởng, đạo đức, lối sống, tác phong, lề lối làm việc, ý thức tổ chức kỷ luật: {info['tieu_chi_chung']}"

        # Hàng 12
        ws["B12"] = inputs["dm_gd"]
        ws["C12"] = inputs["th_gd"]
        ws["D12"] = round(kq["D12"], 2) if kq["D12"] is not None else ""
        ws["E12"] = inputs["nguong_nckh"]
        ws["F12"] = inputs["diem_nckh"]
        ws["G12"] = round(kq["G12"], 2) if kq["G12"] is not None else ""
        ws["H12"] = inputs["dm_khac"]
        ws["I12"] = inputs["th_khac"]
        ws["J12"] = round(kq["J12"], 2) if kq["J12"] is not None else ""
        ws["K12"] = round(kq["K12"], 2)
        ws["L12"] = kq["L12_sau_vp"]

        # Hàng 16
        ws["B16"] = inputs["chuc_danh"]
        ws["C16"] = inputs["wos_q12"]
        ws["D16"] = inputs["hd_wos_q12"]
        ws["E16"] = inputs["hd_sv_giai"]
        ws["F16"] = inputs["bang_sang_che"]
        ws["G16"] = inputs["ct_3_5_b1"]
        ws["H16"] = inputs["ct_1_6_b1"]
        ws["I16"] = kq["I16"]
        ws["J16"] = kq["J16"]
        ws["K16"] = inputs["diem_quy_doi_nckh"]
        ws["L16"] = kq["L16"]

        # Hàng 17
        ws["K17"] = round(kq["K17"], 2)
        ws["L17"] = kq["L17"]

        # Bảng vi phạm (hàng 20-22)
        for idx, row in enumerate(df_vp.itertuples(index=False), start=20):
            if idx <= 22:
                ws[f"C{idx}"] = getattr(row, "Nội_dung_vi_phạm", "")
                ws[f"I{idx}"] = getattr(row, "Điểm_vi_phạm", 0)
                ws[f"K{idx}"] = getattr(row, "Ghi_chú", "")
        ws["I23"] = float(df_vp["Điểm_vi_phạm"].sum()) if not df_vp.empty else 0.0

        # Phần tự đánh giá (hàng 27, 29) & Thủ trưởng (hàng 37, 39)
        ws["B27"] = info["tu_xl"]
        ws["B29"] = info["tu_tntt"]
        ws["B37"] = info["tt_xl"]
        ws["B39"] = info["tt_tntt"]
        ws["I31"] = info["ngay_ky"]

        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # Dự phòng nếu chưa có file mẫu gốc
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KetQuaDanhGia"
    ws.append(["Chỉ tiêu", "Giá trị"])
    ws.append(["Họ và tên", info["ho_ten"]])
    ws.append(["Đơn vị", info["don_vi"]])
    ws.append(["Chức danh", info["chuc_danh_text"]])
    ws.append(["Tổng % đạt được (K12)", round(kq["K12"], 2)])
    ws.append(["Xếp loại (L12)", kq["L12"]])
    ws.append(["Xếp loại sau điểm vi phạm", kq["L12_sau_vp"]])
    ws.append(["Điểm TNTT (K17)", round(kq["K17"], 2)])
    ws.append(["Xếp loại TNTT (L17)", kq["L17"]])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


st.title("PHIẾU ĐÁNH GIÁ XẾP LOẠI CHẤT LƯỢNG VIÊN CHỨC (GIẢNG VIÊN)")
st.caption("Mô phỏng biểu mẫu & công thức tính toán từ file **Phieu danh gia GV.xlsx**")

# --- PHẦN THÔNG TIN CHUNG ---
with st.expander("Thông tin viên chức & Tiêu chí chung", expanded=True):
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        ho_ten = st.text_input("Họ và tên", placeholder="Nguyễn Văn A")
        don_vi = st.text_input("Đơn vị", placeholder="Khoa ...")
    with c2:
        chuc_danh_map = {
            1: "1 - GS / PGS",
            2: "2 - Tiến sĩ (TS)",
            3: "3 - Đại học / Thạc sĩ (ĐH/ThS)",
        }
        chuc_danh = st.selectbox(
            "Chức danh (mã B16)",
            options=[1, 2, 3],
            format_func=lambda x: chuc_danh_map[x],
            index=0,
        )
        chuc_vu = st.text_input("Chức vụ", placeholder="Giảng viên / Phó Trưởng bộ môn...")
    with c3:
        nam_dg = st.text_input("Năm đánh giá", value="2026")
        ngay_ky = st.text_input("Ngày ký", value="Đà Nẵng, ngày ... tháng ... năm 2026")

    tieu_chi_chung = st.text_area(
        "1. Tiêu chí chung (chính trị tư tưởng, đạo đức, lối sống, tác phong, ý thức tổ chức kỷ luật)",
        height=70,
    )

# --- PHẦN 2: KẾT QUẢ THỰC HIỆN CHỨC TRÁCH NHIỆM VỤ ĐƯỢC GIAO ---
st.subheader("2. Kết quả thực hiện chức trách nhiệm vụ được giao")

col_gd, col_nckh, col_khac = st.columns(3)

with col_gd:
    st.markdown("#### GIẢNG DẠY (40%)")
    dm_gd = st.number_input("Định mức giảng dạy (B12)", min_value=0.0, value=270.0, step=10.0)
    th_gd = st.number_input("Thực hiện giảng dạy (C12)", min_value=0.0, value=270.0, step=10.0)

with col_nckh:
    st.markdown("#### NCKH (40%)")
    nguong_nckh = st.number_input("Ngưỡng NCKH - điểm PL02 (E12)", min_value=0.0, value=600.0, step=10.0)
    diem_nckh = st.number_input("Điểm NCKH thực hiện (F12)", min_value=0.0, value=600.0, step=10.0)

with col_khac:
    st.markdown("#### NHIỆM VỤ KHÁC (20%)")
    dm_khac = st.number_input("Định mức nhiệm vụ khác (H12)", min_value=0.0, value=350.0, step=10.0)
    th_khac = st.number_input("Thực hiện nhiệm vụ khác (I12)", min_value=0.0, value=350.0, step=10.0)

st.markdown("#### Tiêu chí bổ sung & Định mức NCKH theo chức danh (Hàng 14–16)")
tc_cols = st.columns(6)
with tc_cols[0]:
    wos_q12 = st.selectbox("WoS/Scopus Q1/Q2 (C16)", [0, 1], index=0)
with tc_cols[1]:
    hd_wos_q12 = st.selectbox("HD SV/HV Q1/Q2 (D16)", [0, 1], index=0)
with tc_cols[2]:
    hd_sv_giai = st.selectbox("HD SV giải 1–3 (E16)", [0, 1], index=0)
with tc_cols[3]:
    bang_sang_che = st.selectbox("Bằng sáng chế (F16)", [0, 1], index=0)
with tc_cols[4]:
    ct_3_5_b1 = st.selectbox("Mục 3 hoặc 5 B1 (G16)", [0, 1], index=0)
with tc_cols[5]:
    ct_1_6_b1 = st.selectbox("Mục 1–6 B1 (H16)", [0, 1], index=0)

diem_quy_doi_nckh = st.number_input(
    "NCKH - điểm quy đổi Bảng 1 PL02 (K16)", min_value=0.0, value=2.0, step=0.25
)

# --- VI PHẠM ---
st.markdown("#### Vi phạm & Điểm trừ (*)")
df_vp_default = pd.DataFrame(
    [
        {"TT": 1, "Nội_dung_vi_phạm": "", "Điểm_vi_phạm": 0.0, "Ghi_chú": ""},
        {"TT": 2, "Nội_dung_vi_phạm": "", "Điểm_vi_phạm": 0.0, "Ghi_chú": ""},
        {"TT": 3, "Nội_dung_vi_phạm": "", "Điểm_vi_phạm": 0.0, "Ghi_chú": ""},
    ]
)
df_vp = st.data_editor(
    df_vp_default,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
)
tong_diem_vi_pham = float(pd.to_numeric(df_vp["Điểm_vi_phạm"], errors="coerce").fillna(0.0).sum())
st.caption(
    f"**Tổng điểm vi phạm (*): {tong_diem_vi_pham:.2f}** — "
    "Quy tắc: ≥ 10 hạ 1 mức; ≥ 20 hạ 2 mức; ≥ 30 xếp loại Không hoàn thành nhiệm vụ."
)

# --- TÍNH TOÁN & HIỂN THỊ KẾT QUẢ ---
kq = tinh_ket_qua(
    dm_gd=dm_gd,
    th_gd=th_gd,
    nguong_nckh=nguong_nckh,
    diem_nckh=diem_nckh,
    dm_khac=dm_khac,
    th_khac=th_khac,
    chuc_danh=chuc_danh,
    wos_q12=wos_q12,
    hd_wos_q12=hd_wos_q12,
    hd_sv_giai=hd_sv_giai,
    bang_sang_che=bang_sang_che,
    ct_3_5_b1=ct_3_5_b1,
    ct_1_6_b1=ct_1_6_b1,
    diem_quy_doi_nckh=diem_quy_doi_nckh,
    tong_diem_vi_pham=tong_diem_vi_pham,
)

st.divider()
st.subheader("Kết quả tổng hợp (giống bảng tính Excel)")

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Giảng dạy (D12)", f"{kq['D12']:.2f}" if kq["D12"] is not None else "-")
m2.metric("NCKH (G12)", f"{kq['G12']:.2f}" if kq["G12"] is not None else "-")
m3.metric("Nhiệm vụ khác (J12)", f"{kq['J12']:.2f}" if kq["J12"] is not None else "-")
m4.metric("Tổng % đạt được (K12)", f"{kq['K12']:.2f}")
m5.metric("Xếp loại (L12)", kq["L12_sau_vp"], delta=f"Gốc: {kq['L12']}")

t1, t2, t3, t4 = st.columns(4)
t1.metric("Tiêu chí bổ sung (I16)", kq["I16"])
t2.metric("Ngưỡng PL1/PL02 (J16)", f"{kq['J16']:.2f}")
t3.metric("Đủ ĐM NCKH (L16)", kq["L16"])
t4.metric("Thu nhập tăng thêm", f"{kq['K17']:.2f} | Loại {kq['L17']}")

bang_tom_tat = pd.DataFrame(
    [
        {
            "Định mức GD (B12)": dm_gd,
            "Thực hiện GD (C12)": th_gd,
            "(% đạt) x 0.4 GD (D12)": round(kq["D12"], 2) if kq["D12"] is not None else None,
            "Ngưỡng NCKH (E12)": nguong_nckh,
            "Điểm NCKH (F12)": diem_nckh,
            "(% đạt) x 0.4 NCKH (G12)": round(kq["G12"], 2) if kq["G12"] is not None else None,
            "Định mức khác (H12)": dm_khac,
            "Thực hiện khác (I12)": th_khac,
            "(% đạt) x 0.2 Khác (J12)": round(kq["J12"], 2) if kq["J12"] is not None else None,
            "Tổng % (K12)": round(kq["K12"], 2),
            "Xếp loại (L12)": kq["L12_sau_vp"],
            "Điểm TNTT (K17)": round(kq["K17"], 2),
            "Xếp TNTT (L17)": kq["L17"],
        }
    ]
)
st.dataframe(bang_tom_tat, use_container_width=True, hide_index=True)

# --- TỰ ĐÁNH GIÁ & THỦ TRƯỞNG ---
with st.expander("4. Kết quả tự đánh giá & Nhận xét của Thủ trưởng đơn vị"):
    c_tu, c_tt = st.columns(2)
    with c_tu:
        st.markdown("**Viên chức tự đánh giá**")
        tu_xl = st.text_input("a) Xếp loại chất lượng viên chức (tự ĐG)", value=kq["L12_sau_vp"])
        tu_tntt = st.text_input("b) Xếp thu nhập tăng thêm (A; B; C; D) (tự ĐG)", value=kq["L17"])
    with c_tt:
        st.markdown("**Thủ trưởng đơn vị đánh giá**")
        tt_xl = st.text_input("a) Xếp loại chất lượng viên chức (Thủ trưởng)", value="")
        tt_tntt = st.text_input("b) Xếp lương/TNTT (A; B; C; D) (Thủ trưởng)", value="")

# --- NÚT TẢI FILE KẾT QUẢ ---
info_payload = {
    "nam": nam_dg,
    "ho_ten": ho_ten,
    "don_vi": don_vi,
    "chuc_danh_text": chuc_danh_map[chuc_danh],
    "chuc_vu": chuc_vu,
    "tieu_chi_chung": tieu_chi_chung,
    "tu_xl": tu_xl,
    "tu_tntt": tu_tntt,
    "tt_xl": tt_xl,
    "tt_tntt": tt_tntt,
    "ngay_ky": ngay_ky,
}
inputs_payload = {
    "dm_gd": dm_gd,
    "th_gd": th_gd,
    "nguong_nckh": nguong_nckh,
    "diem_nckh": diem_nckh,
    "dm_khac": dm_khac,
    "th_khac": th_khac,
    "chuc_danh": chuc_danh,
    "wos_q12": wos_q12,
    "hd_wos_q12": hd_wos_q12,
    "hd_sv_giai": hd_sv_giai,
    "bang_sang_che": bang_sang_che,
    "ct_3_5_b1": ct_3_5_b1,
    "ct_1_6_b1": ct_1_6_b1,
    "diem_quy_doi_nckh": diem_quy_doi_nckh,
}

excel_bytes = tao_file_excel(info_payload, inputs_payload, kq, df_vp)
file_name_out = f"Phieu_danh_gia_{ho_ten.strip().replace(' ', '_') or 'GV'}_{nam_dg}.xlsx"

st.download_button(
    label="Tải kết quả về (.xlsx)",
    data=excel_bytes,
    file_name=file_name_out,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
