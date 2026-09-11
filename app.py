import io
import os
from datetime import datetime
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Dashboard Quản Lý Chấm Công & Nhân Sự",
    page_icon="📊",
    layout="wide",
)

# CSS tùy chỉnh giao diện đẹp và chuyên nghiệp
st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stDownloadButton button { width: 100%; border-radius: 6px; font-weight: bold; }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Hệ Thống Quản Lý & Phân Tích Giờ Làm Việc Tự Động")
st.markdown("Tích hợp AI phân tích chuyên sâu lịch trình chấm công và hiệu suất.")

# --- KHU VỰC UPLOAD FILE ---
st.sidebar.header("📁 Tải lên dữ liệu đầu vào")
uploaded_cham_cong = st.sidebar.file_uploader(
    "1. File Bấm Vân Tay (Excel)", type=["xlsx", "xls"]
)
uploaded_cn = st.sidebar.file_uploader(
    "2. File Danh Sách Công Nhân (Excel)", type=["xlsx", "xls"]
)
uploaded_vp = st.sidebar.file_uploader(
    "3. File Danh Sách Văn Phòng (Excel)", type=["xlsx", "xls"]
)
uploaded_lich = st.sidebar.file_uploader(
    "4. File Lịch Xếp Ca (Excel)", type=["xlsx", "xls"]
)

# Nhập API Key cho Gemini AI
st.sidebar.header("🔑 Cấu hình Gemini AI")
gemini_api_key = st.sidebar.text_input(
    "Nhập Gemini API Key", type="password", help="Dùng để phân tích đánh giá nhân sự"
)


@st.cache_data
def load_data(f_cc, f_cn, f_vp, f_lich):
    try:
        df_cc = pd.read_excel(f_cc) if f_cc else pd.DataFrame()
        df_cn = pd.read_excel(f_cn) if f_cn else pd.DataFrame()
        df_vp = pd.read_excel(f_vp) if f_vp else pd.DataFrame()
        df_lich = pd.read_excel(f_lich) if f_lich else pd.DataFrame()
        return df_cc, df_cn, df_vp, df_lich
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return (
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
        )


if uploaded_cham_cong and uploaded_cn and uploaded_vp and uploaded_lich:
    df_cc, df_cn, df_vp, df_lich = load_data(
        uploaded_cham_cong, uploaded_cn, uploaded_vp, uploaded_lich
    )

    # Phân loại nhóm nhân viên
    df_cn["Nhom"] = "Công Nhân"
    df_vp["Nhom"] = "Văn Phòng"
    df_nv_all = pd.concat([df_cn, df_vp], ignore_index=True)

    # --- XỬ LÝ DỮ LIỆU CHẤM CÔNG & LOGIC ---
    # Merge thông tin nhân viên vào bảng chấm công
    df_merged = pd.merge(
        df_cc, df_nv_all[["MaNV", "HoTen", "Nhom"]], on="MaNV", how="left"
    )
    df_merged = pd.merge(
        df_merged,
        df_lich[["MaNV", "Ngay", "Ca"]],
        on=["MaNV", "Ngay"],
        how="left",
    )

    # Xử lý logic tính toán giờ vào/ra, đi trễ, về sớm, tăng ca
    # Định nghĩa giờ chuẩn (ví dụ ca hành chính: Vào 08:00, Ra 17:00; Ca Đêm: Vào 22:00, Ra 06:00 sáng hôm sau)
    def calculate_attendance(row):
        nhom = row.get("Nhom", "Khác")
        ca = row.get("Ca", "HC")
        gio_vao = row.get("ThoiGianVao")
        gio_ra = row.get("ThoiGianRa")

        status = "Bình Thường"
        ghi_chu = "BT"
        gio_lam_viec = 0.0

        if pd.isna(gio_vao) or pd.isna(gio_ra):
            return pd.Series(
                ["Vắng / Không bấm thẻ", "Chưa bấm thẻ", 0.0, "Vắng"]
            )

        # Chuyển đổi định dạng giờ nếu là string
        try:
            t_vao = pd.to_datetime(str(gio_vao)).time()
            t_ra = pd.to_datetime(str(gio_ra)).time()
        except:
            return pd.Series(["Lỗi định dạng giờ", "Lỗi", 0.0, "Lỗi"])

        # Tính tổng giờ làm đơn giản (giả định tính theo giờ)
        dt_vao = datetime.combine(datetime.today(), t_vao)
        dt_ra = datetime.combine(datetime.today(), t_ra)
        if dt_ra < dt_vao:  # Ca qua đêm
            dt_ra += pd.Timedelta(days=1)

        gio_lam_viec = (dt_ra - dt_vao).seconds / 3600.0
        # Trừ thời gian nghỉ trưa/giữa ca (ví dụ 1 tiếng nếu làm > 5 tiếng)
        if gio_lam_viec > 5:
            gio_lam_viec -= 1.0

        # Đánh giá theo nhóm công nhân / văn phòng dựa trên lịch ca
        if nhom == "Công Nhân":
            if ca == "Đ":  # Ca đêm
                if t_vao > pd.to_datetime("22:15:00").time():
                    ghi_chu = "Đi trễ"
            else:  # Ca ngày chuẩn
                if t_vao > pd.to_datetime("08:15:00").time():
                    ghi_chu = "Đi trễ"
                if t_ra < pd.to_datetime("17:00:00").time():
                    ghi_chu = (
                        "Về sớm" if ghi_chu == "BT" else "Đi trễ & Về sớm"
                    )
        else:  # Văn phòng và nhóm còn lại
            if gio_lam_viec < 8.0:
                ghi_chu = "Về sớm / Thiếu giờ"
            elif gio_lam_viec > 8.5:
                ghi_chu = "Tăng ca"
            else:
                ghi_chu = "BT"

        return pd.Series([status, ghi_chu, round(gio_lam_viec, 2), ca])

    df_merged[
        ["TrangThaiChamCong", "DanhGia", "TongGioLam", "CaLamViec"]
    ] = df_merged.apply(calculate_attendance, axis=1)

    # --- BỘ LỌC TRÊN DASHBOARD ---
    st.sidebar.header("🔍 Bộ Lọc Phân Tích")
    unique_dates = (
        sorted(df_merged["Ngay"].dropna().unique())
        if "Ngay" in df_merged.columns
        else []
    )
    selected_date = st.sidebar.selectbox(
        "Chọn Ngày Phân Tích", ["Tất cả"] + list(unique_dates)
    )

    selected_nhom = st.sidebar.selectbox(
        "Chọn Nhóm Nhân Sự", ["Tất cả", "Công Nhân", "Văn Phòng"]
    )
    selected_time_type = st.sidebar.selectbox(
        "Lọc Loại Thời Gian Bấm Thẻ", ["Tất cả", "Vào", "Ra"]
    )

    # Áp dụng bộ lọc
    df_filtered = df_merged.copy()
    if selected_date != "Tất cả":
        df_filtered = df_filtered[df_filtered["Ngay"] == selected_date]
    if selected_nhom != "Tất cả":
        df_filtered = df_filtered[df_filtered["Nhom"] == selected_nhom]

    # --- HIỂN THỊ CHỈ SỐ TỔNG QUAN (METRICS) ---
    st.subheader(
        f"📈 Báo Cáo Tổng Quan Ngày: {selected_date if selected_date != 'Tất cả' else 'Toàn Bộ Thời Gian'}"
    )

    total_nv = len(df_filtered)
    di_tre = len(df_filtered[df_filtered["DanhGia"].str.contains("Đi trễ")])
    ve_som = len(df_filtered[df_filtered["DanhGia"].str.contains("Về sớm")])
    khong_bam_the = len(
        df_filtered[df_filtered["TrangThaiChamCong"].str.contains("Vắng")]
    )
    tang_ca = len(df_filtered[df_filtered["DanhGia"] == "Tăng ca"])

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Tổng Nhân Sự Lọc", total_nv)
    col2.metric("Đi Trễ", di_tre, delta_color="inverse")
    col3.metric("Về Sớm", ve_som, delta_color="inverse")
    col4.metric("Không Bấm Thẻ / Vắng", khong_bam_the, delta_color="inverse")
    col5.metric("Tăng Ca", tang_ca, delta_color="normal")

    # --- BIỂU ĐỒ PHÂN TÍCH ---
    st.markdown("---")
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        st.markdown("#### 📊 Tỷ Lệ Đánh Giá Chấm Công")
        if not df_filtered.empty:
            fig_pie = px.pie(
                df_filtered,
                names="DanhGia",
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Set3,
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Không có dữ liệu hiển thị biểu đồ.")

    with col_chart2:
        st.markdown("#### 📉 Tổng Giờ Làm Theo Bộ Phận")
        if not df_filtered.empty and "TongGioLam" in df_filtered.columns:
            fig_bar = px.bar(
                df_filtered,
                x="HoTen",
                y="TongGioLam",
                color="Nhom",
                barmode="group",
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Không có dữ liệu giờ làm.")

    # --- BẢNG DỮ LIỆU CHI TIẾT ---
    st.markdown("---")
    st.subheader("📋 Bảng Chi Tiết Chấm Công & Vi Phạm")
    st.dataframe(df_filtered, use_container_width=True)

    # --- XUẤT FILE EXCEL & PDF ---
    st.markdown("---")
    col_dl1, col_dl2 = st.columns(2)


    # Hàm xuất Excel
    def convert_df_to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="BaoCaoChamCong")
        processed_data = output.getvalue()
        return processed_data


    excel_data = convert_df_to_excel(df_filtered)

    with col_dl1:
        st.download_button(
            label="📥 Tải Xuống Báo Cáo Excel",
            data=excel_data,
            file_name=f"BaoCao_ChamCong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col_dl2:
        # Báo cáo dạng text thuần/CSV giả lập cho PDF download đơn giản hoặc có thể dùng thư viện reportlab
        pdf_data = df_filtered.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Tải Xuống Báo Cáo (PDF/CSV Định Dạng Chuẩn)",
            data=pdf_data,
            file_name=f"BaoCao_ChamCong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
        )

    # --- TÍCH HỢP GEMINI AI PHÂN TÍCH ---
    st.markdown("---")
    st.subheader("🤖 Trợ Lý Gemini AI - Phân Tích Đánh Giá Nhân Sự")
    if st.button("Phân tích tự động bằng Gemini AI"):
        if not gemini_api_key:
            st.warning("Vui lòng nhập Gemini API Key ở thanh bên cạnh (Sidebar).")
        else:
            try:
                import google.generativeai as genai

                genai.configure(api_key=gemini_api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")

                summary_text = f"""
                Hãy đóng vai trò là chuyên gia nhân sự và phân tích dữ liệu chấm công sau:
                - Tổng số nhân sự: {total_nv}
                - Số lượng đi trễ: {di_tre}
                - Số lượng về sớm: {ve_som}
                - Không bấm thẻ/Vắng: {khong_bam_the}
                - Số lượng tăng ca: {tang_ca}
                
                Hãy đưa ra nhận xét ngắn gọn, sắc sảo về tình hình chấp hành kỷ luật lao động và đề xuất các giải pháp cải thiện năng suất, quản lý thời gian cho doanh nghiệp.
                """
                response = model.generate_content(summary_text)
                st.success("Kết quả phân tích từ Gemini AI:")
                st.write(response.text)
            except Exception as e:
                st.error(f"Lỗi kết nối hoặc xử lý Gemini AI: {e}")

else:
    st.info(
        "👈 Vui lòng tải lên đầy đủ 4 file Excel ở thanh bên trái (Sidebar) để khởi chạy hệ thống Dashboard."
    )
