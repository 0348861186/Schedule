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
st.markdown(
    "Tích hợp AI phân tích chuyên sâu lịch trình chấm công và hiệu suất."
)

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

if (
    uploaded_cham_cong
    and uploaded_cn
    and uploaded_vp
    and uploaded_lich
):
  try:
    df_cc_raw = pd.read_excel(uploaded_cham_cong)
    df_cn_raw = pd.read_excel(uploaded_cn)
    df_vp_raw = pd.read_excel(uploaded_vp)
    df_lich_raw = pd.read_excel(uploaded_lich)
  except Exception as e:
    st.error(f"Lỗi đọc file Excel: {e}")
    st.stop()

  # --- CẤU HÌNH ÁNH XẠ CỘT LINH HOẠT (GIÚP TRÁNH MỌI LỖI KEYERROR) ---
  st.sidebar.markdown("---")
  st.sidebar.header("⚙️ Tùy Chỉnh Cột Dữ Liệu")

  with st.sidebar.expander("🔗 Khớp cột File Chấm Công", expanded=True):
    cols_cc = list(df_cc_raw.columns)
    c_manv_cc = st.selectbox("Cột Mã NV (Chấm công)", cols_cc, index=0)
    c_hoten_cc = st.selectbox(
        "Cột Họ Tên (Chấm công - nếu có)",
        cols_cc,
        index=min(1, len(cols_cc) - 1),
    )
    c_ngay_cc = st.selectbox(
        "Cột Ngày", cols_cc, index=min(2, len(cols_cc) - 1)
    )
    c_vao_cc = st.selectbox(
        "Cột Giờ Vào", cols_cc, index=min(3, len(cols_cc) - 1)
    )
    c_ra_cc = st.selectbox(
        "Cột Giờ Ra", cols_cc, index=min(4, len(cols_cc) - 1)
    )

  with st.sidebar.expander("🔗 Khớp cột File Công Nhân"):
    cols_cn = list(df_cn_raw.columns)
    c_manv_cn = st.selectbox("Cột Mã NV (Công nhân)", cols_cn, index=0)
    c_hoten_cn = st.selectbox(
        "Cột Họ Tên (Công nhân)", cols_cn, index=min(1, len(cols_cn) - 1)
    )

  with st.sidebar.expander("🔗 Khớp cột File Văn Phòng"):
    cols_vp = list(df_vp_raw.columns)
    c_manv_vp = st.selectbox("Cột Mã NV (Văn phòng)", cols_vp, index=0)
    c_hoten_vp = st.selectbox(
        "Cột Họ Tên (Văn phòng)", cols_vp, index=min(1, len(cols_vp) - 1)
    )

  with st.sidebar.expander("🔗 Khớp cột File Lịch Ca"):
    cols_lich = list(df_lich_raw.columns)
    c_manv_lich = st.selectbox("Cột Mã NV (Lịch ca)", cols_lich, index=0)
    c_ngay_lich = st.selectbox(
        "Cột Ngày (Lịch ca)", cols_lich, index=min(1, len(cols_lich) - 1)
    )
    c_ca_lich = st.selectbox(
        "Cột Ca làm việc", cols_lich, index=min(2, len(cols_lich) - 1)
    )

  # Chuẩn hóa DataFrames dựa trên lựa chọn của người dùng
  df_cc = pd.DataFrame({
      "MaNV": df_cc_raw[c_manv_cc],
      "HoTen": df_cc_raw[c_hoten_cc],
      "Ngay": df_cc_raw[c_ngay_cc],
      "ThoiGianVao": df_cc_raw[c_vao_cc],
      "ThoiGianRa": df_cc_raw[c_ra_cc],
  })

  df_cn = pd.DataFrame(
      {"MaNV": df_cn_raw[c_manv_cn], "HoTen": df_cn_raw[c_hoten_cn]}
  )
  df_cn["Nhom"] = "Công Nhân"

  df_vp = pd.DataFrame(
      {"MaNV": df_vp_raw[c_manv_vp], "HoTen": df_vp_raw[c_hoten_vp]}
  )
  df_vp["Nhom"] = "Văn Phòng"

  df_lich = pd.DataFrame({
      "MaNV": df_lich_raw[c_manv_lich],
      "Ngay": df_lich_raw[c_ngay_lich],
      "Ca": df_lich_raw[c_ca_lich],
  })

  df_nv_all = pd.concat([df_cn, df_vp], ignore_index=True)

  # --- XỬ LÝ GỘP DỮ LIỆU & LOGIC ---
  # Gộp thông tin nhân viên vào bảng chấm công qua MaNV
  df_merged = pd.merge(
      df_cc,
      df_nv_all[["MaNV", "HoTen", "Nhom"]].drop_duplicates(subset=["MaNV"]),
      on="MaNV",
      how="left",
      suffixes=("", "_nv"),
  )
  if "HoTen_nv" in df_merged.columns:
    df_merged["HoTen"] = df_merged["HoTen_nv"].fillna(
        df_merged.get("HoTen", "Không rõ")
    )
    df_merged = df_merged.drop(columns=["HoTen_nv"])

  if "Nhom" not in df_merged.columns:
    df_merged["Nhom"] = "Khác"

  # Gộp lịch ca
  df_merged = pd.merge(
      df_merged,
      df_lich[["MaNV", "Ngay", "Ca"]],
      on=["MaNV", "Ngay"],
      how="left",
  )
  df_merged["Ca"] = df_merged["Ca"].fillna("HC")


  # Xử lý logic tính toán giờ vào/ra, đi trễ, về sớm, tăng ca
  def calculate_attendance(row):
    nhom = row.get("Nhom", "Khác")
    ca = row.get("Ca", "HC")
    gio_vao = row.get("ThoiGianVao")
    gio_ra = row.get("ThoiGianRa")

    ghi_chu = "BT"
    gio_lam_viec = 0.0

    if pd.isna(gio_vao) or pd.isna(gio_ra):
      return pd.Series(["Vắng / Không bấm thẻ", "Chưa bấm thẻ", 0.0, "Vắng"])

    try:
      t_vao = pd.to_datetime(str(gio_vao)).time()
      t_ra = pd.to_datetime(str(gio_ra)).time()
    except:
      return pd.Series(["Lỗi định dạng giờ", "Lỗi", 0.0, "Lỗi"])

    dt_vao = datetime.combine(datetime.today(), t_vao)
    dt_ra = datetime.combine(datetime.today(), t_ra)
    if dt_ra < dt_vao:  # Ca qua đêm
      dt_ra += pd.Timedelta(days=1)

    gio_lam_viec = (dt_ra - dt_vao).seconds / 3600.0
    if gio_lam_viec > 5:
      gio_lam_viec -= 1.0  # Trừ giờ nghỉ trưa/giữa ca

    if str(nhom).strip() == "Công Nhân":
      if str(ca).strip().upper() in ["Đ", "ĐÊM"]:
        if t_vao > pd.to_datetime("22:15:00").time():
          ghi_chu = "Đi trễ"
      else:
        if t_vao > pd.to_datetime("08:15:00").time():
          ghi_chu = "Đi trễ"
        if t_ra < pd.to_datetime("17:00:00").time():
          ghi_chu = "Về sớm" if ghi_chu == "BT" else "Đi trễ & Về sớm"
    else:
      if gio_lam_viec < 8.0:
        ghi_chu = "Về sớm / Thiếu giờ"
      elif gio_lam_viec > 8.5:
        ghi_chu = "Tăng ca"
      else:
        ghi_chu = "BT"

    return pd.Series(["Bình Thường", ghi_chu, round(gio_lam_viec, 2), ca])


  df_merged[
      ["TrangThaiChamCong", "DanhGia", "TongGioLam", "CaLamViec"]
  ] = df_merged.apply(calculate_attendance, axis=1)

  # --- BỘ LỌC TRÊN DASHBOARD ---
  st.sidebar.markdown("---")
  st.sidebar.header("🔍 Bộ Lọc Phân Tích")
  unique_dates = sorted(df_merged["Ngay"].dropna().astype(str).unique())
  selected_date = st.sidebar.selectbox(
      "Chọn Ngày Phân Tích", ["Tất cả"] + list(unique_dates)
  )

  unique_nhom = list(df_merged["Nhom"].dropna().unique())
  selected_nhom = st.sidebar.selectbox(
      "Chọn Nhóm Nhân Sự", ["Tất cả"] + unique_nhom
  )

  # Áp dụng bộ lọc
  df_filtered = df_merged.copy()
  if selected_date != "Tất cả":
    df_filtered = df_filtered[df_filtered["Ngay"].astype(str) == selected_date]
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
    st.markdown("#### 📉 Tổng Giờ Làm Theo Nhân Sự")
    if not df_filtered.empty:
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


  def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df.to_excel(writer, index=False, sheet_name="BaoCaoChamCong")
    processed_data = output.getvalue()
    return processed_data


  if not df_filtered.empty:
    excel_data = convert_df_to_excel(df_filtered)
    with col_dl1:
      st.download_button(
          label="📥 Tải Xuống Báo Cáo Excel",
          data=excel_data,
          file_name=f"BaoCao_ChamCong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

    with col_dl2:
      csv_data = df_filtered.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="📥 Tải Xuống Báo Cáo (CSV Chuẩn)",
          data=csv_data,
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
      "👈 Vui lòng tải lên đầy đủ 4 file Excel ở thanh bên trái (Sidebar) để"
      " khởi chạy hệ thống Dashboard."
  )
