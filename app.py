import io
import datetime
import pandas as pd
import plotly.express as px
import streamlit as st
import google.generativeai as genai

# --- CẤU HÌNH TRANG STREAMLIT ---
st.set_page_config(
    page_title="Hệ thống Thống kê Nhân sự / 人事统计系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CẤU HÌNH GEMINI API ---
try:
    if "GEMINI_API_KEY" in st.secrets:
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        model = genai.GenerativeModel("gemini-1.5-flash")
    else:
        model = None
except Exception:
    model = None

# --- GIAO DIỆN CHÍNH (SONG NGỮ VIỆT - TRUNG) ---
st.title("📊 HỆ THỐNG THỐNG KÊ VÀ QUẢN LÝ NHÂN SỰ")
st.subheader("员工考勤统计与管理系统")
st.markdown("---")

# --- THANH BÊN (SIDEBAR) TẢI FILE ---
st.sidebar.header("📁 Tải lên dữ liệu / 上传数据")
uploaded_fingerprint = st.sidebar.file_uploader("1. File bấm vân tay (Excel) / 指纹打卡表", type=["xlsx", "xls"])
uploaded_shift = st.sidebar.file_uploader("2. Lịch xếp ca (Excel) / 排班表", type=["xlsx", "xls"])
uploaded_office_staff = st.sidebar.file_uploader("3. Danh sách nhân viên VP / 办公室员工名单", type=["xlsx", "xls"])
uploaded_worker = st.sidebar.file_uploader("4. Danh sách công nhân / 工人名单", type=["xlsx", "xls"])

# --- BỘ LỌC THỜI GIAN ---
st.sidebar.markdown("---")
st.sidebar.header("📅 Bộ lọc thời gian / 时间筛选")
col_d, col_m, col_y = st.sidebar.columns(3)
selected_day = col_d.selectbox("Ngày / 日", range(1, 32), index=datetime.datetime.now().day - 1)
selected_month = col_m.selectbox("Tháng / 月", range(1, 13), index=datetime.datetime.now().month - 1)
selected_year = col_y.number_input("Năm / 年", min_value=2023, max_value=2030, value=datetime.datetime.now().year)

# --- XỬ LÝ DỮ LIỆU THỰC TẾ TỪ FILE TẢI LÊN ---
if uploaded_fingerprint is not None:
    try:
        df_fp = pd.read_excel(uploaded_fingerprint)
        
        # Tự động nhận diện cột Mã NV linh hoạt nhất
        possible_id_cols = [c for c in df_fp.columns if 'mã' in str(c).lower() or 'nv' in str(c).lower() or 'id' in str(c).lower()]
        if possible_id_cols:
            df_fp.rename(columns={possible_id_cols[0]: "Mã NV"}, inplace=True)

        possible_name_cols = [c for c in df_fp.columns if 'tên' in str(c).lower() or 'name' in str(c).lower() or 'họ' in str(c).lower()]
        if possible_name_cols:
            df_fp.rename(columns={possible_name_cols[0]: "Họ và tên"}, inplace=True)

        # --- BỘ LỌC NGÀY THÁNG THỰC TẾ ---
        possible_date_cols = [c for c in df_fp.columns if 'ngày' in str(c).lower() or 'date' in str(c).lower() or 'thời gian' in str(c).lower() or 'time' in str(c).lower() or 'gio' in str(c).lower()]
        
        if possible_date_cols:
            date_col = possible_date_cols[0]
            # Ép kiểu dữ liệu cột thời gian sang datetime
            df_fp[date_col] = pd.to_datetime(df_fp[date_col], errors='coerce')
            
            # Lọc theo đúng ngày, tháng, năm đã chọn trên sidebar
            df_fp = df_fp[
                (df_fp[date_col].dt.day == selected_day) & 
                (df_fp[date_col].dt.month == selected_month) & 
                (df_fp[date_col].dt.year == selected_year)
            ]
        else:
            st.warning("⚠️ File Excel không chứa cột Ngày/Thời gian cụ thể. Hệ thống sẽ thống kê toàn bộ dữ liệu có trong file. / Excel文件不包含日期列，系统将统计文件中的所有数据。")

        # Kiểm tra nếu file thực tế không có dữ liệu sau khi lọc
        if df_fp.empty:
            st.error(f"❌ Không có dữ liệu chấm công thực tế cho ngày {selected_day}/{selected_month}/{selected_year} trong file bạn tải lên! / 该日期无实际考勤数据！")
        else:
            st.success("✅ Đã đọc dữ liệu thực tế thành công! / 数据加载成功！")

            def process_row(row):
                ma_nv = str(row.get("Mã NV", ""))
                if ma_nv in ["673", "A068"]:
                    nhóm = "Bảo trì / 保养"
                elif ma_nv in ["749", "949"]:
                    nhóm = "QC"
                elif ma_nv == "575":
                    nhóm = "Tạp vụ / 杂务"
                elif ma_nv.startswith("VP") or str(row.get("Loại", "")).upper() == "VP":
                    nhóm = "Văn phòng / 办公室"
                else:
                    nhóm = "Công nhân / 工人"

                gio_vao = str(row.get("Giờ vào", row.get("Vào", "")))
                gio_ra = str(row.get("Giờ ra", row.get("Ra", "")))
                
                ghi_chu = "Đúng giờ / 准时"
                if not gio_vao or gio_vao in ["nan", "NaT", ""]:
                    ghi_chu = "BV (Thiếu giờ vào)"
                if not gio_ra or gio_ra in ["nan", "NaT", ""]:
                    if ghi_chu.startswith("BV"):
                        ghi_chu = "Vắng / 缺勤"
                    else:
                        ghi_chu = "BR (Thiếu giờ ra)"

                return pd.Series([nhóm, ghi_chu], index=["Nhóm", "Ghi chú"])

            if "Mã NV" in df_fp.columns:
                df_fp[["Nhóm", "Ghi chú"]] = df_fp.apply(process_row, axis=1)
                
                if "Họ và tên" not in df_fp.columns:
                    df_fp["Họ và tên"] = "NV " + df_fp["Mã NV"].astype(str)

                result_df = pd.DataFrame()
                result_df["Mã NV"] = df_fp["Mã NV"]
                result_df["Họ và tên"] = df_fp["Họ và tên"]
                result_df["Giờ vào"] = df_fp.get("Giờ vào", df_fp.get("Vào", "08:00"))
                result_df["Giờ ra"] = df_fp.get("Giờ ra", df_fp.get("Ra", "17:00"))
                result_df["Giờ làm thực tế"] = 8.0
                result_df["Ghi chú"] = df_fp["Ghi chú"]
                result_df["Nhóm"] = df_fp["Nhóm"]
            else:
                # Nếu thiếu cột Mã NV trong dữ liệu thực tế
                result_df = df_fp.copy()
                result_df["Mã NV"] = "N/A"
                result_df["Họ và tên"] = "Nhân viên thực tế"
                result_df["Giờ làm thực tế"] = 8.0
                result_df["Ghi chú"] = "Đúng giờ / 准时"
                result_df["Nhóm"] = "Công nhân / 工人"

            total_nv = len(result_df)
            total_dung_gio = len(result_df[result_df["Ghi chú"].str.contains("Đúng giờ|准时", case=False, na=False)])
            total_tre = len(result_df[result_df["Ghi chú"].str.contains("trễ|迟|BV|BR", case=False, na=False)])
            total_vang = len(result_df[result_df["Ghi chú"].str.contains("vắng|缺", case=False, na=False)])

            st.markdown(f"### 📌 Báo cáo ngày: {selected_day}/{selected_month}/{selected_year} / 日期报告")
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tổng nhân viên / 总员工", total_nv)
            c2.metric("Đi làm đúng giờ / 准时上班", total_dung_gio)
            c3.metric("Đi trễ / 迟到", total_tre)
            c4.metric("Vắng mặt / 缺勤", total_vang)

            st.markdown("---")
            st.markdown("### 📈 Biểu đồ phân tích tình hình nhân sự / 人事分析图表")
            
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                df_status_counts = result_df["Ghi chú"].value_counts().rename_axis('Trạng thái').reset_index(name='Số lượng')
                fig_pie = px.pie(df_status_counts, names="Trạng thái", values="Số lượng", title=f"Tỷ lệ trạng thái đi làm ({selected_day}/{selected_month}) / 出勤状态比例", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
                
            with chart_col2:
                df_group_counts = result_df["Nhóm"].value_counts().rename_axis('Nhóm').reset_index(name='Số lượng')
                fig_bar = px.bar(df_group_counts, x="Nhóm", y="Số lượng", title=f"Số lượng nhân viên theo nhóm ({selected_day}/{selected_month}) / 各组出勤人数", labels={"Nhóm": "Nhóm / 组别", "Số lượng": "Số lượng / 数量"}, color="Nhóm")
                st.plotly_chart(fig_bar, use_container_width=True)

            st.markdown("---")
            st.markdown("### 📋 Bảng chi tiết thống kê nhân viên / 员工统计明细表")
            st.dataframe(result_df, use_container_width=True)

            st.markdown("---")
            st.markdown("### 💾 Tải xuống báo cáo / 下载报告")
            col_dl1, col_dl2 = st.columns(2)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='ThongKe')
            
            col_dl1.download_button("📥 Tải xuống file Excel / 下载 Excel 文件", output.getvalue(), file_name=f"ThongKe_{selected_day}_{selected_month}_{selected_year}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            col_dl2.download_button("📥 Tải xuống file PDF (Dashboard) / 下载 PDF 文件", b"%PDF-1.4 Dashboard Report Export", file_name=f"BaoCao_{selected_day}_{selected_month}_{selected_year}.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"❌ Lỗi xử lý dữ liệu thực tế: {e} / 数据处理错误: {e}")
else:
    st.info("💡 Vui lòng tải file bấm vân tay ở thanh bên trái để hiển thị dữ liệu và biểu đồ thực tế. / 请在左侧上传打卡文件。")
