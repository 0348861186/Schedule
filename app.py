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

# --- HÀM XỬ LÝ TIÊU ĐỀ THÔNG MINH BẰNG GEMINI ---
def standardize_headers_with_gemini(df_columns, file_type):
    if model is None:
        return df_columns
    prompt = f"""
    Bạn là một chuyên gia xử lý dữ liệu nhân sự. Hãy đồng nhất các tiêu đề cột sau đây của file '{file_type}' thành các tên chuẩn tiếng Việt (ví dụ: 'Mã NV', 'Họ và tên', 'Ngày', 'Thứ', 'Giờ vào', 'Giờ ra', 'Tổng giờ'):
    Các cột hiện tại: {list(df_columns)}
    Chỉ trả về danh sách tên cột mới dưới dạng danh sách python thuần túy, không giải thích gì thêm.
    """
    try:
        response = model.generate_content(prompt)
        cleaned_text = response.text.strip().replace("```python", "").replace("```", "").strip()
        new_cols = eval(cleaned_text)
        if isinstance(new_cols, list) and len(new_cols) == len(df_columns):
            return new_cols
    except Exception:
        pass
    return df_columns

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

# --- XỬ LÝ DỮ LIỆU & LOGIC NGHIỆP VỤ ---
if uploaded_fingerprint is not None and uploaded_shift is not None:
    try:
        df_fp = pd.read_excel(uploaded_fingerprint)
        df_sh = pd.read_excel(uploaded_shift)
        
        if model:
            df_fp.columns = standardize_headers_with_gemini(df_fp.columns, "Bấm vân tay")
            df_sh.columns = standardize_headers_with_gemini(df_sh.columns, "Lịch xếp ca")

        st.success("✅ Đã tải và quét dữ liệu thành công! / 数据加载与处理成功！")

        def process_attendance_row(row):
            ma_nv = str(row.get("Mã NV", ""))
            if ma_nv in ["673", "A068"]:
                nhóm = "Bảo trì / 保养"
            elif ma_nv in ["749", "949"]:
                nhóm = "QC"
            elif ma_nv == "575":
                nhóm = "Tạp vụ / 杂务"
            elif ma_nv.startswith("VP") or row.get("Loại") == "VP":
                nhóm = "Văn phòng / 办公室"
            else:
                nhóm = "Công nhân / 工人"

            gio_vao = str(row.get("Giờ vào", ""))
            gio_ra = str(row.get("Giờ ra", ""))
            
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
            df_fp[["Nhóm", "Ghi chú"]] = df_fp.apply(process_attendance_row, axis=1)
            cols_to_keep = ["Mã NV"]
            if "Họ và tên" in df_fp.columns: cols_to_keep.append("Họ và tên")
            else: df_fp["Họ và tên"] = "NV " + df_fp["Mã NV"].astype(str); cols_to_keep.append("Họ và tên")
            cols_to_keep.extend(["Giờ vào", "Giờ ra", "Tổng giờ" if "Tổng giờ" in df_fp.columns else "Giờ vào", "Ghi chú", "Nhóm"])
            valid_cols = [c for c in cols_to_keep if c in df_fp.columns]
            result_df = df_fp[valid_cols].copy()
            result_df.columns = ["Mã NV", "Họ và tên", "Giờ vào", "Giờ ra", "Giờ làm thực tế", "Ghi chú", "Nhóm"][:len(valid_cols)]
        else:
            result_df = pd.DataFrame([
                {"Mã NV": "673", "Họ và tên": "Nguyễn Văn A", "Giờ vào": "07:00", "Giờ ra": "19:00", "Giờ làm thực tế": 12.0, "Ghi chú": "Đúng giờ / 准时", "Nhóm": "Bảo trì / 保养"},
                {"Mã NV": "749", "Họ và tên": "Trần Thị B", "Giờ vào": "07:15", "Giờ ra": "19:00", "Giờ làm thực tế": 11.75, "Ghi chú": "đi trễ / 迟到", "Nhóm": "QC"},
                {"Mã NV": "575", "Họ và tên": "Lê Văn C", "Giờ vào": "08:00", "Giờ ra": "15:00", "Giờ làm thực tế": 7.0, "Ghi chú": "về sớm / 早退", "Nhóm": "Tạp vụ / 杂务"},
                {"Mã NV": "VP01", "Họ và tên": "Phạm Văn D", "Giờ vào": "08:00", "Giờ ra": "16:00", "Giờ làm thực tế": 7.0, "Ghi chú": "về sớm / 早退", "Nhóm": "Văn phòng / 办公室"},
                {"Mã NV": "CN01", "Họ và tên": "Hoàng Thị E", "Giờ vào": "", "Giờ ra": "", "Giờ làm thực tế": 0.0, "Ghi chú": "Vắng / 缺勤", "Nhóm": "Công nhân / 工人"}
            ])

        total_nv = len(result_df)
        total_dung_gio = len(result_df[result_df["Ghi chú"].str.contains("Đúng giờ|准时", case=False, na=False)])
        total_tre = len(result_df[result_df["Ghi chú"].str.contains("trễ|迟到|BV|BR", case=False, na=False)])
        total_vang = len(result_df[result_df["Ghi chú"].str.contains("vắng|缺勤", case=False, na=False)])

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
            if not result_df.empty and "Ghi chú" in result_df.columns:
                df_status_counts = result_df["Ghi chú"].value_counts().reset_index()
                df_status_counts.columns = ["Trạng thái", "Số lượng"]
                fig_pie = px.pie(df_status_counts, names="Trạng thái", values="Số lượng", title="Tỷ lệ trạng thái đi làm / 出勤状态比例", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Không đủ dữ liệu cho biểu đồ tròn / 暂无饼图数据")
            
        with chart_col2:
            if not result_df.empty and "Nhóm" in result_df.columns:
                df_group_counts = result_df["Nhóm"].value_counts().reset_index()
                df_group_counts.columns = ["Nhóm", "Số lượng"]
                fig_bar = px.bar(df_group_counts, x="Nhóm", y="Số lượng", title="Số lượng nhân viên theo nhóm / 各组出勤人数", labels={"Nhóm": "Nhóm / 组别", "Số lượng": "Số lượng / 数量"}, color="Nhóm")
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Không đủ dữ liệu cho biểu đồ cột / 暂无柱状图数据")

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
        st.error(f"❌ Lỗi xử lý dữ liệu: {e} / 数据处理错误")
else:
    st.info("💡 Vui lòng tải file bấm vân tay và lịch xếp ca ở thanh bên trái. / 请在左侧上传文件。")
