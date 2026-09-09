import io
import os
import time
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
from google import genai
import streamlit.components.v1 as components

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="考勤统计仪表盘 / Dashboard Thống Kê Chấm Công",
    page_icon="📊",
    layout="wide"
)

# Cấu hình API Key Gemini ở Sidebar (Không xuất hiện chữ AI trên giao diện)
st.sidebar.header("⚙️ 系统配置 / Cấu hình hệ thống")
api_key = st.sidebar.text_input("Nhập API Key", type="password")

if api_key:
    os.environ["GEMINI_API_KEY"] = api_key

# Tiêu đề Dashboard song ngữ Trung - Việt
st.title("📊 员工考勤与统计仪表盘")
st.markdown("### Dashboard Thống Kê & Phân Tích Chấm Công Nhân Sự")

# --- PHẦN 1: TẢI CÁC FILE DỮ LIỆU LÊN DASHBOARD ---
st.subheader("📁 1. 数据文件上传 / Tải lên tệp dữ liệu")
col1, col2 = st.columns(2)

with col1:
    uploaded_fingerprint = st.file_uploader("1) 上传指纹打卡 Excel 文件 / Tải file Excel bấm vân tay", type=["xlsx", "xls", "csv"])
    uploaded_schedule = st.file_uploader("2) 上传排班表文件 (仅限工人) / Tải file lịch xếp ca (cho công nhân)", type=["xlsx", "xls", "csv"])

with col2:
    uploaded_staff_vp = st.file_uploader("3) 上传办公室员工名单 (NV) / Tải file danh sách nhân viên VP", type=["xlsx", "xls", "csv"])
    uploaded_staff_cn = st.file_uploader("4) 上传工人名单 (CN) / Tải file danh sách công nhân", type=["xlsx", "xls", "csv"])

# Hàm tự động tìm dòng tiêu đề thật
def load_smart_excel(uploaded_file):
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                return pd.read_csv(uploaded_file)
            else:
                df_raw = pd.read_excel(uploaded_file, header=None)
                header_row = 0
                for idx, row in df_raw.iterrows():
                    row_str = " ".join([str(val).lower() for val in row.values])
                    if 'mã' in row_str or 'ngày' in row_str or 'nhân viên' in row_str:
                        header_row = idx
                        break
                return pd.read_excel(uploaded_file, header=header_row)
        except Exception as e:
            st.error(f"文件读取错误 / Lỗi đọc file: {e}")
    return None

df_fp = load_smart_excel(uploaded_fingerprint)
df_sc = load_smart_excel(uploaded_schedule)
df_vp = load_smart_excel(uploaded_staff_vp)
df_cn = load_smart_excel(uploaded_staff_cn)


# --- PHẦN 2: CHỌN NGÀY VÀ PHÂN TÍCH THEO NGÀY ---
st.subheader("📅 2. 选择考勤核对日期 / Chọn ngày kiểm tra từ file vân tay")

selected_date_str = ""
df_fp_filtered = None

if df_fp is not None:
    df_fp.columns = df_fp.columns.astype(str).str.strip()
    date_col = next((col for col in df_fp.columns if any(k in col.lower() for k in ['ngày', 'date', 'ngay'])), None)
    
    if date_col:
        df_fp['Ngày_Clean'] = pd.to_datetime(df_fp[date_col], errors='coerce').dt.date
        available_dates = sorted(df_fp['Ngày_Clean'].dropna().unique())
        
        if available_dates:
            selected_date = st.selectbox("选择文件中的日期 / Chọn ngày có trong file vân tay", available_dates)
            selected_date_str = str(selected_date)
            
            df_fp_filtered = df_fp[df_fp['Ngày_Clean'] == selected_date].copy()
            
            col_vao = next((c for c in df_fp_filtered.columns if 'vào' in c.lower() or 'vao' in c.lower()), None)
            col_ra = next((c for c in df_fp_filtered.columns if 'ra' in c.lower()), None)
            col_dept = next((c for c in df_fp_filtered.columns if 'phòng' in c.lower() or 'bộ phận' in c.lower()), None)
            
            def has_time(val):
                if pd.isna(val):
                    return False
                s = str(val).strip().lower()
                return s not in ['', 'nan', 'none', 'nat', '-', '0:00:00']

            df_fp_filtered['Co_Vao'] = df_fp_filtered[col_vao].apply(has_time) if col_vao else False
            df_fp_filtered['Co_Ra'] = df_fp_filtered[col_ra].apply(has_time) if col_ra else False
            
            def classify_attendance(row):
                if row['Co_Vao'] and row['Co_Ra']:
                    return 'Có mặt đủ giờ / 出勤正常'
                elif row['Co_Vao'] and not row['Co_Ra']:
                    return 'Thiếu giờ ra (BR) / 缺下班卡'
                elif not row['Co_Vao'] and row['Co_Ra']:
                    return 'Thiếu giờ vào (BV) / 缺上班卡'
                else:
                    return 'Vắng / Không bấm thẻ / 缺勤'

            df_fp_filtered['Trạng Thái'] = df_fp_filtered.apply(classify_attendance, axis=1)
            st.success(f"✅ 已选择日期 / Đã chọn ngày: **{selected_date_str}**")
        else:
            st.warning("⚠️ Không nhận diện được dữ liệu ngày trong file.")
    else:
        st.error(f"❌ Không tìm thấy cột 'Ngày' trong file vân tay.")
else:
    st.info("ℹ️ Vui lòng tải file Excel bấm vân tay lên ở bước 1.")


# --- PHẦN 3: THỐNG KÊ BIỂU ĐỒ ---
st.subheader("📈 3. 当日考勤数据统计图表 / Biểu đồ thống kê theo ngày chọn")

if df_fp_filtered is not None:
    total_day_records = len(df_fp_filtered)
    count_full = (df_fp_filtered['Trạng Thái'] == 'Có mặt đủ giờ / 出勤正常').sum()
    count_missing = (df_fp_filtered['Trạng Thái'].isin(['Thiếu giờ ra (BR) / 缺下班卡', 'Thiếu giờ vào (BV) / 缺上班卡'])).sum()
    count_absent = (df_fp_filtered['Trạng Thái'] == 'Vắng / Không bấm thẻ / 缺勤').sum()
    count_active = count_full + count_missing

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("当日总人数 / Tổng nhân sự trong ngày", f"{total_day_records} 人")
    m2.metric("实际出勤 / Có đi làm trong ngày", f"{count_active} 人")
    m3.metric("缺卡人数 / Thiếu giờ (BV/BR)", f"{count_missing} 人")
    m4.metric("缺勤人数 / Vắng mặt trong ngày", f"{count_absent} 人")

    col_c1, col_c2 = st.columns(2)

    with col_c1:
        st.markdown("#### 考勤状态分布 / Tỷ lệ trạng thái đi làm trong ngày")
        status_df = df_fp_filtered['Trạng Thái'].value_counts().reset_index()
        status_df.columns = ['Trạng thái', 'Số lượng']
        fig_pie = px.pie(status_df, values='Số lượng', names='Trạng thái', hole=0.45)
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_c2:
        st.markdown("#### 部门出勤对比 / So sánh đi làm theo bộ phận")
        if col_dept and col_dept in df_fp_filtered.columns:
            dept_series = df_fp_filtered[col_dept].astype(str)
        else:
            dept_series = pd.Series(["Chung"] * len(df_fp_filtered))

        df_fp_filtered['Bộ Phận Chuẩn'] = dept_series.apply(
            lambda x: 'Văn Phòng / 办公室' if any(k in x.lower() for k in ['văn phòng', 'vp', 'office']) else 'Công Nhân / 工人'
        )

        dept_summary = df_fp_filtered.groupby(['Bộ Phận Chuẩn', 'Trạng Thái']).size().reset_index(name='Số lượng')
        fig_bar = px.bar(dept_summary, x='Bộ Phận Chuẩn', y='Số lượng', color='Trạng Thái', barmode='group', text_auto=True)
        st.plotly_chart(fig_bar, use_container_width=True)
else:
    st.info("Vui lòng tải file vân tay để hiển thị biểu đồ thống kê.")


# --- PHẦN 4: ĐỐI CHIẾU THÔNG MINH ---
st.subheader("📋 4. 智能数据核对与分析 / Phân tích & Đối chiếu chuyên sâu")

if st.button("🚀 开始考勤核对分析 / Chạy phân tích chấm công", type="primary"):
    if not api_key:
        st.warning("请输入 API Key 以继续 / Vui lòng nhập API Key ở thanh bên.")
    elif df_fp_filtered is None or df_vp is None or df_cn is None:
        st.warning("请完整上传文件并选择有效日期 / Vui lòng tải đủ file và chọn ngày hợp lệ.")
    else:
        with st.spinner("系统正在智能核对排班与考勤数据，请稍候..."):
            try:
                client = genai.Client()
                
                fp_data = df_fp_filtered.to_string()
                vp_data = df_vp.to_string() if df_vp is not None else "Không có"
                cn_data = df_cn.to_string() if df_cn is not None else "Không có"
                sc_data = df_sc.to_string() if df_sc is not None else "Không có"

                prompt = f"""
                Bạn là hệ thống đối chiếu nhân sự và chấm công tự động thông minh. 
                Hãy thực hiện đối chiếu và phân tích dữ liệu CHO ĐÚNG NGÀY: {selected_date_str}.
                
                DỮ LIỆU ĐẦU VÀO:
                1. File bấm vân tay:
                {fp_data}
                2. Danh sách NV Văn Phòng (VP):
                {vp_data}
                3. Danh sách Công Nhân (CN):
                {cn_data}
                4. Lịch xếp ca CN:
                {sc_data}
                
                QUY TẮC: 
                - VP chuẩn 8h-17h, Chủ Nhật nghỉ. 
                - Mã 575 (7h-15h), 749 & 949 (7h-19h).
                - CN theo lịch ca N (7h-19h) hoặc Đ (19h-7h hôm sau). Thiếu vào = "BV", thiếu ra = "BR", thiếu cả = "Vắng".
                - TRỌNG TÂM: Chỉ đưa ra các trường hợp bất thường ("đi trễ", "về sớm", "làm không đúng lịch", "vắng", "BV", "BR", "Lễ").
                
                YÊU CẦU ĐẦU RA (TABLE MARKDOWN):
                Mã NV | Họ và Tên | Giờ vào | Giờ ra | Giờ làm thực tế | Ghi chú
                """

                max_retries = 3
                response = None
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model='gemini-1.5-flash',
                            contents=prompt,
                        )
                        break
                    except Exception as err:
                        if "503" in str(err) and attempt < max_retries - 1:
                            time.sleep(3)
                            continue
                        else:
                            raise err

                st.session_state['analysis_result'] = response.text
                st.session_state['result_date'] = selected_date_str
                st.success("✅ 考勤核对完成！ / Hoàn tất đối chiếu chấm công!")
                
            except Exception as e:
                st.error(f"处理出错 / Lỗi xử lý: {e}")

if 'analysis_result' in st.session_state:
    res_date = st.session_state.get('result_date', selected_date_str)
    st.markdown(f"### 📋 异常考勤与核对结果 ({res_date})")
    st.markdown(st.session_state['analysis_result'])


# --- PHẦN 5: XUẤT BÁO CÁO EXCEL VÀ IN PDF GIAO DIỆN ---
st.subheader("💾 5. 导出报告 / Xuất báo cáo")

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if df_fp_filtered is not None:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            col_id = next((c for c in df_fp_filtered.columns if 'mã' in c.lower() or 'id' in c.lower()), 'Mã Nhân Viên')
            col_name = next((c for c in df_fp_filtered.columns if 'tên' in c.lower() or 'name' in c.lower()), 'Tên nhân viên')
            col_in = next((c for c in df_fp_filtered.columns if 'vào' in c.lower()), 'Giờ vào')
            col_out = next((c for c in df_fp_filtered.columns if 'ra' in c.lower()), 'Giờ ra')
            col_total = next((c for c in df_fp_filtered.columns if 'tổng' in c.lower()), 'Tổng giờ')

            df_export = pd.DataFrame({
                "Mã NV": df_fp_filtered[col_id] if col_id in df_fp_filtered else "",
                "Họ và Tên": df_fp_filtered[col_name] if col_name in df_fp_filtered else "",
                "Giờ vào": df_fp_filtered[col_in] if col_in in df_fp_filtered else "",
                "Giờ ra": df_fp_filtered[col_out] if col_out in df_fp_filtered else "",
                "Giờ làm thực tế": df_fp_filtered[col_total] if col_total in df_fp_filtered else "",
                "Ghi chú": df_fp_filtered['Trạng Thái']
            })
            df_export.to_excel(writer, index=False, sheet_name=f'ChamCong_{selected_date_str}')
        excel_data = output.getvalue()
        
        st.download_button(
            label=f"📥 下载 Excel 报告 / Tải xuống file Excel",
            data=excel_data,
            file_name=f"Attendance_Report_{selected_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Vui lòng chọn ngày để tải báo cáo Excel.")

with col_dl2:
    st.markdown("#### 🖨️ 导出 PDF 界面 / Lưu trang thành PDF giao diện")
    # Tích hợp nút kích hoạt hộp thoại In/Lưu PDF y chang giao diện màn hình
    pdf_button_html = """
    <button onclick="window.print()" style="
        background-color: #ff4b4b;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        font-size: 16px;
        font-weight: bold;
        width: 100%;
    ">
        📄 打印 / 保存为 PDF (Lưu y chang giao diện Dashboard)
    </button>
    """
    components.html(pdf_button_html, height=50)
