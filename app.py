import io
import os
import time
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
from google import genai
import streamlit.components.v1 as components

# Cấu hình trang
st.set_page_config(
    page_title="Hệ Thống Thống Kê & AI Chấm Công / 考勤统计与AI智能系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện song ngữ và các thẻ thống kê
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border-radius: 4px; padding: 10px 20px; font-weight: bold; border: 1px solid #dee2e6; }
    .stTabs [aria-selected="true"] { background-color: #0d6efd !important; color: white !important; }
    .metric-card { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border-left: 4px solid #0d6efd; }
    </style>
""", unsafe_allow_html=True)

# --- SIDEBAR: CẤU HÌNH & TẢI FILE ---
st.sidebar.header("🔑 系统配置 / Cấu hình hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key", type="password")
if api_key:
    os.environ["GEMINI_API_KEY"] = api_key

st.sidebar.markdown("---")
st.sidebar.header("📁 数据文件上传 / Tải file dữ liệu")
uploaded_fingerprint = st.sidebar.file_uploader("1) File vân tay / 指纹打卡 Excel", type=["xlsx", "xls", "csv"])
uploaded_schedule = st.sidebar.file_uploader("2) Lịch xếp ca / 排班表文件 (CN)", type=["xlsx", "xls", "csv"])
uploaded_staff_vp = st.sidebar.file_uploader("3) Danh sách NV VP / 办公室名单", type=["xlsx", "xls", "csv"])
uploaded_staff_cn = st.sidebar.file_uploader("4) Danh sách CN / 工人名单", type=["xlsx", "xls", "csv"])

# Tiêu đề chính
st.title("📊 HỆ THỐNG THỐNG KÊ & PHÂN TÍCH CHẤM CÔNG AI")
st.markdown("<h3 style='color: #6c757d;'>考勤统计与AI智能核对系统 (Song Ngữ Trung - Việt)</h3>", unsafe_allow_html=True)

# Hàm tự động tìm dòng tiêu đề thật (Smart Load Excel)
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
                    if 'mã' in row_str or 'ngày' in row_str or 'nhân viên' in row_str or 'id' in row_str:
                        header_row = idx
                        break
                return pd.read_excel(uploaded_file, header=header_row)
        except Exception as e:
            st.error(f"Lỗi đọc file / 文件读取错误: {e}")
    return None

df_fp = load_smart_excel(uploaded_fingerprint)
df_sc = load_smart_excel(uploaded_schedule)
df_vp = load_smart_excel(uploaded_staff_vp)
df_cn = load_smart_excel(uploaded_staff_cn)

if df_fp is not None:
    df_fp.columns = df_fp.columns.astype(str).str.strip()
    date_col = next((col for col in df_fp.columns if any(k in col.lower() for k in ['ngày', 'date', 'ngay'])), None)
    
    if date_col:
        df_fp['Ngày_Clean'] = pd.to_datetime(df_fp[date_col], errors='coerce').dt.date
        available_dates = sorted(df_fp['Ngày_Clean'].dropna().unique())
        
        if available_dates:
            st.sidebar.markdown("---")
            st.sidebar.header("📅 Lọc Dữ Liệu / 筛选选项")
            selected_date = st.sidebar.selectbox("Chọn ngày kiểm tra / 选择检查日期", available_dates)
            selected_date_str = str(selected_date)
            
            # --- BỔ SUNG PHẦN CHỌN KHUNG GIỜ THỐNG KÊ ---
            shift_group = st.sidebar.selectbox(
                "Chọn khung giờ thống kê / 选择统计时段",
                [
                    "Tất cả / 全部",
                    "Nhóm vào 7:00 AM / 7:00AM 入场组",
                    "Nhóm vào 19:00 PM / 19:00PM 入场组"
                ]
            )
            
            df_fp_filtered = df_fp[df_fp['Ngày_Clean'] == selected_date].copy()
            
            col_vao = next((c for c in df_fp_filtered.columns if 'vào' in c.lower() or 'vao' in c.lower()), None)
            col_ra = next((c for c in df_fp_filtered.columns if 'ra' in c.lower()), None)
            col_dept = next((c for c in df_fp_filtered.columns if 'phòng' in c.lower() or 'bộ phận' in c.lower() or 'dept' in c.lower()), None)
            
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
            
            # --- LỌC DỮ LIỆU THEO KHUNG GIỜ NẾU CÓ CHỌN ---
            if shift_group == "Nhóm vào 7:00 AM / 7:00AM 入场组" and col_vao:
                df_fp_filtered = df_fp_filtered[df_fp_filtered[col_vao].astype(str).str.contains('07:|7:', na=False)]
            elif shift_group == "Nhóm vào 19:00 PM / 19:00PM 入场组" and col_vao:
                df_fp_filtered = df_fp_filtered[df_fp_filtered[col_vao].astype(str).str.contains('19:', na=False)]
            
            # --- THỐNG KÊ SỐ LIỆU ---
            total_day_records = len(df_fp_filtered)
            count_full = (df_fp_filtered['Trạng Thái'] == 'Có mặt đủ giờ / 出勤正常').sum()
            count_missing = (df_fp_filtered['Trạng Thái'].isin(['Thiếu giờ ra (BR) / 缺下班卡', 'Thiếu giờ vào (BV) / 缺上班卡'])).sum()
            count_absent = (df_fp_filtered['Trạng Thái'] == 'Vắng / Không bấm thẻ / 缺勤').sum()
            count_active = count_full + count_missing

            # --- THIẾT KẾ GIAO DIỆN TAB ---
            tab1, tab2, tab3 = st.tabs([
                "📈 Dashboard Thống Kê / 统计看板",
                "📋 Chi Tiết & Xuất Báo Cáo / 考勤明细与导出",
                "🤖 AI Thông Minh & Đối Chiếu / AI智能核对与分析"
            ])

            with tab1:
                st.markdown(f"### 📊 Tổng Quan Hoạt Động (Ngày: {selected_date_str} - Khung giờ: {shift_group}) / 运营概览")
                m1, m2, m3, m4 = st.columns(4)
                m1.markdown(f'<div class="metric-card"><h4>Tổng Nhân Sự / 总人数</h4><h2>{total_day_records} 人</h2></div>', unsafe_allow_html=True)
                m2.markdown(f'<div class="metric-card"><h4>Thực Tế Đi Làm / 实际出勤</h4><h2>{count_active} 人</h2></div>', unsafe_allow_html=True)
                m3.markdown(f'<div class="metric-card"><h4>Thiếu Giờ / 缺卡人数</h4><h2>{count_missing} 人</h2></div>', unsafe_allow_html=True)
                m4.markdown(f'<div class="metric-card"><h4>Vắng Mặt / 缺勤人数</h4><h2>{count_absent} 人</h2></div>', unsafe_allow_html=True)
                
                st.markdown("---")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 考勤状态分布 / Tỷ lệ trạng thái đi làm")
                    status_df = df_fp_filtered['Trạng Thái'].value_counts().reset_index()
                    status_df.columns = ['Trạng thái', 'Số lượng']
                    fig_pie = px.pie(status_df, values='Số lượng', names='Trạng thái', hole=0.45)
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                with c2:
                    st.markdown("#### 部门出勤对比 / So sánh đi làm theo bộ phận")
                    if col_dept and col_dept in df_fp_filtered.columns:
                        dept_series = df_fp_filtered[col_dept].astype(str)
                    else:
                        dept_series = pd.Series(["Chung / 综合"] * len(df_fp_filtered))
                    
                    df_fp_filtered['Bộ Phận Chuẩn'] = dept_series.apply(
                        lambda x: 'Văn Phòng / 办公室' if any(k in x.lower() for k in ['văn phòng', 'vp', 'office']) else 'Công Nhân / 工人'
                    )
                    dept_summary = df_fp_filtered.groupby(['Bộ Phận Chuẩn', 'Trạng Thái']).size().reset_index(name='Số lượng')
                    fig_bar = px.bar(dept_summary, x='Bộ Phận Chuẩn', y='Số lượng', color='Trạng Thái', barmode='group', text_auto=True)
                    st.plotly_chart(fig_bar, use_container_width=True)

            with tab2:
                st.markdown(f"### 📋 Bảng Chi Tiết Chấm Công Ngày {selected_date_str} / 考勤明细表")
                st.dataframe(df_fp_filtered, use_container_width=True)
                
                st.markdown("---")
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_fp_filtered.to_excel(writer, index=False, sheet_name=f'ChamCong_{selected_date_str}')
                    st.download_button(
                        label="📥 Tải xuống file Excel / 下载 Excel 报告",
                        data=output.getvalue(),
                        file_name=f"Attendance_Report_{selected_date_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col_dl2:
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
                        📄 In / Lưu giao diện thành PDF (打印 / 保存为PDF)
                    </button>
                    """
                    components.html(pdf_button_html, height=50)

            with tab3:
                st.markdown(f"### 🤖 Phân Tích & Đối Chiếu Thông Minh Cùng Gemini AI ({selected_date_str})")
                st.info("Hệ thống sẽ sử dụng Gemini AI (SDK mới nhất) để kiểm tra chéo lịch ca, danh sách nhân viên và vân tay thực tế nhằm tìm ra các trường hợp bất thường.")
                
                if st.button("🚀 Bắt đầu chạy phân tích AI / 开始考勤核对分析", type="primary"):
                    if not api_key:
                        st.warning("Vui lòng nhập API Key ở thanh bên trái! / 请输入 API Key。")
                    else:
                        with st.spinner("AI đang xử lý và đối chiếu dữ liệu, vui lòng đợi trong giây lát..."):
                            try:
                                client = genai.Client()
                                fp_data = df_fp_filtered.to_string()
                                vp_data = df_vp.to_string() if df_vp is not None else "Không có"
                                cn_data = df_cn.to_string() if df_cn is not None else "Không có"
                                sc_data = df_sc.to_string() if df_sc is not None else "Không có"
                                
                                prompt = f"""
                                Bạn là hệ thống đối chiếu nhân sự và chấm công tự động thông minh. 
                                Hãy thực hiện đối chiếu và phân tích dữ liệu CHO ĐÚNG NGÀY: {selected_date_str} (Khung giờ: {shift_group}).
                                
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
                                            model='gemini-2.5-pro',
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
                                st.success("✅ Hoàn tất đối chiếu chấm công! / 考勤核对完成！")
                                
                            except Exception as e:
                                st.error(f"Lỗi xử lý AI / 处理出错: {e}")
                
                if 'analysis_result' in st.session_state:
                    res_date = st.session_state.get('result_date', selected_date_str)
                    st.markdown(f"### 📋 Kết Quả Phân Tích Bất Thường ({res_date})")
                    st.markdown(st.session_state['analysis_result'])
        else:
            st.warning("⚠️ Không nhận diện được dữ liệu ngày trong file vân tay.")
    else:
        st.error("❌ Không tìm thấy cột 'Ngày' trong file vân tay.")
else:
    st.info("👈 Vui lòng tải file bấm vân tay (`.xlsx`, `.xls`, `.csv`) ở thanh bên trái để bắt đầu khởi chạy hệ thống.")
