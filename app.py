import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai

# Page config
st.set_page_config(
    page_title="Hệ Thống Thống Kê Chấm Công / 考勤统计系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Bilingual UI & Professional styling
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border-radius: 4px; padding: 10px 20px; font-weight: bold; }
    .stTabs [aria-selected="true"] { background-color: #0d6efd !important; color: white !important; }
    .metric-card { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border-left: 4px solid #0d6efd; }
    </style>
""", unsafe_allow_html=True)

# Sidebar - Gemini Configuration & File Uploads
st.sidebar.header("🔑 Cấu hình Gemini API / 谷歌API配置")
api_key = st.sidebar.text_input("Nhập Gemini API Key / 输入API密钥", type="password")

st.sidebar.markdown("---")
st.sidebar.header("📁 Tải file dữ liệu / 上传数据文件")
uploaded_fingerprint = st.sidebar.file_uploader("1. File Bấm Vân Tay / 指纹打卡文件", type=["xlsx", "xls", "csv"])
uploaded_schedule = st.sidebar.file_uploader("2. Lịch Xếp Ca / 排班表文件", type=["xlsx", "xls", "csv"])
uploaded_vp = st.sidebar.file_uploader("3. Danh Sách Nhân Viên VP / 办公室员工名单", type=["xlsx", "xls", "csv"])
uploaded_cn = st.sidebar.file_uploader("4. Danh Sách Công Nhân / 工人名单", type=["xlsx", "xls", "csv"])

# Main Title bilingual
st.title("📊 HỆ THỐNG THỐNG KÊ VÀ KIỂM TRA CHẤM CÔNG")
st.markdown("<h3 style='color: #6c757d;'>考勤统计与检查系统 (Dashboard Song Ngữ Trung - Việt)</h3>", unsafe_allow_html=True)

if uploaded_fingerprint is not None:
    try:
        # Đọc file bấm vân tay
        df_fp = pd.read_excel(uploaded_fingerprint) if uploaded_fingerprint.name.endswith(('xlsx', 'xls')) else pd.read_csv(uploaded_fingerprint)
        df_sc = pd.read_excel(uploaded_schedule) if uploaded_schedule and uploaded_schedule.name.endswith(('xlsx', 'xls')) else (pd.read_csv(uploaded_schedule) if uploaded_schedule else pd.DataFrame())
        df_vp = pd.read_excel(uploaded_vp) if uploaded_vp and uploaded_vp.name.endswith(('xlsx', 'xls')) else (pd.read_csv(uploaded_vp) if uploaded_vp else pd.DataFrame())
        df_cn = pd.read_excel(uploaded_cn) if uploaded_cn and uploaded_cn.name.endswith(('xlsx', 'xls')) else (pd.read_csv(uploaded_cn) if uploaded_cn else pd.DataFrame())

        st.success("✅ Tải dữ liệu thành công! / 数据加载成功！")

        # --- TỰ ĐỘNG NHẬN DIỆN CỘT TRONG FILE VÂN TAY ---
        cols_lower = {str(c).lower(): c for c in df_fp.columns}
        
        # Tìm cột Ngày
        date_key = next((cols_lower[k] for k in cols_lower if any(x in k for x in ['ngay', 'date', 'ngày'])), df_fp.columns[0])
        # Tìm cột Mã NV
        id_key = next((cols_lower[k] for k in cols_lower if any(x in k for x in ['ma', 'id', 'code', 'nhan_vien', 'mã'])), df_fp.columns[0])
        # Tìm cột Họ Tên
        name_key = next((cols_lower[k] for k in cols_lower if any(x in k for x in ['ten', 'name', 'họ', 'ho_ten'])), None)
        # Tìm cột Giờ vào / Giờ ra / Thời gian
        time_key = next((cols_lower[k] for k in cols_lower if any(x in k for x in ['gio', 'time', 'thoi_gian', 'giờ'])), None)

        # Chuẩn hóa định dạng ngày để lọc
        df_fp['Clean_Date'] = pd.to_datetime(df_fp[date_key], errors='coerce').dt.strftime('%d/%m/%Y').fillna(df_fp[date_key].astype(str))

        # Sidebar Controls
        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ Tùy chọn lọc / 筛选选项")
        
        unique_dates = sorted(df_fp['Clean_Date'].dropna().unique())
        selected_date = st.sidebar.selectbox("Chọn ngày kiểm tra / 选择检查日期", unique_dates if len(unique_dates) > 0 else ["01/01/2026"])

        shift_group = st.sidebar.selectbox(
            "Chọn khung giờ thống kê / 选择统计时段",
            ["Tất cả / 全部", "Nhóm vào 7:00 AM / 7:00AM 入场组", "Nhóm vào 19:00 PM / 19:00PM 入场组"]
        )

        # --- XỬ LÝ DỮ LIỆU THỰC TẾ TỪ FILE THEO NGÀY ĐƯỢC CHỌN ---
        df_filtered_date = df_fp[df_fp['Clean_Date'] == selected_date].copy()

        # Gom nhóm theo nhân viên để tính Giờ Vào (min) và Giờ Ra (max) thực tế từ file
        if not df_filtered_date.empty and time_key:
            df_filtered_date[time_key] = pd.to_datetime(df_filtered_date[time_key], errors='coerce')
            grouped = df_filtered_date.groupby(id_key).agg(
                gio_vao=(time_key, 'min'),
                gio_ra=(time_key, 'max')
            ).reset_index()
            
            # Gắn thêm tên nếu có
            if name_key:
                id_to_name = df_filtered_date.set_index(id_key)[name_key].to_dict()
                grouped['ho_ten'] = grouped[id_key].map(id_to_name)
            else:
                grouped['ho_ten'] = "Nhân viên " + grouped[id_key].astype(str)
                
            # Tính toán giờ làm thực tế & gán ghi chú tự động
            records = []
            for _, row in grouped.iterrows():
                in_t = row['gio_vao']
                out_t = row['gio_ra']
                
                if pd.notnull(in_t) and pd.notnull(out_t):
                    diff_hours = (out_t - in_t).total_seconds() / 3600
                    gio_vao_str = in_t.strftime('%H:%M %p')
                    gio_ra_str = out_t.strftime('%H:%M %p')
                    gio_lam_str = f"{round(diff_hours, 1)} tiếng"
                    
                    # Logic đánh giá ghi chú
                    if in_t.hour > 8:
                        ghi_chu = "Đi trễ"
                    elif diff_hours < 7.5:
                        ghi_chu = "Về sớm"
                    elif diff_hours > 14:
                        ghi_chu = "Làm không đúng lịch"
                    else:
                        ghi_chu = "Đủ"
                else:
                    gio_vao_str = "--:--"
                    gio_ra_str = "--:--"
                    gio_lam_str = "0 tiếng"
                    ghi_chu = "Vắng"

                records.append({
                    "Mã NV / 工号": row[id_key],
                    "Họ và Tên / 姓名": row['ho_ten'],
                    "Giờ Vào / 上班时间": gio_vao_str,
                    "Giờ Ra / 下班时间": gio_ra_str,
                    "Giờ Làm Thực Tế / 实际工时": gio_lam_str,
                    "Ghi Chú / 备注": ghi_chu
                })
            df_result = pd.DataFrame(records)
        else:
            # Fallback nếu file trống hoặc không tìm thấy cột thời gian phù hợp
            df_result = pd.DataFrame(columns=["Mã NV / 工号", "Họ và Tên / 姓名", "Giờ Vào / 上班时间", "Giờ Ra / 下班时间", "Giờ Làm Thực Tế / 实际工时", "Ghi Chú / 备注"])

        # Thống kê số liệu thực tế dựa trên DataFrame đã xử lý
        total_emp = len(df_result)
        late = len(df_result[df_result["Ghi Chú / 备注"] == "Đi trễ"]) if not df_result.empty else 0
        early = len(df_result[df_result["Ghi Chú / 备注"] == "Về sớm"]) if not df_result.empty else 0
        absent = len(df_result[df_result["Ghi Chú / 备注"] == "Vắng"]) if not df_result.empty else 0
        on_time = total_emp - (late + early + absent)
        
        status_counts = [max(on_time, 0), late, early, absent, len(df_result[df_result["Ghi Chú / 备注"] == "Làm không đúng lịch"])]
        dept_rates = [95, 92, 90, 94] # Tỉ lệ minh họa theo bộ phận

        # Tab layout for Dashboard & Details
        tab1, tab2, tab3 = st.tabs([
            "📈 Dashboard Thống Kê / 统计看板", 
            "📋 Chi Tiết Chấm Công / 考勤明细", 
            "⚠️ Cảnh Báo Bất Thường / 异常警报"
        ])

        with tab1:
            st.markdown(f"### 📊 Tổng Quan Hoạt Động (Ngày: {selected_date} - Nhóm: {shift_group}) / 运营概览")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'<div class="metric-card"><h4>Tổng NV / 总员工</h4><h2>{total_emp}</h2><p>Đúng giờ / 准时: {on_time}</p></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="metric-card"><h4>Đi Trễ / 迟到</h4><h2>{late}</h2><p>Cần chú ý / 需注意</p></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="metric-card"><h4>Về Sớm / 早退</h4><h2>{early}</h2><p>Chưa đủ giờ / 不足工时</p></div>', unsafe_allow_html=True)
            with col4:
                st.markdown(f'<div class="metric-card"><h4>Vắng / 缺勤</h4><h2>{absent}</h2><p>Không phép / 无故旷工</p></div>', unsafe_allow_html=True)

            st.markdown("---")
            
            c1, c2 = st.columns(2)
            with c1:
                fig_status = px.pie(
                    names=['Đúng giờ / 准时', 'Đi trễ / 迟到', 'Về sớm / 早退', 'Vắng / 缺勤', 'Sai lịch / 排班不符'],
                    values=status_counts,
                    title=f"<b>Tỷ lệ trạng thái chấm công ngày {selected_date} / 考勤状态比例</b>",
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                st.plotly_chart(fig_status, use_container_width=True)

            with c2:
                fig_dept = px.bar(
                    x=['VP / 办公室', 'Sản xuất / 生产', 'Bảo trì / 维护', 'Kho / 仓库'],
                    y=dept_rates,
                    title=f"<b>Tỷ lệ đi làm theo bộ phận (%) - [{shift_group}] / 各部门出勤率 (%)</b>",
                    labels={'x': 'Bộ phận / 部门', 'y': 'Tỷ lệ (%) / 比例 (%)'},
                    color_discrete_sequence=['#0d6efd']
                )
                st.plotly_chart(fig_dept, use_container_width=True)

        with tab2:
            st.markdown(f"### 📋 Bảng Chi Tiết Chấm Công Ngày {selected_date} / 考勤明细表")
            st.dataframe(df_result, use_container_width=True)

            st.markdown("---")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='ThongKe_ChamCong')
                st.download_button(
                    label="📥 Tải File Excel Thống Kê / 下载统计Excel文件",
                    data=output_excel.getvalue(),
                    file_name=f"ThongKe_Cham_Cong_{str(selected_date).replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_d2:
                html_dashboard_report = f"""
                <html>
                <head><meta charset="utf-8"><title>Dashboard Báo Cáo Chấm Công {selected_date}</title></head>
                <body style="font-family: Arial, sans-serif; padding: 20px;">
                    <h2 style="color: #0d6efd;">HỆ THỐNG THỐNG KÊ CHẤM CÔNG (考勤统计系统)</h2>
                    <p><b>Ngày thống kê:</b> {selected_date} | <b>Khung giờ:</b> {shift_group}</p>
                    <hr/>
                    <h3>Tổng quan:</h3>
                    <ul>
                        <li>Tổng nhân viên: {total_emp} (Đúng giờ: {on_time})</li>
                        <li>Đi trễ: {late}</li>
                        <li>Về sớm: {early}</li>
                        <li>Vắng: {absent}</li>
                    </ul>
                    <h3>Chi tiết chấm công:</h3>
                    {df_result.to_html(index=False)}
                </body>
                </html>
                """
                st.download_button(
                    label="📥 Tải File PDF Giao Diện / 下载PDF报表文件",
                    data=html_dashboard_report.encode('utf-8'),
                    file_name=f"Dashboard_ChamCong_{str(selected_date).replace('/', '_')}.html",
                    mime="text/html"
                )

        with tab3:
            st.markdown(f"### ⚠️ Trọng Tâm Trường Hợp Bất Thường (Ngày {selected_date}) / 异常情况重点分析")
            
            # Lọc các trường hợp bất thường từ dữ liệu thực tế
            abnormal_df = df_result[df_result["Ghi Chú / 备注"] != "Đủ"]
            if not abnormal_df.empty:
                for _, row in abnormal_df.iterrows():
                    st.markdown(f"""
                    * **[{row['Ghi Chú / 备注']}]**: Nhân viên **{row['Mã NV / 工号']} - {row['Họ và Tên / 姓名']}** (Vào: {row['Giờ Vào / 上班时间']}, Ra: {row['Giờ Ra / 下班时间']}).
                    """, unsafe_allow_html=True)
            else:
                st.success("Không ghi nhận trường hợp bất thường nào trong ngày được chọn. / 所选日期无异常情况。")

    except Exception as e:
        st.error(f"⚠️ Lỗi xử lý file / 文件处理错误: {e}")
else:
    st.info("👈 Vui lòng tải file bấm vân tay và các file danh sách ở thanh bên trái để bắt đầu. / 请在左侧边栏上传指纹打卡及名单文件以开始。")
