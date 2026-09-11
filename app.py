import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, time, timedelta
import io

# Page Config
st.set_page_config(
    page_title="Hệ thống Quản lý Chấm công | 考勤管理系统",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for bilingual styling & clean dashboard layout
st.markdown("""
<style>
    .main { background-color: #f8fafc; }
    .stButton>button {
        width: 100%;
        background-color: #0284c7;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        border: none;
        padding: 0.5rem 1rem;
    }
    .stButton>button:hover {
        background-color: #0369a1;
        color: white;
    }
    .metric-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        text-align: center;
    }
    h1, h2, h3 {
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

# Header Bilingual
st.markdown("<h1 style='text-align: center; color: #0284c7;'>HỆ THỐNG QUẢN LÝ CHẤM CÔNG & ĐIỂM DANH</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #64748b;'>考勤管理与出勤统计系统</h3>", unsafe_allow_html=True)
st.markdown("---")

# Sidebar for File Uploads
st.sidebar.markdown("<h2>📁 Tải File Dữ Liệu / 数据文件上传</h2>", unsafe_allow_html=True)

uploaded_fp = st.sidebar.file_uploader("1. File bấm vân tay (指纹打卡文件 - Excel/CSV)", type=["xlsx", "xls", "csv"])
uploaded_shift = st.sidebar.file_uploader("2. Lịch xếp ca (排班表 - Excel/CSV)", type=["xlsx", "xls", "csv"])
uploaded_vp = st.sidebar.file_uploader("3. Danh sách Nhân viên Văn phòng (办公室员工名单)", type=["xlsx", "xls", "csv"])
uploaded_cn = st.sidebar.file_uploader("4. Danh sách Công nhân (工人名单)", type=["xlsx", "xls", "csv"])

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Quy tắc thiết lập / 规则设置")
holiday_input = st.sidebar.text_input("Ngày lễ (định dạng YYYY-MM-DD, cách nhau bằng dấu phẩy):", "2026-09-02")
holidays = [h.strip() for h in holiday_input.split(",") if h.strip()]

# Helper functions to load data safely
@st.cache_data
def load_data(uploaded_file):
    if uploaded_file is None:
        return None
    try:
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file)
        else:
            return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Lỗi đọc file {uploaded_file.name}: {e}")
        return None

df_fp = load_data(uploaded_fp)
df_shift = load_data(uploaded_shift)
df_vp = load_data(uploaded_vp)
df_cn = load_data(uploaded_cn)

# Demo data generator if files are not uploaded yet so user can immediately test
if df_fp is None:
    st.info("💡 Chưa tải file bấm vân tay. Hệ thống đang hiển thị dữ liệu mẫu để thử nghiệm. Vui lòng tải file thực tế ở thanh bên trái.")
    # Create sample fingerprint data
    np.random.seed(42)
    sample_dates = pd.date_range(start="2026-09-01", end="2026-09-07").strftime("%Y-%m-%d").tolist()
    sample_ids = ["F01", "F02", "VP01", "VP02", "575", "749", "949", "673", "A068"]
    sample_names = ["Nguyễn Văn A", "Trần Thị B", "Lê Văn C", "Phạm Thị D", "Nhân viên 575", "Nhân viên 749", "Nhân viên 949", "Nhân viên 673", "Nhân viên A068"]
    sample_depts = ["Sản xuất", "Sản xuất", "Văn phòng", "Văn phòng", "Đặc biệt", "Đặc biệt", "Đặc biệt", "Đặc biệt", "Đặc biệt"]
    
    fp_data = []
    for d in sample_dates:
        for idx, sid in enumerate(sample_ids):
            # simulate some missing or late
            in_time = "07:00:00" if sid != "VP01" else "08:15:00"
            out_time = "15:00:00" if sid == "575" else ("19:00:00" if sid in ["749","949"] else "17:00:00")
            if sid == "F01": # Night shift worker
                in_time = "19:05:00"
                out_time = "07:00:00"
            fp_data.append({
                "Mã Nhân Viên": sid,
                "Tên nhân viên": sample_names[idx],
                "Phòng ban": sample_depts[idx],
                "Ngày": d,
                "Giờ vào": in_time,
                "Giờ ra": out_time
            })
    df_fp = pd.DataFrame(fp_data)

    # Sample Shift schedule
    shift_data = []
    for d in sample_dates:
        shift_data.append({"Mã số NV": "F01", "Họ Và tên": "Nguyễn Văn A", "Ngày Vào Xưởng": d, "Mỗi ca": "Đ", "Bộ phận": "Sản xuất"})
        shift_data.append({"Mã số NV": "F02", "Họ Và tên": "Trần Thị B", "Ngày Vào Xưởng": d, "Mỗi ca": "N", "Bộ phận": "Sản xuất"})
    df_shift = pd.DataFrame(shift_data)

    # Sample VP list
    df_vp = pd.DataFrame({"Mã số NV": ["VP01", "VP02"], "Họ Và tên": ["Lê Văn C", "Phạm Thị D"], "Bộ phận": ["Kế toán", "Hành chính"]})

    # Sample CN list
    df_cn = pd.DataFrame({"Mã số NV": ["F01", "F02"], "Họ Và tên": ["Nguyễn Văn A", "Trần Thị B"], "Bộ phận": ["Sản xuất", "Sản xuất"]})

# Standardize column names if needed
# Let's inspect fingerprint columns
fp_cols = df_fp.columns.tolist()
st.sidebar.markdown(f"**Cột file vân tay / 指纹列:** {', ' + str(fp_cols) if fp_cols else 'Chưa có'}")

# Date selection on dashboard
if "Ngày" in df_fp.columns:
    df_fp["Ngày"] = pd.to_datetime(df_fp["Ngày"]).dt.strftime("%Y-%m-%d")
    unique_dates = sorted(df_fp["Ngày"].unique().tolist())
else:
    # try finding date column
    date_col = [c for c in df_fp.columns if 'ngày' in c.lower() or 'date' in c.lower()]
    if date_col:
        df_fp["Ngày"] = pd.to_datetime(df_fp[date_col[0]]).dt.strftime("%Y-%m-%d")
        unique_dates = sorted(df_fp["Ngày"].unique().tolist())
    else:
        unique_dates = ["2026-09-06"]
        df_fp["Ngày"] = "2026-09-06"

selected_date = st.selectbox("📅 Chọn ngày kiểm tra / 选择检查日期:", unique_dates)

# Main Processing Logic
def process_attendance(df_fp, df_shift, df_vp, df_cn, target_date, holidays):
    results = []
    
    # Special employees mapping
    special_rules = {
        "575": {"in": time(7, 0), "out": time(15, 0), "hours": 8},
        "749": {"in": time(7, 0), "out": time(19, 0), "hours": 12},
        "949": {"in": time(7, 0), "out": time(19, 0), "hours": 12},
        "673": {"in": time(7, 0), "out": time(15, 0), "hours": 8},
        "A068": {"in": time(10, 0), "out": time(18, 0), "hours": 8}
    }
    
    # Filter fingerprint records for target date
    df_day = df_fp[df_fp["Ngày"] == target_date]
    
    # Get all unique employee IDs from all lists
    all_emp_ids = set()
    
    # Helper to find column containing ID
    def get_id_col(df):
        if df is None: return None
        for c in df.columns:
            if 'mã' in c.lower() or 'id' in c.lower() or 'no' in c.lower() or 'đường' in c.lower():
                return c
        return df.columns[1] if len(df.columns) > 1 else df.columns[0]

    def get_name_col(df):
        if df is None: return None
        for c in df.columns:
            if 'tên' in c.lower() or 'name' in c.lower() or 'họ' in c.lower():
                return c
        return df.columns[2] if len(df.columns) > 2 else df.columns[0]

    cn_id_col = get_id_col(df_cn)
    cn_name_col = get_name_col(df_cn)
    vp_id_col = get_id_col(df_vp)
    vp_name_col = get_name_col(df_vp)
    fp_id_col = get_id_col(df_fp)
    fp_name_col = get_name_col(df_fp)

    cn_ids = set(df_cn[cn_id_col].astype(str).str.strip()) if df_cn is not None else set()
    vp_ids = set(df_vp[vp_id_col].astype(str).str.strip()) if df_vp is not None else set()
    
    # Shift schedule mapping for workers on target_date
    shift_map = {}
    if df_shift is not None:
        s_id_col = get_id_col(df_shift)
        s_date_col = None
        s_shift_col = None
        for c in df_shift.columns:
            if 'ngày' in c.lower() or 'date' in c.lower() or 'vào xưởng' in c.lower():
                s_date_col = c
            if 'ca' in c.lower() or 'mỗi' in c.lower() or 'shift' in c.lower():
                s_shift_col = c
        
        if s_date_col and s_shift_col:
            for _, row in df_shift.iterrows():
                r_date = str(row[s_date_col]).strip()
                # normalize date format if needed
                if target_date in r_date or r_date in target_date:
                    s_id = str(row[s_id_col]).strip()
                    shift_map[s_id] = str(row[s_shift_col]).strip()

    # Combine all employees present in fingerprint or master lists for that day
    fp_records = {}
    for _, row in df_day.iterrows():
        eid = str(row[fp_id_col]).strip()
        fp_records[eid] = row

    # Gather master list of active employees to check
    master_employees = {}
    if df_cn is not None:
        for _, row in df_cn.iterrows():
            eid = str(row[cn_id_col]).strip()
            name = str(row[cn_name_col]).strip() if cn_name_col in df_cn else "N/A"
            master_employees[eid] = {"name": name, "type": "CN"}
            
    if df_vp is not None:
        for _, row in df_vp.iterrows():
            eid = str(row[vp_id_col]).strip()
            name = str(row[vp_name_col]).strip() if vp_name_col in df_vp else "N/A"
            master_employees[eid] = {"name": name, "type": "VP"}

    for eid in special_rules.keys():
        if eid not in master_employees:
            master_employees[eid] = {"name": f"Nhân viên đặc biệt {eid}", "type": "SPECIAL"}

    # Also include any ID present in fingerprint that might not be explicitly listed
    for eid in fp_records.keys():
        if eid not in master_employees:
            name_val = str(fp_records[eid].get(fp_name_col, "N/A"))
            master_employees[eid] = {"name": name_val, "type": "OTHER"}

    # Evaluate each employee for target date
    dt_obj = datetime.strptime(target_date, "%Y-%m-%d")
    is_sunday = dt_obj.weekday() == 6
    is_holiday = target_date in holidays

    for eid, info in master_employees.items():
        name = info["name"]
        emp_type = info["type"]
        
        # Check fingerprint punch
        punch = fp_records.get(eid, None)
        
        in_time_str = str(punch.get("Giờ vào", "")) if punch is not None and "Giờ vào" in punch else ""
        out_time_str = str(punch.get("Giờ ra", "")) if punch is not None and "Giờ ra" in punch else ""
        
        # Clean time strings
        if pd.isna(in_time_str) or in_time_str.lower() in ['nan', 'nat', '']:
            in_time_str = ""
        if pd.isna(out_time_str) or out_time_str.lower() in ['nan', 'nat', '']:
            out_time_str = ""

        note = []
        actual_hours = 0.0
        
        # Holiday check
        if is_holiday:
            results.append({
                "Mã NV": eid,
                "Họ và tên": name,
                "Giờ vào": in_time_str,
                "Giờ ra": out_time_str,
                "Giờ làm thực tế": 0.0,
                "Ghi chú": "Lễ"
            })
            continue

        # Rule processing by employee group
        if eid in special_rules:
            rule = special_rules[eid]
            std_in = rule["in"]
            std_out = rule["out"]
            
            if not in_time_str and not out_time_str:
                note.append("vắng")
            else:
                if not in_time_str:
                    note.append("BV")
                elif not out_time_str:
                    note.append("BR")
                else:
                    try:
                        t_in = pd.to_datetime(in_time_str).time()
                        t_out = pd.to_datetime(out_time_str).time()
                        if t_in > std_in:
                            note.append("đi trễ")
                        if t_out < std_out:
                            note.append("về sớm")
                        
                        # Calculate hours
                        dt_in = datetime.combine(dt_obj.date(), t_in)
                        dt_out = datetime.combine(dt_obj.date(), t_out)
                        if dt_out < dt_in: # overnight case
                            dt_out += timedelta(days=1)
                        actual_hours = round((dt_out - dt_in).total_seconds() / 3600.0, 2)
                    except:
                        actual_hours = rule["hours"]

        elif emp_type == "VP" or eid in vp_ids:
            if is_sunday:
                note.append("Ngày nghỉ")
                actual_hours = 0.0
            else:
                std_in = time(8, 0)
                std_out = time(17, 0)
                if not in_time_str and not out_time_str:
                    note.append("vắng")
                else:
                    if not in_time_str:
                        note.append("BV")
                    elif not out_time_str:
                        note.append("BR")
                    else:
                        try:
                            t_in = pd.to_datetime(in_time_str).time()
                            t_out = pd.to_datetime(out_time_str).time()
                            if t_in > std_in:
                                note.append("đi trễ")
                            if t_out < std_out:
                                note.append("về sớm")
                            
                            dt_in = datetime.combine(dt_obj.date(), t_in)
                            dt_out = datetime.combine(dt_obj.date(), t_out)
                            actual_hours = round((dt_out - dt_in).total_seconds() / 3600.0, 2)
                            if actual_hours < 8:
                                note.append(f"về sớm ({actual_hours}h)")
                        except:
                            actual_hours = 8.0

        elif emp_type == "CN" or eid in cn_ids:
            shift = shift_map.get(eid, "N") # Default N if not specified
            if shift == "Nghỉ" or is_sunday:
                note.append("Ngày nghỉ")
                actual_hours = 0.0
            elif shift == "Đ": # Ca đêm (19:00 to 07:00 next day)
                std_in = time(19, 0)
                std_out = time(7, 0)
                if not in_time_str and not out_time_str:
                    note.append("vắng")
                else:
                    if not in_time_str:
                        note.append("BV")
                    elif not out_time_str:
                        note.append("BR")
                    else:
                        try:
                            t_in = pd.to_datetime(in_time_str).time()
                            t_out = pd.to_datetime(out_time_str).time()
                            if t_in > std_in:
                                note.append("đi trễ")
                            # calculate night shift hours
                            actual_hours = 12.0
                        except:
                            actual_hours = 12.0
            else: # Ca ngày "N" (07:00 to 19:00)
                std_in = time(7, 0)
                std_out = time(19, 0)
                if not in_time_str and not out_time_str:
                    note.append("vắng")
                else:
                    if not in_time_str:
                        note.append("BV")
                    elif not out_time_str:
                        note.append("BR")
                    else:
                        try:
                            t_in = pd.to_datetime(in_time_str).time()
                            t_out = pd.to_datetime(out_time_str).time()
                            if t_in > std_in:
                                note.append("đi trễ")
                            if t_out < std_out:
                                note.append("về sớm")
                            actual_hours = 12.0
                        except:
                            actual_hours = 12.0
        else:
            # General fallback
            if not in_time_str and not out_time_str:
                note.append("vắng")
            else:
                note.append("làm không đúng lịch")

        note_str = ", ".join(note) if note else "Đúng giờ"
        
        results.append({
            "Mã NV": eid,
            "Họ và tên": name,
            "Giờ vào": in_time_str,
            "Giờ ra": out_time_str,
            "Giờ làm thực tế": actual_hours,
            "Ghi chú": note_str
        })

    return pd.DataFrame(results)

df_result = process_attendance(df_fp, df_shift, df_vp, df_cn, selected_date, holidays)

st.markdown(f"### 📋 Kết quả Chấm công Ngày: {selected_date} / 考勤结果")

# Metrics Summary Cards
col1, col2, col3, col4 = st.columns(4)
total_emp = len(df_result)
late_count = df_result["Ghi chú"].str.contains("đi trễ", case=False).sum()
early_count = df_result["Ghi chú"].str.contains("về sớm", case=False).sum()
absent_count = df_result["Ghi chú"].str.contains("vắng", case=False).sum()

with col1:
    st.markdown(f"<div class='metric-card'><h4>Tổng nhân sự / 总人数</h4><h2>{total_emp}</h2></div>", unsafe_allow_html=True)
with col2:
    st.markdown(f"<div class='metric-card'><h4>Đi trễ / 迟到</h4><h2 style='color: #d97706;'>{late_count}</h2></div>", unsafe_allow_html=True)
with col3:
    st.markdown(f"<div class='metric-card'><h4>Về sớm / 早退</h4><h2 style='color: #ea580c;'>{early_count}</h2></div>", unsafe_allow_html=True)
with col4:
    st.markdown(f"<div class='metric-card'><h4>Vắng mặt / 缺勤</h4><h2 style='color: #dc2626;'>{absent_count}</h2></div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Display Table
st.dataframe(df_result, use_container_width=True)

# Analytics Charts
st.markdown("### 📊 Thống kê Phân tích / 统计分析图表")
c1, c2 = st.columns(2)

with c1:
    # Breakdown of notes
    note_counts = df_result["Ghi chú"].value_counts().reset_index()
    note_counts.columns = ["Trạng thái / 状态", "Số lượng / 数量"]
    fig_pie = px.pie(note_counts, names="Trạng thái / 状态", values="Số lượng / 数量", 
                     title="Tỷ lệ Trạng thái Chấm công / 考勤状态比例",
                     color_discrete_sequence=px.colors.qualitative.Set3)
    st.plotly_chart(fig_pie, use_container_width=True)

with c2:
    fig_bar = px.bar(df_result, x="Mã NV", y="Giờ làm thực tế", color="Ghi chú",
                     title="Giờ làm thực tế theo Nhân viên / 实际工作时间统计",
                     color_discrete_sequence=px.colors.qualitative.Bold)
    st.plotly_chart(fig_bar, use_container_width=True)

# Download buttons
st.markdown("### 📥 Tải xuống Báo cáo / 下载报告")
col_down1, col_down2 = st.columns(2)

# Excel export
output_excel = io.BytesIO()
with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    df_result.to_excel(writer, index=False, sheet_name='ThongKeChamCong')
excel_data = output_excel.getvalue()

with col_down1:
    st.download_button(
        label="📥 Tải File Excel Thống Kê / 下载Excel统计表",
        data=excel_data,
        file_name=f"ThongKeChamCong_{selected_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_down2:
    # PDF Export using Weasyprint
    if st.button("🖨️ Xuất Báo Cáo PDF / 导出PDF报告"):
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <style>
            @page {{ size: A4; margin: 15mm; background-color: #ffffff; }}
            body {{ font-family: 'Helvetica', 'Arial', sans-serif; color: #1e293b; font-size: 11pt; line-height: 1.4; }}
            h1 {{ color: #0284c7; text-align: center; font-size: 20pt; margin-bottom: 5px; }}
            h3 {{ color: #64748b; text-align: center; font-size: 12pt; margin-top: 0; }}
            .summary-box {{ display: flex; justify-content: space-around; margin: 20px 0; background: #f1f5f9; padding: 15px; border-radius: 8px; }}
            .stat-item {{ text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #cbd5e1; padding: 8px 10px; text-align: left; font-size: 10pt; }}
            th {{ background-color: #0284c7; color: white; }}
            tr:nth-child(even) {{ background-color: #f8fafc; }}
        </style>
        </head>
        <body>
            <h1>BÁO CÁO THỐNG KÊ CHẤM CÔNG</h1>
            <h3>考勤统计报告 - Ngày / 日期: {selected_date}</h3>
            <hr>
            <div class="summary-box">
                <div class="stat-item"><b>Tổng nhân sự:</b> {total_emp}</div>
                <div class="stat-item"><b>Đi trễ:</b> {late_count}</div>
                <div class="stat-item"><b>Về sớm:</b> {early_count}</div>
                <div class="stat-item"><b>Vắng mặt:</b> {absent_count}</div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>Mã NV</th>
                        <th>Họ và Tên</th>
                        <th>Giờ vào</th>
                        <th>Giờ ra</th>
                        <th>Giờ làm thực tế</th>
                        <th>Ghi chú</th>
                    </tr>
                </thead>
                <tbody>
        """
        for _, row in df_result.iterrows():
            html_content += f"<tr><td>{row['Mã NV']}</td><td>{row['Họ và tên']}</td><td>{row['Giờ vào']}</td><td>{row['Giờ ra']}</td><td>{row['Giờ làm thực tế']}</td><td>{row['Ghi chú']}</td></tr>"
        
        html_content += """
                </tbody>
            </table>
        </body>
        </html>
        """
        
        try:
            from weasyprint import HTML
            pdf_path = f"ThongKeChamCong_{selected_date}.pdf"
            HTML(string=html_content).write_pdf(pdf_path)
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            st.download_button(
                label="📥 Bấm để tải PDF / 点击下载PDF",
                data=pdf_bytes,
                file_name=pdf_path,
                mime="application/pdf"
            )
            st.success("Tạo file PDF thành công! / PDF文件生成成功！")
        except Exception as e:
            st.error(f"Lỗi xuất PDF: {e}")
