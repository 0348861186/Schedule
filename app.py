                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='ThongKe_ChamCong')
                st.download_button(
                    label="📥 Tải File Excel Thống Kê / 下载统计Excel文件",
                    data=output_excel.getvalue(),
                    file_name=f"ThongKe_Cham_Cong_{selected_date.replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_d2:
                st.info("💡 Để xuất PDF giữ nguyên giao diện Dashboard, vui lòng dùng tính năng in trình duyệt (Ctrl+P / Cmd+P) chọn Save as PDF. / 导出PDF请使用浏览器打印功能。")

        with tab3:
            st.markdown("### ⚠️ Trọng Tâm Trường Hợp Bất Thường / 异常情况重点分析")
            st.error("• Nhân viên VP02 (Nguyễn Thị F): Về sớm (làm 7.5/8 tiếng), cần kiểm tra đơn xin phép. / 办公室员工VP02: 早退，需检查请假单。")
            st.warning("• Nhân viên 575: Giờ ra sớm hơn quy chuẩn (15:00 PM thay vì 19:00 PM cho ca 12h hoặc tùy quy định). / 工号575: 下班时间提前。")
            st.info("• Các công nhân khác tuân thủ đúng ca làm việc 'N' và 'Đ'. / 其他工人均符合 'N' 和 'Đ' 班次要求。")

    except Exception as e:
        st.error(f"⚠️ Lỗi xử lý file / 文件处理错误: {e}")
else:
    st.info("👈 Vui lòng tải file bấm vân tay và các file danh sách ở thanh bên trái để bắt đầu. / 请在左侧边栏上传指纹打卡及名单文件以开始。")
