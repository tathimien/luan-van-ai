                f"{parsed_rules.get('margin_right')} – "
                f"{parsed_rules.get('margin_top')} – "
                f"{parsed_rules.get('margin_bottom')} cm"
            )
            st.markdown(
                f"**Kiểu trích dẫn:** "
                f"{parsed_rules.get('citation_style', 'keep')}"
            )
            for requirement in parsed_rules.get(
                "detailed_requirements",
                [],
            ):
                st.markdown(f"- {requirement}")

    st.subheader("2. Tải file Word của học viên")
    st.caption(
        "Hệ thống đọc loại hồ sơ trên bìa, đối chiếu cấu trúc với đúng "
        "template và kiểm tra định dạng theo đúng PDF quy định của loại "
        "hồ sơ đã chọn."
    )
    uploaded_docx = st.file_uploader(
        "Thả file .docx vào đây",
        type=["docx"],
        key=f"uploaded_document_{profile_key}",
    )

    st.sidebar.title("📄 Template mẫu đang chọn")
    st.sidebar.info(profile["label"])
    if template_exists:
        with open(template_path, "rb") as template_file:
            bundled_template_bytes = template_file.read()
        st.sidebar.download_button(
            label="📥 TẢI ĐÚNG TEMPLATE WORD MẪU",
            data=bundled_template_bytes,
            file_name=profile["resolved_template_file"],
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )
    else:
        st.sidebar.error("Template của loại hồ sơ này chưa được cài đặt.")

    st.sidebar.markdown("---")
    st.sidebar.title("⚙️ Cặp tệp hệ thống")
    st.sidebar.markdown(
        f"**Template:** `{profile['resolved_template_file']}`"
    )
    st.sidebar.markdown(
        f"**Quy định:** `{profile['resolved_regulation_file']}`"
    )
    st.sidebar.caption(
        "Học viên không thể thay đổi cặp tệp này. Quản trị viên cập nhật "
        "tệp trong thư mục template và quy_dinh."
    )

    active_rules = copy.deepcopy(parsed_rules)
    profile_ready = template_exists and regulation_exists and bool(pdf_text)

    st.markdown("---")
    if st.button(
        "🔍 KIỂM TRA THEO ĐÚNG LOẠI HỒ SƠ",
        type="primary",
        use_container_width=True,
    ):
        if not profile_ready:
            st.error(
                "❌ Chưa thể kiểm tra vì cặp template/quy định của loại "
                "hồ sơ này chưa đầy đủ hoặc PDF chưa đọc được. Vui lòng "
                "liên hệ quản trị viên."
            )
        elif not uploaded_docx:
            st.error("❌ Vui lòng tải file Word (.docx) ở bước 2.")
        else:
            try:
                with st.spinner(
                    "⏳ Đang đối chiếu template, quy định và đánh dấu màu..."
                ):
                    fixed_stream, error_list = process_docx_file(
                        uploaded_docx.getvalue(),
                        active_rules,
                        template_path=template_path,
                        profile_key=profile_key,
                        profile_label=profile["label"],
                        regulation_filename=(
                            profile["resolved_regulation_file"]
                        ),
                    )
            except Exception as exc:
                st.error(f"❌ Không thể xử lý file Word: {exc}")
            else:
                st.markdown(
                    "### 📋 BÁO CÁO KẾT QUẢ KIỂM TRA VÀ SỬA LỖI"
                )
                for error in error_list:
                    st.write(error)
                st.success(
                    "🎉 Đã hoàn thành. Các vị trí cần rà soát được bôi "
                    "vàng hoặc đỏ trong file Word."
                )
                st.download_button(
                    label="📥 TẢI FILE ĐÃ KIỂM TRA VÀ HIGHLIGHT",
                    data=fixed_stream,
                    file_name=checked_output_filename(
                        uploaded_docx.name
                    ),
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
                    use_container_width=True,
                )


if __name__ == "__main__":
    main()
