import io
import json
import os
import re
import docx
from docx.shared import Cm, Pt
from google import genai
from pypdf import PdfReader
import streamlit as st

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG STREAMLIT & TIÊU ĐỀ MỚI
# ---------------------------------------------------------
st.set_page_config(
    page_title="Đại học Y Hà Nội - Kiểm tra và sửa định dạng luận văn", 
    page_icon="🎓", 
    layout="wide"
)
st.title("🎓 Đại học Y Hà Nội - Kiểm tra và sửa định dạng luận văn")
st.caption(
    "Hệ thống tự động kiểm tra quy định, bảo tồn trang bìa, xử lý mục đạo đức nghiên cứu và chuẩn hóa file Word."
)
API_KEY = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))

# ---------------------------------------------------------
# 2. HÀM ĐỌC VĂN BẢN TỪ FILE PDF QUY ĐỊNH
# ---------------------------------------------------------
def extract_raw_text_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        return ""
    try:
        reader = PdfReader(pdf_path)
        text = "".join([page.extract_text() or "" for page in reader.pages])
        return text.strip()
    except Exception:
        return ""

# ---------------------------------------------------------
# 3. HÀM AI GEMINI PHÂN TÍCH QUY ĐỊNH PDF
# ---------------------------------------------------------
def analyze_rules_with_gemini(pdf_text, api_key):
    default_rules = {
        "font_name": "Times New Roman",
        "font_size": 13.0,
        "line_spacing": 1.5,
        "margin_top": 3.5,
        "margin_bottom": 3.0,
        "margin_left": 3.0,
        "margin_right": 2.0,
        "summary": "Bộ quy định mặc định.",
    }
    if not pdf_text or not api_key:
        return default_rules
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
Phân tích quy định trình bày luận văn trong đoạn văn sau và bóc tách thành JSON.
Nội dung file quy định PDF:
{pdf_text[:6000]}
Trả về DUY NHẤT một chuỗi JSON thuần có cấu trúc chính xác (không kèm mã markdown):
{{
    "font_name": "Times New Roman",
    "font_size": 13.0,
    "line_spacing": 1.5,
    "margin_top": 3.5,
    "margin_bottom": 3.0,
    "margin_left": 3.0,
    "margin_right": 2.0,
    "summary": "Tóm tắt ngắn gọn quy định về Font, Lề, Bảng, Hình và Trích dẫn."
}}
"""
        response = client.models.generate_content(
            model="gemini-2.0-flash", contents=prompt
        )
        raw_text = response.text.strip()
        bt = chr(96) * 3
        clean_json = raw_text.replace(f"{bt}json", "").replace(bt, "").strip()
        return json.loads(clean_json)
    except Exception:
        return default_rules

# ---------------------------------------------------------
# 4. HÀM BẮT LỖI & SỬA MỤC "VẤN ĐỀ ĐẠO ĐỨC TRONG NGHIÊN CỨU"
# ---------------------------------------------------------
def fix_ethics_section(doc):
    detailed_errors = []
    ethics_pattern = re.compile(
        r"(2\.\d+\.?\s*)?(vấn đề đạo đức trong nghiên cứu|đạo đức trong nghiên cứu)", 
        re.IGNORECASE
    )

    for idx, p in enumerate(list(doc.paragraphs)):
        text = p.text
        if not text.strip():
            continue

        match = ethics_pattern.search(text)
        if match:
            start_pos = match.start()
            matched_text = match.group(0)

            # Trường hợp 1: Không ngắt ra đầu dòng (bị dính vào đoạn trước)
            if start_pos > 0 and text[:start_pos].strip() != "":
                detailed_errors.append(
                    f"⚠️ **Đoạn {idx+1}:** Mục `{matched_text}` bị dính chung với câu trước, **chưa được tách ra đầu dòng**! ➡️ **Đã tự động ngắt dòng và in đậm.**"
                )
                text_before = text[:start_pos].rstrip()
                text_after = text[start_pos:].lstrip()

                p.text = text_before

                new_p = doc.add_paragraph()
                p._p.addnext(new_p._p)

                r_title = new_p.add_run(matched_text)
                r_title.bold = True

                rest_text = text_after[len(matched_text):]
                if rest_text:
                    r_rest = new_p.add_run(rest_text)
                    r_rest.bold = False

            # Trường hợp 2: Đã đầu dòng nhưng không được bôi đen
            else:
                is_bold = any(run.bold for run in p.runs if matched_text.lower() in run.text.lower())
                if not is_bold:
                    detailed_errors.append(
                        f"⚠️ **Đoạn {idx+1}:** Tiêu đề `{matched_text}` **chưa được bôi đen**! ➡️ **Đã tự động ép in đậm.**"
                    )
                    p.text = ""
                    r_title = p.add_run(matched_text)
                    r_title.bold = True
                    rest_text = text[len(matched_text):]
                    if rest_text:
                        r_rest = p.add_run(rest_text)
                        r_rest.bold = False

    return detailed_errors

# ---------------------------------------------------------
# 5. HÀM ÉP IN ĐẬM TÊN ĐỀ TÀI
# ---------------------------------------------------------
def enforce_bold_for_thesis_title(p):
    text_lower = p.text.lower().strip()
    is_title = (
        "tên đề tài" in text_lower
        or text_lower.startswith("đề tài:")
        or text_lower.startswith("đề tài :")
        or "tên đề tài:" in text_lower
    )
    if is_title:
        for run in p.runs:
            run.bold = True
        return True
    return False

# ---------------------------------------------------------
# 6. HÀM ĐỔI TRÍCH DẪN [1] THÀNH LŨY THỪA (SUPERSCRIPT)
# ---------------------------------------------------------
def _add_mapped_runs(p, sub_text, sub_map, font_target, size_target):
    if not sub_text or not sub_map:
        return
    curr_str = ""
    curr_bold = sub_map[0][1]
    curr_italic = sub_map[0][2]
    for ch, (_, b, i) in zip(sub_text, sub_map):
        if b == curr_bold and i == curr_italic:
            curr_str += ch
        else:
            r = p.add_run(curr_str)
            r.bold = curr_bold
            r.italic = curr_italic
            r.font.name = font_target
            r.font.size = Pt(size_target)
            curr_str = ch
            curr_bold = b
            curr_italic = i
    if curr_str:
        r = p.add_run(curr_str)
        r.bold = curr_bold
        r.italic = curr_italic
        r.font.name = font_target
        r.font.size = Pt(size_target)

def convert_brackets_to_superscript(p, font_target, size_target):
    text = p.text
    if not text or "[" not in text or "]" not in text:
        return False
    pattern = r"(\[\d+(?:[\s,–-]\d+)*\])"
    if not re.search(pattern, text):
        return False

    char_map = []
    for run in p.runs:
        r_bold = bool(run.bold)
        r_italic = bool(run.italic)
        for ch in run.text:
            char_map.append((ch, r_bold, r_italic))
    if len(char_map) != len(text):
        char_map = [(ch, False, False) for ch in text]

    p.text = ""
    pos = 0
    for match in re.finditer(pattern, text):
        start, end = match.span()
        if start > pos:
            _add_mapped_runs(
                p, text[pos:start], char_map[pos:start], font_target, size_target
            )
        num_str = match.group(1)[1:-1]
        match_bold = char_map[start][1] if start < len(char_map) else False
        match_italic = char_map[start][2] if start < len(char_map) else False
        r = p.add_run(num_str)
        r.font.superscript = True
        r.bold = match_bold
        r.italic = match_italic
        r.font.name = font_target
        r.font.size = Pt(size_target)
        pos = end
    if pos < len(text):
        _add_mapped_runs(p, text[pos:], char_map[pos:], font_target, size_target)
    return True

# ---------------------------------------------------------
# 7. HÀM XÓA KHOẢNG TRẮNG DƯ THỪA (CÓ BẢO VỆ TRANG BÌA & TAB)
# ---------------------------------------------------------
def clean_spaces_and_punctuation(p):
    if not p.runs:
        return False

    text_upper = p.text.upper()
    cover_keywords = ["BỘ GIÁO DỤC", "BỘ Y TẾ", "TRƯỜNG ĐẠI HỌC", "VIỆN NGHIÊN CỨU", "UBND"]
    if any(keyword in text_upper for keyword in cover_keywords):
        return False

    changed = False
    for run in p.runs:
        if not run.text:
            continue
        orig = run.text
        cleaned = orig.replace("\xa0", " ").replace("\u200b", "")

        cleaned = re.sub(r" {2,}", " ", cleaned)
        cleaned = re.sub(r"\s+([:,.,!?,])", r"\1", cleaned)

        if cleaned != orig:
            run.text = cleaned
            changed = True

    runs_with_text = [r for r in p.runs if r.text]
    for idx in range(len(runs_with_text) - 1):
        curr_run = runs_with_text[idx]
        next_run = runs_with_text[idx + 1]
        while (
            curr_run.text
            and next_run.text
            and curr_run.text[-1] == " "
            and next_run.text[0] == " "
        ):
            next_run.text = next_run.text[1:]
            changed = True

    if runs_with_text:
        first_run = runs_with_text[0]
        last_run = runs_with_text[-1]
        l_stripped = first_run.text.lstrip(" \t\xa0")
        if l_stripped != first_run.text:
            first_run.text = l_stripped
            changed = True
        r_stripped = last_run.text.rstrip(" \t\xa0")
        if r_stripped != last_run.text:
            last_run.text = r_stripped
            changed = True

    return changed

# ---------------------------------------------------------
# 8. HÀM TỰ ĐỘNG BỔ SUNG "VÀ CỘNG SỰ" VÀO TRÍCH DẪN
# ---------------------------------------------------------
def fix_author_citations(p):
    text = p.text
    if not text or ("19" not in text and "20" not in text):
        return []
    changed = []
    if re.search(r"\bet\s+al\.?", text, re.IGNORECASE):
        for run in p.runs:
            if run.text:
                new_txt = re.sub(
                    r"\bet\s+al\.?", "và cộng sự", run.text, flags=re.IGNORECASE
                )
                if new_txt != run.text:
                    run.text = new_txt
                    changed.append("Đã chuyển 'et al.' thành 'và cộng sự'")

    pattern_narrative = r"\b([A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+){0,3})\s*\(((?:19|20)\d{2})\)"
    pattern_parenthetical = r"\((?:bởi\s+)?([A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+){0,3})\s*,\s*((?:19|20)\d{2})\)"
    lead_words_pattern = r"^(Theo|Nghiên cứu của|Báo cáo của|Tác giả|Của|Trong)\s+"

    def repl_narrative(m):
        raw_name = m.group(1).strip()
        year = m.group(2)
        lead_match = re.match(lead_words_pattern, raw_name, re.IGNORECASE)
        lead_prefix = lead_match.group(0) if lead_match else ""
        clean_name = raw_name[len(lead_prefix):].strip() if lead_match else raw_name
        if not clean_name:
            clean_name, lead_prefix = raw_name, ""
        if re.search(r"(và\s+cộng\s+sự|và\s+cs\.?|et\s+al\.?)$", clean_name, re.IGNORECASE):
            return m.group(0)
        return f"{lead_prefix}{clean_name} và cộng sự ({year})"

    def repl_parenthetical(m):
        raw_name = m.group(1).strip()
        year = m.group(2)
        if re.search(r"(và\s+cộng\s+sự|và\s+cs\.?|et\s+al\.?)$", raw_name, re.IGNORECASE):
            return m.group(0)
        return f"({raw_name} và cộng sự, {year})"

    run_changed = False
    for run in p.runs:
        if not run.text:
            continue
        r_text = run.text
        new_r_text = re.sub(pattern_narrative, repl_narrative, r_text)
        new_r_text = re.sub(pattern_parenthetical, repl_parenthetical, new_r_text)
        if new_r_text != r_text:
            run.text = new_r_text
            run_changed = True
    if run_changed:
        changed.append("Đã thêm 'và cộng sự' vào trích dẫn Tác Giả (Năm)")
    else:
        full_text = p.text
        new_full_text = re.sub(pattern_narrative, repl_narrative, full_text)
        new_full_text = re.sub(pattern_parenthetical, repl_parenthetical, new_full_text)
        if new_full_text != full_text:
            p.text = new_full_text
            changed.append("Đã thêm 'và cộng sự' vào trích dẫn Tác Giả (Năm)")
    return changed

def has_image(paragraph):
    xml = paragraph._element.xml
    return "w:drawing" in xml or "w:pict" in xml or "a:blip" in xml

# ---------------------------------------------------------
# 9. ĐIỀU HÀNH CHUẨN HÓA TOÀN BỘ FILE DOCX
# ---------------------------------------------------------
def process_docx_file(uploaded_bytes, rules):
    doc = docx.Document(io.BytesIO(uploaded_bytes))
    
    ethics_errors = fix_ethics_section(doc)
    detailed_errors = ethics_errors + []

    font_target = rules["font_name"]
    size_target = float(rules["font_size"])
    line_target = float(rules["line_spacing"])
    target_left = float(rules["margin_left"])
    target_right = float(rules["margin_right"])
    target_top = float(rules["margin_top"])
    target_bottom = float(rules["margin_bottom"])

    for idx, section in enumerate(doc.sections):
        curr_left = round(section.left_margin.cm, 1) if section.left_margin else 0
        curr_right = round(section.right_margin.cm, 1) if section.right_margin else 0
        curr_top = round(section.top_margin.cm, 1) if section.top_margin else 0
        curr_bottom = round(section.bottom_margin.cm, 1) if section.bottom_margin else 0
        if (
            curr_left != target_left
            or curr_right != target_right
            or curr_top != target_top
            or curr_bottom != target_bottom
        ):
            detailed_errors.append(
                f"📌 **Lề Trang (Phần {idx+1}):** Cũ ({curr_left}-{curr_right}-{curr_top}-{curr_bottom}cm) ➡️ **Đã sửa chuẩn** ({target_left}-{target_right}-{target_top}-{target_bottom}cm)"
            )
        section.left_margin = Cm(target_left)
        section.right_margin = Cm(target_right)
        section.top_margin = Cm(target_top)
        section.bottom_margin = Cm(target_bottom)

    spacing_count = 0
    title_bold_count = 0
    fixed_authors = []
    converted_citation_count = 0
    paragraphs = doc.paragraphs

    for i, p in enumerate(paragraphs):
        if clean_spaces_and_punctuation(p):
            spacing_count += 1
        if enforce_bold_for_thesis_title(p):
            title_bold_count += 1
        authors = fix_author_citations(p)
        if authors:
            fixed_authors.extend(authors)
        if convert_brackets_to_superscript(p, font_target, size_target):
            converted_citation_count += 1
        if has_image(p):
            nearby = ""
            if i + 1 < len(paragraphs):
                nearby += paragraphs[i + 1].text.lower() + " "
            if i > 0:
                nearby += paragraphs[i - 1].text.lower() + " "
            if "hình" not in nearby and "sơ đồ" not in nearby:
                detailed_errors.append(
                    f"🖼️ **Đoạn {i+1} (Hình ảnh)**: Thiếu **Chú thích / Tên hình** bên dưới!"
                )
            if "nguồn" not in nearby and "source" not in nearby:
                detailed_errors.append(
                    f"🖼️ **Đoạn {i+1} (Hình ảnh)**: Thiếu **Trích dẫn nguồn**!"
                )

    if title_bold_count > 0:
        detailed_errors.append("🖋️ **Tên Đề Tài:** Đã kiểm tra và **ép buộc In Đậm** cho Tên Đề Tài!")
    if spacing_count > 0:
        detailed_errors.append(f"🧹 **Làm Sạch Dấu Cách:** Đã dọn sạch khoảng trắng thừa tại **{spacing_count} đoạn** văn.")
    if fixed_authors:
        detailed_errors.append(f"✍️ **Trích Dẫn Tác Giả:** Đã bổ sung 'và cộng sự' chuẩn xác tại **{len(fixed_authors)} vị trí**.")
    if converted_citation_count > 0:
        detailed_errors.append(f"✨ **Trích Dẫn Số:** Đã chuyển **{converted_citation_count} đoạn** ngoặc vuông `[1]` thành **lũy thừa (¹, ²)**!")

    for p in paragraphs:
        if not p.text.strip():
            continue
        p.paragraph_format.line_spacing = line_target
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        for run in p.runs:
            run.font.name = font_target
            if not run.font.superscript:
                run.font.size = Pt(size_target)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    clean_spaces_and_punctuation(p)
                    convert_brackets_to_superscript(p, font_target, size_target)
                    p.paragraph_format.line_spacing = line_target
                    for run in p.runs:
                        run.font.name = font_target
                        if not run.font.superscript:
                            run.font.size = Pt(size_target)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output, detailed_errors

# ---------------------------------------------------------
# 10. GIAO DIỆN CHÍNH STREAMLIT
# ---------------------------------------------------------
col_main, col_side = st.columns([2, 1])

with col_main:
    st.subheader("1. Chọn File Quy Định Trình Bày (.PDF)")
    pdf_folder = "quy_dinh"
    pdf_files = (
        [f for f in os.listdir(pdf_folder) if f.endswith(".pdf")]
        if os.path.exists(pdf_folder)
        else []
    )
    selected_pdf = (
        st.selectbox("Chọn file quy định PDF trong thư mục `quy_dinh`:", pdf_files)
        if pdf_files
        else None
    )

    pdf_text = ""
    parsed_rules = {
        "font_name": "Times New Roman",
        "font_size": 13.0,
        "line_spacing": 1.5,
        "margin_top": 3.5,
        "margin_bottom": 3.0,
        "margin_left": 3.0,
        "margin_right": 2.0,
        "summary": "Bộ quy định mặc định.",
    }

    if selected_pdf:
        pdf_path = os.path.join(pdf_folder, selected_pdf)
        pdf_text = extract_raw_text_from_pdf(pdf_path)
        if pdf_text:
            parsed_rules = analyze_rules_with_gemini(pdf_text, API_KEY)
            st.success(f"✅ Đã đọc thành công file `{selected_pdf}`!")
            with st.expander("🔍 Xem nội dung quy định đọc từ PDF"):
                st.text_area("Chữ rút ra từ PDF:", pdf_text, height=200)
        else:
            st.warning("⚠️ File PDF dạng Scan/Ảnh chụp. Bạn có thể tùy chỉnh thông số ở Sidebar bên phải!")

    st.subheader("2. Tải File Luận Văn Của Học Viên (.DOCX)")
    uploaded_docx = st.file_uploader("Thả file .docx luận văn vào đây", type=["docx"])

# ---------------------------------------------------------
# 11. THANH BÊN (SIDEBAR): TẢI TEMPLATE & CHỈNH THÔNG SỐ
# ---------------------------------------------------------
st.sidebar.title("📄 Tải Template Mẫu")
template_path = os.path.join("template", "template.docx")
if os.path.exists(template_path):
    with open(template_path, "rb") as f:
        template_bytes = f.read()
    st.sidebar.download_button(
        label="📥 TẢI TEMPLATE WORD MẪU (.DOCX)",
        data=template_bytes,
        file_name="template.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    st.sidebar.caption("👉 Tải file mẫu về để điền nội dung chuẩn định dạng.")
else:
    st.sidebar.info("ℹ️ Không tìm thấy file `template.docx` trong thư mục `template/`.")

st.sidebar.markdown("---")
st.sidebar.title("⚙️ Bảng Quy Định Định Dạng")

final_font = st.sidebar.text_input("Font Chữ Thân Bài", value=parsed_rules.get("font_name", "Times New Roman"))
final_size = st.sidebar.number_input("Cỡ Chữ (pt)", value=float(parsed_rules.get("font_size", 13.0)), step=0.5)
final_spacing = st.sidebar.number_input("Giãn Dòng (Line Spacing)", value=float(parsed_rules.get("line_spacing", 1.5)), step=0.1)

st.sidebar.markdown("---")
st.sidebar.subheader("📐 Căn Lề Trang (cm)")
final_left = st.sidebar.number_input("Lề Trái (Left)", value=float(parsed_rules.get("margin_left", 3.0)), step=0.5)
final_right = st.sidebar.number_input("Lề Phải (Right)", value=float(parsed_rules.get("margin_right", 2.0)), step=0.5)
final_top = st.sidebar.number_input("Lề Trên (Top)", value=float(parsed_rules.get("margin_top", 3.5)), step=0.5)
final_bottom = st.sidebar.number_input("Lề Dưới (Bottom)", value=float(parsed_rules.get("margin_bottom", 3.0)), step=0.5)

active_rules = {
    "font_name": final_font,
    "font_size": final_size,
    "line_spacing": final_spacing,
    "margin_left": final_left,
    "margin_right": final_right,
    "margin_top": final_top,
    "margin_bottom": final_bottom,
}

# ---------------------------------------------------------
# 12. CHẠY CHƯƠNG TRÌNH VỚI NÚT BẤM MỚI
# ---------------------------------------------------------
with col_main:
    st.markdown("---")
    # Đổi tên nút thành "KIỂM TRA ĐỊNH DẠNG"
    if st.button("🔍 KIỂM TRA ĐỊNH DẠNG", type="primary", use_container_width=True):
        if not uploaded_docx:
            st.error("❌ Vui lòng tải file Luận Văn (.docx) ở Bước 2 trước!")
        else:
            with st.spinner("⏳ Đang tiến hành kiểm tra và chuẩn hóa định dạng luận văn..."):
                fixed_stream, error_list = process_docx_file(
                    uploaded_docx.getvalue(), active_rules
                )
            st.markdown("### 📋 BẢNG BÁO CÁO KẾT QUẢ KIỂM TRA & SỬA LỖI")
            if error_list:
                for err in error_list:
                    st.write(err)
                st.success("🎉 **Đã hoàn thành kiểm tra và tự động sửa các lỗi định dạng!**")
            else:
                st.success("🎉 File luận văn đã hoàn toàn đạt chuẩn định dạng!")
            
            st.download_button(
                label="📥 TẢI FILE LUẬN VĂN ĐÃ ĐƯỢC CHUẨN HÓA",
                data=fixed_stream,
                file_name="LuanVan_DaiHocYHaNoi_DaChuanHoa.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
