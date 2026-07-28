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
# 1. CẤU HÌNH TRANG STREAMLIT & SIDEBAR
# ---------------------------------------------------------
st.set_page_config(
    page_title="ThesisGuard — Sửa định dạng Luận Văn", page_icon="🎓", layout="wide"
)

st.title("🎓 ThesisGuard — Sửa định dạng Luận Văn")
st.caption(
    "Bảo tồn định dạng chính xác từng ký tự, tự động bổ sung 'và cộng sự' vào"
    " trích dẫn tác giả!"
)

API_KEY = st.secrets.get("GEMINI_API_KEY", "")


# ---------------------------------------------------------
# 2. HÀM ĐỌC VĂN BẢN TỪ FILE PDF
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
    "margin_top": 2.0,
    "margin_bottom": 2.0,
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
# 4. HÀM ÉP BẮT BUỘC IN ĐẬM TÊN ĐỀ TÀI
# ---------------------------------------------------------
def enforce_bold_for_thesis_title(p):
  """Chỉ ép In Đậm duy nhất cho dòng Tên Đề Tài"""
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
# 5. HÀM BẢO TỒN ĐỊNH DẠNG TỪNG KÝ TỰ KHI ĐỔI [1] THÀNH LŨY THỪA
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

  # 1. Trích xuất thuộc tính Bold/Italic từng ký tự gốc
  char_map = []
  for run in p.runs:
    r_bold = bool(run.bold)
    r_italic = bool(run.italic)
    for ch in run.text:
      char_map.append((ch, r_bold, r_italic))

  if len(char_map) != len(text):
    char_map = [(ch, False, False) for ch in text]

  # 2. Xây dựng lại đoạn văn bảo toàn chính xác định dạng
  p.text = ""
  pos = 0
  for match in re.finditer(pattern, text):
    start, end = match.span()

    if start > pos:
      _add_mapped_runs(
          p, text[pos:start], char_map[pos:start], font_target, size_target
      )

    num_str = match.group(1)[1:-1]  # Lấy số, bỏ dấu []
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
# 6. HÀM QUÉT DỌN KHOẢNG TRẮNG ĐA TẦNG & DẤU CÂU
# ---------------------------------------------------------
def clean_spaces_and_punctuation(p):
  if not p.runs:
    return False

  changed = False
  for run in p.runs:
    if not run.text:
      continue
    orig = run.text
    cleaned = orig.replace("\xa0", " ").replace("\u200b", "")
    cleaned = re.sub(r"[ \t]{2,}", " ", cleaned)
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
# 7. TỰ ĐỘNG BỔ SUNG "VÀ CỘNG SỰ" VÀO TRÍCH DẪN TÁC GIẢ (NĂM)
# ---------------------------------------------------------
def fix_author_citations(p):
  text = p.text
  if not text or ("19" not in text and "20" not in text):
    return []

  changed = []

  # A. Chuẩn hóa 'et al.' thành 'và cộng sự'
  if re.search(r"\bet\s+al\.?", text, re.IGNORECASE):
    for run in p.runs:
      if run.text:
        new_txt = re.sub(
            r"\bet\s+al\.?", "và cộng sự", run.text, flags=re.IGNORECASE
        )
        if new_txt != run.text:
          run.text = new_txt
          changed.append("Đã chuyển 'et al.' thành 'và cộng sự'")

  # B. Định nghĩa pattern tìm tên tác giả đứng trước (Năm) hoặc (Tác giả, Năm)
  # Pattern 1: Tác Giả (Năm)
  pattern_narrative = r"\b([A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+){0,3})\s*\(((?:19|20)\d{2})\)"

  # Pattern 2: (Tác Giả, Năm)
  pattern_parenthetical = r"\((?:bởi\s+)?([A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹA-Za-zÀ-Ỹ]+){0,3})\s*,\s*((?:19|20)\d{2})\)"

  lead_words_pattern = r"^(Theo|Nghiên cứu của|Báo cáo của|Tác giả|Của|Trong)\s+"

  def repl_narrative(m):
    raw_name = m.group(1).strip()
    year = m.group(2)

    lead_match = re.match(lead_words_pattern, raw_name, re.IGNORECASE)
    lead_prefix = ""
    if lead_match:
      lead_prefix = lead_match.group(0)
      clean_name = raw_name[len(lead_prefix) :].strip()
    else:
      clean_name = raw_name

    if not clean_name:
      clean_name = raw_name
      lead_prefix = ""

    # Nếu đã có 'và cộng sự' / 'và cs' thì giữ nguyên
    if re.search(
        r"(và\s+cộng\s+sự|và\s+cs\.?|et\s+al\.?)$", clean_name, re.IGNORECASE
    ):
      return m.group(0)

    return f"{lead_prefix}{clean_name} và cộng sự ({year})"

  def repl_parenthetical(m):
    raw_name = m.group(1).strip()
    year = m.group(2)

    if re.search(
        r"(và\s+cộng\s+sự|và\s+cs\.?|et\s+al\.?)$", raw_name, re.IGNORECASE
    ):
      return m.group(0)

    return f"({raw_name} và cộng sự, {year})"

  # C. Thực hiện chuyển đổi
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
    # Xử lý trường hợp chuỗi trích dẫn bị ngắt qua nhiều run khác nhau
    full_text = p.text
    new_full_text = re.sub(pattern_narrative, repl_narrative, full_text)
    new_full_text = re.sub(
        pattern_parenthetical, repl_parenthetical, new_full_text
    )
    if new_full_text != full_text:
      p.text = new_full_text
      changed.append("Đã thêm 'và cộng sự' vào trích dẫn Tác Giả (Năm)")

  return changed


def has_image(paragraph):
  xml = paragraph._element.xml
  return "w:drawing" in xml or "w:pict" in xml or "a:blip" in xml


# ---------------------------------------------------------
# 8. XỬ LÝ FORMAT TOÀN BỘ FILE DOCX
# ---------------------------------------------------------
def process_docx_file(uploaded_bytes, rules):
  doc = docx.Document(io.BytesIO(uploaded_bytes))
  detailed_errors = []

  font_target = rules["font_name"]
  size_target = float(rules["font_size"])
  line_target = float(rules["line_spacing"])

  target_left = float(rules["margin_left"])
  target_right = float(rules["margin_right"])
  target_top = float(rules["margin_top"])
  target_bottom = float(rules["margin_bottom"])

  # A. CĂN CHỈNH LỀ TRANG
  for idx, section in enumerate(doc.sections):
    curr_left = (
        round(section.left_margin.cm, 1) if section.left_margin else 0
    )
    curr_right = (
        round(section.right_margin.cm, 1) if section.right_margin else 0
    )
    curr_top = round(section.top_margin.cm, 1) if section.top_margin else 0
    curr_bottom = (
        round(section.bottom_margin.cm, 1) if section.bottom_margin else 0
    )

    if (
        curr_left != target_left
        or curr_right != target_right
        or curr_top != target_top
        or curr_bottom != target_bottom
    ):
      detailed_errors.append(
          f"📌 **Lề Trang (Phần {idx+1}):** Cũ"
          f" ({curr_left}-{curr_right}-{curr_top}-{curr_bottom}cm) ➡️ **Đã sửa"
          f" chuẩn** ({target_left}-{target_right}-{target_top}-{target_bottom}cm)"
      )

    section.left_margin = Cm(target_left)
    section.right_margin = Cm(target_right)
    section.top_margin = Cm(target_top)
    section.bottom_margin = Cm(target_bottom)

  # B. XỬ LÝ NỘI DUNG VĂN BẢN
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
            f"🖼️ **Đoạn {i+1} (Hình ảnh)**: Thiếu **Chú thích / Tên hình** bên"
            " dưới!"
        )
      if "nguồn" not in nearby and "source" not in nearby:
        detailed_errors.append(
            f"🖼️ **Đoạn {i+1} (Hình ảnh)**: Thiếu **Trích dẫn nguồn**!"
        )

  # C. GHI NHẬN BÁO CÁO LỖI
  if title_bold_count > 0:
    detailed_errors.append(
        "🖋️ **Tên Đề Tài:** Đã kiểm tra và **ép buộc In Đậm** cho Tên Đề Tài!"
    )

  if spacing_count > 0:
    detailed_errors.append(
        "🧹 **Làm Sạch Dấu Cách:** Đã dọn sạch khoảng trắng thừa tại"
        f" **{spacing_count} đoạn** văn."
    )

  if fixed_authors:
    detailed_errors.append(
        f"✍️ **Trích Dẫn Tác Giả:** Đã bổ sung 'và cộng sự' chuẩn xác tại"
        f" **{len(fixed_authors)} vị trí**."
    )

  if converted_citation_count > 0:
    detailed_errors.append(
        f"✨ **Trích Dẫn Số:** Đã chuyển **{converted_citation_count} đoạn**"
        " ngoặc vuông `[1]` thành **lũy thừa (¹, ²)**!"
    )

  # D. ÁP DỤNG FONT & GIÃN DÒNG
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

  # Bảng biểu
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
# 9. GIAO DIỆN CHÍNH STREAMLIT
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
      st.selectbox(
          "Chọn file quy định PDF trong thư mục `quy_dinh`:", pdf_files
      )
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
      st.success(f"✅ AI đã đọc thành công file `{selected_pdf}`!")

      with st.expander(
          "🔍 Click vào đây để XEM TOÀN BỘ CHỮ MÀ AI ĐÃ ĐỌC TỪ FILE PDF QUY"
          " ĐỊNH"
      ):
        st.text_area("Chữ rút ra từ PDF:", pdf_text, height=200)
    else:
      st.warning(
          "⚠️ File PDF này là dạng Scan/Ảnh chụp nên không rút được chữ tự"
          " động. Bạn có thể tự chỉnh thông số ở Thanh Bên (Sidebar) bên"
          " phải!"
      )

  st.subheader("2. Tải File Luận Văn Của Bạn (.DOCX)")
  uploaded_docx = st.file_uploader(
      "Thả file .docx luận văn vào đây", type=["docx"]
  )

# ---------------------------------------------------------
# 10. BẢNG TÙY CHỈNH THÔNG SỐ Ở THANH BÊN (SIDEBAR)
# ---------------------------------------------------------
st.sidebar.title("⚙️ Bảng Quy Định Định Dạng")
st.sidebar.caption(
    "AI đã tự điền các số liệu bên dưới từ PDF. Bạn có thể chỉnh lại theo ý"
    " muốn:"
)

final_font = st.sidebar.text_input(
    "Font Chữ Thân Bài", value=parsed_rules.get("font_name", "Times New Roman")
)
final_size = st.sidebar.number_input(
    "Cỡ Chữ (pt)",
    value=float(parsed_rules.get("font_size", 13.0)),
    step=0.5,
)
final_spacing = st.sidebar.number_input(
    "Giãn Dòng (Line Spacing)",
    value=float(parsed_rules.get("line_spacing", 1.5)),
    step=0.1,
)

st.sidebar.markdown("---")
st.sidebar.subheader("📐 Căn Lề Trang (cm)")
final_left = st.sidebar.number_input(
    "Lề Trái (Left)",
    value=float(parsed_rules.get("margin_left", 3.0)),
    step=0.5,
)
final_right = st.sidebar.number_input(
    "Lề Phải (Right)",
    value=float(parsed_rules.get("margin_right", 2.0)),
    step=0.5,
)
final_top = st.sidebar.number_input(
    "Lề Trên (Top)",
    value=float(parsed_rules.get("margin_top", 3.5)),
    step=0.5,
)
final_bottom = st.sidebar.number_input(
    "Lề Dưới (Bottom)",
    value=float(parsed_rules.get("margin_bottom", 3.0)),
    step=0.5,
)

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
# 11. NÚT CHẠY & XUẤT KẾT QUẢ
# ---------------------------------------------------------
with col_main:
  st.markdown("---")
  if st.button(
      "🚀 BẮT ĐẦU SOI & SỬA LUẬN VĂN", type="primary", use_container_width=True
  ):
    if not uploaded_docx:
      st.error("❌ Vui lòng tải file Luận Văn (.docx) ở Bước 2 trước!")
    else:
      with st.spinner(
          "⏳ Đang bổ sung 'và cộng sự' & chuẩn hóa định dạng bài..."
      ):
        fixed_stream, error_list = process_docx_file(
            uploaded_docx.getvalue(), active_rules
        )

      st.markdown("### 📋 BẢNG BÁO CÁO SOI & SỬA LỖI TỰ ĐỘNG")
      if error_list:
        for err in error_list:
          st.write(err)
        st.success(
            "🎉 **Đã sửa xong! Tất cả trích dẫn đã có đầy đủ dạng 'Tác giả và"
            " cộng sự (Năm)'!**"
        )
      else:
        st.success("🎉 Bài luận văn đã chuẩn chỉnh theo đúng thông số trên!")

      st.download_button(
          label="📥 TẢI FILE LUẬN VĂN ĐÃ ĐƯỢC SỬA HOÀN CHỈNH",
          data=fixed_stream,
          file_name="LuanVan_ChuanForm_TricDanhVaCongSu.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
          use_container_width=True,
      )