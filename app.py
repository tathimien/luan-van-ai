import difflib
import copy
import io
import json
import os
import re
import unicodedata
from collections import Counter

import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.text.paragraph import CT_P
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from pypdf import PdfReader

try:
    import streamlit as st
except ImportError:
    st = None

try:
    from google import genai
except ImportError:
    genai = None

try:
    import google.generativeai as genai_legacy
except ImportError:
    genai_legacy = None


DEFAULT_RULES = {
    "font_name": "Times New Roman",
    "font_size": 13.0,
    # Luận văn được phép dùng thống nhất cỡ 13 hoặc cỡ 14.
    "allowed_font_sizes": [13.0, 14.0],
    "line_spacing": 1.5,
    "margin_top": 3.5,
    "margin_bottom": 3.0,
    "margin_left": 3.5,
    "margin_right": 2.0,
    # Các khóa dưới đây quyết định những kiểm tra nghiệp vụ nào được chạy.
    # File PDF có thể ghi đè các giá trị này sau khi được đọc.
    "citation_style": "numeric_superscript",
    "require_image_source": True,
    "require_image_caption": True,
    "check_abbreviations": True,
    "check_subjectless_sentences": True,
    "normalize_references": True,
    "detailed_requirements": [
        "Bảng mã Unicode; Times New Roman cỡ 13 hoặc 14; giãn dòng 1,5.",
        "Lề trên 3,5 cm; dưới 3,0 cm; trái 3,5 cm; phải 2,0 cm.",
        "Tài liệu tham khảo: tên Việt Nam viết đầy đủ; tên nước ngoài "
        "ghi họ đầy đủ, tên đệm/tên gọi viết tắt; trên 3 tác giả ghi 3 "
        "tác giả đầu và cộng sự/et al.; năm trong ngoặc; tên bài in "
        "đứng; tên tạp chí in nghiêng; tập/số in đậm; trang chỉ ghi số.",
    ],
}


# Mỗi lựa chọn luôn gắn với đúng MỘT template và MỘT file quy định.
# Có thể đổi tên file tại đây, nhưng không nên cho học viên tự chọn hai tệp
# độc lập vì rất dễ ghép nhầm quy định của loại này với template của loại khác.
DOCUMENT_PROFILES = {
    "master_research": {
        "label": "Luận văn THS nghiên cứu - NCS - BSCK2",
        "short_label": "THS nghiên cứu - NCS - BSCK2",
        "template_file": (
            "Văn_Luận văn THS nghiên cứu -NCS-BSCK2 "
            "-Template 2026.docx"
        ),
        "template_aliases": [
            "luan_van_THS_nghien_cuu_NCS_BSCK2.docx",
            "Văn_Luận văn Thạc sĩ nghiên cứu -Template 2026.docx",
            "luan_van_thac_si_nghien_cuu.docx",
        ],
        "regulation_file": (
            "quy_dinh_luan_van_THS nghien cuu_NCS_BSCK2.pdf"
        ),
        "regulation_aliases": [
            "quy_dinh_luan_van_THS_nghien_cuu_NCS_BSCK2.pdf",
            "quy_dinh_luan_van_thac_si_nghien_cuu.pdf",
        ],
        "output_file": (
            "LuanVan_THS_NghienCuu_NCS_BSCK2_DaKiemTra.docx"
        ),
    },
    "master_application": {
        "label": "Luận văn thạc sĩ định hướng ứng dụng",
        "short_label": "Thạc sĩ ứng dụng",
        "template_file": (
            "Văn_Luận văn Thạc sĩ ứng dụng -Template 2026.docx"
        ),
        "template_aliases": ["luan_van_thac_si_ung_dung.docx"],
        "regulation_file": "quy_dinh_luan_van_thac_si_ung_dung.pdf",
        "output_file": "LuanVan_ThacSi_UngDung_DaKiemTra.docx",
    },
    "proposal_master_research": {
        "label": "Đề cương luận văn THS nghiên cứu - NCS - BSCK2",
        "short_label": "Đề cương THS nghiên cứu - NCS - BSCK2",
        "template_file": (
            "Văn_Template_Đề cương_Luận văn THS nghiên cứu "
            "-NCS-BSCK2.docx"
        ),
        "template_aliases": [
            "de_cuong_luan_van_THS_nghien_cuu_NCS_BSCK2.docx",
            "Văn_Template_Đề cương_Luận văn Thạc sĩ nghiên cứu "
            "-Template 2026.docx",
            "de_cuong_luan_van_thac_si_nghien_cuu.docx",
        ],
        "regulation_file": (
            "quy_dinh_de_cuong_luan_van_THS nghien cuu_"
            "NCS_BSCK2.pdf"
        ),
        "regulation_aliases": [
            "quy_dinh_de_cuong_luan_van_THS_nghien_cuu_"
            "NCS_BSCK2.pdf",
            "quy_dinh_de_cuong_luan_van_thac_si_nghien_cuu.pdf",
        ],
        "output_file": (
            "DeCuong_THS_NghienCuu_NCS_BSCK2_DaKiemTra.docx"
        ),
    },
    "proposal_master_application": {
        "label": "Đề cương luận văn thạc sĩ định hướng ứng dụng",
        "short_label": "Đề cương thạc sĩ ứng dụng",
        "template_file": (
            "Văn_Template_Đề cương_Luận văn Thạc sĩ ứng dụng.docx"
        ),
        "template_aliases": [
            "Văn_Template_Đề cương_Luận văn Thạc sĩ ứng dụng "
            "-Template 2026.docx",
            "de_cuong_luan_van_thac_si_ung_dung.docx",
        ],
        "regulation_file": (
            "quy_dinh_de_cuong_luan_van_thac_si_ung_dung.pdf"
        ),
        "output_file": "DeCuong_ThacSi_UngDung_DaKiemTra.docx",
    },
}


def _copy_default_rules():
    """Tạo bản sao sâu để các lần chạy Streamlit không dùng chung list."""
    return copy.deepcopy(DEFAULT_RULES)


def _filename_key(filename):
    """So tên file linh hoạt với dấu tiếng Việt và hậu tố (1), (2)."""
    filename_text = str(filename).strip()
    stem, extension = os.path.splitext(filename_text)
    # Trình duyệt/Windows thường tự thêm (1), (2) khi tải lại cùng tên.
    # Bỏ phần này để file mới vẫn khớp cấu hình mà không cần sửa code.
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    normalized = unicodedata.normalize(
        "NFD", f"{stem}{extension}"
    ).replace("Đ", "D")
    normalized = normalized.replace("đ", "d")
    normalized = "".join(
        char for char in normalized if unicodedata.category(char) != "Mn"
    )
    return re.sub(r"\s+", " ", normalized).strip().casefold()


def _filename_copy_number(filename):
    """Lấy số ở hậu tố (1), (2)...; tên gốc được xem là phiên bản 0."""
    stem, _ = os.path.splitext(str(filename).strip())
    match = re.search(r"\s*\((\d+)\)\s*$", stem)
    return int(match.group(1)) if match else 0


def _resolve_named_asset(folder, preferred_name, aliases=None):
    candidates = [preferred_name, *(aliases or [])]
    if os.path.isdir(folder):
        actual_names = {}
        for name in os.listdir(folder):
            if not os.path.isfile(os.path.join(folder, name)):
                continue
            actual_names.setdefault(_filename_key(name), []).append(name)
        for candidate in candidates:
            matching_names = actual_names.get(_filename_key(candidate), [])
            if matching_names:
                # Khi tên gốc và các bản (1), (2) cùng tồn tại, ưu tiên số
                # lớn nhất vì đây thường là file người quản trị mới tải lên.
                actual_name = max(
                    matching_names,
                    key=lambda name: (
                        _filename_copy_number(name),
                        name.casefold(),
                    ),
                )
                return os.path.join(folder, actual_name), actual_name
    return os.path.join(folder, preferred_name), preferred_name


def resolve_profile_paths(app_directory, profile):
    """Trả về đường dẫn tuyệt đối của đúng cặp template – quy định."""
    template_path, actual_template_name = _resolve_named_asset(
        os.path.join(app_directory, "template"),
        profile["template_file"],
        profile.get("template_aliases"),
    )
    regulation_path, actual_regulation_name = _resolve_named_asset(
        os.path.join(app_directory, "quy_dinh"),
        profile["regulation_file"],
        profile.get("regulation_aliases"),
    )
    return {
        **profile,
        "template_path": template_path,
        "regulation_path": regulation_path,
        "resolved_template_file": actual_template_name,
        "resolved_regulation_file": actual_regulation_name,
    }


def extract_raw_text_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        return ""
    try:
        reader = PdfReader(pdf_path)
        full_text = []
        for idx, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text.append(f"--- TRANG {idx + 1} ---\n{text}")
        return "\n".join(full_text).strip()
    except Exception as exc:
        if st is not None:
            st.error(f"Lỗi đọc file PDF: {exc}")
        return ""


def _vi_number(value, fallback=None):
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return fallback


def analyze_rules_locally(pdf_text):
    """Đọc các thông số phổ biến ngay cả khi máy chủ không có khóa AI.

    Đây là lớp dự phòng bắt buộc: trước đây nếu thiếu GEMINI_API_KEY thì PDF
    được hiển thị là "đã đọc" nhưng thực tế chương trình vẫn dùng DEFAULT_RULES.
    """
    rules = _copy_default_rules()
    if not pdf_text:
        return rules

    clean_text = re.sub(r"\s+", " ", pdf_text.replace("\xa0", " "))

    font_match = re.search(
        r"(?i)font(?:\s+chữ)?\s*(?:là|:)?\s*"
        r"(Times\s+New\s+Roman|Arial|Calibri|Tahoma)",
        clean_text,
    )
    if not font_match:
        font_match = re.search(
            r"(?i)\b(Times\s+New\s+Roman|Arial|Calibri|Tahoma)\b",
            clean_text,
        )
    if font_match:
        rules["font_name"] = re.sub(r"\s+", " ", font_match.group(1)).title()
        if rules["font_name"] == "Times New Roman":
            rules["font_name"] = "Times New Roman"

    sizes = []
    # Ưu tiên mẫu nêu rõ hai cỡ được phép cho thân bài. Không gom mọi con số
    # "cỡ" trong PDF vì cỡ tiêu đề/bìa thường khác cỡ chữ nội dung.
    range_match = re.search(
        r"(?i)cỡ(?:\s+chữ)?\s*(\d{1,2}(?:[.,]\d+)?)\s*"
        r"(?:hoặc|và|đến|[-–—/])\s*(\d{1,2}(?:[.,]\d+)?)",
        clean_text,
    )
    if range_match:
        for value in range_match.groups():
            number = _vi_number(value)
            if number and 8 <= number <= 30 and number not in sizes:
                sizes.append(number)
    if not sizes:
        body_size_match = re.search(
            r"(?i)(?:nội\s+dung|thân\s+bài|toàn\s+văn|luận\s+văn|"
            r"đề\s+cương|Times\s+New\s+Roman).{0,100}?"
            r"(?:cỡ|size)\s*(?:chữ\s*)?(?:là|:)?\s*"
            r"(\d{1,2}(?:[.,]\d+)?)",
            clean_text,
        )
        if body_size_match:
            number = _vi_number(body_size_match.group(1))
            if number and 8 <= number <= 30:
                sizes.append(number)
    if sizes:
        rules["allowed_font_sizes"] = sizes
        rules["font_size"] = sizes[0]

    spacing_match = re.search(
        r"(?i)(?:giãn|dãn|khoảng\s+cách)\s*dòng[^\d]{0,30}"
        r"(\d(?:[.,]\d+)?)",
        clean_text,
    )
    if spacing_match:
        spacing = _vi_number(spacing_match.group(1))
        if spacing and 0.8 <= spacing <= 3:
            rules["line_spacing"] = spacing

    margin_patterns = {
        "margin_top": r"(?i)lề\s+trên[^\d]{0,20}(\d+(?:[.,]\d+)?)\s*(?:cm)?",
        "margin_bottom": r"(?i)lề\s+dưới[^\d]{0,20}(\d+(?:[.,]\d+)?)\s*(?:cm)?",
        "margin_left": r"(?i)lề\s+trái[^\d]{0,20}(\d+(?:[.,]\d+)?)\s*(?:cm)?",
        "margin_right": r"(?i)lề\s+phải[^\d]{0,20}(\d+(?:[.,]\d+)?)\s*(?:cm)?",
    }
    for key, pattern in margin_patterns.items():
        match = re.search(pattern, clean_text)
        if match:
            value = _vi_number(match.group(1))
            if value and 0.5 <= value <= 8:
                rules[key] = value

    if re.search(r"(?i)\b(?:APA|Harvard|tác\s+giả\s*[-–]\s*năm)\b", clean_text):
        rules["citation_style"] = "author_year"
    elif re.search(r"(?i)(?:ngoặc\s+vuông|\[\s*\d+\s*\])", clean_text):
        rules["citation_style"] = "numeric_brackets"
    elif re.search(r"(?i)(?:số\s+mũ|lũy\s+thừa|superscript)", clean_text):
        rules["citation_style"] = "numeric_superscript"

    meaningful_lines = []
    for raw_line in pdf_text.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if (
            not line
            or line.startswith("--- TRANG")
            or len(line) < 18
            or line in meaningful_lines
        ):
            continue
        meaningful_lines.append(line)
        if len(meaningful_lines) >= 30:
            break
    if meaningful_lines:
        rules["detailed_requirements"] = meaningful_lines
    return rules


def _merge_rule_payload(base_rules, payload):
    """Chỉ nhận các khóa đã biết và giữ giá trị dự phòng khi AI thiếu khóa."""
    if not isinstance(payload, dict):
        return base_rules
    merged = copy.deepcopy(base_rules)
    allowed_keys = set(DEFAULT_RULES)
    for key, value in payload.items():
        if key in allowed_keys and value is not None:
            merged[key] = value

    numeric_keys = {
        "font_size",
        "line_spacing",
        "margin_top",
        "margin_bottom",
        "margin_left",
        "margin_right",
    }
    for key in numeric_keys:
        parsed = _vi_number(merged.get(key), base_rules.get(key))
        merged[key] = parsed

    raw_sizes = merged.get("allowed_font_sizes", [])
    if not isinstance(raw_sizes, (list, tuple, set)):
        raw_sizes = [raw_sizes]
    sizes = []
    for value in raw_sizes:
        parsed = _vi_number(value)
        if parsed and 8 <= parsed <= 30 and parsed not in sizes:
            sizes.append(parsed)
    merged["allowed_font_sizes"] = sizes or list(
        base_rules.get("allowed_font_sizes", [merged["font_size"]])
    )

    boolean_keys = {
        "require_image_source",
        "require_image_caption",
        "check_abbreviations",
        "check_subjectless_sentences",
        "normalize_references",
    }
    for key in boolean_keys:
        value = merged.get(key)
        if isinstance(value, str):
            merged[key] = value.strip().lower() not in {
                "false", "0", "no", "không", "none",
            }
        else:
            merged[key] = bool(value)

    allowed_citation_styles = {
        "numeric_superscript",
        "numeric_brackets",
        "author_year",
        "keep",
    }
    citation_style = str(merged.get("citation_style", "keep")).lower()
    if citation_style not in allowed_citation_styles:
        citation_style = str(base_rules.get("citation_style", "keep"))
    merged["citation_style"] = citation_style

    requirements = merged.get("detailed_requirements", [])
    if isinstance(requirements, str):
        requirements = [requirements]
    merged["detailed_requirements"] = [
        str(item).strip() for item in requirements if str(item).strip()
    ]
    return merged


def analyze_rules_with_gemini(pdf_text, api_key):
    local_rules = analyze_rules_locally(pdf_text)
    if not pdf_text or not api_key:
        return local_rules

    prompt = f"""
Bạn là chuyên gia kiểm tra định dạng luận văn. Hãy đọc TOÀN BỘ văn bản quy
định dưới đây và trích xuất thông số kỹ thuật chuẩn xác.

NỘI DUNG FILE QUY ĐỊNH PDF:
{pdf_text}

Hãy trả về DUY NHẤT một chuỗi JSON thuần có cấu trúc sau, không dùng mã
Markdown:
{{
  "font_name": "Times New Roman",
  "font_size": 13.0,
  "allowed_font_sizes": [13.0, 14.0],
  "line_spacing": 1.5,
  "margin_top": 3.5,
  "margin_bottom": 3.0,
  "margin_left": 3.5,
  "margin_right": 2.0,
  "citation_style": "numeric_superscript | numeric_brackets | author_year | keep",
  "require_image_source": true,
  "require_image_caption": true,
  "check_abbreviations": true,
  "check_subjectless_sentences": true,
  "normalize_references": true,
  "detailed_requirements": ["Các quy định cụ thể tìm thấy trong PDF"]
}}
"""
    candidate_models = [
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro",
    ]

    for model_name in candidate_models:
        if genai is not None:
            try:
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                )
                if response and response.text:
                    raw_text = response.text.strip()
                    fence = chr(96) * 3
                    clean_json = (
                        raw_text.replace(f"{fence}json", "")
                        .replace(fence, "")
                        .strip()
                    )
                    return _merge_rule_payload(
                        local_rules,
                        json.loads(clean_json),
                    )
            except Exception:
                pass

        if genai_legacy is not None:
            try:
                genai_legacy.configure(api_key=api_key)
                model = genai_legacy.GenerativeModel(model_name)
                response = model.generate_content(prompt)
                if response and response.text:
                    raw_text = response.text.strip()
                    fence = chr(96) * 3
                    clean_json = (
                        raw_text.replace(f"{fence}json", "")
                        .replace(fence, "")
                        .strip()
                    )
                    return _merge_rule_payload(
                        local_rules,
                        json.loads(clean_json),
                    )
            except Exception:
                pass

    return local_rules


def _load_rules_for_pdf(pdf_path, file_mtime, api_key):
    # file_mtime là một phần của khóa cache: thay PDF thì quy định được đọc lại.
    del file_mtime
    pdf_text = extract_raw_text_from_pdf(pdf_path)
    return pdf_text, analyze_rules_with_gemini(pdf_text, api_key)


if st is not None:
    load_rules_for_pdf = st.cache_data(show_spinner=False)(
        _load_rules_for_pdf
    )
else:
    load_rules_for_pdf = _load_rules_for_pdf


def _paragraph_has_hard_page_break(paragraph):
    """Kiểm tra ngắt trang thủ công nằm trong một đoạn Word."""
    xml = paragraph._element.xml
    return "w:br" in xml and 'type="page"' in xml


def _has_page_boundary_before(paragraphs, index):
    """Kiểm tra vị trí đã có ngắt trang trước đoạn ``index`` hay chưa."""
    paragraph = paragraphs[index]
    if paragraph.paragraph_format.page_break_before:
        return True

    # Ngắt trang thường nằm ở đoạn có chữ "HÀ NỘI – 2026"; giữa đoạn
    # đó và tiêu đề trang sau đôi khi có vài đoạn rỗng nên quét lùi.
    for previous_index in range(index - 1, max(-1, index - 6), -1):
        previous = paragraphs[previous_index]
        if _paragraph_has_hard_page_break(previous):
            return True
        if previous.text.strip():
            break
    return False


def _ensure_cover_page_boundaries(doc):
    """Khóa bìa chính, bìa phụ và phần nội dung trên các trang riêng.

    Một số file học viên chỉ có ``lastRenderedPageBreak`` do Word tự dàn
    trang, không có ngắt trang thủ công. Khi chương trình chỉnh cỡ chữ hoặc
    khoảng cách, phần bìa ngắn đi và nội dung phía sau có thể bị kéo lên.
    Hàm này đặt ``pageBreakBefore`` ở đầu bìa phụ và đầu phần nội dung khi
    vị trí đó chưa có ngắt trang thật.
    """
    paragraphs = list(doc.paragraphs)
    if not paragraphs:
        return []

    normalized_headings = [
        _normalize_heading(paragraph.text) for paragraph in paragraphs
    ]
    body_start_keys = {
        "LOI CAM ON",
        "LOI CAM DOAN",
        "MUC LUC",
        "DANH MUC CHU VIET TAT",
        "DANH MUC BANG",
        "DANH MUC HINH",
        "DANH MUC SO DO",
        "DAT VAN DE",
        "TONG QUAN",
        "INTRODUCTION",
    }
    body_start_index = next(
        (
            index
            for index, heading in enumerate(normalized_headings)
            if heading in body_start_keys
        ),
        None,
    )

    # Bìa chính và bìa phụ thường cùng bắt đầu bằng dòng Bộ GD&ĐT/Bộ Y tế.
    # Chỉ tìm trong vùng trước nội dung để không nhận nhầm các trang sau.
    search_end = body_start_index if body_start_index is not None else 80
    cover_start_indices = [
        index
        for index, heading in enumerate(normalized_headings[:search_end])
        if "BO GIAO DUC" in heading
    ]

    protected_boundaries = []
    if len(cover_start_indices) >= 2:
        second_cover_index = cover_start_indices[1]
        if not _has_page_boundary_before(paragraphs, second_cover_index):
            paragraphs[
                second_cover_index
            ].paragraph_format.page_break_before = True
            protected_boundaries.append("bìa phụ")

    if (
        body_start_index is not None
        and not _has_page_boundary_before(paragraphs, body_start_index)
    ):
        paragraphs[body_start_index].paragraph_format.page_break_before = True
        protected_boundaries.append("phần nội dung")

    return protected_boundaries


def optimize_cover_pages(doc):
    cover_errors = []
    cover_p_elements = set()
    cover_paragraphs = []

    protected_boundaries = _ensure_cover_page_boundaries(doc)

    cover_table_keywords = [
        "BỘ GIÁO DỤC",
        "BỘ Y TẾ",
        "TRƯỜNG ĐẠI HỌC",
        "ĐẠI HỌC Y",
        "LUẬN VĂN",
        "LUẬN ÁN",
        "KHÓA LUẬN",
        "TÊN ĐỀ TÀI",
        "NGƯỜI HƯỚNG DẪN",
    ]
    for table in doc.tables[:2]:
        table_text = " ".join(
            cell.text.upper()
            for row in table.rows
            for cell in row.cells
        )
        if any(word in table_text for word in cover_table_keywords):
            for row in table.rows:
                for cell in row.cells:
                    cover_paragraphs.extend(cell.paragraphs)

    page_groups = [[]]
    page_break_count = 0
    for paragraph in doc.paragraphs:
        page_groups[-1].append(paragraph)
        xml = paragraph._element.xml
        is_break = (
            "\x0c" in paragraph.text
            or ("w:br" in xml and 'type="page"' in xml)
            or "w:sectPr" in xml
        )
        if is_break:
            page_break_count += 1
            page_groups.append([])
            if page_break_count >= 2:
                break

    if page_break_count >= 2:
        for page_index in range(2):
            cover_paragraphs.extend(page_groups[page_index])
    else:
        # DOCX không lưu các ngắt trang do Word tự dàn trang. Nếu không
        # có đủ hai ngắt trang thủ công, tuyệt đối không coi toàn bộ luận
        # văn là trang bìa vì như vậy hình trong thân bài sẽ bị bỏ qua.
        body_start_keywords = [
            "LỜI CAM ĐOAN",
            "LỜI CẢM ƠN",
            "MỤC LỤC",
            "DANH MỤC",
            "ĐẶT VẤN ĐỀ",
            "TỔNG QUAN",
            "CHƯƠNG 1",
            "CHAPTER 1",
            "INTRODUCTION",
        ]
        for paragraph in doc.paragraphs[:25]:
            heading = paragraph.text.upper().strip()
            if any(
                heading.startswith(keyword)
                for keyword in body_start_keywords
            ):
                break
            cover_paragraphs.append(paragraph)

    for paragraph in cover_paragraphs:
        cover_p_elements.add(paragraph._element)
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(1)
        if not paragraph.text.strip():
            if not paragraph.runs:
                paragraph.add_run()
            for run in paragraph.runs:
                run.font.size = Pt(6)

    cover_errors.append(
        "🖼️ **Bảo toàn Trang Bìa & Logo:** Đã giữ nguyên khung viền, "
        "logo trường, co nhỏ dòng trống (6pt) và cố định trang bìa."
    )
    if protected_boundaries:
        cover_errors.append(
            "📄 **Khóa ngắt trang:** Đã tạo ngắt trang cố định trước "
            + " và ".join(protected_boundaries)
            + "; nội dung phía sau sẽ không bị đẩy ngược lên trang bìa."
        )
    return cover_errors, cover_p_elements


def is_cover_or_logo_image(p_idx, paragraph, cover_p_elements, paragraphs):
    if paragraph._element in cover_p_elements:
        return True

    if p_idx < 25:
        start_index = max(0, p_idx - 6)
        end_index = min(len(paragraphs), p_idx + 7)
        context_text = " ".join(
            paragraphs[index].text.upper()
            for index in range(start_index, end_index)
        )
        cover_keywords = [
            "BỘ GIÁO DỤC",
            "BỘ Y TẾ",
            "TRƯỜNG ĐẠI HỌC",
            "ĐẠI HỌC Y",
            "LUẬN VĂN",
            "LUẬN ÁN",
            "KHÓA LUẬN",
            "TÊN ĐỀ TÀI",
            "NGƯỜI HƯỚNG DẪN",
        ]
        cover_score = sum(
            keyword in context_text for keyword in cover_keywords
        )
        if cover_score >= 2:
            return True
    return False


def get_reference_heading_format(doc):
    ref_pattern = re.compile(r"^\s*2\.[1-3]\.?", re.IGNORECASE)
    for paragraph in doc.paragraphs:
        if ref_pattern.search(paragraph.text.strip()):
            fmt = paragraph.paragraph_format
            return {
                "left_indent": fmt.left_indent,
                "first_line_indent": fmt.first_line_indent,
                "alignment": fmt.alignment,
                "space_before": fmt.space_before,
                "space_after": fmt.space_after,
            }
    return {
        "left_indent": Pt(0),
        "first_line_indent": Pt(0),
        "alignment": None,
        "space_before": Pt(3),
        "space_after": Pt(3),
    }


def fix_ethics_section(doc):
    detailed_errors = []
    # Chỉ nhận diện khi TOÀN BỘ đoạn là tiêu đề. Không dùng ``search`` vì
    # cụm "đạo đức trong nghiên cứu" có thể xuất hiện giữa câu nội dung.
    ethics_pattern = re.compile(
        r"^\s*(?:2\.\d+\.?\s*)?"
        r"(?:vấn đề\s+)?đạo\s+đức(?:\s+trong)?\s+nghiên\s+cứu"
        r"\s*[.:]?\s*$",
        re.IGNORECASE,
    )
    false_split_pattern = re.compile(
        r"^\s*(?:vấn đề\s+)?đạo\s+đức(?:\s+trong)?\s+nghiên\s+cứu\b",
        re.IGNORECASE,
    )
    ref_fmt = get_reference_heading_format(doc)

    def apply_heading_format(paragraph):
        fmt = paragraph.paragraph_format
        fmt.left_indent = ref_fmt["left_indent"]
        fmt.first_line_indent = ref_fmt["first_line_indent"]
        if ref_fmt["alignment"] is not None:
            fmt.alignment = ref_fmt["alignment"]
        if ref_fmt["space_before"] is not None:
            fmt.space_before = ref_fmt["space_before"]
        if ref_fmt["space_after"] is not None:
            fmt.space_after = ref_fmt["space_after"]

    # Sửa lại tài liệu đã bị phiên bản cũ tách sai: đoạn sau bắt đầu bằng
    # cụm trên, run đầu được in đậm + bôi vàng, còn đoạn trước chưa kết câu.
    paragraphs = list(doc.paragraphs)
    for idx in range(1, len(paragraphs)):
        paragraph = paragraphs[idx]
        previous = paragraphs[idx - 1]
        text = paragraph.text.strip()
        if not text or ethics_pattern.fullmatch(text):
            continue
        first_run = next((run for run in paragraph.runs if run.text), None)
        was_old_false_split = (
            false_split_pattern.match(text)
            and first_run is not None
            and first_run.bold is True
            and first_run.font.highlight_color == WD_COLOR_INDEX.YELLOW
            and previous.text.strip()
            and not re.search(r"[.!?:;]\s*$", previous.text)
        )
        if not was_old_false_split:
            continue

        if not previous.text.endswith((" ", "\t")):
            previous.add_run(" ")
        first_run.text = first_run.text.lstrip()
        first_run.bold = False
        first_run.font.highlight_color = None
        for run in list(paragraph.runs):
            previous._p.append(run._r)
        paragraph._element.getparent().remove(paragraph._element)
        detailed_errors.append(
            f"🧹 **Đoạn {idx + 1}:** Đã nối lại câu từng bị tách sai "
            "tại cụm 'đạo đức trong nghiên cứu'."
        )

    for idx, paragraph in enumerate(list(doc.paragraphs)):
        text = paragraph.text.strip()
        if not text or not ethics_pattern.fullmatch(text):
            continue

        apply_heading_format(paragraph)
        if not paragraph.runs:
            paragraph.add_run(text)
        for run in paragraph.runs:
            if not run.text:
                continue
            run.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        detailed_errors.append(
            f"📏 **Mục {text}:** Đã nhận diện đúng là tiêu đề, căn "
            "thẳng hàng với 2.1, 2.2, 2.3 và bôi vàng; không tách "
            "các câu nội dung có cùng cụm từ."
        )
    return detailed_errors


def enforce_bold_for_thesis_title(paragraph):
    text_lower = paragraph.text.lower().strip()
    is_title = (
        "tên đề tài" in text_lower
        or text_lower.startswith("đề tài:")
        or text_lower.startswith("đề tài :")
    )
    if is_title:
        for run in paragraph.runs:
            run.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return True
    return False


def _style_from_run(run):
    return (
        bool(run.bold) if run.bold is not None else False,
        bool(run.italic) if run.italic is not None else False,
        bool(run.font.superscript)
        if run.font.superscript is not None
        else False,
        bool(run.font.subscript) if run.font.subscript is not None else False,
        run.font.name,
        run.font.size,
        run.font.highlight_color,
    )


def _paragraph_char_records(paragraph):
    records = []
    for run in paragraph.runs:
        style = _style_from_run(run)
        records.extend((char, style) for char in run.text)
    return records


def _style_with(style, superscript=None, subscript=None, highlight=None):
    values = list(style)
    if superscript is not None:
        values[2] = superscript
    if subscript is not None:
        values[3] = subscript
    if highlight is not None:
        values[6] = highlight
    return tuple(values)


def _apply_style(run, style):
    run.bold = style[0]
    run.italic = style[1]
    run.font.superscript = style[2]
    run.font.subscript = style[3]
    if style[4]:
        run.font.name = style[4]
    if style[5]:
        run.font.size = style[5]
    if style[6] is not None:
        run.font.highlight_color = style[6]


def _rebuild_paragraph(paragraph, records):
    paragraph.text = ""
    if not records:
        return

    current_style = records[0][1]
    current_text = ""
    for char, style in records:
        if style == current_style:
            current_text += char
        else:
            run = paragraph.add_run(current_text)
            _apply_style(run, current_style)
            current_text = char
            current_style = style
    if current_text:
        run = paragraph.add_run(current_text)
        _apply_style(run, current_style)


def set_paragraph_text_preserve_formatting(
    paragraph,
    new_text,
    highlight_changes=False,
    highlight_ranges=None,
):
    old_text = paragraph.text
    old_records = _paragraph_char_records(paragraph)
    if not old_records or not new_text:
        paragraph.text = new_text
        return

    default_style = old_records[-1][1]
    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    new_records = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            new_records.extend(old_records[i1:i2])
            continue
        if tag == "delete":
            continue

        source_index = min(i1, len(old_records) - 1)
        if tag == "insert" and i1 > 0:
            source_index = i1 - 1
        style = old_records[source_index][1] if old_records else default_style
        if highlight_changes:
            style = _style_with(style, highlight=WD_COLOR_INDEX.YELLOW)
        new_records.extend((char, style) for char in new_text[j1:j2])

    if len(new_records) != len(new_text):
        paragraph.text = new_text
        return

    for item in highlight_ranges or []:
        if len(item) == 3:
            start, end, highlight_color = item
        else:
            start, end = item
            highlight_color = WD_COLOR_INDEX.YELLOW
        for index in range(max(0, start), min(end, len(new_records))):
            char, style = new_records[index]
            new_records[index] = (
                char,
                _style_with(style, highlight=highlight_color),
            )
    _rebuild_paragraph(paragraph, new_records)


def highlight_paragraph_ranges(
    paragraph,
    ranges,
    highlight_color=WD_COLOR_INDEX.YELLOW,
    preserve_existing=True,
):
    records = _paragraph_char_records(paragraph)
    if not records:
        return False

    changed = False
    for start, end in _merge_ranges(ranges):
        for index in range(max(0, start), min(end, len(records))):
            char, style = records[index]
            if preserve_existing and style[6] is not None:
                continue
            records[index] = (
                char,
                _style_with(style, highlight=highlight_color),
            )
            changed = True
    if changed:
        _rebuild_paragraph(paragraph, records)
    return changed


NUMERIC_CITATION_CONTENT = r"\d+(?:\s*[,;–—-]\s*\d+)*"
NUMERIC_CITATION_PATTERN = re.compile(
    rf"(?:\[\s*{NUMERIC_CITATION_CONTENT}\s*\])"
    rf"(?:\s*[,;]?\s*\[\s*{NUMERIC_CITATION_CONTENT}\s*\])*"
)
NUMERIC_CITATION_ITEM_PATTERN = re.compile(
    rf"\[\s*({NUMERIC_CITATION_CONTENT})\s*\]"
)
SENTENCE_PUNCTUATION = ".!?"
UNICODE_SUPERSCRIPT_DIGITS = {
    "⁰": "0",
    "¹": "1",
    "²": "2",
    "³": "3",
    "⁴": "4",
    "⁵": "5",
    "⁶": "6",
    "⁷": "7",
    "⁸": "8",
    "⁹": "9",
}
SUPERSCRIPT_CITATION_CHARS = set("0123456789,;–—- ")
SUPERSCRIPT_CITATION_CHARS.update(UNICODE_SUPERSCRIPT_DIGITS)


def _is_superscript_citation_record(record):
    char, style = record
    return (
        (style[2] and char in SUPERSCRIPT_CITATION_CHARS)
        or char in UNICODE_SUPERSCRIPT_DIGITS
    )


def _is_probable_superscript_citation(records, start, end):
    """Phân biệt số trích dẫn với số mũ của đơn vị như cm², mm², m³."""
    cluster = "".join(char for char, _ in records[start:end])
    normalized_cluster = "".join(
        UNICODE_SUPERSCRIPT_DIGITS.get(char, char) for char in cluster
    )
    if not any(char.isdigit() for char in normalized_cluster):
        return False

    # Cụm 1,2 hoặc 3-5 gần như chắc chắn là một trích dẫn.
    if any(mark in normalized_cluster for mark in ",;–—-"):
        return True

    prefix = "".join(char for char, _ in records[:start]).rstrip()
    exponent = re.sub(r"\s+", "", normalized_cluster)
    unit_match = re.search(
        r"(?i)(?<!\w)(mm|cm|dm|km|m|nm|µm|μm|kg|mg|g|µg|μg|ml|dl|cl|l|ha)$",
        prefix,
    )
    if exponent in {"2", "3"} and unit_match:
        return False
    if prefix and prefix[-1] in "+−×÷*/=^":
        return False
    return True


def _highlight_record(record):
    char, style = record
    return (
        char,
        _style_with(style, highlight=WD_COLOR_INDEX.YELLOW),
    )


def _highlight_superscript_record(record):
    char, style = record
    return (
        UNICODE_SUPERSCRIPT_DIGITS.get(char, char),
        _style_with(
            style,
            superscript=True,
            subscript=False,
            highlight=WD_COLOR_INDEX.YELLOW,
        ),
    )


def normalize_numeric_citations(paragraph, font_target, size_target):
    """
    Chuẩn hóa trích dẫn số thành lũy thừa và đặt sau dấu chấm.

    Ví dụ:
        Nội dung [1].  -> Nội dung.¹
        Nội dung. [2] -> Nội dung.²
        Nội dung¹,².  -> Nội dung.¹,²
        Nội dung¹,²   -> Nội dung.¹,²
    """
    text = paragraph.text
    records = _paragraph_char_records(paragraph)
    matches = list(NUMERIC_CITATION_PATTERN.finditer(text))
    converted_count = len(matches)
    moved_count = 0

    if records and matches:
        rebuilt = []
        cursor = 0
        for match in matches:
            rebuilt.extend(records[cursor : match.start()])

            after_index = match.end()
            while after_index < len(text) and text[after_index].isspace():
                after_index += 1

            punctuation_after = (
                after_index < len(text)
                and text[after_index] in SENTENCE_PUNCTUATION
            )
            previous_nonspace = match.start() - 1
            while previous_nonspace >= 0 and text[previous_nonspace].isspace():
                previous_nonspace -= 1
            punctuation_before = (
                previous_nonspace >= 0
                and text[previous_nonspace] in SENTENCE_PUNCTUATION
            )
            citation_at_end = after_index >= len(text)

            while rebuilt and rebuilt[-1][0].isspace():
                rebuilt.pop()

            if punctuation_after:
                # Bôi vàng cả dấu chấm vừa được đưa ra trước số trích dẫn.
                punctuation_record = _highlight_record(
                    records[after_index]
                )
                rebuilt.append(punctuation_record)
                cursor = after_index + 1
                moved_count += 1
            else:
                cursor = match.end()
                if citation_at_end and not punctuation_before:
                    base_style = (
                        rebuilt[-1][1]
                        if rebuilt
                        else records[match.start()][1]
                    )
                    normal_style = _style_with(
                        base_style,
                        superscript=False,
                        subscript=False,
                        highlight=WD_COLOR_INDEX.YELLOW,
                    )
                    rebuilt.append((".", normal_style))
                    moved_count += 1
                elif punctuation_before and rebuilt:
                    # Trường hợp đã viết ". [1]": bôi vàng dấu chấm
                    # cùng với số lũy thừa để người dùng thấy chỗ sửa.
                    rebuilt[-1] = _highlight_record(rebuilt[-1])

            # Gộp [1],[2], [1] [2] hoặc [1,2] thành một cụm 1,2.
            citation_parts = []
            for item in NUMERIC_CITATION_ITEM_PATTERN.finditer(
                match.group(0)
            ):
                normalized_item = re.sub(
                    r"\s*([,;–—-])\s*",
                    r"\1",
                    item.group(1),
                )
                citation_parts.append(normalized_item)
            citation_text = ",".join(citation_parts)
            citation_style = _style_with(
                records[match.start()][1],
                superscript=True,
                subscript=False,
                highlight=WD_COLOR_INDEX.YELLOW,
            )
            if font_target:
                citation_style = (
                    citation_style[0],
                    citation_style[1],
                    citation_style[2],
                    citation_style[3],
                    font_target,
                    Pt(size_target),
                    citation_style[6],
                )
            rebuilt.extend((char, citation_style) for char in citation_text)

        rebuilt.extend(records[cursor:])
        records = rebuilt

    if not records:
        return {
            "converted": 0,
            "moved_after_period": 0,
            "changed": False,
        }

    normalized = []
    index = 0
    while index < len(records):
        if not _is_superscript_citation_record(records[index]):
            normalized.append(records[index])
            index += 1
            continue

        end_index = index
        has_digit = False
        while True:
            while (
                end_index < len(records)
                and _is_superscript_citation_record(records[end_index])
            ):
                has_digit = (
                    has_digit or records[end_index][0].isdigit()
                )
                end_index += 1

            # Word có thể tách dấu phẩy thành một run thường:
            # số ¹ (superscript), dấu phẩy (normal), số ² (superscript).
            # Vẫn phải coi toàn bộ ¹,² là một cụm trích dẫn.
            candidate_index = end_index
            while (
                candidate_index < len(records)
                and (
                    records[candidate_index][0].isspace()
                    or records[candidate_index][0] in ",;"
                )
            ):
                candidate_index += 1

            if (
                candidate_index < len(records)
                and _is_superscript_citation_record(
                    records[candidate_index]
                )
                and records[candidate_index][0].isdigit()
            ):
                end_index = candidate_index
                continue
            break

        next_index = end_index
        while next_index < len(records) and records[next_index][0].isspace():
            next_index += 1

        has_digit = has_digit and _is_probable_superscript_citation(
            records,
            index,
            end_index,
        )

        if (
            has_digit
            and next_index < len(records)
            and records[next_index][0] in SENTENCE_PUNCTUATION
            and not records[next_index][1][2]
        ):
            while normalized and normalized[-1][0].isspace():
                normalized.pop()
            normalized.append(_highlight_record(records[next_index]))
            normalized.extend(
                _highlight_superscript_record(record)
                for record in records[index:end_index]
                if not record[0].isspace()
            )
            index = next_index + 1
            moved_count += 1
        elif has_digit and (
            next_index >= len(records)
            or records[next_index][0].isupper()
        ):
            # Học viên đã đặt số ở dạng lũy thừa nhưng quên dấu chấm.
            # Xử lý cả cuối đoạn ("Nội dung¹,²") và giữa đoạn khi câu
            # kế tiếp bắt đầu bằng chữ hoa ("Nội dung¹,² Câu tiếp...").
            previous_nonspace = len(normalized) - 1
            while (
                previous_nonspace >= 0
                and normalized[previous_nonspace][0].isspace()
            ):
                previous_nonspace -= 1
            already_after_punctuation = (
                previous_nonspace >= 0
                and normalized[previous_nonspace][0]
                in SENTENCE_PUNCTUATION
            )
            if not already_after_punctuation:
                while normalized and normalized[-1][0].isspace():
                    normalized.pop()
                base_style = (
                    normalized[-1][1]
                    if normalized
                    else records[index][1]
                )
                normalized.append(
                    (
                        ".",
                        _style_with(
                            base_style,
                            superscript=False,
                            subscript=False,
                            highlight=WD_COLOR_INDEX.YELLOW,
                        ),
                    )
                )
                normalized.extend(
                    _highlight_superscript_record(record)
                    for record in records[index:end_index]
                    if not record[0].isspace()
                )
                # Cuối đoạn: bỏ dấu cách thừa phía sau trích dẫn. Nếu
                # còn câu tiếp theo, luôn tạo đúng một dấu cách thường;
                # dấu cách cũ đôi khi nằm ngay trong run lũy thừa nên
                # sẽ bị mất nếu chỉ loại phần định dạng số.
                if next_index < len(records):
                    normalized.append(
                        (
                            " ",
                            _style_with(
                                base_style,
                                superscript=False,
                                subscript=False,
                            ),
                        )
                    )
                index = next_index
                moved_count += 1
            else:
                normalized.extend(records[index:end_index])
                index = end_index
        else:
            normalized.extend(records[index:end_index])
            index = end_index

    old_signature = [
        (char, style[2], style[6]) for char, style in _paragraph_char_records(paragraph)
    ]
    new_signature = [(char, style[2], style[6]) for char, style in normalized]
    changed = old_signature != new_signature
    if changed:
        _rebuild_paragraph(paragraph, normalized)

    return {
        "converted": converted_count,
        "moved_after_period": moved_count,
        "changed": changed,
    }


def convert_brackets_to_superscript(paragraph, font_target, size_target):
    result = normalize_numeric_citations(
        paragraph,
        font_target,
        size_target,
    )
    return result["changed"]


def clean_spaces_and_punctuation(paragraph):
    if not paragraph.runs:
        return False

    text_upper = paragraph.text.upper()
    cover_keywords = [
        "BỘ GIÁO DỤC",
        "BỘ Y TẾ",
        "TRƯỜNG ĐẠI HỌC",
        "VIỆN NGHIÊN CỨU",
        "UBND",
    ]
    if any(keyword in text_upper for keyword in cover_keywords):
        return False

    changed = False
    for run in paragraph.runs:
        if not run.text:
            continue
        original = run.text
        cleaned = original.replace("\xa0", " ").replace("\u200b", "")
        cleaned = re.sub(r" {2,}", " ", cleaned)
        cleaned = re.sub(r"\s+([:;,.!?])", r"\1", cleaned)
        if cleaned != original:
            run.text = cleaned
            changed = True

    runs_with_text = [run for run in paragraph.runs if run.text]
    for index in range(len(runs_with_text) - 1):
        current_run = runs_with_text[index]
        next_run = runs_with_text[index + 1]
        while (
            current_run.text
            and next_run.text
            and current_run.text[-1] == " "
            and next_run.text[0] == " "
        ):
            next_run.text = next_run.text[1:]
            changed = True

        # Word có thể tách dấu cách và dấu câu sang hai run khác nhau,
        # ví dụ run 1 = "nội dung " và run 2 = ".". Xóa dấu cách thừa
        # nhưng giữ nguyên kiểu chữ của từng run.
        if (
            current_run.text
            and next_run.text
            and current_run.text.endswith(" ")
            and next_run.text[0] in ":;,.!?"
        ):
            stripped = current_run.text.rstrip(" ")
            if stripped != current_run.text:
                current_run.text = stripped
                changed = True

    if runs_with_text:
        first_run = runs_with_text[0]
        last_run = runs_with_text[-1]
        left_stripped = first_run.text.lstrip(" \t\xa0")
        if left_stripped != first_run.text:
            first_run.text = left_stripped
            changed = True
        right_stripped = last_run.text.rstrip(" \t\xa0")
        if right_stripped != last_run.text:
            last_run.text = right_stripped
            changed = True
    return changed


AUTHOR_WORD = r"[A-ZÀ-Ỹ][A-Za-zÀ-ỹ'’.-]*"
AUTHOR_NAME = rf"{AUTHOR_WORD}(?:\s+{AUTHOR_WORD}){{0,4}}"
AUTHOR_SUFFIX = r"(?:\s+và\s+cộng\s+sự)?"
YEAR_PATTERN = r"(?:19|20)\d{2}[a-z]?"
AUTHOR_REPORTING_VERB = (
    r"(?:cho\s+rằng|nhận\s+thấy|ghi\s+nhận|chỉ\s+ra|báo\s+cáo|"
    r"kết\s+luận|đề\s+xuất|mô\s+tả|phát\s+hiện|nghiên\s+cứu|"
    r"thực\s+hiện|công\s+bố|khảo\s+sát|đánh\s+giá|phân\s+tích)"
)


def _looks_like_person_author(author_text):
    normalized = unicodedata.normalize("NFD", author_text.upper())
    normalized = "".join(
        char for char in normalized if unicodedata.category(char) != "Mn"
    )
    words = re.findall(r"[A-Z]+", normalized)
    if not words:
        return False
    normalized_text = " ".join(words)
    if normalized_text in {"THEO", "WHO", "CDC", "UNICEF"}:
        return False
    institution_phrases = {
        "BAO CAO",
        "BO Y TE",
        "SO Y TE",
        "TRUONG DAI HOC",
        "DAI HOC",
        "BENH VIEN",
        "TO CHUC",
        "HIEP HOI",
    }
    if any(phrase in normalized_text for phrase in institution_phrases):
        return False
    return True


def _merge_ranges(ranges):
    if not ranges:
        return []
    merged = []
    for start, end in sorted(ranges):
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [tuple(item) for item in merged]


def _find_author_citations_missing_year(text):
    def has_year_in_same_clause(start, end):
        """Chấp nhận năm ở trước hoặc sau tác giả trong cùng câu/cụm."""
        left_boundary = max(
            text.rfind(mark, 0, start) for mark in (".", "!", "?", ";", "\n")
        )
        right_candidates = [
            position
            for mark in (".", "!", "?", ";", "\n")
            if (position := text.find(mark, end)) != -1
        ]
        right_boundary = min(right_candidates) if right_candidates else len(text)
        clause = text[left_boundary + 1 : right_boundary]
        return bool(re.search(YEAR_PATTERN, clause, re.IGNORECASE))

    ranges = []
    narrative_pattern = re.compile(
        rf"\b(?i:Theo|Nghiên cứu của|Báo cáo của|Tác giả|"
        rf"Công trình của)\s+(?P<author>{AUTHOR_NAME}{AUTHOR_SUFFIX})"
    )
    for match in narrative_pattern.finditer(text):
        if not _looks_like_person_author(match.group("author")):
            continue
        start, end = match.span("author")
        if not has_year_in_same_clause(start, end):
            ranges.append((start, end))

    parenthetical_pattern = re.compile(
        rf"\((?P<author>{AUTHOR_NAME}{AUTHOR_SUFFIX})\)"
    )
    for match in parenthetical_pattern.finditer(text):
        start, end = match.span("author")
        if (
            _looks_like_person_author(match.group("author"))
            and not has_year_in_same_clause(start, end)
        ):
            ranges.append(match.span("author"))

    reporting_pattern = re.compile(
        rf"(?<!Theo )(?<!THEO )\b"
        rf"(?P<author>{AUTHOR_NAME}{AUTHOR_SUFFIX})\s+"
        rf"(?i:{AUTHOR_REPORTING_VERB})"
    )
    for match in reporting_pattern.finditer(text):
        if not _looks_like_person_author(match.group("author")):
            continue
        start, end = match.span("author")
        if not has_year_in_same_clause(start, end):
            ranges.append((start, end))

    # Quét riêng mọi cụm "tên tác giả và cộng sự" để vẫn phát hiện được
    # trường hợp không đi kèm các động từ tường thuật ở trên.
    author_with_suffix_pattern = re.compile(
        rf"\b(?P<author>{AUTHOR_NAME}\s+(?i:và\s+cộng\s+sự))\b"
    )
    for match in author_with_suffix_pattern.finditer(text):
        start, end = match.span("author")
        if (
            _looks_like_person_author(
                re.sub(
                    r"\s+và\s+cộng\s+sự$",
                    "",
                    match.group("author"),
                    flags=re.IGNORECASE,
                )
            )
            and not has_year_in_same_clause(start, end)
        ):
            ranges.append((start, end))
    return _merge_ranges(ranges)


REFERENCE_NUMBER_PATTERN = re.compile(r"^\s*(?P<number>\d+[.)]?)\s*")
REFERENCE_YEAR_PATTERN = re.compile(
    rf"(?P<open>\()?\b(?P<year>{YEAR_PATTERN})\b(?P<close>\))?"
)
REFERENCE_SUFFIX_PATTERN = re.compile(
    r"(?i)(?:\s*,?\s*)(?:và\s+cộng\s+sự|và\s+cs\.?|et\s+al\.?)\s*$"
)
VIETNAMESE_SURNAMES = {
    "BÙI", "CAO", "ĐẶNG", "ĐINH", "ĐỖ", "DƯƠNG", "HOÀNG", "HUỲNH",
    "LÊ", "LÝ", "NGÔ", "NGUYỄN", "PHẠM", "PHAN", "TRẦN", "TRỊNH",
    "VÕ", "VŨ",
}


def _is_vietnamese_author_block(author_text):
    if re.search(r"[À-ỹĐđ]", author_text):
        return True
    first_words = {
        item.strip().split()[0].upper()
        for item in re.split(r"\s*,\s*|\s+và\s+", author_text)
        if item.strip()
    }
    return bool(first_words & VIETNAMESE_SURNAMES)


def _split_reference_authors(author_text):
    without_suffix = REFERENCE_SUFFIX_PATTERN.sub("", author_text).strip()
    authors = [
        item.strip(" ,;")
        for item in re.split(r"\s*,\s*|\s+và\s+", without_suffix)
        if item.strip(" ,;")
    ]
    return authors


def _normalize_reference_authors(author_text):
    """Chuẩn hóa khối tác giả theo quy định 3 tác giả đầu + cộng sự."""
    actions = []
    had_suffix = bool(REFERENCE_SUFFIX_PATTERN.search(author_text))
    is_vietnamese = _is_vietnamese_author_block(author_text)
    authors = _split_reference_authors(author_text)

    # Không tự sửa tên cơ quan/tổ chức vì dấu phẩy có thể không phân cách tác giả.
    if not authors or not all(_looks_like_person_author(item) for item in authors):
        return author_text.strip(" ,.;"), actions, is_vietnamese

    needs_suffix = len(authors) > 3 or had_suffix
    if len(authors) > 3:
        authors = authors[:3]
        actions.append("reference_trimmed_authors")

    if needs_suffix:
        if is_vietnamese:
            normalized = ", ".join(authors) + " và cộng sự"
        else:
            normalized = ", ".join(authors) + ", et al."
        if not had_suffix:
            actions.append("reference_added_et_al")
    else:
        normalized = ", ".join(authors)

    author_warning = False
    if is_vietnamese:
        # Tên Việt Nam phải ghi đầy đủ; các chữ cái đơn/viết tắt là dấu hiệu sai.
        author_warning = any(
            len(re.findall(r"[A-Za-zÀ-ỹĐđ]+", item)) < 2
            or bool(re.search(r"(?<!\w)[A-ZĐ]\.(?:[A-ZĐ]\.)?", item))
            for item in authors
        )
    else:
        # Tên nước ngoài: họ đầy đủ, tên gọi và tên đệm dùng chữ cái viết tắt.
        author_warning = any(
            len(item.split()) >= 2
            and not any(
                re.fullmatch(r"[A-Z](?:\.)?", token)
                or re.fullmatch(r"(?:[A-Z]\.?){1,4}", token)
                for token in item.replace(",", " ").split()[1:]
            )
            for item in authors
        )
    if author_warning:
        actions.append("reference_author_name_warning")

    return normalized, actions, is_vietnamese


def _set_reference_character_style(paragraph, ranges):
    """Áp dụng in đứng/in nghiêng/in đậm theo các khoảng ký tự."""
    records = _paragraph_char_records(paragraph)
    if not records:
        return False
    changed = False
    for start, end, bold, italic in ranges:
        for index in range(max(0, start), min(end, len(records))):
            char, style = records[index]
            values = list(style)
            if values[0] != bold or values[1] != italic:
                values[0] = bold
                values[1] = italic
                records[index] = (char, tuple(values))
                changed = True
    if changed:
        _rebuild_paragraph(paragraph, records)
    return changed


def _format_reference_title_journal_volume(paragraph):
    """Tên bài in đứng, tên tạp chí in nghiêng, tập/số in đậm."""
    text = paragraph.text
    year_match = REFERENCE_YEAR_PATTERN.search(text)
    page_match = re.search(
        r"(?P<pages>\d{1,5}(?:\s*[-–—]\s*\d{1,5})?)\s*\.?\s*$",
        text,
    )
    if not year_match or not page_match:
        return False

    volume_match = re.search(
        r"(?P<volume>\d+(?:\s*\([^)]*\))?)\s*,\s*"
        r"(?P<pages>\d{1,5}(?:\s*[-–—]\s*\d{1,5})?)\s*\.?\s*$",
        text[year_match.end() :],
    )
    if not volume_match:
        return False
    volume_start = year_match.end() + volume_match.start("volume")
    volume_end = year_match.end() + volume_match.end("volume")

    content_start = year_match.end()
    while content_start < len(text) and text[content_start] in " .,:;\t":
        content_start += 1
    title_separator = re.search(r"\.\s+", text[content_start:volume_start])
    if not title_separator:
        return False
    title_end = content_start + title_separator.start()
    journal_start = content_start + title_separator.end()
    journal_end = volume_start
    while journal_end > journal_start and text[journal_end - 1] in " ,.;\t":
        journal_end -= 1

    return _set_reference_character_style(
        paragraph,
        [
            (content_start, title_end, False, False),
            (journal_start, journal_end, False, True),
            (volume_start, volume_end, True, False),
        ],
    )


def normalize_reference_entry(paragraph):
    """Chuẩn hóa một mục tài liệu tham khảo và trả về danh sách thao tác."""
    original_text = paragraph.text
    number_match = REFERENCE_NUMBER_PATTERN.match(original_text)
    if not number_match or not original_text[number_match.end() :].strip():
        return []

    actions = []
    year_match = REFERENCE_YEAR_PATTERN.search(original_text, number_match.end())
    if not year_match:
        highlight_paragraph_ranges(
            paragraph,
            [(number_match.end(), len(original_text))],
            highlight_color=WD_COLOR_INDEX.RED,
            preserve_existing=False,
        )
        return ["reference_missing_year"]

    number_text = number_match.group("number").rstrip(".)") + "."
    author_text = original_text[number_match.end() : year_match.start()].strip(
        " ,.;\t"
    )
    normalized_authors, author_actions, _ = _normalize_reference_authors(
        author_text
    )
    actions.extend(author_actions)

    year = year_match.group("year")
    if not (year_match.group("open") and year_match.group("close")):
        actions.append("reference_parenthesized_year")

    remainder = original_text[year_match.end() :].lstrip(" .,:;\t")
    remainder, page_prefix_count = re.subn(
        r"(?i)\btr\s*\.?\s*(?=\d{1,5}(?:\s*[-–—]\s*\d{1,5})?)",
        "",
        remainder,
    )
    actions.extend(["reference_removed_tr"] * page_prefix_count)
    remainder = re.sub(r"\s*[-–—]\s*", "-", remainder)
    remainder = re.sub(r"\s+([,.;:])", r"\1", remainder)
    remainder = re.sub(r",\s*,", ",", remainder)
    remainder = re.sub(r" {2,}", " ", remainder).strip()

    normalized_text = (
        f"{number_text} {normalized_authors}. ({year}). {remainder}"
    ).strip()
    normalized_text = re.sub(r"\.\s*\.\s*\(", ". (", normalized_text)
    normalized_text = re.sub(r" {2,}", " ", normalized_text)
    if normalized_text != original_text:
        set_paragraph_text_preserve_formatting(
            paragraph,
            normalized_text,
            highlight_changes=True,
        )

    if "reference_author_name_warning" in actions:
        current_year = REFERENCE_YEAR_PATTERN.search(paragraph.text)
        if current_year:
            current_number = REFERENCE_NUMBER_PATTERN.match(paragraph.text)
            highlight_paragraph_ranges(
                paragraph,
                [(current_number.end(), current_year.start())],
                highlight_color=WD_COLOR_INDEX.YELLOW,
                preserve_existing=True,
            )

    if _format_reference_title_journal_volume(paragraph):
        actions.append("reference_styled")
    else:
        actions.append("reference_structure_warning")
        highlight_paragraph_ranges(
            paragraph,
            [(0, len(paragraph.text))],
            highlight_color=WD_COLOR_INDEX.YELLOW,
            preserve_existing=True,
        )
    return actions


def fix_author_citations(paragraph, in_references_section=False):
    text = paragraph.text
    if not text:
        return []

    original_text = text
    working_text = text
    actions = []

    working_text, doi_count = re.subn(
        r"(?i)\s*\bdoi:\s*\S+",
        "",
        working_text,
    )
    actions.extend(["removed_doi"] * doi_count)

    if in_references_section:
        if working_text != original_text:
            set_paragraph_text_preserve_formatting(
                paragraph,
                working_text,
                highlight_changes=False,
            )
        actions.extend(normalize_reference_entry(paragraph))
        return actions

    working_text, et_al_count = re.subn(
        r"(?i)\bet\s+al\.?|\bvà\s+cs\.?",
        "và cộng sự",
        working_text,
    )
    actions.extend(["body_normalized_et_al"] * et_al_count)

    narrative_with_year = re.compile(
        rf"\b(?P<prefix>(?i:Theo|Nghiên cứu của|Báo cáo của|Tác giả|"
        rf"Công trình của))\s+"
        rf"(?P<author>{AUTHOR_NAME})\s*"
        rf"\((?P<year>{YEAR_PATTERN})\)"
    )

    def replace_narrative(match):
        author = match.group("author").strip()
        if not _looks_like_person_author(author):
            return match.group(0)
        if re.search(r"và\s+cộng\s+sự$", author, re.IGNORECASE):
            return match.group(0)
        actions.append("body_added_et_al")
        return (
            f"{match.group('prefix')} {author} và cộng sự "
            f"({match.group('year')})"
        )

    working_text = narrative_with_year.sub(replace_narrative, working_text)

    parenthetical_with_year = re.compile(
        rf"\((?P<author>{AUTHOR_NAME})\s*,\s*"
        rf"(?P<year>{YEAR_PATTERN})\)"
    )

    def replace_parenthetical(match):
        author = match.group("author").strip()
        if not _looks_like_person_author(author):
            return match.group(0)
        if re.search(r"và\s+cộng\s+sự$", author, re.IGNORECASE):
            return match.group(0)
        actions.append("body_added_et_al")
        return f"({author} và cộng sự, {match.group('year')})"

    working_text = parenthetical_with_year.sub(
        replace_parenthetical,
        working_text,
    )

    # Năm đứng trước tác giả, ví dụ:
    # "Năm 2024, Nguyễn Văn A ghi nhận...".
    year_before_author = re.compile(
        rf"\b(?P<prefix>(?i:(?:Vào\s+|Trong\s+)?Năm))\s+"
        rf"(?P<year>{YEAR_PATTERN})(?P<separator>\s*[,;:-]\s*)"
        rf"(?P<author>{AUTHOR_NAME})"
        rf"(?!\s+và\s+cộng\s+sự)"
        rf"(?=\s+(?i:{AUTHOR_REPORTING_VERB})\b)"
    )

    def replace_year_before_author(match):
        author = match.group("author").strip()
        if not _looks_like_person_author(author):
            return match.group(0)
        actions.append("body_added_et_al")
        return (
            f"{match.group('prefix')} {match.group('year')}"
            f"{match.group('separator')}{author} và cộng sự"
        )

    working_text = year_before_author.sub(
        replace_year_before_author,
        working_text,
    )

    standalone_with_year = re.compile(
        rf"\b(?P<author>{AUTHOR_NAME})"
        rf"(?!\s+và\s+cộng\s+sự)\s*"
        rf"\((?P<year>{YEAR_PATTERN})\)"
    )

    def replace_standalone(match):
        if not _looks_like_person_author(match.group("author")):
            return match.group(0)
        actions.append("body_added_et_al")
        return (
            f"{match.group('author')} và cộng sự "
            f"({match.group('year')})"
        )

    working_text = standalone_with_year.sub(
        replace_standalone,
        working_text,
    )

    missing_year_ranges = _find_author_citations_missing_year(working_text)
    actions.extend(["author_missing_year"] * len(missing_year_ranges))
    missing_year_yellow_ranges = [
        (start, end, WD_COLOR_INDEX.YELLOW)
        for start, end in missing_year_ranges
    ]

    if working_text != original_text or missing_year_ranges:
        set_paragraph_text_preserve_formatting(
            paragraph,
            working_text,
            highlight_changes=working_text != original_text,
            highlight_ranges=missing_year_yellow_ranges,
        )
    return actions


def _normalize_heading(text):
    normalized = unicodedata.normalize("NFD", text.upper())
    normalized = "".join(
        char for char in normalized if unicodedata.category(char) != "Mn"
    )
    # Unicode NFD không tách chữ Đ thành D, phải thay riêng để các mục
    # "ĐẶT VẤN ĐỀ" và "ĐỐI TƯỢNG..." được nhận diện đúng.
    normalized = normalized.replace("Đ", "D")
    normalized = re.sub(
        r"^\s*(?:CHUONG\s+)?(?:[IVXLCDM]+|\d+)[.:\s-]+",
        "",
        normalized,
    )
    normalized = re.sub(r"\s+", " ", normalized).strip(" .:-\t")
    return normalized


def is_references_heading(text):
    heading = _normalize_heading(text)
    return heading in {
        "TAI LIEU THAM KHAO",
        "REFERENCES",
        "BIBLIOGRAPHY",
    }


def is_heading_after_references(text):
    heading = _normalize_heading(text)
    return (
        heading.startswith("PHU LUC")
        or heading.startswith("APPENDIX")
        or heading.startswith("DANH MUC BAI BAO")
    )


def has_image(paragraph):
    """Chỉ nhận diện ảnh nhúng thật, không coi shape/đường kẻ là hình ảnh.

    ``w:drawing`` và ``w:pict`` còn được Word dùng cho đường kẻ, khung,
    AutoShape (ví dụ đường chéo trong ô bảng). Ảnh thật phải có phần tử chứa
    dữ liệu ảnh: DrawingML ``a:blip``/``pic:pic`` hoặc VML ``v:imagedata``.
    """
    xml = paragraph._element.xml
    return (
        "a:blip" in xml
        or "pic:pic" in xml
        or "v:imagedata" in xml
    )


def _all_document_paragraphs(doc):
    """Lấy cả paragraph thường, paragraph trong bảng và trong text box."""
    for element in doc.element.body.iter():
        if isinstance(element, CT_P):
            yield Paragraph(element, doc)


def _paragraph_has_superscript_citation(paragraph):
    for run in paragraph.runs:
        if not bool(run.font.superscript):
            continue
        if re.fullmatch(r"\s*\d+(?:\s*[,;–—-]\s*\d+)*\s*", run.text):
            return True
    return False


def _has_author_year_citation(text):
    for match in re.finditer(r"\(([^)]*)\)", text):
        content = match.group(1)
        if (
            re.search(r"(?:19|20)\d{2}[a-z]?", content)
            and re.search(r"[A-Za-zÀ-ỹ]", content)
        ):
            return True
    return False


def _paragraph_has_explicit_image_source(paragraph):
    text = paragraph.text.strip()
    if not text:
        return _paragraph_has_superscript_citation(paragraph)

    explicit_source_patterns = [
        r"(?i)\bnguồn\b",
        r"(?i)\bsource\b",
        r"(?i)\b(?:trích|tham khảo)\s+(?:từ|theo)\b",
        r"(?i)\b(?:tác giả|nhóm nghiên cứu)\s+tự\s+"
        r"(?:chụp|vẽ|xây dựng|tổng hợp|thiết kế)",
        r"(?i)https?://|www\.",
        r"(?i)\b[a-z0-9.-]+\.(?:vn|com|org|edu|gov)\b",
        r"\[\s*\d+(?:\s*[,;–—-]\s*\d+)*\s*\]",
    ]
    if any(re.search(pattern, text) for pattern in explicit_source_patterns):
        return True
    if _has_author_year_citation(text):
        return True
    return _paragraph_has_superscript_citation(paragraph)


def _paragraph_is_image_caption(paragraph):
    return bool(
        re.search(
            r"(?i)^\s*(?:hình|sơ đồ|biểu đồ|đồ thị|bản đồ|ảnh|"
            r"figure|fig\.?|chart)\s*(?:\d+|[:.-])",
            paragraph.text,
        )
    )


def _insert_image_warning(doc, image_paragraph, warning_text, font_target):
    warning_paragraph = doc.add_paragraph()
    image_paragraph._p.addnext(warning_paragraph._p)
    warning_run = warning_paragraph.add_run(warning_text)
    warning_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    warning_run.bold = True
    if font_target:
        warning_run.font.name = font_target
    return warning_paragraph


COVER_END_STRUCTURE_KEYS = {
    "LOI CAM ON",
    "LOI CAM DOAN",
    "MUC LUC",
    "DANH MUC CHU VIET TAT",
    "DANH MUC BANG",
    "DANH MUC BIEU DO",
    "DANH MUC HINH",
    "DANH MUC SO DO",
    "DANH MUC HINH SO DO BIEU DO",
    "DAT VAN DE",
}

IMAGE_WARNING_PARAGRAPH_PATTERN = re.compile(
    r"^\s*(?:⚠️\s*)?\[CẢNH BÁO NGUỒN HÌNH\]\s*:",
    re.IGNORECASE,
)


def _remove_previous_image_warnings(doc):
    """Xóa cảnh báo nguồn hình của lần quét trước để tránh cảnh báo rác."""
    removed_count = 0
    for paragraph in list(_all_document_paragraphs(doc)):
        if not IMAGE_WARNING_PARAGRAPH_PATTERN.search(paragraph.text):
            continue
        parent = paragraph._element.getparent()
        if parent is not None:
            parent.remove(paragraph._element)
            removed_count += 1
    return removed_count


def _find_cover_content_boundary(paragraphs):
    """Trả về vị trí bắt đầu phần nội dung sau bìa chính và bìa phụ.

    Word thường không lưu ngắt trang tự động trong XML, nên không thể chỉ
    đếm ``w:br`` để xác định hai trang bìa. Các ảnh nằm trước tiêu đề đầu
    tiên của phần đầu luận văn được coi là logo/thành phần trang bìa và
    không phải ghi nguồn.
    """
    for index, paragraph in enumerate(paragraphs):
        structure_key = _canonical_structure_key(paragraph.text)
        if structure_key in COVER_END_STRUCTURE_KEYS:
            return index
    return None


def check_image_citations(
    doc,
    cover_p_elements,
    detailed_errors,
    font_target,
    require_source=True,
    require_caption=True,
):
    """
    Quét hình trong thân bài, bảng và text box.

    Chỉ coi là đã có nguồn khi vùng chú thích có "Nguồn/Source", URL,
    trích dẫn số, trích dẫn tác giả-năm hoặc ghi rõ tác giả tự thực hiện.
    Các từ chung chung như "theo", "tác giả", "WHO" không còn tự động
    làm mất cảnh báo.
    """
    removed_warning_count = _remove_previous_image_warnings(doc)
    paragraphs = list(_all_document_paragraphs(doc))
    cover_content_boundary = _find_cover_content_boundary(paragraphs)
    missing_source_count = 0
    missing_caption_count = 0
    skipped_cover_image_count = 0

    for index, paragraph in enumerate(paragraphs):
        if not has_image(paragraph):
            continue
        if (
            cover_content_boundary is not None
            and index < cover_content_boundary
        ):
            skipped_cover_image_count += 1
            continue
        if is_cover_or_logo_image(
            index,
            paragraph,
            cover_p_elements,
            paragraphs,
        ):
            continue

        start_index = max(0, index - 1)
        end_index = min(len(paragraphs), index + 4)
        nearby_paragraphs = paragraphs[start_index:end_index]
        source_paragraphs = paragraphs[index:end_index]
        if (
            index > 0
            and _paragraph_is_image_caption(paragraphs[index - 1])
        ):
            source_paragraphs = [
                paragraphs[index - 1],
                *source_paragraphs,
            ]

        has_caption = any(
            _paragraph_is_image_caption(item)
            for item in nearby_paragraphs
        )
        has_source = any(
            _paragraph_has_explicit_image_source(item)
            for item in source_paragraphs
        )

        caption_ok = has_caption or not require_caption
        source_ok = has_source or not require_source
        if caption_ok and source_ok:
            continue

        already_warned = any(
            "CẢNH BÁO NGUỒN HÌNH" in item.text.upper()
            for item in nearby_paragraphs
        )
        missing_items = []
        if require_source and not has_source:
            missing_source_count += 1
            missing_items.append("chưa có trích dẫn nguồn")
        if require_caption and not has_caption:
            missing_caption_count += 1
            missing_items.append("chưa có tên/chú thích hình")

        error_message = " và ".join(missing_items)
        detailed_errors.append(
            f"🖼️ **Hình ảnh {index + 1}:** {error_message}. "
            "Đã đánh dấu để học viên bổ sung."
        )

        if not already_warned:
            _insert_image_warning(
                doc,
                paragraph,
                "⚠️ [CẢNH BÁO NGUỒN HÌNH]: "
                f"Hình này {error_message}. Vui lòng bổ sung!",
                font_target,
            )

    if missing_source_count:
        detailed_errors.append(
            "🟨 **Kiểm tra nguồn hình:** Phát hiện "
            f"{missing_source_count} hình thiếu trích dẫn nguồn; "
            "đã chèn cảnh báo bôi vàng ngay dưới hình."
        )
    if missing_caption_count:
        detailed_errors.append(
            "🟨 **Kiểm tra chú thích hình:** Phát hiện "
            f"{missing_caption_count} hình thiếu tên/chú thích."
        )
    if skipped_cover_image_count:
        detailed_errors.append(
            "🏫 **Logo/trang bìa:** Đã bỏ qua "
            f"{skipped_cover_image_count} ảnh trên bìa chính/bìa phụ; "
            "không yêu cầu trích dẫn nguồn."
        )
    if removed_warning_count:
        detailed_errors.append(
            "🧹 **Làm sạch cảnh báo nguồn hình:** Đã xóa "
            f"{removed_warning_count} cảnh báo cũ và kiểm tra lại từ đầu."
        )
    return {
        "missing_source": missing_source_count,
        "missing_caption": missing_caption_count,
        "skipped_cover_images": skipped_cover_image_count,
        "removed_previous_warnings": removed_warning_count,
    }


FALLBACK_STRUCTURE = [
    "LOI CAM ON",
    "LOI CAM DOAN",
    "MUC LUC",
    "DANH MUC CHU VIET TAT",
    "DANH MUC BANG",
    "DANH MUC HINH SO DO BIEU DO",
    "DAT VAN DE",
    "TONG QUAN",
    "DOI TUONG VA PHUONG PHAP NGHIEN CUU",
    "KET QUA",
    "BAN LUAN",
    "KET LUAN",
    "KIEN NGHI",
    "TAI LIEU THAM KHAO",
]

STRUCTURE_LABELS = {
    "LOI CAM ON": "Lời cảm ơn",
    "LOI CAM DOAN": "Lời cam đoan",
    "MUC LUC": "Mục lục",
    "DANH MUC CHU VIET TAT": "Danh mục chữ viết tắt",
    "DANH MUC BANG": "Danh mục bảng",
    "DANH MUC BIEU DO": "Danh mục biểu đồ",
    "DANH MUC HINH": "Danh mục hình",
    "DANH MUC SO DO": "Danh mục sơ đồ",
    "DANH MUC HINH SO DO BIEU DO": (
        "Danh mục hình, sơ đồ, biểu đồ"
    ),
    "DAT VAN DE": "Đặt vấn đề",
    "MUC TIEU NGHIEN CUU": "Mục tiêu nghiên cứu",
    "CAU HOI NGHIEN CUU": "Câu hỏi nghiên cứu",
    "GIA THUYET NGHIEN CUU": "Giả thuyết nghiên cứu",
    "TONG QUAN": "Tổng quan",
    "DOI TUONG VA PHUONG PHAP NGHIEN CUU": (
        "Đối tượng và phương pháp nghiên cứu"
    ),
    "KET QUA": "Kết quả",
    "BAN LUAN": "Bàn luận",
    "DU KIEN KET QUA": "Dự kiến kết quả",
    "DU KIEN BAN LUAN": "Dự kiến bàn luận",
    "DU DOAN KET QUA VA BAN LUAN": (
        "Dự đoán kết quả và bàn luận"
    ),
    "DU KIEN KET QUA VA BAN LUAN": (
        "Dự kiến kết quả và bàn luận"
    ),
    "KE HOACH NGHIEN CUU": "Kế hoạch nghiên cứu",
    "TIEN DO THUC HIEN": "Tiến độ thực hiện",
    "DU TRU KINH PHI": "Dự trù kinh phí",
    "SAN PHAM DU KIEN": "Sản phẩm dự kiến",
    "KET LUAN": "Kết luận",
    "KIEN NGHI": "Kiến nghị",
    "TAI LIEU THAM KHAO": "Tài liệu tham khảo",
    "PHU LUC": "Phụ lục",
}

# Template ghi rõ Kiến nghị không bắt buộc. Nếu có thì vẫn kiểm tra vị trí,
# nhưng không báo thiếu khi luận văn không sử dụng phần này.
OPTIONAL_STRUCTURE_KEYS = {"KIEN NGHI", "PHU LUC"}


def _canonical_structure_key(text):
    normalized = _normalize_heading(text)
    proposal_combined_aliases = {
        "DU DOAN KET QUA VA BAN LUAN": (
            "DU DOAN KET QUA VA BAN LUAN"
        ),
        "DU KIEN KET QUA VA BAN LUAN": (
            "DU KIEN KET QUA VA BAN LUAN"
        ),
        "DU KIEN KET QUA NGHIEN CUU VA BAN LUAN": (
            "DU KIEN KET QUA VA BAN LUAN"
        ),
    }
    if normalized in proposal_combined_aliases:
        return proposal_combined_aliases[normalized]
    if re.search(
        r"^(?:KET QUA\s+VA\s+BAN LUAN|"
        r"BAN LUAN\s+VA\s+KET QUA)$",
        normalized,
    ):
        return "MERGED_RESULTS_DISCUSSION"
    if (
        "DOI TUONG" in normalized
        and "PHUONG PHAP" in normalized
        and "NGHIEN CUU" in normalized
    ):
        return "DOI TUONG VA PHUONG PHAP NGHIEN CUU"

    exact_aliases = {
        "LOI CAM ON": "LOI CAM ON",
        "LOI CAM DOAN": "LOI CAM DOAN",
        "MUC LUC": "MUC LUC",
        "DANH MUC CAC CHU VIET TAT": "DANH MUC CHU VIET TAT",
        "DANH MUC CHU VIET TAT": "DANH MUC CHU VIET TAT",
        "DANH MUC BANG": "DANH MUC BANG",
        "DANH MUC CAC BANG": "DANH MUC BANG",
        "DANH MUC BIEU DO": "DANH MUC BIEU DO",
        "DANH MUC CAC BIEU DO": "DANH MUC BIEU DO",
        "DANH MUC HINH": "DANH MUC HINH",
        "DANH MUC CAC HINH": "DANH MUC HINH",
        "DANH MUC HINH ANH": "DANH MUC HINH",
        "DANH MUC SO DO": "DANH MUC SO DO",
        "DANH MUC CAC SO DO": "DANH MUC SO DO",
        "DANH MUC HINH SO DO BIEU DO": (
            "DANH MUC HINH SO DO BIEU DO"
        ),
        "DANH MUC HINH, SO DO, BIEU DO": (
            "DANH MUC HINH SO DO BIEU DO"
        ),
        "DAT VAN DE": "DAT VAN DE",
        "INTRODUCTION": "DAT VAN DE",
        "MUC TIEU": "MUC TIEU NGHIEN CUU",
        "MUC TIEU NGHIEN CUU": "MUC TIEU NGHIEN CUU",
        "OBJECTIVES": "MUC TIEU NGHIEN CUU",
        "CAU HOI NGHIEN CUU": "CAU HOI NGHIEN CUU",
        "RESEARCH QUESTIONS": "CAU HOI NGHIEN CUU",
        "GIA THUYET NGHIEN CUU": "GIA THUYET NGHIEN CUU",
        "RESEARCH HYPOTHESES": "GIA THUYET NGHIEN CUU",
        "TONG QUAN": "TONG QUAN",
        "TONG QUAN TAI LIEU": "TONG QUAN",
        "DOI TUONG VA PHUONG PHAP NGHIEN CUU": (
            "DOI TUONG VA PHUONG PHAP NGHIEN CUU"
        ),
        "DOI TUONG PHUONG PHAP NGHIEN CUU": (
            "DOI TUONG VA PHUONG PHAP NGHIEN CUU"
        ),
        "PHUONG PHAP NGHIEN CUU": (
            "DOI TUONG VA PHUONG PHAP NGHIEN CUU"
        ),
        "KET QUA": "KET QUA",
        "KET QUA NGHIEN CUU": "KET QUA",
        "CAC KET QUA NGHIEN CUU": "KET QUA",
        "BAN LUAN": "BAN LUAN",
        "BAN LUAN KET QUA": "BAN LUAN",
        "BAN LUAN KET QUA NGHIEN CUU": "BAN LUAN",
        "DISCUSSION": "BAN LUAN",
        "DU KIEN KET QUA": "DU KIEN KET QUA",
        "KET QUA DU KIEN": "DU KIEN KET QUA",
        "EXPECTED RESULTS": "DU KIEN KET QUA",
        "DU KIEN BAN LUAN": "DU KIEN BAN LUAN",
        "KE HOACH NGHIEN CUU": "KE HOACH NGHIEN CUU",
        "KE HOACH THUC HIEN": "KE HOACH NGHIEN CUU",
        "TIEN DO THUC HIEN": "TIEN DO THUC HIEN",
        "TIEN DO NGHIEN CUU": "TIEN DO THUC HIEN",
        "DU TRU KINH PHI": "DU TRU KINH PHI",
        "KINH PHI NGHIEN CUU": "DU TRU KINH PHI",
        "SAN PHAM DU KIEN": "SAN PHAM DU KIEN",
        "KET LUAN": "KET LUAN",
        "KIEN NGHI": "KIEN NGHI",
        "TAI LIEU THAM KHAO": "TAI LIEU THAM KHAO",
        "REFERENCES": "TAI LIEU THAM KHAO",
        "PHU LUC": "PHU LUC",
        "APPENDIX": "PHU LUC",
    }
    return exact_aliases.get(normalized)


def _paragraph_style_name(paragraph):
    try:
        return paragraph.style.name if paragraph.style else ""
    except Exception:
        return ""


def _is_toc_paragraph(paragraph):
    normalized_style = _normalize_heading(_paragraph_style_name(paragraph))
    if normalized_style.startswith(
        ("TOC", "MUC LUC", "CONTENTS", "TABLE OF CONTENTS")
    ):
        return True

    # Một số file Word làm mất style TOC nhưng vẫn giữ liên kết của mục
    # lục tới bookmark _Toc... Không được nhầm liên kết này với tiêu đề
    # "TÀI LIỆU THAM KHẢO" thật ở cuối luận văn.
    xml = paragraph._element.xml
    return bool(
        ("w:hyperlink" in xml and 'w:anchor="_Toc' in xml)
        or ("PAGEREF" in xml and "_Toc" in xml)
    )


def _is_actual_references_heading(paragraph):
    """Chỉ nhận tiêu đề tài liệu tham khảo thật, bỏ dòng cùng tên ở mục lục."""
    return (
        is_references_heading(paragraph.text)
        and not _is_toc_paragraph(paragraph)
    )


def _custom_structure_key(text):
    normalized = _normalize_heading(text)
    return f"CUSTOM::{normalized}" if normalized else None


def _structure_label(key):
    if key.startswith("CUSTOM::"):
        return key.split("::", 1)[1].title()
    return STRUCTURE_LABELS.get(key, key)


def _is_template_top_level_heading(paragraph):
    """Nhận tiêu đề cấp 1 riêng có trong template của người dùng.

    Các mục chuẩn vẫn nhận theo từ khóa. Với mục đặc thù chưa có trong danh
    sách (ví dụ một mục riêng của đề cương), chỉ lấy Heading 1 để tránh đưa
    toàn bộ tiểu mục 1.1, 1.2 vào kiểm tra cấu trúc lớn.
    """
    text = paragraph.text.strip()
    normalized_text = _normalize_heading(text)
    if (
        not normalized_text
        or normalized_text.startswith(("NGUON", "SOURCE"))
        or _paragraph_is_image_caption(paragraph)
        or re.fullmatch(r"[.\-–—_…\s]+", text)
    ):
        return False
    style = _normalize_heading(_paragraph_style_name(paragraph))
    return bool(
        re.fullmatch(r"(?:HEADING|TIEU DE)\s*1", style)
        or style in {"CHAPTER TITLE", "TEN CHUONG"}
    )


def _extract_structure_headings(doc, expected_keys=None):
    headings = []
    expected_custom = {
        key for key in (expected_keys or []) if key.startswith("CUSTOM::")
    }
    for paragraph in list(_all_document_paragraphs(doc)):
        # Không lấy các dòng trong Mục lục làm chương thật. Số trang của
        # trường TOC đôi khi không xuất hiện trong paragraph.text, khiến
        # "TÀI LIỆU THAM KHẢO" ở mục lục bị nhận nhầm là phần cuối luận văn.
        if _is_toc_paragraph(paragraph):
            continue
        key = _canonical_structure_key(paragraph.text)
        if key is None and expected_custom:
            candidate = _custom_structure_key(paragraph.text)
            if candidate in expected_custom:
                key = candidate
        if key:
            headings.append((key, paragraph))
    return headings


def _extract_template_structure_headings(template_doc):
    headings = []
    seen = set()
    for paragraph in list(_all_document_paragraphs(template_doc)):
        if _is_toc_paragraph(paragraph):
            continue
        key = _canonical_structure_key(paragraph.text)
        if key is None and _is_template_top_level_heading(paragraph):
            key = _custom_structure_key(paragraph.text)
        if not key or key == "MERGED_RESULTS_DISCUSSION" or key in seen:
            continue
        seen.add(key)
        headings.append((key, paragraph))
    return headings


STRUCTURE_WARNING_PARAGRAPH_PATTERN = re.compile(
    r"^\s*(?:⚠️\s*)?\[CẢNH BÁO CẤU TRÚC\]\s*:",
    re.IGNORECASE,
)


def _remove_previous_structure_warnings(doc):
    """Xóa cảnh báo cấu trúc do lần chạy trước tạo ra trước khi kiểm tra lại."""
    removed_count = 0
    for paragraph in list(_all_document_paragraphs(doc)):
        if not STRUCTURE_WARNING_PARAGRAPH_PATTERN.search(paragraph.text):
            continue
        parent = paragraph._element.getparent()
        if parent is not None:
            parent.remove(paragraph._element)
            removed_count += 1
    return removed_count


def _clear_old_structure_highlight(paragraph):
    for run in paragraph.runs:
        if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
            run.font.highlight_color = None


def _template_structure(template_path=None, template_bytes=None):
    template_requested = bool(template_bytes or template_path)
    try:
        if template_bytes:
            template_doc = docx.Document(io.BytesIO(template_bytes))
        elif template_path and os.path.exists(template_path):
            template_doc = docx.Document(template_path)
        else:
            template_doc = None

        if template_doc is not None:
            keys = []
            for key, _ in _extract_template_structure_headings(template_doc):
                if key not in keys:
                    keys.append(key)
            if len(keys) >= 2:
                return keys
            raise ValueError(
                "Template không có đủ tiêu đề cấp 1 để đối chiếu cấu trúc. "
                "Hãy đặt Style Heading 1 cho các mục/chương chính."
            )
    except ValueError:
        raise
    except Exception as exc:
        if template_requested:
            raise ValueError(f"Không thể đọc cấu trúc template: {exc}") from exc
    return list(FALLBACK_STRUCTURE)


def _highlight_structure_paragraph(paragraph):
    if paragraph.text:
        highlight_paragraph_ranges(
            paragraph,
            [(0, len(paragraph.text))],
            highlight_color=WD_COLOR_INDEX.YELLOW,
            preserve_existing=True,
        )


def check_structure_against_template(
    doc,
    template_path,
    detailed_errors,
    font_target,
    template_bytes=None,
):
    removed_warning_count = _remove_previous_structure_warnings(doc)
    expected_keys = _template_structure(
        template_path=template_path,
        template_bytes=template_bytes,
    )
    expected_positions = {
        key: index for index, key in enumerate(expected_keys)
    }
    actual_headings = _extract_structure_headings(
        doc,
        expected_keys=expected_keys,
    )
    if removed_warning_count:
        # Các tiêu đề từng bị báo sai đã được bôi vàng ở lần chạy trước.
        # Xóa màu cũ trước, sau đó chỉ bôi lại nếu lần kiểm tra mới vẫn sai.
        for _, paragraph in actual_headings:
            _clear_old_structure_highlight(paragraph)
        detailed_errors.append(
            "🧹 **Cảnh báo cấu trúc cũ:** Đã xóa "
            f"{removed_warning_count} cảnh báo của lần kiểm tra trước "
            "và đánh giá lại từ đầu."
        )
    existing_warning_text = " ".join(
        paragraph.text
        for paragraph in _all_document_paragraphs(doc)
        if "CẢNH BÁO CẤU TRÚC" in paragraph.text.upper()
    )

    merged_count = 0
    out_of_order_count = 0
    duplicate_count = 0
    actual_keys = []
    first_paragraph = actual_headings[0][1] if actual_headings else None

    for key, paragraph in actual_headings:
        if key == "MERGED_RESULTS_DISCUSSION":
            merged_count += 1
            _highlight_structure_paragraph(paragraph)
            warning = (
                "Phải tách riêng Chương Kết quả và Chương Bàn luận; "
                "không gộp thành một chương."
            )
            if warning not in existing_warning_text:
                _insert_image_warning(
                    doc,
                    paragraph,
                    "⚠️ [CẢNH BÁO CẤU TRÚC]: " + warning,
                    font_target,
                )
            continue
        actual_keys.append(key)

    seen_keys = set()
    last_position = -1
    for key, paragraph in actual_headings:
        if key == "MERGED_RESULTS_DISCUSSION":
            continue
        if key in seen_keys:
            duplicate_count += 1
            _highlight_structure_paragraph(paragraph)
            continue
        seen_keys.add(key)

        position = expected_positions.get(key)
        if position is None:
            continue
        if position < last_position:
            out_of_order_count += 1
            _highlight_structure_paragraph(paragraph)
            warning = (
                f"Mục {_structure_label(key)} đang sai "
                "thứ tự so với file template."
            )
            if warning not in existing_warning_text:
                _insert_image_warning(
                    doc,
                    paragraph,
                    "⚠️ [CẢNH BÁO CẤU TRÚC]: " + warning,
                    font_target,
                )
        else:
            last_position = position

    missing_keys = [
        key for key in expected_keys
        if key not in seen_keys and key not in OPTIONAL_STRUCTURE_KEYS
    ]
    if merged_count:
        missing_keys = [
            key for key in missing_keys
            if key not in {"KET QUA", "BAN LUAN"}
        ]

    if missing_keys:
        missing_labels = ", ".join(
            _structure_label(key) for key in missing_keys
        )
        warning = (
            "Thiếu các chương/mục theo template: "
            f"{missing_labels}."
        )
        if warning not in existing_warning_text:
            if first_paragraph is not None:
                _insert_image_warning(
                    doc,
                    first_paragraph,
                    "⚠️ [CẢNH BÁO CẤU TRÚC]: " + warning,
                    font_target,
                )
            else:
                warning_paragraph = doc.add_paragraph()
                warning_run = warning_paragraph.add_run(
                    "⚠️ [CẢNH BÁO CẤU TRÚC]: " + warning
                )
                warning_run.bold = True
                warning_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                if font_target:
                    warning_run.font.name = font_target

    if merged_count:
        detailed_errors.append(
            "🟨 **Cấu trúc chương:** Phát hiện "
            f"{merged_count} tiêu đề gộp Kết quả và Bàn luận; "
            "đã bôi vàng và yêu cầu tách thành hai chương riêng."
        )
    if out_of_order_count:
        detailed_errors.append(
            "🟨 **Thứ tự theo template:** Phát hiện "
            f"{out_of_order_count} chương/mục sai thứ tự; "
            "đã bôi vàng."
        )
    if duplicate_count:
        detailed_errors.append(
            "🟨 **Cấu trúc chương:** Phát hiện "
            f"{duplicate_count} chương/mục lớn bị lặp; đã bôi vàng."
        )
    if missing_keys:
        detailed_errors.append(
            "🟨 **Cấu trúc theo template:** Thiếu "
            + ", ".join(
                _structure_label(key)
                for key in missing_keys
            )
            + "; đã chèn cảnh báo bôi vàng."
        )
    return {
        "merged": merged_count,
        "out_of_order": out_of_order_count,
        "duplicates": duplicate_count,
        "missing": len(missing_keys),
    }


ABBREVIATION_PATTERN = re.compile(
    r"(?<![0-9A-Za-zÀ-ỹĐđ])("
    r"(?:GS\.?\s*TS\.?|PGS\.?\s*TS\.?|ThS\.?|"
    r"BS(?:CKI{1,2}|NT)?\.?|"
    r"HbA1c|SpO2|PaO2|FiO2|eGFR|mmHg|SARS-CoV-\d+|"
    r"(?:[A-Za-zÀ-ỹĐđ]\.){2,}|"
    r"[A-Za-zÀ-ỹĐđ]{2,12}"
    r"(?:-[A-Za-zÀ-ỹĐđ0-9]{1,12})*)"
    r")(?![0-9A-Za-zÀ-ỹĐđ])"
)


def _is_all_caps_heading(text):
    stripped = text.strip()
    if not stripped or stripped != stripped.upper():
        return False
    normalized = _normalize_heading(stripped)
    one_line_headings = {
        "KET LUAN",
        "KIEN NGHI",
        "TONG QUAN",
        "PHU LUC",
        "MUC LUC",
        "DAT VAN DE",
    }
    if normalized in one_line_headings:
        return True
    if re.match(
        r"(?i)^\s*(?:CHƯƠNG|CHAPTER|PHẦN|MỤC)\s+"
        r"(?:\d+|[IVXLCDM]+)",
        stripped,
    ):
        return True
    words = re.findall(r"[A-Za-zÀ-ỹĐđ]+", stripped)
    return 2 <= len(words) <= 25


def _is_heading_like(paragraph):
    text = paragraph.text.strip()
    if not text:
        return True
    style_name = ""
    try:
        style_name = paragraph.style.name.lower()
    except Exception:
        pass
    if style_name.startswith(("heading", "title", "subtitle")):
        return True
    if _is_all_caps_heading(text):
        return True
    if (
        len(text.split()) <= 15
        and not re.search(r"[.!?]\s*$", text)
        and re.match(
            r"^\s*(?:\d+(?:\.\d+)*|[IVXLCDM]+)[.)\s-]+",
            text,
        )
    ):
        return True
    return False


def _paragraph_is_in_table(paragraph):
    parent = paragraph._p.getparent()
    while parent is not None:
        if str(parent.tag).endswith("}tc"):
            return True
        parent = parent.getparent()
    return False


def _find_abbreviation_ranges(text):
    ranges = []
    abbreviations = []
    for match in ABBREVIATION_PATTERN.finditer(text):
        abbreviation = match.group(1)
        known_mixed_case = bool(
            re.fullmatch(
                r"(?:ThS\.?|HbA1c|SpO2|PaO2|FiO2|"
                r"eGFR|mmHg|SARS-CoV-\d+)",
                abbreviation,
            )
        )
        if not known_mixed_case and not abbreviation.isupper():
            continue
        roman_candidate = re.sub(r"[^A-Z]", "", abbreviation.upper())
        if roman_candidate and re.fullmatch(
            r"[IVXLCDM]+",
            roman_candidate,
        ):
            continue
        ranges.append(match.span(1))
        abbreviations.append(abbreviation)
    return ranges, abbreviations


def _sentence_spans(text):
    spans = []
    start = 0
    for match in re.finditer(r"[.!?]+(?=\s+|$)", text):
        end = match.end()
        raw = text[start:end]
        left_offset = len(raw) - len(raw.lstrip())
        right_length = len(raw.rstrip())
        if right_length > left_offset:
            spans.append(
                (start + left_offset, start + right_length)
            )
        start = end
    if start < len(text):
        raw = text[start:]
        left_offset = len(raw) - len(raw.lstrip())
        right_length = len(raw.rstrip())
        if right_length > left_offset:
            spans.append(
                (start + left_offset, start + right_length)
            )
    return spans


def _sentence_lacks_subject(sentence):
    cleaned = re.sub(
        r"^\s*(?:[-–—•*]+|\(?\d+(?:\.\d+)*[.)])\s*",
        "",
        sentence,
    ).strip(" \t\"'“”")
    if len(re.findall(r"[A-Za-zÀ-ỹĐđ]+", cleaned)) < 4:
        return False

    lowered = cleaned.lower()
    if re.match(
        r"^(?:hình|bảng|sơ đồ|biểu đồ|đồ thị|chương|mục)\b",
        lowered,
    ):
        return False

    strong_subjectless_start = re.compile(
        r"^(?:cho thấy|nhận thấy|ghi nhận|chỉ ra|khẳng định|"
        r"kết luận(?: rằng)?|đề xuất|tiến hành|thực hiện|"
        r"sử dụng|áp dụng|đánh giá|phân tích|so sánh|"
        r"khảo sát|thu thập|lựa chọn|xác định|mô tả|"
        r"tính toán|kiểm định|phân loại|loại trừ|bao gồm|"
        r"gồm|được tiến hành|được thực hiện|cần(?: phải)?|"
        r"nên|phải)\b",
        re.IGNORECASE,
    )
    if strong_subjectless_start.search(cleaned):
        # Ví dụ "Phân tích hồi quy được sử dụng..." có "Phân tích
        # hồi quy" làm chủ ngữ, không phải câu thiếu chủ ngữ.
        if re.match(
            r"^(?:phân tích|đánh giá|so sánh)\b.{0,80}\bđược\b",
            lowered,
        ):
            return False
        return True

    contextual_error = re.match(
        r"^(?:qua|thông qua|dựa trên|dựa vào|từ kết quả|"
        r"từ phân tích)\b.{0,180}\b"
        r"(?:cho thấy|nhận thấy|ghi nhận|chỉ ra|khẳng định)\b",
        lowered,
    )
    if contextual_error:
        if re.search(
            r"\b(?:chúng tôi|tác giả|nhóm nghiên cứu)\b",
            lowered,
        ):
            return False
        if re.search(
            r",\s*(?:nghiên cứu này|kết quả|số liệu|dữ liệu|"
            r"đối tượng|bệnh nhân)\b",
            lowered,
        ):
            return False
        return True
    return False


def _find_subjectless_sentence_ranges(text):
    ranges = []
    for start, end in _sentence_spans(text):
        if _sentence_lacks_subject(text[start:end]):
            ranges.append((start, end))
    return ranges


def mark_abbreviations_and_subjectless_sentences(
    doc,
    cover_p_elements,
    detailed_errors,
    check_abbreviations=True,
    check_subjectless_sentences=True,
):
    abbreviation_counts = Counter()
    subjectless_count = 0
    in_references = False

    for paragraph in list(_all_document_paragraphs(doc)):
        text = paragraph.text.strip()
        if not text:
            continue
        if _is_actual_references_heading(paragraph):
            in_references = True
            continue
        if in_references and is_heading_after_references(text):
            in_references = False
        if in_references or paragraph._element in cover_p_elements:
            continue
        if (
            "CẢNH BÁO NGUỒN HÌNH" in text.upper()
            or "CẢNH BÁO CẤU TRÚC" in text.upper()
        ):
            continue

        abbreviation_ranges = []
        if check_abbreviations and not _is_all_caps_heading(text):
            abbreviation_ranges, abbreviations = (
                _find_abbreviation_ranges(paragraph.text)
            )
            abbreviation_counts.update(abbreviations)

        subjectless_ranges = []
        if (
            check_subjectless_sentences
            and
            not _is_heading_like(paragraph)
            and not _paragraph_is_image_caption(paragraph)
            and not _paragraph_is_in_table(paragraph)
        ):
            subjectless_ranges = _find_subjectless_sentence_ranges(
                paragraph.text
            )
            subjectless_count += len(subjectless_ranges)

        warning_ranges = abbreviation_ranges + subjectless_ranges
        if warning_ranges:
            highlight_paragraph_ranges(
                paragraph,
                warning_ranges,
                highlight_color=WD_COLOR_INDEX.YELLOW,
                preserve_existing=True,
            )

    if abbreviation_counts:
        abbreviation_list = ", ".join(
            sorted(abbreviation_counts, key=str.casefold)[:30]
        )
        detailed_errors.append(
            "🔠 **Từ viết tắt:** Đã phát hiện và bôi vàng "
            f"{sum(abbreviation_counts.values())} vị trí. "
            f"Các từ gồm: {abbreviation_list}."
        )
    if subjectless_count:
        detailed_errors.append(
            "🟨 **Câu có dấu hiệu thiếu chủ ngữ:** Đã bôi vàng "
            f"{subjectless_count} câu để học viên rà soát."
        )
    return {
        "abbreviations": sum(abbreviation_counts.values()),
        "subjectless_sentences": subjectless_count,
    }


LEGACY_VIETNAMESE_FONT_PATTERN = re.compile(
    r"(?i)(?:^\.?vn|vni[-_ ]?|tcvn|abc[-_ ]?font)"
)


def detect_legacy_vietnamese_fonts(doc, cover_p_elements, detailed_errors):
    """Phát hiện font TCVN3/VNI cũ; bôi đỏ thay vì đổi font gây sai dấu."""
    legacy_runs = set()
    legacy_fonts = Counter()
    for paragraph in _all_document_paragraphs(doc):
        if paragraph._element in cover_p_elements:
            continue
        for run in paragraph.runs:
            font_name = (run.font.name or "").strip()
            if not font_name or not LEGACY_VIETNAMESE_FONT_PATTERN.search(font_name):
                continue
            run.font.highlight_color = WD_COLOR_INDEX.RED
            legacy_runs.add(run._element)
            legacy_fonts[font_name] += 1

    if legacy_runs:
        font_list = ", ".join(sorted(legacy_fonts, key=str.casefold))
        detailed_errors.append(
            "🟥 **Bảng mã Unicode:** Phát hiện "
            f"{len(legacy_runs)} đoạn dùng font/bảng mã cũ ({font_list}); "
            "đã bôi đỏ. Chương trình giữ nguyên font cũ để tránh làm sai "
            "dấu tiếng Việt; học viên cần chuyển nội dung sang Unicode."
        )
    return legacy_runs


def process_docx_file(
    uploaded_bytes,
    rules,
    template_path=None,
    template_bytes=None,
    profile_label=None,
    regulation_filename=None,
):
    doc = docx.Document(io.BytesIO(uploaded_bytes))
    font_target = rules["font_name"]
    size_target = float(rules["font_size"])
    allowed_font_sizes = {
        float(value)
        for value in rules.get("allowed_font_sizes", [13.0, 14.0])
    }
    if not allowed_font_sizes:
        allowed_font_sizes = {size_target}
    if size_target not in allowed_font_sizes:
        size_target = sorted(allowed_font_sizes)[0]
    line_target = float(rules["line_spacing"])
    target_left = float(rules["margin_left"])
    target_right = float(rules["margin_right"])
    target_top = float(rules["margin_top"])
    target_bottom = float(rules["margin_bottom"])

    detailed_errors = []
    if profile_label:
        source_text = f"; quy định: {regulation_filename}" if regulation_filename else ""
        detailed_errors.append(
            f"📘 **Hồ sơ được chọn:** {profile_label}{source_text}."
        )
    check_structure_against_template(
        doc,
        template_path,
        detailed_errors,
        font_target,
        template_bytes=template_bytes,
    )
    cover_errors, cover_p_elements = optimize_cover_pages(doc)
    ethics_errors = fix_ethics_section(doc)
    detailed_errors.extend(cover_errors)
    detailed_errors.extend(ethics_errors)
    legacy_run_elements = detect_legacy_vietnamese_fonts(
        doc,
        cover_p_elements,
        detailed_errors,
    )

    for idx, section in enumerate(doc.sections):
        current_left = (
            round(section.left_margin.cm, 1) if section.left_margin else 0
        )
        current_right = (
            round(section.right_margin.cm, 1) if section.right_margin else 0
        )
        current_top = (
            round(section.top_margin.cm, 1) if section.top_margin else 0
        )
        current_bottom = (
            round(section.bottom_margin.cm, 1)
            if section.bottom_margin
            else 0
        )
        current_margins = (
            current_left,
            current_right,
            current_top,
            current_bottom,
        )
        target_margins = (
            target_left,
            target_right,
            target_top,
            target_bottom,
        )
        if current_margins != target_margins:
            detailed_errors.append(
                f"📌 **Lề trang (phần {idx + 1}):** Cũ "
                f"{current_left}-{current_right}-{current_top}-"
                f"{current_bottom} cm; đã sửa thành "
                f"{target_left}-{target_right}-{target_top}-"
                f"{target_bottom} cm."
            )
        section.left_margin = Cm(target_left)
        section.right_margin = Cm(target_right)
        section.top_margin = Cm(target_top)
        section.bottom_margin = Cm(target_bottom)

    spacing_count = 0
    title_bold_count = 0
    citation_paragraph_count = 0
    citation_moved_count = 0
    action_counts = Counter()
    # Duyệt theo thứ tự XML để xử lý thống nhất cả đoạn thường, đoạn trong
    # bảng và đoạn trong text box. Trước đây các đoạn trong bảng chạy ở một
    # vòng riêng nên số liệu báo cáo và một số thao tác không đồng nhất.
    paragraphs = list(_all_document_paragraphs(doc))
    in_references = False
    citation_style = str(
        rules.get("citation_style", "numeric_superscript")
    ).strip().lower()
    normalize_references = bool(rules.get("normalize_references", True))

    for index, paragraph in enumerate(paragraphs):
        if clean_spaces_and_punctuation(paragraph):
            spacing_count += 1
        if enforce_bold_for_thesis_title(paragraph):
            title_bold_count += 1

        if _is_actual_references_heading(paragraph):
            in_references = True
            continue
        if in_references and is_heading_after_references(paragraph.text):
            in_references = False

        if paragraph._element not in cover_p_elements:
            # Luôn nhận diện các trích dẫn tác giả-năm trong nội dung, kể cả
            # khi bộ quy định ưu tiên trích dẫn số. Nhờ vậy các trường hợp
            # "tác giả (năm)" và "Năm ... tác giả" vẫn được chuẩn hóa đúng.
            should_fix_author_citations = (
                (in_references and normalize_references)
                or not in_references
            )
            actions = []
            if should_fix_author_citations:
                actions = fix_author_citations(
                    paragraph,
                    in_references_section=in_references,
                )
                action_counts.update(actions)

            if not in_references and citation_style == "numeric_superscript":
                citation_result = normalize_numeric_citations(
                    paragraph,
                    font_target,
                    size_target,
                )
                # Chỉ đếm mục này khi thực sự chuyển từ ngoặc vuông;
                # trường hợp số vốn đã là lũy thừa nhưng đặt sai vị trí
                # được báo riêng ở citation_moved_count bên dưới.
                if citation_result["converted"]:
                    citation_paragraph_count += 1
                citation_moved_count += citation_result[
                    "moved_after_period"
                ]

    check_image_citations(
        doc,
        cover_p_elements,
        detailed_errors,
        font_target,
        require_source=bool(rules.get("require_image_source", True)),
        require_caption=bool(rules.get("require_image_caption", True)),
    )
    mark_abbreviations_and_subjectless_sentences(
        doc,
        cover_p_elements,
        detailed_errors,
        check_abbreviations=bool(rules.get("check_abbreviations", True)),
        check_subjectless_sentences=bool(
            rules.get("check_subjectless_sentences", True)
        ),
    )

    if title_bold_count:
        detailed_errors.append(
            "🖋️ **Tên đề tài:** Đã in đậm và bôi vàng."
        )
    if spacing_count:
        detailed_errors.append(
            f"🧹 **Dấu cách thừa:** Đã dọn tại "
            f"{spacing_count} đoạn văn."
        )
    if action_counts["body_added_et_al"]:
        detailed_errors.append(
            "✍️ **Trích dẫn tác giả trong nội dung:** Đã bổ sung "
            f"'và cộng sự' và bôi vàng phần thêm mới tại "
            f"{action_counts['body_added_et_al']} vị trí."
        )
    if action_counts["body_normalized_et_al"]:
        detailed_errors.append(
            "✍️ **Trích dẫn tác giả trong nội dung:** Đã chuẩn hóa "
            f"'et al./và cs.' tại "
            f"{action_counts['body_normalized_et_al']} vị trí."
        )
    if action_counts["author_missing_year"]:
        detailed_errors.append(
            "🟨 **Trích dẫn có tác giả và cộng sự nhưng thiếu năm:** "
            "Cần bổ sung năm; đã bôi vàng toàn bộ cụm trích dẫn tại "
            f"{action_counts['author_missing_year']} vị trí."
        )
    if action_counts["reference_added_et_al"]:
        detailed_errors.append(
            "📚 **Tài liệu tham khảo:** Đã bổ sung 'và cộng sự' hoặc "
            f"'et al.' sau 3 tác giả đầu tại "
            f"{action_counts['reference_added_et_al']} vị trí."
        )
    if action_counts["reference_trimmed_authors"]:
        detailed_errors.append(
            "📚 **Tài liệu tham khảo:** Đã giữ 3 tác giả đầu và rút gọn "
            f"danh sách tác giả tại "
            f"{action_counts['reference_trimmed_authors']} vị trí."
        )
    if action_counts["reference_parenthesized_year"]:
        detailed_errors.append(
            "📅 **Năm xuất bản:** Đã đưa năm vào ngoặc đơn tại "
            f"{action_counts['reference_parenthesized_year']} vị trí."
        )
    if action_counts["reference_missing_year"]:
        detailed_errors.append(
            "🟥 **Tài liệu tham khảo thiếu năm:** Đã bôi đỏ "
            f"{action_counts['reference_missing_year']} mục để bổ sung."
        )
    if action_counts["reference_author_name_warning"]:
        detailed_errors.append(
            "🟨 **Tên tác giả trong tài liệu tham khảo:** Đã bôi vàng "
            f"{action_counts['reference_author_name_warning']} mục có dấu "
            "hiệu viết tắt tên Việt Nam hoặc chưa viết tắt tên nước ngoài."
        )
    if action_counts["reference_styled"]:
        detailed_errors.append(
            "📖 **Định dạng tài liệu tham khảo:** Đã đặt tên bài báo in "
            "đứng, tên tạp chí in nghiêng và tập/số in đậm tại "
            f"{action_counts['reference_styled']} mục."
        )
    if action_counts["reference_structure_warning"]:
        detailed_errors.append(
            "🟨 **Cấu trúc tài liệu tham khảo:** Có "
            f"{action_counts['reference_structure_warning']} mục chưa nhận "
            "diện đủ tên bài báo – tên tạp chí – tập/số – trang; đã bôi vàng."
        )
    if action_counts["reference_removed_tr"]:
        detailed_errors.append(
            "📚 **Tài liệu tham khảo:** Đã bỏ chữ 'tr.' trước số trang "
            f"tại {action_counts['reference_removed_tr']} vị trí."
        )
    if action_counts["removed_doi"]:
        detailed_errors.append(
            f"🔗 **DOI:** Đã xóa {action_counts['removed_doi']} vị trí."
        )
    if citation_paragraph_count:
        detailed_errors.append(
            "✨ **Trích dẫn số:** Đã chuyển ngoặc vuông thành số lũy "
            f"thừa tại {citation_paragraph_count} đoạn."
        )
    if citation_moved_count:
        detailed_errors.append(
            "✨ **Vị trí trích dẫn số:** Đã đưa số lũy thừa ra sau "
            f"dấu chấm và bôi vàng cả dấu chấm cùng số trích dẫn tại "
            f"{citation_moved_count} vị trí."
        )

    font_size_corrected_count = 0
    for paragraph in paragraphs:
        if not paragraph.text.strip():
            continue
        if paragraph._element not in cover_p_elements:
            paragraph.paragraph_format.line_spacing = line_target
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            for run in paragraph.runs:
                if run._element in legacy_run_elements:
                    continue
                run.font.name = font_target
                if not run.font.superscript:
                    current_size = (
                        round(run.font.size.pt, 1)
                        if run.font.size is not None
                        else None
                    )
                    if current_size not in allowed_font_sizes:
                        run.font.size = Pt(size_target)
                        font_size_corrected_count += 1

    if font_size_corrected_count:
        allowed_sizes_text = ", ".join(
            f"{value:g}" for value in sorted(allowed_font_sizes)
        )
        detailed_errors.append(
            f"🔤 **Font và cỡ chữ:** Đã dùng {font_target}; giữ nguyên "
            f"các cỡ chữ được quy định ({allowed_sizes_text} pt) và đưa "
            f"{font_size_corrected_count} đoạn chữ ngoài các cỡ này về "
            f"{size_target:g} pt."
        )

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output, detailed_errors


def main():
    if st is None:
        raise RuntimeError(
            "Chưa cài Streamlit. Hãy chạy: pip install streamlit"
        )

    st.set_page_config(
        page_title=(
            "Đại học Y Hà Nội - Hệ thống Kiểm tra định dạng luận văn"
        ),
        page_icon="🎓",
        layout="wide",
    )

    col_logo, col_title = st.columns([1, 5])
    with col_logo:
        if os.path.exists("logo_hmu.png"):
            st.image("logo_hmu.png", width=125)
        else:
            st.markdown(
                "<h1 style='font-size:80px;margin:0;'>🎓</h1>",
                unsafe_allow_html=True,
            )
    with col_title:
        st.markdown(
            """
            <div style="line-height:1.2;margin-top:5px;">
              <p style="font-size:16px;font-weight:600;color:#666;
                        margin:0;text-transform:uppercase;">
                Trường Đại học Y Hà Nội
              </p>
              <h2 style="font-size:24px;font-weight:700;color:#ad171c;
                         margin:2px 0;">
                Trung tâm Khảo thí & ĐBCLGD
              </h2>
              <h2 style="font-size:24px;font-weight:700;color:#ad171c;
                         margin:2px 0 8px 0;">
                Bộ môn Mắt - Khúc xạ nhãn khoa
              </h2>
              <h3 style="font-size:18px;font-weight:500;color:#333;
                         margin:0;">
                🔬 Hệ thống Kiểm tra Luận văn và Đề cương tự động
              </h3>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.caption(
        "Hệ thống tự động kiểm tra quy định, bảo toàn trang bìa, "
        "khung viền, logo và chuẩn hóa file Word theo quy chế."
    )

    try:
        api_key = st.secrets.get(
            "GEMINI_API_KEY",
            os.environ.get("GEMINI_API_KEY", ""),
        )
    except Exception:
        api_key = os.environ.get("GEMINI_API_KEY", "")

    app_directory = os.path.dirname(os.path.abspath(__file__))
    st.subheader("1. Chọn loại hồ sơ cần kiểm tra")
    profile_key = st.selectbox(
        "Loại hồ sơ",
        options=list(DOCUMENT_PROFILES),
        format_func=lambda key: DOCUMENT_PROFILES[key]["label"],
        help=(
            "Mỗi lựa chọn được gắn cố định với một template Word và một "
            "file quy định PDF tương ứng."
        ),
    )
    profile = resolve_profile_paths(
        app_directory,
        DOCUMENT_PROFILES[profile_key],
    )
    template_path = profile["template_path"]
    regulation_path = profile["regulation_path"]

    template_exists = os.path.isfile(template_path)
    regulation_exists = os.path.isfile(regulation_path)
    parsed_rules = _copy_default_rules()
    pdf_text = ""

    if regulation_exists:
        pdf_text, parsed_rules = load_rules_for_pdf(
            regulation_path,
            os.path.getmtime(regulation_path),
            api_key,
        )

    left_status, right_status = st.columns(2)
    with left_status:
        if template_exists:
            st.success(
                f"✅ Template: {profile['resolved_template_file']}"
            )
        else:
            st.error(f"❌ Thiếu template: {profile['template_file']}")
    with right_status:
        if regulation_exists and pdf_text:
            st.success(
                f"✅ Quy định: {profile['resolved_regulation_file']}"
            )
        elif regulation_exists:
            st.error("❌ PDF quy định không có lớp chữ để đọc.")
        else:
            st.error(f"❌ Thiếu quy định: {profile['regulation_file']}")

    if regulation_exists and pdf_text:
        with st.expander("📌 Quy định đang được áp dụng", expanded=False):
            allowed_sizes = parsed_rules.get(
                "allowed_font_sizes",
                [parsed_rules.get("font_size", 13.0)],
            )
            allowed_sizes_text = ", ".join(
                f"{float(value):g}" for value in allowed_sizes
            )
            st.markdown(
                f"**Font:** {parsed_rules.get('font_name')} | "
                f"**Cỡ:** {allowed_sizes_text} pt | "
                f"**Giãn dòng:** {parsed_rules.get('line_spacing')}"
            )
            st.markdown(
                "**Lề trái – phải – trên – dưới:** "
                f"{parsed_rules.get('margin_left')} – "
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
        "Hệ thống tự đối chiếu cấu trúc với đúng template và định dạng "
        "với đúng PDF quy định của loại hồ sơ đã chọn."
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
                    file_name=profile["output_file"],
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
                    use_container_width=True,
                )


if __name__ == "__main__":
    main()
