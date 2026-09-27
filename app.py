import difflib
import copy
import io
import json
import os
import re
import unicodedata
from collections import Counter

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
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
    "first_line_indent": 1.0,
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
        "Bảng mã Unicode; Times New Roman cỡ 13 hoặc 14; giãn dòng 1,5; "
        "đoạn văn nội dung thụt đầu dòng 1,0 cm.",
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
