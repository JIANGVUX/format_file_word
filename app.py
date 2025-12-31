# app.py
import streamlit as st
from formatter import (
    ReportConfig, DocxReportFormatter,
    PAPER_PRESET, PAGE_NUMBER_POSITION, ALIGN_MAP, PAGE_FMT_MAP,
    load_config_json_bytes, save_config_json_bytes
)

st.set_page_config(page_title="DOCX Report Formatter", layout="wide")

st.title("📄 DOCX Report Formatter (Chuẩn hoá báo cáo)")
st.caption("Upload file .docx → chỉnh chuẩn ở sidebar → Format → Download file output")

# ---------------------------
# State
# ---------------------------
if "cfg" not in st.session_state:
    st.session_state.cfg = ReportConfig()

cfg: ReportConfig = st.session_state.cfg

# ---------------------------
# Sidebar: Config import/export
# ---------------------------
st.sidebar.header("⚙️ Cấu hình")

with st.sidebar.expander("Import/Export Config (.json)", expanded=True):
    up_json = st.file_uploader("Import config JSON", type=["json"], key="cfg_json")
    if up_json is not None:
        try:
            st.session_state.cfg = load_config_json_bytes(up_json.read())
            cfg = st.session_state.cfg
            st.success("Đã import config ✅")
        except Exception as e:
            st.error(f"Import config lỗi: {e}")

    st.download_button(
        "⬇️ Download config hiện tại",
        data=save_config_json_bytes(cfg),
        file_name="report_config.json",
        mime="application/json"
    )

# ---------------------------
# Sidebar: Layout & margins
# ---------------------------
with st.sidebar.expander("Layout & Lề", expanded=True):
    cfg.pagesetup.paper = st.selectbox("Khổ giấy", list(PAPER_PRESET.keys()), index=list(PAPER_PRESET.keys()).index(cfg.pagesetup.paper))
    col1, col2 = st.columns(2)
    with col1:
        cfg.pagesetup.margin_left_cm = st.number_input("Lề trái (cm)", 0.5, 10.0, float(cfg.pagesetup.margin_left_cm), 0.1)
        cfg.pagesetup.margin_top_cm = st.number_input("Lề trên (cm)", 0.5, 10.0, float(cfg.pagesetup.margin_top_cm), 0.1)
        cfg.pagesetup.header_distance_cm = st.number_input("Khoảng header (cm)", 0.0, 5.0, float(cfg.pagesetup.header_distance_cm), 0.05)
    with col2:
        cfg.pagesetup.margin_right_cm = st.number_input("Lề phải (cm)", 0.5, 10.0, float(cfg.pagesetup.margin_right_cm), 0.1)
        cfg.pagesetup.margin_bottom_cm = st.number_input("Lề dưới (cm)", 0.5, 10.0, float(cfg.pagesetup.margin_bottom_cm), 0.1)
        cfg.pagesetup.footer_distance_cm = st.number_input("Khoảng footer (cm)", 0.0, 5.0, float(cfg.pagesetup.footer_distance_cm), 0.05)
    cfg.pagesetup.different_first_page = st.checkbox("Trang đầu khác header/footer", value=bool(cfg.pagesetup.different_first_page))

# ---------------------------
# Helper: style editor
# ---------------------------
def style_editor(title: str, sc):
    st.sidebar.subheader(title)
    sc.font_name = st.sidebar.text_input(f"{title} - Font", value=sc.font_name, key=f"{title}_font")
    sc.font_size_pt = st.sidebar.number_input(f"{title} - Size (pt)", 6.0, 72.0, float(sc.font_size_pt), 0.5, key=f"{title}_size")
    sc.line_spacing = st.sidebar.number_input(f"{title} - Line spacing", 1.0, 3.0, float(sc.line_spacing), 0.1, key=f"{title}_ls")
    c1, c2 = st.sidebar.columns(2)
    with c1:
        sc.space_before_pt = st.number_input(f"{title} - Before (pt)", 0.0, 48.0, float(sc.space_before_pt), 1.0, key=f"{title}_before")
        sc.first_line_indent_cm = st.number_input(f"{title} - Indent (cm)", 0.0, 5.0, float(sc.first_line_indent_cm), 0.1, key=f"{title}_indent")
    with c2:
        sc.space_after_pt = st.number_input(f"{title} - After (pt)", 0.0, 48.0, float(sc.space_after_pt), 1.0, key=f"{title}_after")
        sc.alignment = st.selectbox(f"{title} - Align", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(sc.alignment), key=f"{title}_align")
    b1, b2 = st.sidebar.columns(2)
    with b1:
        sc.bold = st.checkbox(f"{title} - Bold", value=bool(sc.bold), key=f"{title}_bold")
    with b2:
        sc.italic = st.checkbox(f"{title} - Italic", value=bool(sc.italic), key=f"{title}_italic")

# ---------------------------
# Sidebar: Styles
# ---------------------------
with st.sidebar.expander("Styles (Font/Đoạn)", expanded=False):
    style_editor("Normal", cfg.normal)
    style_editor("Title", cfg.title)
    style_editor("Heading 1", cfg.heading1)
    style_editor("Heading 2", cfg.heading2)
    style_editor("Heading 3", cfg.heading3)
    style_editor("Caption", cfg.caption)

# ---------------------------
# Sidebar: Page number
# ---------------------------
with st.sidebar.expander("Đánh số trang", expanded=True):
    cfg.pagenumber.enabled = st.checkbox("Bật số trang", value=bool(cfg.pagenumber.enabled))
    cfg.pagenumber.position = st.selectbox("Vị trí", PAGE_NUMBER_POSITION, index=PAGE_NUMBER_POSITION.index(cfg.pagenumber.position))
    cfg.pagenumber.template = st.text_input("Template", value=cfg.pagenumber.template, help="Dùng {PAGE}, {NUMPAGES}")
    col1, col2 = st.columns(2)
    with col1:
        cfg.pagenumber.start_at = st.number_input("Bắt đầu từ", 1, 999, int(cfg.pagenumber.start_at), 1)
        cfg.pagenumber.number_format = st.selectbox("Định dạng số", list(PAGE_FMT_MAP.keys()), index=list(PAGE_FMT_MAP.keys()).index(cfg.pagenumber.number_format))
    with col2:
        cfg.pagenumber.restart_each_section = st.checkbox("Restart mỗi section", value=bool(cfg.pagenumber.restart_each_section))
        cfg.pagenumber.font_size_pt = st.number_input("Size số trang (pt)", 6.0, 36.0, float(cfg.pagenumber.font_size_pt), 0.5)
    cfg.pagenumber.font_name = st.text_input("Font số trang", value=cfg.pagenumber.font_name)

# ---------------------------
# Sidebar: TOC
# ---------------------------
with st.sidebar.expander("Mục lục (TOC)", expanded=False):
    cfg.toc.insert_toc = st.checkbox("Chèn TOC", value=bool(cfg.toc.insert_toc))
    cfg.toc.heading_levels = st.text_input("Cấp heading (vd 1-3)", value=cfg.toc.heading_levels)
    cfg.toc.title = st.text_input("Tiêu đề TOC", value=cfg.toc.title)
    cfg.toc.title_bold = st.checkbox("Bold tiêu đề", value=bool(cfg.toc.title_bold))
    cfg.toc.title_font_size_pt = st.number_input("Size tiêu đề TOC", 10.0, 24.0, float(cfg.toc.title_font_size_pt), 0.5)
    cfg.toc.title_alignment = st.selectbox("Canh tiêu đề", list(ALIGN_MAP.keys()), index=list(ALIGN_MAP.keys()).index(cfg.toc.title_alignment))

# ---------------------------
# Sidebar: Advanced
# ---------------------------
with st.sidebar.expander("Nâng cao", expanded=False):
    cfg.processing.force_run_font_everywhere = st.checkbox("Ép font cho mọi run (triệt để)", value=bool(cfg.processing.force_run_font_everywhere))
    cfg.processing.force_paragraph_format_everywhere = st.checkbox("Ép format cho mọi đoạn", value=bool(cfg.processing.force_paragraph_format_everywhere))
    cfg.processing.include_tables = st.checkbox("Xử lý cả nội dung trong bảng", value=bool(cfg.processing.include_tables))
    st.info("Tip: Mở file output trong Word → Ctrl+A → F9 để cập nhật số trang / mục lục.")

# ---------------------------
# Main: Upload + Format
# ---------------------------
st.subheader("1) Upload file .docx")
up_docx = st.file_uploader("Chọn file DOCX", type=["docx"], key="docx")

st.subheader("2) Format & Download")

colA, colB = st.columns([1, 1])
with colA:
    st.write("✅ Bạn có thể chỉnh chuẩn ở sidebar.")
    st.write("✅ File xử lý trên server Streamlit, không cần cài Word.")
with colB:
    st.write("⚠️ Word fields (PAGE/TOC) thường cần cập nhật khi mở file.")
    st.write("⚠️ Nếu tài liệu có định dạng đặc thù, tắt 'Ép font cho mọi run' để giữ nguyên một số đoạn.")

if up_docx is None:
    st.warning("Hãy upload 1 file .docx để bắt đầu.")
else:
    input_bytes = up_docx.read()
    in_name = up_docx.name
    base = in_name[:-5] if in_name.lower().endswith(".docx") else in_name
    out_name = f"{base}_FORMATTED.docx"

    if st.button("🚀 FORMAT NGAY", type="primary"):
        try:
            formatter = DocxReportFormatter(cfg)
            output_bytes = formatter.format_docx_bytes(input_bytes)
            st.success("Format xong ✅")
            st.download_button(
                "⬇️ Download file đã chuẩn hoá",
                data=output_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Format lỗi: {e}")
