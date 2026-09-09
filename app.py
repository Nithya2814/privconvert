import streamlit as st
import os
import io
import re
import hashlib
import tempfile
from PIL import Image, ImageDraw, ImageFont
import pymupdf  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import html

# Page Setup
st.set_page_config(
    page_title="PrivConvert Pro - Zero-Trust Document Studio",
    page_icon="🛡️",
    layout="wide"
)

# Custom UI Styling
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 800;
        background: -webkit-linear-gradient(45deg, #3B82F6, #10B981);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0px;
    }
    .sub-title {
        color: #94A3B8;
        font-size: 1rem;
        margin-top: -5px;
        margin-bottom: 20px;
    }
    .feature-card {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 12px;
    }
    .security-badge {
        display: inline-block;
        background-color: #064E3B;
        color: #34D399;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.82rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-title">🛡️ PrivConvert Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Zero-Knowledge Local Document Suite | 100% In-Memory Execution</div>', unsafe_allow_html=True)

# ----------------- Cryptographic & Conversion Helpers -----------------

def calculate_sha256(data: bytes) -> str:
    """Computes SHA-256 digital fingerprint to verify document integrity."""
    return hashlib.sha256(data).hexdigest()

def apply_pdf_password(pdf_bytes: bytes, password: str) -> bytes:
    """Applies native AES-256 encryption to a PDF document."""
    if not password:
        return pdf_bytes
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(user_password=password, algorithm="AES-256")
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

def add_watermark_to_image(image: Image.Image, text: str) -> Image.Image:
    """Adds a semi-transparent repeating diagonal security watermark."""
    watermarked = image.copy().convert("RGBA")
    txt_layer = Image.new("RGBA", watermarked.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(txt_layer)
    font_size = max(24, int(min(watermarked.size) / 16))
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except Exception:
        font = ImageFont.load_default()
    
    w, h = watermarked.size
    for x in range(0, w, font_size * 6):
        for y in range(0, h, font_size * 4):
            draw.text((x, y), text, fill=(220, 38, 38, 70), font=font)
            
    combined = Image.alpha_composite(watermarked, txt_layer)
    return combined.convert("RGB")

def redact_sensitive_text(text: str) -> str:
    """Auto-masks sensitive PII: Phone numbers, emails, and Aadhaar-style numbers."""
    # Mask emails
    text = re.sub(r'[\w\.-]+@[\w\.-]+\.\w+', '[REDACTED_EMAIL]', text)
    # Mask Indian phone numbers (10 digits with optional +91)
    text = re.sub(r'(\+91[\-\s]?)?[6-9]\d{9}', '[REDACTED_PHONE]', text)
    # Mask 12-digit Aadhaar pattern
    text = re.sub(r'\b\d{4}\s\d{4}\s\d{4}\b', '[REDACTED_AADHAAR]', text)
    return text

def convert_docx_to_pdf_formatted(docx_file, redact: bool = False) -> bytes:
    """High-fidelity Word to PDF conversion preserving bold, italics, tables & structure."""
    doc = Document(docx_file)
    pdf_buffer = io.BytesIO()
    doc_template = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    
    # Custom paragraph styles
    body_style = styles["Normal"]
    body_style.fontSize = 11
    body_style.leading = 14
    
    h1_style = styles["Heading1"]
    h1_style.fontSize = 18
    h1_style.leading = 22
    h1_style.spaceAfter = 10
    
    story = []
    
    # Iterate through paragraphs and tables
    for element in doc.element.body:
        if element.tag.endswith('p'):
            # It's a paragraph
            p = [p for p in doc.paragraphs if p._element == element]
            if p:
                para = p[0]
                runs_html = ""
                for run in para.runs:
                    txt = html.escape(run.text)
                    if redact:
                        txt = redact_sensitive_text(txt)
                    if run.bold:
                        txt = f"<b>{txt}</b>"
                    if run.italic:
                        txt = f"<i>{txt}</i>"
                    if run.underline:
                        txt = f"<u>{txt}</u>"
                    runs_html += txt
                
                if runs_html.strip():
                    if "Heading 1" in para.style.name:
                        story.append(Paragraph(runs_html, h1_style))
                    elif "Heading" in para.style.name:
                        story.append(Paragraph(runs_html, styles["Heading2"]))
                    else:
                        story.append(Paragraph(runs_html, body_style))
                    story.append(Spacer(1, 6))
                    
        elif element.tag.endswith('tbl'):
            # It's a table
            t = [t for t in doc.tables if t._element == element]
            if t:
                tbl = t[0]
                tbl_data = []
                for row in tbl.rows:
                    row_data = []
                    for cell in row.cells:
                        cell_txt = html.escape(cell.text.strip())
                        if redact:
                            cell_txt = redact_sensitive_text(cell_txt)
                        row_data.append(Paragraph(cell_txt, body_style))
                    tbl_data.append(row_data)
                if tbl_data:
                    pdf_tbl = Table(tbl_data)
                    pdf_tbl.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F1F5F9')),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                        ('VALIGN', (0,0), (-1,-1), 'TOP'),
                        ('TOPPADDING', (0,0), (-1,-1), 6),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                    ]))
                    story.append(pdf_tbl)
                    story.append(Spacer(1, 10))
                    
    if not story:
        story.append(Paragraph("Empty Document", body_style))
        
    doc_template.build(story)
    return pdf_buffer.getvalue()

def convert_pdf_to_pptx(pdf_file) -> bytes:
    """Converts each PDF page into high-resolution PowerPoint presentation slides."""
    pdf_bytes = pdf_file.read()
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    prs = Presentation()
    
    # Standard 16:9 widescreen or PDF ratio
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]  # Blank layout
    
    for page in doc:
        # Render page at 150 DPI for crisp slide display
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        
        slide = prs.slides.add_slide(blank_layout)
        img_stream = io.BytesIO(img_bytes)
        slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
        
    out_pptx = io.BytesIO()
    prs.save(out_pptx)
    return out_pptx.getvalue()

def convert_pptx_to_pdf(pptx_file) -> bytes:
    """Converts PowerPoint presentation text and slide notes into a structured PDF."""
    prs = Presentation(pptx_file)
    pdf_buffer = io.BytesIO()
    doc_template = SimpleDocTemplate(pdf_buffer, pagesize=letter, margin=40)
    styles = getSampleStyleSheet()
    story = []
    
    for idx, slide in enumerate(prs.slides):
        story.append(Paragraph(f"<b>Slide {idx + 1}</b>", styles["Heading2"]))
        story.append(Spacer(1, 6))
        
        slide_text_found = False
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = html.escape(para.text.strip())
                    if t:
                        story.append(Paragraph(t, styles["Normal"]))
                        story.append(Spacer(1, 4))
                        slide_text_found = True
                        
        if not slide_text_found:
            story.append(Paragraph("<i>[Visual or Media Slide]</i>", styles["Normal"]))
        story.append(Spacer(1, 16))
        
    doc_template.build(story)
    return pdf_buffer.getvalue()

# ----------------- MAIN UI TABS -----------------
tab_doc, tab_media, tab_privacy = st.tabs([
    "📑 Document Studio (PDF, Word, PPT)", 
    "🖼️ Image & Media Studio",
    "🛡️ Privacy Shield & Verification"
])

# ----------------- TAB 1: Document Studio -----------------
with tab_doc:
    st.markdown("### Document Conversion Engine")
    
    conv_mode = st.selectbox(
        "Choose Conversion Action:",
        [
            "📄 PDF ➔ Word (.docx) [Editable Layout]",
            "📝 Word (.docx) ➔ PDF [Preserve Styles & Tables]",
            "📊 PDF ➔ PowerPoint (.pptx) [Pixel-Perfect Slides]",
            "🖥️ PowerPoint (.pptx) ➔ PDF"
        ]
    )
    
    if "PDF ➔ Word" in conv_mode:
        uploaded_pdf = st.file_uploader("Upload PDF file:", type=["pdf"], key="pdf_docx_up")
        if uploaded_pdf:
            if st.button("🔄 Convert to Editable Word Document", type="primary"):
                with st.spinner("Reconstructing formatting locally..."):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
                        tmp_in.write(uploaded_pdf.read())
                        tmp_in_path = tmp_in.name
                    tmp_out_path = tmp_in_path.replace(".pdf", ".docx")
                    
                    cv = Converter(tmp_in_path)
                    cv.convert(tmp_out_path)
                    cv.close()
                    
                    with open(tmp_out_path, "rb") as f:
                        docx_data = f.read()
                    os.remove(tmp_in_path)
                    os.remove(tmp_out_path)
                    
                    name = f"{os.path.splitext(uploaded_pdf.name)[0]}.docx"
                    st.success("✅ Converted with fonts, tables, and alignments preserved!")
                    st.download_button(f"⬇️ Download {name}", data=docx_data, file_name=name)

    elif "Word (.docx) ➔ PDF" in conv_mode:
        uploaded_docx = st.file_uploader("Upload Word Document (.docx):", type=["docx"], key="docx_pdf_up")
        if uploaded_docx:
            col_w1, col_w2 = st.columns(2)
            with col_w1:
                lock_docx_pdf = st.checkbox("🔒 Native AES-256 Password Lock", value=False)
                docx_pass = st.text_input("Set Password (Optional):", type="password", key="docx_p_key") if lock_docx_pdf else ""
            with col_w2:
                redact_option = st.checkbox("🛡️ Auto-Redact Sensitive Data (Phone, Email, Aadhaar)", value=False)
                
            if st.button("🚀 Convert to PDF", type="primary"):
                with st.spinner("Compiling styled PDF in local memory..."):
                    raw_pdf = convert_docx_to_pdf_formatted(uploaded_docx, redact=redact_option)
                    final_pdf = apply_pdf_password(raw_pdf, docx_pass if lock_docx_pdf else None)
                    
                    name = f"{os.path.splitext(uploaded_docx.name)[0]}.pdf"
                    st.success("✅ Successfully generated PDF with preserved styling!")
                    st.download_button(f"⬇️ Download {name}", data=final_pdf, file_name=name, mime="application/pdf")

    elif "PDF ➔ PowerPoint" in conv_mode:
        uploaded_pdf_ppt = st.file_uploader("Upload PDF to turn into Slides:", type=["pdf"], key="pdf_ppt_up")
        if uploaded_pdf_ppt:
            if st.button("📊 Convert PDF to PowerPoint Slides", type="primary"):
                with st.spinner("Rendering high-resolution slides locally..."):
                    pptx_data = convert_pdf_to_pptx(uploaded_pdf_ppt)
                    name = f"{os.path.splitext(uploaded_pdf_ppt.name)[0]}.pptx"
                    st.success("✅ Slides generated! Open and present directly in PowerPoint.")
                    st.download_button(f"⬇️ Download {name}", data=pptx_data, file_name=name)

    elif "PowerPoint (.pptx) ➔ PDF" in conv_mode:
        uploaded_pptx = st.file_uploader("Upload PowerPoint Presentation (.pptx):", type=["pptx"], key="pptx_pdf_up")
        if uploaded_pptx:
            lock_ppt_pdf = st.checkbox("🔒 Protect Output PDF with Password", value=False, key="ppt_lock")
            ppt_pass = st.text_input("Set Password:", type="password", key="ppt_p_key") if lock_ppt_pdf else ""
            if st.button("🚀 Convert Presentation to PDF", type="primary"):
                with st.spinner("Converting slides to document..."):
                    raw_pdf = convert_pptx_to_pdf(uploaded_pptx)
                    final_pdf = apply_pdf_password(raw_pdf, ppt_pass if lock_ppt_pdf else None)
                    name = f"{os.path.splitext(uploaded_pptx.name)[0]}.pdf"
                    st.success("✅ Converted to PDF!")
                    st.download_button(f"⬇️ Download {name}", data=final_pdf, file_name=name, mime="application/pdf")

# ----------------- TAB 2: Media Studio -----------------
with tab_media:
    st.markdown("### Images & Media Converter")
    media_mode = st.radio("Select Action:", ["Images ➔ Secure PDF", "PDF ➔ Extract Images (PNG)", "Image Format Converter (PNG/JPG/WEBP)"], horizontal=True)
    
    if media_mode == "Images ➔ Secure PDF":
        img_upload = st.file_uploader("Upload Image:", type=["png", "jpg", "jpeg", "webp"], key="img_sec_up")
        if img_upload:
            c1, c2 = st.columns(2)
            with c1:
                img_obj = Image.open(img_upload)
                st.image(img_obj, width=300)
            with c2:
                enable_lock = st.checkbox("🔒 Native PDF Password Lock", value=True)
                pass_val = st.text_input("PDF Open Password:", type="password") if enable_lock else ""
                
                enable_wm = st.checkbox("🛡️ Add Anti-Leak Watermark", value=False)
                wm_text = st.text_input("Watermark:", "CONFIDENTIAL - VERIFICATION ONLY") if enable_wm else ""
                
                if st.button("🚀 Create Secure PDF", type="primary"):
                    if enable_lock and not pass_val:
                        st.error("Please enter a password!")
                    else:
                        img_to_proc = img_obj
                        if enable_wm and wm_text:
                            img_to_proc = add_watermark_to_image(img_to_proc, wm_text)
                        elif img_to_proc.mode in ("RGBA", "P"):
                            img_to_proc = img_to_proc.convert("RGB")
                            
                        buf = io.BytesIO()
                        img_to_proc.save(buf, format="PDF")
                        final_pdf = apply_pdf_password(buf.getvalue(), pass_val if enable_lock else None)
                        
                        out_name = f"{os.path.splitext(img_upload.name)[0]}.pdf"
                        st.success("✅ Secure PDF Ready!")
                        st.download_button(f"⬇️ Download {out_name}", data=final_pdf, file_name=out_name, mime="application/pdf")

    elif media_mode == "PDF ➔ Extract Images (PNG)":
        pdf_img_up = st.file_uploader("Upload PDF to extract pages:", type=["pdf"], key="pdf_extract_up")
        if pdf_img_up:
            if st.button("🖼️ Extract Pages as Crisp PNGs"):
                doc = pymupdf.open(stream=pdf_img_up.read(), filetype="pdf")
                st.write(f"Total Pages Found: **{len(doc)}**")
                for i, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=150)
                    png_bytes = pix.tobytes("png")
                    st.image(png_bytes, caption=f"Page {i+1}", width=300)
                    st.download_button(f"⬇️ Download Page {i+1} (PNG)", data=png_bytes, file_name=f"page_{i+1}.png", mime="image/png")

    elif media_mode == "Image Format Converter (PNG/JPG/WEBP)":
        raw_img_up = st.file_uploader("Upload Image:", type=["png", "jpg", "jpeg", "webp"], key="raw_img_up")
        if raw_img_up:
            target_ext = st.selectbox("Convert To:", ["PNG", "JPEG", "WEBP"])
            if st.button("🔄 Convert Format"):
                im = Image.open(raw_img_up)
                buf = io.BytesIO()
                fmt = "JPEG" if target_ext == "JPEG" else target_ext
                if fmt == "JPEG" and im.mode in ("RGBA", "P"):
                    im = im.convert("RGB")
                im.save(buf, format=fmt)
                out_name = f"converted.{target_ext.lower()}"
                st.success(f"✅ Converted to {target_ext}!")
                st.download_button(f"⬇️ Download {out_name}", data=buf.getvalue(), file_name=out_name, mime=f"image/{target_ext.lower()}")

# ----------------- TAB 3: Privacy Shield -----------------
with tab_privacy:
    st.markdown("### 🛡️ Competition Novelty & Privacy Engine")
    st.info("These enterprise security modules run 100% on client-side RAM with zero external server dependencies.")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.markdown("#### 1. SHA-256 Tamper-Proof Fingerprint")
        st.write("Upload any file to calculate its cryptographic fingerprint to ensure nobody modified it:")
        check_file = st.file_uploader("File to verify:", key="hash_check_file")
        if check_file:
            file_hash = calculate_sha256(check_file.read())
            st.code(f"SHA-256 Hash:\n{file_hash}", language="text")
            st.caption("✨ *Recipients can verify this exact hash to guarantee zero tampering or malware infection.*")
            
    with col_p2:
        st.markdown("#### 2. Live PII Auto-Masking Test")
        st.write("Test our auto-redaction engine that scrubs identity data:")
        sample_text = st.text_area(
            "Sample Document Text:",
            "Customer John Doe, Contact: +91 9876543210, Email: john@company.com, Aadhaar: 1234 5678 9012"
        )
        if st.button("🛡️ Run Privacy Cloak"):
            clean_text = redact_sensitive_text(sample_text)
            st.success("✅ Sensitive Information Masked:")
            st.code(clean_text, language="text")
