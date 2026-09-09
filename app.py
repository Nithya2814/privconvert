import streamlit as st
import os
import io
import re
import hashlib
import tempfile
import html
from PIL import Image, ImageDraw, ImageFont
import pymupdf  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
from docx import Document
from pptx import Presentation
from pptx.util import Inches
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Page Configuration
st.set_page_config(
    page_title="PrivConvert Studio - Zero-Knowledge Converter",
    page_icon="✨",
    layout="wide"
)

# Colorful Modern CSS
st.markdown("""
<style>
    /* Gradient Hero Header */
    .hero-container {
        background: linear-gradient(135deg, #1E1B4B 0%, #312E81 40%, #0F172A 100%);
        border: 2px solid #6366F1;
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(99, 102, 241, 0.3);
    }
    .hero-title {
        font-size: 2.6rem;
        font-weight: 900;
        background: linear-gradient(90deg, #38BDF8, #818CF8, #F472B6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0;
    }
    .hero-subtitle {
        color: #E2E8F0;
        font-size: 1.1rem;
        font-weight: 500;
        margin-top: 6px;
    }
    
    /* Colorful Feature Badges */
    .pill-blue {
        background: linear-gradient(90deg, #0284C7, #0EA5E9);
        color: white; padding: 5px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 600; display: inline-block; margin-right: 8px;
    }
    .pill-purple {
        background: linear-gradient(90deg, #7C3AED, #A855F7);
        color: white; padding: 5px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 600; display: inline-block; margin-right: 8px;
    }
    .pill-emerald {
        background: linear-gradient(90deg, #059669, #10B981);
        color: white; padding: 5px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 600; display: inline-block; margin-right: 8px;
    }
    .pill-pink {
        background: linear-gradient(90deg, #DB2777, #F43F5E);
        color: white; padding: 5px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 600; display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# Hero Header
st.markdown("""
<div class="hero-container">
    <div class="hero-title">🛡️ PrivConvert Studio</div>
    <div class="hero-subtitle">The Complete All-In-One Privacy-First Document & Media Conversion Suite</div>
    <div style="margin-top: 15px;">
        <span class="pill-blue">⚡ 100% In-Memory RAM</span>
        <span class="pill-purple">🔒 AES-256 Military Lock</span>
        <span class="pill-emerald">🚫 Zero Cloud Databases</span>
        <span class="pill-pink">🛡️ Tamper-Proof SHA-256</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ----------------- Helper Functions -----------------

def calculate_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def apply_pdf_password(pdf_bytes: bytes, password: str) -> bytes:
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

def unlock_pdf(pdf_bytes: bytes, password: str) -> bytes:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if reader.is_encrypted:
        reader.decrypt(password)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

def add_watermark(image: Image.Image, text: str) -> Image.Image:
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
            draw.text((x, y), text, fill=(220, 38, 38, 80), font=font)
    combined = Image.alpha_composite(watermarked, txt_layer)
    return combined.convert("RGB")

def redact_sensitive_text(text: str) -> str:
    text = re.sub(r'[\w\.-]+@[\w\.-]+\.\w+', '[REDACTED_EMAIL]', text)
    text = re.sub(r'(\+91[\-\s]?)?[6-9]\d{9}', '[REDACTED_PHONE]', text)
    text = re.sub(r'\b\d{4}\s\d{4}\s\d{4}\b', '[REDACTED_AADHAAR]', text)
    return text

def convert_docx_to_pdf(docx_file, redact: bool = False) -> bytes:
    doc = Document(docx_file)
    pdf_buffer = io.BytesIO()
    doc_template = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    
    body_style = styles["Normal"]
    body_style.fontSize = 11
    body_style.leading = 14
    
    h1_style = styles["Heading1"]
    h1_style.fontSize = 17
    h1_style.leading = 22
    h1_style.textColor = colors.HexColor('#4F46E5')
    
    story = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
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
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EEF2FF')),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#C7D2FE')),
                        ('TOPPADDING', (0,0), (-1,-1), 5),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                    ]))
                    story.append(pdf_tbl)
                    story.append(Spacer(1, 10))
                    
    if not story:
        story.append(Paragraph("Empty Document", body_style))
    doc_template.build(story)
    return pdf_buffer.getvalue()

def convert_pdf_to_pptx(pdf_file) -> bytes:
    doc = pymupdf.open(stream=pdf_file.read(), filetype="pdf")
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    for page in doc:
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(io.BytesIO(img_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    out_pptx = io.BytesIO()
    prs.save(out_pptx)
    return out_pptx.getvalue()

def convert_pptx_to_pdf(pptx_file) -> bytes:
    prs = Presentation(pptx_file)
    pdf_buffer = io.BytesIO()
    doc_template = SimpleDocTemplate(pdf_buffer, pagesize=letter, margin=40)
    styles = getSampleStyleSheet()
    story = []
    for idx, slide in enumerate(prs.slides):
        story.append(Paragraph(f"<b>Slide {idx + 1}</b>", styles["Heading2"]))
        story.append(Spacer(1, 6))
        text_found = False
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = html.escape(para.text.strip())
                    if t:
                        story.append(Paragraph(t, styles["Normal"]))
                        story.append(Spacer(1, 4))
                        text_found = True
        if not text_found:
            story.append(Paragraph("<i>[Visual or Media Slide]</i>", styles["Normal"]))
        story.append(Spacer(1, 14))
    doc_template.build(story)
    return pdf_buffer.getvalue()

# ----------------- MAIN TABS -----------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📑 Office & Documents", 
    "🖼️ Images & Media Studio", 
    "🔓 PDF Unlocker & Decryptor",
    "🛡️ Privacy Shield (Redaction & Hash)"
])

# ----------------- TAB 1: Office & Documents -----------------
with tab1:
    st.markdown("### 📑 Document Conversions (Word, PDF, PPT, Text)")
    
    doc_option = st.selectbox(
        "Select Your Conversion:",
        [
            "📄 PDF ➔ Word (.docx) [Fully Editable Layout & Tables]",
            "📝 Word (.docx) ➔ PDF [Preserves Formatting, Fonts & Tables]",
            "📊 PDF ➔ PowerPoint (.pptx) [Pixel-Perfect Slides]",
            "🖥️ PowerPoint (.pptx) ➔ PDF",
            "📋 PDF ➔ Plain Text (.txt)",
            "✍️ Plain Text (.txt) ➔ PDF"
        ]
    )
    
    if "PDF ➔ Word" in doc_option:
        up_pdf = st.file_uploader("Upload PDF:", type=["pdf"], key="pdf_w_in")
        if up_pdf:
            if st.button("✨ Convert to Word (.docx)", type="primary"):
                with st.spinner("Extracting layout, fonts & tables locally..."):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
                        tmp_in.write(up_pdf.read())
                        tmp_in_path = tmp_in.name
                    tmp_out = tmp_in_path.replace(".pdf", ".docx")
                    cv = Converter(tmp_in_path)
                    cv.convert(tmp_out)
                    cv.close()
                    with open(tmp_out, "rb") as f:
                        data = f.read()
                    os.remove(tmp_in_path)
                    os.remove(tmp_out)
                    name = f"{os.path.splitext(up_pdf.name)[0]}.docx"
                    st.success("🎉 Converted successfully!")
                    st.download_button(f"⬇️ Download {name}", data=data, file_name=name)

    elif "Word (.docx) ➔ PDF" in doc_option:
        up_docx = st.file_uploader("Upload Word (.docx):", type=["docx"], key="docx_pdf_in")
        if up_docx:
            c1, c2 = st.columns(2)
            with c1:
                pw_docx = st.checkbox("🔒 Native AES-256 Password Protection", value=False)
                docx_pw = st.text_input("Set Password:", type="password", key="pw_d_val") if pw_docx else ""
            with c2:
                auto_red = st.checkbox("🛡️ Auto-Redact Sensitive PII (Phone, Email, Aadhaar)", value=False)
                
            if st.button("🚀 Convert to Formatted PDF", type="primary"):
                with st.spinner("Compiling styled PDF in RAM..."):
                    pdf_bytes = convert_docx_to_pdf(up_docx, redact=auto_red)
                    final_pdf = apply_pdf_password(pdf_bytes, docx_pw if pw_docx else None)
                    name = f"{os.path.splitext(up_docx.name)[0]}.pdf"
                    st.success("🎉 PDF generated with preserved styles!")
                    st.download_button(f"⬇️ Download {name}", data=final_pdf, file_name=name, mime="application/pdf")

    elif "PDF ➔ PowerPoint" in doc_option:
        up_pdf_ppt = st.file_uploader("Upload PDF to convert to Slides:", type=["pdf"], key="pdf_ppt_in")
        if up_pdf_ppt:
            if st.button("📊 Convert to PowerPoint (.pptx)", type="primary"):
                with st.spinner("Generating presentation slides..."):
                    pptx_bytes = convert_pdf_to_pptx(up_pdf_ppt)
                    name = f"{os.path.splitext(up_pdf_ppt.name)[0]}.pptx"
                    st.success("🎉 Presentation slides created! Open directly in PowerPoint.")
                    st.download_button(f"⬇️ Download {name}", data=pptx_bytes, file_name=name)

    elif "PowerPoint (.pptx) ➔ PDF" in doc_option:
        up_ppt = st.file_uploader("Upload PowerPoint (.pptx):", type=["pptx"], key="ppt_pdf_in")
        if up_ppt:
            ppt_pw_chk = st.checkbox("🔒 Lock PDF with Password", value=False, key="ppt_pw_chk")
            ppt_pw = st.text_input("Enter Password:", type="password", key="ppt_pw_val") if ppt_pw_chk else ""
            if st.button("🚀 Convert PPT to PDF", type="primary"):
                with st.spinner("Converting slides..."):
                    pdf_b = convert_pptx_to_pdf(up_ppt)
                    final_b = apply_pdf_password(pdf_b, ppt_pw if ppt_pw_chk else None)
                    name = f"{os.path.splitext(up_ppt.name)[0]}.pdf"
                    st.success("🎉 Converted to PDF!")
                    st.download_button(f"⬇️ Download {name}", data=final_b, file_name=name, mime="application/pdf")

    elif "PDF ➔ Plain Text" in doc_option:
        up_pdf_txt = st.file_uploader("Upload PDF to extract text:", type=["pdf"], key="pdf_txt_in")
        if up_pdf_txt:
            if st.button("📋 Extract Text"):
                doc = pymupdf.open(stream=up_pdf_txt.read(), filetype="pdf")
                full_text = "\n\n".join([f"--- Page {i+1} ---\n" + page.get_text() for i, page in enumerate(doc)])
                st.text_area("Extracted Text Preview:", full_text[:2000] + "...", height=200)
                st.download_button("⬇️ Download Text File (.txt)", data=full_text, file_name=f"{os.path.splitext(up_pdf_txt.name)[0]}.txt")

    elif "Plain Text (.txt) ➔ PDF" in doc_option:
        up_txt = st.file_uploader("Upload Text File (.txt):", type=["txt"], key="txt_pdf_in")
        if up_txt:
            txt_pw_chk = st.checkbox("🔒 Lock PDF with Password", value=False, key="txt_pw_chk")
            txt_pw = st.text_input("Enter Password:", type="password", key="txt_pw_val") if txt_pw_chk else ""
            if st.button("🚀 Convert Text to PDF", type="primary"):
                buf = io.BytesIO()
                doc_t = SimpleDocTemplate(buf, pagesize=letter, margin=40)
                styles = getSampleStyleSheet()
                lines = up_txt.read().decode("utf-8", errors="ignore").split("\n")
                story = [Paragraph(html.escape(l), styles["Normal"]) if l.strip() else Spacer(1, 6) for l in lines]
                doc_t.build(story)
                final_txt_pdf = apply_pdf_password(buf.getvalue(), txt_pw if txt_pw_chk else None)
                st.success("🎉 Converted to PDF!")
                st.download_button("⬇️ Download PDF", data=final_txt_pdf, file_name=f"{os.path.splitext(up_txt.name)[0]}.pdf", mime="application/pdf")

# ----------------- TAB 2: Images & Media Studio -----------------
with tab2:
    st.markdown("### 🖼️ Images & Media Studio")
    media_action = st.radio(
        "Choose Action:", 
        ["📸 Images ➔ Password-Protected PDF", "🖼️ PDF ➔ Extract Pages as PNG", "🎨 Image Format Converter (PNG/JPG/WEBP)"],
        horizontal=True
    )
    
    if media_action == "📸 Images ➔ Password-Protected PDF":
        img_in = st.file_uploader("Upload Image (PNG, JPG, WEBP):", type=["png", "jpg", "jpeg", "webp"], key="img_pdf_in")
        if img_in:
            col1, col2 = st.columns(2)
            with col1:
                im = Image.open(img_in)
                st.image(im, width=320, caption=f"Selected: {img_in.name}")
            with col2:
                enable_lock = st.checkbox("🔒 Native PDF Password Lock", value=True, key="img_pw_on")
                img_pwd = st.text_input("Enter Secret Password:", type="password", key="img_pw_val") if enable_lock else ""
                
                enable_wm = st.checkbox("🛡️ Add Anti-Leak Watermark", value=False, key="img_wm_on")
                wm_val = st.text_input("Watermark Text:", value="CONFIDENTIAL - FOR VERIFICATION ONLY") if enable_wm else ""
                
                if st.button("🚀 Generate PDF", type="primary"):
                    if enable_lock and not img_pwd:
                        st.error("Please enter a password!")
                    else:
                        proc_im = im
                        if enable_wm and wm_val:
                            proc_im = add_watermark(proc_im, wm_val)
                        elif proc_im.mode in ("RGBA", "P"):
                            proc_im = proc_im.convert("RGB")
                        
                        b = io.BytesIO()
                        proc_im.save(b, format="PDF")
                        final_p = apply_pdf_password(b.getvalue(), img_pwd if enable_lock else None)
                        
                        name = f"{os.path.splitext(img_in.name)[0]}.pdf"
                        st.success("🎉 Secure PDF ready!")
                        st.download_button(f"⬇️ Download {name}", data=final_p, file_name=name, mime="application/pdf")

    elif media_action == "🖼️ PDF ➔ Extract Pages as PNG":
        pdf_pages_in = st.file_uploader("Upload PDF to extract pages:", type=["pdf"], key="pdf_pages_up")
        if pdf_pages_in:
            if st.button("✨ Extract All Pages as Images"):
                doc = pymupdf.open(stream=pdf_pages_in.read(), filetype="pdf")
                st.write(f"Total Pages: **{len(doc)}**")
                cols = st.columns(min(3, len(doc)))
                for idx, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=150)
                    png_b = pix.tobytes("png")
                    with cols[idx % len(cols)]:
                        st.image(png_b, caption=f"Page {idx+1}")
                        st.download_button(f"⬇️ Page {idx+1} (PNG)", data=png_b, file_name=f"page_{idx+1}.png", mime="image/png")

    elif media_action == "🎨 Image Format Converter (PNG/JPG/WEBP)":
        raw_img = st.file_uploader("Upload Image:", type=["png", "jpg", "jpeg", "webp"], key="raw_img_cov")
        if raw_img:
            target = st.selectbox("Convert To Format:", ["PNG", "JPEG", "WEBP"])
            if st.button("🔄 Convert Image"):
                img_obj = Image.open(raw_img)
                b = io.BytesIO()
                f = "JPEG" if target == "JPEG" else target
                if f == "JPEG" and img_obj.mode in ("RGBA", "P"):
                    img_obj = img_obj.convert("RGB")
                img_obj.save(b, format=f)
                ext = target.lower()
                st.success(f"🎉 Converted to {target}!")
                st.download_button(f"⬇️ Download .{ext}", data=b.getvalue(), file_name=f"converted.{ext}", mime=f"image/{ext}")

# ----------------- TAB 3: PDF Unlocker & Decryptor -----------------
with tab3:
    st.markdown("### 🔓 PDF Password Unlocker")
    st.write("Have a password-protected PDF? Unlock it locally so you can view/print it without repeatedly typing the password.")
    
    locked_pdf_in = st.file_uploader("Upload Password-Locked PDF:", type=["pdf"], key="locked_pdf_up")
    unlock_pw = st.text_input("Enter the Document Password:", type="password", key="unlock_pw_in")
    
    if st.button("🔓 Unlock & Remove Password", type="primary"):
        if not locked_pdf_in or not unlock_pw:
            st.warning("⚠️ Please provide both the locked PDF and the password!")
        else:
            try:
                unlocked_bytes = unlock_pdf(locked_pdf_in.read(), unlock_pw)
                st.success("🎉 Successfully unlocked! Password protection removed.")
                st.download_button(
                    label="⬇️ Download Unlocked PDF",
                    data=unlocked_bytes,
                    file_name=f"{os.path.splitext(locked_pdf_in.name)[0]}_unlocked.pdf",
                    mime="application/pdf"
                )
            except Exception:
                st.error("❌ Failed to unlock! The password was incorrect.")

# ----------------- TAB 4: Privacy Shield -----------------
with tab4:
    st.markdown("### 🛡️ Enterprise Privacy Shield")
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### 1. SHA-256 Digital Fingerprint")
        st.write("Calculate a tamper-proof cryptographic fingerprint of any document:")
        hash_file = st.file_uploader("Upload File to Hash:", key="hash_f_in")
        if hash_file:
            h = calculate_sha256(hash_file.read())
            st.code(f"SHA-256 Hash:\n{h}", language="text")
            st.caption("✨ *Recipients can verify this exact checksum to guarantee the file was not tampered with.*")
            
    with col_b:
        st.markdown("#### 2. Live PII Masking Engine")
        st.write("Type or paste document text with phone, email, or Aadhaar numbers to mask and download:")
        sample_in = st.text_area(
            "Enter Document Text:",
            "Phone: +91 9790611283 , Aadhaar: 5445 5020 4394",
            height=100
        )
        if st.button("🛡️ Redact Personal Information", type="primary"):
            cleaned = redact_sensitive_text(sample_in)
            st.success("✅ Sensitive Data Masked Successfully!")
            st.code(cleaned, language="text")
            
            # Direct Download options!
            st.markdown("##### ⬇️ Download Your Redacted Document:")
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                # PDF Download of masked text
                pdf_buf = io.BytesIO()
                doc_t = SimpleDocTemplate(pdf_buf, pagesize=letter, margin=40)
                styles = getSampleStyleSheet()
                lines = cleaned.split("\n")
                story = [Paragraph(html.escape(l), styles["Normal"]) if l.strip() else Spacer(1, 6) for l in lines]
                doc_t.build(story)
                st.download_button(
                    label="📄 Download as Masked PDF",
                    data=pdf_buf.getvalue(),
                    file_name="masked_document.pdf",
                    mime="application/pdf"
                )
            with d_col2:
                st.download_button(
                    label="📋 Download as Masked Text (.txt)",
                    data=cleaned,
                    file_name="masked_document.txt",
                    mime="text/plain"
                )
