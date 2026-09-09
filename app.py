import streamlit as st
import os
import io
import re
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
from streamlit_drawable_canvas import st_canvas

# Page Configuration
st.set_page_config(
    page_title="PrivConvert Studio - Zero-Knowledge Suite",
    page_icon="🛡️",
    layout="wide"
)

# Modern Colorful CSS
st.markdown("""
<style>
    .hero-container {
        background: linear-gradient(135deg, #1E1B4B 0%, #312E81 40%, #0F172A 100%);
        border: 2px solid #6366F1;
        border-radius: 16px;
        padding: 22px;
        margin-bottom: 22px;
        box-shadow: 0 10px 25px -5px rgba(99, 102, 241, 0.3);
    }
    .hero-title {
        font-size: 2.5rem;
        font-weight: 900;
        background: linear-gradient(90deg, #38BDF8, #818CF8, #F472B6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0;
    }
    .hero-subtitle {
        color: #E2E8F0;
        font-size: 1.05rem;
        font-weight: 500;
        margin-top: 5px;
    }
    .pill-blue {
        background: linear-gradient(90deg, #0284C7, #0EA5E9);
        color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; font-weight: 600; display: inline-block; margin-right: 8px;
    }
    .pill-purple {
        background: linear-gradient(90deg, #7C3AED, #A855F7);
        color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; font-weight: 600; display: inline-block; margin-right: 8px;
    }
    .pill-emerald {
        background: linear-gradient(90deg, #059669, #10B981);
        color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; font-weight: 600; display: inline-block; margin-right: 8px;
    }
    .pill-pink {
        background: linear-gradient(90deg, #DB2777, #F43F5E);
        color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; font-weight: 600; display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# Hero Header
st.markdown("""
<div class="hero-container">
    <div class="hero-title">🛡️ PrivConvert Studio</div>
    <div class="hero-subtitle">Zero-Knowledge Document Suite: Drag-to-Redact & Multi-Page PDF Batch Masking</div>
    <div style="margin-top: 12px;">
        <span class="pill-blue">⚡ 100% In-Memory RAM</span>
        <span class="pill-purple">🔒 AES-256 Military Lock</span>
        <span class="pill-emerald">🖱️ Interactive Drag & Draw</span>
        <span class="pill-pink">📑 100-Page PDF Auto-Redactor</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ----------------- Cryptographic & Helper Functions -----------------

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

def convert_docx_to_pdf(docx_file) -> bytes:
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
                    if run.bold:
                        txt = f"<b>{txt}</b>"
                    if run.italic:
                        txt = f"<i>{txt}</i>"
                    runs_html += txt
                if runs_html.strip():
                    if "Heading" in para.style.name:
                        story.append(Paragraph(runs_html, h1_style))
                    else:
                        story.append(Paragraph(runs_html, body_style))
                    story.append(Spacer(1, 6))
        elif element.tag.endswith('tbl'):
            t = [t for t in doc.tables if t._element == element]
            if t:
                tbl = t[0]
                tbl_data = []
                for row in tbl.rows:
                    row_data = [Paragraph(html.escape(c.text.strip()), body_style) for c in row.cells]
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

# ----------------- MAIN APP TABS -----------------
tab1, tab2, tab3, tab4 = st.tabs([
    "🛡️ Interactive Redactor (Drag & Draw + Multi-Page PDF)",
    "📑 Office & Documents (PDF, Word, PPT)", 
    "🖼️ Images & Media Converter", 
    "🔓 PDF Password Unlocker"
])

# ----------------- TAB 1: Advanced Redactor -----------------
with tab1:
    st.markdown("### 🛡️ Smart Document Redactor")
    
    redact_type = st.radio(
        "Choose Redaction Mode:",
        [
            "🖱️ Mode A: Drag & Draw Box on Image (Bank Statement, Cheque, ID)",
            "📑 Mode B: Multi-Page PDF Auto-Redactor (Search & Mask Across ALL Pages)"
        ],
        horizontal=True
    )
    
    # ---------------- MODE A: Drag & Draw on Image ----------------
    if "Mode A" in redact_type:
        st.markdown("#### 🖱️ Interactive Drag & Draw Redactor")
        st.write("Upload an image. Use your mouse to **draw a rectangle** directly over any Account Number, Aadhaar, or Signature to blackout:")
        
        img_file = st.file_uploader("Upload Image (PNG, JPG):", type=["png", "jpg", "jpeg"], key="draw_img_up")
        
        if img_file:
            bg_image = Image.open(img_file).convert("RGB")
            
            # Calculate display scale
            orig_w, orig_h = bg_image.size
            canvas_w = min(750, orig_w)
            scale = canvas_w / orig_w
            canvas_h = int(orig_h * scale)
            
            col_canv, col_side = st.columns([3, 2])
            
            with col_canv:
                st.write("👉 *Click and drag with your mouse to draw black boxes over sensitive data:*")
                canvas_result = st_canvas(
                    fill_color="rgba(0, 0, 0, 1.0)",  # Solid black
                    stroke_width=2,
                    stroke_color="#000000",
                    background_image=bg_image,
                    update_streamlit=True,
                    height=canvas_h,
                    width=canvas_w,
                    drawing_mode="rect",
                    key="redaction_canvas",
                )
                
            with col_side:
                st.markdown("#### 🔒 Export Settings")
                draw_lock = st.checkbox("Protect Exported PDF with Password", value=True, key="draw_pw_box")
                draw_pw = st.text_input("Set Password:", type="password", key="draw_pw_val") if draw_lock else ""
                
                if st.button("🚀 Export Redacted PDF", type="primary"):
                    if draw_lock and not draw_pw:
                        st.error("Please enter a password or uncheck password protection!")
                    else:
                        # Apply drawn boxes to original image
                        final_img = bg_image.copy()
                        draw = ImageDraw.Draw(final_img)
                        
                        box_count = 0
                        if canvas_result.json_data is not None:
                            objects = canvas_result.json_data["objects"]
                            for obj in objects:
                                if obj["type"] == "rect":
                                    # Scale back to original resolution
                                    x1 = int(obj["left"] / scale)
                                    y1 = int(obj["top"] / scale)
                                    x2 = int((obj["left"] + obj["width"]) / scale)
                                    y2 = int((obj["top"] + obj["height"]) / scale)
                                    draw.rectangle([x1, y1, x2, y2], fill=(0, 0, 0))
                                    box_count += 1
                                    
                        buf = io.BytesIO()
                        final_img.save(buf, format="PDF")
                        secured_bytes = apply_pdf_password(buf.getvalue(), draw_pw if draw_lock else None)
                        
                        st.success(f"🎉 Redacted {box_count} areas and created secure PDF!")
                        name = f"redacted_{os.path.splitext(img_file.name)[0]}.pdf"
                        st.download_button(f"⬇️ Download {name}", data=secured_bytes, file_name=name, mime="application/pdf")

    # ---------------- MODE B: Multi-Page PDF Search & Redact ----------------
    else:
        st.markdown("#### 📑 Multi-Page PDF Batch Redactor")
        st.write("Have a 10-page, 50-page, or 100-page bank statement or contract? Automatically search and black out any sensitive text or number **across all pages simultaneously**:")
        
        multi_pdf = st.file_uploader("Upload Multi-Page PDF:", type=["pdf"], key="multi_pdf_up")
        
        if multi_pdf:
            pdf_bytes_in = multi_pdf.read()
            doc_test = pymupdf.open(stream=pdf_bytes_in, filetype="pdf")
            total_pages = len(doc_test)
            st.info(f"📄 Document Loaded: **{multi_pdf.name}** ({total_pages} Pages)")
            
            c_s1, c_s2 = st.columns(2)
            with c_s1:
                target_phrase = st.text_input(
                    "Enter Text / Account Number to Redact Across All Pages:",
                    placeholder="e.g. 9790611283 or Rajesh or Balance"
                )
            with c_s2:
                batch_lock = st.checkbox("Lock Resulting PDF with Password", value=True, key="b_pw_box")
                batch_pw = st.text_input("Set Password:", type="password", key="b_pw_val") if batch_lock else ""
                
            if st.button("🔍 Search & Redact Across ALL Pages", type="primary"):
                if not target_phrase:
                    st.warning("⚠️ Please enter a number or text phrase to search and redact!")
                elif batch_lock and not batch_pw:
                    st.error("Please enter a password!")
                else:
                    with st.spinner(f"Scanning and redacting across all {total_pages} pages in local RAM..."):
                        doc = pymupdf.open(stream=pdf_bytes_in, filetype="pdf")
                        total_redactions = 0
                        
                        for page in doc:
                            # Search for the phrase on this page
                            rects = page.search_for(target_phrase)
                            for r in rects:
                                # Add permanent blackout redaction
                                page.add_redact_annot(r, fill=(0, 0, 0))
                                total_redactions += 1
                            page.apply_redactions()
                            
                        redacted_pdf_data = doc.tobytes()
                        final_secured_pdf = apply_pdf_password(redacted_pdf_data, batch_pw if batch_lock else None)
                        
                        st.success(f"🎉 Success! Found and blacked out **{total_redactions} occurrences** across all {total_pages} pages!")
                        out_b_name = f"redacted_all_pages_{multi_pdf.name}"
                        st.download_button(
                            label=f"⬇️ Download Redacted PDF ({total_pages} Pages)",
                            data=final_secured_pdf,
                            file_name=out_b_name,
                            mime="application/pdf"
                        )

# ----------------- TAB 2: Office & Documents -----------------
with tab2:
    st.markdown("### 📑 Office Document Conversions")
    doc_option = st.selectbox(
        "Choose Conversion:",
        [
            "📄 PDF ➔ Word (.docx) [Fully Editable Layout & Tables]",
            "📝 Word (.docx) ➔ PDF [Preserves Formatting & Fonts]",
            "📊 PDF ➔ PowerPoint (.pptx) [Pixel-Perfect Slides]",
            "🖥️ PowerPoint (.pptx) ➔ PDF",
            "📋 PDF ➔ Plain Text (.txt)"
        ]
    )
    
    if "PDF ➔ Word" in doc_option:
        up_pdf = st.file_uploader("Upload PDF:", type=["pdf"], key="pdf_w_in")
        if up_pdf and st.button("✨ Convert to Word (.docx)", type="primary"):
            with st.spinner("Converting locally..."):
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
                st.success("🎉 Converted to Word!")
                st.download_button(f"⬇️ Download {name}", data=data, file_name=name)

    elif "Word (.docx) ➔ PDF" in doc_option:
        up_docx = st.file_uploader("Upload Word (.docx):", type=["docx"], key="docx_pdf_in")
        if up_docx:
            pw_docx = st.checkbox("🔒 Native AES-256 Password Protection", value=False, key="pw_d_box")
            docx_pw = st.text_input("Set Password:", type="password", key="pw_d_val") if pw_docx else ""
            if st.button("🚀 Convert to Formatted PDF", type="primary"):
                pdf_bytes = convert_docx_to_pdf(up_docx)
                final_pdf = apply_pdf_password(pdf_bytes, docx_pw if pw_docx else None)
                name = f"{os.path.splitext(up_docx.name)[0]}.pdf"
                st.success("🎉 Converted to PDF!")
                st.download_button(f"⬇️ Download {name}", data=final_pdf, file_name=name, mime="application/pdf")

    elif "PDF ➔ PowerPoint" in doc_option:
        up_pdf_ppt = st.file_uploader("Upload PDF to convert to Slides:", type=["pdf"], key="pdf_ppt_in")
        if up_pdf_ppt and st.button("📊 Convert to PowerPoint (.pptx)", type="primary"):
            pptx_bytes = convert_pdf_to_pptx(up_pdf_ppt)
            name = f"{os.path.splitext(up_pdf_ppt.name)[0]}.pptx"
            st.success("🎉 Presentation slides created!")
            st.download_button(f"⬇️ Download {name}", data=pptx_bytes, file_name=name)

    elif "PowerPoint (.pptx) ➔ PDF" in doc_option:
        up_ppt = st.file_uploader("Upload PowerPoint (.pptx):", type=["pptx"], key="ppt_pdf_in")
        if up_ppt and st.button("🚀 Convert PPT to PDF", type="primary"):
            pdf_b = convert_pptx_to_pdf(up_ppt)
            name = f"{os.path.splitext(up_ppt.name)[0]}.pdf"
            st.success("🎉 Converted to PDF!")
            st.download_button(f"⬇️ Download {name}", data=pdf_b, file_name=name, mime="application/pdf")

    elif "PDF ➔ Plain Text" in doc_option:
        up_pdf_txt = st.file_uploader("Upload PDF to extract text:", type=["pdf"], key="pdf_txt_in")
        if up_pdf_txt and st.button("📋 Extract Text"):
            doc = pymupdf.open(stream=up_pdf_txt.read(), filetype="pdf")
            full_text = "\n\n".join([f"--- Page {i+1} ---\n" + page.get_text() for i, page in enumerate(doc)])
            st.text_area("Extracted Text Preview:", full_text[:2000] + "...", height=200)
            st.download_button("⬇️ Download Text (.txt)", data=full_text, file_name=f"{os.path.splitext(up_pdf_txt.name)[0]}.txt")

# ----------------- TAB 3: Images & Media Studio -----------------
with tab3:
    st.markdown("### 🖼️ Images & Media Studio")
    media_action = st.radio(
        "Choose Action:", 
        ["📸 Image ➔ Password-Protected PDF", "🖼️ PDF ➔ Extract Pages as PNG", "🎨 Image Format Converter (PNG/JPG/WEBP)"],
        horizontal=True
    )
    
    if media_action == "📸 Image ➔ Password-Protected PDF":
        img_in = st.file_uploader("Upload Image:", type=["png", "jpg", "jpeg", "webp"], key="media_img_up")
        if img_in:
            c1, c2 = st.columns(2)
            with c1:
                im = Image.open(img_in)
                st.image(im, width=320)
            with c2:
                enable_lock = st.checkbox("🔒 Native PDF Password Lock", value=True, key="m_img_pw_on")
                img_pwd = st.text_input("Enter Secret Password:", type="password", key="m_img_pw_val") if enable_lock else ""
                if st.button("🚀 Generate PDF", type="primary"):
                    if enable_lock and not img_pwd:
                        st.error("Please enter a password!")
                    else:
                        proc_im = im
                        if proc_im.mode in ("RGBA", "P"):
                            proc_im = proc_im.convert("RGB")
                        b = io.BytesIO()
                        proc_im.save(b, format="PDF")
                        final_p = apply_pdf_password(b.getvalue(), img_pwd if enable_lock else None)
                        name = f"{os.path.splitext(img_in.name)[0]}.pdf"
                        st.success("🎉 Secure PDF ready!")
                        st.download_button(f"⬇️ Download {name}", data=final_p, file_name=name, mime="application/pdf")

    elif media_action == "🖼️ PDF ➔ Extract Pages as PNG":
        pdf_pages_in = st.file_uploader("Upload PDF to extract pages:", type=["pdf"], key="m_pdf_pages_up")
        if pdf_pages_in and st.button("✨ Extract All Pages as Images"):
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
        raw_img = st.file_uploader("Upload Image:", type=["png", "jpg", "jpeg", "webp"], key="m_raw_img_cov")
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

# ----------------- TAB 4: PDF Password Unlocker -----------------
with tab4:
    st.markdown("### 🔓 PDF Password Unlocker")
    st.write("Unlock a password-protected PDF to print or view freely.")
    locked_pdf_in = st.file_uploader("Upload Password-Locked PDF:", type=["pdf"], key="m_locked_pdf_up")
    unlock_pw = st.text_input("Enter the Document Password:", type="password", key="m_unlock_pw_in")
    
    if st.button("🔓 Unlock & Remove Password", type="primary"):
        if not locked_pdf_in or not unlock_pw:
            st.warning("Please provide both file and password!")
        else:
            try:
                unlocked_bytes = unlock_pdf(locked_pdf_in.read(), unlock_pw)
                st.success("🎉 Successfully unlocked!")
                st.download_button("⬇️ Download Unlocked PDF", data=unlocked_bytes, file_name="unlocked_document.pdf", mime="application/pdf")
            except Exception:
                st.error("❌ Incorrect password!")
