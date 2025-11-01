import streamlit as st
from docxtpl import DocxTemplate
import pypandoc
import tempfile
import datetime
import os
import zipfile
import mimetypes

# --- Ensure Pandoc is available ---
try:
    pypandoc.get_pandoc_path()
except OSError:
    st.warning("🔧 Pandoc not found. Downloading now... (first run only)")
    pypandoc.download_pandoc()


# --- Helper: list and extract media from DOCX ---
def list_docx_media(docx_path):
    media = []
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            for name in z.namelist():
                if name.startswith("word/media/"):
                    mime_type, _ = mimetypes.guess_type(name)
                    media.append((name, mime_type or "application/octet-stream"))
    except Exception as e:
        st.warning(f"Could not inspect .docx media: {e}")
    return media


def extract_docx_media(docx_path, out_dir):
    extracted = []
    with zipfile.ZipFile(docx_path, 'r') as z:
        for member in z.namelist():
            if member.startswith("word/media/"):
                target_path = os.path.join(out_dir, os.path.basename(member))
                with open(target_path, "wb") as f:
                    f.write(z.read(member))
                extracted.append(target_path)
    return extracted


# --- Template rendering ---
def generate_letter(template_path, context):
    doc = DocxTemplate(template_path)
    doc.render(context)
    return doc


# --- Save DOCX and attempt PDF conversion ---
def save_and_convert_to_pdf(doc, student_name, university_name):
    temp_dir = tempfile.mkdtemp()
    safe_student = student_name.replace(" ", "_").replace("/", "_")
    safe_univ = university_name.replace(" ", "_").replace("/", "_")
    base = f"{safe_student}_{safe_univ}"

    docx_path = os.path.join(temp_dir, f"{base}.docx")
    pdf_path = os.path.join(temp_dir, f"{base}.pdf")

    # Save DOCX
    doc.save(docx_path)

    # Ensure Pandoc is present
    try:
        pypandoc.get_pandoc_path()
    except OSError:
        pypandoc.download_pandoc()

    try:
        # Try PDF conversion (no fixed engine)
        pypandoc.convert_file(docx_path, "pdf", outputfile=pdf_path, extra_args=["--standalone"])
        return docx_path, pdf_path
    except Exception as e:
        st.warning(f"⚠️ PDF conversion failed: {e}")
        media_list = list_docx_media(docx_path)
        emf_files = [m for m, t in media_list if m.lower().endswith(".emf") or (t and "emf" in (t.lower() or ""))]

        if emf_files:
            st.error("🚫 Your Word template contains EMF images that Pandoc cannot convert on Linux.")
            st.info("➡️ Please replace EMF images with PNG or JPG in your Word templates (letterhead/signature).")
            # Extract media for inspection
            media_out = os.path.join(temp_dir, "extracted_media")
            os.makedirs(media_out, exist_ok=True)
            extracted = extract_docx_media(docx_path, media_out)
            st.markdown("### 🔎 Extracted media files:")
            for p in extracted:
                fname = os.path.basename(p)
                try:
                    with open(p, "rb") as mf:
                        st.download_button(label=f"Download {fname}", data=mf.read(), file_name=fname)
                except Exception:
                    st.write(f"- {fname} (unable to show preview)")
        else:
            st.info("No EMF images found. Conversion likely failed due to missing TeX engine or unsupported image type.")
            st.info("➡️ Replace any SVG/WMF images with PNG/JPG in your templates.")

        return docx_path, None


# --- Streamlit App ---
st.set_page_config(page_title="Graduate Recommendation Letter Generator", layout="wide")
st.title("🎓 Graduate School Recommendation Letter Generator")

st.markdown(
    "Automatically generate recommendation letters using official templates. "
    "Ensure your `Male.docx` and `Female.docx` use Jinja2 placeholders such as `{{ Name }}`, `{{ Date }}`, etc."
)

# --- Input Form ---
with st.form("recommendation_form"):
    st.subheader("🧾 Student Details")

    col1, col2 = st.columns(2)
    with col1:
        full_name = st.text_input("Full Name of Student", placeholder="e.g., Jane Doe")
        gender = st.selectbox("Gender", ["Male", "Female"])
        university = st.text_input("University Applying To", placeholder="e.g., Stanford University")

    with col2:
        project_topic = st.text_input("Final Year Project Topic", placeholder="e.g., AI in Renewable Energy")
        grad_class = st.text_input("Graduating Class", placeholder="e.g., First Class Honours")
        cwa = st.text_input("Cumulative Weighted Average (CWA)", placeholder="e.g., 78.5")
        year = st.text_input("Year You Began Teaching the Student", placeholder="e.g., 2021")

    submitted = st.form_submit_button("Generate Letter")

# --- Processing ---
if submitted:
    required = [full_name, gender, university, project_topic, grad_class, cwa, year]
    if not all(required):
        st.warning("⚠️ Please fill in all fields before generating the letter.")
    else:
        with st.spinner("⏳ Generating letter..."):
            try:
                current_date = datetime.date.today().strftime("%B %d, %Y")

                context = {
                    "Name": full_name,
                    "Name_Upper": full_name.upper(),
                    "University_Applying_To": university,
                    "Project_Topic": project_topic,
                    "Graduating_Class": grad_class,
                    "CWA": cwa,
                    "Year": year,
                    "Date": current_date,
                }

                template_file = "Male.docx" if gender == "Male" else "Female.docx"

                if not os.path.exists(template_file):
                    st.error(f"❌ Template file '{template_file}' not found. Please upload it to the app directory.")
                else:
                    doc = generate_letter(template_file, context)
                    docx_path, pdf_path = save_and_convert_to_pdf(doc, full_name, university)

                    st.success("✅ Letter generated successfully!")

                    st.markdown("### 📥 Download Your Files")
                    col1_dl, col2_dl = st.columns(2)

                    with col1_dl:
                        with open(docx_path, "rb") as f_docx:
                            st.download_button(
                                label="⬇️ Download Word File (.docx)",
                                data=f_docx,
                                file_name=os.path.basename(docx_path),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )

                    if pdf_path:
                        with col2_dl:
                            with open(pdf_path, "rb") as f_pdf:
                                st.download_button(
                                    label="⬇️ Download PDF File (.pdf)",
                                    data=f_pdf,
                                    file_name=os.path.basename(pdf_path),
                                    mime="application/pdf",
                                )

            except Exception as e:
                st.error(f"❌ An unexpected error occurred: {e}")
