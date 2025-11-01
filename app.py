import streamlit as st
from docxtpl import DocxTemplate
import pypandoc
import tempfile
import datetime
import os

# --- Ensure Pandoc is available ---
try:
    pypandoc.get_pandoc_path()
except OSError:
    st.warning("🔧 Pandoc not found. Downloading now... (first run only)")
    pypandoc.download_pandoc()


# --- Function to fill the template ---
def generate_letter(template_path, context):
    """Render the DOCX template using docxtpl with all placeholders replaced."""
    doc = DocxTemplate(template_path)
    doc.render(context)
    return doc


# --- Function to save and convert DOCX → PDF ---
def save_and_convert_to_pdf(doc, student_name, university_name):
    """Saves the generated DOCX and converts it to PDF (works on Streamlit Cloud)."""
    temp_dir = tempfile.mkdtemp()

    # Create safe filenames
    safe_student = student_name.replace(" ", "_").replace("/", "_")
    safe_university = university_name.replace(" ", "_").replace("/", "_")
    file_basename = f"{safe_student}_{safe_university}"

    docx_path = os.path.join(temp_dir, f"{file_basename}.docx")
    pdf_path = os.path.join(temp_dir, f"{file_basename}.pdf")

    # Save the filled DOCX
    doc.save(docx_path)

    # Ensure pandoc is installed (in case of rebuild)
    try:
        pypandoc.get_pandoc_path()
    except OSError:
        pypandoc.download_pandoc()

    # Try to convert DOCX → PDF
    try:
        pypandoc.convert_file(
            docx_path,
            "pdf",
            outputfile=pdf_path,
            extra_args=["--standalone", "--pdf-engine=xelatex"],
        )
    except Exception as e:
        st.warning(f"⚠️ PDF conversion failed: {e}")
        st.info("The Word document has been generated successfully.")
        pdf_path = None

    return docx_path, pdf_path


# --- Streamlit App ---
st.set_page_config(page_title="Graduate Recommendation Letter Generator", layout="wide")
st.title("🎓 Graduate School Recommendation Letter Generator")
st.markdown(
    "Automatically generate recommendation letters with your official letterhead and signature. "
    "Ensure your templates (`Male.docx` and `Female.docx`) use Jinja2 placeholders like `{{ Name }}`."
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
    required_fields = [full_name, gender, university, project_topic, grad_class, cwa, year]
    if not all(required_fields):
        st.warning("⚠️ Please fill in all fields before generating the letter.")
    else:
        with st.spinner("⏳ Generating letter... please wait."):
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
                    st.error(
                        f"❌ Template file '{template_file}' not found. "
                        "Please upload or place it in the same folder as this script."
                    )
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
