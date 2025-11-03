import streamlit as st
from docxtpl import DocxTemplate
import tempfile
import datetime
import os

# --- Template rendering ---
def generate_letter(template_path, context):
    """Render the Word (.docx) template with the given context."""
    doc = DocxTemplate(template_path)
    doc.render(context)
    return doc


# --- Save only DOCX ---
def save_docx_only(doc, student_name, university_name):
    """Saves the rendered DOCX file (no PDF conversion)."""
    temp_dir = tempfile.mkdtemp()
    safe_student = student_name.replace(" ", "_").replace("/", "_")
    safe_univ = university_name.replace(" ", "_").replace("/", "_")
    base = f"{safe_student}_{safe_univ}"

    docx_path = os.path.join(temp_dir, f"{base}.docx")
    doc.save(docx_path)
    return docx_path


# --- Streamlit App ---
st.title("🎓 Graduate School Recommendation Letter")

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

                # Context dictionary for the Word template
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

                # Select the appropriate template
                template_file = "Male.docx" if gender == "Male" else "Female.docx"

                # Validate template
                if not os.path.exists(template_file):
                    st.error(f"❌ Template file '{template_file}' not found. Please upload it to the app directory.")
                else:
                    # Generate and save DOCX
                    doc = generate_letter(template_file, context)
                    docx_path = save_docx_only(doc, full_name, university)

                    st.success("✅ Letter generated successfully!")

                    st.markdown("### 📥 Download Your Letter")
                    with open(docx_path, "rb") as f_docx:
                        st.download_button(
                            label="⬇️ Download Word File (.docx)",
                            data=f_docx,
                            file_name=os.path.basename(docx_path),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )

            except Exception as e:
                st.error(f"❌ An unexpected error occurred: {e}")
