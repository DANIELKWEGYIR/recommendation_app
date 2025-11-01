import streamlit as st
from docxtpl import DocxTemplate # Import DocxTemplate
from docx2pdf import convert
import tempfile
import datetime
import os

# Function to replace placeholders in the docx template
def generate_letter(template_path, context):
    """
    Replaces placeholders in a .docx template using docxtpl.
    This method correctly preserves all document formatting.
    """
    # Load the template
    doc = DocxTemplate(template_path)
    
    # Render the context into the template
    doc.render(context)
    
    # Return the rendered document object
    return doc

import pypandoc
def save_and_convert_to_pdf(doc, student_name, university_name):
    """Saves the generated docx and converts it to PDF (cloud-safe)."""
    temp_dir = tempfile.mkdtemp()
    
    safe_student_name = student_name.replace(" ", "_").replace("/", "_")
    safe_university_name = university_name.replace(" ", "_").replace("/", "_")
    file_basename = f"{safe_student_name}_{safe_university_name}"

    docx_path = os.path.join(temp_dir, f"{file_basename}.docx")
    pdf_path = os.path.join(temp_dir, f"{file_basename}.pdf")

    # Save DOCX
    doc.save(docx_path)

    # Convert DOCX → PDF using pypandoc (works on Linux)
    try:
        pypandoc.convert_file(docx_path, "pdf", outputfile=pdf_path, extra_args=["--standalone"])
    except Exception as e:
        st.warning(f"PDF conversion failed: {e}")
        pdf_path = None

    return docx_path, pdf_path


# Streamlit App UI
st.set_page_config(layout="wide")
st.title("🎓 Graduate School Recommendation Letter")
#st.markdown("This app automatically generates recommendation letters with your official letterhead and signature.")
# Updated st.info to include instructions for the new {{ Name_Upper }} placeholder


# Main form
with st.form("recommendation_form"):
    st.subheader("Student Details")
    
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

    # Submit button for the form
    submitted = st.form_submit_button("Generate Letter")

# Processing logic after form submission
if submitted:
    # Check for required fields
    required_fields = [full_name, gender, university, project_topic, grad_class, cwa, year]
    if not all(required_fields):
        st.warning("Please fill in all fields before generating the letter.")
    else:
        with st.spinner("Generating letters... This may take a moment."):
            try:
                # Prepare context for docxtpl (no brackets, no spaces)
                current_date = datetime.date.today().strftime("%B %d, %Y")
                context = {
                    "Name": full_name,
                    "Name_Upper": full_name.upper(), # New key for capitalized name in heading
                    "University_Applying_To": university,
                    "Project_Topic": project_topic,
                    "Graduating_Class": grad_class,
                    "CWA": cwa,
                    "Year": year,
                    "Date": current_date
                }

                # Select template
                template_file = "Male.docx" if gender == "Male" else "Female.docx"
                
                # Check if template files exist
                if not os.path.exists(template_file):
                    st.error(f"Error: Template file '{template_file}' not found. Please make sure it's in the same directory as the script.")
                else:
                    # Generate and save
                    doc = generate_letter(template_file, context)
                    # Pass the university name to the save function
                    docx_path, pdf_path = save_and_convert_to_pdf(doc, full_name, university) 

                    st.success("✅ Letter generated successfully!")
                    
                    # Provide download buttons
                    col1_dl, col2_dl = st.columns(2)
                    with col1_dl:
                        with open(docx_path, "rb") as f_docx:
                            st.download_button(
                                label="⬇️ Download Word File (.docx)",
                                data=f_docx,
                                file_name=os.path.basename(docx_path),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    
                    if pdf_path: # Only show PDF button if conversion succeeded
                        with col2_dl:
                            with open(pdf_path, "rb") as f_pdf:
                                st.download_button(
                                    label="⬇️ Download PDF File (.pdf)",
                                    data=f_pdf,
                                    file_name=os.path.basename(pdf_path),
                                    mime="application/pdf"
                                )
                                
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")

