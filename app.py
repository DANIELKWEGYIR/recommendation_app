import streamlit as st
from docxtpl import DocxTemplate
import tempfile
import datetime
import os
import requests
import base64

# --- Helper: generate letter ---
def generate_letter(template_path, context):
    doc = DocxTemplate(template_path)
    doc.render(context)
    return doc


# --- Save DOCX only ---
def save_docx_only(doc, student_name, university_name):
    temp_dir = tempfile.mkdtemp()
    safe_student = student_name.replace(" ", "_").replace("/", "_")
    safe_univ = university_name.replace(" ", "_").replace("/", "_")
    base = f"{safe_student}_{safe_univ}"
    docx_path = os.path.join(temp_dir, f"{base}.docx")
    doc.save(docx_path)
    return docx_path


# --- Send email with EmailJS ---
def send_email_with_emailjs(student_name, university, grad_class, cwa, docx_path):
    """Send the generated letter to lecturer's email via EmailJS."""
    try:
        with open(docx_path, "rb") as f:
            encoded_file = base64.b64encode(f.read()).decode()

        payload = {
            "service_id": st.secrets["EMAILJS_SERVICE_ID"],
            "template_id": st.secrets["EMAILJS_TEMPLATE_ID"],
            "user_id": st.secrets["EMAILJS_PUBLIC_KEY"],
            "template_params": {
                "student_name": student_name,
                "university": university,
                "grad_class": grad_class,
                "cwa": cwa,
                "to_email": st.secrets["MY_EMAIL"],  # your email
            },
            "attachments": [
                {
                    "name": os.path.basename(docx_path),
                    "data": f"data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{encoded_file}",
                }
            ],
        }

        response = requests.post(
            "https://api.emailjs.com/api/v1.0/email/send",
            json=payload,
            timeout=30,
        )

        if response.status_code == 200:
            st.success("✅ Recommendation letter has been sent to the lecturer's email.")
        else:
            st.warning(f"⚠️ Email sending failed: {response.text}")

    except Exception as e:
        st.error(f"❌ Could not send email: {e}")


# --- Streamlit UI ---
st.set_page_config(page_title="Graduate Recommendation Letter Submission", layout="wide")
st.title("🎓 Graduate School Recommendation Letter Submission Form")

st.markdown(
    """
    Students should complete this form accurately.  
    Once submitted, the system automatically generates the recommendation letter  
    and sends it securely to the lecturer’s official email address.
    """
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
        year = st.text_input("Year Lecturer Began Teaching You", placeholder="e.g., 2021")

    submitted = st.form_submit_button("Submit and Send Letter")

# --- Processing ---
if submitted:
    required = [full_name, gender, university, project_topic, grad_class, cwa, year]
    if not all(required):
        st.warning("⚠️ Please fill in all fields before submitting.")
    else:
        with st.spinner("⏳ Generating and sending letter..."):
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

                # Select appropriate template
                template_file = "Male.docx" if gender == "Male" else "Female.docx"

                if not os.path.exists(template_file):
                    st.error(f"❌ Template file '{template_file}' not found.")
                else:
                    # Generate and email the letter
                    doc = generate_letter(template_file, context)
                    docx_path = save_docx_only(doc, full_name, university)
                    send_email_with_emailjs(full_name, university, grad_class, cwa, docx_path)

            except Exception as e:
                st.error(f"❌ Unexpected error: {e}")
