import streamlit as st
from docxtpl import DocxTemplate
import tempfile
import datetime
import os
import smtplib
from email.message import EmailMessage

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


# --- Send email via Gmail SMTP ---
def send_email_with_gmail(full_name, university, grad_class, cwa, docx_path):
    """Send the generated letter to your Gmail using SMTP."""
    sender = st.secrets["SMTP_EMAIL"]
    password = st.secrets["SMTP_PASS"]
    recipient = st.secrets["SMTP_EMAIL"]  # send to yourself

    # Create the email
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = f"Recommendation Letter: {full_name} ({university})"  # full name used here
    msg.set_content(
        f"""Dear Dr. Kwegyir,

A new recommendation letter has been generated for {full_name} in support of their application to {university}.
Details of the student are as follows;

Student: {full_name}
University: {university}
Graduating Class: {grad_class}
CWA: {cwa}

The Word document is attached.

Regards,
Automated Recommendation Letter System
"""
    )

    # Attach the file
    with open(docx_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(docx_path)
        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file_name,
        )

    # Send via Gmail SMTP
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(sender, password)
            smtp.send_message(msg)
        st.success(f"✅ Your recommendation letter request has been sent successfully to the Lecturer.")
    except Exception as e:
        st.error(f"❌ Email sending failed: {e}")


# --- Streamlit UI ---
st.set_page_config(page_title="Graduate Recommendation Letter Submission", layout="wide")
st.title("Graduate School Recommendation Letter Request Submission Form")

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
        cwa = st.text_input("Cumulative Weighted Average (CWA)", placeholder="e.g., 78.5%")
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

                # Select the appropriate template
                template_file = "Male.docx" if gender == "Male" else "Female.docx"

                if not os.path.exists(template_file):
                    st.error(f"❌ Template file '{template_file}' not found.")
                else:
                    doc = generate_letter(template_file, context)
                    docx_path = save_docx_only(doc, full_name, university)
                    send_email_with_gmail(full_name, university, grad_class, cwa, docx_path)

            except Exception as e:
                st.error(f"❌ Unexpected error: {e}")
