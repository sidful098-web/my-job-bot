import streamlit as st
import pandas as pd
from docx import Document
import google.generativeai as genai
import smtplib
from email.message import EmailMessage
import io

# --- APP CONFIGURATION ---
st.set_page_config(page_title="AI Job Hunter Pro", layout="wide")
st.title("🤖 AI Automated Job Hunter")

# 1. Setup API Keys (User Input)
with st.sidebar:
    st.header("1. Settings")
    gemini_key = st.text_input("Enter Google Gemini API Key", type="password")
    email_user = st.text_input("Your Gmail Address")
    email_pass = st.text_input("Your Gmail App Password (16 digits)", type="password")
    
if gemini_key:
    genai.configure(api_key=gemini_key)

# 2. File Uploads
st.header("2. Upload Your Files")
col1, col2 = st.columns(2)
with col1:
    master_cv_file = st.file_uploader("Upload Master CV (Word .docx)", type="docx")
with col2:
    job_data = st.file_uploader("Upload HR Excel/Sheet", type=["xlsx", "csv"])

# 3. Main Logic
if st.button("🚀 Start Automated Outreach"):
    if not master_cv_file or not job_data or not gemini_key or not email_user or not email_pass:
        st.error("❌ Error: Please fill in all settings and upload both files!")
    else:
        try:
            df = pd.read_excel(job_data) if job_data.name.endswith('xlsx') else pd.read_csv(job_data)
            
            for index, row in df.iterrows():
                # Extract data from Excel columns
                hr_name = str(row.get('HR Name', 'Hiring Manager'))
                hr_email = str(row.get('Email', ''))
                job_desc = str(row.get('Job Description', 'Not provided'))
                
                st.info(f"Processing application for {hr_name}...")

                # STEP A: Tailor CV using AI (Using the updated 1.5-flash model)
                try:
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    prompt = f"Rewrite my professional summary and key skills for this job description. Keep it professional. Job: {job_desc}"
                    ai_response = model.generate_content(prompt).text
                    
                    # STEP B: Prepare the Word Document
                    doc = Document(master_cv_file)
                    doc.add_heading('AI Optimized Skills for this Role', level=1)
                    doc.add_paragraph(ai_response)
                    
                    cv_buffer = io.BytesIO()
                    doc.save(cv_buffer)
                    cv_buffer.seek(0)

                    # STEP C: Send Email
                    if "@" in hr_email:
                        msg = EmailMessage()
                        msg['Subject'] = f"Application for Role - Attn: {hr_name}"
                        msg['From'] = email_user
                        msg['To'] = hr_email
                        msg.set_content(f"Hi {hr_name},\n\nI am interested in this role. Please see my tailored CV attached.\n\nBest regards.")
                        
                        msg.add_attachment(
                            cv_buffer.read(), 
                            maintype='application', 
                            subtype='vnd.openxmlformats-officedocument.wordprocessingml.document', 
                            filename=f"Tailored_CV_{hr_name}.docx"
                        )
                        
                        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                            smtp.login(email_user, email_pass)
                            smtp.send_message(msg)
                        st.success(f"✅ Application sent successfully to {hr_email}")
                    else:
                        st.warning(f"⚠️ Skipped {hr_name}: No valid email found in Excel.")
                except Exception as e:
                    st.error(f"❌ Error with {hr_name}: {str(e)}")
        except Exception as main_e:
            st.error(f"❌ Critical Error reading Excel: {str(main_e)}")
