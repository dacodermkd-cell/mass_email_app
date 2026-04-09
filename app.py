import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import time

st.title("Mass Email Sender with Attachments")

sender_email = st.text_input("Your Email")
app_password = st.text_input("App Password", type="password")

subject = st.text_input("Email Subject")

body_template = st.text_area(
    "Email Body (use {name})",
    "Hello {name},\n\nPlease find attached your agenda.\n\nRegards"
)

cc_input = st.text_area("CC Emails (comma separated)", "")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

uploaded_pdfs = st.file_uploader(
    "Upload PDF Attachments",
    type=["pdf"],
    accept_multiple_files=True
)

delay = st.slider("Delay between emails (seconds)", 1, 10, 3)

if st.button("Send Emails"):

    if uploaded_file is None:
        st.error("Upload Excel file first")
        st.stop()

    df = pd.read_excel(uploaded_file)

    required_cols = ["Name", "Email", "Agenda"]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"Missing column: {col}")
            st.stop()

    # Create PDF dictionary
    pdf_dict = {file.name: file for file in uploaded_pdfs}

    cc_list = [email.strip() for email in cc_input.split(",") if email.strip()]

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, app_password)

        progress = st.progress(0)
        total = len(df)

        for i, row in df.iterrows():
            name = row["Name"]
            recipient = row["Email"]
            filename = str(row["Agenda"])

            body = body_template.replace("{name}", str(name))

            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient
            msg["Subject"] = subject
            msg["Cc"] = ", ".join(cc_list)

            msg.attach(MIMEText(body, "plain"))

            # Attach PDF from uploaded files
            if filename in pdf_dict:
                pdf_file = pdf_dict[filename]
                part = MIMEApplication(pdf_file.read(), _subtype="pdf")
                part.add_header("Content-Disposition", "attachment", filename=filename)
                msg.attach(part)
            else:
                st.warning(f"PDF not found for {name}: {filename}")

            recipients = [recipient] + cc_list

            server.sendmail(sender_email, recipients, msg.as_string())

            progress.progress((i + 1) / total)
            time.sleep(delay)

        server.quit()
        st.success("Emails sent successfully!")

    except Exception as e:
        st.error(f"Error: {e}")
