from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import os
import smtplib
from typing import List


def send_email(
    smtp_server: str,
    smtp_port: int,
    smtp_email_username: str,
    subject: str,
    html_body: str,
    recipient_emails: List[str],
    cc_emails: List[str] = [],
    bcc_emails: List[str] = [],
    display_sender_name: str = "",
    display_sender_email: str = "",
    attachment_path="",
    smtp_email_password=""
):
    recipient_emails = list(set(recipient_emails))
    cc_emails = list(set(cc_emails))
    bcc_emails = list(set(bcc_emails))

    message = MIMEMultipart()
    message["From"] = f"{display_sender_name} <{display_sender_email}>"
    message["To"] = ", ".join(recipient_emails)
    message["Cc"] = ", ".join(cc_emails)
    message["Bcc"] = ", ".join(bcc_emails)
    message["Subject"] = subject
    message.attach(MIMEText(html_body, "html"))

    if attachment_path:
        _, filename = os.path.split(attachment_path)
        file_extension = filename.split(".")[-1]
        with open(attachment_path, "rb") as f:
            attach = MIMEApplication(f.read(),_subtype=file_extension)
            attach.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            message.attach(attach)

    all_recipients = recipient_emails + cc_emails + bcc_emails
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        try:
            server.starttls()
            if smtp_email_password:
                server.login(smtp_email_username, smtp_email_password)
            server.sendmail(smtp_email_username, all_recipients, message.as_string())
        except smtplib.SMTPRecipientsRefused as e:
            print(f"Failed to send email. SMTP Recipients Refused: {e}")


if __name__ == "__main__":
    path_xls = "C:\\Users\\maKse\\Documents\\Sansatech\\sample data\\mutasi\\output\\2024-02-28\\output.xls"
    path_pdf = "C:\\Users\\maKse\\Documents\\Sansatech\\sample data\\mutasi\\output\\2024-02-28\\126823433_PDF\\Transaction Inquiry_Download Single_20240228144303981.pdf"

    send_email(
        {
            "host": "smtp.gmail.com",
            "port": 587,
            "email_address": "kemas.m3@gmail.com",
            "password": "qloo iosx pkom pbee"
        },
        "foo",
        "bar",
        ["kemas.mr@sansatech.co.id"],
        [],
        [],
        "haha",
        "haha@hehe.hihi",
        path_pdf
    )