import smtplib
import win32com.client as win32
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

fileto_attach = "C:/Users/JDeClaro/Documents/python_project/Netsuite/updated_excel_file.csv"
# smtp_server ="smtp.gmail.com"
# smtp_port =587
# sender_email="jerico.declaro08@gmail.com"
# sender_password="Zyd@@Fhae21"
# recipient_email="jerico@oscargroup.com.au"

# message = MIMEMultipart()
# message['From']=sender_email
# message['To']=recipient_email
# message['Subject']="Report"

# email_body = "This is for import schedule"
# message.attach(MIMEText(email_body,'plain'))

# attachment = MIMEApplication(open(fileto_attach,'rb').read(),_subtype='txt')
# attachment.add_header('Content-Disposition', f'attachment; filename={fileto_attach}')
# message.attach(attachment)

# try:
#     server = smtplib.SMTP(host=smtp_server, port=smtp_port)
#     server.ehlo()
#     server.starttls()  # Upgrade the connection to secure (TLS)
#     server.login(sender_email, sender_password)
#     server.sendmail(sender_email, recipient_email, message.as_string())
#     server.quit()
#     print("Email sent successfully!")
# except Exception as e:
#     print(f"Error: {str(e)}")

def send_mail(emailto,subject,body,attachmentdir=None,ccmail=None):
    # Create an instance of the Outlook application
    try:
        outlook = win32.Dispatch('Outlook.Application')

        # Create an email message
        mail = outlook.CreateItem(0)

        # Set email properties
        mail.To =emailto# 'emails.5822203_SB1.2218.768a3fc372@5822203-sb1.email.netsuite.com'
        mail.CC = ccmail
        mail.Subject = subject
        mail.Body = body

        # Add attachments (optional)
        if attachmentdir!="":
            attachment_path = attachmentdir#fileto_attach
            mail.Attachments.Add(attachment_path)

        # Send the email
        mail.Send()

        print("Email sent successfully!")
    except Exception as e:
      print(f"Email Error: {str(e)}")


