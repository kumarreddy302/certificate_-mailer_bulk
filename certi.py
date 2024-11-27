import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

# Email credentials and server settings
smtp_server = "smtp.lmsrscs.me"
smtp_port = 465
username = "no-reply-otp@lmsrscs.me"
password = "Kumar31@"  # Replace with your actual password

# Load cleaned recipient data
data = pd.read_csv('123.csv')

# Initialize SMTP server connection
server = smtplib.SMTP_SSL(smtp_server, smtp_port)
server.login(username, password)

# Send email to each recipient
for index, row in data.iterrows():
    recipient_email = row['Email']
    recipient_name = row['Name']
    file_path = row['FilePath']

    # Create email message
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = recipient_email
    msg['Subject'] = "Thank You for Joining Our Ethical Hacking Workshop!"

    # Email body
    body = f"Thank you for attending our Ethical Hacking Workshop held on 20th and 21st November. We are thrilled to have had the opportunity to share valuable insights and hands-on techniques in cybersecurity with you.We hope you found the sessions informative and engaging, and that the knowledge you gained will empower you in your journey toward mastering ethical hacking and securing systems effectively.Your active participation made this workshop a great success, and we truly appreciate your enthusiasm and curiosity."
    msg.attach(MIMEText(body, 'plain'))

    # Attach file
    try:
        with open(file_path, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename={file_path.split('/')[-1]}'
            )
            msg.attach(part)
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        continue

    # Send email
    try:
        server.sendmail(username, recipient_email, msg.as_string())
        print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

# Close the server connection
server.quit()
