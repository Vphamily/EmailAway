import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# SMTP Configuration for Yahoo (App Password)
smtp_server = "smtp.mail.yahoo.com"
smtp_port = 587  # You can also use 465 with SSL
email_address = "haft_moon15@yahoo.com"
password = "frnhvjurpytfjtsk"  # Use the app password generated from Yahoo

try:
    # Set up the MIME
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = "haft_moon15@yahoo.com"
    msg['Subject'] = "Test Email"
    body = "This is a test email sent from Python using Yahoo."
    msg.attach(MIMEText(body, 'plain'))

    # Start the server
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()  # Secure the connection

    # Log in to the server using the App Password
    server.login(email_address, password)

    # Send the email
    server.sendmail(email_address, "haft_moon15@yahoo.com", msg.as_string())

    # Close the connection
    server.quit()

    print("Email sent successfully!")

except Exception as e:
    print(f"Error: {str(e)}")
