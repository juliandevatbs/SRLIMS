from email.message import EmailMessage
import smtplib


def send_email(to_email: str, subject: str, body: str):

    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587

    EMAIL_FROM = "julianhomezdev@gmail.com"
    EMAIL_PASSWORD = "uygr nlzw yojz gssw"

    try:
        msg = EmailMessage()

        msg["From"] = EMAIL_FROM
        msg["To"] = to_email
        msg["Subject"] = subject

        msg.set_content("This email requires an HTML compatible email client.")
        msg.add_alternative(body, subtype="html")

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.send_message(msg)

        return True

    except Exception as e:
        print(f"Error sending email: {e}")
        return False
