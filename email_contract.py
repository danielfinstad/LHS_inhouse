from os.path import basename
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders


def send_mail(send_from, send_to, subject, text, files=None, username=None,
              password=None, istls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text, 'html'))

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(f, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(basename(f)))
        msg.attach(part)

    server = SMTP('smtp.gmail.com:587')
    if istls:
        server.starttls()
        server.login(username, password)
    server.sendmail(send_from, send_to, msg.as_string())
    server.quit()

# if __name__ == "__main__":
#     send_mail("danielfinstad@gmail.com", ["danielfinstad@gmail.com"], "subject", "text",
#               files=["C:/Users/Daniel/PycharmProjects/mm_contract/test_pdf_out.pdf"],
#               username='',
#               password='')
