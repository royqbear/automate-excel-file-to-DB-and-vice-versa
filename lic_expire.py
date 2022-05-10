import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

class Mail:
    BODY_OK = "Plese see attahced file for expiring licenses.\nto stop this service please close 'expiring_licenses' from services.msc"
    BODY_NOT_OK = '''File is opened therefore latest version attached,\nmake sure the file is closed ('lic1.xlsx')
        \nto stop this service please close 'expiring_licenses' from services.msc'''
    def __init__(self):
        pass
    def send_mail(self, body):
        sent_from = "pythroy@gmail.com"
        to_mail = 'roy.israelit@gmail.com'
        msg = MIMEMultipart()
        msg['From'] = sent_from
        msg['To'] = to_mail
        msg['Subject'] = 'Expiring licenses'
        # body = 'See attached file for expiring licenses'
        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))
        filename = (r"C:\Users\royqb\Desktop\python_learning\softwaremgmt\lic1.xlsx")
        attachment = open(filename, "rb")
        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')
        # To change the payload into encoded form
        p.set_payload((attachment).read())
        # encode into base64
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        # attach the instance 'p' to instance 'msg'
        msg.attach(p)
        # creates SMTP session
        s = smtplib.SMTP('smtp.gmail.com', 587)

        # start TLS for security
        s.starttls()

        # Authentication from sql hashed
        s.login(sent_from, "@WS2ws2ws")

        # Converts the Multipart msg into a string
        text = msg.as_string()

        # sending the mail
        s.sendmail(sent_from, to_mail, text)
        print('Mail sent successfully')
        # terminating the session
        s.quit()

    def get_excel_with_pandas(self):
        data = pd.read_excel(open('lic.xlsx', 'rb'),
                             sheet_name='lic')
        df = pd.DataFrame(data)
        df.sort_values(by=['update_in'], ascending=True)
        try:
            df.to_excel("lic1.xlsx", sheet_name="Expiring Licenses")
        except PermissionError:
            print("File is opened, attached latest version")
            body = self.BODY_NOT_OK
            self.send_mail(body)
        else:
            body = self.BODY_OK
            self.send_mail(body)

if __name__ == "__main__":
    mail = Mail()
    mail.get_excel_with_pandas()