import smtplib
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from FileHelper import FileHelper
from PopupHelper import Popups
from Settings import Settings

"""
This class handles sending emails containing the timesheet
"""


class Email:
    @staticmethod
    def _set_email(settings, excel):
        """Constructs an email using MIME
        :return: Returns MIMEMultipart object
        """
        weekend = excel.week_end()
        week_start = excel.week_start()
        dest = [str(settings.user), str(settings.dest_addr)]
        msg = MIMEMultipart('alternative')
        msg['Subject'] = 'Time sheet'
        msg['From'] = settings.user
        msg['To'] = ", ".join(dest)

        body = """Howdy,\n\nHere is my time sheet from %s to %s. Have a good weekend!\n\nThanks,\n\n%s""" %\
               (week_start, weekend, settings.name)

        derp = MIMEText(body, 'plain')
        msg.attach(derp)

        attach = MIMEBase('application', "octet-stream")
        attach.set_payload(open(settings.active_xl, "rb").read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', 'attachment; filename="Timesheet ' + settings.name + '.xlsx"')

        msg.attach(attach)

        return msg

    @staticmethod
    def _send_email_helper(settings, excel):
        """
        Connects to SMTP email server and sends an email using the users credentials.
        :return:
        """
        try:
            server = smtplib.SMTP(settings.smtp_server, str(settings.smtp_port))
            server.starttls()
            server.login(settings.user,settings.password)
            dest = [str(settings.user), str(settings.dest_addr)]
            server.sendmail(settings.user, dest, Email._set_email(settings,excel).as_string())
            server.quit()

            FileHelper.archive(settings, excel)
            excel.clear_sheet()
            excel.gen_dates()
            Popups.email_sent()
        except Exception, e:
            print(e.message)
            Popups.email_failed()



    @staticmethod
    def send_email(settings, excel):
        """
        Call this to send an email.
        :return:
        """
        Email._set_email(settings, excel)
        Email._send_email_helper(settings, excel)
