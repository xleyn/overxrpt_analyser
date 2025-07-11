import sys
import ctypes
import getpass

from pathlib import Path
import win32com.client
from datetime import datetime

from file_manager import FileManager
from spreadsheet import Spreadsheet


class Email:
    """Class for email output"""

    def __init__(self, cur_report: "Report", subaccount_code: str):
        self.attachments = [rf"{str(cur_report.path)}"]
        template_name = FileManager.paths_from_proj_dir["formal_investigation_template"]
        if any([row.attachTemplate for row in cur_report.rows]):
            self.attachments.append(
                str(Path(sys.argv[0]).parent.joinpath(template_name))
            )

        self.urgent_flag = (
            "URGENT:" if any([row.urgent_flag for row in cur_report.rows]) else ""
        )
        self.subject = " ".join(
            [self.urgent_flag, subaccount_code, "Landauer Overexposure Notification"]
        )
        self.bullets = self.get_bullets(cur_report)
        self.body = self.write_body(subaccount_code)
        self.draft()

    def get_bullets(self, report: "Report") -> str:
        """Processes responses from report rows into a list of bullet points, formatted in HTML

        Args:
            report (Report): report object to send off

        Returns:
            str: Bullet points for report responses, in HTML format

        """

        bullets = "".join(
            [f"<li>{point}</li>" for point in [row.response for row in report.rows]]
        )

        return bullets

    def write_body(self, subaccount_code: str) -> str:
        """Writes body of email in HTML using f-strings

        Args:
            subaccount_code (str): Subaccount code for report

        Returns:
            body (str): Body of email written in HTML

        """
        template_phrase = "I have attached our report template for the exceedance of the annual investigation level as a suggestion for the investigation."

        body = f"""
        <html>
        <head>
        <style>
        body {{
            font-family: Calibri, Arial, sans-serif;
            font-size: 11pt;
        }}
        </style>
        </head>
        <body>
        <p>Good {"morning" if datetime.now().time().hour < 12 else "afternoon"},</p>
        <p>Please find attached an overexposure notification from Landauer for subaccount {subaccount_code}.</p>
        <p>The following dose thresholds have been exceeded:</p>
        <ul>
        {self.bullets}
        </ul>
        <p>Please investigate the reasons behind these high doses and report back to RRPS.</p>
        <p>{template_phrase if len(self.attachments) >= 2 else ""}</p>
        <p>Kind regards,</p>
        <p>{self.get_duty_physicist_first_name()}</p>
        </body>
        </html>
        """

        return body

    @staticmethod
    def get_duty_physicist_first_name():
        name = ctypes.create_unicode_buffer(1024)
        size = ctypes.pointer(ctypes.c_ulong(len(name)))
        result = ctypes.windll.secur32.GetUserNameExW(3, name, size)  # NameDisplay
        if result:
            full_name = name.value
            first_name = full_name.split()[0]  # First word only
            return first_name
        else:
            return getpass.getuser()  # fallback to username

    def draft(self):
        """Writes a draft of email for checking and opens in Outlook

        Returns:
            None.

        """
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.Subject = self.subject
        email.HTMLBody = self.body
        email.SentOnBehalfOfName = "rsc-tr.RadProt@nhs.net"
        for attachment in self.attachments:
            email.Attachments.Add(attachment)
        email.To = Spreadsheet.contact
        email.CC = Spreadsheet.CC
        email.Display()
