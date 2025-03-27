import sys
from pathlib import Path
import win32com.client
from datetime import datetime

from io_manager import Spreadsheet, IO


class Email:
    """Class for email output"""

    def __init__(self, cur_report: "Report", subaccount_code: str):
        self.attachments = [rf"{str(cur_report.path)}"]
        template_name = IO.paths_from_proj_dir["formal_investigation_template"]
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
        bullets = {}
        for row in report.rows:
            bullets.setdefault(row.name, []).append(row.response)

        for key, responses in bullets.items():
            bullets[key] = " ".join(
                [responses[0]]
                + [
                    r.replace(f"For wearer {key}", "Also for this wearer")
                    for r in responses[1:]
                ]
            )
        bullets = "".join([f"<li>{point}</li>" for point in list(bullets.values())])

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
        </body>
        </html>
        """

        return body

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
        email.Display()
