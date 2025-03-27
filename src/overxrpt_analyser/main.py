import pathlib
from tkinter import Tk, filedialog

import itertools

from spreadsheet import Spreadsheet
from row import Row
from report import Report
from email_obj import Email


class LandauerTask:
    """Class for main task. Entry point for report analysis"""

    def run(self):
        """Entry point for running task.

        Returns:
            None.

        """
        self.cur_report_path = self._choose_path()
        self.account_code = self._get_code(self.cur_report_path, "_AC", 6)
        self.subaccount_code = self._get_code(self.cur_report_path, "_SUB", 3)
        Spreadsheet.load_df(self.account_code, self.subaccount_code)
        self.cur_report = Report(self.cur_report_path, pull_levels=True)
        setattr(Row, "past_notifs", self._get_past_notifications())
        self.cur_report.analyse()
        self.email = Email(self.cur_report, self.subaccount_code)

    @staticmethod
    def _choose_path() -> pathlib.Path:
        """Prompts user to choose the Landauer report to analyse using GUI

        Returns:
            path (pathlib.Path): path to the report for analysis

        """
        root = Tk()
        root.withdraw()
        root.call("wm", "attributes", ".", "-topmost", True)
        path = filedialog.askopenfilename(
            title="Please select a Report.", filetypes=[("PDF files", "*.pdf")]
        )
        root.destroy()
        path = pathlib.Path(path)

        return path

    @staticmethod
    def _get_code(reportPath: pathlib.Path, locator: str, targetLen: int) -> str:
        """Gets account or subaccount code using name of pdf path

        Args:
            reportPath (pathlib.Path): path to report
            locator (str): string used to find "zero-point" for finding target string
            targetLen (int): Length of target string

        Returns:
            str: returned account or subaccount code

        """
        index = reportPath.name.find(locator) + len(locator)
        code = reportPath.name[index : index + targetLen]

        return code

    def _get_past_notifications(self) -> list[Row]:
        """Gets all past notifications from same subaccount and year as current report

        Returns:
            list[Row]: List of Row objects representing past notifications

        """
        past_reports = [
            Report(path, pull_levels=False)
            for path in self.cur_report_path.parent.glob("*.pdf")
            if (
                "OVXRPT_" in str(path.name)
                and path.name != self.cur_report_path.name
                and path.stat().st_mtime < self.cur_report_path.stat().st_mtime
            )
        ]
        past_notifications = list(
            itertools.chain(*[report.rows for report in past_reports])
        )

        return past_notifications


if __name__ == "__main__":
    landauer_task = LandauerTask()
    landauer_task.run()
