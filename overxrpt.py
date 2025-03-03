import pathlib
from tkinter import Tk, filedialog
import pdfplumber
import pandas as pd
from datetime import datetime
import win32com.client
import sys
import warnings
import re
import itertools
from spellchecker import SpellChecker
from dataclasses import dataclass, field

warnings.filterwarnings(
    "ignore", message="Conditional Formatting extension is not supported and will be removed"
)


class Spreadsheet:
    """
    Class for investigation levels spreadsheet
    """
    name = "Landauer client list and dose investigation limits.xlsx"
    path = pathlib.Path.cwd().joinpath(name)
    sheets = pd.read_excel(path, sheet_name=None)
    df = None
    row = None

    @classmethod
    def load_df(cls, account_code: str, subaccount_code: str):
        """Pulls the relevant Excel sheet and the subaccount row down into pd DataFrames.

        Args:
            account_code (str): account code for the Landauer report.
            subaccount_code (str): subaccount code for the Landauer report.

        Returns:
            None.

        """
        shortlisted_sheets = [
            sheetName for sheetName in cls.sheets.keys() if account_code in sheetName
        ]
        selected_sheet = sorted(
            shortlisted_sheets, key=lambda sn: len(sn.replace(account_code, ""))
        )[0]
        df = pd.read_excel(cls.path, sheet_name=selected_sheet, header=None, dtype=str)
        df = cls._crop_df(df)
        df = cls._initiate_multi_index(df)
        row = (
            df.loc[subaccount_code]
            if not df.loc[subaccount_code].str.contains("DEFAULT", case=False).any()
            else df.loc[account_code]
        )
        cls.df = df
        cls.row = row

    @staticmethod
    def _crop_df(df: pd.DataFrame) -> pd.DataFrame:
        """Crops df to include only relevant info for further processing

        Args:
            df (pd.DataFrame): Input df

        Returns:
            df (pd.DataFrame): Cropped df

        """

        # set index as code column and subsequently drop code and name columns
        code_col = df.iloc[0].str.contains("code", case=False).idxmax()
        name_col = df.iloc[0].str.contains("name", case=False).idxmax()
        df.index = df.loc[:, code_col].astype(str)
        df = df.drop(df.columns[[code_col, name_col]], axis=1)
        df.columns = range(df.columns.size)

        # Cropping end of df
        contact_col = df.iloc[0].str.contains("contact", case=False).idxmax()
        df = df.iloc[:, :contact_col]

        return df

    @staticmethod
    def _initiate_multi_index(df: pd.DataFrame) -> pd.DataFrame:
        """Sets up multi-indexing for pandas df

        Args:
            df (pd.DataFrame): Input df
        Returns:
            df (pd.DataFrame): Output df with multi-indexing implemented

        """
        categories, subcategories = df.iloc[:2].ffill(axis=1).values
        df.columns = pd.MultiIndex.from_arrays([categories, subcategories])
        df = df.iloc[2:]
        return df

    @classmethod
    def get_levels(cls, badge: str, period_type: str) -> list[str]:
        """Pulls levels from df according to specific badge and period type

        Args:
            badge (str): Badge type for report row
            period_type (str): Type of period for report row (i.e. monthly, quarterly or YTD)

        Returns:
            list[str]: List containing level and urgent level

        """
        hierarchyBadge = BadgeMapper.get_Excel_Hierarchy(badge)
        hierarchyTemporal = TemporalMapper.get_Excel_Hierarchy(period_type)

        def check_substring_in_string(substring: str, string: str):
            """Checks if a substring is in a string

            Args:
                substring (str): Substring to check
                string (str): String to check if contains substring

            Returns:
                bool: True if string contains substring, else False

            """
            string = re.sub(r"[()]", "", string)
            pattern = rf"\b{re.escape(substring)}\b"
            return bool(re.search(pattern, string, re.IGNORECASE))

        multiIndex_badge = cls.df.columns.get_level_values(0)
        multiIndex_period = cls.df.columns.get_level_values(1)

        foundCol = None
        for badgeTest in hierarchyBadge:
            for period_typeTest in hierarchyTemporal:

                bool_mask = multiIndex_badge.map(
                    lambda x: check_substring_in_string(badgeTest, x)
                ) & multiIndex_period.map(lambda x: check_substring_in_string(period_typeTest, x))

                if any(bool_mask):
                    foundCol = cls.df.columns[bool_mask.get_loc(True)]
                    break

            if foundCol is not None:
                break

        if foundCol is None:
            raise ValueError("Badge type and period type not found in spreadsheet!")

        level = cls.row[foundCol]
        urgentLevel = cls.row[foundCol[0], "Urgent"]

        try:
            float(level)
            return level, urgentLevel
        except ValueError:
            raise ValueError("Required investigation level(s) missing from spreadsheet!")


class BadgeMapper:
    """Class for extracting data mapped from badge names"""

    wholeBody = ("collar", "other whole body", "chest", "waist")
    extremity = ("left finger", "right finger")
    lens = "lens"

    hier_Excel_wb = ["DDE", "WHOLE BODY"]
    hier_Excel_lens = ["LDE", "LENS"]
    hier_Excel_extrem = ["EXTREMITY"]

    badge_to_col_mapping = {
        wholeBody: "Whole Body",
        extremity: "Total Extremity",
        lens: "Lens",
    }

    badge_to_hierarchy_mapping = {
        wholeBody: lambda b, hierarchy=hier_Excel_wb: [b] + hierarchy,
        extremity: lambda b, hierarchy=hier_Excel_extrem: hierarchy,
        lens: lambda b, hierarchy=hier_Excel_lens: hierarchy,
    }

    @classmethod
    def get_dose_column(cls, badge: str) -> str:
        """Gets the dose column for a report df.
        Corrected column selected for badge type.

        Args:
            badge (str): Badge type for row

        Returns:
            str:  Column name corresponding to correct dose in df
        """

        col = next(
            (col for badges, col in cls.badge_to_col_mapping.items() if badge in badges),
            None,
        )
        if col is None:
            raise ValueError("Badge type not implemented!")
        return col

    @classmethod
    def get_Excel_Hierarchy(cls, badge: str) -> list[str]:
        """Gets a hierarchy for badge type column selection in Excel multi-index

        Args:
            badge (str): Badge type for row

        Returns:
            list[str]: Hierarchial list for badge type column selection in Excel multi-index
                       Iteration goes from most specific to least specific.

        """
        method = next(
            (
                method
                for badges, method in cls.badge_to_hierarchy_mapping.items()
                if badge in badges
            ),
            None,
        )
        if method is None:
            raise ValueError("Badge type not implemented!")
        hierarchy = method(badge)

        return hierarchy


class TemporalMapper:
    """Class for extracting data mapped from temporal data"""

    period_type_to_hierarchy_mapping = {
        "monthly": ["MONTHLY", "WEAR PERIOD"],
        "quarterly": ["QUARTERLY", "WEAR PERIOD"],
        "YTD": ["ANNUAL"],
    }

    @classmethod
    def get_Excel_Hierarchy(cls, period_type: str) -> list[str]:
        """Gets a hierarchy for period type column selection in Excel multi-index

        Args:
            period_type (str): period type for row

        Returns:
            list[str]: Hierarchial list for period type column selection in Excel multi-index
                       Iteration goes from most specific to least specific.

        """
        hierarchy = cls.period_type_to_hierarchy_mapping[period_type]

        return hierarchy


class Report:
    """Class for a Landauer report pdf from myldr.com"""

    def __init__(self, report_path: pathlib.Path, pull_levels: bool = False):
        """Initialising object

        Args:
            report_path (pathlib.Path): pathlib.Path
            pull_levels (bool, optional): Bool for whether investigation levels should be pulled from Excel spreadsheet. The default is False.

        Returns:
            None.

        """
        self.path = report_path
        self.df = self._pull_dataframe()
        self.rows = self._get_row_objs(pull_levels)

    def _pull_dataframe(self):
        """Returns pd dataframe corresponding to table on Landauer report.

        Returns:
            df (pd.DataFrame): pandas df for table on report

        """
        with pdfplumber.open(self.path.resolve()) as pdf:
            page = pdf.pages[0]

            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "join_x_tolerance": 1000,
                "join_y_tolerance": 1000,
            }
            tableObjs = page.find_tables(table_settings)
            df = sorted(tableObjs, key=lambda t: t.bbox[2] * t.bbox[3], reverse=True)[0]
            df = page.within_bbox(df.bbox).extract_table(table_settings)
            df = pd.DataFrame(df, dtype=str)

            # crop and then fill down name and participant number to account for multiple badges per person
            df = df.iloc[:, 1:-3]
            df.iloc[:, [0, 1]] = df.iloc[:, [0, 1]].replace("", pd.NA).ffill()

            df = self._format_and_assign_header(df)
            df = self._delete_notes_rows(df)
            df = self._standardise_strings(df)
            df["Name"] = df["Name"].map(self._format_name)
            df["Use"] = df["Use"].map(str.lower)

            return df

    @staticmethod
    def _format_and_assign_header(df: pd.DataFrame) -> pd.DataFrame:
        """Formats header of df according to expected formatting

        Args:
            df (pd.DataFrame): Input df

        Returns:
            pd.DataFrame: df with formatted header

        """
        header = df.iloc[1].combine_first(df.iloc[0])
        header = header.str.replace("\n", " ")
        dictionary = SpellChecker()

        def reverse_checker(phrase: str):
            words = phrase.split()
            in_dict = [word.lower() in dictionary for word in words]
            return phrase if all(in_dict) else phrase[::-1]

        header = header.apply(reverse_checker)
        df.columns = header
        df = df.iloc[2:]

        return df

    @staticmethod
    def _delete_notes_rows(df: pd.DataFrame):
        """Deletes unwanted "notes" rows on report. Not needed for analysis.

        Args:
            df (pd.DataFrame): Input df

        Returns:
            df (pd.DataFrame): df with notes rows deleted

        """
        condition = (
            (df["Begin Date"] == df["End Date"])
            & (df["End Date"] == df["Frequency"])
            & (df["Frequency"] == "")
        )
        df = df[~condition]
        df = df.reset_index(drop=True)

        return df

    @staticmethod
    def _standardise_strings(df: pd.DataFrame) -> pd.DataFrame:
        """Standardises strings to make more readable by replacement of abbreviations.

        Args:
            df (pd.DataFrame): Input df

        Returns:
            df (pd.DataFrame): Outputted dataframe with strings standardised

        """
        df = df.replace("\n", "", regex=True)
        string_replacement = {
            "OTHERWHBODY": "OTHER WHOLE BODY",
            "L FINGER": "LEFT FINGER",
            "LFINGER": "LEFT FINGER",
            "R FINGER": "RIGHT FINGER",
            "RFINGER": "RIGHT FINGER",
        }
        df = df.replace(string_replacement)

        return df

    @staticmethod
    def _format_name(name: str) -> str:
        """Formats name string according to required formatting of final email.

        Args:
            name (str): Name string
        Returns:
            str: Formatted name string

        """
        delim = ","
        if delim in name:
            sur, fore = [name.strip() for name in name.split(delim)]
            name = (
                f"{fore.capitalize()} {sur.capitalize()}"
                if fore == "DR"
                else f"{fore[0]}. {sur.capitalize()}" if fore else sur
            )

        return name

    def _get_row_objs(self, pull_levels: bool) -> list["Row"]:
        """Generates a list of row objects, each associated with one row in the report df

        Args:
            pull_levels (bool): Bool describing whether investigation levels should be tracked in Row class.

        Returns:
            list["Row"]: List of row objects for report df.

        """
        rows = []
        for rowIndex in range(len(self.df)):
            name = self.df.at[rowIndex, "Name"]
            badge = self.df.at[rowIndex, "Use"]
            dose = self.df.at[rowIndex, BadgeMapper.get_dose_column(badge)]
            freq = self.df.at[rowIndex, "Frequency"]
            bDate = self.df.at[rowIndex, "Begin Date"]
            eDate = self.df.at[rowIndex, "End Date"]
            rows.append(Row(name, badge, dose, freq, bDate, eDate, pull_levels))

        return rows

    def analyse(self):
        """Called externally to analyse the report.

        Returns:
            None.

        """
        for row in self.rows:
            row.analyse()


@dataclass
class Row:
    """Class for row in report"""

    name: str
    badge: str
    dose: str
    freq: str
    bDate: str
    eDate: str
    pull_levels: bool

    period: str = field(init=False)
    period_type: str = field(init=False)
    level: str = field(init=False)
    urgentLevel: str = field(init=False)

    def __post_init__(self):
        """
        Instance and class attribute initialisation after __init__.
        period, period type, investigation level and urgent level are all defined.

        Returns:
            None.

        """
        if not hasattr(Row, "freq2method_mapping"):
            self.init_class_attributes()
        self.period, self.period_type = self.get_temporal_data()
        if self.pull_levels:
            self.level, self.urgentLevel = Spreadsheet.get_levels(self.badge, self.period_type)

    @classmethod
    def init_class_attributes(cls):
        """Initialises class attributes post __init__
        Dicts initialised to map specific instance attributes of Row obj to strings or methods associated with data processing

        Returns:
            None.

        """
        cls.freq2method_mapping = {
            "1MO": cls.gtd_monthly,
            "3MO": cls.gtd_quarterly,
            "": cls.gtd_ytd,
        }

        cls.month2quarter_mapping = {
            (1, 3): "first",
            (4, 6): "second",
            (7, 9): "third",
            (10, 12): "fourth",
        }

        cls.period_type2method_mapping = {
            "monthly": cls.gr_monthly,
            "quarterly": cls.gr_quarterly,
            "YTD": cls.gr_ytd,
        }

    def gtd_monthly(self) -> list[str]:
        """Method to return period and period type for a frequency = "monthly"

        Returns:
            list[str]: List containing period and period type associated with row

        """
        period_type = "monthly"
        bDateObj = datetime.strptime(self.bDate, "%Y-%m-%d")
        period = bDateObj.strftime("%B") + " " + str(bDateObj.year)

        return period, period_type

    def gtd_quarterly(self) -> list[str]:
        """Method to return period and period type for a frequency = "quarterly"

        Returns:
            list[str]: List containing period and period type associated with row

        """
        period_type = "quarterly"
        bDateObj = datetime.strptime(self.bDate, "%Y-%m-%d")
        eDateObj = datetime.strptime(self.eDate, "%Y-%m-%d")
        quarter = self.month2quarter_mapping[(bDateObj.month, eDateObj.month)]
        period = f"{quarter} quarter of {eDateObj.year}"

        return period, period_type

    def gtd_ytd(self) -> list[str]:
        """
        Method to return period and period type for a frequency = "YTD"

        Returns:
            list[str]: List containing period and period type associated with row

        """
        period_type = "YTD"
        period = self.bDate[:-3]

        return period, period_type

    def get_temporal_data(self) -> list[str]:
        """Returns the period and period type associated with the row

        Returns:
            list[str]: List containing the period and period type for the row

        """
        get_temporal = self.freq2method_mapping[self.freq]
        period, period_type = get_temporal(self)
        return period, period_type

    def gr_monthly(self) -> str:
        """Returns a response string for the analysis of a monthly badge

        Returns:
            str: Response string

        """
        localResponse = {
            True: f"For wearer {self.name} in the month of {self.period}, they exceeded the monthly local review level of {self.level} mSv on their {self.badge} badge.",
            False: f"For wearer {self.name} in the month of {self.period}, the {self.badge} badge alert can be ignored as the monthly local review level is {self.level} mSv on their {self.badge} badge.",
        }

        return localResponse[float(self.dose) >= float(self.level)]

    def gr_quarterly(self) -> str:
        """Returns a response string for the analysis of a quarterly badge

        Returns:
            str: Response string

        """
        localResponse = {
            True: f"For wearer {self.name} in the {self.period}, they exceeded the quarterly local review level of {self.level} mSv on their {self.badge} badge.",
            False: f"For wearer {self.name} in the {self.period}, the {self.badge} badge alert can be ignored as the quarterly local review level is {self.level} mSv on their {self.badge} badge.",
        }

        return localResponse[float(self.dose) >= float(self.level)]

    def gr_ytd(self) -> str:
        """Returns a response string for the analysis of a YTD badge

        Returns:
            str: Response string

        """
        localResponse = {
            True: self.gr_if_ytd_already_raised(),
            False: f"For wearer {self.name}, the year to date alert on this report can be ignored as the annual formal investigation level is {self.level} mSv on their {self.badge} badge.",
        }

        return localResponse[float(self.dose) >= float(self.level)]

    def gr_if_ytd_already_raised(self) -> str:
        """Returns a response string for the YTD case where dose >= level, based on whether the YTD has already been raised

        Returns:
            str: Response string

        """
        localResponse = {
            True: f"For wearer {self.name}, no further investigation is required on the {self.badge} badge year to date alert as this should have already been communicated and investigated.",
            False: f"For wearer {self.name}, they have now exceeded the annual formal investigation level of {self.level} mSv on their {self.badge} badge.",
        }

        for pn in self.past_notifs:
            conditions = [
                pn.name == self.name,
                pn.badge == self.badge,
                pn.period_type == self.period_type,
                pn.period == self.period,
                float(self.dose) >= float(pn.dose),
            ]

            if all(conditions):

                return localResponse[True]
        else:
            self.attachTemplate = True

            return localResponse[False]

    def urgent_flag_query(self) -> bool:
        """Returns a bool showing whether the urgent flag has already been raised in the past

        Returns:
            bool: Bool for whether urgent flag has been raised in past

        """
        for pn in self.past_notifs:
            conditions = [
                pn.name == self.name,
                pn.badge == self.badge,
                pn.period_type == self.period_type,
                pn.period == self.period,
                float(self.dose) >= float(pn.dose) >= float(self.urgentLevel),
            ]

            if all(conditions):

                return True

        return False

    def analyse(self):
        """Performs analysis for particular row by comparing to investigation levels.

        Returns:
            None.

        """
        self.attachTemplate = False
        get_response = self.period_type2method_mapping[self.period_type]
        self.response = get_response(self)
        self.urgent_flag = (
            self.urgent_flag_query() if float(self.dose) >= float(self.urgentLevel) else False
        )


class Email:
    """Class for email output"""

    def __init__(self, cur_report: Report, subaccount_code: str):
        self.attachments = [rf"{str(cur_report.path)}"]
        template_name = "Staff_Dosimetry___Formal_Investigation_Form_v1.1.docx"
        if any([row.attachTemplate for row in cur_report.rows]):
            self.attachments.append(str(pathlib.Path(sys.argv[0]).parent.joinpath(template_name)))

        self.urgent_flag = "URGENT:" if any([row.urgent_flag for row in cur_report.rows]) else ""
        self.subject = " ".join(
            [self.urgent_flag, subaccount_code, "Landauer Overexposure Notification"]
        )
        self.bullets = self.get_bullets(cur_report)
        self.body = self.write_body(subaccount_code)
        self.draft()

    def get_bullets(self, report: Report) -> str:
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
                + [r.replace(f"For wearer {key}", "Also for this wearer") for r in responses[1:]]
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
        for attachment in self.attachments:
            email.Attachments.Add(attachment)
        email.To = "rsc-tr.RadProt@nhs.net"
        email.Display()


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
            str: DESCRIPTION.

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
        past_notifications = list(itertools.chain(*[report.rows for report in past_reports]))

        return past_notifications


if __name__ == "__main__":
    lt = LandauerTask()
    lt.run()
