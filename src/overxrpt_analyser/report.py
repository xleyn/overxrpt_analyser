from pathlib import Path
import pdfplumber
from spellchecker import SpellChecker

import pandas as pd

from mappers import BadgeMapper
from row import Row


class Report:
    """Class for a Landauer report pdf from myldr.com"""

    def __init__(self, report_path: Path, pull_levels: bool = False):
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
