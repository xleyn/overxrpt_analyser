from pathlib import Path
import json
import sys
import time
import warnings

import re
import pandas as pd

from mappers import BadgeMapper, TemporalMapper

warnings.filterwarnings(
    "ignore",
    message="Conditional Formatting extension is not supported and will be removed",
)


class IO:
    """Class for managing I/O operations."""

    if getattr(sys, "frozen", False):
        project_dir = Path(sys.executable).parent
    else:
        project_dir = Path(__file__).parent.parent.parent

    with open(project_dir.joinpath("config/file_structure.json")) as f:
        config = json.load(f)
        paths_from_proj_dir = dict(
            zip(
                config.keys(),
                map(
                    lambda v, project_dir=project_dir: project_dir / v, config.values()
                ),
            )
        )

    @classmethod
    def creation_control(cls):
        """Checks if I/O paths exist. Exits programme if not, rectifying the issues."""
        need_to_quit = False
        for path in cls.paths_from_proj_dir:
            if not path.exists:
                print(f"Does not exist: {path}")
                need_to_quit = True
        if need_to_quit:
            print("Please rectify issues. Quitting application.")
            time.sleep(3)
            sys.exit()
        else:
            print("All I/O paths exist, as expected!")


class Spreadsheet:
    """
    Class for investigation levels spreadsheet
    """

    name = "Landauer client list and dose investigation limits.xlsx"
    path = IO.paths_from_proj_dir["levels_spreadsheet"]
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
            sheet_name for sheet_name in cls.sheets.keys() if account_code in sheet_name
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
        contact = df.loc[subaccount_code].values[-1]
        if "@" in contact:
            cls.contact = contact
        else:
            raise ValueError("Contact email for subaccount not found!")

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
        slice_end = df.iloc[0].str.contains("email", case=False).idxmax() + 1
        df = df.iloc[:, :slice_end]

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
                ) & multiIndex_period.map(
                    lambda x: check_substring_in_string(period_typeTest, x)
                )

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
            raise ValueError(
                "Required investigation level(s) missing from spreadsheet!"
            )
