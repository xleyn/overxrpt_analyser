import re
import pandas as pd
import warnings

from mappers import BadgeMapper, TemporalMapper
from file_manager import FileManager

warnings.filterwarnings(
    "ignore",
    message="Conditional Formatting extension is not supported and will be removed",
)


class Spreadsheet:
    """
    Class for investigation levels spreadsheet
    """

    name = "Landauer client list and dose investigation limits.xlsx"
    path = FileManager.paths_from_proj_dir["levels_spreadsheet"]
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

        if subaccount_code not in df.index:
            raise KeyError(
                f"Subaccount code {subaccount_code} not found in sheet {selected_sheet}. Please check whether it exists!"
            )

        row = df.loc[subaccount_code]

        cls.contact = row.values[-2]
        cls.CC = row.values[-1]

        if not (isinstance(cls.contact, str) and "@" in cls.contact):
            raise ValueError("Contact email for subaccount not found!")
        if not (isinstance(cls.CC, str) and "@" in cls.CC):
            raise ValueError("Site lead CC not found!")

        if df.loc[subaccount_code].str.contains("DEFAULT", case=False).any():
            row = df.loc[account_code]

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
        slice_end = df.iloc[0].str.contains("CC", case=False).idxmax() + 1
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
