import json
from file_manager import FileManager


class BadgeMapper:
    """Class for extracting data mapped from badge names"""

    with open(FileManager.paths_from_proj_dir["badge_groupings"]) as f:
        badge_groupings = json.load(f)

    hier_Excel_wb = ["DDE", "WHOLE BODY"]
    hier_Excel_lens = ["LDE", "LENS", "EYES"]
    hier_Excel_extrem = ["EXTREMITY"]

    badge_to_col_mapping = dict(
        zip(
            map(tuple, badge_groupings.values()),
            ["Whole Body", "Total Extremity", "Lens"],
        )
    )
    get_hierarchy_funcs = [
        lambda b, hierarchy=hier_Excel_wb: [b] + hierarchy,
        lambda b, hierarchy=hier_Excel_extrem: hierarchy,
        lambda b, hierarchy=hier_Excel_lens: hierarchy,
    ]
    badge_to_hierarchy_mapping = dict(
        zip(map(tuple, badge_groupings.values()), get_hierarchy_funcs)
    )

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
            (
                col
                for badges, col in cls.badge_to_col_mapping.items()
                if badge in badges
            ),
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
