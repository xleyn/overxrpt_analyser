class BadgeMapper:
    """Class for extracting data mapped from badge names"""

    wholeBody = ("collar", "other whole body", "chest", "waist")
    extremity = ("left finger", "right finger")
    lens = "lens"

    hier_Excel_wb = ["DDE", "WHOLE BODY"]
    hier_Excel_lens = ["LDE", "LENS", "EYES"]
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
