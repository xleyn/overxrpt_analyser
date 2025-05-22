from dataclasses import dataclass, field
from datetime import datetime

from spreadsheet import Spreadsheet


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
        if self.dose == "M":
            self.dose = "0"
        self.period, self.period_type = self.get_temporal_data()
        if self.pull_levels:
            self.level, self.urgentLevel = Spreadsheet.get_levels(
                self.badge, self.period_type
            )

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
            True: lambda: self.gr_if_ytd_already_raised(),
            False: lambda: f"For wearer {self.name}, the year to date alert on this report can be ignored as the annual formal investigation level is {self.level} mSv on their {self.badge} badge.",
        }

        # Evaluate the condition and call the appropriate lambda
        return localResponse[float(self.dose) >= float(self.level)]()

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
        if self.past_notifs:

            for pn in self.past_notifs:
                conditions = [
                    pn.name == self.name,
                    pn.badge == self.badge,
                    pn.period_type == self.period_type,
                    pn.period == self.period,
                    float(pn.dose) >= float(self.urgentLevel),
                ]

                if all(conditions):

                    return True

            return False

        else:
            return float(self.dose) >= float(self.urgentLevel)

    def analyse(self):
        """Performs analysis for particular row by comparing to investigation levels.

        Returns:
            None.

        """
        self.attachTemplate = False
        get_response = self.period_type2method_mapping[self.period_type]
        self.response = get_response(self)
        self.urgent_flag = (
            self.urgent_flag_query()
            if float(self.dose) >= float(self.urgentLevel)
            else False
        )
