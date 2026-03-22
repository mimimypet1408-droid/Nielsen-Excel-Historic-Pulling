import datetime
import re

import pandas as pd


class MissingPeriodsColumnError(Exception):
    """Custom exception raised when the 'Periods' column is missing from the DataFrame."""

    pass


def get_min_max_nielsen_periods(df: pd.DataFrame) -> tuple[datetime.date, datetime.date]:
    """Returns the minimum and maximum Nielsen periods from the given DataFrame and date column.

    Args:
        df: A pandas DataFrame containing a column with Nielsen periods in the format "1 w/e MM/DD/YY".

    Returns:
        A tuple containing the minimum and maximum Nielsen periods as datetime.date objects.

    Raises:
        MissingPeriodsColumnError: If the 'Periods' column is missing from the DataFrame.
        ValueError: If any of the period strings are not in the expected format.
    """

    def _parse_period(date_str: str) -> datetime.date:
        match = re.match(r"1 w/e (\d{2}/\d{2}/\d{2})", date_str)
        if not match:
            raise ValueError(f"Invalid period format: {date_str}")
        return datetime.datetime.strptime(match.group(1), "%m/%d/%y").date()

    date_col = "Periods"
    try:
        periods = df[date_col].apply(_parse_period)
    except KeyError:
        raise MissingPeriodsColumnError(f"Missing '{date_col}' column in DataFrame")
    min_period = periods.min()
    max_period = periods.max()
    return min_period, max_period


def get_min_max_nielsen_periods_from_excel_file(
    file_path: str, **kwargs
) -> tuple[datetime.date, datetime.date]:
    """Reads the given Excel file and sheet, and returns the minimum and maximum Nielsen periods."""
    df = pd.read_excel(file_path, **kwargs)
    return get_min_max_nielsen_periods(df)
