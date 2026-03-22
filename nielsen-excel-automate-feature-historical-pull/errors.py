from typing import Optional


class FailureLimitExceededError(Exception):
    def __init__(self, failed_files: list[str], message=None):
        self.failed_files = failed_files
        if message is None:
            message = f"Failed to refresh {len(failed_files)} files"
        self.message = message

    def __str__(self):
        return f"{self.message}: {self.failed_files}"


class DataPlanLimitReachedError(Exception):
    def __init__(self, failed_files: Optional[list[str]] = None, message="Data plan limit reached"):
        self.failed_files = failed_files
        self.message = message

    def __str__(self):
        return self.message


class WeekEndingDateError(Exception):
    def __init__(self, file_name: str, expected_date: str, actual_date: str):
        self.file_name = file_name
        self.expected_date = expected_date
        self.actual_date = actual_date
        self.message = (
            f"Week ending date mismatch in file '{file_name}': "
            f"expected {expected_date}, got {actual_date}"
        )

    def __str__(self):
        return self.message
