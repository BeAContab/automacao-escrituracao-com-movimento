from pathlib import Path


class SpreadsheetReaderException(Exception):
    pass


class NoActiveSpreadsheetException(SpreadsheetReaderException):
    def __init__(self, spreadsheet_path: str | Path):
        msg = f'Could not find active spreadsheet in workbook "{spreadsheet_path}"'
        super().__init__(msg)


class SpreadsheetIsLockedException(SpreadsheetReaderException):
    def __init__(self, spreadsheet_path: str | Path):
        msg = f'Cannot open locked spreadsheet "{spreadsheet_path}"'
        super().__init__(msg)
