from pathlib import Path


class WebDriverException(Exception):
    pass


class WebDriverNotInstantiatedException(WebDriverException):
    pass


class WebDriverNotStartedException(WebDriverException):
    pass


class InvalidFileExtensionException(ValueError):
    def __init__(self, path: str | Path, expected: str | list[str], received: str):
        msg = f'Invalid file extension for file "{path}". Expected: "{expected}", received: "{received}"'
        super().__init__(msg)
