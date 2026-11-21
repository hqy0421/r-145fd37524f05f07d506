"""Google Spreadsheets Python API"""

__version__ = "6.2.1"
__author__ = "Anton Burnashev"


from .auth import (
    api_key,
    authorize,
    oauth,
    oauth_from_dict,
    service_account,
    service_account_from_dict,
)
from .cell import Cell
from .client import Client
from .csv_import import CsvImportResult, import_csv, parse_csv
from .exceptions import (
    CsvCommitError,
    CsvImportError,
    CsvImportOutcomeUnknown,
    CsvParseError,
    CsvStagingError,
    CsvTooLargeError,
    GSpreadException,
    IncorrectCellLabel,
    NoValidUrlKeyFound,
    SpreadsheetNotFound,
    WorksheetNotFound,
)
from .http_client import BackOffHTTPClient, HTTPClient
from .spreadsheet import Spreadsheet
from .worksheet import ValueRange, Worksheet

from . import urls as urls
from . import utils as utils

__all__ = [
    # from .auth
    "api_key",
    "authorize",
    "oauth",
    "oauth_from_dict",
    "service_account",
    "service_account_from_dict",

    # from .cell
    "Cell",

    # from .client
    "Client",

    # from .http_client
    "BackOffHTTPClient",
    "HTTPClient",

    # from .spreadsheet
    "Spreadsheet",

    # from .worksheet
    "Worksheet",
    "ValueRange",

    # from .exceptions
    "CsvCommitError",
    "CsvImportError",
    "CsvImportOutcomeUnknown",
    "CsvParseError",
    "CsvStagingError",
    "CsvTooLargeError",
    "GSpreadException",
    "IncorrectCellLabel",
    "NoValidUrlKeyFound",
    "SpreadsheetNotFound",
    "WorksheetNotFound",

    # from .csv_import
    "CsvImportResult",
    "import_csv",
    "parse_csv",

    # full module imports
    "urls",
    "utils",
]
