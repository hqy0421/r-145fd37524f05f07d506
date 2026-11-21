"""
gspread.exceptions
~~~~~~~~~~~~~~~~~~

Exceptions used in gspread.

"""

from typing import Any, Mapping, Optional

from requests import Response


class UnSupportedExportFormat(Exception):
    """Raised when export format is not supported."""


class GSpreadException(Exception):
    """A base class for gspread's exceptions."""


class WorksheetNotFound(GSpreadException):
    """Trying to open non-existent or inaccessible worksheet."""


class NoValidUrlKeyFound(GSpreadException):
    """No valid key found in URL."""


class IncorrectCellLabel(GSpreadException):
    """The cell label is incorrect."""


class InvalidInputValue(GSpreadException):
    """The provided values is incorrect."""


class APIError(GSpreadException):
    """Errors coming from the API itself,
    such as when we attempt to retrieve things that don't exist."""

    def __init__(self, response: Response):
        try:
            error = response.json()["error"]
        except Exception as e:
            # in case we failed to parse the error from the API
            # build an empty error object to notify the caller
            # and keep the exception raise flow running

            error = {
                "code": -1,
                "message": response.text,
                "status": "invalid JSON: '{}'".format(e),
            }

        super().__init__(error)
        self.response: Response = response
        self.error: Mapping[str, Any] = error
        self.code: int = self.error["code"]

    def __str__(self) -> str:
        return "{}: [{}]: {}".format(
            self.__class__.__name__, self.code, self.error["message"]
        )

    def __repr__(self) -> str:
        return self.__str__()

    def __reduce__(self) -> tuple:
        return self.__class__, (self.response,)


class SpreadsheetNotFound(GSpreadException):
    """Trying to open non-existent or inaccessible spreadsheet."""


class CsvImportError(GSpreadException):
    """Base class for errors raised while importing a CSV file.

    Every subclass other than :class:`~gspread.exceptions.CsvImportOutcomeUnknown`
    guarantees that the target worksheet was left untouched.

    :param str message: a human readable description of the failure.
    :param staging_sheet_id: (optional) id of the temporary worksheet used
        to stage the data, when it was created.
    """

    def __init__(
        self,
        message: str,
        staging_sheet_id: Optional[int] = None,
        staging_cleanup_succeeded: Optional[bool] = None,
    ):
        super().__init__(message)
        self.staging_sheet_id = staging_sheet_id
        self.staging_cleanup_succeeded = staging_cleanup_succeeded


class CsvParseError(CsvImportError):
    """The CSV payload could not be parsed.

    The target worksheet was not modified.
    """


class CsvTooLargeError(CsvImportError):
    """The CSV payload does not fit in the requested data area.

    The target worksheet was not modified.
    """


class CsvStagingError(CsvImportError):
    """The CSV payload could not be staged in a temporary worksheet.

    The target worksheet was not modified.
    """


class CsvCommitError(CsvImportError):
    """The final batch was explicitly rejected by the API.

    The target worksheet was not modified; ``staging_cleanup_succeeded``
    reports whether the temporary worksheet could be removed.
    """


class CsvImportOutcomeUnknown(CsvImportError):
    """The request that commits the import may or may not have been applied.

    This is raised when the connection broke (or the response could not be
    read) while sending the final, atomic batch update. The target worksheet
    must be considered to be in an unknown state: the import may have been
    fully applied, or not applied at all. It must not be reported as success.
    """
