"""
gspread.csv_import
~~~~~~~~~~~~~~~~~~

Helpers to import a UTF-8 CSV payload into a bounded A1 range of an
existing worksheet without destroying the worksheet (or the rest of the
spreadsheet).

The import works in three phases:

1. the payload is parsed and checked locally, before any API call, so a
   malformed or oversized CSV never changes the target worksheet;
2. the records are written to a temporary worksheet in batches;
3. a single ``spreadsheets.batchUpdate`` request clears the target range,
   copies the staged values into it and deletes the temporary worksheet.

The last request is atomic: either every part of it is applied or none of
it is. If its outcome cannot be established with certainty (for example a
connection dropped before the response was read), a
:class:`~gspread.exceptions.CsvImportOutcomeUnknown` exception is raised
instead of reporting success.
"""

import csv
import io
import re
import uuid
from dataclasses import dataclass
from typing import (
    TYPE_CHECKING,
    Any,
    Dict,
    Iterator,
    List,
    Mapping,
    Optional,
    Tuple,
    Union,
)

from .exceptions import (
    APIError,
    CsvCommitError,
    CsvImportOutcomeUnknown,
    CsvParseError,
    CsvStagingError,
    CsvTooLargeError,
)
from .utils import (
    PasteOrientation,
    PasteType,
    ValueInputOption,
    a1_to_rowcol,
    rowcol_to_a1,
)

if TYPE_CHECKING is True:
    from .worksheet import Worksheet

CsvData = Union[str, bytes, bytearray]
CsvRecord = List[str]

#: Default number of CSV logical records written to the staging worksheet
#: in a single ``values.update`` call.
DEFAULT_BATCH_SIZE = 1000

# Keep each ``values.update`` payload comfortably below the API limit.
_MAX_BATCH_CHARS = 2_000_000

_TARGET_RANGE_RE = re.compile(r"^([A-Za-z]+)([1-9]\d*)(?::([A-Za-z]+)([1-9]\d*))?$")
_STAGING_TITLE_PREFIX = "_gspread_csv_staging_"


@dataclass(frozen=True)
class CsvImportResult:
    """Details of a successfully committed CSV import."""

    worksheet_id: int
    spreadsheet_id: str
    range_name: str
    rows: int
    columns: int
    response: Mapping[str, Any]


@dataclass(frozen=True)
class _DataArea:
    """A normalized, bounded rectangular area of a worksheet.

    Rows and columns are 1-based and inclusive, like the rest of the
    public gspread API.
    """

    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def rows(self) -> int:
        return self.end_row - self.start_row + 1

    @property
    def columns(self) -> int:
        return self.end_col - self.start_col + 1

    @property
    def range_name(self) -> str:
        return "{}:{}".format(
            rowcol_to_a1(self.start_row, self.start_col),
            rowcol_to_a1(self.end_row, self.end_col),
        )

    def grid_range(self, sheet_id: int) -> Dict[str, int]:
        """Return the Sheets API ``GridRange`` for this area."""
        return {
            "sheetId": sheet_id,
            "startRowIndex": self.start_row - 1,
            "endRowIndex": self.end_row,
            "startColumnIndex": self.start_col - 1,
            "endColumnIndex": self.end_col,
        }


def parse_csv(data: CsvData) -> List[CsvRecord]:
    """Parse a UTF-8 CSV payload into logical records.

    Parsing follows normal CSV rules: quoted fields may contain commas,
    doubled quotes and line breaks, so a single logical record may span
    several physical lines. Fields are returned as text and are never
    converted to numbers, which keeps leading zeros intact.

    :param data: CSV payload as ``str`` or UTF-8 ``bytes``. An optional
        UTF-8 BOM is discarded.
    :returns: a list of records, each record being a list of strings.
    :raises CsvParseError: if the payload is empty, is not valid UTF-8 or
        is not valid CSV.
    """
    text = _decode_csv(data)

    if text == "":
        raise CsvParseError("CSV payload is empty")

    try:
        records = _read_records(text)
    except csv.Error as err:
        raise CsvParseError("Cannot parse CSV payload: {}".format(err)) from err

    if not records:
        raise CsvParseError("CSV payload contains no records")

    if max(len(record) for record in records) == 0:
        raise CsvParseError("CSV payload contains no fields")

    return records


def import_csv(
    worksheet: "Worksheet",
    data: CsvData,
    range_name: Optional[str] = None,
    batch_size: int = DEFAULT_BATCH_SIZE,
) -> CsvImportResult:
    """Import CSV data into a bounded A1 range of ``worksheet``.

    This is the implementation behind
    :meth:`gspread.worksheet.Worksheet.import_csv`.

    :param worksheet: the worksheet to import into.
    :param data: CSV payload as ``str`` or UTF-8 ``bytes``.
    :param range_name: (optional) bounded A1 range such as ``"B2:F20"``,
        optionally prefixed with the worksheet's title. Defaults to
        ``"A1"``. The range must not be unbounded (``"A1:F"`` is
        rejected) and must exist in the worksheet's grid.
    :param batch_size: number of logical records written per API call
        while staging. Must be a positive integer.
    :returns: a :class:`CsvImportResult` describing the committed import.
    :raises CsvParseError: the payload cannot be parsed. The target
        worksheet is untouched.
    :raises CsvTooLargeError: the payload does not fit in the requested
        data area. The target worksheet is untouched.
    :raises CsvStagingError: the values could not be staged. The target
        worksheet is untouched.
    :raises CsvCommitError: the final batch was explicitly rejected by
        the API. The target worksheet is untouched.
    :raises CsvImportOutcomeUnknown: the final batch may or may not have
        been applied. The target worksheet's state is uncertain and the
        import must not be reported as a success.
    """
    if not isinstance(batch_size, int) or isinstance(batch_size, bool):
        raise TypeError("batch_size must be an integer")

    if batch_size < 1:
        raise ValueError("batch_size must be greater than zero")

    records = parse_csv(data)
    area = _resolve_data_area(worksheet, range_name, records)
    _validate_fit(records, area)

    column_count = max(len(record) for record in records)
    rows = [record + [""] * (column_count - len(record)) for record in records]

    staging = _create_staging_worksheet(worksheet, len(rows), column_count)

    try:
        _write_staging_values(staging, rows, batch_size)
    except Exception as err:
        cleaned_up = _delete_staging_worksheet(worksheet, staging)
        message = (
            "Could not stage CSV data in worksheet {!r}; "
            "the target worksheet was not modified.".format(staging.title)
        )
        if not cleaned_up:
            message += (
                " The staging worksheet (id {}) could not be deleted "
                "and may still be present.".format(staging.id)
            )
        raise CsvStagingError(
            message,
            staging_sheet_id=staging.id,
            staging_cleanup_succeeded=cleaned_up,
        ) from err

    return _commit(worksheet, staging, area, len(rows), column_count)


def _read_records(text: str) -> List[CsvRecord]:
    """Read all records, temporarily lifting the field size limit.

    ``csv`` rejects fields longer than ``csv.field_size_limit()`` (128 KiB
    by default), which would refuse valid CSV payloads. The limit is
    restored afterwards so the global state is left as found.
    """
    previous_limit = csv.field_size_limit()
    csv.field_size_limit(max(previous_limit, len(text)))
    try:
        return list(csv.reader(io.StringIO(text, newline=""), strict=True))
    finally:
        csv.field_size_limit(previous_limit)


def _decode_csv(data: CsvData) -> str:
    if hasattr(data, "read"):
        # Support file-like objects as a convenience.
        data = data.read()  # type: ignore[union-attr]

    if isinstance(data, (bytes, bytearray)):
        try:
            return bytes(data).decode("utf-8-sig")
        except UnicodeDecodeError as err:
            raise CsvParseError(
                "CSV payload is not valid UTF-8: {}".format(err)
            ) from err

    if isinstance(data, str):
        # ``utf-8-sig`` only strips a BOM when decoding bytes, so handle
        # the text case explicitly as well.
        return data[1:] if data.startswith("\ufeff") else data

    raise TypeError(
        "CSV payload must be str, bytes, bytearray or a file-like object, "
        "got {}".format(type(data).__name__)
    )


def _resolve_data_area(
    worksheet: "Worksheet", range_name: Optional[str], records: List[CsvRecord]
) -> _DataArea:
    if range_name is not None:
        return _parse_data_area(worksheet, range_name)

    # No explicit data area: the payload is imported at A1 and the area is
    # exactly the size of the CSV.
    rows = len(records)
    columns = max(len(record) for record in records)
    area = _DataArea(1, 1, rows, columns)
    _check_grid_bounds(worksheet, area)
    return area


def _parse_data_area(worksheet: "Worksheet", range_name: Optional[str]) -> _DataArea:
    if range_name is None:
        range_name = "A1"

    if not isinstance(range_name, str):
        raise TypeError("range_name must be a string or None")

    range_name = range_name.strip()

    if "!" in range_name:
        sheet_title, _, range_name = range_name.rpartition("!")
        if sheet_title.startswith("'") and sheet_title.endswith("'"):
            sheet_title = sheet_title[1:-1].replace("''", "'")
        if sheet_title != worksheet.title:
            raise ValueError(
                "range_name {!r} does not belong to worksheet {!r}".format(
                    range_name, worksheet.title
                )
            )

    match = _TARGET_RANGE_RE.match(range_name)
    if match is None:
        raise ValueError(
            "range_name must be a bounded A1 range such as 'A1:C10', got {!r}".format(
                range_name
            )
        )

    start_column_letters, start_row_text, end_column_letters, end_row_text = (
        match.groups()
    )

    start_row, start_col = a1_to_rowcol(start_column_letters + start_row_text)

    if end_column_letters is None:
        end_row, end_col = start_row, start_col
    else:
        end_row, end_col = a1_to_rowcol(end_column_letters + end_row_text)

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col > end_col:
        start_col, end_col = end_col, start_col

    area = _DataArea(start_row, start_col, end_row, end_col)
    _check_grid_bounds(worksheet, area)

    return area


def _check_grid_bounds(worksheet: "Worksheet", area: _DataArea) -> None:
    if area.end_row > worksheet.row_count:
        raise ValueError(
            "data area {!r} ends at row {}, beyond worksheet row count {}".format(
                area.range_name, area.end_row, worksheet.row_count
            )
        )

    if area.end_col > worksheet.col_count:
        raise ValueError(
            "data area {!r} ends at column {}, beyond worksheet column count {}".format(
                area.range_name, area.end_col, worksheet.col_count
            )
        )


def _validate_fit(records: List[CsvRecord], area: _DataArea) -> None:
    if len(records) > area.rows:
        raise CsvTooLargeError(
            "CSV has {} logical records but the data area {!r} only has "
            "{} row(s)".format(len(records), area.range_name, area.rows)
        )

    for record_number, record in enumerate(records, start=1):
        if len(record) > area.columns:
            raise CsvTooLargeError(
                "CSV logical record {} has {} field(s) but the data area "
                "{!r} only has {} column(s)".format(
                    record_number, len(record), area.range_name, area.columns
                )
            )


def _create_staging_worksheet(
    worksheet: "Worksheet", rows: int, columns: int
) -> "Worksheet":
    title = _STAGING_TITLE_PREFIX + uuid.uuid4().hex[:16]

    try:
        return worksheet.spreadsheet.add_worksheet(title, rows, columns)
    except Exception as err:
        raise CsvStagingError(
            "Could not create a staging worksheet; the target worksheet "
            "was not modified."
        ) from err


def _write_staging_values(
    staging: "Worksheet", rows: List[CsvRecord], batch_size: int
) -> None:
    column_count = len(rows[0])

    for offset, chunk in _iter_batches(rows, batch_size):
        first_cell = rowcol_to_a1(offset + 1, 1)
        last_cell = rowcol_to_a1(offset + len(chunk), column_count)

        range_name = (
            first_cell
            if first_cell == last_cell
            else "{}:{}".format(first_cell, last_cell)
        )

        staging.update(
            chunk,
            range_name=range_name,
            value_input_option=ValueInputOption.raw,
        )


def _iter_batches(
    rows: List[CsvRecord], batch_size: int
) -> Iterator[Tuple[int, List[CsvRecord]]]:
    """Yield ``(offset, chunk)`` pairs bounded by row count and payload size."""
    offset = 0
    chunk: List[CsvRecord] = []
    chunk_chars = 0

    for index, row in enumerate(rows):
        row_chars = sum(len(value) + 3 for value in row)

        limit_reached = len(chunk) >= batch_size or (
            bool(chunk) and chunk_chars + row_chars > _MAX_BATCH_CHARS
        )

        if limit_reached:
            yield offset, chunk
            offset = index
            chunk = []
            chunk_chars = 0

        chunk.append(row)
        chunk_chars += row_chars

    if chunk:
        yield offset, chunk


def _delete_staging_worksheet(worksheet: "Worksheet", staging: "Worksheet") -> bool:
    try:
        worksheet.spreadsheet.del_worksheet_by_id(staging.id)
    except Exception:
        return False
    return True


def _staging_sheet_exists(
    worksheet: "Worksheet", staging: "Worksheet"
) -> Optional[bool]:
    """Return whether the staging worksheet still exists, if determinable.

    The staging worksheet is deleted by the same atomic batch that applies
    the values, so its presence is the only independent evidence that the
    batch was not applied (for example when a retrying HTTP client turned
    an already applied batch into a misleading rejection).
    """
    try:
        metadata = worksheet.spreadsheet.fetch_sheet_metadata()
        sheets = metadata.get("sheets", [])
        return any(
            sheet.get("properties", {}).get("sheetId") == staging.id for sheet in sheets
        )
    except Exception:
        return None


def _commit(
    worksheet: "Worksheet",
    staging: "Worksheet",
    area: _DataArea,
    rows: int,
    columns: int,
) -> CsvImportResult:
    requests = [
        # Clear all old values of the data area without touching its
        # formatting. ``updateCells`` clears the whole requested range for
        # the fields in the mask when ``rows`` does not cover it, which
        # also removes leftovers when the previous content was larger than
        # the imported CSV.
        {
            "updateCells": {
                "range": area.grid_range(worksheet.id),
                "fields": "userEnteredValue",
            }
        },
        # Copy the staged values only, so the target cells keep their
        # original formatting.
        {
            "copyPaste": {
                "source": {
                    "sheetId": staging.id,
                    "startRowIndex": 0,
                    "endRowIndex": rows,
                    "startColumnIndex": 0,
                    "endColumnIndex": columns,
                },
                "destination": {
                    "sheetId": worksheet.id,
                    "startRowIndex": area.start_row - 1,
                    "endRowIndex": area.start_row - 1 + rows,
                    "startColumnIndex": area.start_col - 1,
                    "endColumnIndex": area.start_col - 1 + columns,
                },
                "pasteType": PasteType.values,
                "pasteOrientation": PasteOrientation.normal,
            }
        },
        # Delete the staging worksheet as part of the same atomic batch.
        {"deleteSheet": {"sheetId": staging.id}},
    ]

    try:
        response = worksheet.client.batch_update(
            worksheet.spreadsheet_id, {"requests": requests}
        )
    except APIError as err:
        if 400 <= err.code < 500:
            # The API validated and rejected the batch as a whole, so
            # nothing was applied. The staging worksheet, deleted by the
            # same batch, is an independent check: if it disappeared, the
            # batch was applied anyway (for example by a retrying HTTP
            # client), and the rejection must not be trusted.
            staging_exists = _staging_sheet_exists(worksheet, staging)

            if staging_exists is False:
                raise CsvImportOutcomeUnknown(
                    "The CSV import batch was reported as rejected ({}) but "
                    "its staging worksheet (id {}) no longer exists, which "
                    "means the batch was applied after all; the target "
                    "worksheet's state is uncertain.".format(err, staging.id),
                    staging_sheet_id=staging.id,
                ) from err

            cleaned_up = False
            if staging_exists is True:
                cleaned_up = _delete_staging_worksheet(worksheet, staging)
            elif staging_exists is None:
                # The staging worksheet's presence could not be checked, so
                # a retrying client may have applied the batch already. Do
                # not claim the target was left untouched.
                raise CsvImportOutcomeUnknown(
                    "The CSV import batch was rejected by the API ({}) but "
                    "whether its staging worksheet (id {}) still exists "
                    "could not be checked; the target worksheet's state is "
                    "uncertain.".format(err, staging.id),
                    staging_sheet_id=staging.id,
                ) from err

            message = (
                "The CSV import batch was rejected by the API ({}); "
                "the target worksheet was not modified.".format(err)
            )
            if not cleaned_up:
                message += (
                    " The staging worksheet (id {}) could not be deleted "
                    "and may still be present.".format(staging.id)
                )

            raise CsvCommitError(
                message,
                staging_sheet_id=staging.id,
                staging_cleanup_succeeded=cleaned_up,
            ) from err

        # A server-side error does not prove the request was not applied.
        raise CsvImportOutcomeUnknown(
            "The CSV import batch may or may not have been applied ({}); "
            "the target worksheet's state is uncertain. The staging "
            "worksheet (id {}) is deleted by the same atomic batch, so "
            "check whether it still exists before retrying.".format(err, staging.id),
            staging_sheet_id=staging.id,
        ) from err
    except Exception as err:
        raise CsvImportOutcomeUnknown(
            "The CSV import batch may or may not have been applied ({}); "
            "the target worksheet's state is uncertain. The staging "
            "worksheet (id {}) is deleted by the same atomic batch, so "
            "check whether it still exists before retrying.".format(err, staging.id),
            staging_sheet_id=staging.id,
        ) from err

    if (
        not isinstance(response, dict)
        or not isinstance(response.get("replies"), list)
        or len(response["replies"]) != len(requests)
    ):
        raise CsvImportOutcomeUnknown(
            "The CSV import batch returned an unexpected response, so its "
            "outcome is uncertain; the target worksheet must not be "
            "considered successfully updated.",
            staging_sheet_id=staging.id,
        )

    return CsvImportResult(
        worksheet_id=worksheet.id,
        spreadsheet_id=worksheet.spreadsheet_id,
        range_name=area.range_name,
        rows=rows,
        columns=columns,
        response=response,
    )
