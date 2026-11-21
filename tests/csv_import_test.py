"""Tests for :mod:`gspread.csv_import` and ``Worksheet.import_csv``.

These tests never reach the network: a fake HTTP client records the
requests issued by gspread and returns canned responses.
"""

import unittest
from typing import Any, Dict, List, Optional

from requests import Response

from gspread.csv_import import CsvImportResult, parse_csv
from gspread.exceptions import (
    APIError,
    CsvCommitError,
    CsvImportOutcomeUnknown,
    CsvParseError,
    CsvStagingError,
    CsvTooLargeError,
)
from gspread.http_client import HTTPClient
from gspread.worksheet import Worksheet

SPREADSHEET_ID = "spreadsheet-id"


class _FakeResponse:
    def __init__(self, payload: Optional[Dict[str, Any]] = None):
        self._payload = payload if payload is not None else {}
        self.ok = True

    def json(self) -> Dict[str, Any]:
        return self._payload


class _FakeHTTPClient(HTTPClient):
    """An ``HTTPClient`` that records requests instead of sending them."""

    def __init__(self) -> None:
        # Skip HTTPClient.__init__: no session or credentials are needed.
        self.values_updates: List[Dict[str, Any]] = []
        self.batch_updates: List[Dict[str, Any]] = []
        self._batch_response: Any = None
        self._batch_error: Optional[BaseException] = None

    def request(
        self,
        method: str,
        endpoint: str,
        params: Any = None,
        data: Any = None,
        json: Any = None,
        files: Any = None,
        headers: Any = None,
        default_serializer: Any = None,
    ) -> Any:
        if endpoint.endswith(":batchUpdate"):
            self.batch_updates.append(dict(json))
            if self._batch_error is not None:
                raise self._batch_error
            return _FakeResponse(self._batch_response)

        if "/values/" in endpoint:
            self.values_updates.append(
                {
                    "method": method,
                    "endpoint": endpoint,
                    "params": dict(params or {}),
                    "body": dict(json),
                }
            )
            return _FakeResponse({"updatedRange": endpoint})

        raise AssertionError("unexpected request: {} {}".format(method, endpoint))


class _StagingWorksheet(Worksheet):
    """A worksheet whose writes can be made to fail on demand."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self.update_error: Optional[BaseException] = None

    def update(self, *args: Any, **kwargs: Any) -> Any:
        if self.update_error is not None:
            raise self.update_error
        return super().update(*args, **kwargs)


class _FakeSpreadsheet:
    """Minimal stand-in for ``Spreadsheet`` used by ``Worksheet``."""

    def __init__(self, http_client: _FakeHTTPClient) -> None:
        self.id = SPREADSHEET_ID
        self.client = http_client
        self.next_sheet_id = 900
        self.created: List[_StagingWorksheet] = []
        self.deleted: List[int] = []
        self.fail_add = False
        self.fail_delete = False
        self.fail_update: Optional[BaseException] = None
        self.metadata_error: Optional[BaseException] = None

    def add_worksheet(self, title: str, rows: int, cols: int) -> _StagingWorksheet:
        if self.fail_add:
            raise RuntimeError("cannot create staging worksheet")

        self.next_sheet_id += 1
        properties = {
            "sheetId": self.next_sheet_id,
            "title": title,
            "index": 1,
            "gridProperties": {"rowCount": rows, "columnCount": cols},
        }
        worksheet = _StagingWorksheet(self, properties, SPREADSHEET_ID, self.client)
        worksheet.update_error = self.fail_update
        self.created.append(worksheet)
        return worksheet

    def del_worksheet_by_id(self, sheet_id: int) -> None:
        if self.fail_delete:
            raise RuntimeError("cannot delete staging worksheet")
        self.deleted.append(sheet_id)

    def fetch_sheet_metadata(self, params: Any = None) -> Dict[str, Any]:
        if self.metadata_error is not None:
            raise self.metadata_error

        sheets = [{"properties": {"sheetId": 42, "title": "Sheet1"}}]
        for staging in self.created:
            if staging.id not in self.deleted:
                sheets.append({"properties": {"sheetId": staging.id}})
        return {"sheets": sheets}


class CsvImportTest(unittest.TestCase):
    def setUp(self) -> None:
        self.http_client = _FakeHTTPClient()
        self.spreadsheet = _FakeSpreadsheet(self.http_client)
        properties = {
            "sheetId": 42,
            "title": "Sheet1",
            "index": 0,
            "gridProperties": {"rowCount": 100, "columnCount": 26},
        }
        self.worksheet = Worksheet(
            self.spreadsheet,  # type: ignore[arg-type]
            properties,
            SPREADSHEET_ID,
            self.http_client,
        )
        self.http_client._batch_response = {"replies": [{}, {}, {}]}

    # -- parsing ---------------------------------------------------------

    def test_parse_keeps_leading_zeros_and_formula_like_text(self):
        records = parse_csv("id,value\n007,=1+1\n")
        self.assertEqual(records, [["id", "value"], ["007", "=1+1"]])

    def test_parse_multiline_quoted_field_is_one_logical_record(self):
        records = parse_csv('name,note\n"a\nb",c\n')
        self.assertEqual(records, [["name", "note"], ["a\nb", "c"]])

    def test_parse_accepts_utf8_bytes_and_strips_bom(self):
        records = parse_csv("héllo,wörld\n".encode("utf-8"))
        self.assertEqual(records, [["héllo", "wörld"]])

        records = parse_csv(b"\xef\xbb\xbfname,city\n")
        self.assertEqual(records, [["name", "city"]])

    def test_parse_rejects_invalid_utf8(self):
        with self.assertRaises(CsvParseError):
            parse_csv(b"name,\xff\n")

    def test_parse_rejects_empty_payload(self):
        with self.assertRaises(CsvParseError):
            parse_csv("")

    def test_parse_rejects_malformed_quotes(self):
        with self.assertRaises(CsvParseError):
            parse_csv('a,"b\n')

    # -- data area validation -------------------------------------------

    def test_csv_larger_than_area_is_rejected_before_any_request(self):
        with self.assertRaises(CsvTooLargeError):
            self.worksheet.import_csv("a,b\nc,d\ne,f\n", "A1:B2")

        self.assertEqual(self.http_client.batch_updates, [])
        self.assertEqual(self.http_client.values_updates, [])
        self.assertEqual(self.spreadsheet.created, [])

    def test_record_with_too_many_columns_is_rejected(self):
        with self.assertRaises(CsvTooLargeError):
            self.worksheet.import_csv("a,b,c\n", "A1:B5")

        self.assertEqual(self.spreadsheet.created, [])

    def test_unbounded_or_foreign_range_is_rejected(self):
        with self.assertRaises(ValueError):
            self.worksheet.import_csv("a\n", "A1:B")

        with self.assertRaises(ValueError):
            self.worksheet.import_csv("a\n", "Other!A1")

        with self.assertRaises(ValueError):
            self.worksheet.import_csv("a\n", "A200")

    def test_parse_error_leaves_worksheet_untouched(self):
        with self.assertRaises(CsvParseError):
            self.worksheet.import_csv('a,"b\n', "A1:B2")

        self.assertEqual(self.spreadsheet.created, [])
        self.assertEqual(self.http_client.batch_updates, [])

    # -- staging ---------------------------------------------------------

    def test_records_are_staged_in_batches(self):
        csv_payload = "\n".join("row{},value{}".format(i, i) for i in range(1, 6))

        result = self.worksheet.import_csv(csv_payload, "B2:C10", batch_size=2)

        self.assertIsInstance(result, CsvImportResult)
        self.assertEqual(result.rows, 5)
        self.assertEqual(result.columns, 2)
        self.assertEqual(result.range_name, "B2:C10")

        staging = self.spreadsheet.created[0]
        self.assertEqual(staging.row_count, 5)
        self.assertEqual(staging.col_count, 2)

        self.assertEqual(len(self.http_client.values_updates), 3)
        bodies = [
            update["body"]["values"] for update in self.http_client.values_updates
        ]
        self.assertEqual([len(body) for body in bodies], [2, 2, 1])
        self.assertEqual(bodies[0][0], ["row1", "value1"])

        ranges = [update["endpoint"] for update in self.http_client.values_updates]
        self.assertTrue(ranges[0].endswith("A1%3AB2"), ranges[0])
        self.assertTrue(ranges[1].endswith("A3%3AB4"), ranges[1])
        self.assertTrue(ranges[2].endswith("A5%3AB5"), ranges[2])

    def test_staged_values_use_raw_input(self):
        self.worksheet.import_csv("007,=SUM(A1:A2)\n", "A1:B1")

        update = self.http_client.values_updates[0]
        self.assertEqual(update["params"]["valueInputOption"], "RAW")
        self.assertEqual(update["body"]["values"], [["007", "=SUM(A1:A2)"]])

    def test_staging_failure_keeps_target_and_cleans_up(self):
        self.spreadsheet.fail_update = RuntimeError("staging update failed")

        with self.assertRaises(CsvStagingError) as ctx:
            self.worksheet.import_csv("a\n", "A1")

        self.assertEqual(self.http_client.batch_updates, [])
        self.assertEqual(self.spreadsheet.deleted, [self.spreadsheet.created[0].id])
        self.assertTrue(ctx.exception.staging_cleanup_succeeded)

    def test_staging_creation_failure_keeps_target(self):
        self.spreadsheet.fail_add = True

        with self.assertRaises(CsvStagingError):
            self.worksheet.import_csv("a\n", "A1")

        self.assertEqual(self.http_client.batch_updates, [])
        self.assertEqual(self.http_client.values_updates, [])

    # -- commit ----------------------------------------------------------

    def test_commit_is_a_single_atomic_batch(self):
        self.worksheet.import_csv("a,b\nc,d\n", "B2:C5")

        self.assertEqual(len(self.http_client.batch_updates), 1)
        requests = self.http_client.batch_updates[0]["requests"]
        self.assertEqual(
            [list(request)[0] for request in requests],
            ["updateCells", "copyPaste", "deleteSheet"],
        )

        clear = requests[0]["updateCells"]
        self.assertEqual(clear["fields"], "userEnteredValue")
        self.assertEqual(
            clear["range"],
            {
                "sheetId": 42,
                "startRowIndex": 1,
                "endRowIndex": 5,
                "startColumnIndex": 1,
                "endColumnIndex": 3,
            },
        )

        copy = requests[1]["copyPaste"]
        self.assertEqual(copy["pasteType"], "PASTE_VALUES")
        self.assertEqual(copy["source"]["endRowIndex"], 2)
        self.assertEqual(copy["source"]["endColumnIndex"], 2)
        self.assertEqual(copy["destination"]["startRowIndex"], 1)
        self.assertEqual(copy["destination"]["endRowIndex"], 3)
        self.assertEqual(copy["destination"]["startColumnIndex"], 1)
        self.assertEqual(copy["destination"]["endColumnIndex"], 3)

        staging_id = self.spreadsheet.created[0].id
        self.assertEqual(requests[2]["deleteSheet"]["sheetId"], staging_id)

    def test_clear_covers_whole_area_so_old_leftovers_are_removed(self):
        # Two new rows inside a ten-row data area: rows 5 to 11 are
        # cleared as well, so older, longer content cannot survive.
        self.worksheet.import_csv("a\nb\n", "A1:B11")

        clear_range = self.http_client.batch_updates[0]["requests"][0]["updateCells"][
            "range"
        ]
        self.assertEqual(clear_range["startRowIndex"], 0)
        self.assertEqual(clear_range["endRowIndex"], 11)
        self.assertEqual(clear_range["endColumnIndex"], 2)

    def test_explicit_rejection_keeps_target_and_cleans_up_staging(self):
        response = Response()
        response.status_code = 400
        response._content = (
            b'{"error": {"code": 400, "message": "bad", "status": "BAD"}}'
        )
        self.http_client._batch_error = APIError(response)

        with self.assertRaises(CsvCommitError) as ctx:
            self.worksheet.import_csv("a\n", "A1")

        staging_id = self.spreadsheet.created[0].id
        self.assertEqual(self.spreadsheet.deleted, [staging_id])
        self.assertTrue(ctx.exception.staging_cleanup_succeeded)

    def test_server_error_reports_unknown_outcome(self):
        response = Response()
        response.status_code = 500
        response._content = (
            b'{"error": {"code": 500, "message": "oops", "status": "ERROR"}}'
        )
        self.http_client._batch_error = APIError(response)

        with self.assertRaises(CsvImportOutcomeUnknown):
            self.worksheet.import_csv("a\n", "A1")

    def test_connection_error_reports_unknown_outcome(self):
        self.http_client._batch_error = ConnectionError("connection reset")

        with self.assertRaises(CsvImportOutcomeUnknown):
            self.worksheet.import_csv("a\n", "A1")

    def test_unexpected_response_reports_unknown_outcome(self):
        self.http_client._batch_response = {}

        with self.assertRaises(CsvImportOutcomeUnknown):
            self.worksheet.import_csv("a\n", "A1")

    def test_default_range_is_a1(self):
        result = self.worksheet.import_csv("a,b\n")

        self.assertEqual(result.range_name, "A1:B1")
        clear_range = self.http_client.batch_updates[0]["requests"][0]["updateCells"][
            "range"
        ]
        self.assertEqual(clear_range["endRowIndex"], 1)
        self.assertEqual(clear_range["endColumnIndex"], 2)

    def test_ragged_records_are_padded_with_empty_cells(self):
        self.worksheet.import_csv("a,b\nc\n", "A1:B2")

        values = self.http_client.values_updates[0]["body"]["values"]
        self.assertEqual(values, [["a", "b"], ["c", ""]])

    def test_empty_payload_is_rejected(self):
        with self.assertRaises(CsvParseError):
            self.worksheet.import_csv("", "A1")

    def test_data_area_of_exact_csv_size_accepts_a_full_rectangle(self):
        # Every record has the widest width: the area is exactly full.
        self.worksheet.import_csv("a,b,c\nd,e,f\n", "A1:C2")

        copy = self.http_client.batch_updates[0]["requests"][1]["copyPaste"]
        self.assertEqual(copy["source"]["endRowIndex"], 2)
        self.assertEqual(copy["source"]["endColumnIndex"], 3)

    def test_area_larger_than_csv_clears_only_the_area(self):
        self.worksheet.import_csv("a\n", "B2:D4")

        clear_range = self.http_client.batch_updates[0]["requests"][0]["updateCells"][
            "range"
        ]
        self.assertEqual(
            clear_range,
            {
                "sheetId": 42,
                "startRowIndex": 1,
                "endRowIndex": 4,
                "startColumnIndex": 1,
                "endColumnIndex": 4,
            },
        )
        copy = self.http_client.batch_updates[0]["requests"][1]["copyPaste"]
        # Only the staged 1x1 block is copied.
        self.assertEqual(copy["destination"]["endRowIndex"], 2)
        self.assertEqual(copy["destination"]["endColumnIndex"], 2)

    def test_qualified_range_with_matching_title_is_accepted(self):
        result = self.worksheet.import_csv("a,b\n", "Sheet1!B2:C10")

        self.assertEqual(result.range_name, "B2:C10")

    def test_qualified_range_with_quoted_title_is_accepted(self):
        self.worksheet.spreadsheet.add_worksheet("Sheet1", 1, 1)  # unused
        self.worksheet._properties["title"] = "My Sheet"

        result = self.worksheet.import_csv("a\n", "'My Sheet'!A1:B1")

        self.assertEqual(result.range_name, "A1:B1")

    def test_reversed_range_is_normalized(self):
        result = self.worksheet.import_csv("a,b\nc,d\n", "D10:C9")

        self.assertEqual(result.range_name, "C9:D10")

    def test_non_string_range_is_rejected(self):
        with self.assertRaises(TypeError):
            self.worksheet.import_csv("a\n", 42)  # type: ignore[arg-type]

    def test_invalid_batch_size_is_rejected(self):
        with self.assertRaises(ValueError):
            self.worksheet.import_csv("a\n", "A1", batch_size=0)

        with self.assertRaises(TypeError):
            self.worksheet.import_csv("a\n", "A1", batch_size="10")  # type: ignore[arg-type]

    def test_long_payload_is_split_on_payload_size_not_only_row_count(self):
        # Two fields of ~1.2M characters each cannot share one request.
        long_value = "x" * 1_200_000
        payload = "{},a\n{},b\n".format(long_value, long_value)

        self.worksheet.import_csv(payload, "A1:B2", batch_size=1000)

        self.assertEqual(len(self.http_client.values_updates), 2)

    def test_bytes_with_bom_and_crlf_are_supported(self):
        payload = b"\xef\xbb\xbfname,value\r\n007,x\r\n"
        result = self.worksheet.import_csv(payload, "A1:B2")

        self.assertEqual(result.rows, 2)
        values = self.http_client.values_updates[0]["body"]["values"]
        self.assertEqual(values, [["name", "value"], ["007", "x"]])

    def test_quoting_keeps_equal_sign_text_literal(self):
        payload = '"=1+1",=SUM(A1:B1)\n'
        self.worksheet.import_csv(payload, "A1:B1")

        update = self.http_client.values_updates[0]
        self.assertEqual(update["params"]["valueInputOption"], "RAW")
        self.assertEqual(update["body"]["values"], [["=1+1", "=SUM(A1:B1)"]])

    def test_staging_cleanup_failure_is_reported(self):
        self.spreadsheet.fail_update = RuntimeError("staging update failed")
        self.spreadsheet.fail_delete = True

        with self.assertRaises(CsvStagingError) as ctx:
            self.worksheet.import_csv("a\n", "A1")

        self.assertFalse(ctx.exception.staging_cleanup_succeeded)
        self.assertEqual(ctx.exception.staging_sheet_id, self.spreadsheet.created[0].id)

    def test_commit_cleanup_failure_is_reported(self):
        response = Response()
        response.status_code = 403
        response._content = (
            b'{"error": {"code": 403, "message": "denied", "status": "DENIED"}}'
        )
        self.http_client._batch_error = APIError(response)
        self.spreadsheet.fail_delete = True

        with self.assertRaises(CsvCommitError) as ctx:
            self.worksheet.import_csv("a\n", "A1")

        self.assertFalse(ctx.exception.staging_cleanup_succeeded)

    def test_rejection_is_a_no_op_when_staging_still_exists(self):
        response = Response()
        response.status_code = 400
        response._content = (
            b'{"error": {"code": 400, "message": "bad", "status": "BAD"}}'
        )
        self.http_client._batch_error = APIError(response)

        with self.assertRaises(CsvCommitError) as ctx:
            self.worksheet.import_csv("a\n", "A1")

        self.assertTrue(ctx.exception.staging_cleanup_succeeded)
        self.assertEqual(self.spreadsheet.deleted, [self.spreadsheet.created[0].id])

    def test_rejection_after_an_applied_retry_is_reported_as_unknown(self):
        # Simulate a retrying HTTP client: the batch was applied (so the
        # staging sheet is gone) but the last answer was a rejection.
        response = Response()
        response.status_code = 409
        response._content = (
            b'{"error": {"code": 409, "message": "conflict", "status": "CONFLICT"}}'
        )
        self.http_client._batch_error = APIError(response)
        self.spreadsheet.deleted = [901]

        with self.assertRaises(CsvImportOutcomeUnknown):
            self.worksheet.import_csv("a\n", "A1")

    def test_rejection_with_unverifiable_staging_is_reported_as_unknown(self):
        # If the staging worksheet cannot be inspected, a retrying client
        # may already have applied the batch, so a rejection must not be
        # trusted as a no-op.
        response = Response()
        response.status_code = 400
        response._content = (
            b'{"error": {"code": 400, "message": "bad", "status": "BAD"}}'
        )
        self.http_client._batch_error = APIError(response)
        self.spreadsheet.metadata_error = RuntimeError("metadata unavailable")

        with self.assertRaises(CsvImportOutcomeUnknown):
            self.worksheet.import_csv("a\n", "A1")

    def test_unknown_outcome_exposes_staging_sheet_id(self):
        self.http_client._batch_error = TimeoutError("timed out")

        with self.assertRaises(CsvImportOutcomeUnknown) as ctx:
            self.worksheet.import_csv("a\n", "A1")

        self.assertEqual(ctx.exception.staging_sheet_id, self.spreadsheet.created[0].id)

        # The commit batch was attempted exactly once and never reported
        # success, so the caller knows the outcome is unknown.
        self.assertEqual(len(self.http_client.batch_updates), 1)


if __name__ == "__main__":
    unittest.main()
