import io
import pytest
from openpyxl import load_workbook
from spreadsheet_server import SpreadsheetServer
from pathlib import Path
from unittest.mock import patch, MagicMock

@pytest.mark.asyncio
async def test_create_spreadsheet_excel(tmp_path):
    spreadsheet = SpreadsheetServer()
    fake_path = MagicMock()
    fake_path.exists.return_value = False
    fake_path.__str__.return_value = "/fake/path/test.xlsx"
    expected_path = tmp_path / "test.xlsx"
    expected_result = {
                "success": True,
                "filename": 'test.xlsx',
                "path": str(expected_path),
                "sheet_name": 'Sheet1',
                "original_sheet_name": None
            }
    with patch.object(spreadsheet, "_resolve_path", return_value=expected_path):
        actual_result = await spreadsheet.create_spreadsheet(filename='test.xlsx', format='xlsx', headers=None, sheet_name='Sheet1')

    assert actual_result == expected_result, f'Actual : {repr(actual_result)} is not matching the Expected: {repr(expected_result)}'