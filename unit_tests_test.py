import io
import pytest
from openpyxl import load_workbook
from spreadsheet_server import SpreadsheetServer
from pathlib import Path
from unittest.mock import patch, MagicMock, create_autospec

@pytest.mark.asyncio
async def test_create_spreadsheet_excel(tmp_path):
    filename = "text.xlsx"
    spreadsheet = SpreadsheetServer()
    fake_path = MagicMock(spec=Path)
    fake_path.exists.return_value = False
    fake_path.__str__.return_value = f"/fake/path/{filename}"
    expected_path = tmp_path / filename
    expected_result = {
                "success": True,
                "filename": filename,
                "path": str(expected_path),
                "sheet_name": 'Sheet1',
                "original_sheet_name": None
            }
    with patch.object(spreadsheet, "_resolve_path", return_value=expected_path):
        actual_result = await spreadsheet.create_spreadsheet(filename=filename, format='xlsx', headers=None, sheet_name='Sheet1')

    assert actual_result == expected_result, f'Actual : {repr(actual_result)} is not matching the Expected: {repr(expected_result)}'

@pytest.mark.asyncio
async def test_create_spreadsheet_csv(tmp_path):
    filename = "text.csv"
    csv = SpreadsheetServer()
    fake_path = MagicMock(spec=Path)
    fake_path.exists.return_value = False
    fake_path.__str__.return_value = f"/fake/path/{filename}"
    expected_path = tmp_path / filename
    expected_result = {
        "success": True,
        "filename": filename,
        "path": str(expected_path)
    }
    with patch.object(csv, "_resolve_path", return_value=expected_path):
        actual_result = await csv.create_spreadsheet(filename=filename, format='csv', headers=None)

    assert actual_result == expected_result, f'Actual : {repr(actual_result)} is not matching the Expected: {repr(expected_result)}'


@pytest.mark.asyncio
async def test_rename_file(tmp_path):
    spreadsheet = SpreadsheetServer()
    current_fake_path = create_autospec(Path, instance=True)
    current_fake_path.exists.return_value = True
    current = "test.xlsx"
    current_fake_path.name = current
    new_fake_path = create_autospec(Path, instance=True)
    new = "test1.xlsx"
    new_fake_path.name = new
    new_fake_path.exists.return_value = False

    with patch.object(spreadsheet, "_resolve_path", side_effect=[current_fake_path, new_fake_path]):
        result = await spreadsheet.rename_file(old_filename=current, new_filename=new)

    assert result == {"success": True, "old": current, "new": new}
    current_fake_path.rename.assert_called_once_with(new_fake_path)
