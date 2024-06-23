import context
import pytest
import sheets
import sheets.CellErrorType as CellErrorType


@pytest.fixture
def wb():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    return wb



