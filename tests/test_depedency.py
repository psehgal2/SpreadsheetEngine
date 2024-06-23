import context

# Now we can import other things
import sheets
from tests.project_1.wrongNamesList import  wrongSheetNames
import pytest

from sheets.utils import get_hidden_name
from decimal import Decimal

from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType

def test_not_in_dependency():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(name, "A1", "1")
    assert wb.get_cell_value(name, "A1") == Decimal(1)
