import context
import pytest
import sheets
import sheets.CellErrorType as CellErrorType
from decimal import Decimal

@pytest.fixture
def wb():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    return wb

# 3
# 5 3
def test_no_shift_simple(wb):
    wb.set_cell_contents("Sheet1", "A1", "3")
    wb.set_cell_contents("Sheet1", "A2", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "5")
    wb.set_cell_contents("Sheet1", "B2", "=A1")
    wb.sort_region("Sheet1", "A1", "B2", [1,2])
    
    assert wb.get_cell_contents("Sheet1", "A1") == "3"
    assert wb.get_cell_contents("Sheet1", "A2") == "=B1"
    assert wb.get_cell_value("Sheet1", "A2") == Decimal('5')
    assert wb.get_cell_contents("Sheet1", "B1") == "5"
    assert wb.get_cell_contents("Sheet1", "B2") == "=A1"
    assert wb.get_cell_value("Sheet1", "B2") == Decimal('3')
# A1 B1
# 5 3
# =B1 =A1
def test_shift_simple(wb):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "A2", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "3")
    wb.set_cell_contents("Sheet1", "B2", "=A1")
    wb.sort_region("Sheet1", "A1", "B2", [1,2])
    
    assert wb.get_cell_contents("Sheet1", "A1") == "=B2"
    assert wb.get_cell_value("Sheet1", "A1") == Decimal('3')
    assert wb.get_cell_contents("Sheet1", "A2") == "5"
    assert wb.get_cell_value("Sheet1", "A2") == Decimal('5')
    assert wb.get_cell_contents("Sheet1", "B1") == "=A2"
    assert wb.get_cell_value("Sheet1", "B1") == Decimal('5')
    assert wb.get_cell_contents("Sheet1", "B2") == "3"
    assert wb.get_cell_value("Sheet1", "B2") == Decimal('3')