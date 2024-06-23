
import context
import pytest
import sheets
import sheets.CellErrorType as CellErrorType

@pytest.fixture
def get_workbook():
    wb = sheets.Workbook()
    wb.new_sheet()
    return wb

def test_basic_cell_change(get_workbook):
    wb = get_workbook
    wb.set_cell_contents("Sheet1", "A1", "123")
    assert wb.get_cell_contents("Sheet1", "A1") == "123"

def test_basic_formula(get_workbook):
    wb = get_workbook
    wb.set_cell_contents("Sheet1", "A1", "=123")
    assert wb.get_cell_contents("Sheet1", "A1") == "=123"
    assert wb.get_cell_value("Sheet1", "A1") == 123

def test_3_cycle(get_workbook):
    wb = get_workbook
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=C1")
    wb.set_cell_contents("Sheet1", "C1", "=A1")
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "C1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_update_3_cycle(get_workbook):
    wb = get_workbook
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=C1")
    wb.set_cell_contents("Sheet1", "C1", "=A1")

    wb.set_cell_contents("Sheet1", "A1", "123")
    assert wb.get_cell_value("Sheet1", "A1") == 123
    assert wb.get_cell_value("Sheet1", "B1") == 123
    assert wb.get_cell_value("Sheet1", "C1") == 123

def test_works_with_equals(get_workbook):    
    wb = get_workbook
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=C1")
    wb.set_cell_contents("Sheet1", "C1", "=A1")

    wb.set_cell_contents("Sheet1", "A1", "=123")
    assert wb.get_cell_value("Sheet1", "A1") == 123
    assert wb.get_cell_value("Sheet1", "B1") == 123
    assert wb.get_cell_value("Sheet1", "C1") == 123
