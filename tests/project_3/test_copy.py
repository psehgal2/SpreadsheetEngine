import context
import pytest
import sheets
import sheets.CellErrorType as CellErrorType

@pytest.fixture
def wb():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    return wb


def test_move_cells_simple(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "A2", "World")
    wb.move_cells("Sheet1", "A1", "A2", "A3", 'Sheet1')
    assert wb.get_cell_contents("Sheet1", "A3") == "Hello"
    assert wb.get_cell_contents("Sheet1", "A4") == "World"
