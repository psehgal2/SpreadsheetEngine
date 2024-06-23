import context
import pytest
import sheets
from sheets.CellErrorType import CellErrorType
from sheets.CellError import CellError
from sheets.Workbook import Workbook
import decimal

@pytest.fixture
def workbook_with_notifier():
    wb = sheets.Workbook()
    wb.new_sheet()
    wb.notify_cells_changed(on_cells_changed)
    return wb

# Mock notification function to capture cell changes
changed_cells = []
def on_cells_changed(workbook, cells):
    changed_cells.extend(cells)

# Test for basic cell value change
def test_basic_cell_change(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")

    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
    assert changed_cells == []

# Test for cell change with formula
def test_cell_change_with_formula(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=(B2 + 1) * (B3 - 1)")
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "1")
    workbook_with_notifier.set_cell_contents("Sheet1", "B2", "=B1")
    workbook_with_notifier.set_cell_contents("Sheet1", "B3", "=B1")
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "-1")
    assert ("Sheet1", "b1") in changed_cells
    assert ("Sheet1", "b2") in changed_cells
    assert ("Sheet1", "b3") in changed_cells
    assert ("Sheet1", "a1") not in changed_cells

def test_cell_with_functions(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A2", "=TRUE")
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "10")
    workbook_with_notifier.set_cell_contents("Sheet1", "C1", "10")
    workbook_with_notifier.set_cell_contents("Sheet1", "A5", "=IF(A2, A1, C1)")
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A2", "=FALSE")
    assert set(changed_cells) == {("Sheet1", "a2")}

# Test for functions with if else which don't actually update the cell
# Test which circ ref but a plus 1 so the actual value does not change
    
def test_circ_ref_change_no_cell_notif(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=B1")
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "=A1")
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "=(A1 + 1)")
    assert set(changed_cells) == {("Sheet1", "b1")}
