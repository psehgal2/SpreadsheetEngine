import context
import pytest
import sheets
from sheets.CellErrorType import CellErrorType
from sheets.CellError import CellError
from sheets.Workbook import Workbook
import decimal

@pytest.fixture
def wb():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    return wb


def test_copy_sheet():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=Sheet1_1!A1 + 1")
    wb.set_cell_contents("Sheet1", "A2", "=A3")
    wb.set_cell_contents("Sheet1", "A3", "10")

    original_value = wb.get_cell_value("Sheet1", "A1")

    print(wb.graph.children_to_parents)
    print(wb.get_cell_value("Sheet1", "A1"))

    index, name = wb.copy_sheet("Sheet1")
    assert wb.list_sheets() == ["Sheet1", "Sheet1_1"]
    
    print(wb.graph.children_to_parents)
    print(wb.get_cell_value("Sheet1", "A1"))
    print(wb.get_cell_value("Sheet1_1", "A1"))

    # wb.set_cell_contents("Sheet1", "A1", "20")
    copied_value = wb.get_cell_value(name, "A1")
    assert original_value != copied_value

def test_pro_copy_massive_sheet2():
    # Specific case where no sheet is referencing this sheet
    # lets us cache our results
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet", "A1", "1")
    for i in range(1, 9999):
        wb.set_cell_contents("test_sheet", f"A{i+1}", f"=A{i}+1")
    wb.copy_sheet("test_sheet")
    assert wb.get_cell_value("test_sheet_1", "A9999") == 9999
