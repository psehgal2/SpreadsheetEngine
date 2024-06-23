import context
import pytest
import sheets
from sheets.CellErrorType import CellErrorType
from sheets.CellError import CellError
import decimal

@pytest.fixture
def wb():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    return wb
#[-5] A copied spreadsheet might have a bad name
# error that references a future copy. When this copy is made it could cause a circular reference error
def test_copy_sheet_circ_ref():
    workbook = sheets.Workbook()
    workbook.new_sheet("Sheet1")
    workbook.set_cell_contents("Sheet1", "A1", '=Sheet1_1!A1')
    assert workbook.get_cell_value("Sheet1", "A1").get_type() ==  CellErrorType.BAD_REFERENCE
    workbook.copy_sheet("Sheet1")
    assert workbook.get_cell_value("Sheet1_1", "A1").get_type()  == CellErrorType.CIRCULAR_REFERENCE
    assert workbook.get_cell_value("Sheet1", "A1").get_type()  == CellErrorType.CIRCULAR_REFERENCE


#[-2] Test multiplication and division with an unset cell (should become 0)

def test_donnie_example_1(wb):
    wb.set_cell_contents("Sheet1", "A1", "=-3*Q1")
    print(wb.get_cell_value("Sheet1", "A1"))
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(0)

def test_donnie_example_2(wb):
    wb.set_cell_contents("Sheet1", "A1", "=-3.5*Q1")
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(0)

def test_donnie_example_3(wb):
    wb.set_cell_contents("Sheet1", "A1", "=Q1/-3.5")
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(0)

def test_donnie_example_4(wb):
    wb.set_cell_contents("Sheet1", "A1", "=Q1/3.5")
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(0)

def test_multiply_with_unset_cell(wb):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "B1", "=A1*B2")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(0)

def test_multiply_with_unset_cell_decimal(wb):
    wb.set_cell_contents("Sheet1", "A1", "=5.2 * B2")
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(0)

def test_multiply_with_unset_cell_negative(wb):
    wb.set_cell_contents("Sheet1", "A1", "-5")
    wb.set_cell_contents("Sheet1", "B1", "=A1*B2")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(0)

def test_multiply_with_unset_cell_negative_on_cell(wb):
    wb.set_cell_contents("Sheet1", "A1", "=-5* B2")
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(0)

def test_divide_with_unset_cell(wb):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "B1", "=A1/B2")
    assert isinstance(wb.get_cell_value("Sheet1", "B1"),CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.DIVIDE_BY_ZERO

def test_divide_unset_cell(wb):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "B1", "=B2/A1")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(0)

def test_divide_with_unset_cell_negative(wb):
    wb.set_cell_contents("Sheet1", "A1", "-5")
    wb.set_cell_contents("Sheet1", "B1", "=A1/B2")
    assert isinstance(wb.get_cell_value("Sheet1", "B1"),CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.DIVIDE_BY_ZERO

def test_divide_unset_cell_negative(wb):
    wb.set_cell_contents("Sheet1", "A1", "-5")
    wb.set_cell_contents("Sheet1", "B1", "=B2/A1")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(0)

def test_divide_with_unset_cell_decimal(wb):
    wb.set_cell_contents("Sheet1", "A1", "5.2")
    wb.set_cell_contents("Sheet1", "B1", "=A1/B2")
    assert isinstance(wb.get_cell_value("Sheet1", "B1"),CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.DIVIDE_BY_ZERO

def test_divide_unset_cell_decimal(wb):
    wb.set_cell_contents("Sheet1", "A1", "5.2")
    wb.set_cell_contents("Sheet1", "B1", "=B2/A1")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(0)

# [-2] Test a variety of formulas that include multiple operators between both cells and numbers 
def test_multiple_ops(wb):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "A2", "10")
    wb.set_cell_contents("Sheet1", "B1", "=A1 + A2 * 2 - 3 / 2")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(23.5)

def test_multiple_ops2(wb):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "3")
    wb.set_cell_contents("Sheet1", "B1", "=A1 * 2 + A2 / 3 - 1")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(4)

def test_multiple_ops3(wb):
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "5")
    wb.set_cell_contents("Sheet1", "B1", "=A1 - A2 * 2 + 3 / 4")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(.75)

#[-2] Test a formula that multiplies a cell and a number, with various whitespaces. 
    

def test_wt_spc_donnie_example_1(wb):
    wb.set_cell_contents("Sheet1", "A1", "7.5")
    wb.set_cell_contents("Sheet1", "A2", "=A1*3.2")
    assert wb.get_cell_value("Sheet1", "A2") == decimal.Decimal(24)



def test_various_wt_spc_cell_number_multiply(wb):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "B1", "=A1 * 5")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B2", "=A1       * 5  ")
    assert wb.get_cell_value("Sheet1", "B2") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B3", "=A1*5")
    assert wb.get_cell_value("Sheet1", "B3") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B4", "=A1*     5 ")
    assert wb.get_cell_value("Sheet1", "B4") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B5", "=           A1*     5  ")
    assert wb.get_cell_value("Sheet1", "B5") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B6", "        =A1*5")
    assert wb.get_cell_value("Sheet1", "B6") == decimal.Decimal(25)


    wb.set_cell_contents("Sheet1", "A1", "    5    ")
    wb.set_cell_contents("Sheet1", "B1", "=A1 * 5")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B2", "=A1       * 5     ")
    assert wb.get_cell_value("Sheet1", "B2") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B3", "=A1*5")
    assert wb.get_cell_value("Sheet1", "B3") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B4", "=A1*     5     ")
    assert wb.get_cell_value("Sheet1", "B4") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B5", "=           A1*     5")
    assert wb.get_cell_value("Sheet1", "B5") == decimal.Decimal(25)
    wb.set_cell_contents("Sheet1", "B6", "         =    A1    *    5    ")
    assert wb.get_cell_value("Sheet1", "B6") == decimal.Decimal(25)

# Test Fixture to setup a Workbook with notification function
@pytest.fixture
def workbook_with_notifier():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.notify_cells_changed(on_cells_changed)
    return wb

# Mock notification function to capture cell changes
changed_cells = []
def on_cells_changed(workbook, cells):
    changed_cells.extend(cells)

# [-5] Test that when a sheet is copied, notifications are sent for all non- empty cells in the new sheet.
def test_copied_sheet(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "123")
    assert changed_cells == [("Sheet1", "a1")]
    index, name = workbook_with_notifier.copy_sheet("Sheet1")
    assert changed_cells == [("Sheet1", "a1"), ("Sheet1_1", "a1")]
 
 # [-5] Tests that the copied sheet gets the correct value for cells that have formulas that need evaluation when the formulas evaluate to errors. 
def test_copied_sheet_error(wb):
    wb.set_cell_contents("Sheet1", "A1", "=Sheet1_1!A1")
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE
    index, name = wb.copy_sheet("Sheet1")
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    wb.set_cell_contents("Sheet1_1", "A1", "123")
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(123)

# [-5] Loading empty JSON should result in an empty workbook
def test_loading_empty_to_empty():
    wb = sheets.Workbook()
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        wb.save_workbook(json_file)
    loaded_wb = sheets.Workbook()
    with open(json_file_path, "r") as json_file:
        loaded_wb = sheets.Workbook.load_workbook(json_file)
    print(loaded_wb.list_sheets())
    assert loaded_wb.list_sheets() == []

# [-5] Test that rename-sheet adds/removes quotes around sheet names in formulas as appropriate.
def test_rename_fixes_inaccurate_sheet_name():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.new_sheet("Sheet2")
    wb.new_sheet("Sheet4")

    wb.set_cell_contents("Sheet1","A1","='Sheet2'!A1 + 'Sheet4'!A1 + 1")
    wb.rename_sheet("Sheet2", "Sheet3")
    assert wb.get_cell_contents("Sheet1","A1") == "=Sheet3!A1+Sheet4!A1+1"
