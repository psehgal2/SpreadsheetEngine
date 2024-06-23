
import context
import pytest
from sheets.CellError import CellError
from sheets.CellErrorType import *
import sheets
from sheets import Workbook
import decimal


def test_move_workbooks():
    wb = sheets.Workbook()
    wb.new_sheet("nil")
    wb.new_sheet("uno")
    wb.new_sheet("dos")
    wb.new_sheet("tres")
    wb.new_sheet("quatro")
    wb.new_sheet("cinco")
    wb.new_sheet("seis")
    wb.new_sheet("siete")
    wb.new_sheet("ocho")
    wb.move_sheet("tres",3)
    assert wb.list_sheets()[3] ==  "tres"
    wb.move_sheet("siete",3)
    assert wb.list_sheets()[3] == "siete"
    assert wb.list_sheets()[4] == "tres"
#
def test_rename_sheet_success():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.rename_sheet("Sheet1", "NewSheet1")
    assert "NewSheet1" in workbook.list_sheets()

def test_rename_nonexistent_sheet():
    workbook = Workbook()
    with pytest.raises(KeyError):
        workbook.rename_sheet("NonExistentSheet", "NewName")

def test_rename_sheet_invalid_name():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    with pytest.raises(ValueError):
        workbook.rename_sheet("Sheet1", "")

def test_rename_sheet_to_existing_name():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.new_sheet("Sheet2")
    with pytest.raises(ValueError):
        workbook.rename_sheet("Sheet1", "Sheet2")

def test_rename_sheet_updates_formulas():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.new_sheet("Sheet2")

    workbook.set_cell_contents("Sheet2", "A1", "='Sheet1'!A2")
    workbook.rename_sheet("Sheet1", "NewSheet1")
    # Assuming you have a method to get the raw formula of a cell
    assert workbook.get_cell_contents("Sheet2", "A1") == "=NewSheet1!A2"

def test_rename_sheet_leaves_literals_unchanged():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.new_sheet("Sheet2")
    workbook.set_cell_contents("Sheet1", "A1", "' bruh")

    workbook.set_cell_contents("Sheet2", "A1", "=\"Sheet1!A1\" & Sheet1!A1")
    # workbook.set_cell_contents("Sheet2", "A2", "='Sheet1!A2"'" & ")
    # workbook.set_cell_contents("Sheet2", "A3", "='Sheet1 bruh'")

    workbook.rename_sheet("Sheet1", "NewSheet1")
    # Assuming you have a method to get the raw formula of a cell
    assert workbook.get_cell_contents("Sheet2", "A1") == "=\"Sheet1!A1\" & NewSheet1!A1".replace(' ','')
    #why is the cell value here none?
    assert workbook.get_cell_value("Sheet2", "A1") == "Sheet1!A1 bruh"
    # assert workbook.get_cell_contents("Sheet2", "A2") == "='Sheet1!'"
    # assert workbook.get_cell_contents("Sheet2", "A3") == "='Sheet1 bruh'"


def test_rename_sheet_updates_values():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.set_cell_contents("Sheet1", "A1", 'bruh')
    workbook.new_sheet("Sheet2")
    workbook.set_cell_contents("Sheet2","A1","=Sheet3!A1")
    assert isinstance(workbook.get_cell_value("Sheet2", "A1"),CellError)
    assert workbook.get_cell_value("Sheet2", "A1").get_type() ==  CellErrorType.BAD_REFERENCE
    workbook.rename_sheet("Sheet1", "Sheet3")
    assert workbook.get_cell_value("Sheet2", "A1") == 'bruh'


# Test for moving sheets
def test_move_sheet_basic():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.new_sheet("Sheet2")
    wb.move_sheet("Sheet1", 1)
    assert wb.list_sheets() == ["Sheet2", "Sheet1"]

def test_move_sheet_invalid_index():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    with pytest.raises(IndexError):
        wb.move_sheet("Sheet1", 2)  # Invalid index

def test_move_sheet_nonexistent_sheet():
    wb = Workbook()
    with pytest.raises(KeyError):
        wb.move_sheet("NonExistentSheet", 0)

# Test for copying sheets
def test_copy_sheet_basic():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    index, name = wb.copy_sheet("Sheet1")
    assert name == "Sheet1_1" and index == 1

def test_copy_sheet_increment_name():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.copy_sheet("Sheet1")  # Creates "Sheet1_1"
    index, name = wb.copy_sheet("Sheet1")
    assert name == "Sheet1_2" and index == 2

def test_copy_sheet_nonexistent_sheet():
    wb = Workbook()
    with pytest.raises(KeyError):
        wb.copy_sheet("NonExistentSheet")

def test_copy_sheet_independence():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "10")
    index, name = wb.copy_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "20")
    original_value = wb.get_cell_value("Sheet1", "A1")
    copied_value = wb.get_cell_value(name, "A1")
    assert original_value != copied_value

def test_copy_sheet():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "10")
    index, name = wb.copy_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "20")
    original_value = wb.get_cell_value("Sheet1", "A1")
    copied_value = wb.get_cell_value(name, "A1")
    assert original_value != copied_value


def test_copy_sheet_preserve_order():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.new_sheet("Sheet2")
    wb.copy_sheet("Sheet1")
    assert wb.list_sheets() == ["Sheet1", "Sheet2", "Sheet1_1"]

def test_copy_sheet_updates_values():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.set_cell_contents("Sheet1", "A1", 'bruh')
    workbook.new_sheet("Sheet2")
    workbook.set_cell_contents("Sheet2","A1","=Sheet1_1!A1")
    assert isinstance(workbook.get_cell_value("Sheet2", "A1"),CellError)
    assert workbook.get_cell_value("Sheet2", "A1").get_type() ==  CellErrorType.BAD_REFERENCE
    workbook.copy_sheet("Sheet1")
    assert workbook.get_cell_value("Sheet2", "A1") == 'bruh'

def test_copy_sheet_circ_ref():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.set_cell_contents("Sheet1", "A1", '=Sheet1_1!A1')
    assert workbook.get_cell_value("Sheet1", "A1").get_type() ==  CellErrorType.BAD_REFERENCE
    workbook.copy_sheet("Sheet1")
    assert workbook.get_cell_value("Sheet1_1", "A1").get_type()  == CellErrorType.CIRCULAR_REFERENCE
    assert workbook.get_cell_value("Sheet1", "A1").get_type()  == CellErrorType.CIRCULAR_REFERENCE

def test_copy_sheet_name_correct():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.copy_sheet("Sheet1")
    workbook.copy_sheet("Sheet1")
    workbook.copy_sheet("Sheet1")
    workbook.copy_sheet("Sheet1")
    assert workbook.list_sheets() == ["Sheet1", "Sheet1_1", "Sheet1_2", "Sheet1_3", "Sheet1_4"]
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.new_sheet("Sheet1_1")
    workbook.copy_sheet("Sheet1")
    assert workbook.list_sheets() == ["Sheet1", "Sheet1_1", "Sheet1_2"]
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.copy_sheet("Sheet1")
    workbook.copy_sheet("Sheet1_1")
    assert workbook.list_sheets() == ["Sheet1", "Sheet1_1", "Sheet1_1_1"]



def test_move_sheet_to_same_position():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.new_sheet("Sheet2")
    wb.move_sheet("Sheet1", 0)
    assert wb.list_sheets() == ["Sheet1", "Sheet2"]


def test_new_sheet_updates_values():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.set_cell_contents("Sheet1", "A1", 'bruh')
    workbook.new_sheet("Sheet2")
    workbook.set_cell_contents("Sheet2","A1","=Sheet3!A1")
    assert isinstance(workbook.get_cell_value("Sheet2", "A1"),CellError)
    assert workbook.get_cell_value("Sheet2", "A1").get_type() ==  CellErrorType.BAD_REFERENCE
    workbook.new_sheet("Sheet3")
    assert workbook.get_cell_value("Sheet2", "A1") == decimal.Decimal(0)

def test_delete_sheet_bad_ref():
    workbook = Workbook()
    workbook.new_sheet("Sheet1")
    workbook.set_cell_contents("Sheet1", "A1", 'bruh')
    workbook.new_sheet("Sheet2")
    workbook.set_cell_contents("Sheet2","A1","=Sheet1!A1")
    assert workbook.get_cell_value("Sheet2", "A1") == 'bruh'
    workbook.del_sheet("Sheet1")
    assert workbook.get_cell_value("Sheet2", "A1").get_type() ==  CellErrorType.BAD_REFERENCE


    