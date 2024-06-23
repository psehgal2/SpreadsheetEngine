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

def write_2d_array_to_sheet(workbook, sheet_name, top_left_cell, data_2d):
    """
    Writes a 2D array to a specified sheet in the workbook, starting from a given top left cell.

    :param workbook: Instance of the Workbook class where the data will be written.
    :param sheet_name: String, the name of the sheet where data will be written.
    :param top_left_cell: String, the location of the top left cell where writing starts (e.g., "A1").
    :param data_2d: List of lists, the 2D array data to write to the sheet.
    """
    # Calculate starting row and column from top_left_cell
    start_col = ord(top_left_cell[0].upper()) - ord('A')
    start_row = int(top_left_cell[1:]) - 1  # Assuming the cell is in format like "A1", "B2", etc.

    for row_index, row in enumerate(data_2d):
        for col_index, value in enumerate(row):
            # Calculate target cell address
            col_letter = chr(ord('A') + start_col + col_index)
            cell_location = f"{col_letter}{start_row + row_index + 1}"
            # Write value to the cell
            workbook.set_cell_contents(sheet_name, cell_location, str(value))



def test_sum(wb):
    wb.set_cell_contents("Sheet1", "B1" , "1")
    wb.set_cell_contents("Sheet1", "B2" , "2")
    wb.set_cell_contents("Sheet1", "B3" , "3")
    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(B1:B3)")
    assert wb.get_cell_value("Sheet1", "A1") == 6


def test_sum_disjoint_contents(wb):
    wb.set_cell_contents("Sheet1", "B1" , "1")
    wb.set_cell_contents("Sheet1", "B2" , "2")
    wb.set_cell_contents("Sheet1", "B3" , "3")
    wb.set_cell_contents("Sheet1", "B5" , "1")
    wb.set_cell_contents("Sheet1", "B6" , "2")
    wb.set_cell_contents("Sheet1", "B7" , "3")
    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(B1:B3, B5:B7)")
    assert wb.get_cell_value("Sheet1", "A1") == 12


def test_sum_array(wb):
        wb.set_cell_contents("Sheet1", "B1", "1")
        wb.set_cell_contents("Sheet1", "B2", "2")
        wb.set_cell_contents("Sheet1", "B3", "3")
        wb.set_cell_contents("Sheet1", "C1", "1")
        wb.set_cell_contents("Sheet1", "C2", "2")
        wb.set_cell_contents("Sheet1", "C3", "3")
        wb.set_cell_contents("Sheet1" ,"A1", "=SUM(B1:C3)")
        assert wb.get_cell_value("Sheet1", "A1") == 12


def test_sum_with_nones(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    # wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "C1", "1")
    wb.set_cell_contents("Sheet1", "C2", "2")
    wb.set_cell_contents("Sheet1", "C3", "3")
    wb.set_cell_contents("Sheet1", "A1", "=SUM(B1:C3)")
    assert wb.get_cell_value("Sheet1", "A1") == 9


def test_sum_with_no_args(wb):
    wb.set_cell_contents("Sheet1", "A1", "=SUM()")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

def test_circular_sum(wb):
    wb.set_cell_contents("Sheet1", "B5", "=SUM(A1: C10)")
    assert isinstance(wb.get_cell_value("Sheet1", "B5"), CellError)
    assert wb.get_cell_value("Sheet1", "B5").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_single_error_sum(wb):
    wb.set_cell_contents("Sheet1", "A1", "1")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "A3", "3")
    wb.set_cell_contents("Sheet1", "A4", "#DIV/0!")
    wb.set_cell_contents("Sheet1", "A5", "5")
    wb.set_cell_contents("Sheet1", "A6", "6")

    wb.set_cell_contents("Sheet1", "B1", "=SUM(A1:A6)")
    assert isinstance(wb.get_cell_value("Sheet1", "B1"), CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.DIVIDE_BY_ZERO

def test_sum_on_a_different_sheet(wb):
    wb.set_cell_contents("Sheet1", "A1", "1")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "A3", "3")
    wb.set_cell_contents("Sheet1", "A4", "4")
    wb.set_cell_contents("Sheet1", "A5", "5")
    wb.set_cell_contents("Sheet1", "A6", "6")
    wb.new_sheet("Sheet2")
    wb.set_cell_contents("Sheet2", "B1", "=SUM(Sheet1!A1:A6)")
    assert wb.get_cell_value("Sheet2", "B1") == 21

def test_max(wb):
    wb.set_cell_contents("Sheet1", "B1" , "1")
    wb.set_cell_contents("Sheet1", "B2" , "-2")
    wb.set_cell_contents("Sheet1", "B3" , "6")
    wb.set_cell_contents("Sheet1" ,"A1", "=MAX(B1:B3)")
    assert wb.get_cell_value("Sheet1", "A1") == 6



def test_max_disjoint_contents(wb):
    wb.set_cell_contents("Sheet1", "B1" , "1")
    wb.set_cell_contents("Sheet1", "B2" , "2")
    wb.set_cell_contents("Sheet1", "B3" , "3")
    wb.set_cell_contents("Sheet1", "B5" , "11")
    wb.set_cell_contents("Sheet1", "B6" , "2")
    wb.set_cell_contents("Sheet1", "B7" , "3")
    wb.set_cell_contents("Sheet1" ,"A1", "=MAX(B1:B3, B5:B7)")
    assert wb.get_cell_value("Sheet1", "A1") == 11

def test_max_array(wb):
        wb.set_cell_contents("Sheet1", "B1", "1")
        wb.set_cell_contents("Sheet1", "B2", "100")
        wb.set_cell_contents("Sheet1", "B3", "3")
        wb.set_cell_contents("Sheet1", "C1", "1")
        wb.set_cell_contents("Sheet1", "C2", "-2")
        wb.set_cell_contents("Sheet1", "C3", "3")
        wb.set_cell_contents("Sheet1" ,"A1", "=MAX(B1:C3)")
        assert wb.get_cell_value("Sheet1", "A1") == 100

def test_max_2(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "A1", "=MAX(B1:B3)")
    assert wb.get_cell_value("Sheet1", "A1") == 3

def test_max_disjoint_contents_2(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "B5", "5")
    wb.set_cell_contents("Sheet1", "B6", "6")
    wb.set_cell_contents("Sheet1", "B7", "7")
    wb.set_cell_contents("Sheet1", "A1", "=MAX(B1:B3, B5:B7)")
    assert wb.get_cell_value("Sheet1", "A1") == 7

def test_max_array(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "C1", "4")
    wb.set_cell_contents("Sheet1", "C2", "5")
    wb.set_cell_contents("Sheet1", "C3", "6")
    wb.set_cell_contents("Sheet1", "A1", "=MAX(B1:C3)")
    assert wb.get_cell_value("Sheet1", "A1") == 6

def test_max_with_nones(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    # wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "C1", "4")
    wb.set_cell_contents("Sheet1", "C2", "5")
    wb.set_cell_contents("Sheet1", "C3", "6")
    wb.set_cell_contents("Sheet1", "A1", "=MAX(B1:C3)")
    assert wb.get_cell_value("Sheet1", "A1") == 6

def test_max_with_no_args(wb):
    wb.set_cell_contents("Sheet1", "A1", "=MAX()")
    assert isinstance(wb.get_cell_value("Sheet1", "A1") , CellError)
    assert wb.get_cell_value("Sheet1" , "A1").get_type() == CellErrorType.TYPE_ERROR
# Assuming MAX with no args returns 0, or could be an error depending on implementation

def test_circular_max(wb):
    wb.set_cell_contents("Sheet1", "B5", "=MAX(A1: C10)")
    assert isinstance(wb.get_cell_value("Sheet1", "B5"), CellError)
    assert wb.get_cell_value("Sheet1", "B5").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_single_error_max(wb):
    wb.set_cell_contents("Sheet1", "A1", "1")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "A3", "3")
    wb.set_cell_contents("Sheet1", "A4", "#DIV/0!")
    wb.set_cell_contents("Sheet1", "A5", "5")
    wb.set_cell_contents("Sheet1", "A6", "6")

    wb.set_cell_contents("Sheet1", "B1", "=MAX(A1:A6)")
    assert isinstance(wb.get_cell_value("Sheet1", "B1"), CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.DIVIDE_BY_ZERO

def test_min(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "A1", "=MIN(B1:B3)")
    assert wb.get_cell_value("Sheet1", "A1") == 1

def test_min_disjoint_contents(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "B5", "5")
    wb.set_cell_contents("Sheet1", "B6", "6")
    wb.set_cell_contents("Sheet1", "B7", "7")
    wb.set_cell_contents("Sheet1", "A1", "=MIN(B1:B3, B5:B7)")
    assert wb.get_cell_value("Sheet1", "A1") == 1

def test_min_array(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "B3", "3")
    wb.set_cell_contents("Sheet1", "C1", "4")
    wb.set_cell_contents("Sheet1", "C2", "5")
    wb.set_cell_contents("Sheet1", "C3", "6")
    wb.set_cell_contents("Sheet1", "A1", "=MIN(B1:C3)")
    assert wb.get_cell_value("Sheet1", "A1") == 1

def test_min_with_nones(wb):
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "2")
    # Omitting B3 to mimic none value
    wb.set_cell_contents("Sheet1", "C1", "4")
    wb.set_cell_contents("Sheet1", "C2", "5")
    wb.set_cell_contents("Sheet1", "C3", "6")
    wb.set_cell_contents("Sheet1", "A1", "=MIN(B1:C3)")
    assert wb.get_cell_value("Sheet1", "A1") == 1

def test_min_with_no_args(wb):
    wb.set_cell_contents("Sheet1", "A1", "=MIN()")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"),CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

def test_circular_min(wb):
    wb.set_cell_contents("Sheet1", "B5", "=MIN(A1: C10)")
    assert isinstance(wb.get_cell_value("Sheet1", "B5"), CellError)
    assert wb.get_cell_value("Sheet1", "B5").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_single_error_min(wb):
    wb.set_cell_contents("Sheet1", "A1", "1")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "A3", "3")
    wb.set_cell_contents("Sheet1", "A4", "#DIV/0!")
    wb.set_cell_contents("Sheet1", "A5", "5")
    wb.set_cell_contents("Sheet1", "A6", "6")

    wb.set_cell_contents("Sheet1", "B1", "=MIN(A1:A6)")
    assert isinstance(wb.get_cell_value("Sheet1", "B1"), CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.DIVIDE_BY_ZERO


def test_average(wb):
    wb.set_cell_contents("Sheet1", "B1" , "2")
    wb.set_cell_contents("Sheet1", "B2" , "4")
    wb.set_cell_contents("Sheet1", "B3" , "6")
    wb.set_cell_contents("Sheet1" ,"A1", "=AVERAGE(B1:B3)")
    assert wb.get_cell_value("Sheet1", "A1") == 4

def test_average_disjoint(wb):
    wb.set_cell_contents("Sheet1", "B1" , "2")
    wb.set_cell_contents("Sheet1", "B2" , "4")
    wb.set_cell_contents("Sheet1", "B3" , "6")
    wb.set_cell_contents("Sheet1", "B5" , "2")
    wb.set_cell_contents("Sheet1", "B6" , "4")
    wb.set_cell_contents("Sheet1", "B7" , "6")
    wb.set_cell_contents("Sheet1" ,"A1", "=AVERAGE(B1:B3, B5:B7)")
    assert wb.get_cell_value("Sheet1", "A1") == 4

def test_indirect_range(wb):
    wb.set_cell_contents("Sheet1", "B1" , "2")
    wb.set_cell_contents("Sheet1", "B2" , "4")
    wb.set_cell_contents("Sheet1", "B3" , "6")

    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(INDIRECT(\"B1:B3\"))")
    assert wb.get_cell_value("Sheet1", "A1") == 12
    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(INDIRECT(\"$B$1:B3\"))")
    assert wb.get_cell_value("Sheet1", "A1") == 12
    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(INDIRECT(\"B$1:$B3\"))")
    assert wb.get_cell_value("Sheet1", "A1") == 12
    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(INDIRECT(\"$B$1:$B$3\"))")
    assert wb.get_cell_value("Sheet1", "A1") == 12
    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(INDIRECT(\"Sheet1!B1:B3\"))")
    assert wb.get_cell_value("Sheet1", "A1") == 12
    wb.set_cell_contents("Sheet1" ,"A1", "=SUM(INDIRECT(\"Sheet1!$B$1:B$3\"))")
    assert wb.get_cell_value("Sheet1", "A1") == 12