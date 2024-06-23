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

def set_it(wb, contents):
    wb.set_cell_contents("Sheet1", "A1", contents)
def get_it(wb):
    return wb.get_cell_value("Sheet1", "A1")

def transpose(array):
    # Use list comprehension to iterate over the columns and rows
    return [list(row) for row in zip(*array)]

def write_2d_array_to_sheet(workbook, top_left_cell, data_2d, sheet_name="Sheet1"):
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
            if row_index == 0 and col_index == 0:
                start_loc = cell_location
            # Write value to the cell
            workbook.set_cell_contents(sheet_name, cell_location, str(value))
    return f"{sheet_name}!{start_loc}:{cell_location}"

def test_H_lookup(wb):
    test_data = [["Name:", "Jimmy", "Bob", "Alice", "Frank"],
                 ["Job:", "Carpenter", "Cleaner", "Carpenter", "Doctor"],
                 ["Net Worth:", "50", "100", "125", "65"],
                 ["Married:", "True", "False", "False", "True"]]
    cell_range = write_2d_array_to_sheet(wb, "B1", test_data)
    wb.set_cell_contents("Sheet1", "A1", f"=HLOOKUP(\"Bob\", {cell_range}, 3)")
    assert wb.get_cell_value("Sheet1", "A1") == 100
    wb.set_cell_contents("Sheet1", "A1", f"=HLOOKUP(\"Frank\", {cell_range}, 1)")
    assert wb.get_cell_value("Sheet1", "A1") == "Frank"
    wb.set_cell_contents("Sheet1", "A1", f"=HLOOKUP(\"Alice\", {cell_range}, 4)")
    assert wb.get_cell_value("Sheet1", "A1") == False
    wb.set_cell_contents("Sheet1", "A1", f"=HLOOKUP(\"jimmy\", {cell_range}, 4)")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR
    wb.set_cell_contents("Sheet1", "A1", f"=HLOOKUP(\"Jimmy\", {cell_range}, 5)")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR


def test_V_lookup(wb):
    test_data = [["Name:", "Jimmy", "Bob", "Alice", "Frank"],
                 ["Job:", "Carpenter", "Cleaner", "Carpenter", "Doctor"],
                 ["Net Worth:", "50", "100", "125", "65"],
                 ["Married:", "True", "False", "False", "True"]]
    test_data = transpose(test_data)
    cell_range = write_2d_array_to_sheet(wb, "B1", test_data)
    wb.set_cell_contents("Sheet1", "A1", f"=VLOOKUP(\"Bob\", {cell_range}, 3)")
    assert wb.get_cell_value("Sheet1", "A1") == 100
    wb.set_cell_contents("Sheet1", "A1", f"=VLOOKUP(\"Frank\", {cell_range}, 1)")
    assert wb.get_cell_value("Sheet1", "A1") == "Frank"
    wb.set_cell_contents("Sheet1", "A1", f"=VLOOKUP(\"Alice\", {cell_range}, 4)")
    assert wb.get_cell_value("Sheet1", "A1") == False
    wb.set_cell_contents("Sheet1", "A1", f"=VLOOKUP(\"jimmy\", {cell_range}, 4)")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR
    wb.set_cell_contents("Sheet1", "A1", f"=VLOOKUP(\"Jimmy\", {cell_range}, 5)")
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

def test_H_lookup_with_conditionals(wb):
    test_data = [["Name:", "Jimmy", "Bob", "Alice", "Frank"],
                 ["Job:", "Carpenter", "Cleaner", "Carpenter", "Doctor"],
                 ["Net Worth:", "50", "100", "125", "65"],
                 ["Married:", "True", "False", "False", "True"]]
    cell_range = write_2d_array_to_sheet(wb, "B1", test_data)
    wb.set_cell_contents("Sheet1", "A1", f"=3+3 > 4")

    wb.set_cell_contents("Sheet1", "A2", f"=HLOOKUP(\"Bob\", if(A1, {cell_range}, B1:E4), 3)")
    assert wb.get_cell_value("Sheet1", "A2") == 100
    wb.set_cell_contents("Sheet1", "A2", f"=HLOOKUP(\"Frank\", if(A1, {cell_range}, B1:E4), 3)")
    assert wb.get_cell_value("Sheet1", "A2") == 65
    wb.set_cell_contents("Sheet1", "A1", "= 2+2 > 4")
    wb.set_cell_contents("Sheet1", "A2", f"=HLOOKUP(\"Frank\", if(A1, {cell_range}, B1:E4), 3)")
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError)
    assert wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.TYPE_ERROR

    wb.set_cell_contents("Sheet1", "A2", f"=HLOOKUP(\"Frank\", if(A1, {cell_range}, B1:D4), 4)")
    wb.set_cell_contents("Sheet1", "A1", "= 2+2 > 4")

    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError)
    assert wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.TYPE_ERROR

    wb.set_cell_contents("Sheet1", "A2", f"=HLOOKUP(\"Frank\", IFERROR(A1, {cell_range}), 4)")
    wb.set_cell_contents("Sheet1", "A1", "= 1/0")
    assert wb.get_cell_value("Sheet1", "A2") == True

    #TODO write choose test


def test_H_lookup_with_cycles(wb):
    test_data = [["Name:", "Jimmy", "Bob", "=A1", "Frank"],
                 ["Job:", "Carpenter", "Cleaner", "Carpenter", "Doctor"],
                 ["Net Worth:", "50", "100", "125", "65"],
                 ["Married:", "True", "False", "False", "True"]]
    cell_range = write_2d_array_to_sheet(wb, "B1", test_data)
    set_it(wb,f"=HLOOKUP(\"Jimmy\", {cell_range}, 2)")
    assert isinstance(get_it(wb), CellError)
    assert get_it(wb).get_type() == CellErrorType.CIRCULAR_REFERENCE

    test_data = [["Name:", "Jimmy", "Bob", "Alice", "Frank"],
                 ["Job:", "Carpenter", "Cleaner", "Carpenter", "Doctor"],
                 ["Net Worth:", "50", "100", "125", "=A1"],
                 ["Married:", "True", "False", "False", "True"]]
    cell_range = write_2d_array_to_sheet(wb, "B1", test_data)
    set_it(wb,f"=HLOOKUP(\"Jimmy\", {cell_range}, 2)")
    assert get_it(wb) == "Carpenter"
    set_it(wb,f"=HLOOKUP(\"Frank\", {cell_range}, 3)")
    assert isinstance(get_it(wb),CellError)
    assert get_it(wb).get_type() == CellErrorType.CIRCULAR_REFERENCE


