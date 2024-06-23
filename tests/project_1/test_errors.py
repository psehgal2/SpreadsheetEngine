# Now we can import other things
import sheets
import pytest
import context

from sheets.utils import get_hidden_name
from decimal import Decimal

from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType
from sheets.Workbook import Workbook


def test_parse_error():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=INVALID+5")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.PARSE_ERROR

def test_circular_reference():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_bad_reference():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=INVALID!B1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE

def test_bad_reference_2():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=INVALID!B1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE


# def test_bad_name():
#     wb = Workbook()
#     wb.new_sheet("Sheet1")
#     wb.set_cell_contents("Sheet1", "A1", "=UNKNOWN_FUNCTION(5)")
#
#     assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
#     assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_NAME

def test_type_error():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "B1", "Hello")

    wb.set_cell_contents("Sheet1", "A1", "=B1 + 5")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

def error_artihmetic():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1" , "A1", "= #REF! + 1")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE
    wb.set_cell_contents("Sheet1" , "A1", "= #REF! * 1")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE
    wb.set_cell_contents("Sheet1" , "A1", "= #REF! / 1")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE
    wb.set_cell_contents("Sheet1" , "A1", "= #REF! - 1")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE


def test_divide_by_zero():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=5/0")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO

def test_error_propagation():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=5/0")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO

def test_multiple_errors_priority():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=B1/C1")
    wb.set_cell_contents("Sheet1", "B1", "=INVALID")
    wb.set_cell_contents("Sheet1", "C1", "=0")

    # Should prioritize PARSE_ERROR over DIVIDE_BY_ZERO
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.PARSE_ERROR

def test_error_literal_in_formula():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=#REF! + 5")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE

def test_string_mult():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1" , "=3*\"abc\"")
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

def test_string_representation_of_errors():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "=5/0")

    cell_value = wb.get_cell_value("Sheet1", "A1")
    assert isinstance(cell_value, CellError)
    # assert str(cell_value) == "ERROR[DIVIDE_BY_ZERO, \"#DIV/0!\"]"

def test_empty_and_whitespace_cell_contents():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "   ")
    wb.set_cell_contents("Sheet1", "B1", "")

    assert wb.get_cell_value("Sheet1", "A1") is None
    assert wb.get_cell_value("Sheet1", "B1") is None

def test_invalid_cell_location():
    wb = Workbook()

    wb.new_sheet("Sheet1")

    with pytest.raises(ValueError):
        wb.set_cell_contents("Sheet1", "INVALID", "123")

    with pytest.raises(ValueError):
        wb.get_cell_value("Sheet1", "INVALID")

# def test_large_spreadsheet_behavior():
#     wb = Workbook()
#     wb.new_sheet("Sheet1")
#
#     for row in range(1, 1001):
#         wb.set_cell_contents("Sheet1", f"A{row}", f"=B{row}+1")
#         wb.set_cell_contents("Sheet1", f"B{row}", "1")
#
#     for row in range(1, 1001):
#         assert wb.get_cell_value("Sheet1", f"A{row}") == 2

def test_case_insensitivity():
    wb = Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "a1", "123")

    assert wb.get_cell_contents("sheet1", "A1") == "123"
    assert wb.get_cell_value("SHEET1", "a1") == 123
