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


# Test Results: leopard Score: 485/487 (100%) Test Failures: [-2] Test that after renaming Some -> Other, Other has all of Some's errors

def test_sheet_rename_carry_error(wb):
    wb.new_sheet("some")
    wb.set_cell_contents("some", "a1", "=b1")
    wb.set_cell_contents("some", "b1", "=a1")
    wb.set_cell_contents("some", "b2", "=sheet5!b1")
    wb.set_cell_contents("some", "b3", "=BADFUNCTION(b1)")
    wb.set_cell_contents("some", "b4", "=1 + \"five\"")
    wb.set_cell_contents("some", "b5", "=1/0")
    wb.set_cell_contents("some", "b6", "=1 +hiya1 * garbage / \"garbage\"")
    assert wb.get_cell_value("some", "A1").get_type() ==  CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("some", "b2").get_type() ==  CellErrorType.BAD_REFERENCE
    assert wb.get_cell_value("some", "b3").get_type() ==  CellErrorType.BAD_NAME
    assert wb.get_cell_value("some", "b4").get_type() ==  CellErrorType.TYPE_ERROR
    assert wb.get_cell_value("some", "b5").get_type() ==  CellErrorType.DIVIDE_BY_ZERO
    assert wb.get_cell_value("some", "b6").get_type() ==  CellErrorType.PARSE_ERROR

    wb.rename_sheet("some", "other")
    assert wb.get_cell_value("other", "A1").get_type() ==  CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("other", "b2").get_type() ==  CellErrorType.BAD_REFERENCE
    assert wb.get_cell_value("other", "b3").get_type() ==  CellErrorType.BAD_NAME
    assert wb.get_cell_value("other", "b4").get_type() ==  CellErrorType.TYPE_ERROR
    assert wb.get_cell_value("other", "b5").get_type() ==  CellErrorType.DIVIDE_BY_ZERO
    assert wb.get_cell_value("other", "b6").get_type() ==  CellErrorType.PARSE_ERROR

    


    




