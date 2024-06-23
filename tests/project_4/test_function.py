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


def test_sum(wb):
    set_it(wb,"=SUM(5,6,7,8)")
    assert get_it(wb) == decimal.Decimal(5+6+7+8)

def test_and(wb):
    set_it(wb,"=AND(TRUE,TRUE,FALSE)")
    assert get_it(wb) == False
    set_it(wb,"=AND(FALSE,TRUE,TRUE)")
    assert get_it(wb) == False
    set_it(wb,"=AND(TRUE,TRUE,TRUE)")
    assert get_it(wb) == True

def test_or(wb):
    set_it(wb,"=oR(TRUE,FALSE,FALSE)")
    assert get_it(wb) == True
    set_it(wb,"=OR(TRUE,FALSE,TRUE)")
    assert get_it(wb) == True
    set_it(wb,"=OR(FALSE,FALSE, FALSE)")
    assert get_it(wb) == False

def test_not(wb):
    set_it(wb,"=NOT(TRUE)")
    assert get_it(wb) == False
    set_it(wb,"=NOT(FALSE)")
    assert get_it(wb) == True

def test_xor(wb):
    set_it(wb,"=XOR(TRUE, FALSE, TRUE)")
    assert get_it(wb) == True
    set_it(wb,"=XOR(FALSE, FALSE, FALSE)")
    assert get_it(wb) == False
    set_it(wb,"=XOR(TRUE, TRUE, TRUE)")
    assert get_it(wb) == False

def test_if(wb):
    set_it(wb,"=IF(TRUE, \"bruh\", 99)")
    assert get_it(wb) == "bruh"
    set_it(wb,"=IF(FALSE, \"bruh\", 99)")
    assert get_it(wb) == decimal.Decimal(99)
    set_it(wb,"=IF(FALSE, \"bruh\")")
    assert get_it(wb) == False

def test_iferror(wb):
    set_it(wb,"=IFERROR(#REF!, \"bruh\")")
    assert get_it(wb) == "bruh"
    set_it(wb,"=IFERROR(TRUE, \"bruh\")")
    assert get_it(wb) == True
    set_it(wb,"=IFERROR(1/0)")
    assert get_it(wb) == ""

def test_iferror2(wb):
    set_it(wb,"=IFERROR(#ref!, \"bruh\")")
    assert get_it(wb) == "bruh"
    set_it(wb,"=IFERROR(\"rain drop drop top\", \"bruh\")")
    assert get_it(wb) == "rain drop drop top"
    set_it(wb,"=IFERROR(TRUE, \"bruh\")")
    assert get_it(wb) == True

def test_choice(wb):
    set_it(wb,"=CHOICE(1,1,2,3,4,5,6,7)")
    get_it(wb) == 1
    set_it(wb,"=CHOICE(2,1,2,3,4,5,6,7)")
    get_it(wb) == 2
    set_it(wb,"=CHOICE(3,1,2,3,4,5,6,7)")
    get_it(wb) == 3
    set_it(wb,"=CHOICE(4,1,2,3,4,5,6,7)")
    get_it(wb) == 4
    set_it(wb,"=CHOICE(5,1,2,3,4,5,6,7)")
    get_it(wb) == 5

def test_iserror1(wb):
    set_it(wb,"=ISERROR(#NAME?)")
    assert get_it(wb) == True
    set_it(wb,"=ISERROR(TRUE)")
    assert get_it(wb) == False


def test_isError_example(wb):
    wb.set_cell_contents("Sheet1", "C1", "=ISeRROR(B1)")
    assert wb.get_cell_value('Sheet1', 'C1') == False
    wb.set_cell_contents("Sheet1", "B1", "=A1")
    wb.set_cell_contents("Sheet1", "A1", "=B1")

    assert wb.get_cell_value("Sheet1","C1") == True
    assert wb.get_cell_value("Sheet1","B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1","A1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_isError_example2(wb):
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=A1")
    wb.set_cell_contents("Sheet1", "C1", "=ISERROR(B1)")
    assert wb.get_cell_contents("Sheet1", "B1") 
    assert wb.get_cell_value("Sheet1","C1") == True
    assert wb.get_cell_value("Sheet1","B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1","A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_isError_example3(wb):
    wb.set_cell_contents("Sheet1","B1", "=A1")
    wb.set_cell_contents("Sheet1", "A1", "=ISERROR(B1)")
    assert  wb.get_cell_value("Sheet1","A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    wb.new_sheet("Sheet2")
    wb.set_cell_contents("Sheet2", "A1", "=ISErROR(B1)")
    wb.set_cell_contents("Sheet2", "B1", "=A1")
    assert wb.get_cell_value("Sheet1","A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_isblank(wb):
    wb.set_cell_contents("Sheet1", "A1", "=ISBLANK(B1)")
    assert wb.get_cell_value("Sheet1","A1") == True
    wb.set_cell_contents("Sheet1", "B1", "5")
    assert wb.get_cell_value("Sheet1","A1") == False
    wb.set_cell_contents("Sheet1", "B1", "")
    assert wb.get_cell_value("Sheet1","A1") == True

def test_version(wb):
    pass

def test_indirect(wb):
    wb.set_cell_contents("Sheet1", "B1", "=5")
    wb.set_cell_contents("Sheet1", "C1", "=6")
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"B\" & \"1\")")
    assert get_it(wb) == decimal.Decimal(5)
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"C\" & \"1\")")
    assert get_it(wb) == decimal.Decimal(6)

def test_indirect_example(wb):
    wb.set_cell_contents("sheet1", "A1", "=B1")
    # functions are case insensitive
    wb.set_cell_contents("sheet1", "B1", "=indirect(\"A1\")")
    assert wb.get_cell_value("sheet1", "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    wb.set_cell_contents("sheet1", "A1", "5")
    assert wb.get_cell_value("sheet1", "B1") == decimal.Decimal(5)
    assert wb.get_cell_value("sheet1", "A1") == decimal.Decimal(5)


def test_conditional_functions(wb):
    wb.set_cell_contents("Sheet1", "a1", "=0")
    wb.set_cell_contents("Sheet1", "A2", "=If(a1, 5, 6)")
    assert wb.get_cell_value("Sheet1", "A2") == decimal.Decimal(6)
    wb.set_cell_contents("Sheet1", "A1", "=1")
    assert wb.get_cell_value("Sheet1", "A2") == decimal.Decimal(5)

def test_conditional_functions2(wb):
    wb.set_cell_contents("Sheet1", "A1", "=IF(A2, A1, c1)")
    wb.set_cell_contents("Sheet1", "A2", "False")
    wb.set_cell_contents("Sheet1", "B1", "A1")
    wb.set_cell_contents("Sheet1", "C1", "6")
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(6)

    wb.set_cell_contents("Sheet1", "A2", "True")

    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "A2") == True
    # wb.set_cell_contents("Sheet1", "A1", "=1")
    # assert wb.get_cell_value("Sheet1", "A2") == decimal.Decimal(5)

def test_isBlank_swallows_errors(wb):
    wb.set_cell_contents("sheet1", "A1", "=ISBLANK(B1)")
    assert wb.get_cell_value("sheet1", "A1") == True
    wb.set_cell_contents("sheet1", "B1", "#REF!")
    assert wb.get_cell_value("sheet1", "A1") == False

    # In this case, the cell-reference “Nonexistent!B1” would evaluate to #REF! since the sheet doesn’t exist. But, then #REF! is not blank, so the result is FALSE.

    wb.set_cell_contents("sheet1", "A1", "=ISBLANK(Sheet2!B1)")
    assert wb.get_cell_value("sheet1", "A1") == False

    wb.set_cell_contents("sheet1", "A1", "=ISBLANK(B1+)")
    assert wb.get_cell_value("sheet1", "A1").get_type() == CellErrorType.PARSE_ERROR

    wb.set_cell_contents("sheet1", "A1", "=ISBLANK(B1)")
    wb.set_cell_contents("sheet1", "B1", "=ISBLANK(A1)")
    assert wb.get_cell_value("sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("sheet1", "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_IndirectForSheetNames(wb):
    wb.new_sheet("Sheet!2")
    wb.set_cell_contents("Sheet!2", "A1", "=4")
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"'Sheet!2'!A1\")")
    assert wb.get_cell_value("Sheet1", "A1") == 4