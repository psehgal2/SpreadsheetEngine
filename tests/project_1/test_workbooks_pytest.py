import context

# Now we can import other things
import sheets
from tests.project_1.wrongNamesList import  wrongSheetNames
import pytest

from sheets.utils import get_hidden_name
from decimal import Decimal

from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType


def test_sheet_list_del():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    sheet_name = "Sheet234"
    idx, name = wb.new_sheet(sheet_name)
    sheet_name = None
    idx, name = wb.new_sheet(sheet_name)
    sheet_name = "asdflkjshadkjfh"
    idx, name = wb.new_sheet(sheet_name)
    sheet_name = "w984y578923y4"
    idx, name = wb.new_sheet(sheet_name)
    sheet_name = "jkdjddj"
    idx, name = wb.new_sheet(sheet_name)
    print(wb.list_sheets())
    assert wb.list_sheets() == ["Sheet1", "Sheet234", "Sheet2", "asdflkjshadkjfh", "w984y578923y4", "jkdjddj"]
    # assert wb.list_sheets() == ["Sheet1", "Sheet234", "bruhasdf", "asdflkjshadkjfh", "w984y578923y4", "jkdjddj"]

def setUp() -> None:
    wb = sheets.Workbook()

def test_new_sheet_with_proper_name():
    sheet_name = "Sheet1"
    wb = sheets.Workbook()
    idx, name = wb.new_sheet(sheet_name)
    assert sheet_name == name
    assert idx == 0
    assert sheet_name.lower() in wb.sheets

def test_new_sheet_already_exists():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    assert sheet_name == name
    assert idx == 0
    assert sheet_name.lower() in wb.sheets
    with pytest.raises(ValueError):
        idx, name = wb.new_sheet(sheet_name)
        wb.new_sheet("shEEt1")
        wb.new_sheet("   shEEt1  ")
        wb.new_sheet("   shEEt1")
        wb.new_sheet(" shEEt1  ")

def test_list_sheets():
    wb = sheets.Workbook()
    idx, name = wb.new_sheet("Sheet1")
    idx, name = wb.new_sheet("Sheet2")
    idx, name = wb.new_sheet("Sheet3")

    assert "Sheet3" == name
    assert 2 == idx
    assert name in wb.list_sheets()
    assert get_hidden_name(name) in wb.sheets.keys() 

def test_num_sheets():
    wb = sheets.Workbook()
    idx, name = wb.new_sheet("Sheet1")
    idx, name = wb.new_sheet("Sheet2")
    idx, name = wb.new_sheet("Sheet3")

    with pytest.raises(ValueError):
        temp_idx, n = wb.new_sheet(name)
    assert 2 == idx
    assert wb.num_sheets() == 3

def test_del_sheet():
    wb = sheets.Workbook()
    sheet_name = "Sheet4"
    idx, name = wb.new_sheet(sheet_name)
    assert(sheet_name == name)

    # Here index is zero based
    assert (0 == idx)
    assert sheet_name in wb.list_sheets()
    wb.del_sheet(sheet_name)
    assert sheet_name not in wb.list_sheets()
    assert len(wb.list_sheets()) == 0
    assert wb.num_sheets() == 0

    # Test deleting a sheet that doesn't exist
    with pytest.raises(KeyError):
        wb.del_sheet(sheet_name)

def test_new_sheet_with_improper_name():
    wb = sheets.Workbook()
    for sheet_name in wrongSheetNames():
        with pytest.raises(ValueError):
            workbookLength = wb.num_sheets()
            idx, name = wb.new_sheet(sheet_name)
            assert (sheet_name not in  wb.sheets)
            assert (workbookLength == wb.num_sheets())
    with pytest.raises(ValueError):
        wb.new_sheet("   sheas")
        wb.new_sheet("")


def test_complex_cycle_detection():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    wb.new_sheet(sheet_name)

    # Set up a complex cycle A1 -> B1 -> B2 -> A2 -> A1 and A3 -> B3 -> A1
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    wb.set_cell_contents(sheet_name, "B1", "=B2")
    wb.set_cell_contents(sheet_name, "B2", "=A2")
    wb.set_cell_contents(sheet_name, "A2", "=A1")
    wb.set_cell_contents(sheet_name, "A3", "=B3")
    wb.set_cell_contents(sheet_name, "B3", "=A1")

    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "B2").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "A3").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "B3").get_type() == CellErrorType.CIRCULAR_REFERENCE

    # Fixing A1 should fix all of them
    wb.set_cell_contents(sheet_name, "A1", "=1")
    assert wb.get_cell_value(sheet_name, "A1") == 1
    assert wb.get_cell_value(sheet_name, "B1") == 1
    assert wb.get_cell_value(sheet_name, "B2") == 1
    assert wb.get_cell_value(sheet_name, "A2") == 1
    assert wb.get_cell_value(sheet_name, "A3") == 1
    assert wb.get_cell_value(sheet_name, "B3") == 1


def test_self_referential_cycle():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=A1")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_whitespace():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "'   loll")
    assert wb.get_cell_value(sheet_name, "A1") == "   loll"
    wb.set_cell_contents(sheet_name, "A1", "   loll")
    assert wb.get_cell_contents(sheet_name, "A1") == "loll"

    wb.set_cell_contents(sheet_name, "A1", "      ")
    assert wb.get_sheet_extent(sheet_name) == (0,0)
    assert wb.get_cell_value(sheet_name, "A1") == None


def test_reset_cycles():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    wb.new_sheet(sheet_name)

    wb.set_cell_contents(sheet_name, "A1", "=B1")
    wb.set_cell_contents(sheet_name, "B1", "=A1")
    wb.set_cell_contents(sheet_name, "B1", "")
    assert(1==1)

# def test_complex_formula_evaluator():
    # wb = sheets.Workbook()
    # sheet_name = "Sheet1"
    # wb.new_sheet(sheet_name)

    # # Set up some cells with complex formulas
    # wb.set_cell_contents(sheet_name, "A1", "10")
    # wb.set_cell_contents(sheet_name, "A2", "20")
    # wb.set_cell_contents(sheet_name, "A3", "=A1 + A2 * 2")
    # wb.set_cell_contents(sheet_name, "A4", "=A3 * 2 - A1")
    # wb.set_cell_contents(sheet_name, "A5", "=A4 / (2 + A2)")
    # wb.set_cell_contents(sheet_name, "A6", "=A5 * (A1 - A2) + A3")
    # wb.set_cell_contents(sheet_name, "A7", "=A6 / A5 - A4 + A3 * A2 - A1")

    # # Check that the formulas are evaluated correctly
    # assert wb.get_cell_value(sheet_name, "A3") == 50
    # assert wb.get_cell_value(sheet_name, "A4") == 90
    # print(wb.get_cell_value(sheet_name, "A5"))
    # print(wb.get_cell_contents(sheet_name, "A5"))
    # assert wb.get_cell_value(sheet_name, "A5") == 4.090909090909091
    # assert wb.get_cell_value(sheet_name, "A6") == -20
    # assert wb.get_cell_value(sheet_name, "A7") == -30

def test_formula_evaluator():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    wb.new_sheet(sheet_name)

    # Set up some cells with formulas
    wb.set_cell_contents(sheet_name, "A1", "10")
    wb.set_cell_contents(sheet_name, "A2", "20")
    wb.set_cell_contents(sheet_name, "A3", "=A1 + A2")
    wb.set_cell_contents(sheet_name, "A4", "=A3 * 2")
    wb.set_cell_contents(sheet_name, "A5", "=A4 / 2")

    # Check that the formulas are evaluated correctly
    assert wb.get_cell_value(sheet_name, "A3") == 30
    assert wb.get_cell_value(sheet_name, "A4") == 60
    assert wb.get_cell_value(sheet_name, "A5") == 30

def remove_sheet_with_parents():
    wb = sheets.Workbook()
    sheet1_name = "Sheet1"
    sheet2_name = "Sheet2"
    wb.new_sheet(sheet1_name)
    wb.new_sheet(sheet2_name)
    wb.set_cell_contents(sheet1_name, "A1", "=Sheet2!A1+1")
    wb.set_cell_contents(sheet2_name, "A1", "=2")

    assert wb.get_cell_value(sheet1_name, "A1") == 3
    wb.del_sheet(sheet2_name)
    assert wb.get_cell_value(sheet1_name, "A1").get_type() == CellErrorType.BAD_REFERENCE

    # adding back sheet 2 should fix it
    wb.new_sheet(sheet2_name)
    wb.set_cell_contents(sheet2_name, "A1", "=2")
    assert wb.get_cell_value(sheet1_name, "A1") == 3
    

def test_complex_cycle_across_sheets():
    wb = sheets.Workbook()
    sheet1_name = "Sheet1"
    sheet2_name = "Sheet2"
    wb.new_sheet(sheet1_name)
    wb.new_sheet(sheet2_name)

    # Set up a complex cycle A1 -> B1 -> Sheet2!A1 -> A2 -> A1 and A3 -> B3 -> A1
    wb.set_cell_contents(sheet1_name, "A1", "=B1")
    wb.set_cell_contents(sheet1_name, "B1", "=" + sheet2_name + "!A1")
    wb.set_cell_contents(sheet2_name, "A1", "=" + sheet1_name + "!A2")
    wb.set_cell_contents(sheet1_name, "A2", "=A1")
    wb.set_cell_contents(sheet1_name, "A3", "=B3")
    wb.set_cell_contents(sheet1_name, "B3", "=A1")

    assert wb.get_cell_value(sheet1_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet1_name, "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet2_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet1_name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet1_name, "A3").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet1_name, "B3").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_trailing_zeros():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "1.0")
    wb.set_cell_contents(sheet_name, "A2", "1.000")
    wb.set_cell_contents(sheet_name, "A3", "1.0000")
    wb.set_cell_contents(sheet_name, "A4", "10e+4")

    assert wb.get_cell_contents(sheet_name, "A1") == "1.0"
    assert wb.get_cell_contents(sheet_name, "A2") == "1.000" 
    assert wb.get_cell_contents(sheet_name, "A3") == "1.0000"
    assert wb.get_cell_contents(sheet_name, "A4") == "10e+4"

    assert wb.get_cell_value(sheet_name, "A1") == 1.0
    assert wb.get_cell_value(sheet_name, "A2") == 1.0
    assert wb.get_cell_value(sheet_name, "A3") == 1.0
    assert wb.get_cell_value(sheet_name, "A4") == 100000.0

def test_concatenation():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "'1.0")
    wb.set_cell_contents(sheet_name, "A2", "'1.000")
    wb.set_cell_contents(sheet_name, "A3", "=A1&A2")
    assert wb.get_cell_value(sheet_name, "A3") == "1.01.000"
    wb.set_cell_contents(sheet_name, "A4", "=A1&A2&\"asdf\"")
    assert wb.get_cell_value(sheet_name, "A4") == "1.01.000asdf"

    wb.set_cell_contents(sheet_name, "A4", "=A1&A2&1")
    assert wb.get_cell_value(sheet_name, "A4") == "1.01.0001"

    # test string concatenation

    wb.set_cell_contents(sheet_name, "A1", "'bruh")
    wb.set_cell_contents(sheet_name, "B1", "'lol")
    wb.set_cell_contents(sheet_name, "C1", "=A1&B1")
    assert wb.get_cell_value(sheet_name, "C1") == "bruhlol"

    wb.set_cell_contents(sheet_name, "A1", "'bruh")
    wb.set_cell_contents(sheet_name, "B1", "'   lol")
    wb.set_cell_contents(sheet_name, "C1", "=A1&B1")
    assert wb.get_cell_value(sheet_name, "C1") == "bruh   lol"

    # None cells should be treated as empty strings
    wb.set_cell_contents(sheet_name, "A1", None)
    wb.set_cell_contents(sheet_name, "B1", "'   lol")
    wb.set_cell_contents(sheet_name, "C1", "=A1&B1")
    assert wb.get_cell_value(sheet_name, "C1") == "   lol"

    wb.set_cell_contents(sheet_name, "A1", None)
    wb.set_cell_contents(sheet_name, "B1", None)
    wb.set_cell_contents(sheet_name, "C1", "=A1&B1")
    assert wb.get_cell_value(sheet_name, "C1") == ""

def test_formula_not_in_sheet():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=Sheet2!B1")
    wb.set_cell_contents(sheet_name, "B1", "=Sheet2!C1")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.BAD_REFERENCE
    assert wb.get_cell_value(sheet_name, "B1").get_type() == CellErrorType.BAD_REFERENCE

    # create that sheet, should update sheet1:A1
    wb.new_sheet("Sheet2")
    assert wb.get_cell_value(sheet_name, "A1") == 0
    wb.set_cell_contents("Sheet2", "B1", "1")
    assert wb.get_cell_value(sheet_name, "A1") == 1
    wb.set_cell_contents("Sheet2", "C1", "2")
    assert wb.get_cell_value(sheet_name, "B1") == 2



def test_cycle_multiple():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    wb.set_cell_contents(sheet_name, "B1", "=A1")
    wb.set_cell_contents(sheet_name, "C1", "=B1")

    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE


    wb.set_cell_contents(sheet_name, "B1", "=1")
    # should fix all the cells
    assert wb.get_cell_value(sheet_name, "A1") == 1
    assert wb.get_cell_value(sheet_name, "B1") == 1
    assert wb.get_cell_value(sheet_name, "C1") == 1

    # Create a 3-cycle with C1 A1 and B1
    wb.set_cell_contents(sheet_name, "C1", "=A1")
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    wb.set_cell_contents(sheet_name, "B1", "=C1")

    
    # create a cycle with c2
    wb.set_cell_contents(sheet_name, "C2", "=C1")
    assert wb.get_cell_value(sheet_name, "C2").get_type() == CellErrorType.CIRCULAR_REFERENCE

    # updating C1 should fix all of them
    wb.set_cell_contents(sheet_name, "C1", "=1")
    assert wb.get_cell_value(sheet_name, "C1") == 1
    assert wb.get_cell_value(sheet_name, "C2") == 1
    assert wb.get_cell_value(sheet_name, "A1") == 1
    assert wb.get_cell_value(sheet_name, "B1") == 1

def test_nones():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    assert wb.get_cell_value(sheet_name, "A1")  == 0
    wb.set_cell_contents(sheet_name, "C1", "asdf")
    assert wb.get_cell_value(sheet_name, "C1")  == "asdf"
    wb.set_cell_contents(sheet_name, "A3", "C1 & B1")


def test_cycle_same_sheet():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=123")
    wb.set_cell_contents(sheet_name, "B1", "=C1")
    wb.set_cell_contents(sheet_name, "C1", "=B1")
    assert wb.get_cell_value(sheet_name, "C1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    assert wb.get_cell_value(sheet_name, "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    # Point to cycle
    wb.set_cell_contents(sheet_name, "D1", "=B1")
    assert wb.get_cell_value(sheet_name, "D1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    # Test multiple cycles
    wb.set_cell_contents(sheet_name, "E1", "=F1")
    wb.set_cell_contents(sheet_name, "F1", "=G1")
    wb.set_cell_contents(sheet_name, "G1", "=E1")
    assert wb.get_cell_value(sheet_name, "E1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "F1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "G1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    wb.set_cell_contents(sheet_name, "H1", "=E1")
    wb.set_cell_contents(sheet_name, "I1", "=H1")
    assert wb.get_cell_value(sheet_name, "H1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "I1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "G1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    # Fixing E1 should fix all of them
    wb.set_cell_contents(sheet_name, "E1", "=123")
    assert wb.get_cell_value(sheet_name, "E1") == 123
    assert wb.get_cell_value(sheet_name, "F1") == 123
    assert wb.get_cell_value(sheet_name, "G1") == 123
    assert wb.get_cell_value(sheet_name, "H1") == 123
    assert wb.get_cell_value(sheet_name, "I1") == 123


def test_type_conversions():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "'     123")
    wb.set_cell_contents(sheet_name, "B1", "5.3")
    wb.set_cell_contents(sheet_name, "C1", "=A1*B1")

    # what do we do about floating point arithmetic?
    assert wb.get_cell_value(sheet_name, "C1") == Decimal('123')*Decimal('5.3')

def test_propogate_cell_errors():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    wb.set_cell_contents(sheet_name, "B1", "=C1")
    wb.set_cell_contents(sheet_name, "C1", "=B1/0")

    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    wb.set_cell_contents(sheet_name, "C1", "=1")
    assert wb.get_cell_value(sheet_name, "A1") == 1

    # Manually setting C1 to an error should propogate it to A1
    wb.set_cell_contents(sheet_name, "C1", "=#DIV/0!")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO 

    wb.set_cell_contents(sheet_name, "C1", "=#REF!")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.BAD_REFERENCE 

def test_cell_contents():

    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "1")

    wb.set_cell_contents(sheet_name, "D14", "Green")
    assert wb.get_sheet_extent(sheet_name) == (4, 14) 
    wb.set_cell_contents(sheet_name, "D14", None)
    assert wb.get_sheet_extent(sheet_name) == (1, 1)


    wb.set_cell_contents(sheet_name, "BB2", "3")
    wb.set_cell_contents(sheet_name, "C3", "Hello")
    wb.set_cell_contents(sheet_name, "D4", "World")

    assert "3" == wb.get_cell_contents("Sheet1", "BB2")
    assert 3 == wb.get_cell_value("Sheet1", "BB2")

    wb.set_cell_contents(sheet_name, "BB2", "3.0")
    assert 3 == wb.get_cell_value("Sheet1", "BB2")
    assert "3.0" == wb.get_cell_contents("Sheet1", "BB2")

    assert "1" == wb.get_cell_contents("Sheet1", "A1")
    assert "Hello" == wb.get_cell_contents("Sheet1", "C3")
    assert "World" == wb.get_cell_contents("Sheet1", "D4")
    
    with pytest.raises(ValueError):
        wb.set_cell_contents(sheet_name, "LMAOHAHAH696969", "World")
        wb.set_cell_contents(sheet_name, "!2", "World")

    wb.set_cell_contents(sheet_name, "ZZL32", None)

    assert wb.get_sheet_extent(sheet_name) == (26*2 + 2, 4) 
    wb.del_sheet("sheet1")
    assert len(wb.list_sheets()) == 0
    with pytest.raises(KeyError):
        wb.set_cell_contents(sheet_name, "A1", "1")


def test_new_sheet_with_empty_name():
    wb = sheets.Workbook()
    with pytest.raises(ValueError):
        idx, name = wb.new_sheet("")
    idx, name = wb.new_sheet(None)
    assert name == "Sheet1"
    idx, name = wb.new_sheet(None)
    assert name == "Sheet2"
    idx, name = wb.new_sheet(None)
    assert name == "Sheet3"

      

def test_get_sheet_extent():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    assert sheet_name == name
    assert (0 == idx)
    assert sheet_name in wb.list_sheets()
    assert wb.get_sheet_extent(sheet_name) == (0,0)

    wb.set_cell_contents(sheet_name, "A1", "1")
    assert wb.get_sheet_extent(sheet_name) == (1,1)
    wb.set_cell_contents(sheet_name, "A1", None)
    assert wb.get_sheet_extent(sheet_name) == (0,0)

    wb.set_cell_contents(sheet_name, "D14", "Green")
    assert wb.get_sheet_extent(sheet_name) == (4, 14) 
    
    wb.set_cell_contents(sheet_name, "A12", "Green")
    
    assert wb.get_sheet_extent(sheet_name) == (4, 14) 
    
    wb.set_cell_contents(sheet_name, "D14", None)
    assert wb.get_sheet_extent(sheet_name) == (1, 12)


def test_set_none():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    print("NOW WE ADD THE 1")
    wb.set_cell_contents(sheet_name, "B1", "1")
    sheet = wb.sheets["sheet1"].data

    for key, cell in sheet.items():
        assert(cell.get_value() == 1)
    print(sheet)
    assert (1==1)

def test_reset_circ_ref():
    wb = sheets.Workbook()
    sheet_name = "Sheet1".lower()
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    wb.set_cell_contents(sheet_name, "B1", "=A1")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    wb.set_cell_contents(sheet_name, "A1", "")


    assert wb.get_cell_value(sheet_name, "B1") == 0
    assert wb.get_cell_value(sheet_name, "A1") is None

    print("!!!!!!!!!!!!!!!!")
    print(wb.get_cell_contents(sheet_name, "A1"))
    print(wb.sheets[sheet_name].get_cell("A1"))
    print("NOW WE ADD CONTENTS")

    hidden_name = get_hidden_name(sheet_name)
    location = "A1"

    if hidden_name + "!" + location in wb.dependencies:
        print("Cell", location)
        print("LKKSDKFJSn WE ARE IN THE DEP LIST")
        cell = wb.dependencies[hidden_name+"!"+location]
        print(cell.get_contents())
        print(cell)
    wb.set_cell_contents(sheet_name, "A1", "=B1")
    
    print(wb.get_cell_value(sheet_name, "A1"))
    print(wb.get_cell_value(sheet_name, "B1"))
    
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value(sheet_name, "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_setting_errors():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "#div/0!")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    wb.set_cell_contents(sheet_name, "A1", "#ref!")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.BAD_REFERENCE
    wb.set_cell_contents(sheet_name, "A1", "=#ref!")
    assert wb.get_cell_value(sheet_name, "A1").get_type() == CellErrorType.BAD_REFERENCE

def test_set_empty_cell():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "A1", "")
    assert wb.get_cell_value(sheet_name, "A1") is None

    wb.set_cell_contents(sheet_name, "A1", "=B1")
    assert wb.get_cell_value(sheet_name, "A1") == 0

    wb.set_cell_contents(sheet_name, "B1", "")
    assert wb.get_cell_value(sheet_name, "A1") == 0

def test_lower_case():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "a1", "1")
    assert wb.get_cell_value(sheet_name, "A1") == 1
    assert wb.get_cell_value(sheet_name, "a1") == 1

def test_lower_case_formula_cells():
    wb = sheets.Workbook()
    sheet_name = "Sheet1"
    idx, name = wb.new_sheet(sheet_name)
    wb.set_cell_contents(sheet_name, "a1", "=B1")
    wb.set_cell_contents(sheet_name, "b1", "1")
    assert wb.get_cell_value(sheet_name, "A1") == 1
    assert wb.get_cell_value(sheet_name, "a1") == 1

    wb.set_cell_contents(sheet_name, "a1", "1")
    wb.set_cell_contents(sheet_name, "b1", "3")

    wb.set_cell_contents(sheet_name, "a2", "=a1+b1")
    assert wb.get_cell_value(sheet_name, "A2") == 4
    assert wb.get_cell_value(sheet_name, "a2") == 4
    assert wb.get_cell_value(sheet_name, "B1") == 3
    assert wb.get_cell_value(sheet_name, "b1") == 3

    wb.set_cell_contents(sheet_name, "b1", "7")
    assert wb.get_cell_value(sheet_name, "A2") == 8
    assert wb.get_cell_value(sheet_name, "a2") == 8
    assert wb.get_cell_value(sheet_name, "B1") == 7
