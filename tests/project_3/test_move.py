
import context
import pytest
import sheets
import sheets.CellErrorType as CellErrorType

@pytest.fixture
def wb():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    return wb

def test_move_cells_simple(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "A2", "World")
    wb.move_cells("Sheet1", "A1", "A2", "A3", 'Sheet1')
    assert wb.get_cell_contents("Sheet1", "A3") == "Hello"
    assert wb.get_cell_contents("Sheet1", "A4") == "World"

def test_move_overlapping(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "A2", "World")
    wb.set_cell_contents("Sheet1", "A3", "Hi")
    wb.move_cells("Sheet1", "A1", "A3", "A2", 'Sheet1')
    assert wb.get_cell_contents("Sheet1", "A2") == "Hello"
    assert wb.get_cell_contents("Sheet1", "A3") == "World"
    assert wb.get_cell_contents("Sheet1", "A4") == "Hi"

def test_different_corners(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "B2", "World")
    wb.move_cells("Sheet1", "A2", "B1", "B3", 'Sheet1')
    assert wb.get_cell_contents("Sheet1", "B3") == "Hello"
    assert wb.get_cell_contents("Sheet1", "C4") == "World"

def test_move_cells_across_sheets(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "A2", "World")
    wb.new_sheet("Sheet2")
    wb.move_cells("Sheet1", "A1", "A2", "A1", "Sheet2")
    assert wb.get_cell_contents("Sheet2", "A1") == "Hello"

def test_overlap(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "A2", "World")
    wb.set_cell_contents("Sheet1", "B1", "asdf")
    wb.set_cell_contents("Sheet1", "B2", "jkl;")
    wb.move_cells("Sheet1", "A1", "B2", "B2", 'Sheet1')

    assert wb.get_cell_contents("Sheet1", "B2") == "Hello"
    assert wb.get_cell_contents("Sheet1", "B3") == "World"
    assert wb.get_cell_contents("Sheet1", "C2") == "asdf"
    assert wb.get_cell_contents("Sheet1", "C3") == "jkl;"

def test_valid_move_range(wb):
    wb.set_cell_contents("Sheet1", "ZZZY9998", "Hello")
    wb.set_cell_contents("Sheet1", "ZZZY9999", "World")
    wb.set_cell_contents("Sheet1", "ZZZZ9998", "asdf")
    wb.set_cell_contents("Sheet1", "ZZZZ9999", "jkl;")

    with pytest.raises(ValueError) as e:
        wb.move_cells("Sheet1", "ZZZY9998", "ZZZZ9999", "ZZZY9999", 'Sheet1')
    wb.move_cells("Sheet1", "ZZZY9998", "ZZZY9998", "ZZZZ9999", 'Sheet1')

    # move the cells to A1
    wb.move_cells("Sheet1", "ZZZZ9999", "ZZZZ9999", "A1", 'Sheet1')
    assert wb.get_cell_contents("Sheet1", "A1") == "Hello"

    # wb.set_cell_contents("Sheet1", "A1", 'lol')
    # wb.move_cells("Sheet1", "A1", "B2", "ZZZ", 'Sheet1')



def test_none(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "A2", "World")
    wb.set_cell_contents("Sheet1", "B1", "asdf")
    wb.set_cell_contents("Sheet1", "B2", "jkl;")
    wb.move_cells("Sheet1", "A1", "B2", "B2", None)

    assert wb.get_cell_contents("Sheet1", "B2") == "Hello"
    assert wb.get_cell_contents("Sheet1", "B3") == "World"
    assert wb.get_cell_contents("Sheet1", "C2") == "asdf"
    assert wb.get_cell_contents("Sheet1", "C3") == "jkl;"

    wb.set_cell_contents("Sheet1", "B2", None)

    wb.move_cells("Sheet1", "A1", "B2", "B2", 'Sheet1')
    assert wb.get_cell_contents("Sheet1", "C3") == None
    assert wb.get_cell_contents("Sheet1", "C2") == None
    assert wb.get_cell_contents("Sheet1", "B2") == None
    assert wb.get_cell_contents("Sheet1", "B3") == None
def test_move_multiple(wb):
    wb.set_cell_contents("Sheet1", "A1", "Hello")
    wb.set_cell_contents("Sheet1", "A2", "World")
    wb.set_cell_contents("Sheet1", "B1", "asdf")
    wb.set_cell_contents("Sheet1", "B2", "jkl;")
    wb.move_cells("Sheet1", "A1", "B2", "B2", 'Sheet1')

    assert wb.get_cell_contents("Sheet1", "B2") == "Hello"
    assert wb.get_cell_contents("Sheet1", "B3") == "World"
    assert wb.get_cell_contents("Sheet1", "C2") == "asdf"
    assert wb.get_cell_contents("Sheet1", "C3") == "jkl;"

    wb.set_cell_contents("Sheet1", "B2", "Bruh")

    wb.move_cells("Sheet1", "A1", "B2", "B2", 'Sheet1')
    assert wb.get_cell_contents("Sheet1", "C3") == "Bruh"
    assert wb.get_cell_contents("Sheet1", "C2") == None
    assert wb.get_cell_contents("Sheet1", "B2") == None
    assert wb.get_cell_contents("Sheet1", "B3") == None

@pytest.fixture
def workbook1():
    wb = sheets.Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.set_cell_contents('Sheet1', 'A1', "'123")
    wb.set_cell_contents('Sheet1', 'B1', "5.3")
    wb.set_cell_contents('Sheet1', 'C1', "=A1*B1")
    return wb

def test_copy_cells_same_sheet_formula(workbook1):
    workbook1.copy_cells('Sheet1', 'A1', 'C1', 'A2')
    assert workbook1.get_cell_contents('Sheet1', 'A2') == "'123"
    assert workbook1.get_cell_contents('Sheet1', 'B2') == "5.3"
    assert workbook1.get_cell_contents('Sheet1', 'C2') == "=A2*B2"

def test_copy_cells_different_sheet_formula(workbook1):
    workbook1.copy_cells('Sheet1', 'A1', 'C1',  'A1', 'Sheet2')
    assert workbook1.get_cell_contents('Sheet2', 'A1') == "'123"
    assert workbook1.get_cell_contents('Sheet2', 'B1') == "5.3"
    assert workbook1.get_cell_contents('Sheet2', 'C1') == "=A1*B1"

def test_move_cells_same_sheet_formula(workbook1):
    workbook1.move_cells('Sheet1', 'A1', 'C1', 'A2', None)
    assert workbook1.get_cell_contents('Sheet1', 'A2') == "'123"
    assert workbook1.get_cell_contents('Sheet1', 'B2') == "5.3"
    assert workbook1.get_cell_contents('Sheet1', 'C2') == "=A2*B2"

def test_move_cells_different_sheet_formula(workbook1):
    workbook1.move_cells('Sheet1', 'A1', 'C1', 'A1', 'Sheet2')
    assert workbook1.get_cell_contents('Sheet2', 'A1') == "'123"
    assert workbook1.get_cell_contents('Sheet2', 'B1') == "5.3"
    assert workbook1.get_cell_contents('Sheet2', 'C1') == "=A1*B1"

def test_move_cells_diagonal(workbook1):
    workbook1.move_cells('Sheet1', 'A1', 'C1', 'B2', None)
    assert workbook1.get_cell_contents('Sheet1', 'B2') == "'123"
    assert workbook1.get_cell_contents('Sheet1', 'C2') == "5.3"
    assert workbook1.get_cell_contents('Sheet1', 'D2') == "=B2*C2"

def test_move_cells_diagonal_different_sheet(workbook1):
    workbook1.move_cells('Sheet1', 'A1', 'C1', 'B2', "Sheet2")
    assert workbook1.get_cell_contents('Sheet2', 'B2') == "'123"
    assert workbook1.get_cell_contents('Sheet2', 'C2') == "5.3"
    assert workbook1.get_cell_contents('Sheet2', 'D2') == "=B2*C2"

def test_copy_cells_diagonal(workbook1):
    workbook1.copy_cells('Sheet1', 'A1', 'C1', 'B2')
    assert workbook1.get_cell_contents('Sheet1', 'B2') == "'123"
    assert workbook1.get_cell_contents('Sheet1', 'C2') == "5.3"
    assert workbook1.get_cell_contents('Sheet1', 'D2') == "=B2*C2"

def test_copy_cells_diagonal_different_sheet(workbook1):
    workbook1.copy_cells('Sheet1', 'A1', 'C1', 'B2', 'Sheet2')
    assert workbook1.get_cell_contents('Sheet2', 'B2') == "'123"
    assert workbook1.get_cell_contents('Sheet2', 'C2') == "5.3"
    assert workbook1.get_cell_contents('Sheet2', 'D2') == "=B2*C2"


def test_move_cell_rel_ref():
    #Tests moving various cells to a different column
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","A1" , "=C1")
    wb.set_cell_contents("Sheet1","A2" , "=$C2")
    wb.set_cell_contents("Sheet1","A3" , "=C$3")
    wb.set_cell_contents("Sheet1","A4" , "=$C$4")
    wb.move_cells("Sheet1","A1", "A4", "C8")
    assert wb.get_cell_contents("Sheet1", "C8") == "=E8"
    assert wb.get_cell_contents("Sheet1", "C9") == "=$C9" # THIS SHOULD BE CIRC REF
    assert wb.get_cell_value("Sheet1", "C9").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_contents("Sheet1", "C10") =="=E$3"
    assert wb.get_cell_contents("Sheet1", "C11") == "=$C$4"

def test_move_cell_abs_ref_circle_ref():
    #Tests moving various cells to a different column
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","A2" , "=A1")
    wb.set_cell_contents("Sheet1","A1" , "=$A$2")
    wb.move_cells("Sheet1","A1", "A1", "A2")
    assert wb.get_cell_contents("Sheet1", "A2") == "=$A$2"
    assert wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE



def test_move_cell_circle_ref():
    #Tests moving various cells to a different column
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","A1" , "=A2")
    wb.set_cell_contents("Sheet1","A2" , "=A3")
    wb.set_cell_contents("Sheet1","A3" , "=A1")
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "A3").get_type() == CellErrorType.CIRCULAR_REFERENCE
    wb.move_cells("Sheet1","A1", "A3", "C8")
    assert wb.get_cell_value("Sheet1", "C8").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "C9").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "C10").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_contents("Sheet1", "C8") == "=C9"
    assert wb.get_cell_contents("Sheet1", "C9") == "=C10"
    assert wb.get_cell_contents("Sheet1", "C10") =="=C8"

def test_copy_cell_circle_ref():
    #Tests moving various cells to a different column
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","A1" , "=A2")
    wb.set_cell_contents("Sheet1","A2" , "=A3")
    wb.set_cell_contents("Sheet1","A3" , "=A1")
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "A3").get_type() == CellErrorType.CIRCULAR_REFERENCE
    wb.copy_cells("Sheet1","A1", "A3", "C8")
    assert wb.get_cell_value("Sheet1", "C8").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "C9").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1", "C10").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_contents("Sheet1", "C8") == "=C9"
    assert wb.get_cell_contents("Sheet1", "C9") == "=C10"
    assert wb.get_cell_contents("Sheet1", "C10") =="=C8"


def test_move_out_of_bounds():
    #Tests moving various cells to a different column
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")

    wb.set_cell_contents("Sheet1","B1" , "=A1 + 6")
    wb.move_cells("Sheet1","B1", "B1", "A1")
    assert wb.get_cell_contents("Sheet1", "A1") == "=#REF!+6"

    wb.set_cell_contents("Sheet1","A2" , "=A1 + 7")
    wb.move_cells("Sheet1","A2", "A2", "A1")
    assert wb.get_cell_contents("Sheet1", "A1") == "=#REF!+7"

    wb.set_cell_contents("Sheet1","A1" , "=ZZZZ1 + 8")
    wb.move_cells("Sheet1","A1", "A1", "B1")
    assert wb.get_cell_contents("Sheet1", "B1") == "=#REF!+8"

    wb.set_cell_contents("Sheet1","A1" , "=A9999 + 9")
    wb.move_cells("Sheet1","A1", "A1", "A2")
    assert wb.get_cell_contents("Sheet1", "A2") == "=#REF!+9"

def test_move_cell_to_make_circref():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","B1" , "5")
    wb.set_cell_contents("Sheet1","A1" , "=B1")
    wb.set_cell_contents("Sheet1","C1" , "=$A$1")

    wb.move_cells("Sheet1","C1", "C1", "B1")
    assert wb.get_cell_value("Sheet1","B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1","A1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_copy_cell_to_make_circref():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","B1" , "5")
    wb.set_cell_contents("Sheet1","A1" , "=B1")
    wb.set_cell_contents("Sheet1","C1" , "=$A$1")

    wb.copy_cells("Sheet1","C1", "C1", "B1")
    assert wb.get_cell_value("Sheet1","B1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert wb.get_cell_value("Sheet1","A1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_copy_cell_to_sheet_circref():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.new_sheet("Sheet2")

    wb.set_cell_contents("Sheet1","A1" , "=Sheet2!A1")
    wb.copy_cells("Sheet1","A1", "A1", "A1", to_sheet= "Sheet2")
    assert wb.get_cell_value("Sheet2","A1").get_type() == CellErrorType.CIRCULAR_REFERENCE

def test_copy_cells_to_out_of_world():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","A1" , "=1")
    wb.set_cell_contents("Sheet1","B1" , "=2")
    wb.set_cell_contents("Sheet1","A2" , "=3")
    wb.set_cell_contents("Sheet1","B2" , "=4")
    with pytest.raises(ValueError) as e:
        wb.move_cells("Sheet1", "A1", "B1", "ZZZZ9999")

def test_copy_extremely_complicated_formula():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1","A1" , "=1")
    wb.set_cell_contents("Sheet1","B1" , "=A1")
    wb.set_cell_contents("Sheet1","C1" , "=A1")
    wb.set_cell_contents("Sheet1","B1" , "=A1")
    wb.set_cell_contents("Sheet1","B1" , "=A1")
