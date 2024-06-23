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

def test_decimal_comps(wb):
    set_it(wb,"=7>2")
    assert get_it(wb) == True
    set_it(wb,"=2>5")
    assert get_it(wb) == False
    set_it(wb,"=3<9")
    assert get_it(wb) == True
    set_it(wb,"=5<1")
    assert get_it(wb) == False

    set_it(wb,"= 5 >=5")
    assert get_it(wb) == True

    set_it(wb,"= 5>=4 ")
    assert get_it(wb) == True
    set_it(wb,"= 4 >= 9")
    assert get_it(wb) == False
    set_it(wb,"=8 == 8")
    assert get_it(wb) == True
    set_it(wb,"= 8 = 8")
    assert get_it(wb) == True
    set_it(wb,"=9 == 3")
    assert get_it(wb) == False
    set_it(wb,"= 9 <> 2")
    assert get_it(wb) == True
    set_it(wb,"= 9 != 2")
    assert get_it(wb) == True
    set_it(wb,"= 9 <> 5")
    assert get_it(wb) == True
    set_it(wb,"= 9 != 9")
    assert get_it(wb) == False
    set_it(wb,"= 4.5 <> 4.5")
    assert get_it(wb) == False

def test_string_comps(wb):
    # Testing basic string equality and inequality
    set_it(wb, '="hello" == "hello"')
    assert get_it(wb) == True
    set_it(wb, '="Hello" == "hello"')
    assert get_it(wb) == True  # Case-insensitive comparison
    set_it(wb, '="apple" != "Apple"')
    assert get_it(wb) == False  # Case-insensitive comparison
    set_it(wb, '="test" <> "test"')
    assert get_it(wb) == False

    # Testing greater than and less than comparisons
    set_it(wb, '="banana" < "cherry"')
    assert get_it(wb) == True
    set_it(wb, '="banana" > "Apple"')
    assert get_it(wb) == True  # 'b' has a higher ASCII value than 'A' in lowercase
    set_it(wb, '="grape" < "Grape"')
    assert get_it(wb) == False  # Case-insensitive comparison

    # Testing greater than or equal to and less than or equal to comparisons
    set_it(wb, '="alpha" >= "alpha"')
    assert get_it(wb) == True
    set_it(wb, '="Alpha" <= "alpha"')
    assert get_it(wb) == True  # Case-insensitive comparison
    set_it(wb, '="beta" >= "gamma"')
    assert get_it(wb) == False
    set_it(wb, '="delta" <= "epsilon"')
    assert get_it(wb) == True



def test_string_comps_extended(wb):
    set_it(wb, '= "A" != "B"')
    get_it(wb) == True
    set_it(wb, '= "A" != "A"')
    get_it(wb) == False
    set_it(wb, '= "A" <> "B"')
    get_it(wb) == True
    set_it(wb, '= "A" <> "A"')
    get_it(wb) == False

    # Testing comparisons with empty strings
    set_it(wb, '="" == ""')
    assert get_it(wb) == True
    set_it(wb, '=" " == ""')
    assert get_it(wb) == False  # A space is not equal to an empty string
    set_it(wb, '="" < "a"')
    assert get_it(wb) == True  # Empty string is less than any character

    # Testing strings with spaces
    set_it(wb, '="hello world" == "hello world"')
    assert get_it(wb) == True
    set_it(wb, '=" hello" == "hello "')
    assert get_it(wb) == False  # Space at different positions
    set_it(wb, '="hello world" > "Hello World"')
    assert get_it(wb) == False  # Case-insensitive comparison

    # Testing strings with numbers
    set_it(wb, '="123" == "123"')
    assert get_it(wb) == True
    set_it(wb, '="123" < "124"')
    assert get_it(wb) == True
    set_it(wb, '="123" > "122"')
    assert get_it(wb) == True
    set_it(wb, '="123abc" > "123ABC"')
    assert get_it(wb) == False  # Case-insensitive comparison

    # Testing strings that are numerically equal but typecast as strings
    set_it(wb, '="123" == "123.0"')
    assert get_it(wb) == False  # String comparison, not numerical comparison

    # Testing case sensitivity explicitly
    set_it(wb, '="Case" == "case"')
    assert get_it(wb) == True  # Case-insensitive comparison
    set_it(wb, '="case" > "CASE"')
    assert get_it(wb) == False  # Case-insensitive comparison assumes they are equal

    # Edge cases with special characters
    set_it(wb, '="!" < "@"')
    assert get_it(wb) == True
    set_it(wb, '="$" == "$"')
    assert get_it(wb) == True
    set_it(wb, '="~" > "^"')
    assert get_it(wb) == True

    # Comparison of strings with mixed characters
    set_it(wb, '="abc123" == "abc123"')

    assert get_it(wb) == True
    set_it(wb, '="abc123" < "abc124"')
    assert get_it(wb) == True
    set_it(wb, '="abc" < "abc123"')
    assert get_it(wb) == True  # String with only letters is less than string with letters followed by numbers


def test_equals(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet1", "A2", "5")
    assert wb.get_cell_value("Sheet1", "A1") == 5
    wb.set_cell_contents("Sheet1", "A3", "=A1=A2")
    assert wb.get_cell_value("Sheet1", "A3") 

def test_bools(wb):
    wb.set_cell_contents("Sheet1", "A1", '="bruh"')
    assert wb.get_cell_value("Sheet1", "A1") == "bruh"
    wb.set_cell_contents('Sheet1', 'A1', 'true')
    assert isinstance(wb.get_cell_value('Sheet1', 'A1'), bool)
    assert wb.get_cell_value('Sheet1', 'A1')

    wb.set_cell_contents('Sheet1', 'A2', '=FaLsE')
    assert isinstance(wb.get_cell_value('Sheet1', 'A2'), bool)
    assert not wb.get_cell_value('Sheet1', 'A2')

def test_number_comparison(wb):
    wb.set_cell_contents('Sheet1', 'A1', '5')
    wb.set_cell_contents('Sheet1', 'A2', '5')
    wb.set_cell_contents('Sheet1', 'A3', '6')
    wb.set_cell_contents('Sheet1', 'A4', '4')
    wb.set_cell_contents('Sheet1', 'A5', '=A1=A2')
    assert wb.get_cell_value('Sheet1', 'A5')
    wb.set_cell_contents('Sheet1', 'A6', '=A1==A3')
    assert not wb.get_cell_value('Sheet1', 'A6')
    wb.set_cell_contents('Sheet1', 'A7', '=A1>A4')
    assert wb.get_cell_value('Sheet1', 'A7')
    wb.set_cell_contents('Sheet1', 'A8', '=A1<A3')
    assert wb.get_cell_value('Sheet1', 'A8')
    wb.set_cell_contents('Sheet1','b7' ,'=A1 <> A3')
    assert wb.get_cell_value('Sheet1', 'b7')

def test_boolean_presedence(wb):
    """what are these rules????"""
    wb.set_cell_contents('Sheet1', 'A1', '\'12')
    wb.set_cell_contents('Sheet1', 'A2', '12')
    wb.set_cell_contents('Sheet1', 'A3', '=A1>A2')
    assert wb.get_cell_value('Sheet1', 'A3')

    wb.set_cell_contents('Sheet1', 'A1', '\'TRUE')
    wb.set_cell_contents('Sheet1', 'A2', 'FALSE')
    wb.set_cell_contents('Sheet1', 'A3', '=A1>A2')
    assert not wb.get_cell_value('Sheet1', 'A3')

def test_case_insensitivity(wb):
    wb.set_cell_contents('Sheet1', 'A1', 'BLUE')
    wb.set_cell_contents('Sheet1', 'A2', 'blue')
    wb.set_cell_contents('Sheet1', 'A3', '=A1=A2')
    wb.set_cell_contents('Sheet1', 'A4', '=A1<>A2')
    assert wb.get_cell_value('Sheet1', 'A3')
    assert not wb.get_cell_value('Sheet1', 'A4')

def test_lowercase_version(wb):
    wb.set_cell_contents('Sheet1', 'A1', '="A" < "["')
    assert not wb.get_cell_value('Sheet1', 'A1')
    wb.set_cell_contents('Sheet1', 'A2', '="a" < "["')
    assert not wb.get_cell_value('Sheet1', 'A1')

def test_empty_cell_bools(wb):
    assert not wb.get_cell_value('Sheet1', 'A1')
    wb.set_cell_contents('Sheet1', 'A1', '')
    assert not wb.get_cell_value('Sheet1', 'A1')

def test_cell_addition_bools(wb):
    """ Booleans should be converted to numbers when added"""
    wb.set_cell_contents('Sheet1', 'A1', 'true')
    wb.set_cell_contents('Sheet1', 'A2', 'true')
    wb.set_cell_contents('Sheet1', 'A3', '=A1+A2')
    assert wb.get_cell_value('Sheet1', 'A3') == 2

    wb.set_cell_contents('Sheet1', 'A4', 'false')
    wb.set_cell_contents('Sheet1', 'A5', '=A1+A4')
    assert wb.get_cell_value('Sheet1', 'A5') == 1
    
def test_comparison_precedence(wb):
    # TODO: Somehow we are auto stripping whitespaces from string literals.
    # Is this the way it should?
    wb.set_cell_contents('Sheet1', 'A2', 'type')
    wb.set_cell_contents('Sheet1', 'A1', "= A2 = B2 & \"type\"")
    # B2 & " type" is evaluated first
    assert wb.get_cell_value('Sheet1', 'A1')
    wb.set_cell_contents('Sheet1', 'A2', 'lol')
    assert not wb.get_cell_value('Sheet1', 'A1')

def test_functions_in_functions(wb):
    wb.set_cell_contents('Sheet1', 'A1', '5')
    wb.set_cell_contents('Sheet1', 'B1', '1')
    wb.set_cell_contents('Sheet1', 'C1', '2')
    wb.set_cell_contents('Sheet1', 'D1', '14')


    wb.set_cell_contents('Sheet1', 'A2', '=OR(AND(A1 > 5, B1 < 2), AND(C1 < 6, D1 = 14))')
    assert wb.get_cell_value('Sheet1', 'A2')

def test_implicit_conversion(wb):
    wb.set_cell_contents('Sheet1', "A2", '=true')
    set_it(wb,'= A2 &  " bruh"')
    assert get_it(wb) == "TRUE bruh"
    set_it(wb,'= A2 + "1"')
    assert get_it(wb) == 2
    set_it(wb,'=IF("true", 1, 5) ')
    assert get_it(wb) == 1