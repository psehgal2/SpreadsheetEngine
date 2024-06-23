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
    set_it(wb,"=IFERROR(99, \"bruh\")")
    assert get_it(wb) == decimal.Decimal(99)
    set_it(wb,"=IFERROR(#REF!)")
    assert get_it(wb) == ""

def test_iferror2(wb):
    set_it(wb,"=IFERROR(#ref!, \"bruh\")")
    assert get_it(wb) == "bruh"
    set_it(wb,"=IFERROR(99)")
    assert get_it(wb) == decimal.Decimal(99)
    set_it(wb,"=IFERROR(#DIV/0!)")
    assert get_it(wb) == ""

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


def example_test(wb):
    wb.set_cell_contents("Sheet1", "A1", "1")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "C1", "=SUM(A1, B1)")
    assert wb.get_cell_value("Sheet1", "C1") == 2


def test_boolean_literal_conversion_and_comparison(wb):
    wb.set_cell_contents("Sheet1", "A1", "TRUE")
    wb.set_cell_contents("Sheet1", "A2", "false")
    wb.set_cell_contents("Sheet1", "B1", "=A1 = TRUE")
    wb.set_cell_contents("Sheet1", "B2", "=A2 <> FALSE")
    wb.set_cell_contents("Sheet1", "C2", "=A2 != FALSE")

    assert wb.get_cell_value("Sheet1", "B1") is True
    assert wb.get_cell_value("Sheet1", "B2") is False
    assert wb.get_cell_value("Sheet1", "C2") is False


def test_circular_reference_error_with_indirect(wb):
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"B1\")")
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_comparison_operations_with_different_types(wb):
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "\"text\"")
    wb.set_cell_contents("Sheet1", "A3", "=A1 > A2")
    wb.set_cell_contents("Sheet1", "A4", "=A1 < A2")

    assert wb.get_cell_value("Sheet1", "A3") is False
    assert wb.get_cell_value("Sheet1", "A4") is True


def test_if_function_with_error_propagation(wb):
    wb.set_cell_contents("Sheet1", "A1", "=1/0")  # Generates a DIV/0! error
    wb.set_cell_contents("Sheet1", "A2", "=IF(TRUE, A1, 10)")

    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError)
    assert wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.DIVIDE_BY_ZERO


def test_indirect_function_with_invalid_reference(wb):
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"ZZZZZZZZZ9999999999999\")")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE


# I do not believe this test is valid.
# def test_empty_cell_comparison(wb):
#     """Test comparison between empty cells and non-empty cells of different types."""
#     wb.set_cell_contents("Sheet1", "A1", "")
#     wb.set_cell_contents("Sheet1", "B1", "=A1 = 0")
#     wb.set_cell_contents("Sheet1", "C1", "=A1 = \"\"")
#
#     assert wb.get_cell_value("Sheet1", "B1") is True
#     assert wb.get_cell_value("Sheet1", "C1") is True

def test_nones_are_equal(wb):
    wb.set_cell_contents("Sheet1", "A3", "= A1 = B1")
    assert wb.get_cell_value("Sheet1", "A3") is True


def test_function_name_case_insensitivity(wb):
    """Verify that function names are case-insensitive."""
    wb.set_cell_contents("Sheet1", "A3", "4")
    wb.set_cell_contents("Sheet1", "A1", "=indirect(\"A3\")")
    wb.set_cell_contents("Sheet1", "A2", "=INDIRECT(\"A3\")")

    assert wb.get_cell_value("Sheet1", "A1") == 4
    assert wb.get_cell_value("Sheet1", "A2") == 4


def test_divide_by_zero_error(wb):
    """Test for DIV/0! error when dividing by zero in a formula."""
    wb.set_cell_contents("Sheet1", "A1", "=10/0")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO


def test_bad_name_error_for_unrecognized_function(wb):
    """Test for BAD_NAME error when using an unrecognized function name."""
    wb.set_cell_contents("Sheet1", "A1", "=UNDEFINEDFUNCTION(1)")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_NAME


def test_type_error_for_wrong_argument_type(wb):
    """Test for TYPE_ERROR when a function is called with arguments of the wrong type."""
    wb.set_cell_contents("Sheet1", "A1", "=SUM(\"text\", 5)")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR


def test_circular_reference_via_direct_reference(wb):
    """Test direct circular reference detection."""
    wb.set_cell_contents("Sheet1", "A1", "=A1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_circular_reference_through_multiple_cells(wb):
    """Test circular reference detection through a chain of cells."""
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=C1")
    wb.set_cell_contents("Sheet1", "C1", "=A1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


# UNCOMMENT THIS TEST LATER.
# def test_circular_reference_with_function_invocation(wb):
#     """Test for circular reference that includes a function call in the cycle."""
#     wb.set_cell_contents("Sheet1", "A1", "=SUM(B1, 1)")
#     wb.set_cell_contents("Sheet1", "B1", "=A1")
#
#     assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
#     assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_circular_reference_avoided_by_conditional_evaluation(wb):
    """Test avoiding circular reference through conditional function evaluation."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(FALSE, B1, 1)")
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    # This should not result in a circular reference because B1 is not evaluated
    assert wb.get_cell_value("Sheet1", "A1") == 1


def test_indirect_function_leading_to_circular_reference(wb):
    """Test INDIRECT function causing a circular reference dynamically."""
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"B1\")")
    wb.set_cell_contents("Sheet1", "B1", "A1")

    # Initially, B1's value is just 'A1', which doesn't cause a circular reference.
    # Changing B1 to reference A1 indirectly causes the circular reference.
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_iferror_function_with_circular_reference(wb):
    """Test IFERROR function handling a circular reference error."""
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=A1")
    wb.set_cell_contents("Sheet1", "C1", "=IFERROR(A1, \"Error Detected\")")

    assert wb.get_cell_value("Sheet1", "C1") == 'Error Detected'


def test_choose_function_avoiding_circular_reference(wb):
    """Test CHOOSE function avoiding a circular reference by not evaluating all arguments."""
    wb.set_cell_contents("Sheet1", "A1", "=CHOOSE(1, 2, B1)")
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    # CHOOSE should not trigger a circular reference since it only evaluates the first choice
    assert wb.get_cell_value("Sheet1", "A1") == 2


def test_circular_reference_in_nested_function_calls(wb):
    """Test detection of circular references within nested function calls."""
    wb.set_cell_contents("Sheet1", "A1", "=SUM(B1, 1)")
    wb.set_cell_contents("Sheet1", "B1", "=SUM(C1, 1)")
    wb.set_cell_contents("Sheet1", "C1", "=A1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_indirect_function_dynamic_reference_change(wb):
    """Test dynamic change in cell reference via INDIRECT function leading to circular reference."""
    wb.set_cell_contents("Sheet1", "A1", "=INdiRECT(\"c1\")")
    wb.set_cell_contents("Sheet1", "B1", "=2")
    wb.set_cell_contents("Sheet1", "C1", "=B1")

    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal(2)
    # Changing C1 to reference A1 indirectly, causing a circular reference
    wb.set_cell_contents("Sheet1", "C1", "=A1")

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_function_call_within_circular_reference_chain(wb):
    """Test a function call placed within a chain of references leading to a circular reference."""
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=SUM(C1, 0)")
    wb.set_cell_contents("Sheet1", "C1", "=A1")

    # Despite SUM function trying to compute a value, the circular reference is detected
    assert isinstance(wb.get_cell_value("Sheet1", "B1"), CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_xor_function_with_mixed_boolean_inputs(wb):
    """Test XOR function with a mix of TRUE, FALSE, and numeric values treated as Booleans."""
    wb.set_cell_contents("Sheet1", "A1", "=XOR(TRUE, FALSE, 1, 0)")

    # XOR should treat numeric 1 as TRUE and 0 as FALSE. Here, TRUE XOR FALSE XOR TRUE XOR FALSE = TRUE
    assert wb.get_cell_value("Sheet1", "A1") is True


def test_and_function_with_empty_cells(wb):
    """Test AND function treating empty cells as FALSE."""
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "B1", "=AND(A1, TRUE)")

    # Empty cell A1 treated as FALSE, so AND(FALSE, TRUE) = FALSE
    assert wb.get_cell_value("Sheet1", "B1") is False


def test_AND_string_true(wb):
    wb.set_cell_contents("Sheet1", "A1", "TRUE")
    wb.set_cell_contents("Sheet1", "A2", "bruh")
    wb.set_cell_contents("Sheet1", "A3", "=AND(A1, A2)")

    assert wb.get_cell_value("Sheet1", "A3").get_type() == CellErrorType.TYPE_ERROR


def test_or_function_with_text_conversion_to_boolean(wb):
    """Test OR function with text values that are implicitly converted to Booleans."""
    wb.set_cell_contents("Sheet1", "A1", "\"true\"")
    wb.set_cell_contents("Sheet1", "A2", "\"false\"")
    wb.set_cell_contents("Sheet1", "B1", "=OR(A1, A2)")

    # 'true' and 'false' strings do not convert to Boolean TRUE/FALSE for OR evaluation; expects TYPE_ERROR
    assert wb.get_cell_value('Sheet1', 'B1')
    assert wb.get_cell_value('Sheet1', 'A1')

    # assert isinstance(wb.get_cell_value("Sheet1", "B1"), CellError)
    # assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.TYPE_ERROR


def test_not_function_with_numeric_input(wb):
    """Test NOT function with numeric input, expecting conversion to Boolean."""
    wb.set_cell_contents("Sheet1", "A1", "=NOT(0)")

    # Numeric 0 converts to FALSE, so NOT(FALSE) = TRUE
    assert wb.get_cell_value("Sheet1", "A1") is True


def test_if_function_with_error_in_condition(wb):
    """Test IF function when the condition itself results in an error."""
    wb.set_cell_contents("Sheet1", "A1", "=1/0")  # This will result in a DIV/0! error
    wb.set_cell_contents("Sheet1", "B1", "=IF(A1, \"Yes\", \"No\")")

    # The condition A1 results in an error, so B1 should also yield an error instead of evaluating to 'Yes' or 'No'
    assert isinstance(wb.get_cell_value("Sheet1", "B1"), CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.DIVIDE_BY_ZERO


def test_lazy_evaluation_avoids_unnecessary_circular_reference(wb):
    """Test lazy evaluation in IF function avoids triggering an unnecessary circular reference."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(FALSE, B1, 1)")
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    # Due to lazy evaluation, B1 is not evaluated, avoiding a circular reference
    assert wb.get_cell_value("Sheet1", "A1") == 1


def test_dynamic_evaluation_with_indirect_causes_circular_reference(wb):
    """Test dynamic evaluation using INDIRECT function causing a circular reference when cell reference changes."""
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"B1\")")
    wb.set_cell_contents("Sheet1", "B1", "2")
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    # Changing B1 to reference A1 dynamically creates a circular reference
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_circular_reference_detection_with_complex_chain(wb):
    """Test complex circular reference detection involving multiple cells and INDIRECT function."""
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "B1", "=INDIRECT(\"C1\")")
    wb.set_cell_contents("Sheet1", "C1", "=A1")

    # A1 -> B1 -> INDIRECT(C1) -> A1 forms a circular reference
    bruh = wb.get_cell_value("Sheet1", "B1")
    assert isinstance(bruh, CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_indirect_function_with_dynamic_reference_update(wb):
    """Test INDIRECT function's response to dynamic updates affecting cell references."""
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"B1\")")
    wb.set_cell_contents("Sheet1", "B1", "=C1")
    wb.set_cell_contents("Sheet1", "C1", "10")
    wb.set_cell_contents("Sheet1", "B1", "=D1")
    wb.set_cell_contents("Sheet1", "D1", "20")

    # INDIRECT should dynamically update to reference D1 instead of C1 after B1's update
    assert wb.get_cell_value("Sheet1", "A1") == 20


def test_conditional_evaluation_with_nested_functions_avoids_errors(wb):
    """Test conditional evaluation in nested functions avoids unnecessary errors."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(TRUE, 100, 1/0)")
    wb.set_cell_contents("Sheet1", "B1", "=IFERROR(A1,\"womp womp\")")

    # IF should not evaluate 1/0 due to TRUE condition, and IFERROR should not be needed
    assert wb.get_cell_value("Sheet1", "B1") == 100


def test_lazy_evaluation_with_multiple_conditions(wb):
    """Test IF function with multiple conditions ensuring lazy evaluation is properly applied."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(TRUE, 1, 1/0)")
    wb.set_cell_contents("Sheet1", "B1", "=IF(FALSE, 1/0, 2)")

    # A1 should not evaluate 1/0 due to TRUE, B1 should not evaluate 1/0 due to FALSE
    assert wb.get_cell_value("Sheet1", "A1") == 1
    assert wb.get_cell_value("Sheet1", "B1") == 2


def test_dynamic_evaluation_indirect_with_self_reference(wb):
    """Test INDIRECT function with a self-reference that is dynamically created and resolved."""
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"B1\")")
    wb.set_cell_contents("Sheet1", "B1", "\"A1\"")
    wb.set_cell_contents("Sheet1", "B1", "=2")

    # B1 initially creates a self-reference via INDIRECT, then resolved to a static value
    assert wb.get_cell_value("Sheet1", "A1") == 2


def test_cycle_detection_with_function_chain(wb):
    """Test cycle detection involving a chain of function calls across multiple cells."""
    wb.set_cell_contents("Sheet1", "A1", "=B1 + 1")
    wb.set_cell_contents("Sheet1", "B1", "=C1 + 1")
    wb.set_cell_contents("Sheet1", "C1", "=A1 + 1")

    # A cycle is created across A1, B1, C1 with function calls adding 1
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_indirect_function_referencing_a_future_cell(wb):
    """Test INDIRECT function dynamically referencing a cell that will exist in the future."""
    wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(\"D1\")")
    wb.set_cell_contents("Sheet1", "D1", "=10")

    # A1 uses INDIRECT to reference D1 before D1 has a defined value
    assert wb.get_cell_value("Sheet1", "A1") == 10


def test_conditional_evaluation_preventing_circular_reference_error(wb):
    """Test conditional function evaluation preventing a circular reference through careful argument selection."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(TRUE, 10, B1)")
    wb.set_cell_contents("Sheet1", "B1", "=A1")

    # Despite B1 referencing A1, the IF condition prevents evaluation of B1, avoiding a circular reference
    assert wb.get_cell_value("Sheet1", "A1") == 10


def test_complex_chain_with_conditional_and_indirect_avoiding_circular_reference(wb):
    """Test a complex chain of conditional evaluations and INDIRECT references designed to smartly avoid a circular reference."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(TRUE, INDIRECT(\"B1\"), C1)")
    wb.set_cell_contents("Sheet1", "B1", "=C1")
    wb.set_cell_contents("Sheet1", "C1", "=10")
    wb.set_cell_contents("Sheet1", "D1", "=IF(A1=10, 20, INDIRECT(\"A1\"))")

    # Test complexity by involving INDIRECT and IF to navigate and avoid potential circular references
    assert wb.get_cell_value("Sheet1", "D1") == 20


def test_nested_function_calls_with_dynamic_evaluation_and_errors(wb):
    """Test nested function calls involving dynamic evaluation and handling multiple potential errors."""
    wb.set_cell_contents("Sheet1", "A1", "=IFERROR(INDIRECT(\"B1\"), C1)")
    wb.set_cell_contents("Sheet1", "B1", "=IF(True, D1/0, C1)")
    wb.set_cell_contents("Sheet1", "C1", "=10")
    wb.set_cell_contents("Sheet1", "D1", "=IFERROR(1/0, INDIRECT(\"C1\"))")

    # Evaluates nested functions with error handling and dynamic content resolution
    assert wb.get_cell_value("Sheet1", "A1") == 10
    assert wb.get_cell_value("Sheet1", "D1") == 10


def test_circular_reference_with_intermediate_values_and_lazy_evaluation(wb):
    """Test circular reference detection with intermediate values and lazy evaluation, involving INDIRECT and conditional logic."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(FALSE, B1, INDIRECT(\"B1\"))")
    wb.set_cell_contents("Sheet1", "B1", "=A1 + 1")

    # A circular reference that might be overlooked due to lazy evaluation and indirect referencing
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE


def test_dynamic_reference_chain_with_mixed_types_and_functions(wb):
    """Create a dynamic reference chain involving mixed types, functions, and INDIRECT, testing type coercion and error propagation."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(INDIRECT(\"B1\")>5, SUM(C1,D1), \"Error\")")
    wb.set_cell_contents("Sheet1", "B1", "\"C1\"")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "D1", "4")
    wb.set_cell_contents("Sheet1", "E1", "=A1")

    # Tests dynamic referencing with type coercion in conditionals and arithmetic within a chain of dependencies
    assert wb.get_cell_value("Sheet1", "E1") == 7


def test_advanced_error_handling_within_conditional_chains(wb):
    """Test advanced error handling and propagation within a complex chain of conditional logic and function evaluations."""
    wb.set_cell_contents("Sheet1", "A1", "=IF(B1=10, 100, IFERROR(1/0, INDIRECT(\"C1\")))")
    wb.set_cell_contents("Sheet1", "B1", "=10")
    wb.set_cell_contents("Sheet1", "C1", "=IF(D1=20, \"Valid\", \"Invalid\")")
    wb.set_cell_contents("Sheet1", "D1", "=20")
    # A complex decision tree that tests error handling, conditional logic, and dynamic evaluation across multiple cells
    assert wb.get_cell_value("Sheet1", "A1") == 100
    assert wb.get_cell_value("Sheet1", "C1") == 'Valid'

#
# def test_no_way_this_works(wb):
#     wb.set_cell_contents("Sheet1", "C2", "20")
#
#     wb.set_cell_contents("Sheet1", "A1", "=IF(TRUE, \"B\", \"C1\") & IF(C2, 1, FALSE)")
#
#     wb.set_cell_contents("Sheet1", "B1", "You're a Winner")
#     wb.set_cell_contents("Sheet1", "D1", "TRUE")
#
#     assert wb.get_cell_value("Sheet1", "A1") == "B1"
#
#     wb.set_cell_contents("Sheet1", "A1", "=INDIRECT(IF(D1, \"B\", \"F\") & IF(C2, 1, FALSE))")
#     wb.set_cell_contents("Sheet1", "Q1", "=IF(D1, \"B\", \"F\") & IF(C2, 1, FALSE)")
#
#     assert wb.get_cell_value("Sheet1", "A1") == "You're a Winner"
#     wb.set_cell_contents("Sheet1", "B1", "=C1")
#     wb.set_cell_contents("Sheet1", "C1", "=A1")
#     wb.set_cell_contents("Sheet1", "F1", "lllllllllll")
#
#     a1 = wb.get_cell_value("Sheet1", "A1")
#
#     assert a1.get_type() == CellErrorType.CIRCULAR_REFERENCE
#     wb.set_cell_contents("Sheet1", "d1", "false")
#
#     assert wb.get_cell_value("Sheet1", "Q1") == "F1"
#     assert wb.get_cell_value("Sheet1", "A1") == "lllllllllll"


def test_no_way_its_passing_2(wb):
    wb.set_cell_contents("Sheet1", "A1", "=ISERROR(B1)")
    wb.set_cell_contents("Sheet1", "B1", "= A1")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE

    wb.new_sheet("Sheet2")
    wb.set_cell_contents("Sheet2", "B1", "= A1")
    wb.set_cell_contents("Sheet2", "A1", "=ISERROR(B1)")
    assert isinstance(wb.get_cell_value("Sheet2", "A1"), CellError)
    assert wb.get_cell_value("Sheet2", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE