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

# No change but we're passing this
# 1. Test that False converts to "FALSE" when used in a context where a string is needed.
def test_boolean_false_to_string_conversion(wb):
    wb.set_cell_contents("Sheet1", "A1", "False")
    wb.set_cell_contents("Sheet1", "A2", "=A1 & \"_bruh\"")
    assert wb.get_cell_value("Sheet1", "A2") == "FALSE_bruh"
    wb.set_cell_contents("Sheet1", "A2", "=\"bruh_\" & A1")
    assert wb.get_cell_value("Sheet1", "A2") == "bruh_FALSE"



# PASSING WITHOUT CHANGES TO CODE
# 2. Test that True converts to "TRUE" when used in a context where a string is needed.
def test_boolean_true_to_string_conversion(wb):
    wb.set_cell_contents("Sheet1", "A1", "TRue")
    wb.set_cell_contents("Sheet1", "A2", "=A1 & \"_bruh\"")

    assert wb.get_cell_value("Sheet1", "A2") == "TRUE_bruh"

# PASSING WITHOUT CHANGES TO CODE
# 3. Test that conditional functions update propagated errors upon changing cell contents.
def test_conditional_functions_update_errors(wb):
    wb.set_cell_contents("Sheet1","B1", "TRUE")
    wb.set_cell_contents("Sheet1", "A1", "=If(B1, A2, A3)")
    wb.set_cell_contents("Sheet1", "A2", "#div/0!")  # Should produce a DIV/0 error
    wb.set_cell_contents("Sheet1", "A3", "=\"bruh\"")  # Should produce a DIV/0 error

    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    wb.set_cell_contents("Sheet1","B1", "False")

    assert wb.get_cell_value("Sheet1", "A1") == "bruh"

# PASSING WITHOUT CHANGES TO CODE
# 3. Test that conditional functions update propagated errors upon changing cell contents.
def test_what_what(wb):
    wb.set_cell_contents("Sheet1","B1", "false")
    wb.set_cell_contents("Sheet1", "A1", "=if(b1, A2, A3)")
    wb.set_cell_contents("Sheet1", "A2", "#REF!")  # Should produce a DIV/0 error
    wb.set_cell_contents("Sheet1", "A3", "=\"bruh\"")  # Should produce a DIV/0 error

    assert wb.get_cell_value("Sheet1", "A1") == "bruh"
    wb.set_cell_contents("Sheet1","B1", "#div/0!")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"),CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    wb.set_cell_contents("Sheet1","B1", "True")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"),CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE

# FAILS BEFORE CHANGES. 0 Index doesn't give an error.
# 4. CHOOSE() function should report an error if first argument isn't a valid integer index.
def test_choose_function_invalid_index(wb):
    wb.set_cell_contents("Sheet1", "A1", '=CHooSE("one", 1, 2, 3)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR
    wb.set_cell_contents("Sheet1", "A1", '=CHooSE(-1, 1, 2, 3)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR
    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(0, 1, 2, 3)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(9999, 1, 2, 3)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

# PASSES WITHOUT CHANGES
# 5. CHOOSE() function error propagation - only the chosen option should propagate.
def test_choose_function_error_propagation(wb):
    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(2, 1, 1/0, 3)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(3, 1, 1/0, 3)')
    assert wb.get_cell_value("Sheet1", "A1") == 3

    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(2, 1, #div/0!, 3)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(3, 1, #div/0!, 3)')
    assert wb.get_cell_value("Sheet1", "A1") == 3


# PASSES WITHOUT CHANGES
# 5. CHOOSE() function error propagation - only the chosen option should propagate.
def test_choose_function_error_propagation_lol(wb):

    wb.set_cell_contents("Sheet1", "A2", '=\"lol"')

    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(2, 1, A2, 3)')
    # assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1") == 'lol'
    
    wb.set_cell_contents("Sheet1", "A2", '#div/0!')

    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(3, A2, 1, 3)')
    assert wb.get_cell_value("Sheet1", "A1") == 3

    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(2, 1, A2, 3)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO

    wb.set_cell_contents("Sheet1", "A1", '=CHOOSE(3, 1, A2, 3)')
    assert wb.get_cell_value("Sheet1", "A1") == 3

# UNCLEAR WHAT THIS TEST MEANS
# 6. IF() function behavior with various kinds of bad inputs.
def test_if_function_bad_inputs(wb):
    wb.set_cell_contents("Sheet1", "A1", '=IF(A2, "good", "bad")')
    wb.set_cell_contents("Sheet1", "A2", "nonsense")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

# PASSING WITHOUT CHANGES
# 7. IF() function error propagation - errors should only propagate out of the arguments that are evaluated.
def test_if_function_error_propagation(wb):
    wb.set_cell_contents("Sheet1", "A2", 'TRuE')
    wb.set_cell_contents("Sheet1", "A1", '=IF(A2, #div/0!, 2)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    wb.set_cell_contents("Sheet1", "A2", 'FalSE')
    assert wb.get_cell_value("Sheet1", "A1") == 2

# PASSING WITH NO CHANGE.
# 8. Basic functionality of the IFERROR() function.
def test_iferror_function_basic(wb):
    wb.set_cell_contents("Sheet1", "A2", '0')
    wb.set_cell_contents("Sheet1", "A1", '=IFERROR(A2, "ERROR")')
    assert wb.get_cell_value("Sheet1", "A1") == 0
    wb.set_cell_contents("Sheet1", "A2", '=1/0')
    assert wb.get_cell_value("Sheet1", "A1") == "ERROR"


# UNCLEAR WHAT THIS TEST IS ACTING OR.
# 9. IFERROR() function behavior with various kinds of bad inputs.
def test_iferror_function_bad_inputs(wb):
    wb.set_cell_contents("Sheet1", "A1", '=IFERROR("bad")')
    wb.get_cell_value("Sheet1", "A1") == "bad"

# PASSING WITH NO CHANGES
# 10. IFERROR() function behavior when inside and outside of a cycle - cycle through both arguments.
def test_iferror_function_in_cycle(wb):
    wb.set_cell_contents("Sheet1", "A1", '=IFERROR(A2, A3)')
    wb.set_cell_contents("Sheet1", "A2", '=A1')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    wb.set_cell_contents("Sheet1", "A1", '=A2')
    wb.set_cell_contents("Sheet1", "A2", '=A1')
    wb.set_cell_contents("Sheet1", "A3", '=IFERROR(A2, "BRUH")')

    assert wb.get_cell_value("Sheet1", "A3") == "BRUH"

# UNCLEAR WHAT THIS QUESTION IS TRYING TO AASK
# 11. IFERROR() function should map empty-cell arguments to 0 value.
# def test_iferror_function_empty_cell_to_zero(wb):
#     wb.set_cell_contents("Sheet1", "A1", '=IFERROR(A2, 0)')
#     assert wb.get_cell_value("Sheet1", "A1") == 0

# PASSING WITHOUT MODIFICATION
# 12. IFERROR() should not create cycles when first argument is not an error, and therefore second argument should not be evaluated.
def test_iferror_function_no_cycle_on_valid_input(wb):
    wb.set_cell_contents("Sheet1", "A1", '=IFERROR(1, A2)')
    wb.set_cell_contents("Sheet1", "A2", '=A1')

    assert wb.get_cell_value("Sheet1", "A1") == 1

# PASSING WITHOUT MODIFICATION
# 13. IFERROR() function behavior when referencing things that aren't errors (should detect no errors).
def test_iferror_function_no_error_detection(wb):
    wb.set_cell_contents("Sheet1", "A2", '5')
    wb.set_cell_contents("Sheet1", "A1", '=IFerROR(A2, "Error")')
    assert wb.get_cell_value("Sheet1", "A1") == 5

# PASSING WITH NO  MODIFICATION
# 14. Test that function calls in formulas can include a variety of expressions and different function types.
def test_function_calls_with_varied_expressions(wb):
    # Test with arithmetic expressions
    wb.set_cell_contents("Sheet1", "A1", '=SUM(1+1, 2+2)')
    assert wb.get_cell_value("Sheet1", "A1") == 6

    # Test with comparison and logical operators
    wb.set_cell_contents("Sheet1", "A2", '=AND(4 + 7 > 6, 3*12 - 1 < 1000)')
    assert wb.get_cell_value("Sheet1", "A2") == True

    # Test with nested functions and indirect references
    wb.set_cell_contents("Sheet1", "A3", "10")
    wb.set_cell_contents("Sheet1", "A4", '=IF(A3 > 5, SUM(INDIRECT("A3") + 2), 0)')
    assert wb.get_cell_value("Sheet1", "A4") == 12

    # Test with conditional expressions and error handling
    wb.set_cell_contents("Sheet1", "A5", '=IFERROR(1/0, "ErrorHandled")')
    assert wb.get_cell_value("Sheet1", "A5") == "ErrorHandled"

    # Test with string concatenation and comparison
    wb.set_cell_contents("Sheet1", "A6", '=EXACT("test" & "ing", "testing")')
    assert wb.get_cell_value("Sheet1", "A6") == True

    # Test using a boolean function with a string to boolean conversion
    wb.set_cell_contents("Sheet1", "A7", '=OR(4 < 2, 4 > 2)')
    assert wb.get_cell_value("Sheet1", "A7") == True

    # Test using XOR with boolean values directly
    wb.set_cell_contents("Sheet1", "A8", '=XOR(3 + 3 > 6, 2 + 2 < 5)')
    assert wb.get_cell_value("Sheet1", "A8") == True



# PASSING WITHOUT MODICATION
# 15. Test that function calls in formulas can include other function calls as part of expressions in their arguments.
def test_nested_function_calls_of_various_types_in_expressions(wb):
    # Using a combination of SUM, IF, AND, and INDIRECT functions in nested calls
    wb.set_cell_contents("Sheet1", "A2", "4")
    wb.set_cell_contents("Sheet1", "A3", "TRUE")
    wb.set_cell_contents("Sheet1", "A4", "=A2")
    wb.set_cell_contents("Sheet1", "A5",'=IF(AND(A3, TRUE), INdIRECT("A4"), 0)')
    wb.set_cell_contents("Sheet1", "A1", '=SUM(SUM(1,1), IF(AND(A3, TRUE), INDIRECT("A4"), 0), SUM(2,2))')
    # The expected value breakdown:
    # SUM(1,1) evaluates to 2
    # IF(AND(TRUE, TRUE), INDIRECT("A4"), 0) evaluates to the value in A2, which is 4
    # SUM(2,2) evaluates to 4
    # The overall SUM then adds up 2 + 4 + 4, which should be 10
    assert wb.get_cell_value("Sheet1", "A1") == 10


# 16. Test that functions can be used in more complex expressions involving a variety of function types and logical operations.
def test_functions_in_complex_expressions_with_variety(wb):
    # Combine arithmetic operations with SUM and INDIRECT functions
    wb.set_cell_contents("Sheet1", "B1", "5")
    wb.set_cell_contents("Sheet1", "A1", '=SUM(1, INDIRECT("B1")) * 2')
    assert wb.get_cell_value("Sheet1", "A1") == 12  # (1 + 5) * 2 = 12

    # Use IF function within arithmetic expressions
    wb.set_cell_contents("Sheet1", "A2", '=IF(SUM(B1, 5) > 10, 3, 4) * 10')
    assert wb.get_cell_value("Sheet1", "A2") == 40  # (5 + 5) > 10 ? 3 : 4 * 10 = 30

    # Incorporate logical operations with AND and OR functions
    wb.set_cell_contents("Sheet1", "A3", '=AND(2 > 1, 3 < 4)')
    wb.set_cell_contents("Sheet1", "A4", '=OR(A3, FALSE) * 100')
    assert wb.get_cell_value("Sheet1", "A4") == 100  # AND(2 > 1, 3 < 4) = TRUE, OR(TRUE, FALSE) = TRUE * 100 = 100

    # Combining IFERROR with arithmetic operations
    wb.set_cell_contents("Sheet1", "A5", '=IFERROR(1/0, 5) + 5')
    assert wb.get_cell_value("Sheet1", "A5") == 10  # IFERROR(divide by zero, 5) + 5 = 10

    # Nested function calls involving string and comparison operations
    wb.set_cell_contents("Sheet1", "A6", '=NOT(AND(EXACT("a", "a") , (1 + 1 == 2)))')
    assert wb.get_cell_value("Sheet1", "A6") == False  # EXACT("a", "a") AND (1 + 1 = 2) * 10 = TRUE AND TRUE * 10 = 10

    # Complex formula involving multiple function types and a comparison operation
    wb.set_cell_contents("Sheet1", "A7", '=SUM(IF(TRUE, 4, 5), XOR(FALSE, TRUE)) * IFERROR(SUM(1,2), 0)')
    assert wb.get_cell_value("Sheet1", "A7") == 15

# 17. INDIRECT() function should propagate errors in input.
def test_indirect_function_error_propagation(wb):
    wb.set_cell_contents("Sheet1", "A1", '=INDIRECT("A2")')
    wb.set_cell_contents("Sheet1", "A2", '=1/0')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO


# 18. Basic functionality of the EXACT() function.
def test_exact_function_basic(wb):
    wb.set_cell_contents("Sheet1", "A1", '=EXACT("text", "text")')
    assert wb.get_cell_value("Sheet1", "A1") == True

# 19. EXACT() function error propagation.
def test_exact_function_error_propagation(wb):
    wb.set_cell_contents("Sheet1", "A1", '=EXACT(A2, "text")')
    wb.set_cell_contents("Sheet1", "A2", '=1/0')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO

# 20. ISBLANK() function behavior with various kinds of bad inputs.
def test_isblank_function_bad_inputs(wb):
    wb.set_cell_contents("Sheet1", "A1", '=ISBLANK(A2)')
    assert wb.get_cell_value("Sheet1", "A1") == True

# TODO: This originally failed beacause of a bug with the Version function
#  not actually being in the dictionary. Now it fails for a different more complicated reason.
# 21. Basic functionality of the VERSION() function.
def test_version_function_basic(wb):
    wb.set_cell_contents("Sheet1", "A1", '=VERSION()')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), str)

# TODO: Passes after change.
# 22. VERSION() function with bad inputs.
def test_version_function_bad_inputs(wb):
    wb.set_cell_contents("Sheet1", "A1", '=VERSION(A2)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR


# Unsure what this means?
# 23. AND() function behavior with various kinds of bad inputs.
def test_and_function_bad_inputs(wb):
    wb.set_cell_contents("Sheet1", "A1", '=AND(A2, "True")')
    assert wb.get_cell_value("Sheet1", "A1") == False

# Originally just gave a type error.
# 24. AND() function error propagation, and should not short-circuit.
def test_and_function_error_propagation_no_short_circuit(wb):
    wb.set_cell_contents("Sheet1", "A1", '=AND(TRUE, 1/0)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO

#  Originally also just a type error. Is this not what donnie wants?
# 25. NOT() function error propagation.
def test_not_function_error_propagation(wb):
    wb.set_cell_contents("Sheet1", "A1", '=NOT(1/0)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO

# 26. OR() function error propagation, and should not short-circuit.
def test_or_function_error_propagation_no_short_circuit(wb):
    wb.set_cell_contents("Sheet1", "A1", '=OR(FALSE, 1/0)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO

# 27. Basic functionality of the XOR() function.
def test_xor_function_basic(wb):
    wb.set_cell_contents("Sheet1", "A1", '=XOR(TRUE, FALSE)')
    assert wb.get_cell_value("Sheet1", "A1") == True

# 28. XOR() function error propagation.
def test_xor_function_error_propagation(wb):
    wb.set_cell_contents("Sheet1", "A1", '=XOR(TRUE, 1/0)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.TYPE_ERROR

# 29. Renaming a sheet should update cell-references in function arguments.
def test_rename_sheet_updates_function_references(wb):
    wb.set_cell_contents("Sheet1", "A1", '=SUM(Sheet1!A2)')
    wb.new_sheet("Sheet2")
    wb.rename_sheet("Sheet1", "RenamedSheet")
    assert wb.get_cell_contents("RenamedSheet", "A1") == '=SUM(RenamedSheet!A2)'

def test_rename_sheet_updates_function_references_(wb):
    wb.set_cell_contents("Sheet1", "A1", '=SUM(Sheet1!A2)')
    wb.new_sheet("Sheet2")
    wb.rename_sheet("Sheet1", "RenamedSheet")
    assert wb.get_cell_contents("RenamedSheet", "A1") == '=SUM(RenamedSheet!A2)'

def test_INDIRECT(wb):
    wb.set_cell_contents("Sheet1", "A1", '=INdiRECT(FAKEFUNCTION())')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_NAME
    wb.set_cell_contents("Sheet1", "A1", '=INDIRECT(1/0)')
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    wb.set_cell_contents("Sheet1", "A1", '=INDIRECT(\"NonExistant!A1\")')
    assert isinstance(wb.get_cell_value("Sheet1","A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE

def test_bruh(wb):
    wb.set_cell_contents("Sheet1", "A1", '="bruh"')
    assert wb.get_cell_value("Sheet1", "A1") == "bruh"

def test_version(wb):
    wb.set_cell_contents("Sheet1", "A1","=VERSION()")
    assert wb.get_cell_value("Sheet1", "A1") == '1.4'