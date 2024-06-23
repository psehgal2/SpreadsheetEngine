import context
import pytest
import sheets
from decimal import Decimal
from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType
from sheets.Workbook import Workbook

# Test 1: BAD_REFERENCE error resolved by adding missing sheet
def test_bad_reference_resolved():
    # Corresponds to acceptance test 1
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=sheet2!A1')
    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

    workbook.new_sheet('Sheet2')
    # After adding Sheet2, A1 should no longer have a BAD_REFERENCE error
    assert not isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)

# Test 2: BAD_REFERENCE error due to deletion of a sheet
def test_bad_reference_due_to_sheet_deletion():
    # Corresponds to acceptance test 3
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.new_sheet('Sheet2')
    workbook.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1')

    workbook.del_sheet('Sheet2')
    # After deleting Sheet2, A1 should have a BAD_REFERENCE error
    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 3: Propagation of BAD_REFERENCE error
def test_bad_reference_propagation():
    # Corresponds to acceptance test 4
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=NonExistent!A1')
    workbook.set_cell_contents('Sheet1', 'B1', '=-A1')

    # Both A1 and B1 should have a BAD_REFERENCE error
    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE
    assert isinstance(workbook.get_cell_value('Sheet1', 'B1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'B1').get_type() == CellErrorType.BAD_REFERENCE

# Test 4: Propagation of CIRCULAR_REFERENCE error
def test_circular_reference_propagation():
    # Corresponds to acceptance test 5
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=B1')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1')
    workbook.set_cell_contents('Sheet1', 'C1', '=-A1')

    # A1 and B1 should have a CIRCULAR_REFERENCE error, and so should C1
    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert isinstance(workbook.get_cell_value('Sheet1', 'B1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'B1').get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert isinstance(workbook.get_cell_value('Sheet1', 'C1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'C1').get_type() == CellErrorType.CIRCULAR_REFERENCE
# Test 5: Handling multiple loops
def test_multiple_loops():
    # Corresponds to acceptance test 6
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=B1')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1')
    workbook.set_cell_contents('Sheet1', 'C1', '=D1')
    workbook.set_cell_contents('Sheet1', 'D1', '=C1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert isinstance(workbook.get_cell_value('Sheet1', 'C1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'C1').get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 6: Cycles spanning multiple sheets
def test_cycles_across_sheets():
    # Corresponds to acceptance test 7
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.new_sheet('Sheet2')
    workbook.set_cell_contents('Sheet1', 'A1', '=Sheet2!B1')
    workbook.set_cell_contents('Sheet2', 'B1', '=Sheet1!A1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 7: Cells not in cycles
def test_cells_not_in_cycles():
    # Corresponds to acceptance test 8
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=B1')
    workbook.set_cell_contents('Sheet1', 'B1', '5')
    workbook.set_cell_contents('Sheet1', 'C1', '=A1')

    assert workbook.get_cell_value('Sheet1', 'C1') == Decimal('5')

# Test 8: Cell pointing to a cell in a cycle
def test_cell_pointing_to_cycle():
    # Corresponds to acceptance test 9
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=B1')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1')
    workbook.set_cell_contents('Sheet1', 'C1', '=A1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'C1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'C1').get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 9: Priority of circular reference error in multiple errors
def test_circular_reference_priority_in_multiple_errors():
    # Corresponds to acceptance test 10
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=B1/0') # Divide by zero error
    workbook.set_cell_contents('Sheet1', 'B1', '=A1') # Circular reference

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 10: Single loop detection
def test_single_loop_detection():
    # Corresponds to acceptance test 11
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=A1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 11: Cell directly referring to itself
def test_cell_referring_to_itself():
    # Corresponds to acceptance test 12
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=A1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 12: Propagation of DIVIDE_BY_ZERO error
def test_divide_by_zero_propagation():
    # Corresponds to acceptance test 13
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=1/0')
    workbook.set_cell_contents('Sheet1', 'B1', '=-A1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.DIVIDE_BY_ZERO
    assert isinstance(workbook.get_cell_value('Sheet1', 'B1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'B1').get_type() == CellErrorType.DIVIDE_BY_ZERO

# Test 13: Formulas with addition involving error literals
def test_addition_with_error_literals():
    # Corresponds to acceptance test 14
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=4 + #REF!')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 14: Formulas with division involving error literals
def test_division_with_error_literals():
    # Corresponds to acceptance test 15
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=4 / #REF!')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 15: Formulas with multiplication involving error literals
def test_multiplication_with_error_literals():
    # Corresponds to acceptance test 16
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=4 * #REF!')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 16: Formulas with parentheses involving error literals
def test_parentheses_with_error_literals():
    # Corresponds to acceptance test 17
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=(#REF!)')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 17: Simple formulas with error literals
def test_simple_formula_with_error_literals():
    # Corresponds to acceptance test 18
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=#REF!')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 18: Formulas with subtraction involving error literals
def test_subtraction_with_error_literals():
    # Corresponds to acceptance test 19
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=4 - #REF!')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 19: Formulas with unary operators involving error literals
def test_unary_operator_with_error_literals():
    # Corresponds to acceptance test 20
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=-#REF!')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.BAD_REFERENCE

# Test 20: Formula adding cell and number with whitespace
def test_add_cell_and_number_with_whitespace():
    # Corresponds to acceptance test 21
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 123 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= A1 + 456 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('579')

# Test 21: Formula adding number and cell with whitespace
def test_add_number_and_cell_with_whitespace():
    # Corresponds to acceptance test 22
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 123 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= 456 + A1 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('579')

# Test 22: Concatenating number and string removing trailing zeros
def test_concatenate_number_and_string_removing_trailing_zeros():
    # Corresponds to acceptance test 23
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '123.00')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1 & " apples"')

    assert workbook.get_cell_value('Sheet1', 'B1') == '123 apples'

# Test 23: Formula dividing cell and number with whitespace
def test_divide_cell_and_number_with_whitespace():
    # Corresponds to acceptance test 24
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 246 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= A1 / 2 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('123')

# Test 24: Formula dividing number and cell with whitespace
def test_divide_number_and_cell_with_whitespace():
    # Corresponds to acceptance test 25
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 2 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= 246 / A1 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('123')

# Test 25: Formulas with multiple operators and only cells
def test_multiple_operators_with_cells():
    # Corresponds to acceptance test 27
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '3')
    workbook.set_cell_contents('Sheet1', 'B1', '4')
    workbook.set_cell_contents('Sheet1', 'C1', '=A1 + B1 * A1 - B1')

    assert workbook.get_cell_value('Sheet1', 'C1') == Decimal('11')

# Test 26: Formulas with multiple operators, cells, and numbers
def test_multiple_operators_with_cells_and_numbers():
    # Corresponds to acceptance test 28
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '3')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1 + 2 * 5 - 3')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('10')

# Test 27: Formulas with multiple operators and only numbers
def test_multiple_operators_with_only_numbers():
    # Corresponds to acceptance test 29
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=3 + 4 * 5 - 2')

    assert workbook.get_cell_value('Sheet1', 'A1') == Decimal('21')

# Test 28: Formula multiplying cell and number with whitespace
def test_multiply_cell_and_number_with_whitespace():
    # Corresponds to acceptance test 30
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 3 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= A1 * 4 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('12')

# Test 29: Formula multiplying number and cell with whitespace
def test_multiply_number_and_cell_with_whitespace():
    # Corresponds to acceptance test 31
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 4 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= 3 * A1 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('12')

# Test 30: Complex formula spanning multiple sheets with multiple operators
def test_complex_formula_across_sheets():
    # Corresponds to acceptance test 32
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.new_sheet('Sheet2')
    workbook.set_cell_contents('Sheet1', 'A1', '2')
    workbook.set_cell_contents('Sheet2', 'B1', '3')
    workbook.set_cell_contents('Sheet1', 'C1', '=Sheet2!B1 * 2 + A1')

    assert workbook.get_cell_value('Sheet1', 'C1') == Decimal('8')

# Test 31: Formula subtracting cell and number with whitespace
def test_subtract_cell_and_number_with_whitespace():
    # Corresponds to acceptance test 33
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 5 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= A1 - 2 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('3')

# Test 32: Formula subtracting number and cell with whitespace
def test_subtract_number_and_cell_with_whitespace():
    # Corresponds to acceptance test 34
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', ' 2 ')
    workbook.set_cell_contents('Sheet1', 'B1', '= 5 - A1 ')

    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('3')

# Test 33: Propagation of PARSE_ERROR
def test_parse_error_propagation():
    # Corresponds to acceptance test 35
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '=+')
    workbook.set_cell_contents('Sheet1', 'B1', '=-A1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'A1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.PARSE_ERROR
    assert isinstance(workbook.get_cell_value('Sheet1', 'B1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'B1').get_type() == CellErrorType.PARSE_ERROR

# Test 34: Propagation of TYPE_ERROR
def test_type_error_propagation():
    # Corresponds to acceptance test 36
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '="abc"')
    workbook.set_cell_contents('Sheet1', 'B1', '=-A1')

    assert isinstance(workbook.get_cell_value('Sheet1', 'B1'), CellError)
    assert workbook.get_cell_value('Sheet1', 'B1').get_type() == CellErrorType.TYPE_ERROR

# Test 35: String operations resulting in TYPE_ERROR
def test_string_operations_type_error():
    # Corresponds to acceptance test 37
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    tests = ['="abc" + 123', '="abc" * 123', '="abc" - 123', '="abc" / 123']
    for i, test in enumerate(tests, start=1):
        cell = 'A' + str(i)
        workbook.set_cell_contents('Sheet1', cell, test)
        assert isinstance(workbook.get_cell_value('Sheet1', cell), CellError)
        assert workbook.get_cell_value('Sheet1', cell).get_type() == CellErrorType.TYPE_ERROR

# Test 36: String-to-string operations resulting in TYPE_ERROR
def test_string_to_string_operations_type_error():
    # Corresponds to acceptance test 38
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    tests = ['="abc" + "def"', '="abc" * "def"', '="abc" - "def"', '="abc" / "def"']
    for i, test in enumerate(tests, start=1):
        cell = 'A' + str(i)
        workbook.set_cell_contents('Sheet1', cell, test)
        assert isinstance(workbook.get_cell_value('Sheet1', cell), CellError)
        assert workbook.get_cell_value('Sheet1', cell).get_type() == CellErrorType.TYPE_ERROR

# Test 37: Setting cell to error-value string
def test_setting_cell_to_error_value_string():
    # Corresponds to acceptance test 39
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    error_values = ['#ERROR!', '#CIRCREF!', '#REF!', '#NAME?', '#VALUE!', '#DIV/0!']
    for i, error in enumerate(error_values, start=1):
        cell = 'A' + str(i)
        workbook.set_cell_contents('Sheet1', cell, error)
        assert isinstance(workbook.get_cell_value('Sheet1', cell), CellError)

# Test 38: Preserving whitespace in quoted-string values
def test_preserving_whitespace_in_quoted_strings():
    # Corresponds to acceptance test 40
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', "'   hello")
    assert workbook.get_cell_value('Sheet1', 'A1') == '   hello'

# Test 39: "Diamond" dependency pattern
def test_diamond_dependency_pattern():
    # Corresponds to acceptance test 41
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '5')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1 * 2')
    workbook.set_cell_contents('Sheet1', 'C1', '=A1 * 3')
    workbook.set_cell_contents('Sheet1', 'D1', '=B1 + C1')

    workbook.set_cell_contents('Sheet1', 'A1', '6')
    assert workbook.get_cell_value('Sheet1', 'D1') == Decimal('30')

# Test 40: Direct dependency update
def test_direct_dependency_update():
    # Corresponds to acceptance test 42
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '5')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1 * 2')

    workbook.set_cell_contents('Sheet1', 'A1', '6')
    assert workbook.get_cell_value('Sheet1', 'B1') == Decimal('12')

# Test 41: Indirect dependency update
def test_indirect_dependency_update():
    # Corresponds to acceptance test 43
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.set_cell_contents('Sheet1', 'A1', '5')
    workbook.set_cell_contents('Sheet1', 'B1', '=A1 * 2')
    workbook.set_cell_contents('Sheet1', 'C1', '=B1 + 3')

    workbook.set_cell_contents('Sheet1', 'A1', '6')
    assert workbook.get_cell_value('Sheet1', 'C1') == Decimal('15')

# Test 42: Reflecting delete and new operations in Workbook.list_sheets()
def test_reflect_delete_and_new_operations_in_list_sheets():
    # Corresponds to acceptance test 44
    workbook = Workbook()
    workbook.new_sheet('Sheet1')
    workbook.new_sheet('Sheet2')
    workbook.del_sheet('Sheet1')
    workbook.new_sheet('Sheet3')

    assert set(workbook.list_sheets()) == {'Sheet2', 'Sheet3'}

# Test 43: Adding a mix of named and unnamed sheets to a workbook
def test_adding_named_and_unnamed_sheets():
    # Corresponds to acceptance test 45
    workbook = Workbook()
    for _ in range(3):
        workbook.new_sheet()
    workbook.new_sheet('NamedSheet')

    assert set(workbook.list_sheets()) == {'Sheet1', 'Sheet2', 'Sheet3', 'NamedSheet'}

# Test 44: Adding named and unnamed sheets with overlapping names
# Test 44: Attempting to add sheets with case-insensitive overlapping names should raise ValueError
def test_adding_sheets_with_case_insensitive_names_raises_error():
    # Corresponds to revised acceptance test 46
    workbook = Workbook()
    workbook.new_sheet('Sheet')

    with pytest.raises(ValueError):
        workbook.new_sheet('sheet')  # This should raise a ValueError

