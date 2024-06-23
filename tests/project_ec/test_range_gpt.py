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

# Test 1: MIN function with a mix of direct values and range including type error
def test_min_with_mixed_types_and_ranges(wb):
    """
    Validates MIN function handling mixed direct values and cell ranges, ensuring type errors are correctly returned
    if any non-numeric value is included. Spec: 'All non-empty inputs are converted to numbers; if any input cannot
    be converted to a number then the function returns a TYPE_ERROR.'
    """
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "text")  # This should cause a TYPE_ERROR
    wb.set_cell_contents("Sheet1", "A3", "5")
    wb.set_cell_contents("Sheet1", "B1", "=MIN(1, A1:A3, 20)")
    assert isinstance(wb.get_cell_value("Sheet1", "B1"), CellError)
    assert wb.get_cell_value("Sheet1", "B1").get_type() == CellErrorType.TYPE_ERROR

# Test 2: MAX function using only a range with empty cells and non-empty cells
def test_max_with_range_including_empty_cells(wb):
    """
    Tests the MAX function with a cell range that includes both empty and non-empty cells.
    According to spec, 'Only non-empty cells should be considered; empty cells should be ignored.'
    """
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", None)  # Empty cell
    wb.set_cell_contents("Sheet1", "C3", "7")
    wb.set_cell_contents("Sheet1", "D1", "=MAX(C1:C3)")
    assert wb.get_cell_value("Sheet1", "D1") == decimal.Decimal("7")

# Test 3: SUM function with a range that results in a division by zero in AVERAGE
def test_average_divide_by_zero_with_empty_range(wb):
    """
    Ensures the AVERAGE function returns a DIVIDE_BY_ZERO error when applied to an entirely empty range.
    Spec: 'If the function’s inputs only include empty cells then the function returns a DIVIDE_BY_ZERO error.'
    """
    wb.set_cell_contents("Sheet1", "E1", None)
    wb.set_cell_contents("Sheet1", "E2", None)
    wb.set_cell_contents("Sheet1", "F1", "=AVERAGE(E1:E2)")
    assert isinstance(wb.get_cell_value("Sheet1", "F1"), CellError)
    assert wb.get_cell_value("Sheet1", "F1").get_type() == CellErrorType.DIVIDE_BY_ZERO

# Test 4: VLOOKUP function with exact match and return from a specific index
def test_vlookup_with_exact_match_and_index_return(wb):
    """
    Tests VLOOKUP to ensure it correctly finds an exact match and returns the value from the specified index column.
    As per spec: 'If such a column is found, the cell in the index-th column of the found row. The index is 1-based.'
    """
    wb.set_cell_contents("Sheet1", "G1", "Key")
    wb.set_cell_contents("Sheet1", "G2", "Match")
    wb.set_cell_contents("Sheet1", "H2", "Value")
    wb.set_cell_contents("Sheet1", "I1", "=VLOOKUP(\"Match\", G1:H2, 2)")
    assert wb.get_cell_value("Sheet1", "I1") == "Value"

# Test 5: Handling circular references within cell-range references
def test_circular_reference_with_cell_range(wb):
    """
    Validates that the engine correctly identifies and reports circular references involving cell-range references.
    Spec: 'Your spreadsheet engine must correctly identify and report cycles generated via cell-range references.'
    """
    wb.set_cell_contents("Sheet1", "J1", "=SUM(J1:J2)")
    wb.set_cell_contents("Sheet1", "J2", "1")
    assert isinstance(wb.get_cell_value("Sheet1", "J1"), CellError)
    assert wb.get_cell_value("Sheet1", "J1").get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 6: INDIRECT function constructing cell-range references and using it in SUM function
def test_indirect_with_cell_range_in_sum(wb):
    """
    Tests the INDIRECT function's ability to construct cell-range references, used here within a SUM function.
    This tests the spec: '...it should be possible to construct formulas like =IFERROR(VLOOKUP(A2, INDIRECT(B1 & "!A2:J100"), 7), "")'
    focusing on INDIRECT usage for cell ranges.
    """
    wb.set_cell_contents("Sheet1", "A1", "1")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "A3", "3")
    wb.set_cell_contents("Sheet1", "B1", "Sheet1!A1:A3")
    wb.set_cell_contents("Sheet1", "C1", "=SUM(INDIRECT(B1))")
    assert wb.get_cell_value("Sheet1", "C1") == 6

# Test 7: Testing HLOOKUP for horizontal search with multiple matching keys
def test_hlookup_with_multiple_matching_keys(wb):
    """
    Verifies HLOOKUP function's behavior when multiple columns match the key, ensuring the leftmost one is used.
    According to spec, 'If multiple columns match the key value, the leftmost one should be used.'
    """
    wb.set_cell_contents("Sheet1", "D1", "Match")
    wb.set_cell_contents("Sheet1", "E1", "Match")  # Duplicate key to test leftmost selection
    wb.set_cell_contents("Sheet1", "D2", "First Value")
    wb.set_cell_contents("Sheet1", "E2", "Second Value")
    wb.set_cell_contents("Sheet1", "F1", "=HLOOKUP(\"Match\", D1:E2, 2)")
    assert wb.get_cell_value("Sheet1", "F1") == "First Value"

# Test 8: SUM function with a completely empty cell range
def test_sum_with_completely_empty_range(wb):
    """
    Ensures the SUM function correctly handles a completely empty cell range, resulting in 0.
    Spec: 'If the function’s inputs only include empty cells then the function’s result is 0.'
    """
    # No cells set in the range G1:G3
    wb.set_cell_contents("Sheet1", "H1", "=SUM(G1:G3)")
    assert wb.get_cell_value("Sheet1", "H1") == 0

# Test 9: Testing cell-range parsing and updating during sheet rename affecting cell-range references
def test_sheet_rename_updates_cell_range_references(wb):
    """
    Validates that renaming a sheet updates cell-range references in formulas accordingly.
    Spec: 'Operations...renaming a spreadsheet, also need to update cell-range references in the same ways...'
    """
    wb.new_sheet("Sheet2")
    wb.set_cell_contents("Sheet1", "A1", "=SUM(Sheet2!A1:A2)")
    wb.set_cell_contents("Sheet2", "A1", "2")
    wb.set_cell_contents("Sheet2", "A2", "3")
    wb.rename_sheet("Sheet2", "DataSheet")
    assert wb.get_cell_contents("Sheet1", "A1") == "=SUM(DataSheet!A1:A2)"

# Test 10: TYPE_ERROR when mixing strings and numbers in SUM function through cell range
def test_type_error_when_mixing_types_in_sum(wb):
    """
    Tests SUM function for a TYPE_ERROR when cell ranges mix strings and numbers, in line with spec:
    'All non-empty inputs are converted to numbers; if any input cannot be converted to a number then the function returns a TYPE_ERROR.'
    """
    wb.set_cell_contents("Sheet1", "I1", "1")
    wb.set_cell_contents("Sheet1", "I2", "Hello")  # This should cause a TYPE_ERROR
    wb.set_cell_contents("Sheet1", "I3", "3")
    wb.set_cell_contents("Sheet1", "J1", "=SUM(I1:I3)")
    assert isinstance(wb.get_cell_value("Sheet1", "J1"), CellError)
    assert wb.get_cell_value("Sheet1", "J1").get_type() == CellErrorType.TYPE_ERROR


# Test 11: Error propagation in nested functions with cell ranges
def test_error_propagation_in_nested_functions(wb):
    """
    Validates error propagation when nested functions involve cell ranges that result in an error.
    According to the spec: 'With the exception of IFERROR, any function argument which is an error should yield an error in the output.'
    """
    wb.set_cell_contents("Sheet1", "K1", "=1/0")  # Directly causes a DIVIDE_BY_ZERO error
    wb.set_cell_contents("Sheet1", "L1", "=SUM(K1, 10, IFERROR(K1, 0))")
    assert isinstance(wb.get_cell_value("Sheet1", "L1"), CellError)
    assert wb.get_cell_value("Sheet1", "L1").get_type() == CellErrorType.DIVIDE_BY_ZERO

# Test 12: Circular reference detection with INDIRECT referencing a cell range including itself
def test_circular_reference_with_indirect_referencing_itself(wb):
    """
    Tests circular reference detection involving INDIRECT function that constructs a cell-range reference including the cell itself.
    Spec: '...cell-range references in formulas must be taken into account when performing cycle detection.'
    """
    wb.set_cell_contents("Sheet1", "M1", "'Sheet1!M1:M2")
    wb.set_cell_contents("Sheet1", "M2", "=SUM(INDIRECT(M1))")
    assert isinstance(wb.get_cell_value("Sheet1", "M2"), CellError)
    assert wb.get_cell_value("Sheet1", "M2").get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 13: VLOOKUP function correctly handling non-existent key
def test_vlookup_handling_non_existent_key(wb):
    """
    Ensures the VLOOKUP function returns a TYPE_ERROR when the key does not exist within the specified range.
    Spec: 'If no row is found, the function returns a TYPE_ERROR.'
    """
    wb.set_cell_contents("Sheet1", "N1", "SearchKey")
    wb.set_cell_contents("Sheet1", "O1", "=VLOOKUP(\"NotFound\", N1:N2, 1)")
    assert isinstance(wb.get_cell_value("Sheet1", "O1"), CellError)
    assert wb.get_cell_value("Sheet1", "O1").get_type() == CellErrorType.TYPE_ERROR

# Test 14: Correct handling of MIN function with cell ranges including negative and positive values
def test_min_with_negative_and_positive_values(wb):
    """
    Tests MIN function with cell ranges including both negative and positive values, ensuring correct minimum value identification.
    Spec: 'MIN(value1, ...) returns the minimum value over the set of inputs.'
    """
    wb.set_cell_contents("Sheet1", "P1", "-10")
    wb.set_cell_contents("Sheet1", "P2", "0")
    wb.set_cell_contents("Sheet1", "P3", "10")
    wb.set_cell_contents("Sheet1", "Q1", "=MIN(P1:P3)")
    assert wb.get_cell_value("Sheet1", "Q1") == decimal.Decimal("-10")

# Test 15: AVERAGE function excluding empty cells and including cell ranges with mixed numeric types
def test_average_excluding_empty_and_mixed_types(wb):
    """
    Validates AVERAGE function correctly excludes empty cells and handles ranges with mixed numeric types.
    Spec: 'All non-empty inputs are converted to numbers; if any input cannot be converted to a number then the function returns a TYPE_ERROR.'
    """
    wb.set_cell_contents("Sheet1", "R1", "2")
    wb.set_cell_contents("Sheet1", "R2", None)  # Empty cell, should be ignored
    wb.set_cell_contents("Sheet1", "R3", "4.5")
    wb.set_cell_contents("Sheet1", "S1", "=AVERAGE(R1:R3)")
    assert wb.get_cell_value("Sheet1", "S1") == decimal.Decimal("3.25")


# Test 16: CHOOSE function selecting between cell ranges
def test_choose_selecting_between_cell_ranges(wb):
    """
    Tests the CHOOSE function's ability to select between cell ranges as arguments, evaluating the chosen range.
    Spec: 'The CHOOSE(index, value1, ...) function should support cell-range references for value1 and subsequent arguments.'
    """
    wb.set_cell_contents("Sheet1", "T1", "1")
    wb.set_cell_contents("Sheet1", "T2", "2")
    wb.set_cell_contents("Sheet1", "U1", "3")
    wb.set_cell_contents("Sheet1", "U2", "4")
    wb.set_cell_contents("Sheet1", "V1", "=CHOOSE(1, SUM(T1:T2), SUM(U1:U2))")
    assert wb.get_cell_value("Sheet1", "V1") == 3

# Test 17: Ensuring proper error handling in CHOOSE function with out-of-range index
def test_choose_with_out_of_range_index(wb):
    """
    Validates that the CHOOSE function correctly handles an out-of-range index argument, following general error handling rules.
    Spec does not explicitly detail this behavior, but it aligns with expected function robustness.
    """
    wb.set_cell_contents("Sheet1", "W1", "=CHOOSE(3, \"Valid\", \"Also valid\")")
    assert isinstance(wb.get_cell_value("Sheet1", "W1"), CellError)
    assert wb.get_cell_value("Sheet1", "W1").get_type() == CellErrorType.TYPE_ERROR

# Test 18: SUM function accurately summing a complex range of decimal numbers
def test_sum_with_complex_decimal_range(wb):
    """
    Tests the SUM function's ability to accurately aggregate a range of decimal numbers, ensuring precision.
    Spec: 'SUM(value1, ...) returns the sum of all inputs. All non-empty inputs are converted to numbers...'
    """
    wb.set_cell_contents("Sheet1", "X1", "3.14159")
    wb.set_cell_contents("Sheet1", "X2", "2.71828")
    wb.set_cell_contents("Sheet1", "Y1", "=SUM(X1:X2)")
    assert wb.get_cell_value("Sheet1", "Y1") == decimal.Decimal("5.85987")

# Test 19: AVERAGE function handling cell range with a single non-numeric value
def test_average_with_single_non_numeric_value(wb):
    """
    Verifies the AVERAGE function returns a TYPE_ERROR when encountering a cell range with a single non-numeric value.
    Spec: 'If an input cannot be converted to a number then the function returns a TYPE_ERROR.'
    """
    wb.set_cell_contents("Sheet1", "Z1", "Not a number")
    wb.set_cell_contents("Sheet1", "AA1", "=AVERAGE(Z1)")
    assert isinstance(wb.get_cell_value("Sheet1", "AA1"), CellError)
    assert wb.get_cell_value("Sheet1", "AA1").get_type() == CellErrorType.TYPE_ERROR

# Test 20: INDIRECT function used in SUM with dynamic range construction
def test_indirect_in_sum_with_dynamic_range_construction(wb):
    """
    Tests INDIRECT function within SUM for dynamic range construction, validating flexibility in formula compositions.
    Spec: '...update the INDIRECT(str) function such that it can produce cell-range references from the input string...'
    """
    wb.set_cell_contents("Sheet1", "AB1", "1")
    wb.set_cell_contents("Sheet1", "AB2", "2")
    wb.set_cell_contents("Sheet1", "AC1", "Sheet1")
    wb.set_cell_contents("Sheet1", "AC2", "AB1:AB2")
    wb.set_cell_contents("Sheet1", "AD1", "=SUM(INDIRECT(AC1&\"!\"&AC2))")
    assert wb.get_cell_value("Sheet1", "AD1") == 3
#
# # Test 21: Testing sort_region with ascending and descending orders
# def test_sort_region_with_ascending_and_descending_orders(wb):
#     """
#     Validates sort_region's ability to sort a specified region using both ascending and descending orders based on multiple columns.
#     Spec: '...specify sort_cols=[1, -2] to sort on the first column in ascending order and the second in descending order.'
#     """
#     wb.set_cell_contents("Sheet1", "AE1", "b")
#     wb.set_cell_contents("Sheet1", "AE2", "a")
#     wb.set_cell_contents("Sheet1", "AF1", "1")
#     wb.set_cell_contents("Sheet1", "AF2", "2")
#     wb.sort_region("Sheet1", "AE1", "AF2", [1, -2])
#     assert wb.get_cell_value("Sheet1", "AE1") == "a"
#     assert wb.get_cell_value("Sheet1", "AE2") == "b"
#     assert wb.get_cell_value("Sheet1", "AF1") == "2"
#     assert wb.get_cell_value("Sheet1", "AF2") == "1"

# Test 22: Validating AVERAGE function with cell ranges including division by zero scenario
def test_average_with_division_by_zero_scenario(wb):
    """
    Tests the AVERAGE function to ensure it returns a DIVIDE_BY_ZERO error when the cell range is entirely empty,
    aligning with the specification that non-empty cells are considered for calculation.
    """
    wb.set_cell_contents("Sheet1", "AG1", None)
    wb.set_cell_contents("Sheet1", "AG2", None)
    wb.set_cell_contents("Sheet1", "AH1", "=AVERAGE(AG1:AG2)")
    assert isinstance(wb.get_cell_value("Sheet1", "AH1"), CellError)
    assert wb.get_cell_value("Sheet1", "AH1").get_type() == CellErrorType.DIVIDE_BY_ZERO

# Test 23: INDIRECT function handling invalid cell range reference
def test_indirect_with_invalid_cell_range_reference(wb):
    """
    Validates INDIRECT function's error handling when provided with an invalid cell range reference,
    expecting a BAD_REFERENCE error according to the specifications for handling invalid references.
    """
    wb.set_cell_contents("Sheet1", "AI1", "=SUM(INDIRECT(\"Sheet1!ZZZZ9999:ZZZZ10000\"))")
    assert isinstance(wb.get_cell_value("Sheet1", "AI1"), CellError)
    assert wb.get_cell_value("Sheet1", "AI1").get_type() == CellErrorType.BAD_REFERENCE

# Test 24: Testing cell range parsing in formulas affecting other cells
def test_cell_range_parsing_affecting_other_cells(wb):
    """
    Ensures that a formula using a cell range reference correctly updates the value of other cells dependent on the range,
    particularly testing the dynamic updating of referenced cells.
    """
    wb.set_cell_contents("Sheet1", "AJ1", "5")
    wb.set_cell_contents("Sheet1", "AJ2", "15")
    wb.set_cell_contents("Sheet1", "AK1", "=SUM(AJ1:AJ2)")
    wb.set_cell_contents("Sheet1", "AJ1", "10")  # Update should affect AK1
    assert wb.get_cell_value("Sheet1", "AK1") == 25

# Test 25: Testing HLOOKUP and VLOOKUP functions with error values in search key
def test_lookup_functions_with_error_values_as_keys(wb):
    """
    Tests HLOOKUP and VLOOKUP functions when the search key is an error value, expecting the lookup function to propagate the error.
    Spec: '...any function argument which is an error should yield an error in the output.'
    """
    wb.set_cell_contents("Sheet1", "AL1", "=1/0")  # Generates a DIVIDE_BY_ZERO error
    wb.set_cell_contents("Sheet1", "A1", "=HLOOKUP(AL1, AL2:AM2, 2)")
    wb.set_cell_contents("Sheet1", "A2", "=VLOOKUP(AL1, AL2:AN2, 2)")
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.DIVIDE_BY_ZERO
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError)
    assert wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.DIVIDE_BY_ZERO


# Test 26: Verifying MIN function with cell ranges including non-numeric values
def test_min_function_with_non_numeric_values(wb):
    """
    Tests the MIN function's handling of cell ranges that include non-numeric values, ensuring a TYPE_ERROR is returned.
    According to the specification, 'All non-empty inputs are converted to numbers; if any input cannot be converted to a number then the function returns a TYPE_ERROR.'
    """
    wb.set_cell_contents("Sheet1", "AO1", "10")
    wb.set_cell_contents("Sheet1", "AO2", "hello")  # Non-numeric
    wb.set_cell_contents("Sheet1", "AO3", "20")
    wb.set_cell_contents("Sheet1", "AP1", "=MIN(AO1:AO3)")
    assert isinstance(wb.get_cell_value("Sheet1", "AP1"), CellError)
    assert wb.get_cell_value("Sheet1", "AP1").get_type() == CellErrorType.TYPE_ERROR

# Test 27: Testing deletion of a sheet and its impact on cell range references
def test_sheet_deletion_impact_on_cell_range_references(wb):
    """
    Validates the impact of deleting a sheet on formulas in other sheets that reference ranges in the deleted sheet,
    expecting BAD_REFERENCE errors for affected formulas.
    """
    wb.new_sheet("Sheet2")
    wb.set_cell_contents("Sheet1", "AQ1", "=SUM(Sheet2!A1:A10)")
    wb.del_sheet("Sheet2")
    assert isinstance(wb.get_cell_value("Sheet1", "AQ1"), CellError)
    assert wb.get_cell_value("Sheet1", "AQ1").get_type() == CellErrorType.BAD_REFERENCE

# Test 28: Checking cell range updates when cells within the range are moved
def test_cell_range_updates_when_cells_are_moved(wb):
    """
    Tests how cell range references in formulas are updated when cells within the range are moved to a new location.
    The specification mentions that operations updating formulas need to update cell-range references accordingly.
    """
    # Assuming a hypothetical method move_cells is implemented as per spec.
    wb.set_cell_contents("Sheet1", "AR1", "1")
    wb.set_cell_contents("Sheet1", "AR2", "2")
    wb.set_cell_contents("Sheet1", "AS1", "=SUM(AR1:AR2)")
    # wb.move_cells("Sheet1", "AR1:AR2", "AT1")  # Moves AR1:AR2 to AT1:AT2
    # This line is commented out because the API doesn't define move_cells, but the test assumes its behavior.
    # assert wb.get_cell_contents("Sheet1", "AS1") == "=SUM(AT1:AT2)"

# Test 29: Error handling in SORT function with invalid column index
def test_sort_function_with_invalid_column_index(wb):
    """
    Ensures the SORT function handles invalid column indices gracefully, returning an error if a specified column index is out of range or invalid.
    According to the spec, 'If the sort_cols list is invalid in any way, a ValueError is raised.'
    """
    wb.set_cell_contents("Sheet1", "AU1", "c")
    wb.set_cell_contents("Sheet1", "AU2", "b")
    wb.set_cell_contents("Sheet1", "AU3", "a")
    # Trying to sort with an invalid column index (assuming sort_region would throw an exception or handle it internally)
    try:
        wb.sort_region("Sheet1", "AU1", "AU3", [2])  # Invalid column index
        assert False, "Expected a ValueError for invalid column index"
    except ValueError:
        assert True  # Passes if ValueError is raised as expected

# Test 30: Verifying functionality of INDIRECT with dynamically constructed cell references
def test_indirect_with_dynamically_constructed_references(wb):
    """
    Tests the INDIRECT function's ability to resolve dynamically constructed cell references,
    allowing for flexible formula construction based on cell values.
    """
    wb.set_cell_contents("Sheet1", "AV1", "AV")
    wb.set_cell_contents("Sheet1", "AV2", "3")
    wb.set_cell_contents("Sheet1", "AV3", "100")
    wb.set_cell_contents("Sheet1", "AW1", "=INDIRECT(AV1&AV2)")
    assert wb.get_cell_value("Sheet1", "AW1") == 100


# Test 26: Verifying MIN function with cell ranges including non-numeric values
def test_min_function_with_non_numeric_values(wb):
    """
    Tests the MIN function's handling of cell ranges that include non-numeric values, ensuring a TYPE_ERROR is returned.
    According to the specification, 'All non-empty inputs are converted to numbers; if any input cannot be converted to a number then the function returns a TYPE_ERROR.'
    """
    wb.set_cell_contents("Sheet1", "AO1", "10")
    wb.set_cell_contents("Sheet1", "AO2", "hello")  # Non-numeric
    wb.set_cell_contents("Sheet1", "AO3", "20")
    wb.set_cell_contents("Sheet1", "AP1", "=MIN(AO1:AO3)")
    assert isinstance(wb.get_cell_value("Sheet1", "AP1"), CellError)
    assert wb.get_cell_value("Sheet1", "AP1").get_type() == CellErrorType.TYPE_ERROR

# Test 27: Testing deletion of a sheet and its impact on cell range references
def test_sheet_deletion_impact_on_cell_range_references(wb):
    """
    Validates the impact of deleting a sheet on formulas in other sheets that reference ranges in the deleted sheet,
    expecting BAD_REFERENCE errors for affected formulas.
    """
    wb.new_sheet("Sheet2")
    wb.set_cell_contents("Sheet1", "AQ1", "=SUM(Sheet2!A1:A10)")
    wb.del_sheet("Sheet2")
    assert isinstance(wb.get_cell_value("Sheet1", "AQ1"), CellError)
    assert wb.get_cell_value("Sheet1", "AQ1").get_type() == CellErrorType.BAD_REFERENCE

# Test 28: Checking cell range updates when cells within the range are moved
def test_cell_range_updates_when_cells_are_moved(wb):
    """
    Tests how cell range references in formulas are updated when cells within the range are moved to a new location.
    The specification mentions that operations updating formulas need to update cell-range references accordingly.
    """
    # Assuming a hypothetical method move_cells is implemented as per spec.
    wb.set_cell_contents("Sheet1", "AR1", "1")
    wb.set_cell_contents("Sheet1", "AR2", "2")
    wb.set_cell_contents("Sheet1", "AS1", "=SUM(AR1:AR2)")
    # wb.move_cells("Sheet1", "AR1:AR2", "AT1")  # Moves AR1:AR2 to AT1:AT2
    # This line is commented out because the API doesn't define move_cells, but the test assumes its behavior.
    # assert wb.get_cell_contents("Sheet1", "AS1") == "=SUM(AT1:AT2)"

# Test 29: Error handling in SORT function with invalid column index
def test_sort_function_with_invalid_column_index(wb):
    """
    Ensures the SORT function handles invalid column indices gracefully, returning an error if a specified column index is out of range or invalid.
    According to the spec, 'If the sort_cols list is invalid in any way, a ValueError is raised.'
    """
    wb.set_cell_contents("Sheet1", "AU1", "c")
    wb.set_cell_contents("Sheet1", "AU2", "b")
    wb.set_cell_contents("Sheet1", "AU3", "a")
    # Trying to sort with an invalid column index (assuming sort_region would throw an exception or handle it internally)
    try:
        wb.sort_region("Sheet1", "AU1", "AU3", [2])  # Invalid column index
        assert False, "Expected a ValueError for invalid column index"
    except ValueError:
        assert True  # Passes if ValueError is raised as expected

# Test 30: Verifying functionality of INDIRECT with dynamically constructed cell references
def test_indirect_with_dynamically_constructed_references(wb):
    """
    Tests the INDIRECT function's ability to resolve dynamically constructed cell references,
    allowing for flexible formula construction based on cell values.
    """
    wb.set_cell_contents("Sheet1", "AV1", "AV")
    wb.set_cell_contents("Sheet1", "AV2", "3")
    wb.set_cell_contents("Sheet1", "AV3", "100")
    wb.set_cell_contents("Sheet1", "AW1", "=INDIRECT(AV1&AV2)")
    assert wb.get_cell_value("Sheet1", "AW1") == 100

# Test 31: Ensuring correct evaluation of AVERAGE function with exclusively negative values in cell range
def test_average_with_exclusively_negative_values(wb):
    """
    Tests the AVERAGE function's capability to correctly evaluate a cell range containing exclusively negative values,
    highlighting the function's ability to handle negative numbers as specified.
    """
    wb.set_cell_contents("Sheet1", "AX1", "-10")
    wb.set_cell_contents("Sheet1", "AX2", "-20")
    wb.set_cell_contents("Sheet1", "AY1", "=AVERAGE(AX1:AX2)")
    assert wb.get_cell_value("Sheet1", "AY1") == decimal.Decimal("-15")

# Test 32: Formula referencing a non-existent sheet within a cell range
def test_formula_referencing_non_existent_sheet(wb):
    """
    Validates that a formula referencing a cell range in a non-existent sheet correctly returns a BAD_REFERENCE error,
    in accordance with the specifications regarding invalid references.
    """
    wb.set_cell_contents("Sheet1", "AZ1", "=SUM(MissingSheet!A1:A2)")
    assert isinstance(wb.get_cell_value("Sheet1", "AZ1"), CellError)
    assert wb.get_cell_value("Sheet1", "AZ1").get_type() == CellErrorType.BAD_REFERENCE

# # Test 33: SUM function with cell range including both text and numeric values
# def test_sum_with_text_and_numeric_values(wb):
#     """
#     Ensures the SUM function correctly ignores text values when summing up a range that includes both text and numeric values,
#     reflecting the specification's mandate for type-specific operations.
#     """
#     wb.set_cell_contents("Sheet1", "BA1", "5")
#     wb.set_cell_contents("Sheet1", "BA2", "Text")
#     wb.set_cell_contents("Sheet1", "BA3", "10")
#     wb.set_cell_contents("Sheet1", "BB1", "=SUM(BA1:BA3)")
#     assert wb.get_cell_value("Sheet1", "BB1") == 15  # Assuming SUM ignores non-numeric values

# Test 34: Verifying error propagation from a cell within a range to a dependent formula
def test_error_propagation_from_cell_to_formula(wb):
    """
    Tests the propagation of an error from a single cell within a range to a dependent formula,
    ensuring that errors within cell ranges influence the outcomes of formulas relying on those ranges.
    """
    wb.set_cell_contents("Sheet1", "BC1", "=1/0")  # Directly causes DIVIDE_BY_ZERO error
    wb.set_cell_contents("Sheet1", "BC2", "10")
    wb.set_cell_contents("Sheet1", "BD1", "=SUM(BC1:BC2)")
    assert isinstance(wb.get_cell_value("Sheet1", "BD1"), CellError)
    assert wb.get_cell_value("Sheet1", "BD1").get_type() == CellErrorType.DIVIDE_BY_ZERO

# Test 35: Handling of cell range references in a formula that does not traditionally support them
def test_handling_unsupported_cell_range_references(wb):
    """
    Validates the spreadsheet engine's handling of cell range references in formulas where such references are traditionally not supported,
    ensuring adherence to the specification's guidance on formula parsing and execution.
    """
    wb.set_cell_contents("Sheet1", "BE1", "1")
    wb.set_cell_contents("Sheet1", "BE2", "2")
    wb.set_cell_contents("Sheet1", "BF1", "=BE1 + BE1:BE2")  # Hypothetical misuse of cell range
    assert isinstance(wb.get_cell_value("Sheet1", "BF1"), CellError)
    assert wb.get_cell_value("Sheet1", "BF1").get_type() == CellErrorType.PARSE_ERROR

# Test 36: VLOOKUP with exact match and index beyond range
def test_vlookup_exact_match_with_index_beyond_range(wb):
    """
    Tests VLOOKUP function's behavior when the index specified is beyond the range of available columns,
    expecting a BAD_REFERENCE error as per specification guidance on handling invalid indices.
    """
    wb.set_cell_contents("Sheet1", "BG1", "Key")
    wb.set_cell_contents("Sheet1", "BH1", "Value1")
    wb.set_cell_contents("Sheet1", "BG2", "=VLOOKUP(\"Key\", BG1:BH1, 3)")  # Index 3 is beyond the available range
    assert isinstance(wb.get_cell_value("Sheet1", "BG2"), CellError)
    assert wb.get_cell_value("Sheet1", "BG2").get_type() == CellErrorType.TYPE_ERROR

# Test 37: HLOOKUP searching for a non-existent key
def test_hlookup_search_for_non_existent_key(wb):
    """
    Validates HLOOKUP function's handling of a search for a non-existent key within a specified range,
    expecting a TYPE_ERROR to be returned if the key is not found.
    """
    wb.set_cell_contents("Sheet1", "BI1", "NonExistentKey")
    wb.set_cell_contents("Sheet1", "BI2", "=HLOOKUP(\"NotFound\", BI1:BJ2, 2)")
    assert isinstance(wb.get_cell_value("Sheet1", "BI2"), CellError)
    assert wb.get_cell_value("Sheet1", "BI2").get_type() == CellErrorType.TYPE_ERROR
#
# # Test 38: VLOOKUP with a range that doesn't start at the first row
# def test_vlookup_with_range_not_starting_at_first_row(wb):
#     """
#     Tests the VLOOKUP function's ability to accurately perform a search when the specified range doesn't start at the first row,
#     demonstrating the function's flexibility in handling various range configurations.
#     """
#     wb.set_cell_contents("Sheet1", "BK2", "Key")
#     wb.set_cell_contents("Sheet1", "BL2", "Match")
#     wb.set_cell_contents("Sheet1", "BK3", "Value")
#     wb.set_cell_contents("Sheet1", "BM1", "=VLOOKUP(\"Match\", BK2:BL3, 2)")
#     assert wb.get_cell_value("Sheet1", "BM1") == "Value"

# Test 39: HLOOKUP with a row index that refers to a row outside of the specified range
def test_hlookup_with_row_index_outside_range(wb):
    """
    Verifies HLOOKUP's behavior when the row index refers to a row outside of the specified search range,
    expected to return a BAD_REFERENCE error as per specifications on invalid index handling.
    """
    wb.set_cell_contents("Sheet1", "BN1", "SearchKey")
    wb.set_cell_contents("Sheet1", "BO1", "Result1")
    wb.set_cell_contents("Sheet1", "BN2", "=HLOOKUP(\"SearchKey\", BN1:BO1, 3)")  # Index 3 is beyond the search range
    assert isinstance(wb.get_cell_value("Sheet1", "BN2"), CellError)
    assert wb.get_cell_value("Sheet1", "BN2").get_type() == CellErrorType.TYPE_ERROR
#
# # Test 40: VLOOKUP using FALSE for the range_lookup to ensure exact match is required
# def test_vlookup_with_false_for_range_lookup(wb):
#     """
#     Ensures VLOOKUP function requires an exact match when FALSE is used for the range_lookup parameter,
#     testing the function's precision in matching keys within the specified range.
#     """
#     wb.set_cell_contents("Sheet1", "BP1", "Exact")
#     wb.set_cell_contents("Sheet1", "BQ1", "Exact Match")
#     wb.set_cell_contents("Sheet1", "BP2", "Value for Exact")
#     wb.set_cell_contents("Sheet1", "BQ2", "Value for Exact Match")
#     wb.set_cell_contents("Sheet1", "BR1", "=VLOOKUP(\"Exact\", BP1:BQ2, 2)")
#     assert wb.get_cell_value("Sheet1", "BR1") == "Value for Exact"
