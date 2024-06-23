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

def test_example_sort(wb : Workbook):
    wb.set_cell_contents("Sheet1" , "A1", "1")
    wb.set_cell_contents("Sheet1" , "A2", "2")
    wb.set_cell_contents("Sheet1" , "A3", "3")
    wb.set_cell_contents("Sheet1" , "A4", "5")
    wb.set_cell_contents("Sheet1" , "A5", "4")
    wb.sort_region("Sheet1","A1", "A5", [1])
    assert wb.get_cell_value("Sheet1" , "A1") == 1
    assert wb.get_cell_value("Sheet1" , "A2") == 2
    assert wb.get_cell_value("Sheet1" , "A3") == 3
    assert wb.get_cell_value("Sheet1" , "A4") == 4
    assert wb.get_cell_value("Sheet1" , "A5") == 5

# Test 2: Sort on multiple columns, first ascending, second descending
def test_sort_multiple_columns(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "Apple")
    wb.set_cell_contents("Sheet1", "B1", "3")
    wb.set_cell_contents("Sheet1", "A2", "Banana")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "A3", "Apple")
    wb.set_cell_contents("Sheet1", "B3", "1")
    wb.sort_region("Sheet1", "A1", "B3", [1, -2])
    assert wb.get_cell_value("Sheet1", "A1") == "Apple"
    assert wb.get_cell_value("Sheet1", "B1") == 3
    assert wb.get_cell_value("Sheet1", "A2") == "Apple"
    assert wb.get_cell_value("Sheet1", "B2") == 1
    assert wb.get_cell_value("Sheet1", "A3") == "Banana"
    assert wb.get_cell_value("Sheet1", "B3") == 2

# Test 3: Sort with mixed data types
def test_sort_with_mixed_types(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "Banana")
    wb.set_cell_contents("Sheet1", "A3", "")
    wb.set_cell_contents("Sheet1", "A4", "=2+2")
    wb.sort_region("Sheet1", "A1", "A4", [1])
    assert wb.get_cell_value("Sheet1", "A1") is None
    assert wb.get_cell_value("Sheet1", "A2") == 4  # Assuming formulas evaluate immediately
    assert wb.get_cell_value("Sheet1", "A3") == 10
    assert wb.get_cell_value("Sheet1", "A4") == "Banana"

# Test 4: Sorting empty and error-valued cells
def test_sort_empty_and_error_cells(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "A2", "=#REF!")
    wb.set_cell_contents("Sheet1", "A3", "Apple")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert wb.get_cell_value("Sheet1", "A1") is None
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError)
    assert wb.get_cell_value("Sheet1", "A3") == "Apple"

# Test 5: Sorting with descending order
def test_sort_descending_order(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "A2", "3")
    wb.set_cell_contents("Sheet1", "A3", "4")
    wb.set_cell_contents("Sheet1", "A4", "2")
    wb.set_cell_contents("Sheet1", "A5", "1")
    wb.sort_region("Sheet1", "A1", "A5", [-1])
    assert wb.get_cell_value("Sheet1", "A1") == 5
    assert wb.get_cell_value("Sheet1", "A2") == 4
    assert wb.get_cell_value("Sheet1", "A3") == 3
    assert wb.get_cell_value("Sheet1", "A4") == 2
    assert wb.get_cell_value("Sheet1", "A5") == 1

# Test 6: Testing the stability of the sort
def test_sort_stability(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "Apple")
    wb.set_cell_contents("Sheet1", "B1", "2")
    wb.set_cell_contents("Sheet1", "A2", "Banana")
    wb.set_cell_contents("Sheet1", "B2", "1")
    wb.set_cell_contents("Sheet1", "A3", "Apple")
    wb.set_cell_contents("Sheet1", "B3", "1")
    wb.sort_region("Sheet1", "A1", "B3", [-1])
    assert wb.get_cell_value("Sheet1", "A1") == "Banana"
    assert wb.get_cell_value("Sheet1", "B1") == 1
    assert wb.get_cell_value("Sheet1", "A2") == "Apple"
    assert wb.get_cell_value("Sheet1", "B2") == 2
    assert wb.get_cell_value("Sheet1", "A3") == "Apple"
    assert wb.get_cell_value("Sheet1", "B3") == 1


# Test 7: Attempting to sort with an invalid column index
def test_sort_invalid_column_index(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "1")
    wb.set_cell_contents("Sheet1", "A2", "2")
    with pytest.raises(ValueError):
        wb.sort_region("Sheet1", "A1", "A2", [3])  # Invalid column index, beyond the extent

# Test 8: Attempting to sort a non-existent sheet
def test_sort_non_existent_sheet(wb: Workbook):
    with pytest.raises(KeyError):
        wb.sort_region("Sheet2", "A1", "A2", [1])  # Sheet2 does not exist

# Test 9: Sorting a single column with mixed numeric and string data, ascending
def test_sort_mixed_data_ascending(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "apple")
    wb.set_cell_contents("Sheet1", "A3", "2")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert wb.get_cell_value("Sheet1", "A1") ==  2
    assert wb.get_cell_value("Sheet1", "A2") == 10
    assert wb.get_cell_value("Sheet1", "A3") == "apple"

def test_sort_mixed_data_simple(wb: Workbook):
    val_1 = "TRUE"
    val_2 = "bruh"
    wb.set_cell_contents("Sheet1", "A1", val_1)
    wb.set_cell_contents("Sheet1", "A2", val_2)
    starting_order = [wb.get_cell_value("Sheet1", "A1") , wb.get_cell_value("Sheet1", "A2")]

    wb.sort_region("Sheet1", "A1", "A2", [-1])
    new_order_1 = [wb.get_cell_value("Sheet1", "A1") , wb.get_cell_value("Sheet1", "A2")]
    wb.set_cell_contents("Sheet1", "A1", val_1)
    wb.set_cell_contents("Sheet1", "A2", val_2)
    wb.sort_region("Sheet1", "A1", "A2", [1])
    new_order_2 = [wb.get_cell_value("Sheet1", "A1") , wb.get_cell_value("Sheet1", "A2")]
    assert new_order_1 != starting_order or new_order_2 != starting_order


# Test 10: Sorting a single column with mixed numeric and string data, descending
def test_sort_mixed_data_descending(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "apple")
    wb.set_cell_contents("Sheet1", "A3", "2")
    wb.sort_region("Sheet1", "A1", "A3", [-1])
    assert wb.get_cell_value("Sheet1", "A1") == "apple"
    assert wb.get_cell_value("Sheet1", "A2") == 10
    assert wb.get_cell_value("Sheet1", "A3") == 2

# Test 11: Sorting stability with multiple sort keys, including descending order for one key
def test_sort_stability_multiple_keys_descending(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "Banana")
    wb.set_cell_contents("Sheet1", "B1", "2")
    wb.set_cell_contents("Sheet1", "A2", "Apple")
    wb.set_cell_contents("Sheet1", "B2", "3")
    wb.set_cell_contents("Sheet1", "A3", "Banana")
    wb.set_cell_contents("Sheet1", "B3", "1")
    wb.sort_region("Sheet1", "A1", "B3", [1, -2])
    assert wb.get_cell_value("Sheet1", "A1") == "Apple"
    assert wb.get_cell_value("Sheet1", "B1") == 3
    assert wb.get_cell_value("Sheet1", "A2") == "Banana"
    assert wb.get_cell_value("Sheet1", "B2") == 2
    assert wb.get_cell_value("Sheet1", "A3") == "Banana"
    assert wb.get_cell_value("Sheet1", "B3") == 1


# Test 12: Sorting with complex column order and mixed ascending/descending
def test_sort_complex_column_order_mixed_directions(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "Banana")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "C1", "Zebra")
    wb.set_cell_contents("Sheet1", "A2", "Apple")
    wb.set_cell_contents("Sheet1", "B2", "2")
    wb.set_cell_contents("Sheet1", "C2", "Yak")
    wb.sort_region("Sheet1", "A1", "C2", [1, -2, 3])
    assert wb.get_cell_value("Sheet1", "A1") == "Apple"
    assert wb.get_cell_value("Sheet1", "B1") == 2
    assert wb.get_cell_value("Sheet1", "C1") == "Yak"
    assert wb.get_cell_value("Sheet1", "A2") == "Banana"
    assert wb.get_cell_value("Sheet1", "B2") == 1
    assert wb.get_cell_value("Sheet1", "C2") == "Zebra"

# Test 13: Sorting a large dataset to check for performance issues
def test_sort_large_dataset(wb: Workbook):
    # Assuming set_cell_contents can handle batch input for testing large datasets
    for i in range(1, 101):
        wb.set_cell_contents("Sheet1", f"A{i}", str(101 - i))
    wb.sort_region("Sheet1", "A1", "A100", [1])
    assert wb.get_cell_value("Sheet1", "A1") == 1
    assert wb.get_cell_value("Sheet1", "A50") == 50
    assert wb.get_cell_value("Sheet1", "A100") == 100

# Test 14: Ensure sorting one sheet does not affect another
def test_sort_one_sheet_does_not_affect_another(wb: Workbook):
    wb.new_sheet("Sheet2")
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "1")
    wb.set_cell_contents("Sheet2", "A1", "B")
    wb.set_cell_contents("Sheet2", "A2", "A")
    wb.sort_region("Sheet1", "A1", "A2", [1])
    assert wb.get_cell_value("Sheet1", "A1") == 1
    assert wb.get_cell_value("Sheet1", "A2") == 2
    assert wb.get_cell_value("Sheet2", "A1") == "B"  # Should remain unchanged
    assert wb.get_cell_value("Sheet2", "A2") == "A"  # Should remain unchanged

# Test 15: Sorting correctly updates formulas referencing sorted cells
def test_sort_updates_formulas(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "1")
    wb.set_cell_contents("Sheet1", "B1", "=A1")
    wb.set_cell_contents("Sheet1", "B2", "=A2")
    wb.sort_region("Sheet1", "A1", "A2", [-1])
    # Assuming formulas are dynamically updated to reference the new cell positions
    assert wb.get_cell_value("Sheet1", "B1") == 2  # Now B1 should reference A1 which is 2
    assert wb.get_cell_value("Sheet1", "B2") == 1  # Now B2 should reference A2 which is 1

# Test 16: Sorting with all cells in a column being identical
def test_sort_all_cells_identical_in_column(wb: Workbook):
    for i in range(1, 11):
        wb.set_cell_contents("Sheet1", f"A{i}", "Apple")
        wb.set_cell_contents("Sheet1", f"B{i}", str(i))
    wb.sort_region("Sheet1", "A1", "B10", [1])
    # Since all A values are identical, B column should determine the order post-sort
    for i in range(1, 11):
        assert wb.get_cell_value("Sheet1", f"B{i}")

# Test 17: Sorting a region with merged cells
# def test_sort_with_merged_cells(wb: Workbook):
#     wb.set_cell_contents("Sheet1", "A1", "1")
#     wb.set_cell_contents("Sheet1", "A2", "2")
#     wb.set_cell_contents("Sheet1", "B1", "3")
#     wb.set_cell_contents("Sheet1", "B2", "4")
#     # Assuming a hypothetical merge_cells function for the sake of this test
#     wb.merge_cells("Sheet1", "A1", "B1")
#     with pytest.raises(ValueError):
#         wb.sort_region("Sheet1", "A1", "B2", [1])  # Should raise due to merged cells complicating the sort

# Test 18: Sorting based on a mix of numeric and formula cells in the same column
def test_sort_numeric_and_formula_mixed(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "=5+5")  # Evaluates to 10
    wb.set_cell_contents("Sheet1", "A3", "5")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert wb.get_cell_value("Sheet1", "A1") ==5
    assert wb.get_cell_value("Sheet1", "A2") == 10  # Assuming the value is fetched post-evaluation
    assert wb.get_cell_value("Sheet1", "A3") == 10

# Test 19: Sort leading to formula inversion
def test_sort_leading_to_formula_inversion(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=B1")
    wb.set_cell_contents("Sheet1", "A2", "=B2")
    wb.set_cell_contents("Sheet1", "B1", "2")
    wb.set_cell_contents("Sheet1", "B2", "1")
    wb.sort_region("Sheet1", "A1", "B2", [-2])
    # Verifying that formulas still reference correctly after sort
    assert wb.get_cell_contents("Sheet1", "A1") == "=B1"  # Should still reference the correct cell
    assert wb.get_cell_contents("Sheet1", "A2") == "=B2"



# Test 21: Case sensitivity in string sorting
def test_sort_case_sensitivity(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "banana")
    wb.set_cell_contents("Sheet1", "A2", "Apple")
    wb.set_cell_contents("Sheet1", "A3", "apple")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    # Assuming case sensitivity affects the sort, 'Apple' should come before 'apple' and 'banana'
    assert wb.get_cell_value("Sheet1", "A1") == "Apple"
    assert wb.get_cell_value("Sheet1", "A2") == "apple"
    assert wb.get_cell_value("Sheet1", "A3") == "banana"

# Test 22: Sorting with formulas referencing sorted cells
def test_sort_formulas_referencing_sorted_cells(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "3")
    wb.set_cell_contents("Sheet1", "A2", "1")
    wb.set_cell_contents("Sheet1", "A3", "=A1+A2")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert wb.get_cell_value("Sheet1", "A1") == 1
    assert wb.get_cell_value("Sheet1", "A2") == 3
    # Assuming the engine updates formulas based on sorted references, otherwise, it's a design decision
    assert wb.get_cell_contents("Sheet1", "A3") == "=A1+A2"
    assert wb.get_cell_value("Sheet1", "A3") == 4  # This assumes formulas are recalculated after sort

# Test 23: Sorting regions that include cells with identical formulas
def test_sort_regions_with_identical_formulas(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=2*2")
    wb.set_cell_contents("Sheet1", "A2", "=2*2")
    wb.set_cell_contents("Sheet1", "A3", "=3+1")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    # Since all formulas evaluate to the same value, their order should remain as is due to sort stability
    assert wb.get_cell_contents("Sheet1", "A1") == "=2*2"
    assert wb.get_cell_contents("Sheet1", "A2") == "=2*2"
    assert wb.get_cell_contents("Sheet1", "A3") == "=3+1"

# Test 24: Sorting cells containing strings resembling numbers
def test_sort_strings_resembling_numbers(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "100")
    wb.set_cell_contents("Sheet1", "A2", "'100")
    wb.set_cell_contents("Sheet1", "A3", "20")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert wb.get_cell_value("Sheet1", "A1") == 20
    assert wb.get_cell_value("Sheet1", "A2") == 100  # Assuming '100' (string) is treated differently than 100 (number)
    assert wb.get_cell_value("Sheet1", "A3") == "100"

# Test 25: Sorting with date-formatted strings
def test_sort_date_formatted_strings(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2022-01-01")
    wb.set_cell_contents("Sheet1", "A2", "2021-12-31")
    wb.set_cell_contents("Sheet1", "A3", "2022-01-02")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert wb.get_cell_value("Sheet1", "A1") == "2021-12-31"
    assert wb.get_cell_value("Sheet1", "A2") == "2022-01-01"
    assert wb.get_cell_value("Sheet1", "A3") == "2022-01-02"

# Test 27: Sorting cells that contain error values alongside valid numbers
def test_sort_error_values_with_numbers(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=1/0")  # This would generate a DIV/0 error
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "A3", "=A2/0")  # Another DIV/0 error
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError)
    assert isinstance(wb.get_cell_value("Sheet1", "A3"), decimal.Decimal)


# Test 28: Sorting cells where sorting causes a formula to reference an error cell
def test_sort_causing_formula_reference_error(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=B1")  # Initially valid
    wb.set_cell_contents("Sheet1", "B1", "5")
    wb.set_cell_contents("Sheet1", "A2", "=1/0")  # Error cell
    wb.sort_region("Sheet1", "A1", "A2", [1])
    # After sorting, A1 might reference the error from A2
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)

# Test 29: Sorting with multiple types of error values
def test_sort_multiple_error_types(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=1/0")  # DIV/0
    wb.set_cell_contents("Sheet1", "A2", "=#REF!")  # REF error directly inputted
    wb.set_cell_contents("Sheet1", "A3", "=NONEXISTENT()")  # BAD_NAME error
    wb.sort_region("Sheet1", "A1", "A3", [1])
    # Assuming error types are sorted based on their enumeration values
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)  # BAD_NAME is lowest
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError)  # REF
    assert isinstance(wb.get_cell_value("Sheet1", "A3"), CellError)  # DIV/0 is highest

# Test 30: Ensuring error cells sort as 'less than' any valid cell
def test_sort_error_cells_as_less_than(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=1/0")  # Error
    wb.set_cell_contents("Sheet1", "A2", "apple")  # String
    wb.set_cell_contents("Sheet1", "A3", "10")  # Number
    wb.sort_region("Sheet1", "A1", "A3", [1])
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError)  # Error sorts first
    assert wb.get_cell_value("Sheet1", "A2") == decimal.Decimal("10")  # Number comes next
    assert wb.get_cell_value("Sheet1", "A3") == "apple"  # String comes last

# Test 31: Sorting does not remove or alter error types but relocates them appropriately
def test_sort_preserves_error_types(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=1/0")  # DIV/0 error
    wb.set_cell_contents("Sheet1", "A2", "=#NAME?")  # BAD_NAME error, directly inputted
    wb.set_cell_contents("Sheet1", "A3", "=Sheet2!A4")  # Potential REF error if A4 does not exist
    wb.sort_region("Sheet1", "A1", "A3", [1])
    # After sorting, check if errors preserve their type and detail
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError) and wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError) and wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.BAD_NAME
    assert isinstance(wb.get_cell_value("Sheet1", "A3"), CellError) and wb.get_cell_value("Sheet1", "A3").get_type() == CellErrorType.DIVIDE_BY_ZERO

# Test 32: Formula becoming a circular reference after sorting
def test_sort_maintains_circular_reference(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet1", "A2", "=A1")
    wb.sort_region("Sheet1", "A1", "A2", [-1])
    # Check if the circular reference error is correctly identified
    assert isinstance(wb.get_cell_value("Sheet1", "A1"), CellError) and wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError) and wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE

# Test 33: Sorting with extremely large numbers
def test_sort_extremely_large_numbers(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", str(10**308))  # Near the upper limit of floating point
    wb.set_cell_contents("Sheet1", "A2", str(-10**308))  # Near the lower limit
    wb.sort_region("Sheet1", "A1", "A2", [1])
    print(wb.get_cell_value("Sheet1", "A1"))
    assert wb.get_cell_value("Sheet1", "A1") == "-100000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
    assert wb.get_cell_value("Sheet1", "A2") == "100000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"

# Test 34: Sorting very close numeric values to test precision
def test_sort_close_numeric_values(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "1.0000000000000001")  # Precision beyond typical float
    wb.set_cell_contents("Sheet1", "A2", "1.0000000000000002")
    wb.sort_region("Sheet1", "A1", "A2", [1])
    # Ensure that sorting distinguishes between very close values, thanks to decimal.Decimal precision
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal("1.0000000000000001")
    assert wb.get_cell_value("Sheet1", "A2") == decimal.Decimal("1.0000000000000002")

# Test 35: Sorting regions overlapping with fixed cell references in formulas
def test_sort_overlapping_fixed_references(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "A2", "=A$1")  # Fixed reference to A1
    wb.set_cell_contents("Sheet1", "A3", "10")
    wb.sort_region("Sheet1", "A1", "A3", [-1])

    assert wb.get_cell_contents("Sheet1", "A3") == "=A$1"
    assert wb.get_cell_value("Sheet1", "A3") == decimal.Decimal("10")
    assert wb.get_cell_contents("Sheet1", "A1") == "10"
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal("10")
    assert wb.get_cell_contents("Sheet1", "A2") == "5"
    assert wb.get_cell_value("Sheet1", "A2") == decimal.Decimal("5")


# Test 36: Sorting cells with boolean values
def test_sort_boolean_values(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "TRUE")
    wb.set_cell_contents("Sheet1", "A2", "FALSE")
    wb.set_cell_contents("Sheet1", "A3", "TRUE")
    wb.sort_region("Sheet1", "A1", "A3", [1])
    # Assuming boolean values are treated distinctly and sorted logically with FALSE before TRUE
    assert wb.get_cell_contents("Sheet1", "A1") == "FALSE"
    assert wb.get_cell_contents("Sheet1", "A2") == "TRUE"
    assert wb.get_cell_contents("Sheet1", "A3") == "TRUE"

# Test 37: Sorting cells with boolean values
def test_sort_to_bad_ref(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "5")
    wb.set_cell_contents("Sheet1", "A2", "=A1 + 5")

    wb.sort_region("Sheet1", "A1", "A2", [-1])
    # Assuming boolean values are treated distinctly and sorted logically with FALSE before TRUE
    assert isinstance(wb.get_cell_value("Sheet1", "A1") , CellError) and (wb.get_cell_value("Sheet1", "A1").get_type() == CellErrorType.BAD_REFERENCE )


def test_empty_cols(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "A2", "")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    try:
        wb.sort_region("Sheet1", "A1", "B2", [])
    except ValueError:
        assert True
    else:
        assert False

def test_invalid_cols(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "A2", "")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    try:
        wb.sort_region("Sheet1", "A1", "B2", ["a"])
    except ValueError:
        assert True
    else:
        assert False


def test_equal_string(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "a")
    wb.set_cell_contents("Sheet1", "A2", "a")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    wb.sort_region("Sheet1", "A1", "B2", [1, 2])
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal("0")

        
def test_sort_creates_circular_reference(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "=C1")
    wb.set_cell_contents("Sheet1", "A2", "=0")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    wb.sort_region("Sheet1", "A1", "B2", [2])
    # Check if the circular reference error is correctly identified
    assert isinstance(wb.get_cell_value("Sheet1", "A2"), CellError) and wb.get_cell_value("Sheet1", "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
    

def test_sort_one_row_and_col(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", "=C3")
    wb.sort_region("Sheet1", "B1", "B1", [1])
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal("1")


def test_sort_on_second_row(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", "1")
    wb.set_cell_contents("Sheet1", "C3", "0")
    wb.sort_region("Sheet1", "C2", "C3", [1])
    assert wb.get_cell_value("Sheet1", "C3") == decimal.Decimal("1")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal("1")


def test_equal_values(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", "1")
    wb.set_cell_contents("Sheet1", "C3", "0")
    wb.sort_region("Sheet1", "B1", "C3", [1])
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal("2")
    assert wb.get_cell_value("Sheet1", "B1") is None
    assert wb.get_cell_value("Sheet1", "C1") == decimal.Decimal("0")


def test_extra_cell_values(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", "1")
    wb.set_cell_contents("Sheet1", "C3", "0")
    wb.sort_region("Sheet1", "B1", "C3", [1])
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal("2")
    assert wb.get_cell_value("Sheet1", "B1") is None
    assert wb.get_cell_value("Sheet1", "C1") == decimal.Decimal("0")


def test_equal_values(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", "1")
    wb.set_cell_contents("Sheet1", "C3", "0")
    wb.sort_region("Sheet1", "A1", "C2", [1])
    assert wb.get_cell_value("Sheet1", "A1") == decimal.Decimal("2")
    assert wb.get_cell_value("Sheet1", "B1") == decimal.Decimal("1")
    assert wb.get_cell_value("Sheet1", "C1") == decimal.Decimal("3")


def test_bad_to_good_ref_sort(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", "=C3")
    assert wb.get_cell_value("Sheet1", "C2") == decimal.Decimal("0")
    wb.sort_region("Sheet1", "B1", "C2", [1])
    assert wb.get_cell_value("Sheet1", "C1") == decimal.Decimal("3")
    

def test_good_to_bad_ref_sort(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "2")
    wb.set_cell_contents("Sheet1", "A2", "2")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "3")
    wb.set_cell_contents("Sheet1", "C2", "=A1")
    wb.sort_region("Sheet1", "B1", "C2", [1])
    assert isinstance(wb.get_cell_value("Sheet1", "C1"), CellError) and wb.get_cell_value("Sheet1", "C1").get_type() == CellErrorType.BAD_REFERENCE


def test_none_sort(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "A2", "")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    wb.sort_region("Sheet1", "A1", "B2", [1])
    assert wb.get_cell_value("Sheet1", "A1") is None and wb.get_cell_value("Sheet1", "B1") == decimal.Decimal('1')

import unittest

def test_duplicate_sort_cols(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "A2", "")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    try:
        wb.sort_region("Sheet1", "A1", "B2", [1, -1])
    except ValueError:
        assert True
    else:
        assert False

    
def test_sort_cols_out_of_range(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "A2", "")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    try:
        wb.sort_region("Sheet1", "A1", "B2", [3])
    except ValueError:
        assert True
    else:
        assert False


def test_key_error(wb: Workbook):
    wb.set_cell_contents("Sheet1", "A1", "")
    wb.set_cell_contents("Sheet1", "A2", "")
    wb.set_cell_contents("Sheet1", "B1", "1")
    wb.set_cell_contents("Sheet1", "B2", "0")
    wb.set_cell_contents("Sheet1", "C1", "")
    wb.set_cell_contents("Sheet1", "C2", "=A2")
    try:
        wb.sort_region("Sheet3", "A1", "B2", [3])
    except KeyError:
        assert True
    else:
        assert False


# check continuing if nothing is less for one of the columns
    
# Test two cycles and comparing two cycles 
    
# Test multiple sort cols 

# Test errors
    
# test multiple jumps
    
# test large scale
    


