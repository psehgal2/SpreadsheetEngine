import pytest

from sheets.Workbook import Workbook
from decimal import Decimal
from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType

# Test 1: Test moving a single cell with a literal value
def test_move_single_cell_literal_value():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', 'Test')
    wb.move_cells('Sheet1', 'A1', 'A1', 'B2')
    assert wb.get_cell_contents('Sheet1', 'B2') == 'Test', "The literal value should be moved correctly"

# Test 2: Test copying a cell with a formula that includes relative references
def test_copy_cell_formula_relative_references():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '2')
    wb.set_cell_contents('Sheet1', 'B1', '3')
    wb.set_cell_contents('Sheet1', 'C1', '=A1+B1')
    wb.copy_cells('Sheet1', 'C1', 'C1', 'C2')
    assert wb.get_cell_contents('Sheet1', 'C2') == '=A2+B2', "The formula should update with relative references"

# Test 3: Test moving a range of cells with mixed content
def test_move_range_cells_mixed_content():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', 'Hello')
    wb.set_cell_contents('Sheet1', 'A2', '=A1')
    wb.move_cells('Sheet1', 'A1', 'A2', 'B1')
    assert wb.get_cell_contents('Sheet1', 'B1') == 'Hello', "The string should be moved correctly"
    assert wb.get_cell_contents('Sheet1', 'B2') == '=B1', "The formula should update to new location"

# Test 4: Test copying cells with an absolute reference in a formula
def test_copy_cells_absolute_reference():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=B$1+1')
    wb.set_cell_contents('Sheet1', 'B1', '10')
    wb.copy_cells('Sheet1', 'A1', 'A1', 'A2')
    assert wb.get_cell_contents('Sheet1', 'A2') == '=B$1+1', "Absolute row references should not change on copy"

# Test 5: Test moving cells where the destination overlaps the source
def test_move_cells_overlap_source_destination():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '1')
    wb.set_cell_contents('Sheet1', 'A2', '2')
    wb.move_cells('Sheet1', 'A1', 'A2', 'A2')
    assert wb.get_cell_contents('Sheet1', 'A2') == '1', "First cell should be moved to the second cell's position"
    assert wb.get_cell_contents('Sheet1', 'A3') == '2', "Second cell should be moved down due to overlap"

# Test 6: Test moving cells with a formula resulting in a CellError due to out-of-bounds reference
def test_move_cells_formula_resulting_in_cell_error():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'B1', '=A1')
    wb.move_cells('Sheet1', 'B1', 'B1', 'A1')
    assert isinstance(wb.get_cell_value('Sheet1', 'A1'), CellError), "Moving should result in a CellError for out-of-bounds references"

# Test 7: Test copying a cell that has a circular reference error but doesn't now
def test_copy_cells_circular_reference_error():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=B1')
    wb.set_cell_contents('Sheet1', 'B1', '=A1')

    wb.copy_cells('Sheet1', 'A1', 'A1', 'C1')
    error = wb.get_cell_value('Sheet1', 'C1')
    assert not isinstance(error, CellError)
    assert  wb.get_cell_value('Sheet1', 'A1').get_type() == CellErrorType.CIRCULAR_REFERENCE
    assert  wb.get_cell_value('Sheet1', 'B1').get_type() == CellErrorType.CIRCULAR_REFERENCE



# Test 8: Test error handling when moving cells to an invalid location
def test_move_cells_to_invalid_location():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', 'Valid Content')
    with pytest.raises(ValueError):
        wb.move_cells('Sheet1', 'A1', 'A1', 'InvalidLocation')

# Test 9: Check for #REF! error when moving a cell reference out of bounds
def test_move_out_of_bounds_reference_error():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'ZZZZ9998', '= ZZZZ9999')
    wb.move_cells('Sheet1', 'ZZZZ9998', 'ZZZZ9998', 'ZZZZ9999')
    cell_contents = wb.get_cell_contents('Sheet1', 'ZZZZ9999')
    assert cell_contents == '=#REF!', "Moving a cell out of bounds should update the formula to a #REF! error"

# Test 10: Test that copying cells updates relative references but preserves absolute references
def test_copy_cells_preserve_absolute_update_relative():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=B1+C$1')
    wb.set_cell_contents('Sheet1', 'B1', '10')
    wb.set_cell_contents('Sheet1', 'C1', '20')
    wb.copy_cells('Sheet1', 'A1', 'A1', 'A2')
    expected_formula = '=B2+C$1'  # B1 should change to B2, C$1 remains unchanged
    actual_formula = wb.get_cell_contents('Sheet1', 'A2')
    assert actual_formula == expected_formula, "Copying should update relative references and preserve absolute references"

# Test 11: Test moving a cell with a division by zero error
def test_move_cell_with_division_by_zero_error():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=1/0')
    wb.move_cells('Sheet1', 'A1', 'A1', 'B1')
    error = wb.get_cell_value('Sheet1', 'B1')
    assert isinstance(error, CellError) and error.get_type() == CellErrorType.DIVIDE_BY_ZERO, "Moving a cell with a division by zero should result in a CellError with DIVIDE_BY_ZERO type"

# Test 12: Test moving cells with formulas that reference each other correctly updates references
def test_move_cells_with_interdependent_formulas():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=B1+1')
    wb.set_cell_contents('Sheet1', 'B1', '=A1+2')
    wb.move_cells('Sheet1', 'A1', 'B1', 'C1')
    assert wb.get_cell_contents('Sheet1', 'C1') == '=D1+1' and wb.get_cell_contents('Sheet1', 'D1') == '=C1+2', "Moving cells with interdependent formulas should correctly update references"

# Test 13: Ensure moving a cell updates its formula references but does not affect cells referencing it
def test_move_cell_updates_formula_references_but_not_reverse():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '5')
    wb.set_cell_contents('Sheet1', 'B1', '=A1*2')
    wb.move_cells('Sheet1', 'A1', 'A1', 'A2')
    assert wb.get_cell_contents('Sheet1', 'B1') == '=A1*2', "Moving a cell should not update formulas in cells that reference the moved cell"

# Test 14: Test moving cells with mixed references updates only the relative part of the reference
def test_move_cells_with_mixed_references():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=B$1+C1')
    wb.move_cells('Sheet1', 'A1', 'A1', 'A2')
    expected_formula_after_move = '=B$1+C2'  # C1 should change to C2, B$1 remains unchanged
    actual_formula = wb.get_cell_contents('Sheet1', 'A2')
    assert actual_formula == expected_formula_after_move, "Moving cells with mixed references should update only the relative part of the reference"

# Test 15: Test moving a block of cells where some formulas become out of bounds
def test_move_block_cells_formulas_become_out_of_bounds():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=B1')
    wb.set_cell_contents('Sheet1', 'A2', '=B2')
    wb.move_cells('Sheet1', 'A1', 'A2', 'ZZZZ998')  # Assuming ZZ999 is the last valid cell
    assert wb.get_cell_contents('Sheet1', 'ZZZZ998') == '=#REF!', "Formula should become a #REF! error when moved out of bounds"
    assert wb.get_cell_contents('Sheet1', 'ZZZZ999') == '=#REF!', "Formula should become a #REF! error when moved out of bounds"

# Test 16: Test moving cells to a location that causes a circular reference due to relative reference updates
def test_move_cells_circular_reference_due_to_update():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=1')
    wb.set_cell_contents('Sheet1', 'A2', '=C2')
    wb.set_cell_contents('Sheet1', 'C2', '1')
    wb.set_cell_contents('Sheet1', 'C1', '=A1+1')
    wb.move_cells('Sheet1', 'C1', 'C1', 'C2')
    error = wb.get_cell_value('Sheet1', 'C2')
    assert isinstance(error, CellError) and error.get_type() == CellErrorType.CIRCULAR_REFERENCE, "Moving should create a circular reference error due to relative reference updates"
    error = wb.get_cell_value('Sheet1', 'A2')
    assert isinstance(error, CellError) and error.get_type() == CellErrorType.CIRCULAR_REFERENCE, "Moving should create a circular reference error due to relative reference updates"

# Test 17: Test moving a range that includes a cell with an error and a cell that depends on it
def test_move_range_including_error_and_dependent_cell():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=1/0')  # This will cause a DIV/0 error
    wb.set_cell_contents('Sheet1', 'B1', '=A1+2')
    wb.move_cells('Sheet1', 'A1', 'B1', 'A2')
    assert isinstance(wb.get_cell_value('Sheet1', 'A2'), CellError), "Moved cell A1 should still have DIV/0 error"
    assert isinstance(wb.get_cell_value('Sheet1', 'B2'), CellError), "B2 should have an error because it depends on A2 which has an error"

# Test 18: Ensure moving cells does not alter unaffected formulas in the workbook
def test_move_cells_unaffected_formulas_remain_unchanged():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '10')
    wb.set_cell_contents('Sheet1', 'A2', '=A1*2')
    wb.set_cell_contents('Sheet1', 'B1', '=A2+1')
    wb.move_cells('Sheet1', 'A1', 'A1', 'C1')
    assert wb.get_cell_contents('Sheet1', 'B1') == '=A2+1', "Formula in B1 should remain unchanged after moving A1"

# Test 19: Test copying a range of cells that includes both formulas and literals to a new sheet
def test_copy_range_to_new_sheet_including_formulas_and_literals():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.set_cell_contents('Sheet1', 'A1', '100')
    wb.set_cell_contents('Sheet1', 'A2', '=A1*2')
    wb.copy_cells('Sheet1', 'A1', 'A2', 'B1', 'Sheet2')
    assert wb.get_cell_contents('Sheet2', 'B1') == '100', "Literal value should be copied correctly to new sheet"
    assert wb.get_cell_contents('Sheet2', 'B2') == '=B1*2', "Formula should be copied and updated correctly to new sheet"

# Test 20: Validate that moving a range where the destination overlaps the source handles formulas correctly
def test_move_range_with_overlap_handles_formulas_correctly():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '1')
    wb.set_cell_contents('Sheet1', 'B1', '=A1+1')
    wb.set_cell_contents('Sheet1', 'C1', '=B1+1')
    wb.move_cells('Sheet1', 'A1', 'B1', 'B1')
    assert wb.get_cell_contents('Sheet1', 'B1') == '1', "A1 should have been moved to B1 correctly"
    assert wb.get_cell_contents('Sheet1', 'C1') == '=B1+1', "C1's formula should be overwritten"

# Test 21: Test copying cells with a complex formula involving multiple operations and references
def test_copy_cells_with_complex_formula():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '10')
    wb.set_cell_contents('Sheet1', 'B1', '20')
    wb.set_cell_contents('Sheet1', 'C1', '=A1+B1*2-(B1/A1)')
    wb.copy_cells('Sheet1', 'C1', 'C1', 'C2')
    assert wb.get_cell_contents('Sheet1', 'C2') == '=A2+B2*2-(B2/A2)', "The complex formula should be copied with relative references updated"

# Test 22: Ensure that moving a range of cells updates formulas within the range but keeps external references correct
def test_move_range_updates_internal_but_not_external_references():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '=D1*2')
    wb.set_cell_contents('Sheet1', 'B1', '=A1+2')
    wb.set_cell_contents('Sheet1', 'C1', '=B1+A1')
    wb.set_cell_contents('Sheet1', 'D1', '5')
    wb.move_cells('Sheet1', 'A1', 'C1', 'A2')
    assert wb.get_cell_contents('Sheet1', 'A2') == '=D2*2', "Internal reference in moved formula should update"
    assert wb.get_cell_contents('Sheet1', 'B2') == '=A2+2', "Internal reference in moved formula should update"
    assert wb.get_cell_contents('Sheet1', 'C2') == '=B2+A2', "Internal reference in moved formula should update"
    assert wb.get_cell_contents('Sheet1', 'D1') == '5', "External reference should remain unchanged"

# Test 23: Test copying a cell with a formula that references another sheet
# NOT ACTUALLY A VALID TEST. COPYING CELLS THAT REFERENCE OTHER SHEETS SHOULD STILL UPDATE STUFF.
# def test_copy_cell_formula_referencing_another_sheet():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     wb.new_sheet('Sheet2')
#     wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1*2')
#     wb.set_cell_contents('Sheet2', 'A1', '5')
#     wb.copy_cells('Sheet1', 'A1', 'A1', 'B1')
#     assert wb.get_cell_contents('Sheet1', 'B1') == '=Sheet2!A1*2', "Formula referencing another sheet should be copied without change"

# Test 24: Validate handling of a move operation that results in a #REF! error in a formula due to out-of-bounds move
# TEST IS BROKEN AND REDUNDANT.
# def test_move_operation_causing_ref_error_in_formula():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     wb.set_cell_contents('Sheet1', 'A1', '=B1+1')
#     wb.set_cell_contents('Sheet1', 'B1', '10')
#     wb.move_cells('Sheet1', 'B1', 'B1', 'ZZZ9999')  # Assuming ZZZ9999 is the last valid cell
#     wb.move_cells('Sheet1', 'A1', 'A1', 'C1')
#     assert '=#REF!+1' == wb.get_cell_contents('Sheet1', 'C1'), "Formula should have a #REF! error after moving B1 out of bounds"

# Test 25: Test that a series of moves and copies maintains correct formula relationships across multiple sheets
# test incorrect. Fix later
def test_series_of_moves_and_copies_across_sheets_maintains_formulas():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.set_cell_contents('Sheet1', 'A1', '100')
    wb.set_cell_contents('Sheet1', 'A2', '=A1*2')
    wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A2*2+10')
    wb.copy_cells('Sheet1', 'A1', 'A2', 'B1', 'Sheet2')
    wb.move_cells('Sheet2', 'A1', 'A1', 'B2')
    assert wb.get_cell_contents('Sheet2', 'B1') == '100', "Literal value should be copied correctly"
    assert wb.get_cell_contents('Sheet2', 'B2') == '=Sheet1!B3*2+10', "Formula should be updated correctly after a series of moves and copies, maintaining relationships"

# Test 31: Ensure moving a cell updates dependent formulas across sheets
def test_move_cell_updates_dependent_formulas_across_sheets():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.set_cell_contents('Sheet1', 'A1', '5')
    wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1*2')
    wb.move_cells('Sheet1', 'A1', 'A1', 'B1')
    assert wb.get_cell_value('Sheet2', 'A1') == 0, "Dependent formula on another sheet should update correctly after move"

# Test 32: Test copying cells with a formula that references a now-deleted sheet
def test_copy_cells_formula_referencing_deleted_sheet():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1+1')
    wb.del_sheet('Sheet2')
    wb.copy_cells('Sheet1', 'A1', 'A1', 'B1')
    error = wb.get_cell_value('Sheet1', 'B1')
    assert isinstance(error, CellError) and error.get_type() == CellErrorType.BAD_REFERENCE, "Formula referencing a deleted sheet should result in a BAD_REFERENCE error"

# Test 33: Validate moving a range of cells updates named ranges within formulas
# Assuming the spreadsheet engine supports named ranges
# def test_move_cells_updates_named_ranges_in_formulas():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     # Assuming set_named_range is a method to define a named range, for the purpose of this test
#     wb.set_named_range('DataRange', 'Sheet1', 'A1:A2')
#     wb.set_cell_contents('Sheet1', 'B1', '=SUM(DataRange)')
#     wb.set_cell_contents('Sheet1', 'A1', '2')
#     wb.set_cell_contents('Sheet1', 'A2', '3')
#     wb.move_cells('Sheet1', 'A1', 'A2', 'A2')
#     # Assuming get_cell_value also evaluates named ranges within formulas
#     assert wb.get_cell_value('Sheet1', 'B1') == 5, "Formula using a named range should update and evaluate correctly after moving the cells within the range"

# Test 34: Check behavior when copying cells to a range that would cause an overlap in destination sheet
def test_copy_cells_to_overlap_in_destination_sheet():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', '1')
    wb.set_cell_contents('Sheet1', 'A2', '2')
    wb.copy_cells('Sheet1', 'A1', 'A2', 'A2')
    # Assuming the engine handles overlaps by first copying to a temporary location or similar logic
    assert wb.get_cell_value('Sheet1', 'A2') == 1 and wb.get_cell_value('Sheet1', 'A3') == 2, "Copying cells to an overlapping range should handle overlap correctly"

# Test 35: Ensure moving cells with formulas referencing multiple sheets updates references correctly
def test_move_cells_with_formulas_referencing_multiple_sheets():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.new_sheet('Sheet3')
    wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1 + Sheet3!A1')
    wb.set_cell_contents('Sheet2', 'B1', '10')
    wb.set_cell_contents('Sheet3', 'B1', '20')
    wb.move_cells('Sheet1', 'A1', 'A1', 'B1')
    # Check that the formula in B1 correctly references the values in Sheet2!A1 and Sheet3!A1
    assert wb.get_cell_value('Sheet1', 'B1') == 30, "Formula referencing multiple sheets should update and evaluate correctly after move"

# Test 36: Comprehensive test for moving cells with mixed references, error handling, and cross-sheet updates
# def test_comprehensive_move_with_mixed_references_errors_and_cross_sheet_updates():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     wb.new_sheet('Sheet2')
#     wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!B1+10')
#     wb.set_cell_contents('Sheet1', 'A2', '=$A$1+1')
#     wb.set_cell_contents('Sheet2', 'B1', '20')
#     wb.set_cell_contents('Sheet2', 'B2', '=NONEXISTENT()')
#     wb.move_cells('Sheet1', 'A1', 'A2', 'C1')
#     wb.move_cells('Sheet2', 'B1', 'B2', 'C1')
#     assert wb.get_cell_value('Sheet1', 'C1') == 30, "Cross-sheet reference should update correctly after move"
#     assert isinstance(wb.get_cell_value('Sheet2', 'C2'), CellError) and wb.get_cell_value('Sheet2', 'C2').get_type() == CellErrorType.BAD_NAME, "Function error should persist after move"
#     assert wb.get_cell_contents('Sheet1', 'C2') == '=$C$1+1', "Absolute reference in formula should remain unchanged after move"
#
# # Test 37: Comprehensive test for copying cells involving formulas, absolute and relative references, and error propagation
# def test_comprehensive_copy_with_formulas_references_and_errors():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     wb.set_cell_contents('Sheet1', 'A1', '=B1+C1')
#     wb.set_cell_contents('Sheet1', 'B1', '5')
#     wb.set_cell_contents('Sheet1', 'C1', '=D1/0')  # This will cause a DIV/0 error
#     wb.set_cell_contents('Sheet1', 'D1', '10')
#     wb.copy_cells('Sheet1', 'A1', 'C1', 'A2')
#     assert wb.get_cell_contents('Sheet1', 'A2') == '=B2+C2', "Formula should be copied with relative references updated"
#     assert wb.get_cell_value('Sheet1', 'C2') is not None and isinstance(wb.get_cell_value('Sheet1', 'C2'), CellError), "Error should be copied correctly"
#     assert wb.get_cell_value('Sheet1', 'A2') is not None and isinstance(wb.get_cell_value('Sheet1', 'A2'), CellError), "Resultant cell should inherit error from copied formula"
#
# # Test 38: Comprehensive test for operations involving cell deletion, formula updates, and error checks
# def test_operations_with_deletion_formula_updates_and_error_checks():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     wb.set_cell_contents('Sheet1', 'A1', '10')
#     wb.set_cell_contents('Sheet1', 'B1', '=A1*2')
#     wb.set_cell_contents('Sheet1', 'C1', '=B1+10')
#     wb.del_sheet('Sheet1')
#     wb.new_sheet('Sheet1')
#     assert wb.num_sheets() == 1, "Sheet deletion and recreation should be handled correctly"
#     wb.set_cell_contents('Sheet1', 'A1', '=B1*2')
#     error = wb.get_cell_value('Sheet1', 'A1')
#     assert isinstance(error, CellError) and error.get_type() == CellErrorType.BAD_REFERENCE, "Reference to deleted cell should result in an error"
#
# # Test 39: Comprehensive test involving moving cells, circular references detection, and cross-sheet formula evaluation
# def test_move_with_circular_references_and_cross_sheet_evaluation():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     wb.new_sheet('Sheet2')
#     wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1+1')
#     wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1+1')
#     wb.move_cells('Sheet2', 'A1', 'A1', 'B1')
#     circular_error_sheet1 = wb.get_cell_value('Sheet1', 'A1')
#     circular_error_sheet2 = wb.get_cell_value('Sheet2', 'B1')
#     assert isinstance(circular_error_sheet1, CellError) and circular_error_sheet1.get_type() == CellErrorType.CIRCULAR_REFERENCE, "Circular reference should be detected in Sheet1"
#     assert isinstance(circular_error_sheet2, CellError) and circular_error_sheet2.get_type() == CellErrorType.CIRCULAR_REFERENCE, "Circular reference update should be detected in Sheet2 after move"
#
# # Test 40: Comprehensive edge case test for formula updates, error literals, and sheet reordering impacts
# def test_formula_updates_error_literals_and_sheet_reordering_impacts():
#     wb = Workbook()
#     wb.new_sheet('Sheet1')
#     wb.new_sheet('Sheet2')
#     wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!B1*2')
#     wb.set_cell_contents('Sheet2', 'B1', '5')
#     wb.move_cells('Sheet1', 'A1', 'A1', 'A2')
#     wb.del_sheet('Sheet2')
#     wb.new_sheet('Sheet2')
#     wb.set_cell_contents('Sheet2', 'B1', '10')
#     wb.move_cells('Sheet1', 'A2', 'A2', 'A3')
#     error = wb.get_cell_value('Sheet1', 'A3')
#     assert isinstance(error, CellError) and error.get_type() == CellErrorType.BAD_REFERENCE, "Reference to a cell in a deleted and recreated sheet should result in a BAD_REFERENCE error"
#     assert wb.get_cell_value('Sheet2', 'B1') == 10, "Newly added data in a recreated sheet should be correctly evaluated"
#
#
#
