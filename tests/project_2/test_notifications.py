
import context
import pytest
import sheets

def test_cell_notifs():
    wb = sheets.Workbook()
    def func(workbook, changed_cells):
        print(changed_cells)
        print("num changed cells", len(changed_cells))

    wb.notify_cells_changed(func)
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", "A1", "10")
    wb.set_cell_contents("Sheet1", "A2", "=A1")
    wb.set_cell_contents("Sheet1", "A3", "=A1")

    wb.set_cell_contents("Sheet1", "A2", "")
    wb.set_cell_contents("Sheet1", "A3", "")

    wb.set_cell_contents("Sheet1", "A1", "20")
    assert 0==0

# Test Fixture to setup a Workbook with notification function
@pytest.fixture
def workbook_with_notifier():
    wb = sheets.Workbook()
    wb.new_sheet()
    wb.notify_cells_changed(on_cells_changed)
    return wb

# Mock notification function to capture cell changes
changed_cells = []
def on_cells_changed(workbook, cells):
    changed_cells.extend(cells)

# Test for basic cell value change
def test_basic_cell_change(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
    assert changed_cells == [("Sheet1", "a1")]

# Test for cell change with formula
def test_cell_change_with_formula(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "123")
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "=A1")
    assert ("Sheet1", "b1") in changed_cells

# Test for multiple cell changes
def test_multiple_cell_changes(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "=A1")
    workbook_with_notifier.set_cell_contents("Sheet1", "C1", "=A1+B1")
    assert set(changed_cells) == {("Sheet1", "a1"), ("Sheet1", "b1"), ("Sheet1", "c1")}

# Test for cell change back to original value
def test_cell_change_back_to_original(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'456")
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
    assert changed_cells.count(("Sheet1", "a1")) >= 1

# Test for no cell change notification if value remains same
def test_no_change_no_notification(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
    assert not changed_cells

# Test for changes in cells due to sheet operations (add, rename, delete)
# Add more tests here depending on how these operations affect cell values

# Test for exception handling in notification function
# def test_exception_handling_in_notifier(workbook_with_notifier):
#     # Redefine notifier to throw an exception
#     def notifier_throws(workbook, cells):
#         raise Exception("Test exception")

#     workbook_with_notifier.notify_cells_changed(notifier_throws)
#     try:
#         workbook_with_notifier.set_cell_contents("Sheet1", "A1", "'123")
#     except Exception:
#         pytest.fail("Exception in notifier should not propagate")

#     # Restore original notifier
#     workbook_with_notifier.notify_cells_changed(on_cells_changed)

# Test for cell value change due to sheet renaming
def test_cell_change_due_to_sheet_rename(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.new_sheet("Sheet2")
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=Sheet3!B1")
    changed_cells.clear()
    workbook_with_notifier.rename_sheet("Sheet2", "Sheet3")
    assert ("Sheet1", "a1") in changed_cells

# Test for cell value change due to sheet deletion
def test_cell_change_due_to_sheet_deletion(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.new_sheet("Sheet3")
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=Sheet3!B1")
    changed_cells.clear()
    workbook_with_notifier.del_sheet("Sheet3")
    assert ("Sheet1", "a1") in changed_cells

# Test for cell value change due to sheet addition
def test_cell_change_due_to_sheet_addition(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=Sheet3!B1")
    changed_cells.clear()
    workbook_with_notifier.new_sheet("Sheet3")
    assert ("Sheet1", "a1") in changed_cells

# Test for cell value change due to sheet copy
def test_cell_change_due_to_sheet_copy(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.new_sheet("Sheet2")
    workbook_with_notifier.set_cell_contents("Sheet2", "B1", "'value")
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=Sheet2_1!B1")
    changed_cells.clear()
    workbook_with_notifier.copy_sheet("Sheet2")
    assert ("Sheet1", "a1") in changed_cells


# Test for cell value change due to sheet copy
def test_cell_change_due_to_sheet_cop_2(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.new_sheet("Sheet2")
    workbook_with_notifier.set_cell_contents("Sheet2", "B1", "'value")
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=Sheet2_1!B1")
    changed_cells.clear()
    workbook_with_notifier.copy_sheet("Sheet2")
    assert ("Sheet2_1", "b1") in changed_cells

# More tests can be added to test for more complex sheet operations,
# and their impact on cells with formula references.

# Test for cell value change due to circular reference
def test_cell_change_due_to_circular_reference(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=B1")
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "=A1")
    # Expecting both A1 and B1 to be reported as changed due to circular reference
    assert set(changed_cells) == {("Sheet1", "a1"), ("Sheet1", "b1")}

# Test for cell value change due to resolving circular reference
def test_resolve_circular_reference(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=B1")
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "=A1")
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "B1", "'5")
    # Expecting A1 and B1 to be reported as changed after resolving circular reference
    print(changed_cells)

    assert set(changed_cells) == {("Sheet1", "a1"), ("Sheet1", "b1")}

# Test for cell value change due to formula parsing error
def test_cell_change_due_to_formula_parsing_error(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=SUM(A2:A3")
    # Expecting A1 to be reported as changed due to formula parsing error
    assert ("Sheet1", "a1") in changed_cells

# Test for cell value change due to resolving formula parsing error
def test_resolve_formula_parsing_error(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=SUM(A2:A3")
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=SUM(A2:A3)")
    # Expecting A1 to be reported as changed after resolving formula parsing error
    assert ("Sheet1", "a1") in changed_cells

def test_cell_notifs_bad_ref(workbook_with_notifier):
    changed_cells.clear()
    workbook_with_notifier.set_cell_contents("Sheet1", "A1", "=Sheet2!A1")
    assert ("Sheet1", "a1") in changed_cells
    changed_cells.clear()
    workbook_with_notifier.new_sheet("Sheet2")

    assert ("Sheet1", "a1") in changed_cells
    changed_cells.clear()

    workbook_with_notifier.del_sheet("Sheet2")
    assert ("Sheet1", "a1") in changed_cells
    changed_cells.clear()

# Additional tests can be added to cover other types of complex scenarios
# such as reference to non-existing cells, type errors in formulas, etc.
