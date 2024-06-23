from typing import Any, List, Optional, Tuple, Callable, Iterable
from sheets.utils import (
    check_new_name,
    get_hidden_name,
    check_valid_cell_location,
    is_formula,
    convert_location_to_idx,
    convert_idx_to_location,
)
from sheets.Sheet import Sheet
from sheets.Cell import Cell
from sheets.Parser import FormulaEvaluator, parser
from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType
from sheets.Parser import FormulaFixer, MoveFormula
from sheets.Graph import Graph
from sheets.Row import Row
from sheets.Functions import function_directory
import logging
import copy

import collections
import lark
from lark.reconstruct import Reconstructor
import json
from decimal import Decimal


def notify_cell_changes(function: object) -> object:
    """A decorator which handles error propogation"""

    # Error Presedence is such that: An error will be propogated if its enum value
    # is higher in the CellErrorType than any other error in values.
    # If all errors in value are of equivalent type, then the left most Error (in terms of
    # the formula) will be propogated.
    def run_func(self, *args, **kwargs):
        result = function(self, *args, **kwargs)
        for func in self.cell_notification_functions:
            try:
                func(self, set(self.cells_changed.keys()))
            except Exception as e:
                logging.error("Error occurred:", e)
                continue
        self.cells_changed = {}
        return result

    return run_func


class Workbook:
    """A workbook containing zero or more named spreadsheets."""

    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    def __init__(self):
        """Initialize a new empty workbook."""
        # self.num_sheets = 0
        self.sheets = (
            collections.OrderedDict()
        )  # {sheet_name (sheet.hidden_name): Sheet}
        self.dependencies = {}  # (location : cell)
        # path = os.path.join(os.getcwd(),"CS130-Wi24", "sheets\\formulas.lark") if "CS130-Wi24" not in os.getcwd() else os.path.join(os.getcwd(),"sheets\\formulas.lark") # noqa: E501
        # Make a parser global variable once the module is loaded and then just call the global
        # WB can delegate the responsbalities to the parser
        self.cells_changed = {}  # set of tuples (sheet_name, location, oldest contents)
        self.cell_notification_functions = []  # : Callable[["Workbook", Iterable[Tuple[str, str]]], None]

        self.json = None

        self.graph = Graph()
        self.formula_functions = function_directory

    @staticmethod
    def load_workbook(fp):
        """This is a static method (not an instance method) to load a workbook
        from a text file or file-like object in JSON format, and return the
        new Workbook instance.  Note that the _caller_ of this function is
        expected to have opened the file; this function merely reads the file.

        If the contents of the input cannot be parsed by the Python json
        module then a json.JSONDecodeError should be raised by the method.
        (Just let the json module's exceptions propagate through.)  Similarly,
        if an IO read error occurs (unlikely but possible), let any raised
        exception propagate through.

        If any expected value in the input JSON is missing (e.g. a sheet
        object doesn't have the "cell-contents" key), raise a KeyError with
        a suitably descriptive message.

        If any expected value in the input JSON is not of the proper type
        (e.g. an object instead of a list, or a number instead of a string),
        raise a TypeError with a suitably descriptive message."""

        ## CLEAN UP. No need to catch the exceptions just to reraise them

        fp.seek(0)
        val = fp.read()
        workbook = json.loads(val)

        if not isinstance(workbook, dict):
            raise TypeError("Workbook is not a dictionary")
        sheets = workbook["sheets"]

        if not isinstance(sheets, list):
            raise TypeError("Sheets is not a list")

        wb = Workbook()

        # need for stuff to get updated
        for sheet in sheets:
            if not isinstance(sheet, dict):
                raise TypeError("Sheet is not a dictionary")

            if "name" not in sheet:
                raise KeyError("Name key not present in sheet: ", sheet)
            else:
                sheet_name = sheet["name"]

            if not isinstance(sheet_name, str):
                raise TypeError("Sheet name is not a string")

            if "cell-contents" not in sheet:
                raise KeyError("Cell contents not present in sheet: ", sheet)
            else:
                cell_contents = sheet["cell-contents"]

            if not isinstance(cell_contents, dict):
                raise TypeError("Cell contents is not a dictionary")

            # may be a problem?
            wb.new_sheet(sheet_name)
            for loc in cell_contents:
                if not isinstance(loc, str):
                    raise TypeError("Cell location is not a string")
                if not isinstance(cell_contents[loc], str):
                    raise TypeError("Cell value is not a string")
                wb.set_cell_contents(sheet_name, loc, cell_contents[loc])
        return wb

    def sort_region(self, sheet_name: str, start_location: str, end_location: str, sort_cols: List[int]):
        # Sort the specified region of a spreadsheet with a stable sort, using
        # the specified columns for the comparison.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be sorted.  Both corners are included in the
        # area being sorted; for example, sorting the region including cells B3
        # to J12 would be done by specifying start_location="B3" and
        # end_location="J12".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to sort, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to sort.
        #
        # The sort_cols argument specifies one or more columns to sort on.  Each
        # element in the list is the one-based index of a column in the region,
        # with 1 being the leftmost column in the region.  A column's index in
        # this list may be positive to sort in ascending order, or negative to
        # sort in descending order.  For example, to sort the region B3..J12 on
        # the first two columns, but with the second column in descending order,
        # one would specify sort_cols=[1, -2].
        #
        # The sorting implementation is a stable sort:  if two rows compare as
        # "equal" based on the sorting columns, then they will appear in the
        # final result in the same order as they are at the start.
        #
        # If multiple columns are specified, the behavior is as one would
        # expect:  the rows are ordered on the first column indicated in
        # sort_cols; when multiple rows have the same value for the first
        # column, they are then ordered on the second column indicated in
        # sort_cols; and so forth.
        #
        # No column may be specified twice in sort_cols; e.g. [1, 2, 1] or
        # [2, -2] are both invalid specifications.
        #
        # The sort_cols list may not be empty.  No index may be 0, or refer
        # beyond the right side of the region to be sorted.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        # If the sort_cols list is invalid in any way, a ValueError is raised.

        sheet_name = sheet_name.lower()
        if sheet_name not in self.sheets:
            raise KeyError
        
        if not check_valid_cell_location(start_location.lower()):
            raise ValueError("Start cell location not valid")
        
        if not check_valid_cell_location(end_location.lower()):
            raise ValueError("End cell location not valid")
        
        if len(sort_cols) <= 0:
            raise ValueError("Sort_cols can't be empty")
        
        cols = []
        for col in sort_cols:
            if col != 1 and col != 2 and col != 3 and col != 4 and col != 5 and col != 6 and col != 7 and col != 8 and col != 9 and col != -1 and col != -2 and col != -3 and col != -4 and col != -5 and col != -6 and col != -7 and col != -8 and col != -9:
                raise ValueError("Columns must be integers")
            col = abs(col)
            cols.append(col)
        cols_set = set(cols)

        if len(cols_set) < len(sort_cols):
            raise ValueError("Contains duplicate sort_cols values")
    
        start_row, start_col = convert_location_to_idx(start_location)
        end_row, end_col = convert_location_to_idx(end_location)
        top_row = min(start_row, end_row)
        bottom_row = max(start_row, end_row)
        left_col = min(start_col, end_col)
        right_col = max(start_col, end_col)

        for idx in sort_cols:
            if idx == 0 or idx > right_col - left_col + 1:
                raise ValueError

        # iterate 
        adapter_objs = []
        for row in range(top_row, bottom_row + 1):
            adapter_obj = Row(row, left_col, right_col, sort_cols)
            for col in sort_cols:
                col = abs(col)
                curr_loc = convert_idx_to_location(row, left_col + col - 1)
                cell_val = self.get_cell_value(sheet_name, curr_loc)
                adapter_obj.add_col(cell_val)
            adapter_objs.append(adapter_obj)
        adapter_objs.sort()

        cell_contents_to_copy = []
        for i in range(len(adapter_objs)):
            adapter_obj = adapter_objs[i]
            new_row = top_row + i
            old_row = adapter_obj.get_og_row()
            col_cells = []
            for col in range(left_col, right_col + 1):
                curr_loc = convert_idx_to_location(new_row, col)
                old_loc = convert_idx_to_location(old_row, col)
                col_cells.append(self.get_cell_contents(sheet_name, old_loc))
            # the last value of the list will just contain the delta shift
            col_cells.append(new_row-old_row)
            cell_contents_to_copy.append(col_cells)

        recon = Reconstructor(parser)
        for row in range(len(cell_contents_to_copy)):
            # minus one since have the delta row at the end
            delta_row = cell_contents_to_copy[row][len(cell_contents_to_copy[0]) - 1]
            for col in range(len(cell_contents_to_copy[0]) - 1):
                new_loc = convert_idx_to_location(top_row + row, left_col + col)
                contents = cell_contents_to_copy[row][col]
                contents = contents.lstrip().rstrip() if contents else contents
                if contents and contents[0] == "=":
                    parsed_formula = parser.parse(contents)
                    mf = MoveFormula(delta_row, 0)
                    moved_formula = mf.transform(parsed_formula)
                    fixed_formula = recon.reconstruct(
                        moved_formula, insert_spaces=False
                    )
                    self.set_cell_contents(sheet_name, new_loc, "=" + fixed_formula)
                elif contents:
                    self.set_cell_contents(sheet_name, new_loc, contents)
                else:
                    self.set_cell_contents(sheet_name, new_loc, None)

    
    def transfer_cells(
        self, sheet_name, start_location, end_location, to_location, to_sheet, isCopy
    ):
        if not to_sheet:
            to_sheet = sheet_name
        sheet_name = sheet_name.lower()
        to_sheet = to_sheet.lower()
        start_location = start_location.lower()
        end_location = end_location.lower()
        to_location = to_location.lower()

        if sheet_name not in self.sheets or to_sheet not in self.sheets:
            raise KeyError(f"{sheet_name} or {to_sheet} was not found in the Workbook.")

        if (
            not check_valid_cell_location(start_location)
            or not check_valid_cell_location(end_location)
            or not check_valid_cell_location(to_location)
        ):
            raise ValueError(
                f"{start_location} or {end_location} or {to_location} is not a valid cell location."
            )

        start_row, start_col = convert_location_to_idx(start_location)
        end_row, end_col = convert_location_to_idx(end_location)
        move_row, move_col = convert_location_to_idx(to_location)

        # Has not been handled:
        # If a formula being moved contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.

        # get cell at certain location
        # Fixes formulas for cells referencing the old sheet name
        top_row = min(start_row, end_row)
        bottom_row = max(start_row, end_row)
        left_col = min(start_col, end_col)
        right_col = max(start_col, end_col)

        shift_row = move_row - top_row
        shift_col = move_col - left_col

        extent_row = move_row + (bottom_row - top_row)
        extent_col = move_col + (right_col - left_col)
        extent_loc = convert_idx_to_location(extent_row, extent_col)
        if not check_valid_cell_location(extent_loc):
            raise ValueError(
                f"Move would put furthest extent of cell location to {extent_loc} which is not valid"
            )

        # Copies the cells
        # MAY NOT WORK WITH DEPENDENCIES GRAPH
        # Test with None values

        cell_contents_to_copy = []
        for row in range(top_row, bottom_row + 1):
            col_cells = []
            for col in range(left_col, right_col + 1):
                curr_loc = convert_idx_to_location(row, col)
                col_cells.append(self.get_cell_contents(sheet_name, curr_loc))
                if not isCopy:
                    self.set_cell_contents(sheet_name, curr_loc, None)
            cell_contents_to_copy.append(col_cells)

        recon = Reconstructor(parser)
        for row in range(len(cell_contents_to_copy)):
            for col in range(len(cell_contents_to_copy[0])):
                new_loc = convert_idx_to_location(row + move_row, col + move_col)
                contents = cell_contents_to_copy[row][col]
                contents = contents.lstrip().rstrip() if contents else contents
                if contents and contents[0] == "=":
                    parsed_formula = parser.parse(contents)
                    mf = MoveFormula(shift_row, shift_col)
                    moved_formula = mf.transform(parsed_formula)
                    fixed_formula = recon.reconstruct(
                        moved_formula, insert_spaces=False
                    )
                    self.set_cell_contents(to_sheet, new_loc, "=" + fixed_formula)
                elif contents:
                    self.set_cell_contents(to_sheet, new_loc, contents)
                else:
                    self.set_cell_contents(to_sheet, new_loc, None)

    def move_cells(
        self,
        sheet_name: str,
        start_location: str,
        end_location: str,
        to_location: str,
        to_sheet: Optional[str] = None,
    ) -> None:
        # Move cells from one location to another, possibly moving them to
        # another sheet.  All formulas in the area being moved will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) will
        # become empty due to the move operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be moved.  The to_location specifies the
        # top-left corner of the target area to move the cells to.
        #
        # Both corners are included in the area being moved; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to move, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to move.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being moved to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being moved contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        self.transfer_cells(
            sheet_name, start_location, end_location, to_location, to_sheet, False
        )

    def copy_cells(
        self,
        sheet_name: str,
        start_location: str,
        end_location: str,
        to_location: str,
        to_sheet: Optional[str] = None,
    ) -> None:
        # Copy cells from one location to another, possibly copying them to
        # another sheet.  All formulas in the area being copied will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) are
        # left unchanged by the copy operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be copied.  The to_location specifies the
        # top-left corner of the target area to copy the cells to.
        #
        # Both corners are included in the area being copied; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to copy, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to copy.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being copied to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being copied contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        self.transfer_cells(
            sheet_name, start_location, end_location, to_location, to_sheet, True
        )

    def save_workbook(self, fp):
        """Instance method (not a static/class method) to save a workbook to a
        text file or file-like object in JSON format.  Note that the _caller_
        of this function is expected to have opened the file; this function
        merely writes the file.

        If an IO write error occurs (unlikely but possible), let any raised
        exception propagate through."""

        workbook_dict = {"sheets": []}
        for hidden_name in self.sheets:
            # order may mess up
            sheet_dict = {}
            sheet = self.sheets[hidden_name]
            sheet_name = sheet.get_sheet_name()
            sheet_dict["name"] = sheet_name
            cell_dict = {}
            cells_dict = sheet.get_sheet_cells()
            for cell in cells_dict:
                content = sheet.get_cell_contents(cell)
                cell_dict[cell] = content
            # what to do if there are no cell contents
            sheet_dict["cell-contents"] = cell_dict
            workbook_dict["sheets"].append(sheet_dict)

        try:
            workbook_json = json.dumps(workbook_dict)
            fp.write(workbook_json)
            self.json = workbook_dict
        except TypeError as t:
            raise t("Unable to convert to JSON")

    def num_sheets(self) -> int:
        """Return the number of spreadsheets in the workbook."""
        return len(self.sheets.keys())  # DON'T need keys() here

    def list_sheets(self) -> List[str]:
        """Return a list of the spreadsheet names in the workbook, with the
        capitalization specified at creation, and in the order that the sheets
        appear within the workbook."""
        #
        # In this project, the sheet names appear in the order that the user
        # created them; later, when the user is able to move and copy sheets,
        # the ordering of the sheets in this function's result will also reflect
        # such operations.
        #
        # A user should be able to mutate the return-value without affecting the
        # workbook's internal state.

        return [val.sheet_name for val in self.sheets.values()]

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        """Return a tuple (num-cols, num-rows) indicating the current extent of
        the specified spreadsheet."""
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.

        hidden_name = get_hidden_name(sheet_name)
        if hidden_name in self.sheets:
            return self.sheets[hidden_name].get_extent()
        else:
            raise KeyError(f"{sheet_name} was not found in the Workbook.")
        # DUPLICATE THE KEY ERROR, can put into a function because its called a lot and the funct can handle the raising error

    def get_sheet_name_from_hidden(self, hidden_name: str) -> str:
        sheet = self.sheets[get_hidden_name(hidden_name)]
        return sheet.get_sheet_name()

    def get_formula_value(self, hidden_name: str, cell: Cell, force_recompute: bool):
        """Computes the value of the formula in the given cell.
        And also returns the cells referenced in the computation of the formula."""
        formula_evaluator = FormulaEvaluator(self, hidden_name, self.formula_functions)
        try:
            if not cell.get_cached_formula() or force_recompute:
                parsed_formula = parser.parse(cell.get_contents())
                cell.cache_formula(parsed_formula)
            else:
                parsed_formula = cell.get_cached_formula()

        except lark.exceptions.LarkError as E:
            formula_value = CellError(
                CellErrorType.PARSE_ERROR,
                f"The formula {cell.get_contents()} is invalid and generated a parser error.",
                E,
            )
            cell_refs = set({})
        else:
            formula_value = formula_evaluator.visit(parsed_formula)
            cell_refs = formula_evaluator.get_cell_refs()

        if formula_value is None:
            formula_value = Decimal(0)
        if cell_refs is None:
            cell_refs = set({})

        cell.set_children(cell_refs)

        return formula_value, cell_refs

    def get_cell_from_location(self, sheet_name, location):
        """Returns the cell from the location."""
        hidden_name = get_hidden_name(sheet_name)
        if hidden_name not in self.sheets:
            return None
        return self.sheets[hidden_name].get_cell(location)

    def recompute_cell_value(self, cell_loc):
        """Recomputes the value of the given cell based on its formula and the value of the dependencies."""
        # Recompute the value of the specified cell.

        sheet_name, location = cell_loc  # cell.get_location().lower().split("!")
        hidden_name = get_hidden_name(sheet_name)
        cell = self.get_cell_from_location(sheet_name, location)
        bruh = False
        if not cell or hidden_name not in self.sheets:
            return
        if is_formula(cell.get_contents()):
            old_children = cell.get_children()
            formula_value, cell_refs = self.get_formula_value(hidden_name, cell, False)
            # if cell_refs != set({}):
            #     self.add_children_cells(hidden_name, location, cell_refs)
            if old_children != cell_refs:
                bruh = True
                self.recompute_cell_and_parents(
                    hidden_name, location, cell.get_contents()
                )
            else:
                old_value = self.sheets[hidden_name].get_cell_value(location)
                self.sheets[hidden_name].set_cell_value(location, formula_value)
                new_value = self.sheets[hidden_name].get_cell_value(location)
                if new_value != old_value:
                    sheet_name = self.get_sheet_name_from_hidden(hidden_name)
                    if (sheet_name, location) not in self.cells_changed:
                        self.cells_changed[(sheet_name, location)] = old_value
                    elif self.cells_changed[(sheet_name, location)] == new_value:
                        del self.cells_changed[(sheet_name, location)]
                    
                    
        return bruh

    def _set_cycle_detected(
        self,
        capture,
    ):
        for sheet_name, loc in capture:
            cell = self.get_cell_from_location(sheet_name, loc)
            if cell:
                cell_val = cell.get_value()
                if (
                    not isinstance(cell_val, CellError)
                    or not cell_val.get_type() == CellErrorType.CIRCULAR_REFERENCE
                ):
                    sheet_name = self.get_sheet_name_from_hidden(sheet_name)
                    if (sheet_name, loc) not in self.cells_changed:
                        self.cells_changed[(sheet_name, loc)] = cell_val
                    elif self.cells_changed[(sheet_name, loc)] == cell_val:
                       del self.cells_changed[(sheet_name, loc)]
                cell.set_cell_value(
                    cell.get_contents(),
                    CellError(CellErrorType.CIRCULAR_REFERENCE, "Cycle detected", None),
                )

    def recompute_cell_and_parents(self, hidden_name, location, contents):
        """Helper function to recompute the parents of cells."""

        cell = self.sheets[hidden_name].get_cell(location)
        # Compute value if the cell is a formula
        if is_formula(contents):
            formula_value, cell_refs = self.get_formula_value(hidden_name, cell, True)

            # Set the cell value
            old_value = self.sheets[hidden_name].get_cell_value(location)
            self.sheets[hidden_name].set_cell_value(location, formula_value)
            new_value = self.sheets[hidden_name].get_cell_value(location)
            if old_value != new_value:
                # self.cells_changed.add(
                #     (self.get_sheet_name_from_hidden(hidden_name), location)
                # )
                sheet_name = self.get_sheet_name_from_hidden(hidden_name)
                if (sheet_name, location) not in self.cells_changed:
                    self.cells_changed[(sheet_name, location)] = old_value
                elif self.cells_changed[(sheet_name, location)] == new_value:
                    del self.cells_changed[(sheet_name, location)]

            # We still need to add the children and parent cells if it's a valid formula
            if cell_refs != set({}):
                self.add_children_cells(hidden_name, location, cell_refs)
        # else:
        # Currently Unnecessary, we aren't resetting parents but we may need
        # to in the future

        # Recompute the value of the specified cell, and the cells that depend on it.
        # cycle_detected, capture = self.graph.topological_sort(hidden_name, location)
        capture, cycle_detected = self.graph.tarjans(hidden_name, location)

        # Technically we don't need to recompute the first cell as it's
        # already been recomputed
        # for i in range(len(capture)):
        #     if len(capture[i]) == 1:
        #         for cell in capture[i]:
        #             self.recompute_cell_value(cell)

        for i in range(len(capture)):
            if len(capture[i]) > 1 or cycle_detected:
                self._set_cycle_detected(capture[i])
        for i in range(len(capture)):
            if len(capture[i]) == 1:
                for cell in capture[i]:
                    self.recompute_cell_value(cell)
                    # break

    def _get_sheet_name_location(self, child, parent_hidden_name):
        if "!" in child:
            # Need the child hidden name
            name_components = child.split("!")
            child_sheet_name = "!".join(name_components[:-1])
            child_location = name_components[-1]
            child_sheet_name = get_hidden_name(child_sheet_name)
        else:
            # we assume the sheet is in the same sheet as the parent
            child_sheet_name = parent_hidden_name
            child_location = child
        return child_sheet_name.lower(), child_location.lower()

    def add_children_cells(self, hidden_name, location, cell_refs):
        """Adds the cells referenced by the formula."""
        # Again unnecessary, we aren't resetting parents but we may need later
        # self.graph.update_children(hidden_name, location)

        for child in cell_refs:
            # We have to update the graph with the new children
            child_sheet_name, child_location = self._get_sheet_name_location(
                child, hidden_name
            )
            self.graph.add_connection(
                hidden_name, location, child_sheet_name, child_location
            )

    def clean_children_cells(self, hidden_name, location):
        """Cleans the children of the cell."""
        cell = self.sheets[hidden_name].get_cell(location)
        if cell is not None:
            children = cell.get_children()
            for child in children:
                child_hidden_name, child_location = self._get_sheet_name_location(
                    child, hidden_name
                )
                self.graph.update_child(
                    hidden_name, location, child_hidden_name, child_location
                )

    @notify_cell_changes
    def set_cell_contents(
        self, sheet_name: str, location: str, contents: Optional[str]
    ) -> None:
        """Set the contents of the specified cell on the specified sheet."""
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # A cell may be set to "empty" by specifying a contents of None.
        #
        # Leading and trailing whitespace are removed from the contents before
        # storing them in the cell.  Storing a zero-length string "" (or a
        # string composed entirely of whitespace) is equivalent to setting the
        # cell contents to None.
        #
        # If the cell contents appear to be a formula, and the formula is
        # invalid for some reason, this method does not raise an exception;
        # rather, the cell's value will be a CellError object indicating the
        # nature of the issue.
        location = location.lower()
        hidden_name = get_hidden_name(sheet_name)
        if hidden_name not in self.sheets:
            raise KeyError(f"{sheet_name} was not found in the Workbook.")
        if not check_valid_cell_location(location):
            raise ValueError(f"{location} is not a valid cell location.")

        contents = contents.lstrip().rstrip() if contents is not None else None

        old_value = self.sheets[hidden_name].get_cell_value(location)

        # Clean the children of the cell before setting the contents incase we set the contents to None
        self.clean_children_cells(hidden_name, location)

        self.sheets[hidden_name].set_cell_contents(hidden_name, location, contents)
        new_value = self.sheets[hidden_name].get_cell_value(location)
        if new_value != old_value:
            sheet_name = self.get_sheet_name_from_hidden(sheet_name)
            if (sheet_name, location) not in self.cells_changed:
                self.cells_changed[(sheet_name, location)] = old_value
            elif self.cells_changed[(sheet_name, location)] == new_value:
                del self.cells_changed[(sheet_name, location)]

        # Recompute cell and parents now actually recomputes the current cell value

        self.recompute_cell_and_parents(hidden_name, location, contents)



    def get_cell_contents(self, sheet_name: str, location: str) -> Optional[str]:
        """Return the contents of the specified cell on the specified sheet."""
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # Any string returned by this function will not have leading or trailing
        # whitespace, as this whitespace will have been stripped off by the
        # set_cell_contents() function.
        #
        # This method will never return a zero-length string; instead, empty
        # cells are indicated by a value of None.

        hidden_name = get_hidden_name(sheet_name)  # TODO pick upper or lower case
        location = location.lower()

        # can consolidate the errors
        if hidden_name not in self.sheets:
            raise KeyError(f"{sheet_name} was not found in the Workbook.")
        if not check_valid_cell_location(location):
            raise ValueError(f"{location} is not a valid cell location.")

        return self.sheets[hidden_name].get_cell_contents(location)

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        """Return the evaluated value of the specified cell on the specified sheet."""
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # The value of empty cells is None.  Non-empty cells may contain a
        # value of str, decimal.Decimal, or CellError.
        #
        # Decimal values will not have trailing zeros to the right of any
        # decimal place, and will not include a decimal place if the value is a
        # whole number.  For example, this function would not return
        # Decimal('1.000'); rather it would return Decimal('1').

        hidden_name = get_hidden_name(sheet_name)
        location = location.lower()

        if hidden_name not in self.sheets:
            raise KeyError(f"{sheet_name} was not found in the Workbook.")
        if not check_valid_cell_location(location):
            raise ValueError(f"{location} is not a valid cell location.")
        return self.sheets[hidden_name].get_cell_value(location)

    def notify_cells_changed(
        self, notify_function: Callable[["Workbook", Iterable[Tuple[str, str]]], None]
    ) -> None:
        """Request that all changes to cell values in the workbook are reported
        to the specified notify_function.  The values passed to the notify
        function are the workbook, and an iterable of 2-tuples of strings,
        of the form ([sheet name], [cell location]).  The notify_function is
        expected not to return any value; any return-value will be ignored.

        Multiple notification functions may be registered on the workbook;
        functions will be called in the order that they are registered.

        A given notification function may be registered more than once; it
        will receive each notification as many times as it was registered.

        If the notify_function raises an exception while handling a
        notification, this will not affect workbook calculation updates or
        calls to other notification functions.

        A notification function is expected to not mutate the workbook or
        iterable that it is passed to it.  If a notification function violates
        this requirement, the behavior is undefined."""
        self.cell_notification_functions.append(notify_function)

    def recompute_sheet_parents(self, sheet_name):
        """
        Updates cells that referenced this new sheet before it actually existed.
        """
        cells_refing_sheet = self.graph.get_sheet_parents(sheet_name)
        # Iterate through these cells to update the values of their parents.
        for sheet_name, loc in cells_refing_sheet:
            self.recompute_cell_and_parents(
                sheet_name, loc, self.get_cell_contents(sheet_name, loc)
            )

    @notify_cell_changes
    def new_sheet(self, sheet_name: Optional[str] = None, copy_sheet = False) -> Tuple[int, str]:
        """Add a new sheet to the workbook.  If the sheet name is specified, it
        # must be unique.  If the sheet name is None, a unique sheet name is
        # generated."""
        # Uniqueness" is determined in a case-insensitive manner,
        # but the case specified for the sheet name is preserved.
        #
        # The function returns a tuple with two elements:
        # (0-based index of sheet in workbook, sheet name).  This allows the
        # function to report the sheet's name when it is auto-generated.
        #
        # If the spreadsheet name is an empty string (not None), or it is
        # otherwise invalid, a ValueError is raised.

        if sheet_name is None:
            counter = 1
            while True:
                if f"Sheet{counter}".lower() in self.sheets:
                    counter += 1
                else:
                    sheet_name = f"Sheet{counter}"
                    break
        elif (
            sheet_name == ""
        ):  # Raise exceptions upfront so we don't end up with nested stuff
            raise ValueError("Sheet name string cannot be empty.")
        # MAKE SURE WE DONT HAVE TWO SHEETS WITH NAME sheEt1 and SHEET1 ebcause not alloed
        if not check_new_name(sheet_name):  # noqa: F405
            raise ValueError(f"{sheet_name} is an Invalid sheet name.")

        hidden_name = get_hidden_name(sheet_name)  # noqa: F405
        if hidden_name in self.sheets:
            raise ValueError(f"{sheet_name} already exists in the Workbook.")
        self.sheets[hidden_name] = Sheet(sheet_name)
        if not self.graph.has_sheet(hidden_name):
            self.graph.add_sheet(hidden_name)
        if not copy_sheet:
            self.recompute_sheet_parents(hidden_name)
        return (self.num_sheets() - 1, sheet_name)

    @notify_cell_changes
    def del_sheet(self, sheet_name: str) -> None:
        """Delete the spreadsheet with the specified name."""
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.

        hidden_name = get_hidden_name(sheet_name)
        if hidden_name in self.sheets:
            # del parent
            # all of the cells which reference this sheet need to have that refrence deleted

            # self.graph.del_parents(hidden_name)

            # Graph.del_sheet(hidden_name)
            # sheet_parents_set = self.graph.get_sheet_connections(hidden_name)
            # if sheet_parents_set:

            del self.sheets[hidden_name]
            self.recompute_sheet_parents(hidden_name)
        else:
            # Can rely on the built in error message from key error
            # Do we want really clear error messages or do we want to rely on the built in ones?
            # THis can give us exact location of the error
            raise KeyError(f"{sheet_name} was not found in the Workbook.")

    @notify_cell_changes
    def rename_sheet(self, old_sheet_name: str, new_sheet_name: str) -> None:
        """Rename the specified sheet to the new sheet name.  Additionally, all
        cell formulas that referenced the original sheet name are updated to
        reference the new sheet name (using the same case as the new sheet
        name, and single-quotes iff [if and only if] necessary).

        The sheet_name match is case-insensitive; the text must match but the
        case does not have to.

        As with new_sheet(), the case of the new_sheet_name is preserved by
        the workbook.

        If the sheet_name is not found, a KeyError is raised.

        If the new_sheet_name is an empty string or is otherwise invalid, a
        ValueError is raised.
        Check for invalid new sheet name"""
        if not new_sheet_name or not check_new_name(old_sheet_name):
            raise ValueError("New sheet name cannot be empty or invalid.")

        # Check if the sheet exists (case-insensitive)
        lower_old_sheet_name = old_sheet_name.lower()
        if lower_old_sheet_name not in self.sheets:
            raise KeyError("Sheet name not found.")

        # Check if the new sheet name is redundant:
        if new_sheet_name.lower() in list(self.sheets.keys()):
            raise ValueError(
                "The new sheet name cannot be the same as an already existing sheet name."
            )

        idx = list(self.sheets.keys()).index(lower_old_sheet_name)

        # Get the sheet object

        sheet = self.sheets.pop(lower_old_sheet_name)

        # Rename the sheet object
        sheet.set_sheet_name(new_sheet_name)
        # Replace sheet name in dependencies:
        # Update the ordered dictionary with the new name
        self.sheets[get_hidden_name(new_sheet_name)] = sheet
        self.move_sheet(get_hidden_name(new_sheet_name), idx)

        recon = Reconstructor(parser)
        fixer = FormulaFixer(old_sheet_name, new_sheet_name)
        cells_refing_renamed_sheet = self.graph.get_sheet_parents(lower_old_sheet_name)
        # Fixes formulas for cells referencing the old sheet name

        for sheet_name, cell_loc in cells_refing_renamed_sheet:
            cell = self.get_cell_from_location(sheet_name, cell_loc)
            if cell is None:
                cell = self.get_cell_from_location(new_sheet_name, cell_loc)
            contents = cell.get_contents()
            parsed_formula = parser.parse(contents)
            fixed_parsed_formula = fixer.transform(parsed_formula)
            fixed_formula = recon.reconstruct(fixed_parsed_formula, insert_spaces=False)
            cell.set_contents("=" + fixed_formula, keep_value=True)

        # Track what cells are already refing the new name.
        self.recompute_sheet_parents(get_hidden_name(new_sheet_name))
        # Update the graph
        self.graph.rename_sheet(
            get_hidden_name(old_sheet_name), get_hidden_name(new_sheet_name)
        )

    @notify_cell_changes
    def move_sheet(self, sheet_name: str, index: int) -> None:
        """Move the specified sheet to the specified index in the workbook's
        ordered sequence of sheets.  The index can range from 0 to
        workbook.num_sheets() - 1.  The index is interpreted as if the
        specified sheet were removed from the list of sheets, and then
        re-inserted at the specified index.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        If the specified sheet name is not found, a KeyError is raised.

        If the index is outside the valid range, an IndexError is raised.
                Convert all keys to lowercase for case-insensitive comparison"""

        # Check if the sheet exists (case-insensitive)
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError("Sheet name not found.")

        # Validate the index
        if not (0 <= index < self.num_sheets()):
            raise IndexError("Index out of range.")

        # Get the original sheet name (with case)
        original_sheet_name = self.sheets[sheet_name.lower()].get_sheet_name()

        # Remove the sheet and store its value
        sheet_value = self.sheets.pop(original_sheet_name.lower())

        # Reinsert the sheet at the specified index
        items = list(self.sheets.items())
        items.insert(index, (original_sheet_name.lower(), sheet_value))
        self.sheets = collections.OrderedDict(items)

    @notify_cell_changes
    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        """Make a copy of the specified sheet, storing the copy at the end of the
        workbook's sequence of sheets.  The copy's name is generated by
        appending "_1", "_2", ... to the original sheet's name (preserving the
        original sheet name's case), incrementing the number until a unique
        name is found.  As usual, "uniqueness" is determined in a
        case-insensitive manner.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        The copy should be added to the end of the sequence of sheets in the
        workbook.  Like new_sheet(), this function returns a tuple with two
        elements:  (0-based index of copy in workbook, copy sheet name).  This
        allows the function to report the new sheet's name and index in the
        sequence of sheets.

        If the specified sheet name is not found, a KeyError is raised.

        Check if the sheet exists (case-insensitive)
        Convert all keys to lowercase for case-insensitive comparison"""
        # TODO redo with old implementation and just check for self referential cells.

        # Check if the sheet exists (case-insensitive)
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError("Sheet name not found.")

        # Get the original sheet name (with case)
        original_sheet_name = self.get_sheet_name_from_hidden(
            get_hidden_name(sheet_name)
        )

        # Copy the content of the found sheet

        # Generate a unique name for the copy
        copy_index = 1

        while True:
            new_sheet_name = f"{original_sheet_name}_{copy_index}"
            if new_sheet_name.lower() not in list(self.sheets.keys()):
                break
            copy_index += 1

        prior_parents = set(self.graph.children_to_parents.keys())

        self.new_sheet(new_sheet_name, copy_sheet=True)

        sheet_data = self.sheets[get_hidden_name(original_sheet_name)].get_sheet_cells().items()
        old_name = get_hidden_name(sheet_name)
        if new_sheet_name.lower() not in prior_parents:
            for loc, cell in sheet_data:
                self.set_cell_contents_copy_sheet(old_name, new_sheet_name, loc, cell)
        else:
            for loc, cell in sheet_data:
                self.set_cell_contents(new_sheet_name, loc, cell.get_contents())
       
        # Return the index and the new sheet name
        new_index = self.num_sheets() - 1
        return new_index, new_sheet_name
    
    @notify_cell_changes
    def set_cell_contents_copy_sheet(
        self, old_sheet_name: str, sheet_name: str, location: str, cell: Cell
    ) -> None:
    
        location = location.lower()
        hidden_name = get_hidden_name(sheet_name)
        if hidden_name not in self.sheets:
            raise KeyError(f"{sheet_name} was not found in the Workbook.")
        if not check_valid_cell_location(location):
            raise ValueError(f"{location} is not a valid cell location.")            



        contents = cell.get_contents()
        contents = contents.lstrip().rstrip() if contents is not None else None

        old_value = self.sheets[hidden_name].get_cell_value(location)

        # Clean the children of the cell before setting the contents incase we set the contents to None
        self.clean_children_cells(hidden_name, location)

        self.sheets[hidden_name].set_cell_contents(hidden_name, location, contents)

        new_value = self.sheets[hidden_name].get_cell_value(location)
        if new_value != old_value:
            sheet_name = self.get_sheet_name_from_hidden(sheet_name)
            if (sheet_name, location) not in self.cells_changed:
                self.cells_changed[(sheet_name, location)] = old_value
            elif self.cells_changed[(sheet_name, location)] == new_value:
                del self.cells_changed[(sheet_name, location)]

        new_cell = self.sheets[hidden_name].get_cell(location)
        new_cell.cache_formula(cell.get_cached_formula())
        # Compute value if the cell is a formula
        if is_formula(contents):
            formula_value, cell_refs = self.get_formula_value(old_sheet_name, new_cell, False)

            # Set the cell value
            old_value = self.sheets[hidden_name].get_cell_value(location)
            self.sheets[hidden_name].set_cell_value(location, formula_value)

            new_value = self.sheets[hidden_name].get_cell_value(location)
            if old_value != new_value:
                sheet_name = self.get_sheet_name_from_hidden(hidden_name)
                if (sheet_name, location) not in self.cells_changed:
                    self.cells_changed[(sheet_name, location)] = old_value
                elif self.cells_changed[(sheet_name, location)] == new_value:
                    del self.cells_changed[(sheet_name, location)]

            # We still need to add the children and parent cells if it's a valid formula
            if cell_refs != set({}):
                self.add_children_cells(hidden_name, location, cell_refs)
        # Recompute cell and parents now actually recomputes the current cell value

    def tarjans(self):
        """Runs Tarjan's algorithm to find the strongly connected components in the graph."""
        return self.graph.tarjans()
