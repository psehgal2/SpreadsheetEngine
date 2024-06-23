from typing import Tuple, Optional, Any
from sheets.utils import convert_location_to_idx
from sheets.Cell import Cell

import heapq


class Sheet_location_obj:
    def __init__(self, sheet_name):
        self.sheet_name = sheet_name

    def get_sheet_location(self):
        return self.sheet_name_obj


class cell_location_obj:
    def __init__(self, cell_location):
        self.cell_location = cell_location

    def get_cell_location(self):
        return self.cell_location_obj


class Sheet:
    def __init__(self, sheet_name) -> None:
        """Initialize a new Sheet object."""
        #
        # The sheet name must be unique within the workbook.  The sheet name
        # match is case-insensitive; the text must match but the case does not
        # have to.
        #
        # If the specified sheet name is not unique, a ValueError is raised.
        #
        # The workbook parameter is the workbook to which the sheet belongs.
        # This parameter must not be None.
        self.row_heap = []
        self.col_heap = []

        self.extent = (0, 0)
        self.sheet_name = sheet_name

        self.data = {}  # String ("A1"): Cell

    def get_extent(self) -> Tuple[int, int]:
        """Return a tuple (num-cols, num-rows) indicating the current extent of
        the specified spreadsheet."""
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        return self.extent

    def get_sheet_name(self):
        return self.sheet_name

    def get_sheet_cells(self):
        return self.data

    def set_cell_contents(
        self, sheet_name, location: str, contents: Optional[str]
    ) -> None:
        """Return a tuple (num-cols, num-rows) indicating the current extent of
        the specified spreadsheet."""
        row, col = convert_location_to_idx(location)
        if location not in self.data and (
            contents is not None and contents.strip() != ""
        ):
            self.data[location] = Cell(sheet_name + "!" + location, contents)
            heapq.heappush(self.row_heap, -row)
            heapq.heappush(self.col_heap, -col)
        elif location not in self.data and (contents is None or contents.strip() == ""):
            pass
        else:
            if contents is None or contents.strip() == "":
                self._shrink_sheet(row, col)
                self.data[location].set_contents(None)
                del self.data[location]
            else:
                self.data[location].set_contents(contents.strip())
        self.update_extent()

    def drop_in_cell(self, location: str, cell: Cell) -> None:
        """Drop the cell into the minheap."""
        row, col = convert_location_to_idx(location)
        if location not in self.data:
            self.data[location] = cell
            heapq.heappush(self.row_heap, -row)
            heapq.heappush(self.col_heap, -col)

        self.extent = (-self.col_heap[0], -self.row_heap[0])

    def get_cell_value(self, location: str) -> Any:
        location = location.lower()
        """Get the value of the cell at the specified location."""
        cell = self.data.get(location, None)
        return cell.get_value() if cell else None

    def get_cell_contents(self, location: str) -> Any:
        """Get the cell contents at the specified location."""
        cell = self.data.get(location, None)
        return cell.get_contents() if cell else None

    def set_cell_value(self, location: str, formula_value: Any) -> None:
        """Get the value of the cell at the specified location."""
        cell = self.data.get(location, None)
        if cell:
            cell.set_cell_value(cell.get_contents(), formula_value=formula_value)

    # min/ max heap for popping and popping off the heap for the new extent
    def _shrink_sheet(self, row, col):
        """
        Shrink the sheet to deallocate memory for deleted (None) cells.
        """
        self.col_heap.remove(-col)
        self.row_heap.remove(-row)
        heapq.heapify(self.col_heap)
        heapq.heapify(self.row_heap)

    def get_cell(self, location: str) -> Cell:
        """Get the cel at the specified location."""
        return self.data.get(location, None)

    def update_extent(self):
        """Update the extent of the sheet. Where the extent is a measure of the
        most distant cell that's occupied by a
        a non Null value."""
        self.extent = (
            (-self.col_heap[0], -self.row_heap[0])
            if len(self.col_heap) > 0 and len(self.row_heap) > 0
            else (0, 0)
        )

    def set_sheet_name(self, name: str):
        self.sheet_name = name
        cells = self.data.values()
        for cell in cells:
            cell.set_location(name, cell.get_location().split("!")[1])

    def get_cells_to_be_moved(self, start_location: str, end_location: str):
        start_row, start_col = convert_location_to_idx(start_location)
        end_row, end_col = convert_location_to_idx(end_location)
        cells = []
        for cell_loc, cell in self.data.items():
            row, col = convert_location_to_idx(cell_loc.split("!")[1])
            if start_row <= row <= end_row and start_col <= col <= end_col:
                cells.append(cell)
        return cells

    ## This Code is maybe deprecated because it can't handle references to cells that haven't been initialized yet.
    # def update_fomula_sheet_names(self,old_sheet_name, new_sheet_name):
    #     for cell in self.data.values():
    #         parents = cell.get_parents()
    #         if len(parents) > 0:
    #             for parent in parents:
    #                 parent.replace_sheet_name(old_sheet_name, new_sheet_name)
