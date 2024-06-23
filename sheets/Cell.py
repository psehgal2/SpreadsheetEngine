import decimal
from sheets.utils import remove_exponent, error_matcher
import re
from sheets.CellError import CellError

# there will only be 4 types of cells
# Decimals (numbers)
# Strings (text, dates, etc) [For these we don't need to adjust
# their representation]
# Formulas (strings that start with "=" and which we need to calculate
# /evaluate)
# CellErrors these we need to handle as well
ALLOWED_PUNCTUATION = [
    " ",
    ".",
    "?",
    "!",
    ",",
    ":",
    ";",
    "!",
    "@",
    "#",
    "$",
    "%",
    "^",
    "&",
    "*",
    "(",
    ")",
    "-",
]


class Cell_location_obj:
    def __init__(self, cell_location):
        self.cell_location_obj = cell_location

    def get_cell_location_obj(self):
        return self.cell_location_obj


class Cell:
    """Cell Class: The Fundamental unit of our Spread Sheet."""

    def __init__(self, location, contents: str):
        """
        args:
            location: the location of the cell given in spreadsheet format
            contents: the user input string representing the cell
        """
        self.children = {}  # str object
        self.contents = None
        self.value = None
        self.location = location
        self.parsed_formula = None
        self.set_contents(contents)

    def cache_formula(self, formula):
        self.parsed_formula = formula

    def get_cached_formula(self):
        return self.parsed_formula

    def set_contents(self, contents: str, keep_value=False):
        """
        Sets the contents of the cell.
        """
        if not isinstance(contents, str) and contents is not None:
            raise TypeError("Contents must be a string")
        if contents is not None and contents.strip() == "":
            self.contents = None
        elif contents is not None:
            self.contents = contents
            if not keep_value:
                self.set_cell_value(contents)
        else:
            self.contents = None
            self.value = None

    def set_cell_value(self, contents: str, formula_value=None):
        """Sets the value of the cell. Includes a keyword arg for the case
        where there is a formula."""
        if contents is not None:
            if contents[0] == "'":
                self.value = str(contents[1:])
            elif contents[0] == "=":
                if isinstance(formula_value, bool):
                    self.value = formula_value
                elif contents.strip().lower() == "=version()":
                    self.value = formula_value
                elif formula_value is not None and not isinstance(
                    formula_value, CellError
                ):
                    try:
                        # if isinstance(formula_value,bool):
                        #     raise decimal.InvalidOperation("We refuse to convert bools")
                        decimal_value = decimal.Decimal(formula_value)
                        if decimal_value.is_finite():
                            self.value = remove_exponent(decimal_value)
                        else:
                            self.value = formula_value
                        if self.value == -0:
                            self.value = decimal.Decimal(0)
                    except decimal.InvalidOperation:
                        self.value = formula_value
                else:
                    self.value = formula_value
            elif error_matcher(contents) is not None:
                self.value = error_matcher(contents)
                pass
            elif contents.lower() == "false" or contents.lower() == "true":
                self.value = contents.lower() == "true"

            else:
                try:
                    decimal_value = decimal.Decimal(contents)
                    if decimal_value.is_finite():
                        self.value = remove_exponent(decimal_value)
                    else:
                        self.value = contents
                except decimal.InvalidOperation:
                    self.value = contents

    def get_value(self):
        return self.value

    def get_contents(self) -> str:
        return self.contents

    def set_children(self, children):
        """sets the children of the cell"""
        self.children = children

    def get_children(self):
        """returns the cells which the formula in this cell references"""
        return self.children

    def add_parent(self, parent):
        """adds a cell to the parent list"""
        if parent not in self.parents:
            self.parents.append(parent)

    def add_child(self, child):
        """adds a cell to the child list"""
        if child not in self.children:
            self.children.append(child)

    def get_location(self):
        """returns the location of the cell in teh workbook."""
        return self.location

    def set_location(self, sheet_name, loc):
        self.location = sheet_name + "!" + loc

    def replace_sheet_name(self, old_sheet_name: str, new_sheet_name: str):
        """Go through and replace occurences of the old sheet name with the new sheet name
        in the contents. Also update all uncessarily single quoted sheet names to remove
        their single quotes."""

        # Replace occurences of the old sheet name with the new sheet name
        pattern = re.compile(
            "('?)("
            + re.escape(old_sheet_name)
            + ")('?)"
            + "(?=!)"
            + "([a-zA-Z]+)"
            + "([0-9]+)",
            re.IGNORECASE,
        )
        self.contents = pattern.sub(new_sheet_name, self.contents)

        # Iterate through and fix all the uncessarily single quoted sheet names
        potential_bad_names = re.findall("'.*'?=!", self.contents)
        for name in potential_bad_names:
            if not (
                (not name[0].isalpha() and not name[0] == "_")
                or set(ALLOWED_PUNCTUATION).isdisjoint(set(new_sheet_name))
            ):
                self.contents.replace(name, name[1:-1])
