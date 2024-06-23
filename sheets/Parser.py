import lark
import decimal
from lark.visitors import visit_children_decor
from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType
from sheets.utils import (
    requires_single_quotes,
    convert_location_to_idx,
    convert_idx_to_location,
    check_valid_cell_location,
    get_highest_precedence_error,
)
import re

ERROR_MAPPING = {
    "#error!": CellError(CellErrorType.PARSE_ERROR, "Oopsy"),
    "#circref!": CellError(CellErrorType.CIRCULAR_REFERENCE, "Your Mom's a Circle"),
    "#ref!": CellError(
        CellErrorType.BAD_REFERENCE,
        "Look bud. All I've ever seen you do is copy and paste from GPT4, and you used your free time to have an affair with my wife. As your project manager I will not be giving you a reference.",
    ),
    "#name?": CellError(CellErrorType.BAD_NAME, "Lmao bro. That is not a sheet."),
    "#value!": CellError(
        CellErrorType.TYPE_ERROR, "I'm pretty sure that's not a real number"
    ),
    "#div/0!": CellError(
        CellErrorType.DIVIDE_BY_ZERO, "WHY WOULD YOU TRY TO DIVIDE BY 0"
    ),
}


# Donnie doens't like this. Its overvly complicated and can easilty be redone better.
def formula_decor(function):
    """A decorator which handles error propogation"""

    # Error precedence is such that: An error will be propogated if its enum value
    # is higher in the CellErrorType than any other error in values.
    # If all errors in value are of equivalent type, then the left most Error (in terms of
    # the formula) will be propogated.
    @visit_children_decor
    def fix_formula(cls: lark.Visitor, values: list):
        highest_precedence_error = get_highest_precedence_error(values)

        if highest_precedence_error is None:
            return_val = function(cls, values)
            return return_val
        else:
            return highest_precedence_error

    return fix_formula


def strip_trailing_zeros(value):
    """
    Strips trailing zeros from a Decimal value and returns it as a string.
    If the input is not a Decimal, it returns the input as is.
    """
    if isinstance(value, decimal.Decimal):
        # Convert to string and strip trailing zeros
        return str(value).rstrip("0").rstrip(".") if "." in str(value) else str(value)
    else:
        # Return non-Decimal values as is
        return value


def convert_nones(l, r):
    if l is None and r is None:
        return (l, r)
    elif l is None:
        if isinstance(r, decimal.Decimal):
            return (decimal.Decimal(0), r)
        if type(r) == str:
            return ("", r)
        if type(r) == bool:
            return (False, r)
    elif r is None:
        if isinstance(l, decimal.Decimal):
            return (l, decimal.Decimal(0))
        if type(l) == str:
            return (l, "")
        if type(l) == bool:
            return (l, False)
    else:
        return (l, r)


def convert_types(l):
    if type(l) == bool:
        return 3
    if type(l) == str:
        return 2
    if type(l) == decimal.Decimal:
        return 1


class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, workbook, sheet_name, formula_functions: dict):
        """Evaluates formulas that have already been parsed by lark.
        args:
            workbook -- the workbook so that the Evaluator can access cell values
            sheet_name -- the name of the sheet in which the formula is located"""
        super(FormulaEvaluator, self).__init__()
        self.workbook = workbook
        self.sheet_name = sheet_name
        self.formula_functions = formula_functions
        self.refs = set({})

    def get_cell_refs(self):
        return self.refs

    @formula_decor
    def parens(self, values):
        """Handles the parentheses nonterminal."""
        return values[0]

    @formula_decor
    def comp_expr(self, values):
        l, op, r = values
        l, r = convert_nones(l, r)
        if type(l) != type(r):
            l = convert_types(l)
            r = convert_types(r)
        if type(l) is str:
            l = l.lower()
            r = r.lower()
        if op in ["<>", "!="]:
            return l != r
        if op == ">":
            return l > r
        if op == "<":
            return l < r
        if op == ">=":
            return l >= r
        if op == "<=":
            return l <= r
        if op in ["==", "="]:
            return l == r

    @formula_decor
    def concat_expr(self, values):
        """Handles the concatenation nonterminal."""
        if values[0] is None:
            values[0] = ""
        if values[1] is None:
            values[1] = ""
        if isinstance(values[0], bool):
            values[0] = "TRUE" if values[0] else "FALSE"
        if isinstance(values[1], bool):
            values[1] = "TRUE" if values[1] else "FALSE"

        if not isinstance(values[0], str) and not isinstance(
            values[0], decimal.Decimal
        ):
            return CellError(
                CellErrorType.TYPE_ERROR,
                f"Concatenation only functions with strings. {values[0]} is not a string.",
            )
        elif not isinstance(values[1], str) and not isinstance(
            values[1], decimal.Decimal
        ):
            return CellError(
                CellErrorType.TYPE_ERROR,
                f"Concatenation only functions with strings. {values[1]} is not a string.",
            )

        # Must handle implicit cell types
        return "".join([str(strip_trailing_zeros(val)) for val in values])

    @formula_decor
    def add_expr(self, values):
        """Handles the addition/subtraction nonterminal."""
        (left, op, r) = values

        try:
            left = decimal.Decimal(0) if left is None else decimal.Decimal(left)
        except decimal.InvalidOperation:
            return CellError(
                CellErrorType.TYPE_ERROR,
                f"Addition/Subtraction only works with decimals. {left} in {left} {op} {r} is not a decimal.",
            )

        try:
            r = decimal.Decimal(0) if r is None else decimal.Decimal(r)
        except decimal.InvalidOperation:
            return CellError(
                CellErrorType.TYPE_ERROR,
                f"Addition/Subtraction only works with decimals. {r} in  {left} {op} {r} is not a decimal.",
            )

        if op == "+":
            return left + r
        else:
            # Handle case of negative numbers?
            return left - r

    @formula_decor
    def unary_op(self, values):
        (op, r) = values
        try:
            r = decimal.Decimal(0) if r is None else decimal.Decimal(r)
        except decimal.InvalidOperation:
            return CellError(
                CellErrorType.TYPE_ERROR,
                f"Unary operators only work with decimals. {r} in {op} {r} is not a decimal.",
            )
        if op == "+":
            return r
        elif op == "-":
            return -r

    @formula_decor
    def mul_expr(self, values):
        """Handles the multiplication/division nonterminal."""
        (left, op, r) = values

        try:
            left = decimal.Decimal(0) if left is None else decimal.Decimal(left)
        except decimal.InvalidOperation:
            return CellError(
                CellErrorType.TYPE_ERROR,
                f"Multiplication/Addition only works with decimals. {left} in {left} {op} {r} is not a decimal.",
            )

        try:
            r = decimal.Decimal(0) if r is None else decimal.Decimal(r)
        except decimal.InvalidOperation:
            return CellError(
                CellErrorType.TYPE_ERROR,
                f"Multiplication/Addition only works with decimals. {r} in  {left} {op} {r} is not a decimal.",
            )

        if op == "*":
            return left * r
        else:
            if r == decimal.Decimal(0):
                return CellError(
                    CellErrorType.DIVIDE_BY_ZERO,
                    f"In this expression you try to divide {left} by 0",
                )
            else:
                return left / r

    @formula_decor
    def number(self, values):
        """Handles the number nonterminal."""
        return decimal.Decimal(values[0])

    @formula_decor
    def string(self, values):
        """Handles the string nonterminal"""
        return values[0].value[1:-1]

    @formula_decor
    def cell(self, values):
        """Handles the cell nonterminal. The cell nonterminal here means another cell referenced in the formula."""
        if len(values) == 2:
            cell_sheet = values[0].value
            cell_loc = values[1].value.lower()
        else:
            cell_sheet = self.sheet_name
            cell_loc = values[0].value.lower()
        cell_loc = re.sub("\$", "", cell_loc)

        self.refs.add(str(cell_sheet) + "!" + cell_loc)
        #TODO:
        try:
            return self.workbook.get_cell_value(cell_sheet, cell_loc)
        except ValueError as E:
            return CellError(
                CellErrorType.BAD_REFERENCE,
                "During parsing an invalid cell location was discovered.",
                E,
            )
        except KeyError as E:
            return CellError(
                CellErrorType.BAD_REFERENCE,
                "During parsing a sheet was given that was not found in the workbook.",
                E,
            )

    @formula_decor
    def error(
        self, values
    ):  # CAN MAKE A DATA MAP A THE TOP LIKE A GLOBAL. DATA DRIVEN APRROACH
        """Handles the error nonterminal by returning the appropriate cell error which is propogated up the tree."""
        error = values[0].value.lower()
        return ERROR_MAPPING[error]

    @formula_decor
    def boolean(self, values):
        value = values[0].value
        if value.lower() == "false":
            return False
        elif value.lower() == "true":
            return True
        else:
            raise ValueError("Invalid boolean value???")

    def range_expr(self, tree):
        children = tree.children
        if len(children) == 3:
            sheet_name = children[0]
            loc_1 = children[1].value
            loc_2 = children[2].value
        else:
            if requires_single_quotes(self.sheet_name):
                sheet_name = lark.Token('QUOTED_SHEET_NAME', self.sheet_name)
            else:
                sheet_name = lark.Token('SHEET_NAME', self.sheet_name)
            loc_1 = children[0].value
            loc_2 = children[1].value

        row_1 , col_1 = convert_location_to_idx(loc_1)
        row_2 , col_2 = convert_location_to_idx(loc_2)
        start_row = min(row_1,row_2)
        end_row = max(row_1,row_2)
        start_col = min(col_1,col_2)
        end_col = max(col_1, col_2)

        cell_arr = []

        for row in range(start_row, end_row + 1):
            cell_row = []
            for col in range(start_col, end_col + 1):
                cell_row.append(lark.Tree(lark.Token('RULE', 'cell'), [sheet_name, lark.Token("CELLREF_NO_ABS", convert_idx_to_location(row,col))]))
            cell_arr.append(cell_row)

        return cell_arr

    def function(self, tree):
        children = tree.children
        func_token = children[0]
        func_name = func_token.value
        try:
            function = self.formula_functions[func_name.upper()]
        except KeyError:
            return CellError(
                CellErrorType.BAD_NAME, f"{func_name} is not a valid function name"
            )
        # I know this is disgusting but bear with me
        # TODO: Fix this whole gross thing
        func = function(self)
        value = func(children[1:])
        return value


#
# # This is stolen from lecture code.
# class CellRefFinder(lark.Visitor):
#     """Finds all cells referenced in the formula."""
#
#     def __init__(self) -> None:
#         super().__init__()
#         self.refs = []
#
#     def cell(self, tree: lark.Tree):
#         """Handles the cell token."""
#         if len(tree.children) == 1:
#             ref = str(tree.children[0])
#             ref = re.sub("\$", "", ref)
#             self.refs.append(ref)
#         else:
#             assert len(tree.children)
#             ref = str(tree.children[-1])
#             ref = re.sub("\$", "", ref)
#             self.refs.append(str(tree.children[0]) + "!" + ref)
#
#     def get_refs(self):
#         """Returns the list of cell references after they've been collected."""
#         return self.refs
#
#     def clear_refs(self):
#         """Resets the list of cell references."""
#         self.refs = []


class FormulaFixer(lark.Transformer):
    """Fixes incorrect references."""

    def __init__(self, old_sheet_name: str, new_sheet_name: str) -> None:
        super().__init__()
        self.old_sheet_name = old_sheet_name
        self.new_sheet_name = new_sheet_name

    def QUOTED_SHEET_NAME(self, sheet_name_tok):
        sheet_name = sheet_name_tok.value
        if not requires_single_quotes(sheet_name_tok):
            sheet_name = sheet_name[1:-1]
            sheet_name_tok.type = "SHEET_NAME"

        if sheet_name.lower().strip("'") == self.old_sheet_name.lower().strip("'"):
            sheet_name = self.new_sheet_name

        sheet_name_tok = lark.Token("SHEET_NAME", sheet_name)
        return sheet_name_tok

    def SHEET_NAME(self, sheet_name_tok):
        sheet_name = sheet_name_tok.value

        if sheet_name.lower().strip("'") == self.old_sheet_name.lower().strip("'"):
            sheet_name = self.new_sheet_name
        sheet_name_tok = lark.Token("SHEET_NAME", sheet_name)
        return sheet_name_tok


class MoveFormula(lark.Transformer):
    """Fixes incorrect references."""

    def __init__(self, delta_row, delta_col) -> None:
        self.delta_row = delta_row
        self.delta_col = delta_col

    def cell(self, values):
        loc = values[-1]
        if loc == lark.Token("ERROR_VALUE", "#REF!"):
            return lark.Tree("error", [loc])
        else:
            return lark.Tree(lark.Token("RULE", "cell"), values)

    def CELLREF_NO_ABS(self, cell_ref_tok):
        (row, col) = convert_location_to_idx(cell_ref_tok.value)
        new_loc = (row + self.delta_row, col + self.delta_col)
        new_loc = convert_idx_to_location(*new_loc)
        if not check_valid_cell_location(new_loc):
            value = lark.Token("ERROR_VALUE", "#REF!")
            return value
        value = lark.Token("CELLREF_NO_ABS", new_loc.upper())
        return value

    def CELLREF_ROW_ABS(self, cell_ref_tok):
        (row, col) = convert_location_to_idx(cell_ref_tok.value)
        new_loc = (row, col + self.delta_col)
        new_loc = convert_idx_to_location(*new_loc)
        if not check_valid_cell_location(new_loc):
            value = lark.Token("ERROR_VALUE", "#REF!")
            return value
        new_loc = re.sub("(?<=[a-zA-Z])[0-9]+$", "$\g<0>", new_loc)
        value = lark.Token("CELLREF_ROW_ABS", new_loc.upper())
        return value

    def CELLREF_COL_ABS(self, cell_ref_tok):
        (row, col) = convert_location_to_idx(cell_ref_tok.value)
        new_loc = (row + self.delta_row, col)
        new_loc = convert_idx_to_location(*new_loc)
        if not check_valid_cell_location(new_loc):
            value = lark.Token("ERROR_VALUE", "#REF!")
            return value

        new_loc = re.sub("^[a-zA-Z]+(?=[0-9])", "$\g<0>", new_loc)

        value = lark.Token("CELLREF_COL_ABS", new_loc.upper())
        return value

    def CELLREF_BOTH_ABS(self, cell_ref_tok):
        return cell_ref_tok


parser = lark.Lark.open(
    "formulas.lark", rel_to=__file__, start="formula", maybe_placeholders=False
)
