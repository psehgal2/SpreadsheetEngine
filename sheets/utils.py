import decimal
from decimal import Decimal

# import deprecation
import re
import json
from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType

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

def convert_to(item, type, error_detail = ""):
    if type == bool:
        if isinstance(item, str):
            if item.lower() == 'true':
                return True
            elif item.lower() == 'false':
                return False
            else:
                return CellError(CellErrorType.TYPE_ERROR, error_detail + "\n Failed to convert string to bool")
        if isinstance(item,Decimal):
            if item == Decimal(0):
                return False
            else:
                return True
        if isinstance(item,bool):
            return item
        if isinstance(item,CellError):
            item : CellError
            return CellError(item.get_type(), f"failed to convert cell error of type {item.get_type()} to bool" + item.get_detail())
        if item is None:
            return False
    if type == str:
        if isinstance(item,Decimal):
            return str(Decimal)
        if isinstance(item,bool):
            if item == False:
                return "FALSE"
            if item == True:
                return "TRUE"
        if isinstance(item,CellError):
            return item
        if isinstance(item,str):
            return item
    if type == Decimal:
        if isinstance(item,str):
            try:
                return decimal.Decimal(float(item))
            except ValueError:
                return CellError(CellErrorType.TYPE_ERROR, "Tried to convert improper string to a float.")
        if isinstance(item,bool):
            if item:
                return 1
            else:
                return 0
        if isinstance(item, CellError):
            return CellError
        if isinstance(item, Decimal):
            return item
        if item is None:
            return 0
    if type == CellError:
        raise ValueError("Tried to Convert to Cell Error. Not yet implemented.")

def get_hidden_name(sheet_name: str) -> str:
    """Return a hidden name for the sheet."""
    #
    # The hidden name is used internally to identify the sheet.  It is not
    # displayed to the user.
    #
    # The hidden name is guaranteed to be unique within the workbook.
    #
    # The sheet name match is case-insensitive; the text must match but the
    # case does not have to.
    #
    # If the specified sheet name is not unique, a ValueError is raised.
    if sheet_name[0] == "'" and sheet_name[-1] == "'":
        sheet_name = sheet_name[1:-1]

    return sheet_name.lower()


def check_new_name(sheet_name: str) -> bool:
    """Return True if the specified sheet name is valid for a new sheet."""
    #
    # User-specified spreadsheet names can be comprised of letters, numbers,
    # spaces, and these punctuation characters: .?!,:;!@#$%^&*()-_
    allowed = ".?!,:;!@#$%^&*()-_ "
    if sheet_name == "":
        return False
    if sheet_name[0] == " ":
        return False
    if sheet_name[-1] == " ":
        return False
    for char in sheet_name:
        if char not in allowed and not char.isalnum():
            return False

    return True


def requires_single_quotes(sheet_name):
    # Strip single quotes if present
    stripped_name = sheet_name.strip("'")

    # Check if the sheet name starts with a non-alphabetical and non-underscore character
    if not re.match("^[a-zA-Z_]", stripped_name):
        return True

    # Check if the sheet name contains characters other than allowed ones
    if not re.match("^[a-zA-Z0-9_]+$", stripped_name):
        return True

    # If none of the conditions are met, single quotes are not required
    return False


def check_valid_cell_location(location: str) -> bool:
    """Return True if the specified cell location is valid."""
    #
    # The cell location is specified as a string in the form 'A1', 'B2', etc.
    # The column must be a capital letter from A to Z, and the row must be a
    # number from 1 to 1000.
    #
    # If the specified cell location is not valid, a ValueError is raised.
    #
    # The cell location match is case-insensitive; the text must match but the
    # case does not have to.

    # check if the location is within A1 to ZZZZ9999

    # The regular expression pattern for Excel cell location
    pattern = r"^[a-z]{1,4}[1-9][0-9]{0,3}$"

    # TODO: change to use the pattern
    # Check if the cell matches the pattern
    if not re.match(pattern, location):
        return False
    return True

    # if len(location) < 2:
    #     return False

    # location = location.upper()

    # if location[0] < 'a' or location[0] > 'Z':
    #     return False

    # if not location.isalnum():
    #     return False

    # endOfLetters = False

    # for c in location:
    #     if endOfLetters and c.isalpha():
    #         return False
    #     if c.isdigit() and not endOfLetters:
    #         if c == '0':
    #             return False
    #         else:
    #             endOfLetters = True

    # return True


def convert_location_to_idx(cell_reference):
    """Extract letters and numbers from the cell reference"""
    letters = "".join(filter(str.isalpha, cell_reference))
    numbers = "".join(filter(str.isdigit, cell_reference))

    # Convert letters to column index
    column_index = 0
    for char in letters:
        column_index = column_index * 26 + (ord(char.upper()) - ord("A")) + 1

    # Convert numbers to row index
    row_index = int(numbers)
    return row_index, column_index


def convert_idx_to_location(row_index, column_index):
    """Convert row and column index to cell reference"""
    # Convert column index to letters
    letters = ""
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        letters = chr(97 + remainder) + letters

    # Convert row index to numbers
    numbers = str(row_index).upper()
    return letters + numbers


def remove_exponent(d):
    """Takes care of zeroes."""
    return d.quantize(Decimal(1)) if d == d.to_integral() else d.normalize()


def is_formula(contents):
    if contents is not None and len(contents) > 1:
        return contents[0] == "="
    return False


# @deprecation
def tarjan_strongly_connected_components(cell):
    """Find the strongly connected components of the graph containing cell"""
    index_counter = [0]
    index = {}
    lowlink = {}
    stack = []
    result = []

    to_visit = [(cell, 0)]  # nodes to visit in DFS, with node and 'i' index
    while to_visit:
        v, i = to_visit[-1]  # peek current node and 'i'
        if i == 0:  # first time seeing the node
            index[v] = index_counter[0]
            lowlink[v] = index_counter[0]
            index_counter[0] += 1
            stack.append(v)
        remove_node = True
        for j in range(i, len(cell.get_parents())):
            w = cell.get_parents()[j]
            if w not in lowlink:
                to_visit[-1] = (v, j + 1)  # next time continue from next node
                to_visit.append((w, 0))  # new node to be visited
                remove_node = False
                break
        if remove_node:
            while stack and index[stack[-1]] >= index[v]:
                w = stack.pop()
                lowlink[w] = min(lowlink[w], lowlink[v])
                if lowlink[w] == index[w]:
                    result.append(w)
                    break
            to_visit.pop()
    return result


# Deprecated we need to perform this inside the workbook load json function
def check_n_load_valid_workbook_json(workbook_json: str) -> json:
    """Check if the workbook json file is valid"""
    try:
        # Try parsing the JSON data
        data = json.loads(workbook_json)

        # Check if the required structure is present
        if "sheets" in data and isinstance(data["sheets"], list):
            for sheet in data["sheets"]:
                if (
                    "name" not in sheet
                    or "cell-contents" not in sheet
                    or not isinstance(sheet["cell-contents"], dict)
                ):
                    return None
        else:
            return None

        return data

    except json.JSONDecodeError:
        return None


ERROR_PRECEDENCE = {
    CellErrorType.PARSE_ERROR: 1,
    CellErrorType.CIRCULAR_REFERENCE: 2,
    CellErrorType.BAD_REFERENCE: 3,
    CellErrorType.BAD_NAME: 4,
    CellErrorType.TYPE_ERROR: 5,
    CellErrorType.DIVIDE_BY_ZERO: 6,
}


def get_error_precedence(
    error: CellError,
):  # spell check presendence, CAN GET RID OF BECUSE ERRORS THEMSELVS ARE ALREADY NUMBERED
    return ERROR_PRECEDENCE[error.get_type()]


def get_highest_precedence_error(values):
    highest_precedence_error = None
    error_precedence = 99999
    for value in values:
        if isinstance(value, CellError):
            current_precedence = get_error_precedence(value)
            if current_precedence < error_precedence:
                highest_precedence_error = value
                error_precedence = current_precedence
    return highest_precedence_error


def error_matcher(text: str) -> CellErrorType:
    if text.lower().strip() == "#ERROR!".lower():
        return CellError(CellErrorType.PARSE_ERROR, "This Error was set by the user.")
    elif text.lower().strip() == "#CIRCREF!".lower():
        return CellError(
            CellErrorType.CIRCULAR_REFERENCE, "This Error was set by the user."
        )
    elif text.lower().strip() == "#REF!".lower():
        return CellError(CellErrorType.BAD_REFERENCE, "This Error was set by the user.")
    elif text.lower().strip() == "#NAME?".lower():
        return CellError(CellErrorType.BAD_NAME, "This Error was set by the user.")
    elif text.lower().strip() == "#VALUE!".lower():
        return CellError(CellErrorType.TYPE_ERROR, "This Error was set by the user.")
    elif text.lower().strip() == "#DIV/0!".lower():
        return CellError(
            CellErrorType.DIVIDE_BY_ZERO, "This Error was set by the user."
        )
    else:
        return None
