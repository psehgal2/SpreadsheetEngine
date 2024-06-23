import decimal
from sheets.CellError import CellError
from sheets.CellErrorType import CellErrorType
from sheets.utils import get_highest_precedence_error
import lark
from lark import Token, Tree
import re
from sheets.utils import convert_to

# TODO Error SxTuff:
#  1. WHat if evaluation fails (parse error).
#  2. What if args are wrong type (Type error, but annoying to check).
#  3. What if args are a cell erorr? What do we return.

def flatten_2d_list(two_d_list):
    # Use a list comprehension to flatten the 2D list
    return [item for sublist in two_d_list for item in sublist]

def get_column(my_list, column_index):
    return [row[column_index] for row in my_list if len(row) > column_index]


def convert_to_string(arg):
    if isinstance(arg, str):
        return arg
    elif isinstance(arg, CellError):
        return


# Send errors up the change broadly.
def handle_error(extra_detail, error: CellError):
    return CellError(error.get_type(), extra_detail + "\n" + "\n" + error.get_detail())


def convert_to_bool(arg):
    if isinstance(arg, str):
        if arg.lower() == "false":
            return False
        elif arg.lower() == "true":
            return True
        else:
            return CellError(
                CellErrorType.TYPE_ERROR, f"Invalid Non Boolean Argument - {arg}"
            )
    else:
        return bool(arg)


def reparse_cellrange(match_object):
    if not match_object:
        return None

    # Extracting groups from the match object
    sheet_name_group = match_object.group(1) if match_object.group(1) else ""
    start_cell_ref = match_object.group(3)
    end_cell_ref = match_object.group(6)

    # Sanitizing the sheet name: remove leading/trailing quotes and exclamation mark
    sheet_name = sheet_name_group.replace("'", "").strip("!")

    # Sanitizing cell references: remove dollar signs
    start_cell_ref_clean = start_cell_ref.replace("$", "")
    end_cell_ref_clean = end_cell_ref.replace("$", "")

    # Creating Tokens
    # For the sheet name, create a token only if the sheet name exists
    if sheet_name:
        sheet_name_token = Token("SHEET_NAME", sheet_name)
        start_cell_token = Token("CELLREF_NO_ABS", start_cell_ref_clean)
        end_cell_token = Token("CELLREF_NO_ABS", end_cell_ref_clean)
        ast = Tree("range_expr", [sheet_name_token, start_cell_token, end_cell_token])
    else:
        start_cell_token = Token("CELLREF_NO_ABS", start_cell_ref_clean)
        end_cell_token = Token("CELLREF_NO_ABS", end_cell_ref_clean)
        ast = Tree("range_expr", [start_cell_token, end_cell_token])

    return ast


def reparse_cellref(match_object):
    # argument to self.interpreter.cell instead of calling directly to teh workbook.
    sheet_name = (
        match_object.group(1)[:-1] if match_object.group(1) else None
    )  # Remove trailing '!' from sheet name
    cell_ref = match_object.group(3)
    ref_type = "CELL_REF_NO_ABS"
    if match_object.group(4) == "$":
        ref_type = "CELL_REF_ROW_ABS"
    if match_object.group(5) == "$":
        if ref_type == "CELL_REF_ROW_ABS":
            ref_type = "CELL_REF_BOTH_ABS"
        else:
            ref_type = "CELL_REF_COL_ABS"
    if sheet_name:
        sheet_type = "QUOTED_SHEET_NAME" if "'" in sheet_name else "SHEET_NAME"
        return lark.Tree(
            "cell", [lark.Token(sheet_type, sheet_name), lark.Token(ref_type, cell_ref)]
        )
    else:
        return lark.Tree("cell", [lark.Token(ref_type, cell_ref)])

class SheetFunction:
    def __init__(self, interpeter: lark.visitors.Interpreter, lazy=False, flatten=False):
        self.interpeter = interpeter
        self.lazy = lazy
        self.flatten = flatten

    def eval_arg(self, arg):
        return self.interpeter.visit(arg)

    def __call__(self, args):
        if not self.lazy:
            return self.func(args)

        args = [self.eval_arg(arg) for arg in args]
        if self.flatten:
            new_args = []
            for arg in args:
                if isinstance(arg,list):
                    new_args.extend(self.eval_arg(sub_arg) for sub_arg in flatten_2d_list(arg))
                else:
                    new_args.append(arg)
            args = new_args

        highest_precedence_error = get_highest_precedence_error(args)
        if highest_precedence_error is not None:
            return highest_precedence_error
        else:
            return self.func(tuple(args))

    def func(self, args):
        pass


class AND(SheetFunction):
    def __init__(self, interpeter: lark.visitors.Interpreter):
        super().__init__(interpeter, lazy=True)

    def func(self, args):
        final_val = True
        for arg in args:
            arg = convert_to_bool(arg)
            if isinstance(arg, CellError):
                return arg
            final_val = final_val and arg
        return final_val


class OR(SheetFunction):
    def __init__(self, interpeter: lark.visitors.Interpreter):
        super().__init__(interpeter, lazy=True)

    def func(self, args):
        if len(args) < 1:
            return CellError(
                CellErrorType.TYPE_ERROR, "Or function requires at least 1 argument."
            )
        final_val = False
        for arg in args:
            arg = convert_to_bool(arg)
            if isinstance(arg, CellError):
                return arg
            final_val = final_val or arg
        return final_val


class NOT(SheetFunction):
    def __init__(self, interpeter: lark.visitors.Interpreter):
        super().__init__(interpeter, lazy=True)

    def func(self, args):
        if len(args) != 1:
            return CellError(
                CellErrorType.TYPE_ERROR, "NOT function requires a single argument"
            )
        arg = convert_to_bool(args[0])
        if isinstance(arg, CellError):
            return arg
        else:
            return not bool(args[0])


class XOR(SheetFunction):
    def __init__(self, interpeter: lark.visitors.Interpreter):
        super().__init__(interpeter, lazy=False)
        self.and_func = AND(interpeter)
        self.or_func = OR(interpeter)

    # So this defo isn't the most efficient but rn I'm too lazy to speed it up
    def func(self, args):
        if len(args) < 1:
            return CellError(
                CellErrorType.TYPE_ERROR, "Or function requires at least 1 argument."
            )
        anded = self.and_func(args)
        ored = self.or_func(args)
        if isinstance(anded, CellError) or isinstance(ored, CellError):
            return CellError(CellErrorType.TYPE_ERROR, "Unsupported type condition")
        if isinstance(anded, CellError):
            return anded
        if isinstance(ored, CellError):
            return ored

        return not anded and ored


class EXACT(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=True)

    def func(self, args):
        if len(args) != 2:
            return CellError(
                CellErrorType.TYPE_ERROR,
                "Incorrect number of arguments given to NOT function.",
            )
        # TODO: address this more robustly
        l, r = args
        return convert_to(l,str) == convert_to(r,str)


class SUM(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=True, flatten=True)

    def func(self, args):
        if len(args)==0:
            return CellError(CellErrorType.TYPE_ERROR,"SUM requires at least 1 argument")
        try:
            return sum([convert_to(arg, decimal.Decimal) for arg in args ])
        except TypeError:
            return CellError(CellErrorType.TYPE_ERROR, "Arguments to SUM were invalid")


class MAX(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=True, flatten=True)
    def func(self, args):
        if len(args) == 0:
            return CellError(CellErrorType.TYPE_ERROR, "MAX requires at least 1 argument")
        args = list(filter(lambda x: x != None, args))
        if len(args)==0:
            return 0
        try:
            return max([convert_to(arg, decimal.Decimal) for arg in args ])
        except TypeError:
            return CellError(CellErrorType.TYPE_ERROR, "Arguments to MAX were invalid")


class MIN(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=True, flatten=True)
    def func(self, args):
        if len(args) == 0:
            return CellError(CellErrorType.TYPE_ERROR, "MIN requires at least 1 argument")
        args = list(filter(lambda x: x != None, args))

        if len(args)==0:
            return 0
        try:
            return min([convert_to(arg, decimal.Decimal) for arg in args ])
        except TypeError:
            return CellError(CellErrorType.TYPE_ERROR, "Arguments to MIN were invalid")


class AVERAGE(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=True, flatten=True)
    def func(self, args):
        if len(args) == 0:
            return CellError(CellErrorType.TYPE_ERROR, "Avg requires at least 1 argument")
        args = list(filter(lambda x: x != None, args))
        if len(args) == 0:
            return CellError(CellErrorType.DIVIDE_BY_ZERO, "Averaging over no elements yields div by 0.")
        try:
            return sum([convert_to(arg, decimal.Decimal) for arg in args ]) / len(args)
        except TypeError:
            return CellError(CellErrorType.TYPE_ERROR, "Arguments to Average were invalid")


class HLOOKUP(SheetFunction):
    def __init__(self,interpreter : lark.visitors.Interpreter):
        super().__init__(interpreter,lazy=False)

    def func(self, args):
        key = self.eval_arg(args[0])
        arg_range = self.eval_arg(args[1])
        index = self.eval_arg(args[2])
        error =  get_highest_precedence_error([key, arg_range, index])
        if error:
            return error

        if not isinstance(arg_range, list):
            return CellError(CellErrorType.TYPE_ERROR, "second argument to lookup must be a cell range.")

        if (
                not isinstance(index, decimal.Decimal)
                or index <= 0
                or index.as_integer_ratio()[1] != 1
                or index > len(arg_range)
        ):
            return CellError(CellErrorType.TYPE_ERROR, "Index must be a positive integer")

        search_row = arg_range[0]
        search_row = [self.eval_arg(elem) for elem in search_row]

        error = get_highest_precedence_error(search_row)
        if error:
            return error

        try:
            col_idx = search_row.index(key)
        except ValueError:
            return CellError(CellErrorType.TYPE_ERROR, "Key in HLOOKUP not found")
        result_col = get_column(arg_range, col_idx)
        result_col = [self.eval_arg(elem) for elem in result_col]
        return result_col[int(index)-1]


class VLOOKUP(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=False)

    def func(self, args):
        key = self.eval_arg(args[0])
        arg_range = self.eval_arg(args[1])
        index = self.eval_arg(args[2])

        error =  get_highest_precedence_error([key, arg_range, index])
        if error:
            return error

        if not isinstance(arg_range, list):
            return CellError(CellErrorType.TYPE_ERROR, "second argument to lookup must be a cell range.")

        if (
            not isinstance(index, decimal.Decimal)
            or index <= 0
            or index.as_integer_ratio()[1] != 1
            or index > len(arg_range[0])

        ):
            return CellError(CellErrorType.TYPE_ERROR, "Index must be a positive integer")

        search_col = get_column(arg_range,0)
        search_col = [self.eval_arg(elem) for elem in search_col ]
        error = get_highest_precedence_error(search_col)

        if error:
            return error

        try:
            row_idx = search_col.index(key)
        except ValueError:
            return CellError(CellErrorType.TYPE_ERROR, "Key in VLOOKUP not found")

        result_row = arg_range[row_idx]
        result_row = [self.eval_arg(elem) for elem in result_row]
        return result_row[int(index)-1]


class IF(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=False)

    def func(self, args):
        """In this case"""
        num_args = len(args)
        if num_args != 2 and num_args != 3:
            return CellError(
                CellErrorType.TYPE_ERROR,
                "Incorrect number of arguments given to IF function.",
            )
        cond = self.eval_arg(args[0])
        cond = convert_to(cond,bool)
        if isinstance(cond, CellError):
            return cond
        if cond:
            val = self.eval_arg(args[1])
            return val
        elif num_args == 3:
            val = self.eval_arg(args[2])
            return val
        else:
            return False


class IFERROR(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=False)

    def func(self, args):
        """In this case"""
        num_args = len(args)
        if num_args != 2 and num_args != 1:
            return CellError(
                CellErrorType.TYPE_ERROR,
                "Incorrect number of arguments given to IF function.",
            )
        cond = self.eval_arg(args[0])
        if isinstance(cond, CellError):
            if len(args) > 1:
                return self.eval_arg(args[1])
            else:
                return ""
        else:
            return cond


class ISERROR(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=False)

    def func(self, args):
        """In this case"""
        num_args = len(args)
        if num_args != 1:
            return CellError(
                CellErrorType.TYPE_ERROR,
                "Incorrect number of arguments given to ISERROR function.",
            )
        arg = self.eval_arg(args[0])
        return isinstance(arg, CellError)


class CHOOSE(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=False)

    def func(self, args):
        """In this case"""
        num_args = len(args)
        if num_args < 2:
            return CellError(
                CellErrorType.TYPE_ERROR,
                "Incorrect number of arguments given to IF function.",
            )
        index = self.eval_arg(args[0])
        if isinstance(index, CellError):
            return index

        if (
            not isinstance(index, decimal.Decimal)
            or index <= 0
            or index.as_integer_ratio()[1] != 1
            or index > len(args) - 1
        ):
            return CellError(CellErrorType.TYPE_ERROR, f"Index must be an integer between 1 and {len(args)-1}")

        return self.eval_arg(args[int(index)])


class ISBLANK(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=False)

    def func(self, args):
        value = self.eval_arg(args[0])
        if value == False or value == "" or value == 0:
            return False
        elif value is None:
            return True
        else:
            return False


class VERSION(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=True)

    def func(self, args):
        """In this case"""
        version = "1.4"
        num_args = len(args)
        if num_args != 0:
            return CellError(
                CellErrorType.TYPE_ERROR,
                "Incorrect number of arguments given to IF function.",
            )
        return str(version)


class INDIRECT(SheetFunction):
    def __init__(self, interpreter: lark.visitors.Interpreter):
        super().__init__(interpreter, lazy=False)

    def func(self, args):
        if len(args) != 1:
            return CellError(
                CellErrorType.TYPE_ERROR, "INDIRECT only accepts one argument. "
            )
        arg = self.eval_arg(args[0])

        if isinstance(arg,CellError):
            return arg

        # TODO: This is not right, we can pass INDIRECT(A1), need to convert it to a string or something
        if not isinstance(arg, str):
            return CellError(
                CellErrorType.BAD_REFERENCE,
                "Argument to function Indirect must be a string",
            )

        arg = arg.strip()
        # TODO: Precompile regex: use sheet name part
        # Precompile
        match_object = re.match(
            r"^(([A-Za-z_][A-Za-z0-9_]*|\'[^']*\')!)?((\$)?[A-Za-z]+(\$)?[1-9][0-9]*)\:((\$)?[A-Za-z]+(\$)?[1-9][0-9]*$)",
            arg,
        )
        if match_object:
            cell_range = reparse_cellrange(match_object)
            return self.interpeter.range_expr(cell_range)


        match_object = re.match(
            r"^(([A-Za-z_][A-Za-z0-9_]*|\'[^']*\')!)?((\$)?[A-Za-z]+(\$)?[1-9][0-9]*$)",
            arg,
        )
        if match_object:
            cell_ref = reparse_cellref(match_object)
            return self.interpeter.cell(cell_ref)
        else:
            return CellError(
                CellErrorType.BAD_REFERENCE,
                f"Argument to function given as {arg} is not a valid cell location!",
            )


function_directory = {
    "AND": AND,
    "OR": OR,
    "NOT": NOT,
    "XOR": XOR,
    "EXACT": EXACT,
    "SUM": SUM,
    "IF": IF,
    "IFERROR": IFERROR,
    "ISBLANK": ISBLANK,
    "ISERROR": ISERROR,
    "CHOOSE": CHOOSE,
    "INDIRECT": INDIRECT,
    "MAX" : MAX,
    "MIN": MIN,
    "AVERAGE": AVERAGE,
    "VLOOKUP" : VLOOKUP,
    "HLOOKUP" : HLOOKUP,
    "VERSION" : VERSION
}
