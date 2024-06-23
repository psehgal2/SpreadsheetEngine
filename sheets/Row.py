from enum import Enum
from decimal import Decimal
from sheets.CellError import CellError
import functools
NONE_TYPE = type(None)

class ORDER(Enum):
    NONE_TYPE = -2
    str = 1
    Decimal = 0
    CellError = -1
    bool = 2

order_dict = {type(None): -2,CellError : -1, Decimal : 0 ,str : 1, bool : 2}


@functools.total_ordering
class Row:
    def __init__(self, row, left_col, right_col, sort_cols):
        self.curr_row = row
        self.left_col = left_col
        self.right_col = right_col
        self.sort_cols = sort_cols
        self.sort_cols_ascend = [True if col > 0 else False for col in sort_cols]
        self.col_vals = []

    def get_col_val(self):
        return self.col_vals

    def add_col(self, val):
        self.col_vals.append(val)

    def __eq__(self, row2):
        return self.col_vals == row2.get_col_val()

    def get_og_row(self):
        return self.curr_row
    
    def __lt__(self, row2):
        for i in range(len(self.sort_cols)):
            ascend = self.sort_cols_ascend[i]
            val1 = self.col_vals[i]
            val2 = row2.get_col_val()[i]
            if type(val1) != type(val2):
                if ascend:
                    return order_dict[type(val1)] < order_dict[type(val2)]
                else:
                    return order_dict[type(val1)] > order_dict[type(val2)]
            else:
                if val1 is None:
                    continue 
                elif isinstance(val2, CellError):
                    type1 = val1.get_type().value
                    type2 = val2.get_type().value
                    if type1 != type2:
                        if ascend:
                            return type1 < type2
                        else:
                            return type1 > type2
                elif isinstance(val2, str):
                    if val1.lower() != val2.lower():
                        if ascend:
                            return val1.lower() < val2.lower()
                        else:
                            return val1.lower() > val2.lower()
                elif isinstance(val2, bool) or isinstance(val2, Decimal):
                    if val1 != val2:
                        if ascend:
                            return val1 < val2
                        else:
                            return val1 > val2
                else:
                    raise TypeError
        return False