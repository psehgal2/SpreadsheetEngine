import context

# Now we can import other things
# import sheets
from tests.project_1.wrongNamesList import  wrongSheetNames, VALID_JSON_DATA
import pytest

from sheets.utils import check_n_load_valid_workbook_json
# from decimal import Decimal

# from sheets.CellError import CellError
# from sheets.CellErrorType import CellErrorType
import json

def setUp():
    """Setup function for the tests"""
    pass

def test_check_n_load_valid_workbook_json():
    """Test if the function check_n_load_valid_workbook_json works correctly"""
    assert check_n_load_valid_workbook_json(VALID_JSON_DATA) != None
    assert check_n_load_valid_workbook_json(VALID_JSON_DATA) == json.loads(VALID_JSON_DATA)
