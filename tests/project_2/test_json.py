import context
import json
import sheets
import json
import sheets
from sheets import Workbook
from sheets import CellError
from sheets import CellErrorType
import pytest
import io
from io import StringIO

# Test incorrect IO 


def test_workbook_to_json():
    # Make a new empty workbook
    wb = sheets.Workbook()
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        print(wb.save_workbook(json_file))
    (index, name) = wb.new_sheet()
    # Should print:  New spreadsheet "Sheet1" at index 0
    # print(f'New spreadsheet "{name}" at index {index}')

    wb.set_cell_contents(name, 'A1', "'123")
    wb.set_cell_contents(name, 'b1', "5.3")
    wb.set_cell_contents(name, 'c1', "=A1*B1")

    (index, name) = wb.new_sheet()
    # Should print:  New spreadsheet "Sheet2" at index 1
    # print(f'New spreadsheet "{name}" at index {index}')

    with open(json_file_path, "w") as json_file:
        wb.save_workbook(json_file)
        workbook_dict = wb.json

    example_workbook = {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "a1": "'123",
                        "b1": "5.3",
                        "c1": "=A1*B1"
                    }
                },
                {
                    "name": "Sheet2",
                    "cell-contents": {}
                }
            ]
    }

    # example_workbook = json.dumps(example_workbook)
    print("Example workbook", workbook_dict)
    print("Actual workbook", example_workbook)
    assert(example_workbook == workbook_dict)


# def test_check_is_json():
#     wb = sheets.Workbook()
#     wb.load_workbook()
#     (index, name) = wb.new_sheet()
#     # Should print:  New spreadsheet "Sheet1" at index 0
#     print(f'New spreadsheet "{name}" at index {index}')

#     wb.set_cell_contents(name, 'a1', "'123")
#     wb.set_cell_contents(name, 'b1', "5.3")
#     wb.set_cell_contents(name, 'c1', "=A1*B1")

#     (index, name) = wb.new_sheet()
#     # Should print:  New spreadsheet "Sheet1" at index 0
#     print(f'New spreadsheet "{name}" at index {index}')

#     workbook_dict = wb.load_workbook()
#     print(type(workbook_dict))
#     # if isinstance(workbook_dict, str):
#     try:
#         # Attempt to decode the string as JSON
#         print("inside2")
#         json.loads(workbook_dict)
#         assert True
#     except json.JSONDecodeError:
#         assert(False)


# works
def test_save_formula():
    print("inside")
    wb = sheets.Workbook()
    print("inside2")
    (index,name) = wb.new_sheet()

    wb.set_cell_contents(name, 'a1', "5")
    wb.set_cell_contents(name, 'b1', "=a1 + 1")
    wb.set_cell_contents(name, 'c1', '=b1 + a1 / 5')
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        print("opening json file")
        wb.save_workbook(json_file)
    with open(json_file_path, "r") as json_file:
        wb_2 = wb.load_workbook(json_file)
        print("new workbook: ", wb_2)
        assert wb_2.get_cell_value(name, 'a1') == wb.get_cell_value(name, 'a1')
        assert wb_2.get_cell_value(name, 'b1') == wb.get_cell_value(name, 'b1')
        assert wb_2.get_cell_value(name, 'c1') == wb.get_cell_value(name, 'c1')


# works
def test_save_error():
    wb = sheets.Workbook()
    (index, name) = wb.new_sheet()
    wb.set_cell_contents(name, 'a1', "#DIV/0!")
    wb.set_cell_contents(name, 'b1', "=#REF!")
    wb.set_cell_contents(name, 'c1', '=D1')
    wb.set_cell_contents(name, 'd1', '=C1')
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        wb.save_workbook(json_file)
    with open(json_file_path, "r") as json_file:
        wb_2 = wb.load_workbook(json_file)

    with open(json_file_path, 'r') as json_file:
        print(type(json_file))
        data = json.load(json_file)
        print("Data:", data)
    print(wb_2.get_cell_value(name, 'a1'))
    print(wb.get_cell_value(name, 'a1'))
    assert wb_2.get_cell_value(name, 'a1').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO
    assert wb.get_cell_value(name, 'a1').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO
    assert wb_2.get_cell_value(name, 'c1').get_type() is wb.get_cell_value(name, 'c1').get_type()
    assert wb_2.get_cell_value(name, 'b1').get_type() is wb.get_cell_value(name, 'b1').get_type()
    assert wb_2.get_cell_value(name, 'd1').get_type() is wb.get_cell_value(name, 'd1').get_type()



# fails when have with StringIO a second time
def test_save_sheets():
    wb = sheets.Workbook()
    sheet_name_list = ["taco", "nuke", "space ship", "watermole", "123153156", "AQUARIUM"]
    for sheet_name in sheet_name_list:
        wb.new_sheet(sheet_name)
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        print("opening json file")
        print("JSON FIle:", json_file)
        wb.save_workbook(json_file)
        val = wb.json
    with open(json_file_path, "r") as json_file:
        print("Output of save workbook", val)
        wb_2 = wb.load_workbook(json_file)
    assert wb.list_sheets() == wb_2.list_sheets()


# works 
# Utility function for creating a sample workbook
def create_sample_workbook():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.set_cell_contents('Sheet1', 'A1', "'123")
    wb.set_cell_contents('Sheet1', 'B1', "5.3")
    wb.set_cell_contents('Sheet1', 'C1', "=A1*B1")
    wb.new_sheet('Sheet2')
    return wb

# Test saving workbook to JSON
def test_save_workbook():
    wb = create_sample_workbook()
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        print("new_inside")
        wb.save_workbook(json_file)
        json_file.seek(0)
    with open(json_file_path, "r") as json_file:
        data = json.load(json_file)
    assert 'sheets' in data
    assert len(data['sheets']) == 2
    assert data['sheets'][0]['name'] == 'Sheet1'
    assert data['sheets'][0]['cell-contents']['a1'] == "'123"


# # Test loading invalid JSON format
def test_load_invalid_json():
    wb = Workbook()
    json_data = '''
    {
        "sheets": "invalid data"
    }
    '''
    json_file_path = "test_workbook.json"

    with open(json_file_path, "r") as json_file:
        json_data = json.loads(json_data)
    with open(json_file_path, "w") as json_file:   
        json.dump(json_data, json_file)
    
    with pytest.raises(TypeError):
        with open(json_file_path, "r") as json_file:
            print("in here4")
            wb.load_workbook(json_file)


# Test loading workbook with missing fields
def test_load_missing_fields():
    wb = Workbook()
    json_data = '''
    {
        "sheets":[
            {
                "cell-contents":{
                    "A1":"123"
                }
            }
        ]
    }
    '''
    json_file_path = "test_workbook.json"
    with open(json_file_path, "r") as json_file:
        json_data = json.loads(json_data)
    with open(json_file_path, "w") as json_file:
        json.dump(json_data, json_file)
    with pytest.raises(KeyError):
        with open(json_file_path, "r") as json_file:
            loaded_data = wb.load_workbook(json_file) 
            assert (loaded_data is None)
    

# Test saving and loading empty workbook
def test_empty_workbook():
    wb = Workbook()
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        wb.save_workbook(json_file)
        val = wb.json
        print("Output of save workbook", val)
        json_file.seek(0)
    with open(json_file_path, "r") as json_file:
        loaded_wb = wb.load_workbook(json_file)
    assert loaded_wb.num_sheets() == 0

# Test saving and loading workbook with a cell containing bad formula
    


# NOT SURE WHAT THIS SHOULD OUTPUT
# def test_workbook_with_bad_formula():
#     wb = create_sample_workbook()
#     wb.set_cell_contents('Sheet1', 'd1', "=BADFORMULA()")
#     json_file_path = "test_workbook.json"
#     with open(json_file_path, "w") as json_file:
#         wb.save_workbook(json_file)
#         json_file.seek(0)
#     with open(json_file_path, "r") as json_file:
#         data = json.load(json_file)
#         assert data['sheets'][0]['cell-contents']['d1'] == "=BADFORMULA()"
#         loaded_wb = wb.load_workbook(json_file)
#     print("Testing load:", loaded_wb.get_cell_value('Sheet1', 'd1'))
#     assert loaded_wb.get_cell_value('Sheet1', 'd1').get_type() == sheets.CellErrorType.PARSE_ERROR

# Add more test cases as needed to cover additional edge cases and scenarios.


