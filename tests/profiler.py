import context
import sheets
from sheets.Workbook import Workbook
import string
import random
import cProfile as cp
from typing import List
import pstats

#This is a decorator
def profile_it(func):
    def run_then_print():
        profile = func()
        stats = profile.getstats()
        sum_time = 0
        for stat in stats:
            sum_time += stat.totaltime
        print(f"for test {func} total time: {sum_time}")
        profile.disable()

        return profile
    return run_then_print

#Make a "Test_Utils" and put this there.
def convert_to_excel_coordinates(row: int, col: int) -> str:
    """
    Convert a 1-indexed (row, col) coordinate to an Excel-style coordinate.

    Parameters:
    - row: Integer representing the row number (1-indexed).
    - col: Integer representing the column number (1-indexed).

    Returns:
    - A string representing the Excel-style coordinate (e.g., "A1").
    """
    excel_col = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        excel_col = chr(65 + remainder) + excel_col
    return f"{excel_col}{row}"


def generate_random_string(length=10):
    """Generate a random string of fixed length."""
    letters = string.ascii_letters + string.digits
    return ''.join(random.choice(letters) for i in range(length))

# Each of these profiler tests do load tests, they then print the overall runtime, and return a cProfiler object.

@profile_it
def pro_load_workbook():
    wb = Workbook()
    (index, name) = wb.new_sheet()
    wb.set_cell_contents(name, 'a1', "5")
    wb.set_cell_contents(name, 'b1', "=a1 + 1")
    wb.set_cell_contents(name, 'c1', '=b1 + a1 / 5')
    json_file_path = "test_workbook.json"
    with open(json_file_path, "w") as json_file:
        wb.save_workbook(json_file)
    ITERATIONS = 100
    with open(json_file_path, "r") as json_file:
        with cp.Profile() as profile:
            for i in range(ITERATIONS):
                wb_2 = Workbook.load_workbook(json_file)
        return profile

# loading a workbook, copying or renaming a sheet, and moving or copying an area of cells.
    
@profile_it
def pro_copy_sheet():
    wb = Workbook()
    ITERATIONS = 100
    wb.new_sheet("Sheet1")
    wb.set_cell_contents("Sheet1", 'a1', "5")
    wb.set_cell_contents("Sheet1", 'b1', "=a1 + 1")
    with cp.Profile() as profile:
        for i in range(ITERATIONS):
            wb.copy_sheet("Sheet1")  # Creates "Sheet1_1"
    return profile

@profile_it
def pro_rename_sheet():
    wb = Workbook()
    ITERATIONS = 100
    with cp.Profile() as profile:
        for i in range(ITERATIONS):
            wb.new_sheet("Sheet" + str(2*i))
            wb.rename_sheet("Sheet" + str(2*i), "Sheet" + str(2*i + 1))  # Creates "Sheet1_1"
    return profile


@profile_it
def pro_copy_cells():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.set_cell_contents('Sheet1', 'A1', "'123")
    wb.set_cell_contents('Sheet1', 'B1', "5.3")
    wb.set_cell_contents('Sheet1', 'C1', "=A1*B1")
    ITERATIONS = 100
    with cp.Profile() as profile:
        for i in range(1, ITERATIONS):
            wb.copy_cells('Sheet1', 'A1', 'C1',  'A1', 'Sheet2')
    return profile


@profile_it
def pro_copy_cells2():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    ITERATIONS = 100
    for i in range(1, ITERATIONS):
        wb.set_cell_contents('Sheet1', 'A' + str(i), "'123")
        wb.set_cell_contents('Sheet1', 'B' + str(i), "=(A1 + 5) / 2 - 6 * IF(TRUE, 2, 4) - 1/7 "
                                                     "+ IFERROR(False, 1) - 6  + IF(FALSE, 0 , (1*6-4 * 1/2 + (6*1 - 2/(5+1) * 1000)))")

    with cp.Profile() as profile:
        wb.copy_cells('Sheet1', 'A1', 'B' + str(ITERATIONS - 1),  'A1', 'Sheet2')
    # check if working as expected 
    return profile

# check multiple move cells operations 
@profile_it
def pro_move_cells():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    wb.set_cell_contents('Sheet1', 'A1', "'123")
    wb.set_cell_contents('Sheet1', 'B1', "5.3")
    wb.set_cell_contents('Sheet1', 'C1', "=A1*B1")
    ITERATIONS = 100
    with cp.Profile() as profile:
        for i in range(ITERATIONS):
            wb.move_cells('Sheet1', 'A1', 'C1',  'A1', 'Sheet2')
    return profile

@profile_it
def pro_move_cells2():
    wb = Workbook()
    wb.new_sheet('Sheet1')
    wb.new_sheet('Sheet2')
    ITERATIONS = 100
    for i in range(1, ITERATIONS):
        wb.set_cell_contents('Sheet1', 'A' + str(i), "'123")
        wb.set_cell_contents('Sheet1', 'B' + str(i), "=A1")
    with cp.Profile() as profile:
        wb.move_cells('Sheet1', 'A1', 'B' + str(ITERATIONS - 1),  'A1', 'Sheet2')
    # check if working as expected 
    return profile  


@profile_it
def pro_create_sheets():
    ITERATIONS = 100
    STRING_LENGTH = 50
    wb = Workbook()
    with cp.Profile() as profile:
        for i in range(ITERATIONS):
            wb.new_sheet(generate_random_string(STRING_LENGTH))
    return profile

@profile_it
def pro_set_cell_contents():
    ITERATIONS = 100000
    STRING_LENGTH = 500
    wb = Workbook()
    wb.new_sheet("test_sheet")

    with cp.Profile() as profile:
        for i in range(1, 10000):
            row, col = (i % 9998 + 1, i // 9998 + 1)
            sheet_coords = convert_to_excel_coordinates(row, col)
            wb.set_cell_contents("test_sheet", f"{sheet_coords}", generate_random_string(STRING_LENGTH))
    return profile

@profile_it
def pro_update_formula_chain():
    ITERATIONS = 9997
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = convert_to_excel_coordinates(1,1)
    for i in range(1,10000):
        row,col = ( i % 9998 + 1,i // 9998 + 1)
        sheet_coords = convert_to_excel_coordinates(row,col)
        wb.set_cell_contents("test_sheet", f"{sheet_coords}", f"={old_sheet_coords}")
        old_sheet_coords = sheet_coords
    with cp.Profile() as profile:
        wb.set_cell_contents("test_sheet", f"A1", "2")
    # profile.print_stats()
    return profile

@profile_it
def pro_update_formula_chain():
    ITERATIONS = 9997
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = convert_to_excel_coordinates(1,1)
    for i in range(1,10000):
        row,col = ( i % 9998 + 1,i // 9998 + 1)
        sheet_coords = convert_to_excel_coordinates(row,col)
        wb.set_cell_contents("test_sheet", f"{sheet_coords}", f"={old_sheet_coords}")
        old_sheet_coords = sheet_coords
    with cp.Profile() as profile:
        wb.set_cell_contents("test_sheet", f"A1", "2")
    # profile.print_stats()
    return profile

@profile_it
def pro_write_formula_chain():
    ITERATIONS = 9997
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = convert_to_excel_coordinates(1,1)
    with cp.Profile() as profile:
        for i in range(1,10000):
            row,col = ( i % 9998 + 1,i // 9998 + 1)
            sheet_coords = convert_to_excel_coordinates(row,col)
            wb.set_cell_contents("test_sheet", f"{sheet_coords}", f"= (3 + {old_sheet_coords})/2")
            old_sheet_coords = sheet_coords
    return profile


@profile_it
def pro_cyle_formula_chain():
    ITERATIONS = 500
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = convert_to_excel_coordinates(1,1)
    for i in range(1,10000):
            row,col = ( i % 9998 + 1,i // 9998 + 1)
            sheet_coords = convert_to_excel_coordinates(row,col)
            wb.set_cell_contents("test_sheet", f"{sheet_coords}", f"={old_sheet_coords}")
            old_sheet_coords = sheet_coords
    with cp.Profile() as profile:
        wb.set_cell_contents("test_sheet", "A1", f"={old_sheet_coords}")

    return profile


@profile_it
def pro_update_formula_pyramid():
    ITERATIONS = 100
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = [convert_to_excel_coordinates(1,1)]
    for i in range(1,ITERATIONS):
        row,col = ( i % 9998 + 1,i // 9998 + 1)
        sheet_coords = convert_to_excel_coordinates(row,col)
        wb.set_cell_contents("test_sheet", f"{sheet_coords}", "=("+"+".join(old_sheet_coords) + f")/ {i}")
        old_sheet_coords = sheet_coords
    with cp.Profile() as profile:
        wb.set_cell_contents("test_sheet", f"A1", "2")
    # profile.print_stats()
    return profile

@profile_it
def pro_write_formula_pyramid():
    ITERATIONS = 100
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = [convert_to_excel_coordinates(1,1)]
    with cp.Profile() as profile:
        for i in range(1,ITERATIONS):
            row,col = ( i % 9998 + 1,i // 9998 + 1)
            sheet_coords = convert_to_excel_coordinates(row,col)
            wb.set_cell_contents("test_sheet", f"{sheet_coords}", "=("+"+".join(old_sheet_coords) + f")/ {i}")
            old_sheet_coords = sheet_coords

    # profile.print_stats()
    return profile

@profile_it
def pro_cycle_formula_pyramid():
    ITERATIONS = 100
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = [convert_to_excel_coordinates(1,1)]
    for i in range(1,ITERATIONS):
            row,col = ( i % 9998 + 1,i // 9998 + 1)
            sheet_coords = convert_to_excel_coordinates(row,col)
            wb.set_cell_contents("test_sheet", f"{sheet_coords}", "=("+"+".join(old_sheet_coords) + f")/ {i}")
            old_sheet_coords.append(sheet_coords)

    with cp.Profile() as profile:
        wb.set_cell_contents("test_sheet", "A1", "=("+"+".join(old_sheet_coords) + f")/ {100}")

    return profile


#All the cycle stuff
@profile_it
def pro_cycle_break_unbreak():
    CYCLE_SIZE = 1000
    ITERATIONS = 100
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet",f"A1","1")
    old_sheet_coords = convert_to_excel_coordinates(1,1)
    for i in range(1,CYCLE_SIZE):
            row,col = ( i % 9998 + 1,i // 9998 + 1)
            sheet_coords = convert_to_excel_coordinates(row,col)
            wb.set_cell_contents("test_sheet", f"{sheet_coords}", f"={old_sheet_coords}")
            old_sheet_coords = sheet_coords
    with cp.Profile() as profile:
        for i in range(1, ITERATIONS):
            wb.set_cell_contents("test_sheet", "A1", f"={old_sheet_coords}")
            wb.set_cell_contents("test_sheet", "A1", f"=1")


    return profile


# @profile_it
# Get this profiler test working at somepoint.
# def pro_many_cycles():
#     wb = Workbook()
#     wb.new_sheet("test_sheet")
#     wb.set_cell_contents("test_sheet", f"A1", "1")
#     last_element = []
#     for j in range(100):
#         old_sheet_coords = convert_to_excel_coordinates(1, 1)
#         for i in range(1, 100):
#             row, col = ((j*100 + i) % 9998 + 1, (j*100 + i) // 9998 + 1)
#             sheet_coords = convert_to_excel_coordinates(row, col)
#             wb.set_cell_contents("test_sheet", f"{sheet_coords}", f"={old_sheet_coords}")
#             old_sheet_coords = sheet_coords
#         last_element.append(old_sheet_coords)
#     with cp.Profile() as profile:
#         for i in range(1, 10000):
#             wb.set_cell_contents("test_sheet", "A1", f"={old_sheet_coords}")
#             wb.set_cell_contents("test_sheet", "A1", "=("+"+".join(old_sheet_coords) + f")/100")
#     return profile

@profile_it
def pro_copy_massive_sheet():
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet", f"A1", "1")
    STRING_LENGTH = 1000
    for i in range(1, 100000):
        row, col = (i % 9998 + 1, i // 9998 + 1)
        sheet_coords = convert_to_excel_coordinates(row, col)
        wb.set_cell_contents("test_sheet", f"{sheet_coords}", generate_random_string(STRING_LENGTH))
    with cp.Profile() as profile:
        wb.copy_sheet("test_sheet")
    return profile

# Function to generate Pascal's triangle
def generate_pascals_triangle(n: int) -> List[List[int]]:
    triangle = [[1]]
    for i in range(1, n):
        row = [1]
        for j in range(1, i):
            row.append(triangle[i-1][j-1] + triangle[i-1][j])
        row.append(1)
        triangle.append(row)
    return triangle

# Test to check the speed of updating formulas in a massive Pascal's triangle
def pro_pascals_triangle_speed():
    n = 50
    workbook = Workbook()
    sheet_name = "PascalsTriangle"
    workbook.new_sheet(sheet_name)

    # Generate Pascal's triangle
    triangle = generate_pascals_triangle(n)

    # Set up formulas for each cell in the spreadsheet
    for i in range(n):
        for j in range(i+1):
            cell_location = f"{chr(ord('A')+(j % 26))}{i+1}"
            col_index = j // 26
            if col_index > 0:
                cell_location = f"{chr(ord('A')+(col_index-1))}{cell_location}"
            formula = f"={sheet_name}!{chr(ord('A')+(j % 26))}{i}" if j > 0 else "=1"
            workbook.set_cell_contents(sheet_name, cell_location, formula)

    # Start the timer
    with cp.Profile() as profile:
        # Update the bottom of the triangle with the actual values
        for i in range(n):
            for j in range(i+1):
                cell_location = f"{chr(ord('A')+(j % 26))}{i+1}"
                col_index = j // 26
                if col_index > 0:
                    cell_location = f"{chr(ord('A')+(col_index-1))}{cell_location}"
                value = triangle[i][j]
                workbook.set_cell_contents(sheet_name, cell_location, str(value))

    return profile

@profile_it
def pro_copy_massive_sheet2():
    # Specific case where no sheet is referencing this sheet
    # lets us cache our results
    wb = Workbook()
    wb.new_sheet("test_sheet")
    wb.set_cell_contents("test_sheet", "A1", "1")
    for i in range(1, 9999):
        wb.set_cell_contents("test_sheet", f"A{i+1}", f"=A{i}+1")
    with cp.Profile() as profile:
        wb.copy_sheet("test_sheet")
    assert wb.get_cell_value("test_sheet_1", "A9999") == 9999
    return profile



from datetime import datetime
import sys
import os
if __name__ == "__main__":

    old_stdout = sys.stdout
    save_path = os.path.join("profile_logs", f"profile_{str(datetime.now()).replace(' ' , '_' ).replace(':', ';')}.txt")
    log_file = open(save_path, "w+")
    name = input("who is running this? ")

    sys.stdout = log_file
    print(f"beginning of profiler log written on {datetime.now()} ran on the computer of {name}")

    # Iterate over all items in the global scope
    func_list = list(globals().items())
    for name, obj in func_list:
        # Check if the object is a callable (i.e., a function), but not a built-in function
        if callable(obj) and name.startswith("pro") and not name == "profile_it" and not name.startswith("__"):
            sys.stdout = old_stdout
            print(f"running {name}")
            sys.stdout = log_file
            print(f"\n##################Running {name}... #######################")
            profile = obj()
            print("Total Logs")
            profile.print_stats('tottime')
    log_file.close()