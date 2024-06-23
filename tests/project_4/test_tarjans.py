import context
import pytest
import sheets
import sheets.CellErrorType as CellErrorType

@pytest.fixture
def wb():
    wb = sheets.Workbook()
    wb.new_sheet("Sheet1")
    wb.new_sheet("Sheet2")
    return wb

# def test_move_cells_simple(wb):
#     wb.set_cell_contents("Sheet1", "A1", "=A2 + A4")
#     wb.set_cell_contents("Sheet1", "A2", "=A3 + A5")
#     wb.set_cell_contents("Sheet1", "A3", "=A4")
#     wb.set_cell_contents("Sheet1", "A4", "5")
#     wb.set_cell_contents("Sheet1", "A5", "=A2+A1")
#     results, cycle = wb.graph.tarjans("sheet1", "a2")
#     print(results)
#     assert results == [[('sheet1', 'a2'), ('sheet1', 'a1'), ('sheet1', 'a5')]]
#     assert 1 == 2

def test_move_cells_simple(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet1", "A2", "=A3")
    wb.set_cell_contents("Sheet1", "A3", "5")
    wb.set_cell_contents("Sheet1", "A4", "=A1")
    results, cycle = wb.graph.tarjans("sheet1", "a2")
    print(results)
    assert results == [[('sheet1', 'a2')], [('sheet1', 'a1')], [('sheet1', 'a4')]]

def test_move_cells_simple2(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet1", "A2", "=A3")
    wb.set_cell_contents("Sheet1", "A3", "5")
    wb.set_cell_contents("Sheet1", "A4", "=A1")
    results, cycle = wb.graph.tarjans("sheet1", "a1")
    print(results)
    assert results == [[('sheet1', 'a1')], [('sheet1', 'a4')]]


def test_move_cells_simple3(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet1", "A2", "=A3")
    wb.set_cell_contents("Sheet1", "A3", "5")
    wb.set_cell_contents("Sheet1", "A4", "=A1")
    results, cycle = wb.graph.tarjans("sheet1", "a1")
    print(results)
    assert results == [[('sheet1', 'a1')], [('sheet1', 'a4')]]


def test_move_cells_simple4(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet2", "A3", "=Sheet1!A1")
    results, cycle = wb.graph.tarjans("sheet1", "a1")
    print(results)
    assert results == [[('sheet1', 'a1')], [('sheet2', 'a3')]]
    

def test_move_cells_simple5(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet2", "A3", "=Sheet1!A1")
    results, cycle = wb.graph.tarjans("sheet1", "a1")
    print(results)
    assert results == [[('sheet1', 'a1')], [('sheet2', 'a3')]]

def test_move_cells_simple6(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    wb.set_cell_contents("Sheet2", "A3", "=Sheet1!A1")
    wb.set_cell_contents("Sheet1", "A1", "=Sheet2!A3")
    results, cycle = wb.graph.tarjans("sheet1", "a1")
    print(results)
    assert results == [[('sheet1', 'a1'), ('sheet2', 'a3')]]

def test_move_cells_simple7(wb):
    wb.set_cell_contents("Sheet1", "A1", "=A2")
    results, cycle = wb.graph.tarjans("sheet1", "a2")
    print(results)
    assert results == [[('sheet1', 'a2')],[('sheet1', 'a1')]]


def test_move_cells_simple8(wb):
    results, cycle = wb.graph.tarjans("sheet1", "a1")
    print("Test 8: ", results)
    assert results == [[]]



# def test_move_cells_simple9(wb):
#     wb.set_cell_contents("Sheet1", "A1", "=A2")
#     wb.set_cell_contents("Sheet1", "A3", "=Sheet1!A1")
#     wb.set_cell_contents("Sheet1", "A1", "=Sheet2!A3")
#     results, cycle = wb.graph.tarjans("sheet1", "a1")
#     print(results)
#     assert results == [[('sheet1', 'a1'), ('sheet2', 'a3')]]