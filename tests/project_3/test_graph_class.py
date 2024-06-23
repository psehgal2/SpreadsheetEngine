import context
from sheets.Graph import Graph


def test_graph_init():
    g = Graph()
    assert g.children_to_parents == {}
    assert g.parents_to_children == {}
    # assert g.sheet_map == {}

def test_graph_add_connection():
    g = Graph()
    # case of B1 = A1
    g.add_connection("Sheet1", "A1", "Sheet2", "B1")
    assert g.children_to_parents == {"Sheet2": {"B1": [("Sheet1", "A1")]}}

