class Node:
    def __init__(self, sheet_location, cell_location) -> None:
        self.children = []  # List of Nodes
        self.parents = []  # List of Nodes
        self.sheet_location = sheet_location
        self.cell_location = cell_location


# A1 = B1
# If you're the forumula cell, you're the parent
# if you are being referenced, you are the child


class Graph:
    def __init__(self):
        # key is child: value is list of parents
        # {Sheet_Name : {location : node}}
        # CHILD LOCATION -> PARENT LOCATION()
        # A1 -> B1 | B1:[A1]
        self.children_to_parents = {}  # Dict{sheet_name: Dict{cell_location: List[(sheet_location_object, cell_location_objects)]}}

        # child -> parent
        # key is parent: value is list of children
        self.parents_to_children = {}  # Dict{sheet_name: Dict{cell_location: List[(sheet_location_object, cell_location_objects)]}}

        # self.sheet_map = {} # Dict{sheet_name: List of Nodes}

    def rename_sheet(self, old_sheet_name, new_sheet_name):
        """
        Does all the necessary updates to change the old_sheet_name to the new_sheet_name.

        """
        # just handles children_to_parents for now
        # collects the children in the new sheet name
        children_of_overlapping_sheet = self.children_to_parents.get(new_sheet_name)
        if old_sheet_name in self.children_to_parents:
            self.children_to_parents[new_sheet_name] = self.children_to_parents.get(
                old_sheet_name
            )
        else:
            self.add_sheet(new_sheet_name)
        # This merging part here is actually pretty expensive
        if children_of_overlapping_sheet is not None:
            for location, parents in children_of_overlapping_sheet.items():
                if location in self.children_to_parents.get(new_sheet_name):
                    self.children_to_parents[new_sheet_name][location].extend(
                        [
                            parent
                            for parent in parents
                            if parent
                            not in self.children_to_parents[new_sheet_name][location]
                        ]
                    )
                else:
                    self.children_to_parents[new_sheet_name][location] = parents

        # Now the children list has been properly renamed and merged with the parent list
        # We need to go through and update the sheet names of the parents.
        for sheet_name, sheet in self.children_to_parents.items():
            for child_loc, list in sheet.items():
                for parent_loc, parent_sheet in list:
                    if parent_sheet == old_sheet_name:
                        list.remove((parent_loc, parent_sheet))
                        list.append((parent_loc, new_sheet_name))

        # Now everything on the children_to_parents side is complete.

    def add_connection(
        self, parent_sheet_name, parent_cell_loc, child_sheet_name, child_cell_loc
    ):
        """Adds a parent in the graph data structure.
        sheet_name: the name of the sheet the referenced cell is in
        cell_loc : the location of the referenced cell
        parent_sheet_name : the name of the sheet of the referencing cell
        parent_cell_loc : the location of the referencing cell"""
        if child_sheet_name not in self.children_to_parents:
            self.children_to_parents.update({child_sheet_name: {}})

        sheet = self.children_to_parents[child_sheet_name]

        if child_cell_loc in sheet:
            sheet[child_cell_loc].append((parent_sheet_name, parent_cell_loc))
        else:
            sheet.update({child_cell_loc: [(parent_sheet_name, parent_cell_loc)]})

        # Uncommented to add to parents
        # if parent_sheet_name not in self.parents_to_children:
        #     self.parents_to_children.update({parent_sheet_name : {}})

        # sheet = self.parents_to_children[child_sheet_name]
        # if parent_cell_loc in sheet:
        #     sheet[parent_cell_loc].append((parent_sheet_name , child_cell_loc))
        # else:
        #     sheet.update({parent_cell_loc : [(child_sheet_name , child_cell_loc)]})

    def update_child(self, hidden_name, location, child_hidden_name, child_location):
        """Updates the references to a cell."""

        # if hidden_name not in self.children_to_parents:
        #     return
        if (hidden_name, location) in self.children_to_parents[child_hidden_name][
            child_location
        ]:
            self.children_to_parents[child_hidden_name][child_location].remove(
                (hidden_name, location)
            )

    # def update_children(self, sheet_name, location):
    #     """Deletes all the references to a cell."""
    #     # for sheet in self.children_to_parents:
    #     #     for cell in self.children_to_parents[sheet]:
    #     #         for (parent_sheet, cell_loc) in self.children_to_parents[sheet][cell]:
    #     #             if (parent_sheet, cell_loc) == (sheet_name, location):
    #     #                 self.children_to_parents[sheet][cell].remove((parent_sheet, cell_loc))
    #     pass
    #     # if sheet_name not in self.children_to_parents or \
    #     # location not in self.children_to_parents[sheet_name]:
    #     #     return
    #     # self.children_to_parents[sheet_name][location] = []

    def has_sheet(self, sheet_name):
        return sheet_name in self.children_to_parents.keys()

    def add_sheet(self, sheet_name):
        """Adds a sheet to the graph."""
        self.children_to_parents[sheet_name] = {}

    def get_sheet_parents(self, sheet_name):
        """Returns the parents for every cell in a given sheet."""
        set_of_parents = set()
        if sheet_name in self.children_to_parents:
            for cell in self.children_to_parents[sheet_name]:
                set_of_parents.update(set(self.children_to_parents[sheet_name][cell]))

        return set_of_parents

    def get_parents_from_cell(self, sheet_name: str, cell_loc: str):
        """Returns the parents of a cell."""
        if sheet_name not in self.children_to_parents:
            return []
        if cell_loc not in self.children_to_parents[sheet_name]:
            return []
        return self.children_to_parents[sheet_name][cell_loc]

    # def get_children_from_cell(self, sheet_name: str, cell_loc: str):
    #     """Returns the children of a cell."""
    #     if sheet_name not in self.parents_to_children:
    #         return []
    #     if cell_loc not in self.parents_to_children[sheet_name]:
    #         return []
    #     return self.parents_to_children[sheet_name][cell_loc]

    def get_circ_refs(self, cycle_check: list):
        """Performs a BFS to find all cells in a cycle."""
        queue = []
        visited = set()
        circ_refs = []

        for sheet_location, cell_location in cycle_check:
            queue.append((sheet_location, cell_location))
            visited.add((sheet_location, cell_location))

        while queue:
            sheet_location, cell_location = queue.pop(0)
            circ_refs.append((sheet_location, cell_location))
            parents = self.get_parents_from_cell(sheet_location, cell_location)
            for parent_sheet, parent_cell in parents:
                if (parent_sheet, parent_cell) not in visited:
                    queue.append((parent_sheet, parent_cell))
                    visited.add((parent_sheet, parent_cell))
        return circ_refs

    def topological_sort(self, start_sheet_name: str, start_cell_location: str):
        """Performs a topological sort of the DAG containing the cell."""

        if start_sheet_name not in self.children_to_parents:
            return False, [(start_sheet_name, start_cell_location)]
        if start_cell_location not in self.children_to_parents[start_sheet_name]:
            return False, [(start_sheet_name, start_cell_location)]
        visited = []
        stack = []
        cycle_check = set()
        cycle_detected = False

        stack.append((start_sheet_name, start_cell_location))
        while stack:
            sheet_location, cell_location = stack[-1]
            cycle_check.add((sheet_location, cell_location))
            # c1 -> c2 <-> c3
            parents_not_visited = []
            for s_name, c_loc in self.get_parents_from_cell(
                sheet_location, cell_location
            ):
                if (s_name, c_loc) not in visited:
                    parents_not_visited.append((s_name, c_loc))

            if parents_not_visited:
                parent = parents_not_visited[0]
                if parent in cycle_check:
                    cycle_detected = True
                    visited.append((sheet_location, cell_location))
                    # return True, self.get_circ_refs(cycle_check)
                stack.append(parent)
            else:
                visited.append((sheet_location, cell_location))
                cycle_check.remove((sheet_location, cell_location))
                stack.pop()
                # Recompute the cell value here
        return cycle_detected, visited[::-1]

    def tarjansCopilot(self, start_sheet_name: str, start_cell_location: str):
        result = []
        stack = []
        index = 0

        # {cell_location : (index, lowlink, onstack)}
        cell_nodes = {}
        cycle_detected = False

        stack.append((start_sheet_name, start_cell_location))

        while stack:
            sheet_name, cell_location = stack[-1]
            curr_cell = (sheet_name, cell_location)

            if curr_cell not in cell_nodes:
                cell_nodes[curr_cell] = (index, index, True)
                index += 1

                for child_sheet, child_cell in self.get_parents_from_cell(
                    sheet_name, cell_location
                ):
                    child = (child_sheet, child_cell)
                    if child == curr_cell and not cycle_detected:
                        cycle_detected = True

                    if child not in cell_nodes:
                        stack.append(child)
                        break
                    elif child in stack:
                        cycle_detected = True
                        lowlink = min(cell_nodes[curr_cell][1], cell_nodes[child][0])
                        cell_nodes[curr_cell] = (
                            cell_nodes[curr_cell][0],
                            lowlink,
                            True,
                        )
            else:
                stack.pop()
                if stack:
                    parent = stack[-1]
                    lowlink = min(cell_nodes[parent][1], cell_nodes[curr_cell][1])
                    cell_nodes[parent] = (cell_nodes[parent][0], lowlink, True)

                # if lowlink = index, pop the stack and generate the SCC
                if cell_nodes[curr_cell][0] == cell_nodes[curr_cell][1]:
                    scc = []
                    while stack and stack[-1] != curr_cell:
                        scc.append(stack.pop())
                    scc.append(curr_cell)
                    result.append(scc[::-1])

        return result[::-1], cycle_detected

    # REC OG Tarjans
    def tarjansREC(self, start_sheet_name: str, start_cell_location: str):
        if start_sheet_name not in self.children_to_parents:
            raise ValueError("Sheet does not exist")
        if start_cell_location not in self.children_to_parents[start_sheet_name]:
            raise ValueError("Cell does not exist")

        result = []
        stack = []
        index = 0

        # {cell_location : (index, lowlink, onstack)}
        cell_nodes = {}
        cycle_detected = False

        def strongConnect(sheet_name, cell_location):
            nonlocal index
            nonlocal result
            nonlocal stack
            nonlocal cell_nodes
            nonlocal cycle_detected

            if (sheet_name, cell_location) not in cell_nodes:
                cell_nodes[(sheet_name, cell_location)] = (index, index, True)
                index += 1
                stack.append((sheet_name, cell_location))

                curr_cell = (sheet_name, cell_location)

                for child_sheet, child_cell in self.get_parents_from_cell(
                    sheet_name, cell_location
                ):
                    child = (child_sheet, child_cell)
                    if child == curr_cell and not cycle_detected:
                        cycle_detected = True

                    if (child_sheet, child_cell) not in cell_nodes:
                        strongConnect(child_sheet, child_cell)
                        lowlink = min(cell_nodes[curr_cell][1], cell_nodes[child][1])
                        cell_nodes[curr_cell] = (
                            cell_nodes[curr_cell][0],
                            lowlink,
                            True,
                        )
                    elif child in stack:
                        cycle_detected = True
                        lowlink = min(cell_nodes[curr_cell][1], cell_nodes[child][0])
                        cell_nodes[curr_cell] = (
                            cell_nodes[curr_cell][0],
                            lowlink,
                            True,
                        )

                # if lowlink = index, pop the stack and generate the SCC
                if cell_nodes[curr_cell][0] == cell_nodes[curr_cell][1]:
                    scc = []
                    while stack[-1] != curr_cell:
                        scc.append(stack.pop())  # NEED TO HAVE NAMING CORRECT
                    scc.append(stack.pop())
                    result.append(scc[::-1])

        strongConnect(start_sheet_name, start_cell_location)

        return result[::-1], cycle_detected

    def tarjans(self, start_sheet_name: str, start_cell_location: str):
        if start_sheet_name not in self.children_to_parents:
            return [[]], False
        if start_cell_location not in self.children_to_parents[start_sheet_name]:
            return [[]], False

        result = []
        stack = []
        index = 0
        call_stack = []

        # {cell_location : (index, lowlink, onstack)}
        cell_nodes = {}
        cycle_detected = False
        call_stack.append((start_sheet_name, start_cell_location, 0))

        while call_stack:
            sheet_name, cell_location, pi = call_stack.pop()
            curr_cell = (sheet_name, cell_location)
            # pi = 0 means seen for the first time
            if pi == 0:
                if (sheet_name, cell_location) not in cell_nodes:
                    cell_nodes[(sheet_name, cell_location)] = (index, index, True)
                    index += 1
                    stack.append((sheet_name, cell_location))
                    call_stack.append((sheet_name, cell_location, 1))

                    for child_sheet, child_cell in self.get_parents_from_cell(
                        sheet_name, cell_location
                    ):
                        child = (child_sheet, child_cell)
                        if child == curr_cell and not cycle_detected:
                            cycle_detected = True
                        if (child_sheet, child_cell) not in cell_nodes:
                            call_stack.append((child_sheet, child_cell, 0))
                        elif child in stack:
                            cycle_detected = True
                            lowlink = min(
                                cell_nodes[curr_cell][1], cell_nodes[child][0]
                            )
                            cell_nodes[curr_cell] = (
                                cell_nodes[curr_cell][0],
                                lowlink,
                                True,
                            )

                else:
                    continue
            elif pi > 0:
                for child_sheet, child_cell in self.get_parents_from_cell(
                    sheet_name, cell_location
                ):
                    child = (child_sheet, child_cell)
                    if child == curr_cell and not cycle_detected:
                        cycle_detected = True
                    if child in stack:
                        cell_nodes[curr_cell] = (
                            cell_nodes[curr_cell][0],
                            lowlink,
                            True,
                        )
                    if child in stack:
                        lowlinkv1 = min(cell_nodes[curr_cell][1], cell_nodes[child][1])
                        lowlinkv2 = min(cell_nodes[curr_cell][1], cell_nodes[child][0])
                        lowlink = min(lowlinkv1, lowlinkv2)
                        cell_nodes[curr_cell] = (
                            cell_nodes[curr_cell][0],
                            lowlink,
                            True,
                        )

                # if lowlink = index, pop the stack and generate the SCC
                if cell_nodes[curr_cell][0] == cell_nodes[curr_cell][1]:
                    scc = []
                    while stack[-1] != curr_cell:
                        scc.append(stack.pop())  # NEED TO HAVE NAMING CORRECT
                    scc.append(stack.pop())
                    result.append(scc[::-1])
        return result[::-1], cycle_detected
