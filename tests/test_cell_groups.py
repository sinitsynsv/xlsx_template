from xlsx_template.runtime.cell_groups import (
    CellGroup,
    Cell,
    LoopCellGroup,
    FuncArg,
    FuncCell,
    Size,
    SheetCellGroup,
)
from xlsx_template.consts import LoopDirection


def test_cell_group_with_cells():
    cell_group = CellGroup(initial_size=Size(height=1, width=3))
    cell_group.add_cell(Cell(0, 0, "s1", 1, 0, 0))
    cell_group.add_cell(Cell(0, 1, "s1", 2, 0, 0))
    cell_group.add_cell(Cell(0, 2, "s1", 3, 0, 0))
    valid_result = {
        (0, 0): [Cell(0, 0, "s1", 1, 0, 0)],
        (0, 1): [Cell(0, 1, "s1", 2, 0, 0)],
        (0, 2): [Cell(0, 2, "s1", 3, 0, 0)],
    }
    assert valid_result == cell_group.get_final_cells()


def test_loop_cell_group():
    cell_group0 = CellGroup(initial_size=Size(height=3, width=3))
    cell_group0.add_cell(Cell(0, 0, "s1", "Header1", 0, 0))
    cell_group0.add_cell(Cell(0, 1, "s1", "Header2", 0, 0))
    cell_group0.add_cell(Cell(0, 2, "s1", "Header3", 0, 0))

    cell_group1 = LoopCellGroup(
        initial_size=Size(height=1, width=3), direction=LoopDirection.DOWN
    )
    for row_index in range(1, 4):
        cell_group2 = CellGroup(initial_size=Size(height=1, width=3))
        cell_group2.add_cell(Cell(0, 0, "s1", row_index * 10, 0, 0))
        cell_group2.add_cell(Cell(0, 1, "s1", row_index * 100, 0, 0))
        cell_group2.add_cell(Cell(0, 2, "s1", row_index * 1000, 0, 0))
        cell_group1.add_cell_group(cell_group2)

    cell_group0.add_cell_group(1, 0, cell_group1)

    cell_group0.add_cell(Cell(2, 0, "s1", "Summary1", 0, 0))
    cell_group0.add_cell(Cell(2, 1, "s1", "Summary2", 0, 0))
    cell_group0.add_cell(Cell(2, 2, "s1", "Summary3", 0, 0))

    valid_cells = {
        (0, 0): [Cell(0, 0, "s1", "Header1", 0, 0)],
        (0, 1): [Cell(0, 1, "s1", "Header2", 0, 0)],
        (0, 2): [Cell(0, 2, "s1", "Header3", 0, 0)],
        (1, 0): [
            Cell(1, 0, "s1", 10, 0, 0),
            Cell(2, 0, "s1", 20, 0, 0),
            Cell(3, 0, "s1", 30, 0, 0),
        ],
        (1, 1): [
            Cell(1, 1, "s1", 100, 0, 0),
            Cell(2, 1, "s1", 200, 0, 0),
            Cell(3, 1, "s1", 300, 0, 0),
        ],
        (1, 2): [
            Cell(1, 2, "s1", 1000, 0, 0),
            Cell(2, 2, "s1", 2000, 0, 0),
            Cell(3, 2, "s1", 3000, 0, 0),
        ],
        (2, 0): [Cell(4, 0, "s1", "Summary1", 0, 0)],
        (2, 1): [Cell(4, 1, "s1", "Summary2", 0, 0)],
        (2, 2): [Cell(4, 2, "s1", "Summary3", 0, 0)],
    }
    final_cells = cell_group0.get_final_cells()
    assert valid_cells == final_cells


def test_table_with_formulas():
    sheet_cell_group = SheetCellGroup(initial_size=Size(3, 5))
    cell_group0 = CellGroup(initial_size=Size(3, 5))

    cell_group0.add_cell(Cell(0, 0, "s1", "Name", 0, 0))
    cell_group0.add_cell(Cell(0, 1, "s1", "Val", 0, 0))
    cell_group0.add_cell(Cell(0, 2, "s1", "Bonus %", 0, 0))
    cell_group0.add_cell(Cell(0, 3, "s1", "Bonus Val", 0, 0))
    cell_group0.add_cell(Cell(0, 4, "s1", "Total", 0, 0))

    cell_group1 = LoopCellGroup(initial_size=Size(1, 5), direction=LoopDirection.DOWN)

    for row_index in range(1, 4):
        cell_group2 = CellGroup(initial_size=Size(1, 5))
        cell_group2.add_cell(Cell(0, 0, "s1", "Name{}".format(row_index), 0, 0))
        cell_group2.add_cell(Cell(0, 1, "s1", row_index * 100, 0, 0))
        cell_group2.add_cell(Cell(0, 2, "s1", row_index * 10, 0, 0))
        func_args = [FuncArg(1, 3, [(0, -2)]), FuncArg(4, 6, [(0, -1)])]
        cell_group2.add_func_cell(FuncCell(0, 3, "s1", "=B2*C2", 0, 0, args=func_args))
        func_args = [FuncArg(1, 3, [(0, -3)]), FuncArg(4, 6, [(0, -1)])]
        cell_group2.add_func_cell(FuncCell(0, 4, "s1", "=B2+D2", 0, 0, args=func_args))
        cell_group1.add_cell_group(cell_group2)

    cell_group0.add_cell_group(1, 0, cell_group1)

    cell_group0.add_cell(Cell(2, 0, "s1", "Total", 0, 0))

    func_args = [FuncArg(5, 7, [(-1, 0)])]
    cell_group0.add_func_cell(FuncCell(2, 1, "s1", "=SUM(B2)", 0, 0, args=func_args))

    func_args = [FuncArg(5, 7, [(-1, 0)])]
    cell_group0.add_func_cell(FuncCell(2, 2, "s1", "=SUM(C2)", 0, 0, args=func_args))

    func_args = [FuncArg(5, 7, [(-1, 0)])]
    cell_group0.add_func_cell(FuncCell(2, 3, "s1", "=SUM(D2)", 0, 0, args=func_args))

    func_args = [FuncArg(5, 7, [(-1, 0)])]
    cell_group0.add_func_cell(FuncCell(2, 4, "s1", "=SUM(E2)", 0, 0, args=func_args))

    sheet_cell_group.add_cell_group(0, 0, cell_group0)

    valid_result = [
        [None, None, None, None, None, None],
        [None, "Name", "Val", "Bonus %", "Bonus Val", "Total"],
        [None, "Name1", 100, 10, "=B2*C2", "=B2+D2"],
        [None, "Name2", 200, 20, "=B3*C3", "=B3+D3"],
        [None, "Name3", 300, 30, "=B4*C4", "=B4+D4"],
        [None, "Total", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)", "=SUM(E2:E4)"],
    ]
    result = sheet_cell_group.final_result.get_simple_display()
    assert result == valid_result


def test_negative_row_offset():
    sheet_cell_group = SheetCellGroup(initial_size=Size(3, 3))
    cell_group0 = CellGroup(initial_size=Size(3, 3))

    cell_group0.add_cell(Cell(0, 0, "s1", "H1", 0, 0))
    cell_group0.add_cell(Cell(0, 1, "s1", "H2", 0, 0))
    cell_group0.add_cell(Cell(0, 2, "s1", "H3", 0, 0))

    cell_group1 = CellGroup(initial_size=Size(1, 3))

    cell_group0.add_cell_group(1, 0, cell_group1)

    func_args = [FuncArg(4, 6, [(-1, 0)])]

    cell_group0.add_func_cell(
        FuncCell(2, 0, "s1", "SUM(A2)", 0, 0, args=func_args, default_value=0)
    )
    cell_group0.add_func_cell(FuncCell(2, 1, "s1", "SUM(B2)", 0, 0, args=func_args))
    cell_group0.add_func_cell(FuncCell(2, 2, "s1", "SUM(C2)", 0, 0, args=func_args))
    sheet_cell_group.add_cell_group(0, 0, cell_group0)
    valid_result = [
        [None, None, None, None],
        [None, "H1", "H2", "H3"],
        [None, 0, "", ""],
    ]
    result = sheet_cell_group.final_result.get_simple_display()
    assert result == valid_result


def test_offsets():
    """
    Test when one group has smaller size, and another bigger
    """
    cell_group0 = CellGroup(Size(2, 2))

    cell_group1 = CellGroup(Size(1, 2))
    cell_group0.add_cell_group(0, 0, cell_group1)

    cell_group1 = CellGroup(Size(1, 2))
    cell_group1.add_cell(Cell(0, 0, "s1", "value", 0, 0))
    cell_group2 = LoopCellGroup(Size(1, 1), LoopDirection.RIGHT)
    for i in range(3):
        cell_group3 = CellGroup(initial_size=Size(1, 1))
        cell_group3.add_cell(Cell(0, 0, "s1", i, 0, 0))
        cell_group2.add_cell_group(cell_group3)
    cell_group1.add_cell_group(0, 1, cell_group2)
    cell_group0.add_cell_group(1, 0, cell_group1)
    valid_result = {
        (1, 0): [Cell(0, 0, "s1", "value", 0, 0)],
        (1, 1): [
            Cell(0, 1, "s1", 0, 0, 0),
            Cell(0, 2, "s1", 1, 0, 0),
            Cell(0, 3, "s1", 2, 0, 0),
        ],
    }
    assert cell_group0.get_final_cells() == valid_result

    cell_group0 = CellGroup(Size(2, 2))
    cell_group1 = CellGroup(Size(2, 1))
    cell_group0.add_cell_group(0, 0, cell_group1)

    cell_group1 = CellGroup(Size(2, 1))
    cell_group1.add_cell(Cell(0, 0, "s1", "value", 0, 0))
    cell_group2 = LoopCellGroup(Size(1, 1), LoopDirection.DOWN)
    for i in range(3):
        cell_group3 = CellGroup(Size(1, 1))
        cell_group3.add_cell(Cell(0, 0, "s1", i, 0, 0))
        cell_group2.add_cell_group(cell_group3)
    cell_group1.add_cell_group(1, 0, cell_group2)
    cell_group0.add_cell_group(0, 1, cell_group1)
    valid_result = {
        (0, 1): [Cell(0, 0, "s1", "value", 0, 0)],
        (1, 1): [
            Cell(1, 0, "s1", 0, 0, 0),
            Cell(2, 0, "s1", 1, 0, 0),
            Cell(3, 0, "s1", 2, 0, 0),
        ],
    }
    assert cell_group0.get_final_cells() == valid_result
