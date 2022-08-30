from collections import defaultdict, namedtuple
import itertools
import typing

from cached_property import cached_property

from .. import utils, consts


Size = namedtuple("Size", "height,width")


class Cell:
    def __init__(self, row, col, style, value, row_height, col_width):
        self.row = row
        self.col = col
        self.value = value
        self.style = style
        self.row_height = row_height
        self.col_width = col_width

    def move(self, row, col):
        self.row += row
        self.col += col

    def __eq__(self, other):
        return (
            self.row == other.row
            and self.col == other.col
            and self.value == other.value
            and self.style == other.style
        )

    def __str__(self):  # pragma: no cover
        return "Cell({}, {}, {})".format(self.row, self.col, self.value)

    def __repr__(self):  # pragma: no cover
        return str(self)


class Merge:
    def __init__(self, row, col, rows, cols):
        self.row, self.col = row, col
        self.rows, self.cols = rows, cols

    def __str__(self):
        return "{}{}:{}{}"

    def move(self, row, col):
        self.row += row
        self.col += col


class FuncArg:
    def __init__(self, start_index, end_index, cells, direction=None):
        self.start_index = start_index
        self.end_index = end_index
        self.cells = cells
        self.final_cells = []
        self.direction = direction

    def finalize_cells(self, current_row, current_col, initial_cell, final_cells):
        self.cells.remove(initial_cell)
        if self.direction is not None:
            if self.direction == consts.FuncArgDirection.HORIZONTAL:
                cells = [
                    (cell.row, cell.col)
                    for cell in final_cells
                    if cell.row == current_row
                ]
            elif self.direction == consts.FuncArgDirection.VERTICAL:
                cells = [
                    (cell.row, cell.col)
                    for cell in final_cells
                    if cell.col == current_col
                ]
        else:
            cells = [(cell.row, cell.col) for cell in final_cells]
        self.final_cells.extend(cells)


class FuncCell:
    def __init__(
        self,
        row,
        col,
        style,
        initial_value,
        row_height,
        col_width,
        args,
        default_value="",
    ):
        self.row = row
        self.col = col
        self.style = style
        self.initial_value = initial_value
        self.row_height = row_height
        self.col_width = col_width
        self.default_value = default_value
        self.args = {(arg.start_index, arg.end_index): arg for arg in args}
        self.final_args = []

    def move(self, row, col):
        self.row += row
        self.col += col
        for arg in self.final_args:
            arg.final_cells = [
                (arg_row + row, arg_col + col) for arg_row, arg_col in arg.final_cells
            ]

    def finalize_arg(self, arg_key, initial_cell, final_cells):
        arg = self.args[arg_key]
        arg.finalize_cells(self.row, self.col, initial_cell, final_cells)
        if not arg.cells:
            arg = self.args.pop(arg_key)
            self.final_args.append(arg)

    def get_final_value(self):
        str_args = []
        if not self.final_args:
            return self.default_value
        for arg in self.final_args:
            cells = list(sorted(arg.final_cells))
            if not cells:
                return self.default_value
            rectangle_cell_count = (cells[-1][0] - cells[0][0] + 1) * (
                cells[-1][1] - cells[0][1] + 1
            )
            if len(cells) > 1 and len(cells) == rectangle_cell_count:
                start_cell = cells[0]
                start_cell = utils.cell_int_to_str(start_cell[0], start_cell[1])
                end_cell = cells[-1]
                end_cell = utils.cell_int_to_str(end_cell[0], end_cell[1])
                str_value = "{}:{}".format(start_cell, end_cell)
            else:
                str_value = ",".join(
                    utils.cell_int_to_str(cell[0], cell[1]) for cell in cells
                )
            str_args.append((arg.start_index, arg.end_index, str_value))
        value = self.initial_value
        value_parts = []
        value_part_start_index = 0
        for arg in str_args:
            value_parts.append(value[value_part_start_index : arg[0]])
            value_parts.append(arg[2])
            value_part_start_index = arg[1]
        value_parts.append(value[value_part_start_index:])
        return "".join(value_parts)


class CellGroupFinalResult:
    def __init__(self, cells, func_cells, merges, size):
        self.cells = cells
        self.func_cells = func_cells
        self.merges = merges
        self.size = size

    def get_simple_display(self):
        result = [[None] * self.size.width for _ in range(self.size.height)]
        for cell in self.cells:
            result[cell.row][cell.col] = cell.value
        return result


class BaseCellGroup:
    def add_cell(self, row, col, cell):  # pragma: no cover
        raise NotImplementedError()

    def add_cell_group(self, row, col, cell_group):  # pragma: no cover
        raise NotImplementedError()

    @cached_property
    def final_result(self):
        return self.get_final_result()

    def get_final_result(self):  # pragma: no cover
        raise NotImplementedError()

    def get_final_size(self):
        return self.final_result.size

    def get_final_cells(self):
        return self.final_result.cells

    def get_final_func_cells(self):
        return self.final_result.func_cells

    def get_final_merges(self):
        return self.final_result.merges

    def calc_last_cell(self, final_cells, final_func_cells):
        last_cell = (-1, -1)
        for cells in itertools.chain(final_cells.values(), final_func_cells.values()):
            for cell in cells:
                last_cell = (max(last_cell[0], cell.row), max(last_cell[1], cell.col))
        return last_cell


class CellGroup(BaseCellGroup):
    def __init__(self, initial_size: Size):
        self.initial_size = initial_size
        self.cells = []
        self.func_cells = []
        self.cell_groups = {}
        self.merges = []

    def add_merge(self, row, col, rows, cols):
        self.merges.append(Merge(row, col, rows, cols))

    def add_cell(self, cell):
        self.cells.append(cell)

    def add_func_cell(self, cell):
        self.func_cells.append(cell)

    def add_cell_group(self, row, col, cell_group):
        self.cell_groups[(row, col)] = cell_group

    def get_final_result(self):
        row_offsets = [
            [None] * self.initial_size.width
            for _ in range(self.initial_size.height + 1)
        ]
        col_offsets = [
            [None] * self.initial_size.height
            for _ in range(self.initial_size.width + 1)
        ]
        for (row, col), cell_group in self.cell_groups.items():
            final_size = cell_group.get_final_size()
            if final_size.width > cell_group.initial_size.width:
                for i in range(cell_group.initial_size.width):
                    col_offsets[col + i][row] = 0
                col_offsets[col + 1][row] = (
                    final_size.width - cell_group.initial_size.width
                )
            else:
                for i in range(col, col + final_size.width):
                    col_offsets[i][row] = max(col_offsets[i][row] or 0, 0)
                for i in range(
                    col + final_size.width, col + cell_group.initial_size.width
                ):
                    col_offsets[i][row] = -1

            if final_size.height > cell_group.initial_size.height:
                for i in range(cell_group.initial_size.height):
                    row_offsets[row + i][col] = 0
                row_offsets[row + 1][col] = (
                    final_size.height - cell_group.initial_size.height
                )
            else:
                for i in range(row, row + final_size.height):
                    row_offsets[i][col] = max(row_offsets[i][col] or 0, 0)
                for i in range(
                    row + final_size.height, row + cell_group.initial_size.height
                ):
                    row_offsets[i][col] = -1

        for cell in itertools.chain(self.cells, self.func_cells):
            row_offsets[cell.row][cell.col] = max(
                row_offsets[cell.row][cell.col] or 0, 0
            )
            col_offsets[cell.col][cell.row] = max(
                col_offsets[cell.col][cell.row] or 0, 0
            )

        row_offsets = [
            [offset for offset in row if offset is not None] for row in row_offsets
        ]
        row_offsets = [max(row) if row else 0 for row in row_offsets]
        row_offsets = list(itertools.accumulate(row_offsets))

        col_offsets = [
            [offset for offset in col if offset is not None] for col in col_offsets
        ]
        col_offsets = [max(col) if col else 0 for col in col_offsets]
        col_offsets = list(itertools.accumulate(col_offsets))

        final_cells = defaultdict(list)
        final_func_cells = defaultdict(list)
        final_merges = []

        for (cg_row, cg_col), cell_group in self.cell_groups.items():
            for (c_row, c_col), cells in cell_group.get_final_cells().items():
                for cell in cells:
                    cell.move(
                        cg_row + row_offsets[cg_row], cg_col + col_offsets[cg_col]
                    )
                    final_cells[(cg_row + c_row, cg_col + c_col)].append(cell)
            for (c_row, c_col), cells in cell_group.get_final_func_cells().items():
                for cell in cells:
                    cell.move(
                        cg_row + row_offsets[cg_row], cg_col + col_offsets[cg_col]
                    )
                    final_func_cells[(cg_row + c_row, cg_col + c_col)].append(cell)
            for merge in cell_group.get_final_merges():
                merge.move(cg_row + row_offsets[cg_row], cg_col + col_offsets[cg_col])
                final_merges.append(merge)

        for cell in self.cells:
            row, col = cell.row, cell.col
            cell.move(row_offsets[row], col_offsets[col])
            final_cells[(row, col)].append(cell)

        for cell in self.func_cells:
            row, col = cell.row, cell.col
            cell.move(row_offsets[row], col_offsets[col])
            final_func_cells[(row, col)].append(cell)

        for merge in self.merges:
            row, col = merge.row, merge.col
            merge.move(row_offsets[row], col_offsets[col])
            final_merges.append(merge)

        for (row, col), cells in final_func_cells.items():
            finalized_args = []
            for arg_key, arg in cells[0].args.items():
                for cell in arg.cells:
                    final_row = row + cell[0]
                    final_col = col + cell[1]
                    if (
                        final_row >= 0
                        and final_row < self.initial_size.height
                        and final_col >= 0
                        and final_col < self.initial_size.width
                    ):
                        finalized_args.append((arg_key, cell))
            for arg_key, initial_cell in finalized_args:
                for cell in cells:
                    final_row = initial_cell[0] + row
                    final_col = initial_cell[1] + col
                    cell.finalize_arg(
                        arg_key,
                        initial_cell,
                        itertools.chain(
                            final_cells.get((final_row, final_col), []),
                            final_func_cells.get((final_row, final_col), []),
                        ),
                    )

        last_cell = self.calc_last_cell(final_cells, final_func_cells)

        return CellGroupFinalResult(
            cells=final_cells,
            size=Size(width=last_cell[1] + 1, height=last_cell[0] + 1),
            func_cells=final_func_cells,
            merges=final_merges,
        )


class SheetCellGroup(CellGroup):
    def get_final_result(self):
        result = super().get_final_result()
        cells = [cell for _cells in result.cells.values() for cell in _cells]
        func_cells = [cell for _cells in result.func_cells.values() for cell in _cells]
        merges = result.merges

        for cell in itertools.chain(cells, func_cells, merges):
            cell.move(1, 1)

        func_cells = [
            Cell(
                cell.row,
                cell.col,
                cell.style,
                cell.get_final_value(),
                cell.row_height,
                cell.col_width,
            )
            for cell in func_cells
        ]

        return CellGroupFinalResult(
            cells=cells + func_cells,
            func_cells=[],
            merges=merges,
            size=Size(result.size.height + 1, result.size.width + 1),
        )


class LoopCellGroup(BaseCellGroup):
    def __init__(self, initial_size, direction):
        self.initial_size = initial_size
        self.direction = direction
        self.cell_groups = []

    def add_cell_group(self, cell_group):
        self.cell_groups.append(cell_group)

    def get_final_result(self):
        row_offsets = [0]
        col_offsets = [0]
        if self.direction == consts.LoopDirection.DOWN:
            for cell_group in self.cell_groups:
                final_size = cell_group.get_final_size()
                row_offsets.append(final_size.height)
            col_offsets = [0] * len(row_offsets)
        else:
            for cell_group in self.cell_groups:
                final_size = cell_group.get_final_size()
                col_offsets.append(final_size.width)
            row_offsets = [0] * len(col_offsets)

        row_offsets = list(itertools.accumulate(row_offsets))
        col_offsets = list(itertools.accumulate(col_offsets))
        final_cells = defaultdict(list)
        final_func_cells = defaultdict(list)
        final_merges = []
        for index, cell_group in enumerate(self.cell_groups):
            for (row, col), cells in cell_group.get_final_cells().items():
                for cell in cells:
                    cell.move(row_offsets[index], col_offsets[index])
                    final_cells[(row, col)].append(cell)
            for (row, col), cells in cell_group.get_final_func_cells().items():
                for cell in cells:
                    cell.move(row_offsets[index], col_offsets[index])
                    final_func_cells[(row, col)].append(cell)
            for merge in cell_group.get_final_merges():
                merge.move(row_offsets[index], col_offsets[index])
                final_merges.append(merge)

        last_cell = self.calc_last_cell(final_cells, final_func_cells)

        return CellGroupFinalResult(
            cells=final_cells,
            size=Size(width=last_cell[1] + 1, height=last_cell[0] + 1),
            func_cells=final_func_cells,
            merges=final_merges,
        )
