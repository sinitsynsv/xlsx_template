import string
import collections


def col_str_to_int(col):
    final_column = 0
    for letter in col:
        final_column = final_column * 26 + string.ascii_uppercase.index(letter) + 1
    return final_column


def cell_str_to_int(row, col):
    row = int(row)
    return (row, col_str_to_int(col))


def cell_int_to_str(row, col):
    return col_int_to_str(col) + str(row)


def col_int_to_str(col):
    res = ""
    while col > 0:
        col, reminder = divmod(col, 26)
        if reminder == 0:
            reminder = 26
            col -= 1
        res = string.ascii_uppercase[reminder - 1] + res
    return res


# Cell = collections.namedtuple("Cell", "row,col")

# class Cell:
#     def __init__(self, row, col):
#         self.row = row
#         self.col = col

#     def adjust(self, row, col):
#         self.row += row
#         self.col += col

#     def __eq__(self, other):
#         return self.row == other.row and self.col == other.col

#     def __hash__(self):
#         return (self.row, self.col).__hash__()

# def __str__(self):
#     return cell_int_to_str(self.row, self.col)


# class CellRange:
#     def __init__(self, start_cell, end_cell):
#         self.start_cell = start_cell
#         self.end_cell = end_cell
