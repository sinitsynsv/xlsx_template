import io
from collections import defaultdict

from xlsx_template import nodes


class Symbols:

    BUILTINS = ("range", "len")

    def __init__(self):
        self.symbols = defaultdict(list)

    def declare_ref(self, name):
        self.symbols[name].append("{}_{}".format(name, len(self.symbols[name])))
        return self.symbols[name][-1]

    def undeclare_ref(self, name):
        self.symbols[name].pop(-1)
        if not self.symbols[name]:
            self.symbols.pop(name)

    def add_ref(self, name, real_name):
        self.symbols[name].append(real_name)

    def find_ref(self, name):
        if name in self.symbols:
            return self.symbols[name][-1]
        elif name in self.BUILTINS:
            return name


class CodeGenerator:
    def __init__(self):
        self.indent_count = 0
        self.stream = None
        self.symbols = None
        self.cell_group_level = 0
        self.is_new_line = True

    def generate(self, root_node):
        self.indent_count = 0
        self.stream = io.StringIO()
        self.symbols = Symbols()
        self.is_new_line = True
        self.cell_group_level = 0
        self.generate_for(root_node)
        return self.stream.getvalue()

    def write(self, s):
        if self.is_new_line:
            self.stream.write("    " * self.indent_count)
            self.is_new_line = False
        self.stream.write(s)

    def newline(self):
        self.stream.write("\n")
        self.is_new_line = True

    def write_line(self, s):
        self.write(s)
        self.newline()

    def indent(self):
        self.indent_count += 1

    def unindent(self):
        self.indent_count -= 1

    def generate_for(self, node):
        method_name = "generate_for_{}".format(node.__class__.__name__.lower())
        return getattr(self, method_name)(node)

    def generate_for_sheet(self, sheet_node):
        self.write("sheet = wb.create_sheet(")
        self.generate_for(sheet_node.name)
        self.write(")")
        self.newline()
        size = "cg.Size({}, {})".format(
            sheet_node.last_cell[0] + 1, sheet_node.last_cell[1] + 1
        )
        self.write_line("cell_group_0 = cg.SheetCellGroup({})".format(size))
        self.newline()
        for child in sheet_node.body:
            self.generate_for(child)
        self.newline()
        self.write_line("for f_cell in cell_group_0.get_final_cells():")
        self.indent()
        self.write_line("cell = sheet.cell(f_cell.row, f_cell.col, f_cell.value)")
        self.write_line("if f_cell.style is not None:")
        self.indent()
        self.write_line("cell.style = f_cell.style")
        self.unindent()
        self.write_line("sheet.row_dimensions[f_cell.row].height = f_cell.row_height")
        self.write_line(
            "sheet.column_dimensions[utils.col_int_to_str(f_cell.col)].width = f_cell.col_width"
        )
        self.unindent()
        self.write_line("for m in cell_group_0.get_final_merges():")
        self.indent()
        self.write_line(
            "sheet.merge_cells(start_row=m.row, start_column=m.col, end_row=m.row+m.rows-1, end_column=m.col+m.cols-1)"
        )
        self.unindent()

    def generate_for_sheetloop(self, sheet_loop):
        loop_ref = self.symbols.declare_ref("loop")
        if sheet_loop.name:
            self.symbols.add_ref("{}_loop".format(sheet_loop.name), loop_ref)
        self.write("{} = LoopContext(".format(loop_ref))
        self.generate_for(sheet_loop.items)
        self.write_line(")")
        target_ref = self.symbols.declare_ref(sheet_loop.target)
        self.write_line("for {} in {}:".format(target_ref, loop_ref))
        self.indent()
        self.generate_for_sheet(sheet_loop.sheet)
        self.unindent()
        self.symbols.undeclare_ref("loop")
        self.symbols.undeclare_ref(sheet_loop.target)

    def generate_for_remove(self, remove):
        size = "cg.Size({}, {})".format(remove.height, remove.width)
        self.write_line(
            "cell_group_{} = cg.CellGroup(initial_size={})".format(
                self.cell_group_level + 1, size
            )
        )
        self.write_line(
            "cell_group_{}.add_cell_group({}, {}, cell_group_{})".format(
                self.cell_group_level,
                remove.base_cell[0],
                remove.base_cell[1],
                self.cell_group_level + 1,
            )
        )

    def generate_for_if(self, if_d):
        self.cell_group_level += 1
        size = "cg.Size({}, {})".format(if_d.height, if_d.width)
        self.write_line(
            "cell_group_{} = cg.CellGroup(initial_size={})".format(
                self.cell_group_level, size
            )
        )
        self.write("if ")
        self.generate_for(if_d.condition)
        self.write_line(":")
        self.indent()
        for node in if_d.body:
            self.generate_for(node)
        self.unindent()
        if if_d.else_block:
            self.write_line("else:")
            self.indent()
            for node in if_d.else_block:
                self.generate_for(node)
            self.unindent()
        self.write_line(
            "cell_group_{}.add_cell_group({}, {}, cell_group_{})".format(
                self.cell_group_level - 1,
                if_d.base_cell[0],
                if_d.base_cell[1],
                self.cell_group_level,
            )
        )
        self.cell_group_level -= 1

    def generate_for_cellgroup(self, cell_group):
        self.cell_group_level += 1
        size = "cg.Size({}, {})".format(cell_group.height, cell_group.width)
        self.write_line(
            "cell_group_{} = cg.CellGroup(initial_size={})".format(
                self.cell_group_level, size
            )
        )
        for node in cell_group.body:
            self.generate_for(node)
        self.write_line(
            "cell_group_{}.add_cell_group({}, {}, cell_group_{})".format(
                self.cell_group_level - 1,
                cell_group.base_cell[0],
                cell_group.base_cell[1],
                self.cell_group_level,
            )
        )
        self.cell_group_level -= 1

    def generate_for_cellloop(self, cell_loop):
        self.cell_group_level += 1
        size = "cg.Size({}, {})".format(cell_loop.height, cell_loop.width)
        self.write_line(
            "cell_group_{} = cg.LoopCellGroup(initial_size={}, direction={})".format(
                self.cell_group_level, size, cell_loop.direction
            )
        )
        loop_ref = self.symbols.declare_ref("loop")
        if cell_loop.name:
            self.symbols.add_ref("{}_loop".format(cell_loop.name), loop_ref)
        self.write("{} = LoopContext(".format(loop_ref))
        self.generate_for(cell_loop.items)
        self.write_line(")")
        target_ref = self.symbols.declare_ref(cell_loop.target)
        self.write_line("for {} in {}:".format(target_ref, loop_ref))
        self.indent()
        self.cell_group_level += 1
        self.write_line(
            "cell_group_{} = cg.CellGroup(initial_size={})".format(
                self.cell_group_level, size
            )
        )
        for node in cell_loop.body:
            self.generate_for(node)
        self.write_line(
            "cell_group_{}.add_cell_group(cell_group_{})".format(
                self.cell_group_level - 1, self.cell_group_level
            )
        )
        self.cell_group_level -= 1
        self.unindent()
        self.write_line(
            "cell_group_{}.add_cell_group({}, {}, cell_group_{})".format(
                self.cell_group_level - 1,
                cell_loop.base_cell[0],
                cell_loop.base_cell[1],
                self.cell_group_level,
            )
        )
        self.cell_group_level -= 1
        self.symbols.undeclare_ref("loop")
        self.symbols.undeclare_ref(cell_loop.target)

    def generate_for_funccelloutput(self, func_cell_output):
        self.write_line("fargs = [")
        self.indent()
        for arg in func_cell_output.args:
            self.generate_for(arg)
            self.write_line(",")
        self.unindent()
        self.write_line("]")
        self.write("row_height = ")
        self.generate_for(func_cell_output.row_height)
        self.newline()
        self.write("col_width = ")
        self.generate_for(func_cell_output.col_width)
        self.newline()
        style = (
            None
            if func_cell_output.style is None
            else "'{}'".format(func_cell_output.style)
        )
        self.write(
            "cell = cg.FuncCell({}, {}, {}, '{}', row_height, col_width, fargs, ".format(
                func_cell_output.base_cell[0],
                func_cell_output.base_cell[1],
                style,
                func_cell_output.value,
            )
        )
        if func_cell_output.default_value:
            self.generate_for(func_cell_output.default_value)
        else:
            self.write("None")
        self.write_line(")")
        self.write_line(
            "cell_group_{}.add_func_cell(cell)".format(self.cell_group_level)
        )
        if func_cell_output.merge:
            self.generate_for_merge(func_cell_output)

    def generate_for_funcarg(self, func_arg):
        cells = ", ".join(
            "({}, {})".format(cell[0], cell[1]) for cell in func_arg.cells
        )
        direction = "None" if func_arg.direction is None else str(func_arg.direction)
        self.write(
            "cg.FuncArg({}, {}, [{}], {})".format(
                func_arg.start_index, func_arg.end_index, cells, direction
            )
        )

    def generate_for_merge(self, cell_output):
        self.write(
            "cell_group_{}.add_merge({}, {}, ".format(
                self.cell_group_level,
                cell_output.base_cell[0],
                cell_output.base_cell[1],
            )
        )
        if cell_output.merge.rows:
            self.generate_for(cell_output.merge.rows)
        else:
            self.write("1")
        self.write(", ")
        if cell_output.merge.cols:
            self.generate_for(cell_output.merge.cols)
        else:
            self.write("1")
        self.write_line(")")

    def generate_for_getattr(self, get_attr):
        self.write("env.get_attr(")
        self.generate_for(get_attr.obj)
        self.write(', "{}")'.format(get_attr.attr_name))

    def generate_for_getitem(self, get_item):
        self.write("env.get_item(")
        self.generate_for(get_item.obj)
        self.write(", ")
        self.generate_for(get_item.key)
        self.write(")")

    def generate_for_celloutput(self, cell_output):
        style = None if cell_output.style is None else "'{}'".format(cell_output.style)
        self.write("row_height = ")
        self.generate_for(cell_output.row_height)
        self.newline()
        self.write("col_width = ")
        self.generate_for(cell_output.col_width)
        self.newline()
        self.write(
            "cell = cg.Cell({}, {}, {}, ".format(
                cell_output.base_cell[0], cell_output.base_cell[1], style
            )
        )
        if cell_output.value:
            self.generate_for(cell_output.value)
        else:
            self.write("None")
        self.write(", row_height, col_width)")
        self.newline()
        self.write_line("cell_group_{}.add_cell(cell)".format(self.cell_group_level))
        if cell_output.merge:
            self.generate_for_merge(cell_output)

    def generate_for_call(self, call_node):
        self.generate_for(call_node.obj)
        self.write("(")
        for arg in call_node.args:
            self.generate_for(arg)
            self.write(", ")
        self.write(")")

    def generate_for_arg(self, arg):
        self.generate_for(arg.value)

    def generate_for_value(self, value):
        for child in value.body[:-1]:
            self.generate_for(child)
            self.write(" + ")
        self.generate_for(value.body[-1])

    def generate_for_const(self, const):
        self.write(repr(const.value))

    def generate_for_strconst(self, const):
        self.write(const.value)

    def generate_for_filter(self, filter_):
        self.write('env.filters["{}"]'.format(filter_.name))
        self.write("(")
        self.generate_for(filter_.obj)
        self.write(", ")
        for arg in filter_.args:
            self.generate_for(arg)
            self.write(", ")
        self.write(")")

    def generate_for_kwarg(self, kwarg):
        self.write("{}=".format(kwarg.name))
        self.generate_for(kwarg.value)

    def generate_for_tostr(self, to_str):
        self.write("str(")
        self.generate_for(to_str.value)
        self.write(")")

    def generate_for_var(self, var):
        ref = self.symbols.find_ref(var.name)
        if ref is not None:
            res = ref
        else:
            res = 'ctx.resolve("{}")'.format(var.name)
        self.write(res)

    def generate_for_template(self, template_node):
        self.write_line("import io")
        self.write_line("import datetime")
        self.newline()
        self.write_line("import openpyxl")
        self.newline()
        self.write_line(
            "from xlsx_template.runtime import cell_groups as cg, LoopContext"
        )
        self.write_line(
            "from xlsx_template.consts import LoopDirection, FuncArgDirection"
        )
        self.write_line("from xlsx_template import utils")
        self.newline()
        self.newline()
        self.write_line("def root(context, styles, env):")
        self.indent()
        self.write_line("ctx = context")
        self.write_line("wb = openpyxl.Workbook()")
        self.write_line("del wb['Sheet']")
        self.write_line("for style in styles:")
        self.indent()
        self.write_line("wb.add_named_style(style)")
        self.unindent()
        self.newline()
        for child_node in template_node.body:
            self.generate_for(child_node)
            self.newline()
        self.write_line("buf = io.BytesIO()")
        self.write_line("wb.save(buf)")
        self.write_line("return buf.getvalue()")
        self.unindent()
        self.write_line("")
