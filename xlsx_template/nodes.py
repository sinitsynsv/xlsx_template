class NodeMetaclass(type):
    def __new__(cls, name, bases, attrs):
        if bases:
            base = bases[0]
            attributes = base.attributes.copy()
        else:
            attributes = {}
        attributes.update(
            {
                key: value
                for key, value in attrs.items()
                if not key.startswith("_") and key != "attributes"
            }
        )
        attrs["attributes"] = attributes
        return super().__new__(cls, name, bases, attrs)


class Node(metaclass=NodeMetaclass):
    attributes = []

    def __init__(self, **kwargs):
        for attr_name, value in kwargs.items():
            if attr_name not in self.attributes:
                raise RuntimeError("Invalid attribute {}".format(attr_name))
            setattr(self, attr_name, value)


class Const(Node):
    value = None


class StrConst(Const):
    pass


class Var(Node):
    name = None


class Value(Node):
    body = None


class ToStr(Node):
    value = None


class BaseObjNode(Node):
    obj = None


class GetAttr(BaseObjNode):
    attr_name = None


class GetItem(BaseObjNode):
    key = None


class Arg(Node):
    value = None


class Kwarg(Node):
    name = None
    value = None


class Call(BaseObjNode):
    args = None


class Filter(BaseObjNode):
    name = None
    args = None


class Template(Node):
    body = None


class CellOutput(Node):
    base_cell = None
    value = None
    style = None
    row_height = None
    col_width = None
    merge = None

    def adjust(self, row, col):
        self.base_cell = (self.base_cell[0] + row, self.base_cell[1] + col)


class CellGroup(Node):
    base_cell = None
    last_cell = None
    body = None

    def adjust(self, row, col):
        self.base_cell = (self.base_cell[0] + row, self.base_cell[1] + col)
        self.last_cell = (self.last_cell[0] + row, self.last_cell[1] + col)

    @property
    def height(self):
        return self.last_cell[0] - self.base_cell[0] + 1

    @property
    def width(self):
        return self.last_cell[1] - self.base_cell[1] + 1

    def get_cell_range(self):
        for row in range(self.base_cell[0], self.last_cell[0] + 1):
            for col in range(self.base_cell[1], self.last_cell[1] + 1):
                yield (row, col)


class Remove(CellGroup):
    pass


class If(CellGroup):
    condition = None
    else_block = None


class Sheet(CellGroup):
    name = None

    def __init__(self, name, max_row, max_col, body):
        super().__init__(
            base_cell=(0, 0), last_cell=(max_row - 1, max_col - 1), body=body
        )
        self.name = name


class CellLoop(CellGroup):
    target = None
    items = None
    name = None
    body = None
    direction = None


class SheetLoop(CellGroup):
    target = None
    items = None
    name = None
    sheet = None


class FuncArg(Node):
    start_index = None
    end_index = None
    cells = None
    direction = None


class Merge(Node):
    rows = None
    cols = None


class FuncArgDirection(Node):
    direction = None


class FuncCellOutput(CellOutput):
    default_value = None
    args = None

    def adjust(self, row, col):
        for arg in self.args:
            arg.cells = [
                (cell[0] - self.base_cell[0], cell[1] - self.base_cell[1])
                for cell in arg.cells
            ]
        super().adjust(row, col)


class ColWidth(Node):
    value = None


class RowHeight(Node):
    value = None
