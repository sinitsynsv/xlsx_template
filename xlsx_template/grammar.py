import string

import pyparsing as pp
from pyparsing import pyparsing_common as ppc

from . import nodes, utils


LPAR, RPAR, LBRACK, RBRACK, DOT, EQ, COMMA, COLON = [pp.Suppress(_) for _ in "()[].=,:"]


hex_constant = pp.Regex(r"0[xX][0-9a-fA-F]+").addParseAction(
    lambda t: int(t[0][2:], 16)
)
int_constant = hex_constant | ppc.integer
double_constant = ppc.real


def double_constant_pa(r):
    return nodes.Const(value=float(r.asList()[0]))


double_constant.setParseAction(double_constant_pa)
string_constant = pp.quotedString


def string_constant_pa(r):
    res = r.asList()[0]
    return nodes.StrConst(value=res)


string_constant.setParseAction(string_constant_pa)

boolean_constant = (
    pp.Keyword("true") | pp.Keyword("True") | pp.Keyword("false") | pp.Keyword("False")
)


def boolean_constant_pa(r):
    res = r.asList()[0].lower()
    if res == "true":
        return nodes.Const(value=True)
    else:
        return nodes.Const(value=False)


boolean_constant.setParseAction(boolean_constant_pa)


constant = double_constant | int_constant | string_constant | boolean_constant
constant.setResultsName("constant")


def constant_pa(r):
    return nodes.Const(value=r.asList()[0])


int_constant.setParseAction(constant_pa)


name = pp.Word(pp.alphas + "_", pp.alphanums + "_")
get_attr = DOT + name("attr_name")


def get_attr_pa(r):
    return nodes.GetAttr(attr_name=r.attr_name)


get_attr.setParseAction(get_attr_pa)

expr = pp.Forward()
get_item = LBRACK + expr("key") + RBRACK


def get_item_pa(r):
    return nodes.GetItem(key=r.key)


get_item.setParseAction(get_item_pa)

kwarg = name("name") + EQ + expr("value")


def kwarg_pa(r):
    return nodes.Kwarg(name=r.name, value=r.value)


kwarg.setParseAction(kwarg_pa)

arg = expr("arg")


def arg_pa(r):
    return nodes.Arg(value=r.asList()[0])


arg.setParseAction(arg_pa)

call_args = pp.delimitedList(kwarg | arg)
call = LPAR + pp.Group(pp.Optional(call_args))("args") + RPAR


def call_pa(r):
    return nodes.Call(args=r.asList()[0])


call.setParseAction(call_pa)

var = name("var")


def var_pa(r):
    return nodes.Var(name=r.asList()[0])


var.setParseAction(var_pa)

expr << ((constant | var) + pp.ZeroOrMore(get_attr | get_item | call))


def expr_pa(r):
    obj = None
    r = r.asList()
    for part in r:
        if obj is not None:
            part.obj = obj
        obj = part
    return obj


expr.setParseAction(expr_pa)

filter_expr = name("name") + pp.Optional(
    LPAR + pp.Group(pp.Optional(call_args))("args") + RPAR
)


def filter_expr_pa(r):
    return nodes.Filter(name=r.name, args=r.args)


filter_expr.setParseAction(filter_expr_pa)

expr_with_filter = expr + pp.ZeroOrMore(pp.Suppress("|") + filter_expr)


expr_with_filter.setParseAction(expr_pa)


def parse_expr_with_filter(s):
    return expr_with_filter.parseString(s, True).asList()[0]


cell = pp.Word(pp.alphas.upper())("col") + pp.Word(pp.nums)("row")


def cell_pa(r):
    res = (int(r["row"]), utils.col_str_to_int(r["col"]))
    return res


cell.setParseAction(cell_pa)


cell_range = cell("start_cell") + COLON + cell("end_cell")


def cell_range_pa(r):
    return r.asList()


cell_range.setParseAction(cell_range_pa)


def param_pa(r):
    r = r.asList()
    res = {"name": r[0], "value": r[1:]}
    if len(res["value"]) == 1:
        res["value"] = res["value"][0]
    return res


def make_param(param_name, param_value):
    r = pp.Keyword(param_name) + EQ + param_value
    r.setParseAction(param_pa)
    return r


for_statement = (
    pp.Keyword("for").suppress()
    + name("target")
    + pp.Keyword("in").suppress()
    + expr("items")
)


def for_statement_pa(r):
    return {"target": r["target"], "items": r["items"]}


for_statement.setParseAction(for_statement_pa)

cell_loop = for_statement("stmt") + pp.Group(
    pp.Optional(
        COMMA
        + pp.delimitedList(make_param("name", name) | make_param("last_cell", cell))
    )
)("params")


def cell_loop_pa(r):
    keywords = {param["name"]: param["value"] for param in r["params"].asList()}
    return nodes.CellLoop(
        target=r["stmt"]["target"], items=r["stmt"]["items"], **keywords
    )


cell_loop.setParseAction(cell_loop_pa)


def parse_cell_loop_directive(s):
    return cell_loop.parseString(s, True).asList()[0]


sheet_loop = for_statement("stmt") + pp.Group(
    pp.Optional(COMMA + pp.delimitedList(make_param("name", name)))
)("params")


def sheet_loop_pa(r):
    keywords = {param["name"]: param["value"] for param in r["params"].asList()}
    return nodes.SheetLoop(
        target=r["stmt"]["target"], items=r["stmt"]["items"], **keywords
    )


sheet_loop.setParseAction(sheet_loop_pa)


def parse_sheet_loop_directive(s):
    return sheet_loop.parseString(s, True).asList()[0]


merge = pp.delimitedList(make_param("rows", expr) | make_param("cols", expr))


def merge_pa(r):
    return nodes.Merge(**{param["name"]: param["value"] for param in r.asList()})


merge.setParseAction(merge_pa)


def parse_merge(s):
    return merge.parseString(s, True).asList()[0]


group = pp.delimitedList(make_param("last_cell", cell))


def group_pa(r):
    return nodes.CellGroup(**{param["name"]: param["value"] for param in r.asList()})


group.setParseAction(group_pa)


def parse_group(s):
    return group.parseString(s, True).asList()[0]


remove = pp.Optional(make_param("last_cell", cell))


def remove_pa(r):
    params = {param["name"]: param["value"] for param in r.asList()}
    return nodes.Remove(**params)


remove.setParseAction(remove_pa)


def parse_remove(s):
    return remove.parseString(s, True).asList()[0]


excel_func_arg = cell_range | cell


def excel_func_arg_pa(r):
    return r.asList()


excel_func_arg.setParseAction(excel_func_arg_pa)


excel_func_args = pp.ZeroOrMore(excel_func_arg)


def parse_func_args(s):
    return [
        ((start_index, end_index), parse_res)
        for (parse_res, start_index, end_index) in excel_func_args.scanString(s)
    ]


if_d = pp.delimitedList(
    make_param("condition", expr)
    | make_param("last_cell", cell)
    | make_param("else", excel_func_arg)
)


def if_pa(r):
    params = {param["name"]: param["value"] for param in r.asList()}
    params["else_block"] = params.pop("else", None)
    return nodes.If(**params)


if_d.setParseAction(if_pa)


def parse_if(s):
    return if_d.parseString(s, True).asList()[0]


def parse_col_width(s):
    return nodes.ColWidth(value=parse_expr_with_filter(s))


def parse_row_height(s):
    return nodes.RowHeight(value=parse_expr_with_filter(s))
