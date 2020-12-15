import decimal


def generate_for_test_variables():
    class Obj:
        def __getattr__(self, attr_name):
            return "getattr {}".format(attr_name)

        def __getitem__(self, key):
            return "getitem {}".format(repr(key))

    def complex_call(*args, **kwargs):
        res = ", ".join(str(arg) for arg in args)
        if kwargs:
            if res:
                res += ", "
            res += ", ".join(
                "{}={}".format(key, value) for key, value in kwargs.items()
            )
        return res

    return {
        "sheet_name": "My Sheet",
        "int_var": 10,
        "str_var": "String",
        "float_var": 10.0,
        "decimal_var": decimal.Decimal(10.1),
        "simple_call": lambda: "simple_call",
        "complex_call": complex_call,
        "arg1": "1",
        "arg2": True,
        "obj": Obj(),
        "none_variable": None,
    }


def generate_for_test_simple_loop():
    items = list("item{}".format(i) for i in range(1, 6))
    return {"items": items}


def generate_for_test_two_nested_loops():
    table = [[row * 10 + col for col in range(10)] for row in range(10)]
    return {"table": table, "row_count": 10, "column_count": 10}


def generate_for_test_loop_context():
    items = ["item{}".format(i) for i in range(5)]
    return {"items": items}


def generate_for_loop_with_formulas():
    employees = [
        {
            "name": "Employee # {}".format(index + 1),
            "salary": 1000 + ((index % 5) * 100),
            "bonus": decimal.Decimal("0.1") * (5 - (index % 5)),
        }
        for index in range(100)
    ]
    return {"employees": employees}


def generate_for_test_merge():
    return {"items": [i for i in range(3)]}


def generate_for_test_if():
    return {"rows": [{"value": value, "is_red": value % 2 == 0} for value in range(10)]}


def generate_for_test_column_width():
    return {"cells": [{"dim_value": 10 * index} for index in range(1, 6)]}


def generate_for_sheet_loop():
    return {
        "sheets": [
            {
                "name": f"Sheet # {index}",
                "sheet_var": f"Sheet var # {index}",
                "numbers": list(range(index)),
            }
            for index in range(1, 11)
        ]
    }