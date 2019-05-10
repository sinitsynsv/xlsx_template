import os
import shutil
import datetime

import data_generators
from xlsx_template import Template


ROOT_DIR = os.path.dirname(__file__)
EXAMPLE_RES_DIR = os.path.join(ROOT_DIR, "examples")

BASE_TEMPLATE_DIR = os.path.join(ROOT_DIR, "tests", "xlsx")
EXAMPLE_DEF = [
    ("test_variables.xlsx", data_generators.generate_for_test_variables),
    ("test_simple_loop.xlsx", data_generators.generate_for_test_simple_loop),
    ("test_two_nested_loops.xlsx", data_generators.generate_for_test_two_nested_loops),
    (
        "test_two_nested_loops_synt_v2.xlsx",
        data_generators.generate_for_test_two_nested_loops,
    ),
    ("test_loop_context.xlsx", data_generators.generate_for_test_loop_context),
    ("test_merge.xlsx", data_generators.generate_for_test_merge),
    ("test_if.xlsx", data_generators.generate_for_test_if),
    ("loop_with_formulas.xlsx", data_generators.generate_for_loop_with_formulas),
]

EXAMPLE_DEF = [
    (os.path.join(BASE_TEMPLATE_DIR, template), data_generator)
    for (template, data_generator) in EXAMPLE_DEF
]


def main():
    if not os.path.exists(EXAMPLE_RES_DIR):
        os.mkdir(EXAMPLE_RES_DIR)

    for i, (template_name, dg) in enumerate(EXAMPLE_DEF, 1):
        base_template_name = os.path.basename(template_name)
        if base_template_name.startswith("test_"):
            base_template_name = base_template_name[5:]
        base_template_name = "{:0>2d}_{}".format(i, base_template_name)
        file_name, _ = os.path.splitext(base_template_name)
        template = Template(open(template_name, "rb").read(), debug=True)
        source_code_file_name = os.path.join(EXAMPLE_RES_DIR, file_name + ".py")
        with open(source_code_file_name, "w") as f:
            f.write(template.code_source)
        print(
            "{}, start rendering '{}'".format(
                datetime.datetime.now(), base_template_name
            )
        )
        res = template.render(dg())
        print(
            "{}, end rendering '{}'".format(datetime.datetime.now(), base_template_name)
        )
        res_template_name = os.path.join(EXAMPLE_RES_DIR, base_template_name)
        result_file_name = os.path.join(EXAMPLE_RES_DIR, file_name + "_result.xlsx")
        shutil.copy(template_name, res_template_name)
        with open(result_file_name, "wb") as f:
            f.write(res)


if __name__ == "__main__":
    main()
