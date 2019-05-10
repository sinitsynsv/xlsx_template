import os
import io

from openpyxl import load_workbook
import pytest

from xlsx_template.template import Template


@pytest.fixture(scope="session")
def template_dir():
    return os.path.join(os.path.dirname(__file__), "xlsx")


@pytest.fixture(scope="session")
def get_template(template_dir):
    def inner(template_name):
        return open(os.path.join(template_dir, template_name), "rb")

    return inner


@pytest.fixture(scope="session")
def render_template(get_template):
    def inner(template_name, context):
        template = Template(get_template(template_name), debug=True)
        res = template.render(context)
        wb = load_workbook(io.BytesIO(res))
        return wb

    return inner
