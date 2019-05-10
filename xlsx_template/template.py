import tempfile
import os

from xlsx_template.parser import Parser
from xlsx_template.code_generator import CodeGenerator
from xlsx_template.environment import Environment
from xlsx_template.runtime.context import Context


class Template:
    def __init__(self, source, env=None, debug=False):
        if env is None:
            env = Environment()
        self.env = env
        if hasattr(source, "read"):
            source = source.read()
        parser = Parser(source)
        self.root_node, self.styles = parser.parse()
        code_source = CodeGenerator().generate(self.root_node)
        if debug:
            fd, self.debug_file_name = tempfile.mkstemp(suffix=".py")
            filename = self.debug_file_name
            with open(fd, "w") as f:
                f.write(code_source)
            self.code_source = code_source
        else:
            filename = "<template>"
        source = compile(code_source, filename, "exec")
        self.namespace = {}
        exec(source, self.namespace)

    def render(self, context_data):
        return self.namespace["root"](
            Context(context_data, self.env), self.styles, self.env
        )
