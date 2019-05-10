class ParseError(Exception):
    def __init__(self, msg, source_hint, orig_exc=None):
        self.msg = msg
        self.source_hint = source_hint
        self.orig_exc = orig_exc


class TemplateRuntimeException(RuntimeError):
    def __init__(self, msg, orig_exception=None):
        super().__init__(msg)
        self.orig_exception = orig_exception
