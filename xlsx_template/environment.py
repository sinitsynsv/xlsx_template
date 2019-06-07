from .exceptions import TemplateRuntimeException
from . import filters
from .parser import Parser


class StrictResolveStrategy:
    def resolve(self, obj, name, found):
        if not found:
            raise Unresolved("Can not find variable '{}'".format(name))
        return obj


class StrictGetAttrStrategy:
    def get_attr(self, obj, attr_name):
        try:
            return getattr(obj, attr_name)
        except AttributeError as e:
            raise TemplateRuntimeException(
                "Can not get attrib {} from obj {}".format(attr_name, obj), e
            )


class SoftGetAttrStrategy:
    def get_attr(self, obj, attr_name):
        exc = None
        try:
            return getattr(obj, attr_name)
        except AttributeError as e:
            exc = e
        try:
            return obj[attr_name]
        except Exception as e:
            exc = exc or e
            raise TemplateRuntimeException(
                "Can not get attrib {} from obj {}".format(attr_name, obj), exc
            )


class StrictGetItemStrategy:
    def get_item(self, obj, key):
        try:
            return obj[key]
        except Exception as e:
            raise TemplateRuntimeException(
                "Can not get item {} from obj {}".format(key, obj), e
            )


class Environment:

    filters = {"default_if_none": filters.default_if_none, "yes_no": filters.yes_no}

    def __init__(
        self, resolve_strategy=None, get_attr_strategy=None, get_item_strategy=None
    ):
        if resolve_strategy is None:
            resolve_strategy = StrictResolveStrategy()

        if get_attr_strategy is None:
            get_attr_strategy = SoftGetAttrStrategy()

        if get_item_strategy is None:
            get_item_strategy = StrictGetItemStrategy()

        self.resolve_strategy = resolve_strategy
        self.get_attr_strategy = get_attr_strategy
        self.get_item_strategy = get_item_strategy

    def resolve(self, obj, name, found):
        return self.resolve_strategy.resolve(obj, name, found)

    def get_attr(self, obj, attr_name):
        return self.get_attr_strategy.get_attr(obj, attr_name)

    def get_item(self, obj, key):
        return self.get_item_strategy.get_item(obj, key)
