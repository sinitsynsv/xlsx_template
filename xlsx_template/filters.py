def default_if_none(value, default_value):
    return default_value if value is None else value


def yes_no(value, yes_value, no_value):
    if value:
        return yes_value
    else:
        return no_value
