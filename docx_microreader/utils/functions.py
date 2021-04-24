from typing import Callable


def execute_if_not_none(argument, function: Callable, default=None):
    """
    :return: function(argument) if argument is not None else default
    """
    return function(argument) if argument is not None else default
