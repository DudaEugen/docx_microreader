from enum import Enum
from ..properties import PropertyDescription
from typing import Dict
from functools import cache


class ElementPropertyEnum(Enum):
    def __init__(self, key: str, property_description: PropertyDescription, is_style_property: bool = False):
        self.key: str = key
        self.description: PropertyDescription = property_description
        self.is_style_property: bool = is_style_property

    @classmethod
    def get_properties_description_dict(cls, is_only_style_properties: bool = False) -> Dict[str, PropertyDescription]:
        """
        create keys for dict of property descriptions (those keys use also for corresponding properties)
        :param is_only_style_properties: don't using instances of cls for whom is_style_property == False if True
        :return: Dict[el.key, el.description] where el is instance of cls
        """
        result: Dict[str, PropertyDescription] = {}
        for i in cls:
            if not is_only_style_properties or i.is_style_property:
                if i.key in result:
                    raise KeyError(f'duplicate key {i.key} in {cls} Enum')
                result[i.key] = i.description
        return result


@cache
def _construct_dict_of_str_from_enum(enum_cls) -> dict:
    """
    Convert enum to dict
    :param enum_cls: Enum class
    :return: Dict[el.value, el] where el is instance of enum
    """
    result = {}
    for el in enum_cls:
        if isinstance(el.value, str):
            result[el.value] = el
        elif isinstance(el.value, tuple):
            result[el.value[0]] = el
        else:
            raise TypeError(f"can't convert {enum_cls} Enum to dict")
    return result


def convert_to_enum_element(s, enum_cls):
    """
    convert s to value of Enum if it possible. Else pass Error
    :param enum_cls: Enum class
    :param s: str or enum_cls instance
    :return: instance of enum_cls
    """
    if isinstance(s, str):
        return _construct_dict_of_str_from_enum(enum_cls)[s]
    elif isinstance(s, enum_cls):
        return s
    raise TypeError(f"can't convert {s} to {enum_cls} enum")
