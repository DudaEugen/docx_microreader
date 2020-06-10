from typing import Union, Tuple

PROPERTY_TYPES: Tuple[str, str] = ('str', 'bool')


class PropertyDescription:

    def __init__(self, tag_wrap: str, tag: str, tag_property: Union[str, None], is_view: bool):
        self.tag_wrap: str = tag_wrap
        self.tag: str = tag
        self.tag_property: Union[str, None] = tag_property
        self.value_type: str = 'str' if self.tag_property is not None else 'bool'
        self.is_view: bool = is_view

    def get_wrapped_tag(self) -> str:
        return self.tag_wrap + '/' + self.tag

    def get_xml_property_name(self) -> str:
        return self.get_wrapped_tag() + (('/' + self.tag_property) if self.tag_property is not None else '')


class Property:

    def __init__(self, value: Union[str, None, bool], description: PropertyDescription):
        self.description: PropertyDescription = description
        self.value: Union[str, None, bool] = value

    def is_view_and_not_none(self) -> bool:
        return self.description.is_view and (self.value is not None and self.value is not False)
