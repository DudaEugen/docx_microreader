from typing import Union, List


class PropertyDescription:

    def __init__(self, tag_wrap: Union[str, None], tag: Union[List[str], str],
                 tag_property: Union[List[str], str, None], is_view: bool):
        self.tag_wrap: Union[str, None] = tag_wrap
        self.tag: Union[List[str], str] = tag
        self.tag_property: Union[List[str], str, None] = tag_property
        self.value_type: str = 'str' if self.tag_property is not None else 'bool'
        self.is_view: bool = is_view

    def get_wrapped_tags(self) -> Union[List[str], str]:
        if isinstance(self.tag, list):
            return [self.__wrap_for_tag() + tag for tag in self.tag]
        return self.__wrap_for_tag() + self.tag

    def __wrap_for_tag(self) -> str:
        if self.tag_wrap is not None:
            return self.tag_wrap + '/'
        return ''


class Property:

    def __init__(self, value: Union[str, None, bool], description: PropertyDescription):
        self.description: PropertyDescription = description
        self.value: Union[str, None, bool] = value
