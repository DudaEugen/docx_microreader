from typing import Union, List


class PropertyDescription:

    def __init__(self, tag_wrap: Union[List[str], str, None], tag: Union[List[str], str],
                 tag_property: Union[List[str], str, None]):
        self.tag_wrap: Union[List[str], str, None] = tag_wrap
        self.tag: Union[List[str], str] = tag
        self.tag_property: Union[List[str], str, None] = tag_property
        self.value_type: str = 'str' if self.tag_property is not None else 'bool'

    def get_wrapped_tags(self) -> Union[List[str], str]:
        tag_wrap: Union[List[str], str] = self.__wrap_for_tag()
        if isinstance(self.tag, list) and isinstance(tag_wrap, list):
            result: List[str] = []
            for wrap in tag_wrap:
                result.extend([(wrap + tag) for tag in self.tag])
            return result
        if isinstance(self.tag, list):
            return [(tag_wrap + tag) for tag in self.tag]
        if isinstance(tag_wrap, list):
            return [(wrap + self.tag) for wrap in tag_wrap]
        return tag_wrap + self.tag

    def __wrap_for_tag(self) -> Union[List[str], str]:
        if self.tag_wrap is not None:
            if isinstance(self.tag_wrap, list):
                return [(tag_wrap + '/') for tag_wrap in self.tag_wrap]
            return self.tag_wrap + '/'
        return ''


class Property:

    def __init__(self, value: Union[str, None, bool]):
        self.value: Union[str, None, bool] = value
