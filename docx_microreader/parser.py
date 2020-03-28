from .support_classes import ContentInf
import re
from typing import List, Union, Tuple


class XMLement:
    _output_format: str = 'html'
    _tag_name: str
    _is_can_contain_same_elements: bool = False

    def __init__(self, element: ContentInf):
        self._raw_xml: Union[str, None] = element.content
        self._begin: int = element.begin
        self._end: int = element.end

    def __find_property_of_element_that_can_contain_same_elements(self) -> Union[str, None]:
        inner_content = re.sub(rf'<{self._tag_name}([^P\n>%]+)?>', '', self._raw_xml)
        inner_content = re.sub(rf'</{self._tag_name}>', '', inner_content)
        property = re.search(rf'<{self._tag_name}Pr>([^%]+)?</{self._tag_name}Pr>', inner_content)
        inner_same_element = re.search(rf'<{self._tag_name}( [^\n>%]+)?>[^%]+?</{self._tag_name}>', inner_content)
        return property if inner_same_element is None or property.span()[0] < inner_same_element.span()[0] else None

    def __parse_all_elements(self, regular: str) -> List[ContentInf]:
        elements: List[ContentInf] = []
        for o in re.finditer(regular, self._raw_xml):
            elements.append(ContentInf(o.group(0), o.span()))
        return elements

    def __parse_one_element(self, regular: str) -> ContentInf:
        element = re.search(regular, self._raw_xml)
        return ContentInf(element.group(0), element.span())

    def __parse_external_elements(self, tag: str, get_one_element=False) -> Union[List[ContentInf], ContentInf]:
        elements: List[ContentInf] = []
        tags: List[Tuple[ContentInf, bool]] = []
        for o in re.finditer(rf'<{tag}( [^\n>%]+)?>', self._raw_xml):
            tags.append((ContentInf(o.group(0), o.span()), True))
        for o in re.finditer(rf'</{tag}>', self._raw_xml):
            tags.append((ContentInf(o.group(0), o.span()), False))

        tags.sort(key=lambda x: x[0].begin)
        i: int = 0
        while i < len(tags) - 1:
            begin_number = 1
            for j in range(i + 1, len(tags)):
                if tags[j][1]:
                    begin_number += 1
                else:
                    begin_number -= 1
                    if begin_number == 0:
                        element: ContentInf = ContentInf(
                                self._raw_xml[tags[i][0].begin:tags[j][0].end],
                                (tags[i][0].begin, tags[j][0].end)
                        )
                        if not get_one_element:
                            elements.append(element)
                            i = j + 1
                            break
                        else:
                            return element
        return elements

    def _parse_elements(self, element_class, get_one_element=False) -> Union[List[ContentInf], ContentInf]:
        tag: str = element_class._tag_name

        if not element_class._is_can_contain_same_elements:
            regular: str = rf'<{tag}( [^\n>%]+)?>[^%]+?</{tag}>'
            return self.__parse_all_elements(regular) if not get_one_element else self.__parse_one_element(regular)
        else:
            return self.__parse_external_elements(tag, get_one_element)

    def _inner_content(self) -> str:
        inner_content = re.sub(rf'<{self._tag_name}([^\n>%]+)?>', '', self._raw_xml)
        inner_content = re.sub(rf'</{self._tag_name}>', '', inner_content)
        return inner_content

    def _parse_properties(self, tag_name: str) -> Union[str, None]:
        if not self._is_can_contain_same_elements:
            properties = re.search(rf'<{self._tag_name}Pr>([^%]+)?</{self._tag_name}Pr>', self._raw_xml)
        else:
            properties = self.__find_property_of_element_that_can_contain_same_elements()

        if properties:
            prop = re.search(rf'<{tag_name} w:val="([^"%]+)"/>', properties.group(0))
            if prop:
                p = prop.group(0)
                begin = p.find('"') + 1
                end = p.find('"', begin)
                return p[begin: end]

    def _parse_named_value_of_properties(
            self, property_name: str, value: str, wrapper: Union[str, None] = None) -> Union[str, None]:
        if not self._is_can_contain_same_elements:
            properties = re.search(rf'<{self._tag_name}Pr>([^%]+)?</{self._tag_name}Pr>', self._raw_xml)
        else:
            properties = self.__find_property_of_element_that_can_contain_same_elements()

        if properties is not None:
            if wrapper is not None:
                wrap = re.search(rf'<{wrapper}>([^%]+)?</wrapper>', properties.group(0))
                prop = re.search(rf'<{property_name}([^/%]+)?/>', wrap.group(0))
            else:
                prop = re.search(rf'<{property_name}([^/%]+)?/>', properties.group(0))
            if prop is not None:
                v = re.search(rf'{value}="([^"%]+)"', prop.group(0))
                if v:
                    result = v.group(0)
                    begin = result.find('"') + 1
                    end = result.find('"', begin)
                    return result[begin: end]

    def _have_properties(self, tag_name: str) -> bool:
        properties = re.search(rf'<{self._tag_name}Pr([^\n>%]+)?>([^%]+)?</{self._tag_name}Pr>', self._raw_xml).group(0)
        return True if (properties.find(rf'<{tag_name}/>') != -1) else False

    def _remove_raw_xml(self):
        del self._raw_xml
