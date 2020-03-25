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

    def _parse_element(self, element_class) -> ContentInf:
        tag: str = element_class._tag_name
        element = re.search(rf'<{tag}( [^\n>%]+)?>([^%]+)?</{tag}>', self._raw_xml)
        return ContentInf(element.group(0), element.span())

    def _parse_elements(self, element_class) -> List[ContentInf]:
        tag: str = element_class._tag_name
        elements: List[ContentInf] = []
        if not element_class._is_can_contain_same_elements:
            for o in re.finditer(rf'<{tag}( [^\n>%]+)?>[^%]+?</{tag}>', self._raw_xml):
                elements.append(ContentInf(o.group(0), o.span()))
        else:
            tags: List[Tuple[ContentInf, bool]] = []
            for o in re.finditer(rf'<{tag}( [^\n>%]+)?>', self._raw_xml):
                tags.append((ContentInf(o.group(0), o.span()), True))
            for o in re.finditer(rf'</{tag}>', self._raw_xml):
                tags.append((ContentInf(o.group(0), o.span()), False))
            tags.sort(key=lambda x: x[0].begin)
            i: int = 0
            while i < len(tags) - 1:
                begin_number = 1
                for j in range(i+1, len(tags)):
                    if tags[j][1]:
                        begin_number += 1
                    else:
                        begin_number -= 1
                        if begin_number == 0:
                            elements.append(ContentInf(
                                    self._raw_xml[tags[i][0].begin:tags[j][0].end],
                                    (tags[i][0].begin, tags[j][0].end)
                                )
                            )
                            i = j + 1
                            break
        return elements

    def _inner_content(self) -> str:
        inner_content = re.sub(rf'<{self._tag_name}([^\n>%]+)?>', '', self._raw_xml)
        inner_content = re.sub(rf'</{self._tag_name}>', '', inner_content)
        return inner_content

    def _parse_properties(self, tag_name: str) -> Union[str, None]:
        properties = re.search(rf'<{self._tag_name}Pr>([^%]+)?</{self._tag_name}Pr>', self._raw_xml)
        if properties:
            prop = re.search(rf'<{tag_name} w:val="([^"%]+)"/>', properties.group(0))
            if prop:
                p = prop.group(0)
                begin = p.find('"') + 1
                end = p.find('"', begin)
                return p[begin: end]

    def _parse_named_value_of_properties(self, property_name: str, value: str) -> Union[str, None]:
        properties = re.search(rf'<{self._tag_name}Pr>([^%]+)?<{property_name} ([^\n>%]+)?/>([^%]+)?</{self._tag_name}Pr>',
                               self._raw_xml)
        if properties:
            prop = re.search(rf'<{property_name}([^/%]+)?/>', properties.group(0))
            if prop:
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
