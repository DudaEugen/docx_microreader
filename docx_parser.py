import xml.etree.ElementTree as ET
from namespaces import namespaces
from typing import Dict, Union, List, Tuple, Callable
import re
from properties import PropertyDescription, Property
from constants import get_properties_dict


class Parser:
    _all_properties: Dict[str, PropertyDescription] = {}

    def __init__(self, element: ET.Element):
        self._element: ET.Element = element

    def _remove_raw_xml(self):
        del self._element

    @staticmethod
    def _check_namespace(tag: str) -> str:
        if ':' in tag:
            key = re.split(':', tag)[0]
            return tag.replace(key + ':', '{' + namespaces[key] + '}')
        raise ValueError(rf"tag '{tag}' don't have namespace")

    def __find_property_element(self, description: PropertyDescription) -> Union[ET.Element, None]:
        """
        find element by propertyDescription
        """
        if isinstance(description.get_wrapped_tags(), list):
            for tag in description.get_wrapped_tags():
                result: Union[ET.Element, None] = self._element.find(tag, namespaces)
                if result is not None:
                    return result
            return None
        return self._element.find(description.get_wrapped_tags(), namespaces)

    def __find_property(self, property_element: ET.Element, pr: PropertyDescription,
                        tags: Union[List[str], str, None]) -> Property:
        """
        find property in element
        """
        if pr.value_type == 'str':
            if isinstance(tags, list):
                for tag_prop in tags:
                    prop: Union[None, str] = property_element.get(self._check_namespace(tag_prop))
                    if prop is not None:
                        return Property(prop, pr)
                return Property(None, pr)
            else:
                return Property(property_element.get(self._check_namespace(tags)), pr)
        return Property(True, pr)

    def _parse_element(self, element: ET.Element):
        from models import Document, Table, Paragraph

        tags: Dict[str, Callable] = {
            Parser._check_namespace(Document.Body.tag): Document.Body,
            Parser._check_namespace(Table.tag): Table,
            Parser._check_namespace(Table.Row.tag): Table.Row,
            Parser._check_namespace(Table.Row.Cell.tag): Table.Row.Cell,
            Parser._check_namespace(Paragraph.tag): Paragraph,
            Parser._check_namespace(Paragraph.Run.tag): Paragraph.Run,
            Parser._check_namespace(Paragraph.Run.Text.tag): Paragraph.Run.Text,
        }

        return tags[element.tag](element, self) if element.tag in tags else None

    def _get_elements(self, class_of_element):
        if class_of_element._is_unique:
            return class_of_element(self._element.find(class_of_element.tag, namespaces), self)
        else:
            result: list = []
            for el in self._element.findall('./' + class_of_element.tag, namespaces):
                elem = self._parse_element(el)
                if elem is not None:
                    result.append(elem)
            return result

    def _get_all_elements(self):
        result: list = []
        for el in self._element.findall('./'):
            elem = self._parse_element(el)
            if elem is not None:
                result.append(elem)
        return result

    def _parse_properties(self) -> Dict[str, Property]:
        result: Dict[str, Property] = {}
        for key in self._all_properties:
            pr: PropertyDescription = self._all_properties[key]
            property_element: Union[ET.Element, None] = self.__find_property_element(pr)
            if property_element is not None:
                result[key] = self.__find_property(property_element, pr, self._all_properties[key].tag_property)
            else:
                if pr.value_type == 'str':
                    result[key] = Property(None, pr)
                else:
                    result[key] = Property(False, pr)
        return result

    @staticmethod
    def _parse_color_str(color: str) -> str:
        match = re.match('[0-9A-F]{6}', color)
        if match is not None:
            if color == match.group(0):
                return '#' + color
        return color


class DocumentParser(Parser):

    def __init__(self, element: ET.Element):
        super(DocumentParser, self).__init__(element)
        self._init()
        self._remove_raw_xml()

    def _init(self):
        pass

    @staticmethod
    def get_xml_file(path: str, file_name: str) -> ET.Element:
        """
        :param path: path to document (include name of document)
        :param file_name: name of xml file in document in directory word
        :return: ElementTree of xml file
        """
        import xml.dom.minidom
        import zipfile

        raw_xml: Union[str, None] = xml.dom.minidom.parseString(
            zipfile.ZipFile(path).read(rf'word/{file_name}.xml')
        ).toprettyxml()
        return ET.fromstring(raw_xml)

    @staticmethod
    def __parse_style(element: ET.Element, parent):
        from styles import ParagraphStyle

        types: Dict[str, Callable] = {
            ParagraphStyle.type: ParagraphStyle,
        }

        t: str = element.get(Parser._check_namespace('w:type'))
        return types[t](element, parent) if t in types else None

    def _get_styles(self, styles_file: ET.Element):
        from styles import Style

        result: list = []
        for el in styles_file.findall('./' + Style.tag, namespaces):
            elem = DocumentParser.__parse_style(el, self)
            if elem is not None:
                result.append(elem)
        return result


class XMLement(Parser):
    from translators import TranslatorToHTML

    tag: str
    type: str = ''
    translators = {}
    str_format: str = 'html'
    _is_unique: bool = False   # True if parent can containing only one this element
    # all_style_properties: Dict[str, Tuple[str, Union[str, None], bool]] = {}

    # function that return Dict of property descriptions
    _property_descriptions_getter: Callable = get_properties_dict

    # first element of Tuple is correct variant of property value; second element is variants of this value
    # _properties_validate method set correct variant if find value equal of one of variant
    # if value of property not equal one of variants or correct variant set None
    _properties_unificators: Dict[str, List[Tuple[str, List[str]]]] = {}

    def _init(self):
        pass

    def __init__(self, element: ET.Element, parent):
        self.parent: Union[XMLement, None] = parent
        super(XMLement, self).__init__(element)
        self._init()
        self._all_properties = XMLement._property_descriptions_getter(self)
        self._properties: Dict[str, Property] = self._parse_properties()
        self._properties_unificate()
        self._remove_raw_xml()

    def __str__(self):
        return self.translators[self.str_format].translate(self)

    def _properties_unificate(self):
        """
        first element of _properties_validators[key] is correct variant of property value;
        second element is variants of this value
        set correct variant if find value equal of one of variant
        if value of property not equal one of variants or correct variant set None
        """
        for key in self._properties_unificators:
            if key in self._properties:
                is_finding_value: bool = False
                for correct_value, variants in self._properties_unificators[key]:
                    if self._properties[key].value == correct_value:
                        is_finding_value = True
                        break
                    else:
                        for variant in variants:
                            if self._properties[key].value == variant:
                                self._properties[key].value = correct_value
                                is_finding_value = True
                                break
                if not is_finding_value:
                    self._properties[key].value = None
            else:
                raise KeyError(f'not found key "{key}" from _properties_validators in _all_properties')

    def get_inner_text(self) -> Union[str, None]:
        return None

    def get_properties(self) -> List[Property]:
        return [self._properties[k] for k in self._properties]

    def get_property(self, property_name: str) -> Property:
        return self._properties[property_name]

    def get_property_value(self, property_name: str) -> Union[str, bool, None]:
        return self._properties[property_name].value

    def set_property_value(self, property_name: str, value: Union[str, bool, None]):
        self._properties[property_name].value = value

    def set_property_view(self, property_name: str, is_view: bool):
        self._all_properties[property_name].is_view = is_view

    def _is_have_viewing_property(self, property_name: str) -> bool:
        """
        return True if property is not None and this property is viewing
        """
        return self._properties[property_name].is_view_and_not_none()


class XMLcontainer(XMLement):
    """
    this objects can containing Tables, Paragraphs, Images, Lists
    """

    def _init(self):
        from models import Paragraph, Table

        elements: list = self._get_all_elements()
        for element in elements:
            if isinstance(element, Paragraph):
                self.paragraphs.append(element)
                self.elements.append(element)
            elif isinstance(element, Table):
                self.tables.append(element)
                self.elements.append(element)

    def __init__(self, element: ET.Element, parent):
        from models import Paragraph, Table

        self.tables: List[Table] = []
        self.paragraphs: List[Paragraph] = []
        self.elements: List[Union[Table, Paragraph]] = []
        super(XMLcontainer, self).__init__(element, parent)

    def get_inner_text(self) -> Union[str, None]:
        result: str = ''
        for element in self.elements:
            result += str(element)
        return result
