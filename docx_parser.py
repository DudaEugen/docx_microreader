import xml.etree.ElementTree as ET
from namespaces import namespaces
from typing import Dict, Union, List, Tuple
import re
from propepty_models import PropertyDescription, Property


"""
class Style:
    def __init__(self, style_type: str, style_id: str, is_default: bool,
                 style_dictionary: Dict[str, Union[str, bool, None]], base_style_id: Union[str, None]):
        self.style_type: str = style_type
        self.style_id: str = style_id
        self.is_default: bool = is_default
        self.dictionary: Dict[str, Union[str, bool, None]] = style_dictionary
        self.base_style_id: Union[str, None] = base_style_id
"""


class Parser:
    _all_properties: Dict[str, PropertyDescription] = {}

    def __init__(self, element: ET.Element):
        self._element: ET.Element = element

    def _remove_raw_xml(self):
        del self._element

    @staticmethod
    def get_xml_file(doc, file_name: str) -> ET.Element:
        import xml.dom.minidom
        import zipfile

        raw_xml: Union[str, None] = xml.dom.minidom.parseString(
            zipfile.ZipFile(doc.get_path()).read(rf'word/{file_name}.xml')
        ).toprettyxml()
        return ET.fromstring(raw_xml)

    @staticmethod
    def __check_namespace(tag: str) -> str:
        for key in namespaces:
            if key + ':' in tag:
                return tag.replace(key + ':', '{' + namespaces[key] + '}')

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
                    prop: Union[None, str] = property_element.get(self.__check_namespace(tag_prop))
                    if prop is not None:
                        return Property(prop, pr)
                return Property(None, pr)
            else:
                return Property(property_element.get(self.__check_namespace(tags)), pr)
        return Property(True, pr)

    def _parse_element(self, element: ET.Element):
        from models import Document, Table, Paragraph
        tags = {
            Parser.__check_namespace(Document.Body.tag): Document.Body,
            Parser.__check_namespace(Table.tag): Table,
            Parser.__check_namespace(Table.Row.tag): Table.Row,
            Parser.__check_namespace(Table.Row.Cell.tag): Table.Row.Cell,
            Parser.__check_namespace(Paragraph.tag): Paragraph,
            Parser.__check_namespace(Paragraph.Run.tag): Paragraph.Run,
            Parser.__check_namespace(Paragraph.Run.Text.tag): Paragraph.Run.Text,
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

    """
        @staticmethod
        def _parse_styles(raw_styles: ET.Element, models) -> Dict[str, Style]:
            result: Dict[str, Style] = {}
            styles: List[ET.Element] = raw_styles.findall('./w:style', namespaces)
            for style in styles:
                style_type: str = style.get(Parser.__check_namespace('w:type'))
                style_id: str = style.get(Parser.__check_namespace('w:styleId'))
                is_default: bool = True if style.get(Parser.__check_namespace('w:default')) is not None else False
                for model in models:
                    if style_type == model.type:
                        dictionary: Dict[str, Union[str, bool, None]] = {}
                        for key in model.all_style_properties:
                            st = style.find(model.all_style_properties[key][0], namespaces)
                            if st is not None:
                                if model.all_style_properties[key][1] is not None:
                                    dictionary[key] = st.get(Parser.__check_namespace(model.all_style_properties[key][1]))
                                else:
                                    dictionary[key] = True
                            else:
                                if model.all_style_properties[key][1] is not None:
                                    dictionary[key] = None
                                else:
                                    dictionary[key] = False
                        base: Union[ET.Element, None] = style.find('w:basedOn', namespaces)
                        base_style_id: Union[str, None] = base.get(Parser.__check_namespace('w:val')) \
                            if base is not None else None
                        result[style_id] = Style(style_type, style_id, is_default, dictionary, base_style_id)
            return result
    """


class XMLement(Parser):
    from translators import TranslatorToHTML

    tag: str
    type: str
    translators = {}
    str_format: str = 'html'
    _is_unique: bool = False   # True if parent can containing only one this element
    # all_style_properties: Dict[str, Tuple[str, Union[str, None], bool]] = {}

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
