import xml.etree.ElementTree as ET
from constants.namespaces import namespaces
from typing import Union, List, Callable, Dict, Tuple
import re
from properties import Property, PropertyDescription
from constants import properties_consts as p_consts
from constants import keys_consts as k_consts


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
        return tag

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
                        return Property(prop)
                return Property(None)
            else:
                return Property(property_element.get(self._check_namespace(tags)))
        else:
            value: Union[str, None] = property_element.get(self._check_namespace(k_consts.BoolPropertyValue))
            if value is not None and value == '0':
                return Property(False)
            else:
                return Property(True)

    def _parse_element(self, element: ET.Element):
        from models import Document, Table, Paragraph, Drawing

        tags: Dict[str, Callable] = {
            Parser._check_namespace(Document.Body.tag): Document.Body,
            Parser._check_namespace(Table.tag): Table,
            Parser._check_namespace(Table.Row.tag): Table.Row,
            Parser._check_namespace(Table.Row.Cell.tag): Table.Row.Cell,
            Parser._check_namespace(Paragraph.tag): Paragraph,
            Parser._check_namespace(Paragraph.Run.tag): Paragraph.Run,
            Parser._check_namespace(Paragraph.Run.Text.tag): Paragraph.Run.Text,
            Parser._check_namespace(Drawing.tag): Drawing,
            Parser._check_namespace(Drawing.Image.tag): Drawing.Image,
        }

        return tags[element.tag](element, self) if element.tag in tags else None

    def _get_elements(self, class_of_element):
        if class_of_element._is_unique:
            element: Union[ET.Element, None] = self._element.find(class_of_element.tag, namespaces)
            return class_of_element(element, self) if element is not None else None
        else:
            result: list = []
            for el in self._element.findall('./' + class_of_element.tag, namespaces):
                elem = class_of_element(el, self)
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
        for key, pr in self._all_properties.items():
            if pr.tag is not None:
                property_element: Union[ET.Element, None] = self.__find_property_element(pr)
                if property_element is not None:
                    result[key] = self.__find_property(property_element, pr, self._all_properties[key].tag_property)
                else:
                    result[key] = Property(None)
            else:
                result[key] = Property(self._element.get(self._check_namespace(pr.tag_property)))
        return result

    @staticmethod
    def _parse_color_str(color: str) -> str:
        match = re.match('[0-9A-F]{6}', color)
        if match is not None:
            if color == match.group(0):
                return '#' + color
        return color


class DocumentParser(Parser):

    def __init__(self, path: str, path_for_images: Union[None, str] = None):
        from os.path import abspath

        self._path: str = abspath(path).replace('\\', '/')
        super(DocumentParser, self).__init__(self.get_xml_file('document.xml'))
        self._styles: dict = {}
        self._parse_styles()
        self._images_dir: str = abspath(path_for_images).replace('\\', '/') + '/' if path_for_images is not None else \
            self._get_images_directory(False)
        self._images: Dict[str, str] = self._parse_images_relationships()
        self.__images_extraction()

    def get_xml_file(self, file_name: str, is_return_element_tree: bool = True) -> Union[ET.Element, str, None]:
        """
        :param file_name: name of xml file in document in directory word
        :param is_return_element_tree: method return ElementTree if True else return str
        :return: ElementTree of xml file or str
        """
        import xml.dom.minidom
        import zipfile

        raw_xml: Union[str, None] = xml.dom.minidom.parseString(
            zipfile.ZipFile(self._path).read(rf'word/{file_name}')
        ).toprettyxml()
        if is_return_element_tree:
            return ET.fromstring(raw_xml)
        return raw_xml

    def __parse_style(self, element: ET.Element):
        from styles import ParagraphStyle, CharacterStyle, TableStyle, NumberingStyle

        types: Dict[str, Callable] = {
            ParagraphStyle.type: ParagraphStyle,
            CharacterStyle.type: CharacterStyle,
            TableStyle.type: TableStyle,
            NumberingStyle.type: NumberingStyle,
        }

        parameters: Tuple[str, str, bool, bool] = (
            element.get(Parser._check_namespace(p_consts.Style_parameters[p_consts.StyleParam_type])),
            element.get(Parser._check_namespace(p_consts.Style_parameters[p_consts.StyleParam_id])),
            False if element.get(Parser._check_namespace(
                p_consts.Style_parameters[p_consts.StyleParam_is_default])
            ) is None else True,
            False if element.get(Parser._check_namespace(
                p_consts.Style_parameters[p_consts.StyleParam_is_custom])
            ) is None else True,
        )
        return types[parameters[0]](element, self, parameters[1],
                                    parameters[2], parameters[3]) if parameters[0] in types else None

    def _parse_styles(self):
        from styles import Style

        for el in self.get_xml_file('styles.xml').findall('./' + Style.tag, namespaces):
            elem = self.__parse_style(el)
            if elem is not None:
                self._styles[elem.id] = elem

    def _parse_images_relationships(self) -> Dict[str, str]:
        result: Dict[str, str] = {}
        for relationship in re.findall(r'<Relationship ([^>]+)>', self.get_xml_file('_rels/document.xml.rels', False)):
            rel_type: str = re.search(r'Type="([\w/:.]+)"', relationship).group(1)
            if rel_type[-5:] == 'image':
                rel_id: str = re.search(r'Id="(\w+)"', relationship).group(1)
                rel_target: str = re.search(r'Target="([\w./]+)"', relationship).group(1)
                result[rel_id] = rel_target
        return result

    def _get_images_directory(self, is_defined_directory: bool = True) -> str:
        if not is_defined_directory:
            result: str = ''
            path_split: List[str] = self._path.split('/')
            for i in range(len(path_split) - 1):
                result += f'{path_split[i]}/'
            return result
        return self._images_dir

    def __images_extraction(self):
        import zipfile
        from PIL import Image

        docx_media_dir: str = 'word/media/'
        directory: str = self._get_images_directory()
        with zipfile.ZipFile(self._path) as archive:
            for entry in archive.infolist():
                if entry.filename[:len(docx_media_dir)] == docx_media_dir:
                    image_key: Union[None, str] = None
                    for key in self._images:
                        if self._images[key] == entry.filename[5:]:
                            image_key = key
                    if image_key is not None:
                        with archive.open(entry) as file:
                            img = Image.open(file)
                            img.save(f'{directory}{self._images[image_key][5:]}')

    def get_style(self, style_id: str):
        return self._styles[style_id]

    def get_image(self, image_id: str):
        return f'{self._get_images_directory()}{self._images[image_id][6:]}'


class XMLement(Parser):
    tag: str
    type: str = ''
    from constants.translate_formats import TranslateFormat
    translators = {}        # {TranslateFormat: translator}
    translate_format: TranslateFormat = TranslateFormat.HTML
    _is_unique: bool = False   # True if parent can containing only one this element
    # all_style_properties: Dict[str, Tuple[str, Union[str, None], bool]] = {}

    # function that return Dict of property descriptions
    _property_descriptions_getter: Callable = p_consts.get_properties_dict

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
        self._base_style = self._get_style_from_document()
        self._remove_raw_xml()

    def __str__(self):
        return self.translators[self.translate_format].translate(self)

    def _get_document(self):
        return self.parent._get_document()

    def _get_style_id(self) -> Union[str, None]:
        return None

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

    def get_parent(self):
        return self.parent

    def get_property(self, property_name: str) -> Union[str, None, bool]:
        if property_name in self._properties:
            if self._properties[property_name].value is not None:
                return self._properties[property_name].value
        if self._base_style is not None:
            return self._base_style.get_property(property_name)
        return None

    def get_style(self):
        return self._base_style

    def set_style(self, style_id: str):
        self._base_style = self._get_document().get_style(style_id)

    def set_property_value(self, property_name: str, value: Union[str, bool, None]):
        self._properties[property_name].value = value

    def _get_style_from_document(self):
        style_id = self._get_style_id()
        if style_id is not None:
            return self._get_document().get_style(style_id)
        return None


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
