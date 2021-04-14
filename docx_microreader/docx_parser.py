import xml.etree.ElementTree as ET
from .constants.namespaces import namespaces, check_namespace_of_tag
from typing import Union, List, Callable, Dict, Tuple
import re
from .properties import Property, PropertyDescription
from .constants import property_enums as pr_const
from .constants.property_enums import DefaultStyle


class Parser:
    element_description = None

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return []

    def __init__(self, element: ET.Element):
        self._properties: Dict[str, Property] = self.__class__._parse_properties(element)
        self.inner_elements: list = self._parse_all_inner_elements(element)

    def __possible_inner_elements(self) -> list:
        import sys

        result = []
        classes_descriptions = self.__class__._possible_inner_elements_descriptions()
        for v in classes_descriptions:
            if isinstance(v, tuple):
                result.append(getattr(sys.modules[v[0]], v[1]))
            else:
                result.append(v)
        return result

    def _parse_all_inner_elements(self, element: ET.Element) -> list:
        result: list = []
        for el in element.findall('./'):
            elem = self._parse_element(el)
            if elem is not None:
                result.append(elem)
        return result

    def _parse_element(self, element: ET.Element):
        tags: Dict[str, Callable] = {
            check_namespace_of_tag(el.element_description.tag): el for el in self.__possible_inner_elements()
        }
        return tags[element.tag](element, self) if element.tag in tags else None

    @classmethod
    def _parse_properties(cls, element: ET.Element) -> Dict[str, Property]:
        result: Dict[str, Property] = {}
        properties_description_dict = cls.element_description.get_property_descriptions_dict()
        for key, pr in properties_description_dict.items():
            if pr.tag is not None:
                property_element: Union[ET.Element, None] = Parser.__find_property_element(pr, element)
                if property_element is not None:
                    result[key] = Parser.__find_property(property_element, pr,
                                                         properties_description_dict[key].tag_property)
                else:
                    result[key] = Property(None)
            else:
                result[key] = Property(element.get(check_namespace_of_tag(pr.tag_property)))
        return result

    @staticmethod
    def __find_property_element(description: PropertyDescription, element: ET.Element) -> Union[ET.Element, None]:
        """
        find element by propertyDescription
        """
        if isinstance(description.get_wrapped_tags(), list):
            for tag in description.get_wrapped_tags():
                result: Union[ET.Element, None] = element.find(tag, namespaces)
                if result is not None:
                    return result
            return None
        return element.find(description.get_wrapped_tags(), namespaces)

    @staticmethod
    def __find_property(property_element: ET.Element, pr: PropertyDescription,
                        tags: Union[List[str], str, None]) -> Property:
        """
        find property in element
        """
        if isinstance(tags, list):
            for tag_prop in tags:
                prop: Union[None, str] = property_element.get(check_namespace_of_tag(tag_prop))
                if prop is not None:
                    return Property(prop)
            return Property(None)
        else:
            if pr.is_can_be_miss:
                value: [str, None] = property_element.get(check_namespace_of_tag(pr_const.MissedPropertyAttribute))
                if value is None:
                    return Property(Property.Missed())
                if value == '1':
                    return Property(True)
                elif value == '0':
                    return Property(False)
                return Property(value)
            else:
                return Property(property_element.get(check_namespace_of_tag(tags)))


class DocumentParser(Parser):

    def __init__(self, path: str, path_for_images: Union[None, str] = None):
        from os.path import abspath

        self._path: str = abspath(path).replace('\\', '/')
        self._styles: dict = {}
        self._default_styles: dict = {}
        self._parse_default_styles()
        self._parse_styles()
        self._images_dir: str = abspath(path_for_images).replace('\\', '/') + '/' if path_for_images is not None else \
            self._get_images_directory(False)
        self._images: Dict[str, str] = self._parse_images_relationships()
        self.__images_extraction()
        super(DocumentParser, self).__init__(self.get_xml_file('document.xml'))

    @classmethod
    def _parse_properties(cls, element: ET.Element) -> Dict[str, Property]:
        return {}

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

    def __parse_style(self, element: ET.Element, default_type=None):
        from .styles import ParagraphStyle, CharacterStyle, TableStyle, NumberingStyle

        types: Dict[str, Callable] = {
            ParagraphStyle.element_description.type: ParagraphStyle,
            CharacterStyle.element_description.type: CharacterStyle,
            TableStyle.element_description.type: TableStyle,
            NumberingStyle.element_description.type: NumberingStyle,
        }

        parameters: Tuple[str, str, bool, bool] = (
            element.get(check_namespace_of_tag(pr_const.StyleProperty.TYPE.description.tag_property),
                        default=default_type),
            element.get(check_namespace_of_tag(pr_const.StyleProperty.ID.description.tag_property)),
            False if element.get(check_namespace_of_tag(
                pr_const.StyleProperty.DEFAULT.description.tag_property)
            ) is None else True,
            False if element.get(check_namespace_of_tag(
                pr_const.StyleProperty.CUSTOM.description.tag_property)
            ) is None else True,
        )
        return types[parameters[0]](element, self, parameters[1],
                                    parameters[2], parameters[3]) if parameters[0] in types else None

    def _parse_default_styles(self):
        for default_style in DefaultStyle:
            el = self.get_xml_file('styles.xml').find('./w:docDefaults/' + default_style.tag(), namespaces)
            if el is not None:
                elem = self.__parse_style(el, default_style.style_type())
                if elem is not None:
                    self._default_styles[default_style.key] = elem

    def _parse_styles(self):
        for el in self.get_xml_file('styles.xml').findall('./' + pr_const.Style.tag(), namespaces):
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
        return self._styles.get(style_id)

    def get_default_style(self, style_type: DefaultStyle):
        return self._default_styles.get(style_type.key)

    def get_image(self, image_id: str):
        return f'{self._get_images_directory()}{self._images[image_id][6:]}'


class XMLement(Parser):
    element_description: Union[pr_const.Element, pr_const.Style, pr_const.SubStyle]
    from .constants.translate_formats import TranslateFormat

    # {TranslateFormat: translator} tarnslator must have method:
    # def translate(self, xml_element, translated_inner_elements: str)
    translators = {}
    translate_format: TranslateFormat = TranslateFormat.HTML

    # first element of Tuple is correct variant of property value; second element is variants of this value
    # _properties_validate method set correct variant if find value equal of one of variant
    # if value of property not equal one of variants or correct variant set None
    _properties_unificators: Dict[str, List[Tuple[str, List[str]]]] = {}

    _default_style: Union[None, DefaultStyle] = None

    def __init__(self, element: ET.Element, parent):
        self.parent: Union[XMLement, None] = parent
        super(XMLement, self).__init__(element)
        self._properties_unificate()
        self._base_style = self._get_style_from_document()
        self._set_default_style_of_class()

    def translate(self, to_format: Union[TranslateFormat, str, None] = None, is_recursive_translate: bool = True):
        """
        :param to_format: using translate_format of element if None
        :param is_recursive_translate: pass to_format to inner element if True
        """
        from docx_microreader.constants.translate_formats import TranslateFormat
        translator = self.translators[TranslateFormat(to_format)] if to_format is not None else \
                     self.translators[self.translate_format]
        translated_inner_elements = []
        for el in self.inner_elements:
            if is_recursive_translate:
                translated_inner_elements.append(el.translate(to_format, is_recursive_translate))
            else:
                translated_inner_elements.append(el.translate())
        return translator.translate(self, translated_inner_elements)

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

    @staticmethod
    def _key_of_property(property_name) -> str:
        """
        :param property_name: str or instance of Enum from constants.property_enums
        :return: key of property
        """
        return property_name if isinstance(property_name, str) else property_name.key

    def get_parent(self):
        return self.parent

    def get_property(self, property_name, is_find_missed_or_true: bool = True):
        """
        :param property_name: key of property (str or instance of Enum from constants.property_enums)
        :param is_find_missed_or_true: recursive find value if result equal Property.Missed. Return True if not find
        :return: Property.value or None
        """
        result = None
        key: str = XMLement._key_of_property(property_name)

        if key in self._properties:
            result = self._properties[key].value
            if result is not None and (not isinstance(result, Property.Missed) or not is_find_missed_or_true):
                return result

        if self._base_style is not None:
            base_style_result = self._base_style.get_property(key, is_find_missed_or_true)
            if base_style_result is not None:
                result = base_style_result
                if not isinstance(base_style_result, Property.Missed) or not is_find_missed_or_true:
                    return base_style_result

        if isinstance(result, Property.Missed) and is_find_missed_or_true:
            return True
        return result

    def get_style(self):
        return self._base_style

    def set_style(self, style_id: str):
        self._base_style = self._get_document().get_style(style_id)

    def set_property_value(self, property_name, value: Union[str, bool, None]):
        """
        :param property_name: key of property (str or instance of Enum from constants.property_enums)
        :param value: new value
        """
        self._properties[XMLement._key_of_property(property_name)].value = value

    def _get_style_from_document(self):
        style_id = self._get_style_id()
        if style_id is not None:
            return self._get_document().get_style(style_id)
        return None

    def _get_default_style_from_document(self, style: DefaultStyle):
        return self._get_document().get_default_style(style)

    @classmethod
    def _set_default_style_of_class(cls):
        if cls._default_style is None:
            for default_style in DefaultStyle:
                if cls.element_description == default_style.element:
                    cls._default_style = default_style
