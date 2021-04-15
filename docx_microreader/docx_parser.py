import xml.etree.ElementTree as ET
from .constants.namespaces import namespaces, check_namespace_of_tag
from typing import Union, List, Callable, Dict, Tuple, Optional
from .properties import Property, PropertyDescription
from .constants import property_enums as pr_const


class Parser:
    element_description = None
    __possible_inner_elements: Optional[dict] = None

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return []

    def __init__(self, element: ET.Element):
        self._properties: Dict[str, Property] = self.__class__._parse_properties(element)
        self.inner_elements: list = []
        self._parse_all_inner_elements(element)

    @classmethod
    def __set_possible_inner_elements(cls):
        import sys

        cls.__possible_inner_elements = {}
        classes_descriptions = cls._possible_inner_elements_descriptions()
        for v in classes_descriptions:
            if isinstance(v, tuple):
                v = getattr(sys.modules[v[0]], v[1])
            tags = v.element_description.tag.split('/')
            d = cls.__possible_inner_elements
            for tag in tags[:-1]:
                t = check_namespace_of_tag(tag)
                if not (t in d):
                    d[t] = {}
                d = d[t]
            d[check_namespace_of_tag(tags[-1])] = v

    def _parse_all_inner_elements(self, element: ET.Element):
        if self.__class__.__possible_inner_elements is None:
            self.__class__.__set_possible_inner_elements()
        self.__parse_element(element, self.__class__.__possible_inner_elements)

    def __parse_element(self, element, d):
        if not isinstance(d, dict):
            self.inner_elements.append(d(element, self))
            return
        if d:
            for el in element.findall('./'):
                if el.tag in d:
                    self.__parse_element(el, d[el.tag])

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
    document_key: str = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
    styles_key: str = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'

    def __init__(self, path: str, path_for_images: Optional[str] = None):
        from os.path import abspath

        self._path: str = abspath(path).replace('\\', '/')
        self._content: Dict[str, Optional[str]] = {
            DocumentParser.document_key: None,
            DocumentParser.styles_key: None,
        }
        self._relationships: Dict[str, Tuple[str, str]] = {}
        self._extract_content_types()
        self._extract_relationships()

        self._styles: dict = {}
        self._default_styles: dict = {}
        self._parse_default_styles()
        self._parse_styles()
        self._images_dir: str = abspath(path_for_images).replace('\\', '/') + '/' if path_for_images is not None else \
            self._get_images_directory(False)
        self._images_extraction()
        super(DocumentParser, self).__init__(self._get_xml_file(self._content[DocumentParser.document_key]))

    @classmethod
    def _parse_properties(cls, element: ET.Element) -> Dict[str, Property]:
        return {}

    def _get_xml_file(self, file: str) -> Optional[ET.Element]:
        """
        :param file: path to file in document
        :return: ElementTree of file
        """
        import xml.dom.minidom
        import zipfile

        return ET.fromstring(xml.dom.minidom.parseString(zipfile.ZipFile(self._path).read(file)).toprettyxml())

    def _extract_content_types(self):
        content_types: Optional[ET.Element] = self._get_xml_file('[Content_Types].xml')
        if content_types is not None:
            for content_type in content_types.findall('./'):
                if content_type.tag.split('}')[-1] == 'Override':
                    content_type_key: Optional[str] = content_type.get('ContentType')
                    if content_type_key in self._content:
                        self._content[content_type_key] = content_type.get('PartName')[1:]

    def _extract_relationships(self):
        relationships: Optional[ET.Element] = self._get_xml_file('word/_rels/document.xml.rels')
        for rel in relationships.findall('./'):
            self._relationships[rel.get('Id')] = (rel.get('Type').split('/')[-1], rel.get('Target'))

    def _parse_default_styles(self):
        for default_style in pr_const.DefaultStyle:
            el = self._get_xml_file(self._content[DocumentParser.styles_key]).find(
                './w:docDefaults/' + default_style.tag(), namespaces
            )
            if el is not None:
                elem = self.__parse_style(el, default_style.style_type())
                if elem is not None:
                    self._default_styles[default_style.key] = elem

    def _parse_styles(self):
        for el in self._get_xml_file(self._content[DocumentParser.styles_key]).findall(
                './' + pr_const.Style.tag(), namespaces
        ):
            elem = self.__parse_style(el)
            if elem is not None:
                self._styles[elem.id] = elem

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

    def _get_images_directory(self, is_defined_directory: bool = True) -> str:
        if not is_defined_directory:
            result: str = ''
            path_split: List[str] = self._path.split('/')
            for i in range(len(path_split) - 1):
                result += f'{path_split[i]}/'
            return result
        return self._images_dir

    def _images_extraction(self):
        import zipfile

        images: List[str] = []
        for rel_id, rel_v in self._relationships.items():
            if rel_v[0] == 'image':
                images.append(rel_v[1])

        docx_media_dir: str = 'word/'
        directory: str = self._get_images_directory()
        with zipfile.ZipFile(self._path) as z:
            for image in images:
                with open(directory + image.split('/')[-1], 'wb') as f:
                    f.write(z.read(docx_media_dir + image))

    def get_style(self, style_id: str):
        return self._styles.get(style_id)

    def get_default_style(self, style_type: pr_const.DefaultStyle):
        return self._default_styles.get(style_type.key)

    def get_image(self, image_id: str):
        return f'{self._get_images_directory()}{self._relationships[image_id][1].split("/")[-1]}'
