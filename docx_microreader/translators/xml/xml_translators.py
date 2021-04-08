import xml.etree.ElementTree as ET
from typing import List


class TranslatorToXML:

    @staticmethod
    def create_properties_dict(parent: dict, rest_tags: List[str]) -> dict:
        tag: str = rest_tags[0]
        if not (tag in parent):
            parent[tag] = {}
        d: dict = parent[tag]
        if len(rest_tags) > 1:
            return TranslatorToXML.create_properties_dict(d, rest_tags[1:])
        return d

    def add_property(self, element: ET.Element, d: dict) -> bool:
        from ...properties import Property

        is_added_someone: bool = False
        for key, value in d.items():
            is_added: bool = False
            if not isinstance(value, dict):
                if value is not None:
                    if isinstance(value, bool):
                        is_added = value
                    elif isinstance(value, Property.Missed):
                        is_added = True
                    else:
                        element.set(key, value)
                        is_added = True
            else:
                new_el = ET.Element(key)
                is_added = self.add_property(new_el, value)
                if is_added:
                    element.append(new_el)
            is_added_someone = is_added_someone or is_added
        return is_added_someone

    def create_xml(self, element, inner_elements: list) -> ET.Element:
        result: ET.Element = ET.Element(element.element_description.tag)
        properties = {}
        for prop in element.element_description.props:
            tags: List[str] = ''.join(prop.description.get_wrapped_tags()).split('/')
            tag_property = prop.description.tag_property[0] if isinstance(prop.description.tag_property, list) else \
                prop.description.tag_property
            self.create_properties_dict(properties, tags)[tag_property] = element._properties[prop.key].value
        self.add_property(result, properties)
        TranslatorToXML.append_inner_elements(result, inner_elements)
        return result

    @staticmethod
    def append_inner_elements(container, inner_elements):
        for el in inner_elements:
            if isinstance(el, ET.Element):
                container.append(el)
            else:
                if container.text is not None:
                    container.text += str(el)
                else:
                    container.text = str(el)

    def translate(self, element, inner_elements: list) -> ET.Element:
        el = self.create_xml(element, inner_elements)
        return el


class TextTranslatorToXML(TranslatorToXML):
    def create_xml(self, element, inner_elements: list) -> ET.Element:
        result = super(TextTranslatorToXML, self).create_xml(element, inner_elements)
        # TO DO: upgrade create_xml in TranslatorToXML, add xml:space to properties of Text
        if len(element.content) > 0 and (element.content[0] == ' ' or element.content[-1] == ' '):
            result.set('xml:space', 'preserve')
        return result


class RunTranslatorToXML(TranslatorToXML):
    pass


class ParagraphTranslatorToXML(TranslatorToXML):
    pass


class CellTranslatorToXML(TranslatorToXML):
    pass


class RowTranslatorToXML(TranslatorToXML):
    pass


class TableTranslatorToXML(TranslatorToXML):
    pass


class BodyTranslatorToXML(TranslatorToXML):
    pass


class DocumentTranslatorToXML:
    def translate(self, element, inner_elements: list) -> ET.Element:
        from ...constants.namespaces import namespaces

        doc = ET.Element('w:document')
        for key, value in namespaces.items():
            doc.set(f'xmlns:{key}', value)
        doc.set('mc:Ignorable', "w14 w15 w16se w16cid wp14")
        TranslatorToXML.append_inner_elements(doc, inner_elements)
        return doc

    @staticmethod
    def template_directory_name() -> str:
        return 'document template'

    @staticmethod
    def template_directory_path() -> str:
        import pathlib
        current_file_path = pathlib.Path(__file__).parent.absolute()

        return f'{current_file_path}\\{DocumentTranslatorToXML.template_directory_name()}'

    @staticmethod
    def document_header() -> str:
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
