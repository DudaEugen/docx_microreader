import xml.etree.ElementTree as ET
from typing import List


class TranslatorToXML:
    def insert_elements(self, parent_element: ET.Element, d: dict, properties: dict):
        result: bool = False
        for key, value in d.items():
            if isinstance(value, str):
                if isinstance(properties[value].value, str):
                    if properties[value].value is not None:
                        parent_element.set(key, properties[value].value)
                        result = True
                        # return True
                elif isinstance(properties[value].value, bool):
                    if value:
                        parent_element.append(ET.Element(key))
                        result = True
                        # return True
            else:
                element = ET.Element(key)
                res = self.insert_elements(element, value, properties)
                #if res is not None and res:
                if res:
                    parent_element.append(element)
                    result = True
                    # return True
        return result

    def create_xml(self, element, inner_elements: list) -> ET.Element:
        result: ET.Element = ET.Element(element.element_description.tag)
        properties = {}
        for pr_key, pr_description in element._all_properties.items():
            tags_description = pr_description.get_wrapped_tags()
            descr = tags_description if isinstance(tags_description, str) else tags_description[0]
            if isinstance(pr_description.tag_property, list):
                descr += '/' + pr_description.tag_property[0]
            elif isinstance(pr_description.tag_property, str):
                descr += '/' + pr_description.tag_property
            tags: List[str] = descr.split('/')
            d = properties
            for i in range(len(tags) - 1):
                if not tags[i] in d:
                    d[tags[i]] = {}
                d = d[tags[i]]
            d[tags[-1]] = pr_key
        self.insert_elements(result, properties, element._properties)

        for el in inner_elements:
            if isinstance(el, ET.Element):
                result.append(el)
            else:
                if result.text is not None:
                    result.text += str(el)
                else:
                    result.text = str(el)
        return result

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
        for el in inner_elements:
            if isinstance(el, ET.Element):
                doc.append(el)
            else:
                if doc.text is not None:
                    doc.text += str(el)
                else:
                    doc.text = str(el)
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
