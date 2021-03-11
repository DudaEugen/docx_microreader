import xml.etree.ElementTree as ET
from typing import List
from ...constants.translate_formats import TranslateFormat


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

    def create_xml(self, element) -> ET.Element:
        result: ET.Element = ET.Element(element.element_description.tag)
        inner_elements = {}
        for pr_key, pr_description in element._all_properties.items():
            tags_description = pr_description.get_wrapped_tags()
            descr = tags_description if isinstance(tags_description, str) else tags_description[0]
            if isinstance(pr_description.tag_property, list):
                descr += '/' + pr_description.tag_property[0]
            elif isinstance(pr_description.tag_property, str):
                descr += '/' + pr_description.tag_property
            tags: List[str] = descr.split('/')
            d = inner_elements
            for i in range(len(tags) - 1):
                if not tags[i] in d:
                    d[tags[i]] = {}
                d = d[tags[i]]
            d[tags[-1]] = pr_key
        self.insert_elements(result, inner_elements, element._properties)
        self.add_inner_xml(result, element)
        return result

    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        return xml

    def translate(self, element) -> str:
        el = self.create_xml(element)
        return ET.tostring(el, encoding='unicode')


class TextTranslatorToXML(TranslatorToXML):
    def create_xml(self, element) -> ET.Element:
        result = super(TextTranslatorToXML, self).create_xml(element)
        # TO DO: upgrade create_xml in TranslatorToXML, add xml:space to properties of Text
        if len(element.content) > 0 and (element.content[0] == ' ' or element.content[-1] == ' '):
            result.set('xml:space', 'preserve')
        return result

    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        xml.text = element.get_inner_text()
        return xml


class RunTranslatorToXML(TranslatorToXML):
    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        inner_xml = TextTranslatorToXML().create_xml(element.text)
        xml.append(inner_xml)
        return xml


class ParagraphTranslatorToXML(TranslatorToXML):
    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        for run in element.runs:
            run_xml = RunTranslatorToXML().create_xml(run)
            xml.append(run_xml)
        return xml


class CellTranslatorToXML(TranslatorToXML):
    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        for e in element.elements:
            e_xml = e.translators[TranslateFormat.XML].create_xml(e)
            xml.append(e_xml)
        return xml


class RowTranslatorToXML(TranslatorToXML):
    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        for cell in element.cells:
            cell_xml = CellTranslatorToXML().create_xml(cell)
            xml.append(cell_xml)
        return xml


class TableTranslatorToXML(TranslatorToXML):
    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        for row in element.rows:
            row_xml = RowTranslatorToXML().create_xml(row)
            xml.append(row_xml)
        return xml


class BodyTranslatorToXML(TranslatorToXML):
    def add_inner_xml(self, xml: ET.Element, element) -> ET.Element:
        for e in element.elements:
            e_xml = e.translators[TranslateFormat.XML].create_xml(e)
            xml.append(e_xml)
        return xml


class DocumentTranslatorToXML:
    def translate(self, element) -> str:
        from ...constants.namespaces import namespaces

        doc = ET.Element('w:document')
        doc.append(BodyTranslatorToXML().create_xml(element.body))
        for key, value in namespaces.items():
            doc.set(f'xmlns:{key}', value)
        doc.set('mc:Ignorable', "w14 w15 w16se w16cid wp14")
        return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n{ET.tostring(doc, encoding="unicode")}'

    @staticmethod
    def template_directory_name() -> str:
        return 'document template'

    @staticmethod
    def template_directory_path() -> str:
        import pathlib
        current_file_path = pathlib.Path(__file__).parent.absolute()

        return f'{current_file_path}\\{DocumentTranslatorToXML.template_directory_name()}'
