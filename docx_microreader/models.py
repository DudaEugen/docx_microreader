import zipfile
import xml.dom.minidom
import re
from typing import List, Union
from .special_characters import *
from .support_classes import ContentInf


class XMLement:
    _output_format: str = 'html'
    _tag_name: str

    def __init__(self, element: ContentInf):
        self._raw_xml: Union[str, None] = element.content
        self._begin: int = element.begin
        self._end: int = element.end

    def _get_element(self, element_class) -> ContentInf:
        tag: XMLement = element_class._tag_name
        element = re.search(rf'<{tag}( [^\n>%]+)?>([^%]+)?</{tag}>', self._raw_xml)
        return ContentInf(element.group(0), element.span())

    def _get_elements(self, element_class) -> List[ContentInf]:
        tag: XMLement = element_class._tag_name
        elements: List[ContentInf] = []
        for o in re.finditer(rf'<{tag}( [^\n>%]+)?>[^%]+?</{tag}>', self._raw_xml):
            elements.append(ContentInf(o.group(0), o.span()))
        return elements

    def _inner_content(self) -> str:
        inner_content = re.sub(rf'<{self._tag_name}([^\n>%]+)?>', '', self._raw_xml)
        inner_content = re.sub(rf'</{self._tag_name}>', '', inner_content)
        return inner_content

    def _get_properties(self, tag_name: str) -> Union[str, None]:
        properties = re.search(rf'<{self._tag_name}Pr>([^%]+)?</{self._tag_name}Pr>', self._raw_xml).group(0)
        prop = re.search(rf'<{tag_name} w:val="([^"%]+)"/>', properties)
        if prop:
            p = prop.group(0)
            begin = p.find('"') + 1
            end = p.find('"', begin)
            return p[begin: end]

    def _have_properties(self, tag_name: str) -> bool:
        properties = re.search(rf'<{self._tag_name}([^\n>%]+)?>([^%]+)?</{self._tag_name}>', self._raw_xml).group(0)
        return True if (properties.find(rf'<{tag_name}/>') != -1) else False

    def _remove_raw_xml(self):
        self._raw_xml = None


class Text(XMLement):
    _tag_name: str = 'w:t'

    def __init__(self, element: ContentInf):
        super(Text, self).__init__(element)
        self._content: str = self._inner_content()
        self._remove_raw_xml()

    def __str__(self) -> str:
        return self._content


class Run(XMLement):
    _tag_name: str = 'w:r'

    def __init__(self,  element: ContentInf):
        super(Run, self).__init__(element)
        text_tuple = self._get_element(Text)
        self.text: Text = Text(text_tuple)
        self._is_bold: bool = self._have_properties('w:b')
        self._is_italic: bool = self._have_properties('w:i')
        self._underline: str or None = self._get_properties('w:u')
        self._language: str or None = self._get_properties('w:lang')
        self._color: str or None = self._get_properties('w:color')
        self._background: str or None = self._get_properties('w:highlight')
        self._vertical_align: str or None = self._get_properties('w:vertAlign')
        self._remove_raw_xml()

    def __str__(self) -> str:
        if XMLement._output_format == 'html':      # TODO
            result: str = str(self.text)
            for char in characters_html_first:
                result = re.sub(char, characters_html_first[char], result)
            for char in characters_html:
                result = re.sub(char, characters_html[char], result)

            if self._is_bold:
                result = '<b>' + result + '</b>'
            if self._is_italic:
                result = '<i>' + result + '</i>'
            if self._vertical_align is not None:
                if self._vertical_align == 'subscript':
                    result = '<sub>' + result + '</sub>'
                elif self._vertical_align == 'superscript':
                    result = '<sup>' + result + '</sup>'
            return result
        return str(self.text)


class Paragraph(XMLement):
    _tag_name: str = 'w:p'

    def __init__(self, element: ContentInf):
        super(Paragraph, self).__init__(element)
        self.runs: List[Run]
        self.__get_runs()
        self._remove_raw_xml()

    def __get_runs(self):
        self.runs = []
        runs = self._get_elements(Run)
        for r in runs:
            run = Run(r)
            self.runs.append(run)

    def __str__(self) -> str:
        result = ''
        for run in self.runs:
            result += str(run)
        if XMLement._output_format == 'html':
            return '<p>' + result + '</p>'
        return result + '\n'


class TableCell(XMLement):
    _tag_name: str = 'w:tc'

    def __init__(self, element: ContentInf):
        super(TableCell, self).__init__(element)
        self.paragraphs: List[Paragraph]
        self.__get_paragraphs()
        self._remove_raw_xml()

    def __get_paragraphs(self):
        self.paragraphs = []
        paragraphs = self._get_elements(Paragraph)
        for par in paragraphs:
            paragraph = Paragraph(par)
            self.paragraphs.append(paragraph)

    def __str__(self) -> str:
        result = ''
        for paragraph in self.paragraphs:
            result += str(paragraph)
        if XMLement._output_format == 'html':
            return '<td>' + result + '</td>'
        return result


class TableRow(XMLement):
    _tag_name: str = 'w:tr'

    def __init__(self, element: ContentInf):
        super(TableRow, self).__init__(element)
        self.cells: List[TableCell]
        self.__get_cells()
        self._remove_raw_xml()

    def __get_cells(self):
        self.cells = []
        cells = self._get_elements(TableCell)
        for c in cells:
            cell = TableCell(c)
            self.cells.append(cell)

    def __str__(self) -> str:
        result = ''
        for cell in self.cells:
            result += str(cell)
        if XMLement._output_format == 'html':
            return '<tr>' + result + '</tr>'
        return result + '\n'


class Table(XMLement):
    _tag_name: str = 'w:tbl'

    def __init__(self, element: ContentInf):
        super(Table, self).__init__(element)
        self.rows: List[TableRow]
        self.__get_rows()
        self._remove_raw_xml()

    def __get_rows(self):
        self.rows = []
        rows = self._get_elements(TableRow)
        for r in rows:
            row = TableRow(r)
            self.rows.append(row)

    def __str__(self) -> str:
        result = ''
        for row in self.rows:
            result += str(row)
        if XMLement._output_format == 'html':
            return '<table>' + result + '</table>'
        return result


class Document(XMLement):
    _tag_name: str = 'w:body'

    def __init__(self, path: str):
        self._raw_xml: Union[str, None] = xml.dom.minidom.parseString(
            zipfile.ZipFile(path).read('word/document.xml')
        ).toprettyxml()
        doc: ContentInf = self._get_element(Document)
        super(Document, self).__init__(doc)
        self.tables: List[Table]
        self.__get_tables()
        self.paragraphs: List[Paragraph]
        self.__get_paragraphs()
        self.elements: List[Union[Paragraph, Table]]
        self.__create_queue_elements()
        self._remove_raw_xml()

    def __get_paragraphs(self):
        self.paragraphs = []
        paragraphs = self._get_elements(Paragraph)
        for p in paragraphs:
            paragraph = Paragraph(p)
            is_inner_paragraph = False
            for table in self.tables:
                if table._begin < paragraph._begin and table._end > paragraph._end:
                    is_inner_paragraph = True
                    break
            if not is_inner_paragraph:
                self.paragraphs.append(paragraph)

    def __get_tables(self):
        self.tables = []
        tables = self._get_elements(Table)
        for tbl in tables:
            table = Table(tbl)
            self.tables.append(table)

    def __create_queue_elements(self):
        self.elements = []
        self.elements.extend(self.paragraphs)
        self.elements.extend(self.tables)
        self.elements.sort(key=lambda x: x._begin)

