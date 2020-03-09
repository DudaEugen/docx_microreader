import zipfile
import xml.dom.minidom
import re
from typing import List, Union, Tuple
from .special_characters import *
from .support_classes import ContentInf


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
        del self._raw_xml


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
        text_tuple = self._parse_element(Text)
        self.text: Text = Text(text_tuple)
        self._is_bold: bool = self._have_properties('w:b')
        self._is_italic: bool = self._have_properties('w:i')
        self._underline: str or None = self._parse_properties('w:u')
        self._language: str or None = self._parse_properties('w:lang')
        self._color: str or None = self._parse_properties('w:color')
        self._background: str or None = self._parse_properties('w:highlight')
        self._vertical_align: str or None = self._parse_properties('w:vertAlign')
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
        runs = self._parse_elements(Run)
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
    _is_can_contain_same_elements: bool = True

    def __init__(self, element: ContentInf):
        super(TableCell, self).__init__(element)
        self.tables: List[Table]
        self.__parse_tables()
        self.paragraphs: List[Paragraph]
        self.__parse_paragraphs()
        self.elements: List[Union[Paragraph, Table]]
        self.__create_queue_elements()
        self._remove_raw_xml()

    def __parse_paragraphs(self):
        self.paragraphs = []
        paragraphs = self._parse_elements(Paragraph)
        for p in paragraphs:
            paragraph = Paragraph(p)
            is_inner_paragraph = False
            for table in self.tables:
                if table._begin < paragraph._begin and table._end > paragraph._end:
                    is_inner_paragraph = True
                    break
            if not is_inner_paragraph:
                self.paragraphs.append(paragraph)

    def __parse_tables(self):
        self.tables = []
        tables = self._parse_elements(Table)
        for tbl in tables:
            table = Table(tbl)
            self.tables.append(table)

    def __create_queue_elements(self):
        self.elements = []
        self.elements.extend(self.paragraphs)
        self.elements.extend(self.tables)
        self.elements.sort(key=lambda x: x._begin)

    def __str__(self) -> str:
        result = ''
        for element in self.elements:
            result += str(element)
        if XMLement._output_format == 'html':
            return '<td>' + result + '</td>'
        return result


class TableRow(XMLement):
    _tag_name: str = 'w:tr'
    _is_can_contain_same_elements: bool = True

    def __init__(self, element: ContentInf):
        super(TableRow, self).__init__(element)
        self.cells: List[TableCell]
        self.__parse_cells()
        self._remove_raw_xml()

    def __parse_cells(self):
        self.cells = []
        cells = self._parse_elements(TableCell)
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
    _is_can_contain_same_elements: bool = True

    def __init__(self, element: ContentInf):
        super(Table, self).__init__(element)
        self.rows: List[TableRow]
        self.__parse_rows()
        self._remove_raw_xml()

    def __parse_rows(self):
        self.rows = []
        rows = self._parse_elements(TableRow)
        for r in rows:
            row = TableRow(r)
            self.rows.append(row)

    def __str__(self) -> str:
        result = ''
        for row in self.rows:
            result += str(row)
        if XMLement._output_format == 'html':
            return '<table border="1px">' + result + '</table>'
        return result


class Document(XMLement):
    _tag_name: str = 'w:body'

    def __init__(self, path: str):
        self._raw_xml: Union[str, None] = xml.dom.minidom.parseString(
            zipfile.ZipFile(path).read('word/document.xml')
        ).toprettyxml()
        doc: ContentInf = self._parse_element(Document)
        super(Document, self).__init__(doc)
        self.tables: List[Table]
        self.__parse_tables()
        self.paragraphs: List[Paragraph]
        self.__parse_paragraphs()
        self.elements: List[Union[Paragraph, Table]]
        self.__create_queue_elements()
        self._remove_raw_xml()

    def __parse_paragraphs(self):
        self.paragraphs = []
        paragraphs = self._parse_elements(Paragraph)
        for p in paragraphs:
            paragraph = Paragraph(p)
            is_inner_paragraph = False
            for table in self.tables:
                if table._begin < paragraph._begin and table._end > paragraph._end:
                    is_inner_paragraph = True
                    break
            if not is_inner_paragraph:
                self.paragraphs.append(paragraph)

    def __parse_tables(self):
        self.tables = []
        tables = self._parse_elements(Table)
        for tbl in tables:
            table = Table(tbl)
            self.tables.append(table)

    def __create_queue_elements(self):
        self.elements = []
        self.elements.extend(self.paragraphs)
        self.elements.extend(self.tables)
        self.elements.sort(key=lambda x: x._begin)

    def __str__(self) -> str:
        result = ''
        for element in self.elements:
            result += str(element)
        return result
