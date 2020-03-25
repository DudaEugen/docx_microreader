import zipfile
import xml.dom.minidom
import re
from typing import List, Union
from .special_characters import *
from .support_classes import ContentInf
from .parser import XMLement
from .mixins import ContainerMixin


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
        self._underline: Union[str, None] = self._parse_properties('w:u')
        self._language: Union[str, None] = self._parse_properties('w:lang')
        self._color: Union[str, None] = self._parse_properties('w:color')
        self._background: Union[str, None] = self._parse_properties('w:highlight')
        self._vertical_align: Union[str, None] = self._parse_properties('w:vertAlign')
        self._remove_raw_xml()

    def __str__(self) -> str:
        if XMLement._output_format == 'html':      # TODO
            result: str = str(self.text)
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
        self.__parse_runs()
        self._remove_raw_xml()

    def __parse_runs(self):
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


class TableCell(XMLement, ContainerMixin):
    _tag_name: str = 'w:tc'
    _is_can_contain_same_elements: bool = True

    def __init__(self, element: ContentInf):
        super(TableCell, self).__init__(element)
        self._parse_all_elements()
        self._bg_color: Union[str, None] = self._parse_named_value_of_properties('w:shd', 'w:fill')
        self._remove_raw_xml()

    def __str__(self) -> str:
        result = super(TableCell, self).__str__()
        if XMLement._output_format == 'html':
            color = ' bgcolor="#' + self._bg_color + '"' if self._bg_color is not None and self._bg_color != 'auto' \
                else ''
            return rf'<td{color}>' + result + '</td>'
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


class Document(XMLement, ContainerMixin):
    _tag_name: str = 'w:body'

    def __init__(self, path: str):
        self._raw_xml: Union[str, None] = xml.dom.minidom.parseString(
            zipfile.ZipFile(path).read('word/document.xml')
        ).toprettyxml()
        doc: ContentInf = self._parse_element(Document)
        super(Document, self).__init__(doc)
        self._parse_all_elements()
        self._remove_raw_xml()
