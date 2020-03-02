import zipfile
import xml.dom.minidom
import re
from typing import List, Tuple
from .special_characters import *


class XMLement:
    _output_format: str = 'html'
    tag_name: str

    def __init__(self, tag: str, element: Tuple[str, Tuple[int, int]]):
        self._tag: str = tag
        self._content: str = element[0]
        self._begin_position: int = element[1][0]
        self._end_position: int = element[1][1]

    def _get_tag(self, tag: str = '') -> Tuple[str, Tuple[int, int]]:
        _tag = self._tag if tag == '' else tag
        element = re.search(rf'<{_tag}( [^\n>%]+)?>([^%]+)?</{_tag}>', self._content)
        return element.group(0), element.span()

    def _get_tags(self, tag) -> List[Tuple[str, Tuple[int, int]]]:
        tags: List[Tuple[str, Tuple[int, int]]] = []
        for o in re.finditer(rf'<{tag}( [^\n>%]+)?>[^%]+?</{tag}>', self._content):
            tags.append((o.group(0), o.span()))
        return tags

    def _inner_content(self) -> str:
        inner_content = re.sub(rf'<{self._tag}([^\n>%]+)?>', '', self._content)
        inner_content = re.sub(rf'</{self._tag}>', '', inner_content)
        return inner_content

    def _get_properties(self, tag_name: str) -> str or None:
        properties = re.search(rf'<{self.tag_name}Pr>([^%]+)?</{self.tag_name}Pr>', self._content).group(0)
        prop = re.search(rf'<{tag_name} w:val="([^"%]+)"/>', properties)
        if prop:
            p = prop.group(0)
            begin = p.find('"') + 1
            end = p.find('"', begin)
            return p[begin: end]

    def _have_properties(self, tag_name: str) -> bool:
        properties = re.search(rf'<{self.tag_name}([^\n>%]+)?>([^%]+)?</{self.tag_name}>', self._content).group(0)
        return True if (properties.find(rf'<{tag_name}/>') != -1) else False


class Text(XMLement):
    tag_name: str = 'w:t'

    def __init__(self, element: Tuple[str, Tuple[int, int]], output_format: str = 'html'):
        super(Text, self).__init__(Text.tag_name, element)
        self._content: str = self._inner_content()

    def __str__(self) -> str:
        return self._content


class Run(XMLement):
    tag_name: str = 'w:r'

    def __init__(self,  element: Tuple[str, Tuple[int, int]]):
        super(Run, self).__init__(Run.tag_name, element)
        text_tuple = self._get_tag(Text.tag_name)
        self.text: Text = Text(text_tuple)
        self._is_bold: bool = self._have_properties('w:b')
        self._is_italic: bool = self._have_properties('w:i')
        self._underline: str or None = self._get_properties('w:u')
        self._language: str or None = self._get_properties('w:lang')
        self._color: str or None = self._get_properties('w:color')
        self._background: str or None = self._get_properties('w:highlight')
        self._vertical_align: str or None = self._get_properties('w:vertAlign')

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
    tag_name: str = 'w:p'

    def __init__(self, element: Tuple[str, Tuple[int, int]]):
        super(Paragraph, self).__init__(Paragraph.tag_name, element)
        self.runs: List[Run] = []
        run_tuples = self._get_tags(Run.tag_name)
        for r in run_tuples:
            run = Run(r)
            self.runs.append(run)

    def __str__(self) -> str:
        result = ''
        for run in self.runs:
            result += str(run)
        if XMLement._output_format == 'html':
            return '<p>' + result + '</p>'
        return result + '\n'


class Document(XMLement):
    tag_name: str = 'w:body'

    def __init__(self, path: str):
        self._content: str = xml.dom.minidom.parseString(
            zipfile.ZipFile(path).read('word/document.xml')
        ).toprettyxml()
        doc: Tuple[str, Tuple[int, int]] = self._get_tag(Document.tag_name)
        super(Document, self).__init__(Document.tag_name, doc)
        self.paragraphs: List[Paragraph] = []
        paragraph_tuples = self._get_tags(Paragraph.tag_name)
        for p in paragraph_tuples:
            par = Paragraph(p)
            self.paragraphs.append(par)
