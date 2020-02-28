import zipfile
import xml.dom.minidom
import re
from typing import List, Tuple


class XMLement:
    _output_format: str = 'html'

    def __init__(self, tag: str, content: str, position: int = 0):
        self._tag: str = tag
        self._content: str = content
        self._relative_position: int = position

    def _get_tag(self, tag: str = '') -> Tuple[str, int]:
        _tag = self._tag if tag == '' else tag
        element = re.search(rf'<{_tag}( [^\n>%]+)?>([^%]+)?</{_tag}>', self._content)
        return element.group(0), element.span()[0]

    def _get_tags(self, tag) -> List[Tuple[str, int]]:
        tags: List[Tuple[str, int]] = []
        for o in re.finditer(rf'<{tag}( [^\n>%]+)?>[^%]+?</{tag}>', self._content):
            tags.append((o.group(0), o.span()[0]))
        return tags

    def _inner_content(self) -> str:
        inner_content = re.sub(rf'<{self._tag}([^\n>%]+)?>', '', self._content)
        inner_content = re.sub(rf'</{self._tag}>', '', inner_content)
        return inner_content


class Text(XMLement):
    tag_name: str = 'w:t'

    def __init__(self, content: str, position: int, output_format: str = 'html'):
        super(Text, self).__init__(Text.tag_name, content, position)
        self._content: str = self._inner_content()

    def __str__(self) -> str:
        return self._content


class Run(XMLement):
    tag_name: str = 'w:r'

    def __init__(self, content: str, position: int):
        super(Run, self).__init__(Run.tag_name, content, position)
        text_tuple = self._get_tag(Text.tag_name)
        self.text: Text = Text(text_tuple[0], text_tuple[1])

    def __str__(self) -> str:
        if XMLement._output_format == 'html':      # TODO
            return str(self.text)
        return str(self.text)


class Paragraph(XMLement):
    tag_name: str = 'w:p'

    def __init__(self, content: str, position: int):
        super(Paragraph, self).__init__(Paragraph.tag_name, content, position)
        self.runs: List[Run] = []
        run_tuples = self._get_tags(Run.tag_name)
        for r in run_tuples:
            self.runs.append(Run(r[0], r[1]))

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
        self._document: str = xml.dom.minidom.parseString(
            zipfile.ZipFile(path).read('word/document.xml')
        ).toprettyxml()
        super(Document, self).__init__(Document.tag_name, self._document)
        self._content: str
        self._relative_position: int
        self._content, self._relative_position = self._get_tag()
        self.paragraphs: List[Paragraph] = []
        paragraph_tuples = self._get_tags(Paragraph.tag_name)
        for p in paragraph_tuples:
            self.paragraphs.append(Paragraph(p[0], p[1]))
