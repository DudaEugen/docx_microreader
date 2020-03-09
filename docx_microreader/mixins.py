from typing import List, Union


class ContainerMixin:

    def _parse_all_elements(self):
        from .models import Table, Paragraph
        self.tables: List[Table]
        self._parse_tables()
        self.paragraphs: List[Paragraph]
        self._parse_paragraphs()
        self.elements: List[Union[Paragraph, Table]]
        self._create_queue_elements()

    def _parse_paragraphs(self):
        self.paragraphs = []
        from .models import Paragraph
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

    def _parse_tables(self):
        self.tables = []
        from .models import Table
        tables = self._parse_elements(Table)
        for tbl in tables:
            table = Table(tbl)
            self.tables.append(table)

    def _create_queue_elements(self):
        self.elements = []
        self.elements.extend(self.paragraphs)
        self.elements.extend(self.tables)
        self.elements.sort(key=lambda x: x._begin)

    def __str__(self) -> str:
        result = ''
        for element in self.elements:
            result += str(element)
        return result
