from docx_parser import XMLcontainer, DocumentParser
from typing import List
from styles import *
from mixins.getters_setters import *
from constants import keys_consts as k_const


class Paragraph(XMLement, ParagraphPropertiesGetSetMixin):
    tag: str = k_const.Par_tag
    _properties_unificators = {
        k_const.Par_align: [('left', ['start']),
                            ('right', ['end']),
                            ('center', []),
                            ('both', []),
                            ('distribute', [])]
    }
    from translators.html_translators import ParagraphTranslatorToHTML
    translators = {
        'html': ParagraphTranslatorToHTML(),
    }

    def __init__(self, element: ET.Element, parent):
        self.runs: List[Paragraph.Run] = []
        super(Paragraph, self).__init__(element, parent)

    def _init(self):
        self.runs = self._get_elements(Paragraph.Run)

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[k_const.ParStyle].value

    def get_inner_text(self) -> Union[str, None]:
        result: str = ''
        for run in self.runs:
            result += str(run)
        return result

    class Run(XMLement, RunPropertiesGetSetMixin):
        tag: str = k_const.Run_tag

        from translators.html_translators import RunTranslatorToHTML
        translators = {
            'html': RunTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            self.text: Paragraph.Run.Text
            super(Paragraph.Run, self).__init__(element, parent)

        def _init(self):
            self.text = self._get_elements(Paragraph.Run.Text)

        def _get_style_id(self) -> Union[str, None]:
            return self._properties[k_const.CharStyle].value

        def get_inner_text(self) -> Union[str, None]:
            return str(self.text)

        class Text(XMLement):
            tag: str = k_const.Text_tag
            _is_unique = True
            from translators.html_translators import TextTranslatorToHTML
            translators = {
                'html': TextTranslatorToHTML(),
            }

            def __init__(self, element: ET.Element, parent):
                self.content: str = ''
                super(Paragraph.Run.Text, self).__init__(element, parent)

            def _init(self):
                self.content = self._element.text

            def __str__(self):
                if self.str_format in self.translators:
                    return super(Paragraph.Run.Text, self).__str__()
                return self.content

            def get_inner_text(self) -> Union[str, None]:
                return str(self.content)


class Table(XMLement, TablePropertiesGetSetMixin):
    tag: str = k_const.Tab_tag
    _properties_unificators = {
        k_const.Tab_align: [('left', ['start']),
                            ('right', ['end']),
                            ('center', [])]
    }
    from translators.html_translators import TableTranslatorToHTML
    translators = {
        'html': TableTranslatorToHTML(),
    }

    def __init__(self, element: ET.Element, parent):
        self.rows: List[Table.Row] = []
        super(Table, self).__init__(element, parent)

    def __str__(self):
        self.__define_first_and_last_head_rows()
        self.__calculate_rowspan_for_cells()
        self.__set_margins_for_cells()
        self.__set_inside_borders()
        return super(Table, self).__str__()

    def _init(self):
        self.rows = self._get_elements(Table.Row)

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[k_const.TabStyle].value

    def get_inner_text(self) -> Union[str, None]:
        result: str = ''
        for row in self.rows:
            result += str(row)
        return result

    def __define_first_and_last_head_rows(self):
        if self.rows:
            if self.rows[0].is_header():
                self.rows[0].is_first_row_in_header = True
                self.rows[0].is_last_row_in_header = True
                previous_row: Table.Row = self.rows[0]
                for row in self.rows:
                    if row != previous_row and row.is_header():
                        previous_row.is_last_row_in_header = False
                        row.is_last_row_in_header = True
                    previous_row = row

    def __calculate_rowspan_for_cells(self):
        cells: List[Table.Row.Cell] = []
        for row in self.rows:
            cells_for_delete: List[Table.Row.Cell] = []
            j: int = -1
            for i in range(len(row.cells)):
                if len(cells) == (j + 1):
                    col_span_number: int = int(row.cells[i].get_col_span()) \
                        if row.cells[i].get_col_span() is not None else 1
                    for k in range(col_span_number):
                        cells.append(row.cells[i])
                j += col_span_number
                if row.cells[i].get_property(k_const.Cell_vertical_merge) == 'restart':
                    cells[j] = row.cells[i]
                    cells[j].row_span = 1
                elif row.cells[i].get_property(k_const.Cell_vertical_merge) == 'continue':
                    cells[j].row_span += 1
                    cells_for_delete.append(row.cells[i])
            for cell in cells_for_delete:
                row.cells.remove(cell)

    def __set_margins_for_cells(self):
        for direction in "top", "bottom", "left", "right":
            if self.get_property(k_const.get_key('cell_margin', direction, "type")) is not None:
                for row in self.rows:
                    for cell in row.cells:
                        if cell.get_property(k_const.get_key('margin', direction, "size")) is None:
                            cell.set_property_value(k_const.get_key('margin', direction, "size"),
                                        self._properties[k_const.get_key('cell_margin', direction, "size")].value)
                            cell.set_property_value(k_const.get_key('margin', direction, "type"),
                                        self._properties[k_const.get_key('cell_margin', direction, "type")].value)

    def __set_inside_borders(self):
        for direction in "top", "bottom", "left", "right":
            d: str = 'horizontal' if (direction == 'top' or direction == 'bottom') else 'vertical'
            if self.get_property(k_const.get_key('borders_inside', d, "type")) is not None:
                for i in range(len(self.rows)):
                    for j in range(len(self.rows[i].cells)):
                        if not (i == 0 and direction == 'top') and \
                                not (i == (len(self.rows) - 1) and direction == 'bottom') and \
                                not (j == 0 and direction == 'left') and \
                                not (j == (len(self.rows[i].cells) - 1) and direction == 'right'):
                            if self.rows[i].cells[j].get_property(k_const.get_key('border', direction, "color")) is None or \
                                    self.rows[i].cells[j].get_property(k_const.get_key('border', direction, "color")) == 'auto':
                                self.rows[i].cells[j].set_property_value(k_const.get_key('border', direction, "color"),
                                    self._properties[k_const.get_key('borders_inside', d, "color")].value
                                )
                            if self.rows[i].cells[j].get_property(k_const.get_key('border', direction, "type")) is None:
                                self.rows[i].cells[j].set_property_value(k_const.get_key('border', direction, "type"),
                                    self._properties[k_const.get_key('borders_inside', d, "type")].value
                                )
                                self.rows[i].cells[j].set_property_value(k_const.get_key('border', direction, "size"),
                                    self._properties[k_const.get_key('borders_inside', d, "size")].value
                                )

    class Row(XMLement, RowPropertiesGetSetMixin):
        tag: str = k_const.Row_tag
        from translators.html_translators import RowTranslatorToHTML
        translators = {
            'html': RowTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            self.cells: List[Table.Row.Cell] = []
            self.is_first_row_in_header: bool = False
            self.is_last_row_in_header: bool = False
            super(Table.Row, self).__init__(element, parent)
            if self._properties[k_const.Row_is_header].value:
                self.__set_cells_as_header()

        def _init(self):
            self.cells = self._get_elements(Table.Row.Cell)

        def get_inner_text(self) -> Union[str, None]:
            result: str = ''
            for cell in self.cells:
                result += str(cell)
            return result

        def __set_cells_as_header(self):
            for cell in self.cells:
                cell.is_header = True

        def set_as_header(self, is_header: bool = True):
            super(Table.Row, self).set_as_header(is_header)
            self.__set_cells_as_header()

        class Cell(XMLcontainer, CellPropertiesGetSetMixin):
            tag: str = k_const.Cell_tag
            from translators.html_translators import CellTranslatorToHTML
            translators = {
                'html': CellTranslatorToHTML(),
            }

            def __init__(self, element: ET.Element, parent):
                self.row_span: int = 1
                self.is_header = False
                super(Table.Row.Cell, self).__init__(element, parent)


class Document(DocumentParser):

    class Body(XMLcontainer):
        tag: str = k_const.Body_tag
        _is_unique = True

        def __str__(self):
            return self.get_inner_text()

    def __init__(self, path: str):
        self.body: Document.Body
        super(Document, self).__init__(path)
        self.body: Document.Body = self._get_elements(Document.Body)
        self._remove_raw_xml()

    def __str__(self):
        return str(self.body)

    def get_inner_text(self) -> Union[str, None]:
        return str(self.body)

    def _get_document(self):
        return self
