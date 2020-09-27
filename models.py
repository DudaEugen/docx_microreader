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

    def get_property(self, property_name: str) -> Union[str, None, bool]:
        result: Union[str, None, bool] = super(Paragraph, self).get_property(property_name)
        return result if result is not None else self.parent.get_property(property_name)

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

        def get_property(self, property_name: str) -> Union[str, None, bool]:
            result: Union[str, None, bool] = super(Paragraph.Run, self).get_property(property_name)
            return result if result is not None else self.parent.get_property(property_name)

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
        return super(Table, self).__str__()

    def _init(self):
        self.rows = self._get_elements(Table.Row)
        self.__set_index_in_table_for_rows()

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[k_const.TabStyle].value

    def get_inner_text(self) -> Union[str, None]:
        result: str = ''
        for row in self.rows:
            result += str(row)
        return result

    def __set_index_in_table_for_rows(self):
        for index in range(len(self.rows)):
            self.rows[index].index_in_table = index

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
        cell_for_row_span: Dict[int, Table.Row.Cell] = {}
        for row in self.rows:
            cells_for_delete: List[Table.Row.Cell] = []
            col: int = 0
            for cell in row.cells:
                if cell.get_property(k_const.Cell_vertical_merge) == 'restart':
                    cell_for_row_span[col] = cell
                    cell_for_row_span[col].row_span = 1
                elif cell.get_property(k_const.Cell_vertical_merge) == 'continue' or \
                        cell.get_property(k_const.Cell_is_vertical_merge_continue):
                    cell_for_row_span[col].row_span += 1
                    cells_for_delete.append(cell)
                col_span = cell.get_col_span()
                col += int(col_span) if col_span is not None else 1
            for cell in cells_for_delete:
                row.cells.remove(cell)

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
            self.index_in_table: int = -1
            super(Table.Row, self).__init__(element, parent)
            if self._properties[k_const.Row_is_header].value:
                self.__set_cells_as_header()

        def _init(self):
            self.cells = self._get_elements(Table.Row.Cell)
            self.__set_index_in_row_for_cells()

        def get_inner_text(self) -> Union[str, None]:
            result: str = ''
            for cell in self.cells:
                result += str(cell)
            return result

        def __set_cells_as_header(self):
            for cell in self.cells:
                cell.is_header = True

        def __set_index_in_row_for_cells(self):
            for index in range(len(self.cells)):
                self.cells[index].index_in_row = index

        def set_as_header(self, is_header: bool = True):
            super(Table.Row, self).set_as_header(is_header)
            self.__set_cells_as_header()

        def is_first_in_table(self) -> bool:
            return self.index_in_table == 0

        def is_last_in_table(self) -> bool:
            return self.index_in_table == (len(self.get_parent().rows) - 1)

        def is_odd(self) -> bool:
            return self.index_in_table % 2 == 1

        def is_even(self) -> bool:
            return self.index_in_table % 2 == 0

        class Cell(XMLcontainer, CellPropertiesGetSetMixin):
            tag: str = k_const.Cell_tag
            from translators.html_translators import CellTranslatorToHTML
            translators = {
                'html': CellTranslatorToHTML(),
            }

            def __init__(self, element: ET.Element, parent):
                self.row_span: int = 1
                self.is_header: bool = False
                self.index_in_row: int = -1
                super(Table.Row.Cell, self).__init__(element, parent)

            def __get_table_area_style(self, table_area_style_type: str):
                return self.get_parent().get_parent().get_style().get_table_area_style(table_area_style_type)

            def __get_property_of_table_area_style(self, property_name: str, table_area_style_type: str):
                table_area_style = self.__get_table_area_style(table_area_style_type)
                if table_area_style is not None:
                    return table_area_style.get_property(property_name)
                return None

            def get_property(self, property_name: str) -> Union[str, None, bool]:
                result = self._properties.get(property_name)
                if result is not None and result.value is not None:
                    return result.value
                if self.is_top_left():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabTopLeftCellStyle_type)
                    if result is not None:
                        return result
                if self.is_top_right():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabTopRightCellStyle_type)
                    if result is not None:
                        return result
                if self.is_bottom_left():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabBottomLeftCellStyle_type)
                    if result is not None:
                        return result
                if self.is_bottom_right():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabBottomRightCellStyle_type)
                    if result is not None:
                        return result
                if self.is_first_in_row():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabFirstColumnStyle_type)
                    if result is not None:
                        return result
                if self.is_last_in_row():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabLastColumnStyle_type)
                    if result is not None:
                        return result
                if self.get_parent().is_first_in_table():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabFirsRowStyle_type)
                    if result is not None:
                        return result
                if self.get_parent().is_last_in_table():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabLastRowStyle_type)
                    if result is not None:
                        return result
                if self.is_odd():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabOddColumnStyle_type)
                    if result is not None:
                        return result
                if self.is_even():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabEvenColumnStyle_type)
                    if result is not None:
                        return result
                if self.get_parent().is_odd():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabOddRowStyle_type)
                    if result is not None:
                        return result
                if self.get_parent().is_even():
                    result = self.__get_property_of_table_area_style(property_name, k_const.TabEvenRowStyle_type)
                    if result is not None:
                        return result
                if self._base_style is not None:
                    base_style_property = self._base_style.get_property(property_name)
                    if base_style_property is not None:
                        return base_style_property
                return self.get_parent().get_property(property_name)

            def is_first_in_row(self) -> bool:
                return self.index_in_row == 0

            def is_last_in_row(self) -> bool:
                return self.index_in_row == (len(self.get_parent().cells) - 1)

            def is_odd(self) -> bool:
                return self.index_in_row % 2 == 1

            def is_even(self) -> bool:
                return self.index_in_row % 2 == 0

            def is_top_left(self) -> bool:
                return self.is_first_in_row() and self.get_parent().is_first_in_table()

            def is_top_right(self) -> bool:
                return self.is_last_in_row() and self.get_parent().is_first_in_table()

            def is_bottom_left(self) -> bool:
                return self.is_first_in_row() and self.get_parent().is_last_in_table()

            def is_bottom_right(self) -> bool:
                return self.is_last_in_row() and self.get_parent().is_last_in_table()


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
