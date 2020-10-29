from docx_parser import XMLcontainer, DocumentParser, XMLement
import xml.etree.ElementTree as ET
from typing import List, Callable, Dict
from mixins.getters_setters import *
from constants import keys_consts as k_const
from constants.translate_formats import TranslateFormat


class Drawing(XMLement):
    tag = k_const.ElementTag.DRAWING
    _is_unique = True
    from translators.html_translators import ContainerTranslatorToHTML
    translators = {
        TranslateFormat.HTML: ContainerTranslatorToHTML(),
    }

    def __init__(self, element: ET.Element, parent):
        self.image: Union[Drawing.Image, None] = None
        super(Drawing, self).__init__(element, parent)

    def _init(self):
        self.image = self._get_elements(Drawing.Image)

    def get_inner_text(self) -> Union[str, None]:
        if self.image is not None:
            return str(self.image)
        return ''

    def get_size(self, direction: str) -> int:
        return int(self._properties[k_const.get_property_key(
            k_const.Element.DRAWING, direction=direction, property_name=k_const.PropertyName.SIZE
        )].value)

    class Image(XMLement):
        tag = k_const.ElementTag.IMAGE
        _is_unique = True
        from translators.html_translators import ImageTranslatorToHTML
        translators = {
            TranslateFormat.HTML: ImageTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            super(Drawing.Image, self).__init__(element, parent)

        def get_path(self):
            return self._get_document().get_image(self._properties[k_const.IMAGE_ID].value)

        def get_size(self, direction: str) -> int:
            return self.get_parent().get_size(direction)


class Paragraph(XMLement, ParagraphPropertiesGetSetMixin):
    tag = k_const.ElementTag.PARAGRAPH
    _properties_unificators = {
        k_const.PARAGRAPH_ALIGN: [('left', ['start']),
                                  ('right', ['end']),
                                  ('center', []),
                                  ('both', []),
                                  ('distribute', [])]
    }
    from translators.html_translators import ParagraphTranslatorToHTML
    translators = {
        TranslateFormat.HTML: ParagraphTranslatorToHTML(),
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
        tag = k_const.ElementTag.RUN

        from translators.html_translators import RunTranslatorToHTML
        translators = {
            TranslateFormat.HTML: RunTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            self.text: Paragraph.Run.Text
            self.image: Union[Drawing, None] = None
            super(Paragraph.Run, self).__init__(element, parent)

        def _init(self):
            text: Union[str, None] = self._get_elements(Paragraph.Run.Text)
            self.text = text if text is not None else ''
            self.image = self._get_elements(Drawing)

        def _get_style_id(self) -> Union[str, None]:
            return self._properties[k_const.CharStyle].value

        def get_inner_text(self) -> Union[str, None]:
            if self.image is None:
                return str(self.text)
            return str(self.image)

        def get_property(self, property_name: str) -> Union[str, None, bool]:
            result: Union[str, None, bool] = super(Paragraph.Run, self).get_property(property_name)
            return result if result is not None else self.parent.get_property(property_name)

        class Text(XMLement):
            tag = k_const.ElementTag.TEXT
            _is_unique = True
            from translators.html_translators import TextTranslatorToHTML
            translators = {
                TranslateFormat.HTML: TextTranslatorToHTML(),
            }

            def __init__(self, element: ET.Element, parent):
                self.content: str = ''
                super(Paragraph.Run.Text, self).__init__(element, parent)

            def _init(self):
                self.content = self._element.text

            def __str__(self):
                if self.translate_format in self.translators:
                    return super(Paragraph.Run.Text, self).__str__()
                return self.content

            def get_inner_text(self) -> Union[str, None]:
                return str(self.content)


class Table(XMLement, TablePropertiesGetSetMixin):
    tag = k_const.ElementTag.TABLE
    _properties_unificators = {
        k_const.TABLE_ALIGN: [('left', ['start']),
                            ('right', ['end']),
                            ('center', [])]
    }
    from translators.html_translators import TableTranslatorToHTML
    translators = {
        TranslateFormat.HTML: TableTranslatorToHTML(),
    }

    def __init__(self, element: ET.Element, parent):
        self.rows: List[Table.Row] = []
        self.header_row_number: int = 0
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
        self.header_row_number = 0
        if self.rows:
            if self.rows[0].is_header():
                self.header_row_number = 1
                self.rows[0].is_first_row_in_header = True
                self.rows[0].is_last_row_in_header = True
                previous_row: Table.Row = self.rows[0]
                for row in self.rows:
                    if row != previous_row and row.is_header():
                        self.header_row_number += 1
                        previous_row.is_last_row_in_header = False
                        row.is_last_row_in_header = True
                    previous_row = row

    def __calculate_rowspan_for_cells(self):
        cell_for_row_span: Dict[int, Table.Row.Cell] = {}
        for row in self.rows:
            cells_for_delete: List[Table.Row.Cell] = []
            col: int = 0
            for cell in row.cells:
                if cell.get_property(k_const.CELL_VERTICAL_MARGE) == 'restart':
                    cell_for_row_span[col] = cell
                    cell_for_row_span[col].row_span = 1
                elif cell.get_property(k_const.CELL_VERTICAL_MARGE) == 'continue' or \
                        cell.get_property(k_const.CELL_IS_VERTICAL_MARGE_CONTINUE):
                    cell_for_row_span[col].row_span += 1
                    cells_for_delete.append(cell)
                col_span = cell.get_col_span()
                col += int(col_span) if col_span is not None else 1
            for cell in cells_for_delete:
                row.cells.remove(cell)

    def is_use_style_of_first_row(self) -> bool:
        if self._properties[k_const.TABLE_FIRST_ROW_STYLE_LOOK].value is None:
            return False
        return self._properties[k_const.TABLE_FIRST_ROW_STYLE_LOOK].value == '1'

    def set_as_use_style_of_first_row(self, is_use: bool):
        self._properties[k_const.TABLE_FIRST_ROW_STYLE_LOOK].value = None if not is_use else '1'

    def is_use_style_of_first_column(self) -> bool:
        if self._properties[k_const.TABLE_FIRST_COLUMN_STYLE_LOOK].value is None:
            return False
        return self._properties[k_const.TABLE_FIRST_COLUMN_STYLE_LOOK].value == '1'

    def set_as_use_style_of_first_column(self, is_use: bool):
        self._properties[k_const.TABLE_FIRST_COLUMN_STYLE_LOOK].value = None if not is_use else '1'

    def is_use_style_of_last_row(self) -> bool:
        if self._properties[k_const.TABLE_LAST_ROW_STYLE_LOOK].value is None:
            return False
        return self._properties[k_const.TABLE_LAST_ROW_STYLE_LOOK].value == '1'

    def set_as_use_style_of_last_row(self, is_use: bool):
        self._properties[k_const.TABLE_LAST_ROW_STYLE_LOOK].value = None if not is_use else '1'

    def is_use_style_of_last_column(self) -> bool:
        if self._properties[k_const.TABLE_LAST_COLUMN_STYLE_LOOK].value is None:
            return False
        return self._properties[k_const.TABLE_LAST_COLUMN_STYLE_LOOK].value == '1'

    def set_as_use_style_of_last_column(self, is_use: bool):
        self._properties[k_const.TABLE_LAST_COLUMN_STYLE_LOOK].value = None if not is_use else '1'

    def is_use_style_of_horizontal_banding(self) -> bool:
        if self._properties[k_const.TABLE_NO_HORIZONTAL_BANDING].value is None:
            return True
        return self._properties[k_const.TABLE_NO_HORIZONTAL_BANDING].value == '0'

    def set_as_use_style_of_horizontal_banding(self, is_use: bool):
        self._properties[k_const.TABLE_NO_HORIZONTAL_BANDING].value = None if is_use else '0'

    def is_use_style_of_vertical_banding(self) -> bool:
        if self._properties[k_const.TABLE_NO_VERTICAL_BANDING].value is None:
            return True
        return self._properties[k_const.TABLE_NO_VERTICAL_BANDING].value == '0'

    def set_as_use_style_of_vertical_banding(self, is_use: bool):
        self._properties[k_const.TABLE_NO_VERTICAL_BANDING].value = None if is_use else '0'

    class Row(XMLement, RowPropertiesGetSetMixin):
        tag = k_const.ElementTag.ROW
        from translators.html_translators import RowTranslatorToHTML
        translators = {
            TranslateFormat.HTML: RowTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            self.cells: List[Table.Row.Cell] = []
            self.is_first_row_in_header: bool = False
            self.is_last_row_in_header: bool = False
            self.index_in_table: int = -1
            super(Table.Row, self).__init__(element, parent)
            if self._properties[k_const.ROW_IS_HEADER].value:
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

        def get_parent_table(self):
            return self.get_parent()

        def set_as_header(self, is_header: bool = True):
            super(Table.Row, self).set_as_header(is_header)
            self.__set_cells_as_header()

        def is_first_in_table(self) -> bool:
            return self.index_in_table == 0

        def is_last_in_table(self) -> bool:
            return self.index_in_table == (len(self.get_parent_table().rows) - 1)

        def is_odd(self) -> bool:
            if self.is_header():
                return False
            offset: int = 1 - self.get_parent_table().header_row_number
            if self.get_parent_table().is_use_style_of_first_row() and offset == 1:
                offset = 0
            return (self.index_in_table + offset) % 2 == 1

        def is_even(self) -> bool:
            if self.is_header():
                return False
            return not self.is_odd()

        class Cell(XMLcontainer, CellPropertiesGetSetMixin):
            tag = k_const.ElementTag.CELL
            from translators.html_translators import CellTranslatorToHTML
            translators = {
                TranslateFormat.HTML: CellTranslatorToHTML(),
            }

            def __init__(self, element: ET.Element, parent):
                self.row_span: int = 1
                self.is_header: bool = False
                self.index_in_row: int = -1
                super(Table.Row.Cell, self).__init__(element, parent)

            def __get_table_area_style(self, table_area_style_type: str):
                return self.get_parent_table().get_style().get_table_area_style(table_area_style_type)

            def __get_property_of_table_area_style(self, property_name: str, table_area_style_type: str):
                table_area_style = self.__get_table_area_style(table_area_style_type)
                if table_area_style is not None:
                    return table_area_style.get_property(property_name)
                return None

            def __get_property_of_top_left_cell_area_style(self, property_name: str) -> Tuple[Union[str, None, bool],
                                                                                              bool]:
                if self.is_top() and self.get_parent_table().is_use_style_of_first_row() and \
                        self.is_first_in_row() and self.get_parent_table().is_use_style_of_first_column():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabTopLeftCellStyle_type), \
                           True
                return None, False

            def __get_property_of_top_right_cell_area_style(self, property_name: str) -> Tuple[Union[str, None, bool],
                                                                                               bool]:
                if self.is_top() and self.get_parent_table().is_use_style_of_first_row() and \
                        self.is_last_in_row() and self.get_parent_table().is_use_style_of_last_column():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabTopRightCellStyle_type), \
                           True
                return None, False

            def __get_property_of_bottom_left_cell_area_style(self, property_name: str) -> Tuple[Union[str, None, bool],
                                                                                                 bool]:
                if self.is_bottom() and self.get_parent_table().is_use_style_of_last_row() and \
                        self.is_first_in_row() and self.get_parent_table().is_use_style_of_first_column():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabBottomLeftCellStyle_type), \
                           True
                return None, False

            def __get_property_of_bottom_right_cell_area_style(self, property_name: str) -> Tuple[Union[str, None,
                                                                                                        bool], bool]:
                if self.is_bottom() and self.get_parent_table().is_use_style_of_last_row() and \
                        self.is_last_in_row() and self.get_parent_table().is_use_style_of_last_column():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabBottomRightCellStyle_type),\
                           True
                return None, False

            def __get_property_of_first_column_area_style(self, property_name: str) -> Tuple[Union[str, None, bool],
                                                                                             bool]:
                if self.is_first_in_row() and self.get_parent_table().is_use_style_of_first_column():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabFirstColumnStyle_type), \
                           True
                return None, False

            def __get_property_of_last_column_area_style(self, property_name: str) -> Tuple[Union[str, None, bool],
                                                                                            bool]:
                if self.is_last_in_row() and self.get_parent_table().is_use_style_of_last_column():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabLastColumnStyle_type), True
                return None, False

            def __get_property_of_first_row_area_style(self, property_name: str) -> Tuple[Union[str, None, bool], bool]:
                if self.is_top() and self.get_parent_table().is_use_style_of_first_row():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabFirsRowStyle_type), True
                return None, False

            def __get_property_of_last_row_area_style(self, property_name: str) -> Tuple[Union[str, None, bool], bool]:
                if self.is_bottom() and self.get_parent_table().is_use_style_of_last_row():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabLastRowStyle_type), True
                return None, False

            def __get_property_of_odd_row_area_style(self, property_name: str) -> Tuple[Union[str, None, bool], bool]:
                if self.get_parent_row().is_odd() and self.get_parent_table().is_use_style_of_horizontal_banding():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabOddRowStyle_type), True
                return None, False

            def __get_property_of_even_row_area_style(self, property_name: str) -> Tuple[Union[str, None, bool], bool]:
                if self.get_parent_row().is_even() and self.get_parent_table().is_use_style_of_horizontal_banding():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabEvenRowStyle_type), True
                return None, False

            def __get_property_of_odd_column_area_style(self, property_name: str) -> Tuple[Union[str, None, bool],
                                                                                           bool]:
                if self.is_odd() and self.get_parent_table().is_use_style_of_vertical_banding():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabOddColumnStyle_type), True
                return None, False

            def __get_property_of_even_column_area_style(self, property_name: str) -> Tuple[Union[str, None, bool],
                                                                                            bool]:
                if self.is_even() and self.get_parent_table().is_use_style_of_vertical_banding():
                    return self.__get_property_of_table_area_style(property_name, k_const.TabEvenColumnStyle_type), True
                return None, False

            def __define_table_area_style_and_get_property(self, property_name: str) -> Union[str, None, bool]:
                methods: List[Callable] = [
                    self.__get_property_of_top_left_cell_area_style,
                    self.__get_property_of_top_right_cell_area_style,
                    self.__get_property_of_bottom_left_cell_area_style,
                    self.__get_property_of_bottom_right_cell_area_style,
                    self.__get_property_of_first_row_area_style,
                    self.__get_property_of_last_row_area_style,
                    self.__get_property_of_first_column_area_style,
                    self.__get_property_of_last_column_area_style,
                    self.__get_property_of_odd_column_area_style,
                    self.__get_property_of_even_column_area_style,
                    self.__get_property_of_odd_row_area_style,
                    self.__get_property_of_even_row_area_style,
                ]
                for method in methods:
                    result = method(property_name)[0]
                    if result is not None:
                        return result

            def __define_table_area_style_and_get_property_for_horizontal_inside_borders(self, property_name: str) -> \
                    Union[str, None, bool]:
                methods: List[Callable] = [
                    self.__get_property_of_top_left_cell_area_style,
                    self.__get_property_of_top_right_cell_area_style,
                    self.__get_property_of_bottom_left_cell_area_style,
                    self.__get_property_of_bottom_right_cell_area_style,
                    self.__get_property_of_first_row_area_style,
                    self.__get_property_of_last_row_area_style,
                    self.__get_property_of_first_column_area_style,
                    self.__get_property_of_last_column_area_style,
                    self.__get_property_of_odd_column_area_style,
                    self.__get_property_of_even_column_area_style,
                ]
                for method in methods:
                    result, is_met_contition = method(property_name)
                    if result is not None or is_met_contition:
                        return result

            def __define_table_area_style_and_get_property_for_vertical_inside_borders(self, property_name: str) -> \
                    Union[str, None, bool]:
                methods: List[Callable] = [
                    self.__get_property_of_top_left_cell_area_style,
                    self.__get_property_of_top_right_cell_area_style,
                    self.__get_property_of_bottom_left_cell_area_style,
                    self.__get_property_of_bottom_right_cell_area_style,
                    self.__get_property_of_first_row_area_style,
                    self.__get_property_of_last_row_area_style,
                    self.__get_property_of_first_column_area_style,
                    self.__get_property_of_last_column_area_style,
                    self.__get_property_of_odd_row_area_style,
                    self.__get_property_of_even_row_area_style,
                ]
                for method in methods:
                    result, is_met_contition = method(property_name)
                    if result is not None or is_met_contition:
                        return result

            def __define_table_area_style_and_get_property_for_inside_borders_of_header(self, property_name: str) -> \
                    Union[str, None, bool]:
                methods: List[Callable] = [
                    self.__get_property_of_first_column_area_style,
                    self.__get_property_of_last_column_area_style,
                    self.__get_property_of_odd_column_area_style,
                    self.__get_property_of_even_column_area_style,
                ]
                for method in methods:
                    result, is_met_contition = method(property_name)
                    if result is not None or is_met_contition:
                        return result

            def get_property(self, property_name: str) -> Union[str, None, bool]:
                result = self._properties.get(property_name)
                if result is not None and result.value is not None:
                    return result.value
                result = self.__define_table_area_style_and_get_property(property_name)
                if result is not None:
                    return result
                return self.get_parent_row().get_property(property_name)

            def get_border(self, direction: str, property_name: str) -> Union[str, None]:
                """
                :param direction: top, bottom, right, left  (keys of XMLementPropertyDescriptions.Const_directions dict)
                :param property_name: color, size, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
                """
                result = self._properties.get(k_const.get_property_key(
                    k_const.Element.CELL, k_const.SubElement.BORDER, direction=direction, property_name=property_name
                ))
                if result is not None:
                    result = result.value
                if result is None or (result == 'auto' and k_const.PropertyName.COLOR.is_equal(property_name)):
                    if not (self.is_first_in_row() and k_const.Direction.LEFT.is_equal(direction)) and \
                            not (self.is_last_in_row() and k_const.Direction.RIGHT.is_equal(direction)) and \
                            not (self.is_top() and k_const.Direction.TOP.is_equal(direction)) and \
                            not (self.is_bottom() and k_const.Direction.BOTTOM.is_equal(direction)):
                        d = k_const.Direction.horizontal_or_vertical_straight(direction)
                        if k_const.Direction.BOTTOM.is_equal(direction) and \
                                self.get_parent_table().is_use_style_of_first_row() and \
                                self.get_parent_table().header_row_number > 1 and \
                                self.get_parent_row().is_header() and not self.get_parent_row().is_last_row_in_header:
                            return self.__define_table_area_style_and_get_property_for_inside_borders_of_header(
                                        k_const.get_property_key(k_const.Element.TABLE,
                                                                 subelements=[k_const.Element.CELL,
                                                                              k_const.SubElement.BORDER],
                                                                 direction=d, property_name=property_name)
                                    )
                        result = self.get_parent_table().get_inside_border(d, property_name)
                        if result is None:
                            result = self.__define_table_area_style_and_get_property(
                                k_const.get_property_key(k_const.Element.CELL, k_const.SubElement.BORDER,
                                                         direction=direction, property_name=property_name)
                            )
                            if result is None:
                                if k_const.Direction.HORIZONTAL.is_equal(d):
                                    result = self.__define_table_area_style_and_get_property_for_horizontal_inside_borders(
                                        k_const.get_property_key(k_const.Element.TABLE,
                                                                 subelements=[k_const.Element.CELL,
                                                                              k_const.SubElement.BORDER],
                                                                 direction=d, property_name=property_name)
                                    )
                                else:
                                    result = self.__define_table_area_style_and_get_property_for_vertical_inside_borders(
                                        k_const.get_property_key(k_const.Element.TABLE,
                                                                 subelements=[k_const.Element.CELL,
                                                                              k_const.SubElement.BORDER],
                                                                 direction=d, property_name=property_name)
                                    )
                return result

            def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
                """
                :param direction: top, bottom, right, left  (keys of XMLementPropertyDescriptions.Const_directions dict)
                :param property_name: color, size, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
                """
                self.set_property_value(k_const.get_property_key(
                    k_const.Element.CELL, k_const.SubElement.BORDER, direction=direction, property_name=property_name),
                    value
                )

            def get_col_span(self) -> Union[str, None]:
                return self._properties[k_const.CELL_COLUMN_SPAN].value

            def set_col_span_value(self, value: Union[str, int]):
                self._properties[k_const.CELL_COLUMN_SPAN].value = str(value)

            def get_parent_row(self):
                return self.get_parent()

            def get_parent_table(self):
                return self.get_parent_row().get_parent()

            def is_top(self) -> bool:
                return self.get_parent_row().is_first_in_table() or self.is_header

            def is_bottom(self) -> bool:
                return self.get_parent_row().index_in_table + self.row_span >= len(self.get_parent_table().rows)

            def is_first_in_row(self) -> bool:
                return self.index_in_row == 0

            def is_last_in_row(self) -> bool:
                col_span: int = 1 if self.get_col_span() is None else int(self.get_col_span())
                return self.index_in_row + col_span >= len(self.get_parent_row().cells)

            def is_odd(self) -> bool:
                offset: int = 0 if self.get_parent_table().is_use_style_of_first_column() else 1
                return (self.index_in_row + offset) % 2 == 1

            def is_even(self) -> bool:
                return not self.is_odd()


class Document(DocumentParser):

    class Body(XMLcontainer):
        tag = k_const.ElementTag.BODY
        _is_unique = True

        def __str__(self):
            return self.get_inner_text()

    def __init__(self, path: str, path_for_images: Union[None, str] = None):
        self.body: Document.Body
        super(Document, self).__init__(path, path_for_images)
        self.body: Document.Body = self._get_elements(Document.Body)
        self._remove_raw_xml()

    def __str__(self):
        return str(self.body)

    def get_inner_text(self) -> Union[str, None]:
        return str(self.body)

    def _get_document(self):
        return self
