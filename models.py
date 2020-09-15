import xml.etree.ElementTree as ET
from docx_parser import Parser, XMLement, XMLcontainer
from typing import Union, List, Tuple
from properties import Property
from constants import *
from styles import *


class Paragraph(XMLement):
    tag: str = 'w:p'
    _properties_unificators = {
        Par_align: [('left', ['start']),
                    ('right', ['end']),
                    ('center', []),
                    ('both', []),
                    ('distribute', [])]
    }
    from translators import ParagraphTranslatorToHTML
    translators = {
        'html': ParagraphTranslatorToHTML(),
    }

    def __init__(self, element: ET.Element, parent):
        self.runs: List[Paragraph.Run] = []
        super(Paragraph, self).__init__(element, parent)

    def _init(self):
        self.runs = self._get_elements(Paragraph.Run)

    def get_inner_text(self) -> Union[str, None]:
        result: str = ''
        for run in self.runs:
            result += str(run)
        return result

    def get_align(self) -> Property:
        return self._properties[Par_align]

    def set_align_value(self, align: Union[str, None]):
        """
        :param align: 'both' or 'right' or 'center' or 'left
        """
        self._properties[Par_align].value = align

    def get_indent_left(self) -> Property:
        return self._properties[Par_indent_left]

    def set_indent_left_value(self, indent_left: Union[str, None]):
        self._properties[Par_indent_left].value = indent_left

    def get_indent_right(self) -> Property:
        return self._properties[Par_indent_right]

    def set_indent_right_value(self, indent_right: Union[str, None]):
        self._properties[Par_indent_right].value = indent_right

    def get_hanging(self) -> Property:
        return self._properties[Par_hanging]

    def set_hanging_value(self, hanging: Union[str, None]):
        self._properties[Par_hanging].value = hanging

    def get_first_line(self) -> Property:
        return self._properties[Par_first_line]

    def set_first_line_value(self, first_line: Union[str, None]):
        self._properties[Par_first_line].value = first_line

    def is_keep_lines(self) -> bool:
        return self._properties[Par_keep_lines].value

    def set_as_keep_lines(self, is_keep_lines: bool = True):
        self._properties[Par_keep_lines].value = is_keep_lines

    def is_keep_next(self) -> bool:
        return self._properties[Par_keep_next].value

    def set_as_keep_next(self, is_keep_next: bool = True):
        self._properties[Par_keep_next].value = is_keep_next

    def get_outline_level(self) -> Property:
        return self._properties[Par_outline_level]

    def set_outline_level_value(self, level: Union[str, int, None]):
        """
        :param level: 0-9
        """
        self._properties[Par_outline_level].value = level

    def get_border(self, direction: str, property_name: str) -> Property:
        """
        :param direction: top, bottom, right, left     (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, space, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self._properties[get_key('border', direction, property_name)]

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right, left     (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, space, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self._properties[get_key('border', direction, property_name)].value = value

    class Run(XMLement):
        tag: str = 'w:r'

        from translators import RunTranslatorToHTML
        translators = {
            'html': RunTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            self.text: Paragraph.Run.Text
            super(Paragraph.Run, self).__init__(element, parent)

        def _init(self):
            self.text = self._get_elements(Paragraph.Run.Text)

        def get_inner_text(self) -> Union[str, None]:
            return str(self.text)

        def get_size(self) -> Property:
            return self._properties[Run_size]

        def set_size_value(self, value: Union[str, None]):
            self._properties[Run_size].value = value

        def is_bold(self) -> bool:
            return self._properties[Run_is_bold].value

        def set_as_bold(self, is_bold: bool = True):
            self._properties[Run_is_bold].value = is_bold

        def is_italic(self) -> bool:
            return self._properties[Run_is_italic].value

        def set_as_italic(self, is_italic: bool = True):
            self._properties[Run_is_italic].value = is_italic

        def is_strike(self) -> bool:
            return self._properties[Run_is_strike].value

        def set_as_strike(self, is_strike: bool = True):
            self._properties[Run_is_strike].value = is_strike

        def get_language(self) -> Property:
            return self._properties[Run_language]

        def set_language_value(self, value: Union[str, None]):
            self._properties[Run_language].value = value

        def get_vertical_align(self) -> Property:
            return self._properties[Run_vertical_align]

        def set_vertical_align_value(self, value: Union[str, None]):
            """
            param value: 'superscript' or 'subscript'
            """
            self._properties[Run_vertical_align].value = value

        def get_color(self) -> Property:
            return self._properties[Run_color]

        def set_color_value(self, value: Union[str, None]):
            """
            param value: specifies the color as a hex value in RRGGBB format or auto
            """
            self._properties[Run_color].value = value

        def get_theme_color(self) -> Property:
            return self._properties[Run_theme_color]

        def set_theme_color_value(self, value: Union[str, None]):
            self._properties[Run_theme_color].value = value

        def get_background_color(self) -> Property:
            return self._properties[Run_background_color]

        def set_background_color_value(self, value: Union[str, None]):
            """
            param value: possible values: black, blue, cyan, darkBlue, darkCyan, darkGray, darkGreen, darkMagenta,
                                        darkRed, darkYellow, green, lightGray, magenta, none, red, white, yellow
            """
            self._properties[Run_background_color].value = value

        def get_background_fill(self) -> Property:
            return self._properties[Run_background_fill]

        def set_background_fill_value(self, value: Union[str, None]):
            """
            param value: specifies the color as a hex value in RRGGBB format or auto
            """
            self._properties[Run_background_fill].value = value

        def get_underline(self) -> Property:
            return self._properties[Run_underline]

        def set_underline_value(self, value: Union[str, None]):
            self._properties[Run_underline].value = value

        def get_underline_color(self) -> Property:
            return self._properties[Run_underline_color]

        def set_underline_color_value(self, value: Union[str, None]):
            """
            param value: specifies the color as a hex value in RRGGBB format or auto
            """
            self._properties[Run_underline_color].value = value

        def get_border(self, property_name: str) -> Property:
            """
            :param property_name: color, size, space, type
            (keys of XMLementPropertyDescriptions.Const_property_names dict)
            """
            return self._properties[get_key('border', property_name=property_name)]

        def set_border_value(self, property_name: str, value: Union[str, None]):
            """
            :param property_name: color, size, space or type
            (keys of XMLementPropertyDescriptions.Const_property_names dict)
            """
            self._properties[get_key('border', property_name=property_name)].value = value

        class Text(XMLement):
            tag: str = 'w:t'
            _is_unique = True
            from translators import TextTranslatorToHTML
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


class Table(XMLement):
    tag: str = 'w:tbl'
    _properties_unificators = {
        Tab_align: [('left', ['start']),
                    ('right', ['end']),
                    ('center', [])]
    }
    from translators import TableTranslatorToHTML
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
                    col_span_number: int = int(row.cells[i].get_col_span().value) \
                        if row.cells[i].get_col_span().value is not None else 1
                    for k in range(col_span_number):
                        cells.append(row.cells[i])
                j += col_span_number
                if row.cells[i].get_property_value(Cell_vertical_merge) == 'restart':
                    cells[j] = row.cells[i]
                    cells[j].row_span = 1
                elif row.cells[i].get_property_value(Cell_vertical_merge) == 'continue':
                    cells[j].row_span += 1
                    cells_for_delete.append(row.cells[i])
            for cell in cells_for_delete:
                row.cells.remove(cell)

    def __set_margins_for_cells(self):
        for direction in "top", "bottom", "left", "right":
            if self._is_have_viewing_property(get_key('cell_margin', direction, "type")):
                for row in self.rows:
                    for cell in row.cells:
                        if cell.get_property_value(get_key('margin', direction, "size")) is None:
                            cell.set_property_value(get_key('margin', direction, "size"),
                                                    self._properties[get_key('cell_margin', direction, "size")].value)
                            cell.set_property_value(get_key('margin', direction, "type"),
                                                    self._properties[get_key('cell_margin', direction, "type")].value)

    def __set_inside_borders(self):                                              # TO DO: testing this method
        for direction in "top", "bottom", "left", "right":
            d: str = 'horizontal' if (direction == 'top' or direction == 'bottom') else 'vertical'
            if self._is_have_viewing_property(get_key('borders_inside', d, "type")):
                for i in range(len(self.rows)):
                    for j in range(len(self.rows[i].cells)):
                        if not (i == 0 and direction == 'top') and \
                                not (i == (len(self.rows) - 1) and direction == 'bottom') and \
                                not (j == 0 and direction == 'left') and \
                                not (j == (len(self.rows[i].cells) - 1) and direction == 'right'):
                            if self.rows[i].cells[j].get_property_value(get_key('border', direction, "color")) is None or \
                                    self.rows[i].cells[j].get_property_value(get_key('border', direction, "color")) == 'auto':
                                self.rows[i].cells[j].set_property_value(get_key('border', direction, "color"),
                                    self._properties[get_key('borders_inside', d, "color")].value
                                )
                            if self.rows[i].cells[j].get_property_value(get_key('border', direction, "type")) is None:
                                self.rows[i].cells[j].set_property_value(get_key('border', direction, "type"),
                                    self._properties[get_key('borders_inside', d, "type")].value
                                )
                                self.rows[i].cells[j].set_property_value(get_key('border', direction, "size"),
                                    self._properties[get_key('borders_inside', d, "size")].value
                                )

    def get_width(self) -> Tuple[Property, Property, Property]:
        """
        return: Tuple(width, type, layout)
        """
        return self._properties[Tab_width], self._properties[Tab_width_type], self._properties[Tab_layout]

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None], layout: str = 'autofit'):
        self._properties[Tab_width].value = width
        self._properties[Tab_width_type].value = width_type
        self._properties[Tab_layout].value = layout

    def get_align(self) -> Property:
        return self._properties[Tab_align]

    def set_align_value(self, value: Union[str, None]):
        self._properties[Tab_align].value = value

    def get_border(self, direction: str, property_name: str) -> Property:
        """
        :param direction: top, bottom, right, left      (keys of XMLementPropertyDescriptions.Const_property_names dict)
        :param property_name: color, size, type         (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self._properties[get_key('border', direction, property_name)]

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right, left      (keys of XMLementPropertyDescriptions.Const_property_names dict)
        :param property_name: color, size, type         (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self._properties[get_key('border', direction, property_name)].value = value

    def get_indentation(self) -> Tuple[Property, Property]:
        """
        return: Tuple(value, type)
        """
        return self._properties[Tab_indentation], self._properties[Tab_indentation_type]

    def set_indentation_value(self, width: Union[str, None], width_type: Union[str, None]):
        self._properties[Tab_indentation].value = width
        self._properties[Tab_indentation_type].value = width_type

    def get_inside_border(self, direction: str, property_name: str) -> Property:
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color', 'size', 'type'  (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self._properties[get_key('borders_inside', direction, property_name)]

    def set_inside_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color', 'size', 'type'  (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self._properties[get_key('borders_inside', direction, property_name)].value = value

    def get_cells_margin(self, direction: str) -> Tuple[Property, Property]:
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :result: value, value_type
        """
        return self._properties[get_key('cell_margin', direction, "size")], \
               self._properties[get_key('cell_margin', direction, "type")]

    def set_cells_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param value_type: 'dxa' or 'nil'
        """
        self._properties[get_key('cell_margin', direction, "size")].value = value
        self._properties[get_key('cell_margin', direction, "type")].value = value_type

    class Row(XMLement):
        tag: str = 'w:tr'
        from translators import RowTranslatorToHTML
        translators = {
            'html': RowTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            self.cells: List[Table.Row.Cell] = []
            self.is_first_row_in_header: bool = False
            self.is_last_row_in_header: bool = False
            super(Table.Row, self).__init__(element, parent)
            if self._properties[Row_is_header].value:
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

        def is_header(self) -> bool:
            return self._properties[Row_is_header].value

        def set_as_header(self, is_header: bool = True):
            self._properties[Row_is_header].value = is_header
            self.__set_cells_as_header()

        def get_height(self) -> Tuple[Property, Property]:
            """
            return: Tuple(height, rule)
            """
            return self._properties[Row_height], self._properties[Row_height_rule]

        def set_height_value(self, value: Union[str, None], height_rule: str = 'atLeast'):
            self._properties[Row_height].value = value
            self._properties[Row_height_rule].value = height_rule

        class Cell(XMLcontainer):
            tag: str = 'w:tc'
            from translators import CellTranslatorToHTML
            translators = {
                'html': CellTranslatorToHTML(),
            }

            def __init__(self, element: ET.Element, parent):
                self.row_span: int = 1
                self.is_header = False
                super(Table.Row.Cell, self).__init__(element, parent)

            def get_fill_color(self) -> Property:
                return self._properties[Cell_fill_color]

            def set_fill_color_value(self, value: Union[str, None]):
                self._properties[Cell_fill_color].value = value

            def get_fill_theme(self) -> Property:
                return self._properties[Cell_fill_theme]

            def set_fill_theme_value(self, value: Union[str, None]):
                self._properties[Cell_fill_theme].value = value

            def get_border(self, direction: str, property_name: str) -> Property:
                """
                :param direction: top, bottom, right, left  (keys of XMLementPropertyDescriptions.Const_directions dict)
                :param property_name: color, size, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
                """
                return self._properties[get_key('border', direction, property_name)]

            def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
                """
                :param direction: top, bottom, right, left  (keys of XMLementPropertyDescriptions.Const_directions dict)
                :param property_name: color, size, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
                """
                self._properties[get_key('border', direction, property_name)].value = value

            def get_width(self) -> Tuple[Property, Property]:
                """
                return: Tuple(width, type)
                """
                return self._properties[Cell_width], self._properties[Cell_width_type]

            def set_width_value(self, width: Union[str, None], width_type: Union[str, None]):
                self._properties[Cell_width].value = width
                self._properties[Cell_width_type].value = width_type

            def get_col_span(self) -> Property:
                return self._properties[Cell_col_span]

            def set_col_span_value(self, value: int):
                self._properties[Cell_col_span].value = value

            def get_vertical_align(self) -> Property:
                return self._properties[Cell_vertical_align]

            def set_vertical_align_value(self, value: str):
                self._properties[Cell_vertical_align].value = value

            def get_text_direction(self) -> Property:
                return self._properties[Cell_text_direction]

            def set_text_direction_value(self, value: Union[str, None]):
                self._properties[Cell_text_direction].value = value

            def get_margin(self, direction: str) -> Tuple[Property, Property]:
                """
                :param direction: 'top', 'bottom', 'right', 'left'
                (keys of XMLementPropertyDescriptions.Const_directions dict)
                :return: (margin, type)
                """
                return self._properties[get_key('margin', direction, "size")], \
                       self._properties[get_key('margin', direction, "type")]

            def set_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
                """
                :param direction:  'top' or 'bottom' or 'right' or 'left'
                (keys of XMLementPropertyDescriptions.Const_directions dict)
                :param value_type: 'dxa' or 'nil'
                """
                self._properties[get_key('margin', direction, "size")].value = value
                self._properties[get_key('margin', direction, "type")].value = value_type


class Document(XMLement):

    class Body(XMLcontainer):
        tag: str = 'w:body'
        _is_unique = True

        def __str__(self):
            return self.get_inner_text()

    def __init__(self, path: str):
        self.__path: str = path
        self.body: Document.Body
        super(Document, self).__init__(Parser.get_xml_file(self.__path, 'document'), None)

    def _init(self):
        style_file: ET.Element = Parser.get_xml_file(self.__path, 'styles')
        styles: list = self._get_styles(style_file)
        self.body: Document.Body = self._get_elements(Document.Body)

    def __str__(self):
        return str(self.body)

    def get_inner_text(self) -> Union[str, None]:
        return str(self.body)
