import xml.etree.ElementTree as ET
from docx_parser import Parser, XMLement, XMLcontainer
from typing import Union, List, Dict, Tuple
from propepty_models import PropertyDescription, Property


class Paragraph(XMLement):
    tag: str = 'w:p'
    _all_properties = {
        'align': PropertyDescription('w:pPr', 'w:jc', 'w:val', True),
        'indent_left': PropertyDescription('w:pPr', 'w:ind', 'w:left', True),
        'indent_right': PropertyDescription('w:pPr', 'w:ind', 'w:right', True),
        'hanging': PropertyDescription('w:pPr', 'w:ind', 'w:hanging', True),
        'first_line': PropertyDescription('w:pPr', 'w:ind', 'w:firstLine', True),
        'keep_lines': PropertyDescription('w:pPr', 'w:keepLines', None, True),
        'keep_next': PropertyDescription('w:pPr', 'w:keepNext', None, True),
        'outline_level': PropertyDescription('w:pPr', 'w:outlineLvl', 'w:val', True),
        'border_top': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:val', True),
        'border_top_color': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:color', True),
        'border_top_size': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:sz', True),
        'border_top_space': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:space', True),
        'border_bottom': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:val', True),
        'border_bottom_color': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:color', True),
        'border_bottom_size': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:sz', True),
        'border_bottom_space': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:space', True),
        'border_right': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:val', True),
        'border_right_color': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:color', True),
        'border_right_size': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:sz', True),
        'border_right_space': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:space', True),
        'border_left': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:val', True),
        'border_left_color': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:color', True),
        'border_left_size': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:sz', True),
        'border_left_space': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:space', True),
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
        return self._properties['align']

    def set_align_value(self, align: Union[str, None]):
        """
        :param align: 'both' or 'right' or 'center' or 'left
        """
        self._properties['align'].value = align

    def get_indent_left(self) -> Property:
        return self._properties['indent_left']

    def set_indent_left_value(self, indent_left: Union[str, None]):
        self._properties['indent_left'].value = indent_left

    def get_indent_right(self) -> Property:
        return self._properties['indent_right']

    def set_indent_right_value(self, indent_right: Union[str, None]):
        self._properties['indent_right'].value = indent_right

    def get_hanging(self) -> Property:
        return self._properties['hanging']

    def set_hanging_value(self, hanging: Union[str, None]):
        self._properties['hanging'].value = hanging

    def get_first_line(self) -> Property:
        return self._properties['first_line']

    def set_first_line_value(self, first_line: Union[str, None]):
        self._properties['first_line'].value = first_line

    def is_keep_lines(self) -> bool:
        return self._properties['keep_lines'].value

    def set_as_keep_lines(self, is_keep_lines: bool = True):
        self._properties['keep_lines'].value = is_keep_lines

    def is_keep_next(self) -> bool:
        return self._properties['keep_next'].value

    def set_as_keep_next(self, is_keep_next: bool = True):
        self._properties['keep_next'].value = is_keep_next

    def get_outline_level(self) -> Property:
        return self._properties['outline_level']

    def set_outline_level_value(self, level: Union[str, int, None]):
        """
        :param level: 0-9
        """
        self._properties['outline_level'].value = level

    def get_border(self, direction: str, property_name: str) -> Property:
        """
        :param direction: top, bottom, right or left
        :param property_name: color, size, space or type
        """
        property_name = '' if property_name == 'type' else ('_' + property_name)
        return self._properties[rf'border_{direction}{property_name}']

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right or left
        :param property_name: color, size, space or type
        """
        property_name = '' if property_name == 'type' else ('_' + property_name)
        self._properties[rf'border_{direction}{property_name}'].value = value

    class Run(XMLement):
        tag: str = 'w:r'
        _all_properties = {
            'size': PropertyDescription('w:rPr', 'w:sz', 'w:val', True),
            'is_bold': PropertyDescription('w:rPr', 'w:b', None, True),
            'is_italic': PropertyDescription('w:rPr', 'w:i', None, True),
            'vertical_align': PropertyDescription('w:rPr', 'w:vertAlign', 'w:val', True),
            'language': PropertyDescription('w:rPr', 'w:lang', 'w:val', True),
            'color': PropertyDescription('w:rPr', 'w:color', 'w:val', True),
            'theme_color': PropertyDescription('w:rPr', 'w:color', 'w:themeColor', True),
            'background_color': PropertyDescription('w:rPr', 'w:highlight', 'w:val', True),
            'background_fill': PropertyDescription('w:rPr', 'w:shd', 'w:fill', True),
            'underline': PropertyDescription('w:rPr', 'w:u', 'w:val', True),
            'underline_color': PropertyDescription('w:rPr', 'w:u', 'w:color', True),
            'is_strike': PropertyDescription('w:rPr', 'w:strike', None, True),
            'border': PropertyDescription('w:rPr', 'w:bdr', 'w:val', True),
            'border_color': PropertyDescription('w:rPr', 'w:bdr', 'w:color', True),
            'border_size': PropertyDescription('w:rPr', 'w:bdr', 'w:sz', True),
            'border_space': PropertyDescription('w:rPr', 'w:bdr', 'w:space', True),
        }
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
            return self._properties['size']

        def set_size_value(self, value: Union[str, None]):
            self._properties['size'].value = value

        def is_bold(self) -> bool:
            return self._properties['is_bold'].value

        def set_as_bold(self, is_bold: bool = True):
            self._properties['is_bold'].value = is_bold

        def is_italic(self) -> bool:
            return self._properties['is_italic'].value

        def set_as_italic(self, is_italic: bool = True):
            self._properties['is_italic'].value = is_italic

        def is_strike(self) -> bool:
            return self._properties['is_strike'].value

        def set_as_strike(self, is_strike: bool = True):
            self._properties['is_strike'].value = is_strike

        def get_language(self) -> Property:
            return self._properties['language']

        def set_language_value(self, value: Union[str, None]):
            self._properties['language'].value = value

        def get_vertical_align(self) -> Property:
            return self._properties['vertical_align']

        def set_vertical_align_value(self, value: Union[str, None]):
            """
            param value: 'superscript' or 'subscript'
            """
            self._properties['vertical_align'].value = value

        def get_color(self) -> Property:
            return self._properties['color']

        def set_color_value(self, value: Union[str, None]):
            """
            param value: specifies the color as a hex value in RRGGBB format or auto
            """
            self._properties['color'].value = value

        def get_theme_color(self) -> Property:
            return self._properties['theme_color']

        def set_theme_color_value(self, value: Union[str, None]):
            self._properties['theme_color'].value = value

        def get_background_color(self) -> Property:
            return self._properties['background_color']

        def set_background_color_value(self, value: Union[str, None]):
            """
            param value: possible values: black, blue, cyan, darkBlue, darkCyan, darkGray, darkGreen, darkMagenta,
                                        darkRed, darkYellow, green, lightGray, magenta, none, red, white, yellow
            """
            self._properties['background_color'].value = value

        def get_background_fill(self) -> Property:
            return self._properties['background_fill']

        def set_background_fill_value(self, value: Union[str, None]):
            """
            param value: specifies the color as a hex value in RRGGBB format or auto
            """
            self._properties['background_fill'].value = value

        def get_underline(self) -> Property:
            return self._properties['underline']

        def set_underline_value(self, value: Union[str, None]):
            self._properties['underline'].value = value

        def get_underline_color(self) -> Property:
            return self._properties['underline_color']

        def set_underline_color_value(self, value: Union[str, None]):
            """
            param value: specifies the color as a hex value in RRGGBB format or auto
            """
            self._properties['underline_color'].value = value

        def get_border(self, property_name: str) -> Property:
            """
            :param direction: top, bottom, right or left
            :param property_name: color, size, space or type
            """
            property_name = '' if property_name == 'type' else ('_' + property_name)
            return self._properties[rf'border{property_name}']

        def set_border_value(self, property_name: str, value: Union[str, None]):
            """
            :param direction: top, bottom, right or left
            :param property_name: color, size, space or type
            """
            property_name = '' if property_name == 'type' else ('_' + property_name)
            self._properties[rf'border_{property_name}'].value = value

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
    _all_properties = {
        'layout': PropertyDescription('w:tblPr', 'w:tblLayout', 'w:type', True),
        'width': PropertyDescription('w:tblPr', 'w:tblW', 'w:w', True),
        'width_type': PropertyDescription('w:tblPr', 'w:tblW', 'w:type', True),
        'align': PropertyDescription('w:tblPr', 'w:jc', 'w:val', True),
        'borders_inside_horizontal': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:val', True),
        'borders_inside_horizontal_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:color', True),
        'borders_inside_horizontal_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:sz', True),
        'borders_inside_vertical': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:val', True),
        'borders_inside_vertical_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:color', True),
        'borders_inside_vertical_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:sz', True),
        'border_top': PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:val', True),
        'border_top_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:color', True),
        'border_top_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:sz', True),
        'border_bottom': PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:val', True),
        'border_bottom_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:color', True),
        'border_bottom_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:sz', True),
        'border_right': PropertyDescription('w:tblPr/w:tblBorders', 'w:right', 'w:val', True),
        'border_right_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:right', 'w:color', True),
        'border_right_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:right', 'w:sz', True),
        'border_left': PropertyDescription('w:tblPr/w:tblBorders', 'w:left', 'w:val', True),
        'border_left_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:left', 'w:color', True),
        'border_left_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:left', 'w:sz', True),
        'cell_margin_top': PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:w', True),
        'cell_margin_top_type': PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:type', True),
        'cell_margin_left': PropertyDescription('w:tblPr/w:tblCellMar', 'w:left', 'w:w', True),
        'cell_margin_left_type': PropertyDescription('w:tblPr/w:tblCellMar', 'w:left', 'w:type', True),
        'cell_margin_bottom': PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:w', True),
        'cell_margin_bottom_type': PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:type', True),
        'cell_margin_right': PropertyDescription('w:tblPr/w:tblCellMar', 'w:right', 'w:w', True),
        'cell_margin_right_type': PropertyDescription('w:tblPr/w:tblCellMar', 'w:right', 'w:type', True),
        'indentation': PropertyDescription('w:tblPr', 'w:tblInd', 'w:w', True),
        'indentation_type': PropertyDescription('w:tblPr', 'w:tblInd', 'w:type', True),
    }
    '''
    all_style_properties = {
        'borders_inside_horizontal': ('w:tblPr/w:tblBorders/w:insideH', 'w:val', True),
        'borders_inside_horizontal_color': ('w:tblPr/w:tblBorders/w:insideH', 'w:color', True),
        'borders_inside_horizontal_size': ('w:tblPr/w:tblBorders/w:insideH', 'w:size', True),
        'borders_inside_vertical': ('w:tblPr/w:tblBorders/w:insideH', 'w:val', True),
        'borders_inside_vertical_color': ('w:tblPr/w:tblBorders/w:insideH', 'w:color', True),
        'borders_inside_vertical_size': ('w:tblPr/w:tblBorders/w:insideH', 'w:size', True),
        'border_top': ('w:tblPr/w:tblBorders/w:top', 'w:val', True),
        'border_top_color': ('w:tblPr/w:tblBorders/w:top', 'w:color', True),
        'border_top_size': ('w:tblPr/w:tblBorders/w:top', 'w:sz', True),
        'border_bottom': ('w:tblPr/w:tblBorders/w:bottom', 'w:val', True),
        'border_bottom_color': ('w:tblPr/w:tblBorders/w:bottom', 'w:color', True),
        'border_bottom_size': ('w:tblPr/w:tblBorders/w:bottom', 'w:sz', True),
        'border_right': ('w:tblPr/w:tblBorders/w:right', 'w:val', True),
        'border_right_color': ('w:tblPr/w:tblBorders/w:right', 'w:color', True),
        'border_right_size': ('w:tblPr/w:tblBorders/w:right', 'w:sz', True),
        'border_left': ('w:tblPr/w:tblBorders/w:left', 'w:val', True),
        'border_left_color': ('w:tblPr/w:tblBorders/w:left', 'w:color', True),
        'border_left_size': ('w:tblPr/w:tblBorders/w:left', 'w:sz', True),
        'cell_margin_top': ('w:tblPr/w:tblCellMar/w:top', 'w:w', True),
        'cell_margin_top_type': ('w:tblPr/w:tblCellMar/w:top', 'w:type', True),
        'cell_margin_left': ('w:tblPr/w:tblCellMar/w:left', 'w:w', True),
        'cell_margin_left_type': ('w:tblPr/w:tblCellMar/w:left', 'w:type', True),
        'cell_margin_bottom': ('w:tblPr/w:tblCellMar/w:bottom', 'w:w', True),
        'cell_margin_bottom_type': ('w:tblPr/w:tblCellMar/w:bottom', 'w:type', True),
        'cell_margin_right': ('w:tblPr/w:tblCellMar/w:right', 'w:w', True),
        'cell_margin_right_type': ('w:tblPr/w:tblCellMar/w:right', 'w:type', True),
    }
    '''
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
                if row.cells[i].get_property_value('vertical_merge') == 'restart':
                    cells[j] = row.cells[i]
                    cells[j].row_span = 1
                elif row.cells[i].get_property_value('vertical_merge') == 'continue':
                    cells[j].row_span += 1
                    cells_for_delete.append(row.cells[i])
            for cell in cells_for_delete:
                row.cells.remove(cell)

    def __set_margins_for_cells(self):
        directions: Tuple[str, str, str, str] = ('top', 'bottom', 'left', 'right')
        for direction in directions:
            if self._is_have_viewing_property(rf'cell_margin_{direction}'):
                for row in self.rows:
                    for cell in row.cells:
                        if cell.get_property_value(rf'margin_{direction}') is None:
                            cell.set_property_value(rf'margin_{direction}',
                                                    self._properties[rf'cell_margin_{direction}'].value)
                            cell.set_property_value(rf'margin_{direction}_type',
                                                    self._properties[rf'cell_margin_{direction}_type'].value)

    def __set_inside_borders(self):                                              # TO DO: testing this method
        directions: Tuple[str, str, str, str] = ('top', 'bottom', 'left', 'right')
        for direction in directions:
            d: str = 'horizontal' if (direction == 'top' or direction == 'bottom') else 'vertical'
            if self._is_have_viewing_property(rf'borders_inside_{d}'):
                for i in range(len(self.rows)):
                    for j in range(len(self.rows[i].cells)):
                        if not (i == 0 and direction == 'top') and \
                                not (i == (len(self.rows) - 1) and direction == 'bottom') and \
                                not (j == 0 and direction == 'left') and \
                                not (j == (len(self.rows[i].cells) - 1) and direction == 'right'):
                            if self.rows[i].cells[j].get_property_value(rf'border_{direction}_color') is None or \
                                    self.rows[i].cells[j].get_property_value(rf'border_{direction}_color') == 'auto':
                                self.rows[i].cells[j].set_property_value(
                                    rf'border_{direction}_color',
                                    self._properties[rf'borders_inside_{d}_color'].value
                                )
                            if self.rows[i].cells[j].get_property_value(rf'border_{direction}') is None:
                                self.rows[i].cells[j].set_property_value(
                                    rf'border_{direction}',
                                    self._properties[rf'borders_inside_{d}'].value
                                )
                                self.rows[i].cells[j].set_property_value(
                                    rf'border_{direction}_size',
                                    self._properties[rf'borders_inside_{d}_size'].value
                                )

    def get_width(self) -> Tuple[Property, Property, Property]:
        """
        return: Tuple(width, type, layout)
        """
        return self._properties['width'], self._properties['width_type'], self._properties['layout']

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None], layout: str = 'autofit'):
        self._properties['width'].value = width
        self._properties['width_type'].value = width_type
        self._properties['layout'].value = layout

    def get_align(self) -> Property:
        return self._properties['align']

    def set_align_value(self, value: Union[str, None]):
        self._properties['align'].value = value

    def get_border(self, direction: str, property_name: str) -> Property:
        """
        :param direction: top, bottom, right or left
        :param property_name: color, size, or type
        """
        property_name = '' if property_name == 'type' else ('_' + property_name)
        return self._properties[rf'border_{direction}{property_name}']

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right or left
        :param property_name: color, size, or type
        """
        property_name = '' if property_name == 'type' else ('_' + property_name)
        self._properties[rf'border_{direction}{property_name}'].value = value

    def get_indentation(self) -> Tuple[Property, Property]:
        """
        return: Tuple(value, type)
        """
        return self._properties['indentation'], self._properties['indentation_type']

    def set_indentation_value(self, width: Union[str, None], width_type: Union[str, None]):
        self._properties['indentation'].value = width
        self._properties['indentation_type'].value = width_type

    def get_inside_border(self, direction: str, property_name: str) -> Property:
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color' or 'size' or 'type'
        """
        property_name = '' if property_name == 'type' else ('_' + property_name)
        return self._properties[rf'borders_inside_{direction}{property_name}']

    def set_inside_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color' or 'size' or 'type'
        """
        property_name = '' if property_name == 'type' else ('_' + property_name)
        self._properties[rf'borders_inside_{direction}{property_name}'].value = value

    def get_cells_margin(self, direction: str) -> Tuple[Property, Property]:
        """
        :param direction: 'top' or 'bottom' or 'right' or 'left'
        :result: value, value_type
        """
        return self._properties[rf'cell_margin_{direction}'], self._properties[rf'cell_margin_{direction}_type']

    def set_cells_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction:  'top' or 'bottom' or 'right' or 'left'
        :param value_type: 'dxa' or 'nil'
        """
        self._properties[rf'margin_{direction}'].value = value
        self._properties[rf'margin_{direction}_type'].value = value_type

    class Row(XMLement):
        tag: str = 'w:tr'
        _all_properties = {
            'is_header': PropertyDescription('w:trPr', 'w:tblHeader', None, True),
            'height': PropertyDescription('w:trPr', 'w:trHeight', 'w:val', True),
            'height_rule': PropertyDescription('w:trPr', 'w:trHeight', 'w:hRule', True),
        }
        from translators import RowTranslatorToHTML
        translators = {
            'html': RowTranslatorToHTML(),
        }

        def __init__(self, element: ET.Element, parent):
            self.cells: List[Table.Row.Cell] = []
            self.is_first_row_in_header: bool = False
            self.is_last_row_in_header: bool = False
            super(Table.Row, self).__init__(element, parent)
            if self._properties['is_header'].value:
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
            return self._properties['is_header'].value

        def set_as_header(self, is_header: bool = True):
            self._properties['is_header'].value = is_header
            self.__set_cells_as_header()

        def get_height(self) -> Tuple[Property, Property]:
            """
            return: Tuple(height, rule)
            """
            return self._properties['height'], self._properties['height_rule']

        def set_height_value(self, value: Union[str, None], height_rule: str = 'atLeast'):
            self._properties['height'].value = value
            self._properties['height_rule'].value = height_rule

        class Cell(XMLcontainer):
            tag: str = 'w:tc'
            _all_properties = {
                'fill_color': PropertyDescription('w:tcPr', 'w:shd', 'w:fill', True),
                'fill_theme': PropertyDescription('w:tcPr', 'w:shd', 'w:themeFill', True),
                'border_top': PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:val', True),
                'border_top_color': PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:color', True),
                'border_top_size': PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:sz', True),
                'border_bottom': PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:val', True),
                'border_bottom_color': PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:color', True),
                'border_bottom_size': PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:sz', True),
                'border_right': PropertyDescription('w:tcPr/w:tcBorders', 'w:right', 'w:val', True),
                'border_right_color': PropertyDescription('w:tcPr/w:tcBorders', 'w:right', 'w:color', True),
                'border_right_size': PropertyDescription('w:tcPr/w:tcBorders', 'w:right', 'w:sz', True),
                'border_left': PropertyDescription('w:tcPr/w:tcBorders', 'w:left', 'w:val', True),
                'border_left_color': PropertyDescription('w:tcPr/w:tcBorders', 'w:left', 'w:color', True),
                'border_left_size': PropertyDescription('w:tcPr/w:tcBorders', 'w:left', 'w:sz', True),
                'width': PropertyDescription('w:tcPr', 'w:tcW', 'w:w', True),
                'width_type': PropertyDescription('w:tcPr', 'w:tcW', 'w:type', True),
                'col_span': PropertyDescription('w:tcPr', 'w:gridSpan', 'w:val', True),
                'vertical_merge': PropertyDescription('w:tcPr', 'w:vMerge', 'w:val', True),
                'is_vertical_merge_continue': PropertyDescription('w:tcPr', 'w:vMerge', None, True),
                'vertical_align': PropertyDescription('w:tcPr', 'w:vAlign', 'w:val', True),
                'text_direction': PropertyDescription('w:tcPr', 'w:textDirection', 'w:val', True),
                'margin_top': PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:w', True),
                'margin_top_type': PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:type', True),
                'margin_bottom': PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:w', True),
                'margin_bottom_type': PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:type', True),
                'margin_left': PropertyDescription('w:tcPr/w:tcMar', 'w:left', 'w:w', True),
                'margin_left_type': PropertyDescription('w:tcPr/w:tcMar', 'w:left', 'w:type', True),
                'margin_right': PropertyDescription('w:tcPr/w:tcMar', 'w:right', 'w:w', True),
                'margin_right_type': PropertyDescription('w:tcPr/w:tcMar', 'w:right', 'w:type', True)
            }
            from translators import CellTranslatorToHTML
            translators = {
                'html': CellTranslatorToHTML(),
            }

            def __init__(self, element: ET.Element, parent):
                self.row_span: int = 1
                self.is_header = False
                super(Table.Row.Cell, self).__init__(element, parent)

            def get_fill_color(self) -> Property:
                return self._properties['fill_color']

            def set_fill_color_value(self, value: Union[str, None]):
                self._properties['fill_color'].value = value

            def get_fill_theme(self) -> Property:
                return self._properties['fill_theme']

            def set_fill_theme_value(self, value: Union[str, None]):
                self._properties['fill_theme'].value = value

            def get_border(self, direction: str, property_name: str) -> Property:
                """
                :param direction: top, bottom, right or left
                :param property_name: color, size, or type
                """
                property_name = '' if property_name == 'type' else ('_' + property_name)
                return self._properties[rf'border_{direction}{property_name}']

            def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
                """
                :param direction: top, bottom, right or left
                :param property_name: color, size, or type
                """
                property_name = '' if property_name == 'type' else ('_' + property_name)
                self._properties[rf'border_{direction}{property_name}'].value = value

            def get_width(self) -> Tuple[Property, Property]:
                """
                return: Tuple(width, type)
                """
                return self._properties['width'], self._properties['width_type']

            def set_width_value(self, width: Union[str, None], width_type: Union[str, None]):
                self._properties['width'].value = width
                self._properties['width_type'].value = width_type

            def get_col_span(self) -> Property:
                return self._properties['col_span']

            def set_col_span_value(self, value: int):
                self._properties['col_span'].value = value

            def get_vertical_align(self) -> Property:
                return self._properties['vertical_align']

            def set_vertical_align_value(self, value: str):
                self._properties['vertical_align'].value = value

            def get_text_direction(self) -> Property:
                return self._properties['text_direction']

            def set_text_direction_value(self, value: Union[str, None]):
                self._properties['text_direction'].value = value

            def get_margin(self, direction: str) -> Tuple[Property, Property]:
                """
                :param direction:  'top' or 'bottom' or 'right' or 'left'
                :return: (margin, type)
                """
                return self._properties[rf'margin_{direction}'], self._properties[rf'margin_{direction}_type']

            def set_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
                """
                :param direction:  'top' or 'bottom' or 'right' or 'left'
                :param value_type: 'dxa' or 'nil'
                """
                self._properties[rf'margin_{direction}'].value = value
                self._properties[rf'margin_{direction}_type'].value = value_type


class Document(XMLement):

    class Body(XMLcontainer):
        tag: str = 'w:body'
        _is_unique = True

        def __str__(self):
            return self.get_inner_text()

    def __init__(self, path: str):
        self.__path: str = path
        self.body: Document.Body
        super(Document, self).__init__(Parser.get_xml_file(self, 'document'), None)
        # self._parse_styles(self._get_xml_file('styles'), [Table])

    def _init(self):
        self.body: Document.Body = self._get_elements(Document.Body)

    def __str__(self):
        return str(self.body)

    def get_inner_text(self) -> Union[str, None]:
        return str(self.body)

    def get_path(self) -> str:
        return self.__path
