from constants import keys_consts as k_const
from typing import Tuple, Union
from abc import ABC


class GetSetMixin(ABC):
    def get_property(self, property_name: str) -> Union[str, None, bool]:
        raise NotImplementedError('method get_property is not implemented. Your class must implement this method or'
                                  'GetSetMixin must be inherited after the class that implements this method')

    def set_property_value(self, property_name: str, value: Union[str, bool, None]):
        raise NotImplementedError('method set_property_value is not implemented. '
                                  'Your class must implement this method or'
                                  'GetSetMixin must be inherited after the class that implements this method')


class ParagraphPropertiesGetSetMixin(GetSetMixin, ABC):
    def get_align(self) -> Union[str, None]:
        return self.get_property(k_const.PARAGRAPH_ALIGN)

    def set_align_value(self, align: Union[str, None]):
        """
        :param align: 'both' or 'right' or 'center' or 'left
        """
        self.set_property_value(k_const.PARAGRAPH_ALIGN, align)

    def get_indent_left(self) -> Union[str, None]:
        return self.get_property(k_const.PARAGRAPH_INDENT_LEFT)

    def set_indent_left_value(self, indent_left: Union[str, None]):
        self.set_property_value(k_const.PARAGRAPH_INDENT_LEFT, indent_left)

    def get_indent_right(self) -> Union[str, None]:
        return self.get_property(k_const.PARAGRAPH_INDENT_RIGHT)

    def set_indent_right_value(self, indent_right: Union[str, None]):
        self.set_property_value(k_const.PARAGRAPH_INDENT_RIGHT, indent_right)

    def get_hanging(self) -> Union[str, None]:
        return self.get_property(k_const.PARAGRAPH_HANGING)

    def set_hanging_value(self, hanging: Union[str, None]):
        self.set_property_value(k_const.PARAGRAPH_HANGING, hanging)

    def get_first_line(self) -> Union[str, None]:
        return self.get_property(k_const.PARAGRAPH_FIRST_LINE)

    def set_first_line_value(self, first_line: Union[str, None]):
        self.set_property_value(k_const.PARAGRAPH_FIRST_LINE, first_line)

    def is_keep_lines(self) -> bool:
        result = self.get_property(k_const.PARAGRAPH_KEEP_LINES)
        return result if result is not None else False

    def set_as_keep_lines(self, is_keep_lines: Union[bool, None] = True):
        self.set_property_value(k_const.PARAGRAPH_KEEP_LINES, is_keep_lines)

    def is_keep_next(self) -> bool:
        result = self.get_property(k_const.PARAGRAPH_KEEP_NEXT)
        return result if result is not None else False

    def set_as_keep_next(self, is_keep_next: Union[bool, None] = True):
        self.set_property_value(k_const.PARAGRAPH_KEEP_NEXT, is_keep_next)

    def get_outline_level(self) -> Union[str, None]:
        return self.get_property(k_const.PARAGRAPH_OUTLINE_LEVEL)

    def set_outline_level_value(self, level: Union[str, int, None]):
        """
        :param level: 0-9
        """
        self.set_property_value(k_const.PARAGRAPH_OUTLINE_LEVEL, level)

    def get_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left     (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, space, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(k_const.get_property_key(
            k_const.Element.PARAGRAPH, k_const.SubElement.BORDER, direction=direction, property_name=property_name)
        )

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right, left     (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, space, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(k_const.get_property_key(
            k_const.Element.PARAGRAPH, k_const.SubElement.BORDER, direction=direction, property_name=property_name),
            value
        )


class RunPropertiesGetSetMixin(GetSetMixin, ABC):
    def get_size(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_SIZE)

    def set_size_value(self, value: Union[str, None]):
        self.set_property_value(k_const.RUN_SIZE, value)

    def is_bold(self) -> bool:
        result = self.get_property(k_const.RUN_IS_BOLD)
        return result if result is not None else False

    def set_as_bold(self, is_bold: Union[bool, None] = True):
        self.set_property_value(k_const.RUN_IS_BOLD, is_bold)

    def is_italic(self) -> bool:
        result = self.get_property(k_const.RUN_IS_ITALIC)
        return result if result is not None else False

    def set_as_italic(self, is_italic: Union[bool, None] = True):
        self.set_property_value(k_const.RUN_IS_ITALIC, is_italic)

    def is_strike(self) -> bool:
        result = self.get_property(k_const.RUN_IS_STRIKE)
        return result if result is not None else False

    def set_as_strike(self, is_strike: Union[bool, None] = True):
        self.set_property_value(k_const.RUN_IS_STRIKE, is_strike)

    def get_language(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_LANGUAGE)

    def set_language_value(self, value: Union[str, None]):
        self.set_property_value(k_const.RUN_LANGUAGE, value)

    def get_vertical_align(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_VERTICAL_ALIGN)

    def set_vertical_align_value(self, value: Union[str, None]):
        """
        param value: 'superscript' or 'subscript'
        """
        self.set_property_value(k_const.RUN_VERTICAL_ALIGN, value)

    def get_color(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_COLOR)

    def set_color_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(k_const.RUN_COLOR, value)

    def get_theme_color(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_THEME_COLOR)

    def set_theme_color_value(self, value: Union[str, None]):
        self.set_property_value(k_const.RUN_THEME_COLOR, value)

    def get_background_color(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_BACKGROUND_COLOR)

    def set_background_color_value(self, value: Union[str, None]):
        """
        param value: possible values: black, blue, cyan, darkBlue, darkCyan, darkGray, darkGreen, darkMagenta,
                                    darkRed, darkYellow, green, lightGray, magenta, none, red, white, yellow
        """
        self.set_property_value(k_const.RUN_BACKGROUND_COLOR, value)

    def get_background_fill(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_BACKGROUND_FILL)

    def set_background_fill_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(k_const.RUN_BACKGROUND_FILL, value)

    def get_underline(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_UNDERLINE_TYPE)

    def set_underline_value(self, value: Union[str, None]):
        self.set_property_value(k_const.RUN_UNDERLINE_TYPE, value)

    def get_underline_color(self) -> Union[str, None]:
        return self.get_property(k_const.RUN_UNDERLINE_COLOR)

    def set_underline_color_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(k_const.RUN_UNDERLINE_COLOR, value)

    def get_border(self, property_name: str) -> Union[str, None]:
        """
        :param property_name: color, size, space, type
        (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(k_const.get_property_key(
            k_const.Element.RUN, k_const.SubElement.BORDER, property_name=property_name)
        )

    def set_border_value(self, property_name: str, value: Union[str, None]):
        """
        :param property_name: color, size, space or type
        (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(k_const.get_property_key(
            k_const.Element.RUN, k_const.SubElement.BORDER, property_name=property_name), value
        )


class TablePropertiesGetSetMixin(GetSetMixin, ABC):
    def get_width(self) -> Tuple[Union[str, None], Union[str, None], Union[str, None]]:
        """
        return: Tuple(width, type, layout)
        """
        return (
            self.get_property(k_const.TABLE_WIDTH),
            self.get_property(k_const.TABLE_WIDTH_TYPE),
            self.get_property(k_const.TABLE_LAYOUT)
        )

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None], layout: str = 'autofit'):
        self.set_property_value(k_const.TABLE_WIDTH, width)
        self.set_property_value(k_const.TABLE_WIDTH_TYPE, width_type)
        self.set_property_value(k_const.TABLE_LAYOUT, layout)

    def get_align(self) -> Union[str, None]:
        return self.get_property(k_const.TABLE_ALIGN)

    def set_align_value(self, value: Union[str, None]):
        self.set_property_value(k_const.TABLE_ALIGN, value)

    def get_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left      (keys of XMLementPropertyDescriptions.Const_property_names dict)
        :param property_name: color, size, type         (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(k_const.get_property_key(
            k_const.Element.TABLE, k_const.SubElement.BORDER, direction=direction, property_name=property_name)
        )

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right, left      (keys of XMLementPropertyDescriptions.Const_property_names dict)
        :param property_name: color, size, type         (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(k_const.get_property_key(
            k_const.Element.TABLE, k_const.SubElement.BORDER, direction=direction, property_name=property_name), value
        )

    def get_indentation(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(value, type)
        """
        return self.get_property(k_const.TABLE_INDENTATION), self.get_property(k_const.TABLE_INDENTATION_TYPE)

    def set_indentation_value(self, width: Union[str, None], width_type: Union[str, None]):
        self.set_property_value(k_const.TABLE_INDENTATION, width)
        self.set_property_value(k_const.TABLE_INDENTATION_TYPE, width_type)

    def get_inside_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color', 'size', 'type'  (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(
            k_const.get_property_key(k_const.Element.TABLE,
                                     subelements=[k_const.Element.CELL, k_const.SubElement.BORDER],
                                     direction=direction, property_name=property_name)
        )

    def set_inside_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color', 'size', 'type'  (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(
            k_const.get_property_key(k_const.Element.TABLE,
                                     subelements=[k_const.Element.CELL, k_const.SubElement.BORDER],
                                     direction=direction, property_name=property_name),
            value
        )

    def get_cells_margin(self, direction: str) -> Tuple[Union[str, None], Union[str, None]]:
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :result: value, value_type
        """
        return (
            self.get_property(k_const.get_property_key(
                k_const.Element.TABLE, subelements=[k_const.Element.CELL, k_const.SubElement.MARGIN],
                direction=direction, property_name=k_const.PropertyName.SIZE)
            ),
            self.get_property(k_const.get_property_key(
                k_const.Element.TABLE, subelements=[k_const.Element.CELL, k_const.SubElement.MARGIN],
                direction=direction, property_name=k_const.PropertyName.TYPE)
            )
        )

    def set_cells_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param value_type: 'dxa' or 'nil'
        """
        self.set_property_value(k_const.get_property_key(k_const.Element.TABLE,
                                                         subelements=[k_const.Element.CELL, k_const.SubElement.MARGIN],
                                                         direction=direction, property_name=k_const.PropertyName.SIZE),
                                value)
        self.set_property_value(k_const.get_property_key(k_const.Element.TABLE,
                                                         subelements=[k_const.Element.CELL, k_const.SubElement.MARGIN],
                                                         direction=direction, property_name=k_const.PropertyName.TYPE),
                                value_type)


class RowPropertiesGetSetMixin(GetSetMixin, ABC):
    def is_header(self) -> bool:
        result = self.get_property(k_const.ROW_IS_HEADER)
        return result if result is not None else False

    def set_as_header(self, is_header: Union[bool, None] = True):
        self.set_property_value(k_const.ROW_IS_HEADER, is_header)

    def get_height(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(height, rule)
        """
        return self.get_property(k_const.ROW_HEIGHT), self.get_property(k_const.ROW_HEIGHT_RULE)

    def set_height_value(self, value: Union[str, None], height_rule: str = 'atLeast'):
        self.set_property_value(k_const.ROW_HEIGHT, value)
        self.set_property_value(k_const.ROW_HEIGHT_RULE, height_rule)


class CellPropertiesGetSetMixin(GetSetMixin, ABC):
    def get_parent_table(self):
        raise NotImplementedError('method get_parent_table is not implemented. Your class must implement this method or'
                                  'CellPropertiesGetSetMixin must be inherited after the class that implements this '
                                  'method')

    def is_top(self) -> bool:
        raise NotImplementedError('method is_top is not implemented. Your class must implement this method or'
                                  'CellPropertiesGetSetMixin must be inherited after the class that implements this '
                                  'method')

    def is_bottom(self) -> bool:
        raise NotImplementedError('method is_bottom is not implemented. Your class must implement this method or'
                                  'CellPropertiesGetSetMixin must be inherited after the class that implements this '
                                  'method')

    def is_first_in_row(self) -> bool:
        raise NotImplementedError('method is_first_in_row is not implemented. Your class must implement this method or'
                                  'CellPropertiesGetSetMixin must be inherited after the class that implements this '
                                  'method')

    def is_last_in_row(self) -> bool:
        raise NotImplementedError('method is_last_in_row is not implemented. Your class must implement this method or'
                                  'CellPropertiesGetSetMixin must be inherited after the class that implements this '
                                  'method')

    def get_fill_color(self) -> Union[str, None]:
        return self.get_property(k_const.CELL_FILL_COLOR)

    def set_fill_color_value(self, value: Union[str, None]):
        self.set_property_value(k_const.CELL_FILL_COLOR, value)

    def get_fill_theme(self) -> Union[str, None]:
        return self.get_property(k_const.CELL_FILL_THEME)

    def set_fill_theme_value(self, value: Union[str, None]):
        self.set_property_value(k_const.CELL_FILL_THEME, value)

    def get_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        result = self.get_property(k_const.get_property_key(k_const.Element.CELL, k_const.SubElement.BORDER,
                                                            direction=direction, property_name=property_name)
                                   )
        if result is None or (k_const.PropertyName.COLOR.is_equal(property_name) and result == 'auto'):
            if not (self.is_first_in_row() and k_const.Direction.LEFT.is_equal(direction)) and \
               not (self.is_last_in_row() and k_const.Direction.RIGHT.is_equal(direction)) and \
               not (self.is_top() and k_const.Direction.TOP.is_equal(direction)) and \
               not (self.is_bottom() and k_const.Direction.BOTTOM.is_equal(direction)):
                d = k_const.Direction.horizontal_or_vertical_straight(direction)
                return self.get_parent_table().get_inside_border(d, property_name)
        return result

    def get_width(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(width, type)
        """
        return self.get_property(k_const.CELL_WIDTH), self.get_property(k_const.CELL_WIDTH_TYPE)

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None]):
        self.set_property_value(k_const.CELL_WIDTH, width)
        self.set_property_value(k_const.CELL_WIDTH_TYPE, width_type)

    def get_vertical_align(self) -> Union[str, None]:
        return self.get_property(k_const.CELL_VERTICAL_ALIGN)

    def set_vertical_align_value(self, value: str):
        self.set_property_value(k_const.CELL_VERTICAL_ALIGN, value)

    def get_text_direction(self) -> Union[str, None]:
        return self.get_property(k_const.CELL_TEXT_DIRECTION)

    def set_text_direction_value(self, value: Union[str, None]):
        self.set_property_value(k_const.CELL_TEXT_DIRECTION, value)

    def get_margin(self, direction: str) -> Tuple[Union[str, None], Union[str, None]]:
        """
        :param direction: 'top', 'bottom', 'right', 'left'
        (keys of XMLementPropertyDescriptions.Const_directions dict)
        :return: (margin, type)
        """
        if self.get_property(k_const.get_property_key(k_const.Element.CELL, k_const.SubElement.MARGIN,
                                                      direction=direction, property_name=k_const.PropertyName.SIZE)
                             ) is None:
            return (
                self.get_parent_table().get_property(
                    k_const.get_property_key(k_const.Element.TABLE,
                                             subelements=[k_const.Element.CELL, k_const.SubElement.MARGIN],
                                             direction=direction, property_name=k_const.PropertyName.SIZE)
                ),
                self.get_parent_table().get_property(
                    k_const.get_property_key(k_const.Element.TABLE,
                                             subelements=[k_const.Element.CELL, k_const.SubElement.MARGIN],
                                             direction=direction, property_name=k_const.PropertyName.TYPE)
                )
            )
        return (
            self.get_property(k_const.get_property_key(k_const.Element.CELL, k_const.SubElement.MARGIN,
                                                       direction=direction, property_name=k_const.PropertyName.SIZE)),
            self.get_property(k_const.get_property_key(k_const.Element.CELL, k_const.SubElement.MARGIN,
                                                       direction=direction, property_name=k_const.PropertyName.TYPE))
        )

    def set_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction:  'top' or 'bottom' or 'right' or 'left'
        (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param value_type: 'dxa' or 'nil'
        """
        self.set_property_value(k_const.get_property_key(k_const.Element.CELL, k_const.SubElement.MARGIN,
                                                         direction=direction, property_name=k_const.PropertyName.SIZE),
                                value)
        self.set_property_value(k_const.get_property_key(k_const.Element.CELL, k_const.SubElement.MARGIN,
                                                         direction=direction, property_name=k_const.PropertyName.TYPE),
                                value_type)
