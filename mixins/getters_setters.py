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
        return self.get_property(k_const.Par_align)

    def set_align_value(self, align: Union[str, None]):
        """
        :param align: 'both' or 'right' or 'center' or 'left
        """
        self.set_property_value(k_const.Par_align, align)

    def get_indent_left(self) -> Union[str, None]:
        return self.get_property(k_const.Par_indent_left)

    def set_indent_left_value(self, indent_left: Union[str, None]):
        self.set_property_value(k_const.Par_indent_left, indent_left)

    def get_indent_right(self) -> Union[str, None]:
        return self.get_property(k_const.Par_indent_right)

    def set_indent_right_value(self, indent_right: Union[str, None]):
        self.set_property_value(k_const.Par_indent_right, indent_right)

    def get_hanging(self) -> Union[str, None]:
        return self.get_property(k_const.Par_hanging)

    def set_hanging_value(self, hanging: Union[str, None]):
        self.set_property_value(k_const.Par_hanging, hanging)

    def get_first_line(self) -> Union[str, None]:
        return self.get_property(k_const.Par_first_line)

    def set_first_line_value(self, first_line: Union[str, None]):
        self.set_property_value(k_const.Par_first_line, first_line)

    def is_keep_lines(self) -> bool:
        result = self.get_property(k_const.Par_keep_lines)
        return result if result is not None else False

    def set_as_keep_lines(self, is_keep_lines: Union[bool, None] = True):
        self.set_property_value(k_const.Par_keep_lines, is_keep_lines)

    def is_keep_next(self) -> bool:
        result = self.get_property(k_const.Par_keep_next)
        return result if result is not None else False

    def set_as_keep_next(self, is_keep_next: Union[bool, None] = True):
        self.set_property_value(k_const.Par_keep_next, is_keep_next)

    def get_outline_level(self) -> Union[str, None]:
        return self.get_property(k_const.Par_outline_level)

    def set_outline_level_value(self, level: Union[str, int, None]):
        """
        :param level: 0-9
        """
        self.set_property_value(k_const.Par_outline_level, level)

    def get_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left     (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, space, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(k_const.get_key('paragraph_border', direction, property_name))

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right, left     (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, space, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(k_const.get_key('paragraph_border', direction, property_name), value)


class RunPropertiesGetSetMixin(GetSetMixin, ABC):
    def get_size(self) -> Union[str, None]:
        return self.get_property(k_const.Run_size)

    def set_size_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Run_size, value)

    def is_bold(self) -> bool:
        result = self.get_property(k_const.Run_is_bold)
        return result if result is not None else False

    def set_as_bold(self, is_bold: Union[bool, None] = True):
        self.set_property_value(k_const.Run_is_bold, is_bold)

    def is_italic(self) -> bool:
        result = self.get_property(k_const.Run_is_italic)
        return result if result is not None else False

    def set_as_italic(self, is_italic: Union[bool, None] = True):
        self.set_property_value(k_const.Run_is_italic, is_italic)

    def is_strike(self) -> bool:
        result = self.get_property(k_const.Run_is_strike)
        return result if result is not None else False

    def set_as_strike(self, is_strike: Union[bool, None] = True):
        self.set_property_value(k_const.Run_is_strike, is_strike)

    def get_language(self) -> Union[str, None]:
        return self.get_property(k_const.Run_language)

    def set_language_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Run_language, value)

    def get_vertical_align(self) -> Union[str, None]:
        return self.get_property(k_const.Run_vertical_align)

    def set_vertical_align_value(self, value: Union[str, None]):
        """
        param value: 'superscript' or 'subscript'
        """
        self.set_property_value(k_const.Run_vertical_align, value)

    def get_color(self) -> Union[str, None]:
        return self.get_property(k_const.Run_color)

    def set_color_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(k_const.Run_color, value)

    def get_theme_color(self) -> Union[str, None]:
        return self.get_property(k_const.Run_theme_color)

    def set_theme_color_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Run_theme_color, value)

    def get_background_color(self) -> Union[str, None]:
        return self.get_property(k_const.Run_background_color)

    def set_background_color_value(self, value: Union[str, None]):
        """
        param value: possible values: black, blue, cyan, darkBlue, darkCyan, darkGray, darkGreen, darkMagenta,
                                    darkRed, darkYellow, green, lightGray, magenta, none, red, white, yellow
        """
        self.set_property_value(k_const.Run_background_color, value)

    def get_background_fill(self) -> Union[str, None]:
        return self.get_property(k_const.Run_background_fill)

    def set_background_fill_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(k_const.Run_background_fill, value)

    def get_underline(self) -> Union[str, None]:
        return self.get_property(k_const.Run_underline)

    def set_underline_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Run_underline, value)

    def get_underline_color(self) -> Union[str, None]:
        return self.get_property(k_const.Run_underline_color)

    def set_underline_color_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(k_const.Run_underline_color, value)

    def get_border(self, property_name: str) -> Union[str, None]:
        """
        :param property_name: color, size, space, type
        (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(k_const.get_key('border', property_name=property_name))

    def set_border_value(self, property_name: str, value: Union[str, None]):
        """
        :param property_name: color, size, space or type
        (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(k_const.get_key('border', property_name=property_name), value)


class TablePropertiesGetSetMixin(GetSetMixin, ABC):
    def get_width(self) -> Tuple[Union[str, None], Union[str, None], Union[str, None]]:
        """
        return: Tuple(width, type, layout)
        """
        return (
            self.get_property(k_const.Tab_width),
            self.get_property(k_const.Tab_width_type),
            self.get_property(k_const.Tab_layout)
        )

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None], layout: str = 'autofit'):
        self.set_property_value(k_const.Tab_width, width)
        self.set_property_value(k_const.Tab_width_type, width_type)
        self.set_property_value(k_const.Tab_layout, layout)

    def get_align(self) -> Union[str, None]:
        return self.get_property(k_const.Tab_align)

    def set_align_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Tab_align, value)

    def get_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left      (keys of XMLementPropertyDescriptions.Const_property_names dict)
        :param property_name: color, size, type         (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(k_const.get_key('border', direction, property_name))

    def set_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: top, bottom, right, left      (keys of XMLementPropertyDescriptions.Const_property_names dict)
        :param property_name: color, size, type         (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(k_const.get_key('border', direction, property_name), value)

    def get_indentation(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(value, type)
        """
        return self.get_property(k_const.Tab_indentation), self.get_property(k_const.Tab_indentation_type)

    def set_indentation_value(self, width: Union[str, None], width_type: Union[str, None]):
        self.set_property_value(k_const.Tab_indentation, width)
        self.set_property_value(k_const.Tab_indentation_type, width_type)

    def get_inside_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color', 'size', 'type'  (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        return self.get_property(k_const.get_key('borders_inside', direction, property_name))

    def set_inside_border_value(self, direction: str, property_name: str, value: Union[str, None]):
        """
        :param direction: 'horizontal' or 'vertical'
        :param property_name: 'color', 'size', 'type'  (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        self.set_property_value(k_const.get_key('borders_inside', direction, property_name), value)

    def get_cells_margin(self, direction: str) -> Tuple[Union[str, None], Union[str, None]]:
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :result: value, value_type
        """
        return (
            self.get_property(k_const.get_key('cell_margin', direction, "size")),
            self.get_property(k_const.get_key('cell_margin', direction, "type"))
        )

    def set_cells_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param value_type: 'dxa' or 'nil'
        """
        self.set_property_value(k_const.get_key('cell_margin', direction, "size"), value)
        self.set_property_value(k_const.get_key('cell_margin', direction, "type"), value_type)


class RowPropertiesGetSetMixin(GetSetMixin, ABC):
    def is_header(self) -> bool:
        result = self.get_property(k_const.Row_is_header)
        return result if result is not None else False

    def set_as_header(self, is_header: Union[bool, None] = True):
        self.set_property_value(k_const.Row_is_header, is_header)

    def get_height(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(height, rule)
        """
        return self.get_property(k_const.Row_height), self.get_property(k_const.Row_height_rule)

    def set_height_value(self, value: Union[str, None], height_rule: str = 'atLeast'):
        self.set_property_value(k_const.Row_height, value)
        self.set_property_value(k_const.Row_height_rule, height_rule)


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
        return self.get_property(k_const.Cell_fill_color)

    def set_fill_color_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Cell_fill_color, value)

    def get_fill_theme(self) -> Union[str, None]:
        return self.get_property(k_const.Cell_fill_theme)

    def set_fill_theme_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Cell_fill_theme, value)

    def get_border(self, direction: str, property_name: str) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left  (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param property_name: color, size, type (keys of XMLementPropertyDescriptions.Const_property_names dict)
        """
        result = self.get_property(k_const.get_key('cell_border', direction, property_name))
        if result is None or (property_name == 'color' and result == 'auto'):
            if not (self.is_first_in_row() and direction == 'left') and \
               not (self.is_last_in_row() and direction == 'right') and \
               not (self.is_top() and direction == 'top') and \
               not (self.is_bottom() and direction == 'bottom'):
                d: str = 'horizontal' if (direction == 'top' or direction == 'bottom') else 'vertical'
                return self.get_parent_table().get_inside_border(d, property_name)
        return result

    def get_width(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(width, type)
        """
        return self.get_property(k_const.Cell_width), self.get_property(k_const.Cell_width_type)

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None]):
        self.set_property_value(k_const.Cell_width, width)
        self.set_property_value(k_const.Cell_width_type, width_type)

    def get_vertical_align(self) -> Union[str, None]:
        return self.get_property(k_const.Cell_vertical_align)

    def set_vertical_align_value(self, value: str):
        self.set_property_value(k_const.Cell_vertical_align, value)

    def get_text_direction(self) -> Union[str, None]:
        return self.get_property(k_const.Cell_text_direction)

    def set_text_direction_value(self, value: Union[str, None]):
        self.set_property_value(k_const.Cell_text_direction, value)

    def get_margin(self, direction: str) -> Tuple[Union[str, None], Union[str, None]]:
        """
        :param direction: 'top', 'bottom', 'right', 'left'
        (keys of XMLementPropertyDescriptions.Const_directions dict)
        :return: (margin, type)
        """
        if self.get_property(k_const.get_key('margin', direction, "size")) is None:
            return (
                self.get_parent_table().get_property(k_const.get_key('cell_margin', direction, "size")),
                self.get_parent_table().get_property(k_const.get_key('cell_margin', direction, "type"))
            )
        return (
            self.get_property(k_const.get_key('margin', direction, "size")),
            self.get_property(k_const.get_key('margin', direction, "type"))
        )

    def set_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction:  'top' or 'bottom' or 'right' or 'left'
        (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param value_type: 'dxa' or 'nil'
        """
        self.set_property_value(k_const.get_key('margin', direction, "size"), value)
        self.set_property_value(k_const.get_key('margin', direction, "type"), value_type)
