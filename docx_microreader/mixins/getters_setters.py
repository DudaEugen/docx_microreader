from ..constants import property_enums as pr_const
from typing import Tuple, Union
from abc import ABC, abstractmethod


class GetSetMixin(ABC):
    @abstractmethod
    def get_property(self, property_name, is_find_missed_or_true: bool = True):
        pass

    @abstractmethod
    def set_property_value(self, property_name, value: Union[str, bool, None]):
        pass


class ParagraphPropertiesGetSetMixin(GetSetMixin, ABC):
    def get_align(self) -> Union[str, None]:
        return self.get_property(pr_const.ParagraphProperty.ALIGN)

    def set_align_value(self, align: Union[str, None]):
        """
        :param align: 'both' or 'right' or 'center' or 'left
        """
        self.set_property_value(pr_const.ParagraphProperty.ALIGN, align)

    def get_indent_left(self) -> Union[str, None]:
        return self.get_property(pr_const.ParagraphProperty.INDENT_LEFT)

    def set_indent_left_value(self, indent_left: Union[str, None]):
        self.set_property_value(pr_const.ParagraphProperty.INDENT_LEFT, indent_left)

    def get_indent_right(self) -> Union[str, None]:
        return self.get_property(pr_const.ParagraphProperty.INDENT_RIGHT)

    def set_indent_right_value(self, indent_right: Union[str, None]):
        self.set_property_value(pr_const.ParagraphProperty.INDENT_RIGHT, indent_right)

    def get_hanging(self) -> Union[str, None]:
        return self.get_property(pr_const.ParagraphProperty.HANGING)

    def set_hanging_value(self, hanging: Union[str, None]):
        self.set_property_value(pr_const.ParagraphProperty.HANGING, hanging)

    def get_first_line(self) -> Union[str, None]:
        return self.get_property(pr_const.ParagraphProperty.FIRST_LINE)

    def set_first_line_value(self, first_line: Union[str, None]):
        self.set_property_value(pr_const.ParagraphProperty.FIRST_LINE, first_line)

    def is_keep_lines(self) -> bool:
        result = self.get_property(pr_const.ParagraphProperty.KEEP_LINES)
        return result if result is not None else False

    def set_as_keep_lines(self, is_keep_lines: Union[bool, None] = True):
        self.set_property_value(pr_const.ParagraphProperty.KEEP_LINES, is_keep_lines)

    def is_keep_next(self) -> bool:
        result = self.get_property(pr_const.ParagraphProperty.KEEP_NEXT)
        return result if result is not None else False

    def set_as_keep_next(self, is_keep_next: Union[bool, None] = True):
        self.set_property_value(pr_const.ParagraphProperty.KEEP_NEXT, is_keep_next)

    def get_outline_level(self) -> Union[str, None]:
        return self.get_property(pr_const.ParagraphProperty.OUTLINE_LEVEL)

    def set_outline_level_value(self, level: Union[str, int, None]):
        """
        :param level: 0-9
        """
        self.set_property_value(pr_const.ParagraphProperty.OUTLINE_LEVEL, level)

    def get_border(self, direction: Union[str, pr_const.Direction],
                   property_name: [str, pr_const.BorderProperty]) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left     (or corresponding values of property_enums.Direction)
        :param property_name: color, size, space, type (or corresponding values of property_enums.BorderProperty)
        """
        return self.get_property(
            pr_const.ParagraphProperty.get_border_property_enum_value(direction, property_name)
        )

    def set_border_value(self, direction: Union[str, pr_const.Direction],
                         property_name: [str, pr_const.BorderProperty], value: Union[str, None]):
        """
        :param direction: top, bottom, right, left     (or corresponding values of property_enums.Direction)
        :param property_name: color, size, space, type (or corresponding values of property_enums.BorderProperty)
        """
        self.set_property_value(
            pr_const.ParagraphProperty.get_border_property_enum_value(direction, property_name), value
        )


class RunPropertiesGetSetMixin(GetSetMixin, ABC):
    def get_font(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.FONT_ASCII)

    def get_size(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.SIZE)

    def set_size_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.RunProperty.SIZE, value)

    def is_bold(self) -> bool:
        result = self.get_property(pr_const.RunProperty.BOLD)
        return result if result is not None else False

    def set_as_bold(self, is_bold: Union[bool, None] = True):
        self.set_property_value(pr_const.RunProperty.BOLD, is_bold)

    def is_italic(self) -> bool:
        result = self.get_property(pr_const.RunProperty.ITALIC)
        return result if result is not None else False

    def set_as_italic(self, is_italic: Union[bool, None] = True):
        self.set_property_value(pr_const.RunProperty.ITALIC, is_italic)

    def is_strike(self) -> bool:
        result = self.get_property(pr_const.RunProperty.STRIKE)
        return result if result is not None else False

    def set_as_strike(self, is_strike: Union[bool, None] = True):
        self.set_property_value(pr_const.RunProperty.STRIKE, is_strike)

    def get_language(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.LANGUAGE)

    def set_language_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.RunProperty.LANGUAGE, value)

    def get_vertical_align(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.VERTICAL_ALIGN)

    def set_vertical_align_value(self, value: Union[str, None]):
        """
        param value: 'superscript' or 'subscript'
        """
        self.set_property_value(pr_const.RunProperty.VERTICAL_ALIGN, value)

    def get_color(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.COLOR)

    def set_color_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(pr_const.RunProperty.COLOR, value)

    def get_theme_color(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.THEME_COLOR)

    def set_theme_color_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.RunProperty.THEME_COLOR, value)

    def get_background_color(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.BACKGROUND_COLOR)

    def set_background_color_value(self, value: Union[str, None]):
        """
        param value: possible values: black, blue, cyan, darkBlue, darkCyan, darkGray, darkGreen, darkMagenta,
                                    darkRed, darkYellow, green, lightGray, magenta, none, red, white, yellow
        """
        self.set_property_value(pr_const.RunProperty.BACKGROUND_COLOR, value)

    def get_background_fill(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.BACKGROUND_FILL)

    def set_background_fill_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(pr_const.RunProperty.BACKGROUND_FILL, value)

    def get_underline(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.UNDERLINE_TYPE)

    def set_underline_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.RunProperty.UNDERLINE_TYPE, value)

    def get_underline_color(self) -> Union[str, None]:
        return self.get_property(pr_const.RunProperty.UNDERLINE_COLOR)

    def set_underline_color_value(self, value: Union[str, None]):
        """
        param value: specifies the color as a hex value in RRGGBB format or auto
        """
        self.set_property_value(pr_const.RunProperty.UNDERLINE_COLOR, value)

    def get_border(self, property_name: [str, pr_const.BorderProperty]) -> Union[str, None]:
        """
        :param property_name: color, size, space, type (or corresponding value of property_enums.BorderProperty)
        """
        return self.get_property(pr_const.RunProperty.get_border_property_enum_value(None, property_name))

    def set_border_value(self, property_name: [str, pr_const.BorderProperty], value: Union[str, None]):
        """
        :param property_name: color, size, space or type (or corresponding value of property_enums.BorderProperty)
        """
        self.set_property_value(pr_const.RunProperty.get_border_property_enum_value(None, property_name), value)


class TablePropertiesGetSetMixin(GetSetMixin, ABC):
    def get_width(self) -> Tuple[Union[str, None], Union[str, None], Union[str, None]]:
        """
        return: Tuple(width, type, layout)
        """
        return (
            self.get_property(pr_const.TableProperty.WIDTH),
            self.get_property(pr_const.TableProperty.WIDTH_TYPE),
            self.get_property(pr_const.TableProperty.LAYOUT)
        )

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None], layout: str = 'autofit'):
        self.set_property_value(pr_const.TableProperty.WIDTH, width)
        self.set_property_value(pr_const.TableProperty.WIDTH_TYPE, width_type)
        self.set_property_value(pr_const.TableProperty.LAYOUT, layout)

    def get_align(self) -> Union[str, None]:
        return self.get_property(pr_const.TableProperty.ALIGN)

    def set_align_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.TableProperty.ALIGN, value)

    def get_border(self, direction: Union[str, pr_const.Direction],
                   property_name: [str, pr_const.BorderProperty]) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left      (or corresponding value of property_enums.Direction)
        :param property_name: color, size, type         (or corresponding value of property_enums.BorderProperty)
        """
        return self.get_property(pr_const.TableProperty.get_border_property_enum_value(direction, property_name))

    def set_border_value(self, direction: Union[str, pr_const.Direction],
                         property_name: [str, pr_const.BorderProperty], value: Union[str, None]):
        """
        :param direction: top, bottom, right, left      (or corresponding value of property_enums.Direction)
        :param property_name: color, size, type         (or corresponding value of property_enums.BorderProperty)
        """
        self.set_property_value(
            pr_const.TableProperty.get_border_property_enum_value(direction, property_name),
            value
        )

    def get_indentation(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(value, type)
        """
        return (self.get_property(pr_const.TableProperty.INDENTATION),
                self.get_property(pr_const.TableProperty.INDENTATION_TYPE))

    def set_indentation_value(self, width: Union[str, None], width_type: Union[str, None]):
        self.set_property_value(pr_const.TableProperty.INDENTATION, width)
        self.set_property_value(pr_const.TableProperty.INDENTATION_TYPE, width_type)

    def get_inside_border(self, direction: Union[str, pr_const.Direction],
                          property_name: [str, pr_const.BorderProperty]) -> Union[str, None]:
        """
        :param direction: 'horizontal' or 'vertical'    (or corresponding value of property_enums.Direction)
        :param property_name: 'color', 'size', 'type'   (or corresponding value of property_enums.BorderProperty)
        """
        return self.get_property(pr_const.TableProperty.get_border_property_enum_value(direction, property_name))

    def set_inside_border_value(self, direction: Union[str, pr_const.Direction],
                                property_name: [str, pr_const.BorderProperty], value: Union[str, None]):
        """
        :param direction: 'horizontal' or 'vertical'    (or corresponding value of property_enums.Direction)
        :param property_name: 'color', 'size', 'type'   (or corresponding value of property_enums.BorderProperty)
        """
        self.set_property_value(
            pr_const.TableProperty.get_border_property_enum_value(direction, property_name),
            value
        )

    def get_cells_margin(self, direction: str) -> Tuple[Union[str, None], Union[str, None]]:
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (or corresponding value of property_enums.Direction)
        :result: value, value_type
        """
        return (
            self.get_property(
                pr_const.TableProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.SIZE)
            ),
            self.get_property(
                pr_const.TableProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.TYPE)
            )
        )

    def set_cells_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction: 'top', 'bottom', 'right', 'left'  (or corresponding value of property_enums.Direction)
        :param value_type: 'dxa' or 'nil'
        """
        self.set_property_value(
            pr_const.TableProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.SIZE),
            value
        )
        self.set_property_value(
            pr_const.TableProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.TYPE),
            value_type
        )


class RowPropertiesGetSetMixin(GetSetMixin, ABC):
    def is_header(self) -> bool:
        result = self.get_property(pr_const.RowProperty.HEADER)
        return result if result is not None else False

    def set_as_header(self, is_header: Union[bool, None] = True):
        self.set_property_value(pr_const.RowProperty.HEADER, is_header)

    def get_height(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(height, rule)
        """
        return (
            self.get_property(pr_const.RowProperty.HEIGHT),
            self.get_property(pr_const.RowProperty.HEIGHT_RULE)
        )

    def set_height_value(self, value: Union[str, None], height_rule: str = 'atLeast'):
        self.set_property_value(pr_const.RowProperty.HEIGHT, value)
        self.set_property_value(pr_const.RowProperty.HEIGHT_RULE, height_rule)


class CellPropertiesGetSetMixin(GetSetMixin, ABC):
    @abstractmethod
    def get_parent_table(self):
        pass

    @abstractmethod
    def is_top(self) -> bool:
        pass

    @abstractmethod
    def is_bottom(self) -> bool:
        pass

    @abstractmethod
    def is_first_in_row(self) -> bool:
        pass

    @abstractmethod
    def is_last_in_row(self) -> bool:
        pass

    def get_fill_color(self) -> Union[str, None]:
        return self.get_property(pr_const.CellProperty.FILL_COLOR)

    def set_fill_color_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.CellProperty.FILL_COLOR, value)

    def get_fill_theme(self) -> Union[str, None]:
        return self.get_property(pr_const.CellProperty.FILL_THEME)

    def set_fill_theme_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.CellProperty.FILL_THEME, value)

    def get_border(self, direction: Union[str, pr_const.Direction],
                   property_name: Union[str, pr_const.BorderProperty]) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left  (or corresponding value of property_enums.Direction)
        :param property_name: color, size, type     (or corresponding value of property_enums.BorderProperty)
        """
        direct: pr_const.Direction = pr_const.convert_to_enum_element(direction, pr_const.Direction)
        pr_name: pr_const.BorderProperty = pr_const.convert_to_enum_element(property_name, pr_const.BorderProperty)

        maybe_result = self.get_property(pr_const.CellProperty.get_border_property_enum_value(direct, pr_name))
        if maybe_result is None or (pr_name == pr_const.BorderProperty.COLOR and maybe_result == 'auto'):
            if not (self.is_first_in_row() and direct == pr_const.Direction.LEFT) and \
               not (self.is_last_in_row() and direct == pr_const.Direction.RIGHT) and \
               not (self.is_top() and direct == pr_const.Direction.TOP) and \
               not (self.is_bottom() and direct == pr_const.Direction.BOTTOM):
                d: pr_const.Direction = pr_const.Direction.horizontal_or_vertical_straight(direct)
                return self.get_parent_table().get_inside_border(d, pr_name)
        return maybe_result

    def get_width(self) -> Tuple[Union[str, None], Union[str, None]]:
        """
        return: Tuple(width, type)
        """
        return (self.get_property(pr_const.CellProperty.WIDTH),
                self.get_property(pr_const.CellProperty.WIDTH_TYPE))

    def set_width_value(self, width: Union[str, None], width_type: Union[str, None]):
        self.set_property_value(pr_const.CellProperty.WIDTH.type, width)
        self.set_property_value(pr_const.CellProperty.WIDTH_TYPE.type, width_type)

    def get_vertical_align(self) -> Union[str, None]:
        return self.get_property(pr_const.CellProperty.VERTICAL_ALIGN)

    def set_vertical_align_value(self, value: str):
        self.set_property_value(pr_const.CellProperty.VERTICAL_ALIGN, value)

    def get_text_direction(self) -> Union[str, None]:
        return self.get_property(pr_const.CellProperty.TEXT_DIRECTION)

    def set_text_direction_value(self, value: Union[str, None]):
        self.set_property_value(pr_const.CellProperty.TEXT_DIRECTION, value)

    def get_margin(self, direction: str) -> Tuple[Union[str, None], Union[str, None]]:
        """
        :param direction: 'top', 'bottom', 'right', 'left'
        (keys of XMLementPropertyDescriptions.Const_directions dict)
        :return: (margin, type)
        """
        if self.get_property(
                pr_const.CellProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.SIZE)
        ) is None:
            return (
                self.get_parent_table().get_property(
                    pr_const.TableProperty.get_cell_margin_property_enum_value(direction,
                                                                               pr_const.MarginProperty.SIZE)
                ),
                self.get_parent_table().get_property(
                    pr_const.TableProperty.get_cell_margin_property_enum_value(direction,
                                                                               pr_const.MarginProperty.TYPE)
                )
            )
        return (
            self.get_property(
                pr_const.CellProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.SIZE)
            ),
            self.get_property(
                pr_const.CellProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.TYPE)
            )
        )

    def set_margin_value(self, direction: str, value: Union[str, None], value_type: Union[str, None] = 'dxa'):
        """
        :param direction:  'top' or 'bottom' or 'right' or 'left'
        (keys of XMLementPropertyDescriptions.Const_directions dict)
        :param value_type: 'dxa' or 'nil'
        """
        self.set_property_value(
            pr_const.CellProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.SIZE),
            value
        )
        self.set_property_value(
            pr_const.CellProperty.get_cell_margin_property_enum_value(direction, pr_const.MarginProperty.TYPE),
            value_type
        )
