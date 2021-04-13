from .docx_parser import DocumentParser, XMLement
import xml.etree.ElementTree as ET
from typing import List, Callable, Dict, Union, Tuple
from .mixins.getters_setters import ParagraphPropertiesGetSetMixin, RunPropertiesGetSetMixin, \
                                    TablePropertiesGetSetMixin, RowPropertiesGetSetMixin, CellPropertiesGetSetMixin
from .constants import property_enums as pr_const
from .constants.translate_formats import TranslateFormat


class Image(XMLement):
    element_description = pr_const.Element.IMAGE
    _is_unique = True
    from docx_microreader.translators.html.html_translators import ImageTranslatorToHTML
    translators = {
        TranslateFormat.HTML: ImageTranslatorToHTML(),
    }

    def __init__(self, element: ET.Element, parent):
        super(Image, self).__init__(element, parent)

    def translate(self, to_format: Union[TranslateFormat, str, None] = None, is_recursive_translate: bool = True):
        """
        :param to_format: using translate_format of element if None
        :param is_recursive_translate: pass to_format to inner element if True
        """
        translator = self.translators[TranslateFormat(to_format)] if to_format is not None else \
                     self.translators[self.translate_format]
        return translator.translate(self, [])

    def get_path(self):
        return self._get_document().get_image(self._properties[pr_const.ImageProperty.ID.key].value)

    def get_size(self) -> (int, int):
        """
        :return: (horizontal size, vertical size)
        """
        return self.get_parent().get_size()


class Drawing(XMLement):
    element_description = pr_const.Element.DRAWING

    from docx_microreader.translators.html.html_translators import ContainerTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: ContainerTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Image]

    def get_size(self) -> (int, int):
        """
        :return: (horizontal size, vertical size)
        """
        return (
            int(self._properties[pr_const.DrawingProperty.HORIZONTAL_SIZE.key].value),
            int(self._properties[pr_const.DrawingProperty.VERTICAL_SIZE.key].value)
        )


class LineBreak(XMLement):
    element_description = pr_const.Element.LINE_BREAK
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    from docx_microreader.translators.html.html_translators import LineBreakTranslatorToHTML
    translators = {
        TranslateFormat.HTML: LineBreakTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }


class Text(XMLement):
    element_description = pr_const.Element.TEXT

    from docx_microreader.translators.html.html_translators import TextTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: TextTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    def __init__(self, element: ET.Element, parent):
        self.content: str = element.text
        super(Text, self).__init__(element, parent)

    def translate(self, to_format: Union[TranslateFormat, str, None] = None, is_recursive_translate: bool = True):
        """
        :param to_format: using translate_format of element if None
        :param is_recursive_translate: pass to_format to inner element if True
        """
        translator = self.translators[TranslateFormat(to_format)] if to_format is not None else \
                     self.translators[self.translate_format]
        return translator.translate(self, [self.content])


class Run(XMLement, RunPropertiesGetSetMixin):
    element_description = pr_const.Element.RUN

    from docx_microreader.translators.html.html_translators import RunTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: RunTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Text, Drawing, LineBreak]

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[pr_const.RunProperty.STYLE.key].value

    def get_property(self, property_name, is_find_missed_or_true: bool = True):
        """
        :param property_name: key of property (str or instance of Enum from constants.property_enums)
        :param is_find_missed_or_true: recursive find value if result equal Property.Missed. Return True if not find
        :return: Property.value or None
        """
        from .properties import Property

        key: str = XMLement._key_of_property(property_name)

        result = super(Run, self).get_property(key, is_find_missed_or_true)
        if result is not None and (not isinstance(result, Property.Missed) or not is_find_missed_or_true):
            return result

        parent_prop = self.parent.get_property(key, is_find_missed_or_true)
        if parent_prop is not None:
            result = parent_prop
            if not isinstance(parent_prop, Property.Missed) or not is_find_missed_or_true:
                return parent_prop

        if self._default_style is not None:
            default_style = self._get_default_style_from_document(self._default_style)
            if default_style is not None:
                default_style_result = default_style.get_property(key, is_find_missed_or_true)
                result = default_style_result
                if not isinstance(default_style_result, Property.Missed) or not is_find_missed_or_true:
                    return default_style_result

        if isinstance(result, Property.Missed) and is_find_missed_or_true:
            return True
        return result


class Paragraph(XMLement, ParagraphPropertiesGetSetMixin):
    element_description = pr_const.Element.PARAGRAPH

    _properties_unificators = {
        pr_const.ParagraphProperty.ALIGN.key: [('left', ['start']),
                                               ('right', ['end']),
                                               ('center', []),
                                               ('both', []),
                                               ('distribute', [])]
    }

    from docx_microreader.translators.html.html_translators import ParagraphTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: ParagraphTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Run]

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[pr_const.ParagraphProperty.STYLE.key].value

    def get_property(self, property_name, is_find_missed_or_true: bool = True):
        """
        :param property_name: key of property (str or instance of Enum from constants.property_enums)
        :param is_find_missed_or_true: recursive find value if result equal Property.Missed. Return True if not find
        :return: Property.value or None
        """
        from .properties import Property

        key: str = XMLement._key_of_property(property_name)

        result = super(Paragraph, self).get_property(key, is_find_missed_or_true)
        if result is not None and (not isinstance(result, Property.Missed) or not is_find_missed_or_true):
            return result

        parent_prop = self.parent.get_property(key, is_find_missed_or_true)
        if parent_prop is not None:
            result = parent_prop
            if not isinstance(parent_prop, Property.Missed) or not is_find_missed_or_true:
                return parent_prop

        if self._default_style is not None:
            default_style = self._get_default_style_from_document(self._default_style)
            if default_style is not None:
                default_style_result = default_style.get_property(key, is_find_missed_or_true)
                result = default_style_result
                if not isinstance(default_style_result, Property.Missed) or not is_find_missed_or_true:
                    return default_style_result

        if isinstance(result, Property.Missed) and is_find_missed_or_true:
            return True
        return result


class Cell(XMLement, CellPropertiesGetSetMixin):
    element_description = pr_const.Element.CELL

    from docx_microreader.translators.html.html_translators import CellTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: CellTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Paragraph, ('docx_microreader.models', 'Table')]

    def __init__(self, element: ET.Element, parent):
        self.row_span: int = 1
        self.is_header: bool = False
        self.index_in_row: int = -1
        super(Cell, self).__init__(element, parent)

    def __get_table_area_style(self, table_area_style_type: str):
        return self.get_parent_table().get_style().get_table_area_style(table_area_style_type)

    def __get_property_of_table_area_style(self, property_name, table_area_style_type: str):
        table_area_style = self.__get_table_area_style(table_area_style_type)
        if table_area_style is not None:
            return table_area_style.get_property(property_name)
        return None

    def __get_property_of_top_left_cell_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_top() and self.get_parent_table().is_use_style_of_first_row() and \
                self.is_first_in_row() and self.get_parent_table().is_use_style_of_first_column():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.TOP_LEFT_CELL.value), True
        return None, False

    def __get_property_of_top_right_cell_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_top() and self.get_parent_table().is_use_style_of_first_row() and \
                self.is_last_in_row() and self.get_parent_table().is_use_style_of_last_column():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.TOP_RIGHT_CELL.value), True
        return None, False

    def __get_property_of_bottom_left_cell_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_bottom() and self.get_parent_table().is_use_style_of_last_row() and \
                self.is_first_in_row() and self.get_parent_table().is_use_style_of_first_column():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.BOTTOM_LEFT_CELL.value), \
                   True
        return None, False

    def __get_property_of_bottom_right_cell_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_bottom() and self.get_parent_table().is_use_style_of_last_row() and \
                self.is_last_in_row() and self.get_parent_table().is_use_style_of_last_column():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.BOTTOM_RIGHT_CELL.value),\
                   True
        return None, False

    def __get_property_of_first_column_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_first_in_row() and self.get_parent_table().is_use_style_of_first_column():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.FIRST_COLUMN.value), True
        return None, False

    def __get_property_of_last_column_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_last_in_row() and self.get_parent_table().is_use_style_of_last_column():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.LAST_COLUMN.value), True
        return None, False

    def __get_property_of_first_row_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_top() and self.get_parent_table().is_use_style_of_first_row():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.FIRST_ROW.value), True
        return None, False

    def __get_property_of_last_row_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_bottom() and self.get_parent_table().is_use_style_of_last_row():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.LAST_ROW.value), True
        return None, False

    def __get_property_of_odd_row_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.get_parent_row().is_odd() and self.get_parent_table().is_use_style_of_horizontal_banding():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.ODD_ROW.value), True
        return None, False

    def __get_property_of_even_row_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.get_parent_row().is_even() and self.get_parent_table().is_use_style_of_horizontal_banding():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.EVEN_ROW.value), True
        return None, False

    def __get_property_of_odd_column_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_odd() and self.get_parent_table().is_use_style_of_vertical_banding():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.ODD_COLUMN.value), True
        return None, False

    def __get_property_of_even_column_area_style(self, property_name) -> Tuple[Union[str, None, bool], bool]:
        if self.is_even() and self.get_parent_table().is_use_style_of_vertical_banding():
            return self.__get_property_of_table_area_style(property_name, pr_const.TableArea.EVEN_COLUMN.value), True
        return None, False

    def __define_table_area_style_and_get_property(self, property_name) -> Union[str, None, bool]:
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

    def __define_table_area_style_and_get_property_for_horizontal_inside_borders(self, property_name) -> \
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

    def __define_table_area_style_and_get_property_for_vertical_inside_borders(self, property_name) -> \
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

    def __define_table_area_style_and_get_property_for_inside_borders_of_header(self, property_name) -> \
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

    def get_property(self, property_name, is_find_missed_or_true: bool = True):
        """
        :param property_name: key of property (str or instance of Enum from constants.property_enums)
        :param is_find_missed_or_true: recursive find value if result equal Property.Missed. Return True if not find
        :return: Property.value or None
        """
        from .properties import Property

        result = None
        key: str = XMLement._key_of_property(property_name)

        result = super(Cell, self).get_property(key, is_find_missed_or_true)
        if result is not None and (not isinstance(result, Property.Missed) or not is_find_missed_or_true):
            return result

        area_style_result = self.__define_table_area_style_and_get_property(key)
        if area_style_result is not None:
            result = area_style_result
            if not isinstance(area_style_result, Property.Missed) or not is_find_missed_or_true:
                return area_style_result

        parent_prop = self.get_parent_table().get_property(key, is_find_missed_or_true)
        if parent_prop is not None:
            result = parent_prop
            if not isinstance(parent_prop, Property.Missed) or not is_find_missed_or_true:
                return parent_prop

        if isinstance(result, Property.Missed) and is_find_missed_or_true:
            return True
        return result

    def get_border(self, direction: str, property_name) -> Union[str, None]:
        """
        :param direction: top, bottom, right, left  (or corresponding value of property_enums.Direction)
        :param property_name: color, size, type     (or corresponding value of property_enums.BorderProperty)
        """
        direct: pr_const.Direction = pr_const.convert_to_enum_element(direction, pr_const.Direction)
        pr_name: pr_const.BorderProperty = pr_const.convert_to_enum_element(property_name, pr_const.BorderProperty)
        maybe_result = self._properties.get(pr_const.CellProperty.get_border_property_enum_value(direct, pr_name).key)
        if maybe_result is not None:
            maybe_result = maybe_result.value
        if maybe_result is None or (maybe_result == 'auto' and pr_name == pr_const.BorderProperty.COLOR):
            if not (self.is_first_in_row() and direct == pr_const.Direction.LEFT) and \
                    not (self.is_last_in_row() and direct == pr_const.Direction.RIGHT) and \
                    not (self.is_top() and direct == pr_const.Direction.TOP) and \
                    not (self.is_bottom() and direct == pr_const.Direction.BOTTOM):
                d = pr_const.Direction.horizontal_or_vertical_straight(direct)
                if direct == pr_const.Direction.BOTTOM and \
                        self.get_parent_table().is_use_style_of_first_row() and \
                        self.get_parent_table().header_row_number > 1 and \
                        self.get_parent_row().is_header() and not self.get_parent_row().is_last_row_in_header:
                    return self.__define_table_area_style_and_get_property_for_inside_borders_of_header(
                                pr_const.TableProperty.get_border_property_enum_value(d, pr_name)
                            )
                maybe_result = self.get_parent_table().get_inside_border(d, pr_name)
                if maybe_result is None:
                    maybe_result = self.__define_table_area_style_and_get_property(
                        pr_const.CellProperty.get_border_property_enum_value(direct, pr_name)
                    )
                    if maybe_result is None:
                        if d == pr_const.Direction.HORIZONTAL:
                            maybe_result = self.__define_table_area_style_and_get_property_for_horizontal_inside_borders(
                                pr_const.TableProperty.get_border_property_enum_value(d, pr_name)
                            )
                        else:
                            maybe_result = self.__define_table_area_style_and_get_property_for_vertical_inside_borders(
                                pr_const.TableProperty.get_border_property_enum_value(d, pr_name)
                            )
        return maybe_result

    def set_border_value(self, direction: str, property_name, value: Union[str, None]):
        """
        :param direction: top, bottom, right, left  (or corresponding value of property_enums.Direction)
        :param property_name: color, size, type     (or corresponding value of property_enums.BorderProperty)
        """
        self.set_property_value(
            pr_const.CellProperty.get_border_property_enum_value(direction, property_name),
            value
        )

    def get_col_span(self) -> Union[str, None]:
        return self._properties[pr_const.CellProperty.COLUMN_SPAN.key].value

    def set_col_span_value(self, value: Union[str, int]):
        self._properties[pr_const.CellProperty.COLUMN_SPAN.key].value = str(value)

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


class Row(XMLement, RowPropertiesGetSetMixin):
    element_description = pr_const.Element.ROW

    from docx_microreader.translators.html.html_translators import RowTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: RowTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Cell]

    def __init__(self, element: ET.Element, parent):
        super(Row, self).__init__(element, parent)
        self.cells: List[Cell] = self.inner_elements
        self.__set_index_in_row_for_cells()
        self.is_first_row_in_header: bool = False
        self.is_last_row_in_header: bool = False
        self.index_in_table: int = -1
        if self._properties[pr_const.RowProperty.HEADER.key].value:
            self.__set_cells_as_header()

    def __set_cells_as_header(self):
        for cell in self.cells:
            cell.is_header = True

    def __set_index_in_row_for_cells(self):
        for index in range(len(self.cells)):
            self.cells[index].index_in_row = index

    def get_parent_table(self):
        return self.get_parent()

    def set_as_header(self, is_header: bool = True):
        super(Row, self).set_as_header(is_header)
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


class Table(XMLement, TablePropertiesGetSetMixin):
    element_description = pr_const.Element.TABLE

    _properties_unificators = {
        pr_const.TableProperty.ALIGN.key: [('left', ['start']),
                                           ('right', ['end']),
                                           ('center', [])]
    }

    from docx_microreader.translators.html.html_translators import TableTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: TableTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Row]

    def __init__(self, element: ET.Element, parent):
        super(Table, self).__init__(element, parent)
        self.rows: List[Row] = self.inner_elements
        self.__set_index_in_table_for_rows()
        self.header_row_number: int = 0

    def translate(self, to_format: Union[TranslateFormat, str, None] = None, is_recursive_translate: bool = True):
        """
        :param to_format: using translate_format of element if None
        :param is_recursive_translate: pass to_format to inner element if True
        """
        self.__define_first_and_last_head_rows()
        self.__calculate_rowspan_for_cells()
        return super(Table, self).translate(TranslateFormat(to_format), is_recursive_translate)

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[pr_const.TableProperty.STYLE.key].value

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
                previous_row: Row = self.rows[0]
                for row in self.rows:
                    if row != previous_row and row.is_header():
                        self.header_row_number += 1
                        previous_row.is_last_row_in_header = False
                        row.is_last_row_in_header = True
                    previous_row = row

    def __calculate_rowspan_for_cells(self):
        from .properties import Property

        cell_for_row_span: Dict[int, Cell] = {}
        for row in self.rows:
            col: int = 0
            for cell in row.cells:
                vertical_merge = cell.get_property(pr_const.CellProperty.VERTICAL_MERGE, False)
                if vertical_merge == 'restart':
                    cell_for_row_span[col] = cell
                    cell_for_row_span[col].row_span = 1
                elif vertical_merge == 'continue' or isinstance(vertical_merge, Property.Missed):
                    cell_for_row_span[col].row_span += 1
                col_span = cell.get_col_span()
                col += int(col_span) if col_span is not None else 1

    def is_use_style_of_first_row(self) -> bool:
        if self._properties[pr_const.TableProperty.FIRST_ROW_STYLE_LOOK.key].value is None:
            return False
        return self._properties[pr_const.TableProperty.FIRST_ROW_STYLE_LOOK.key].value == '1'

    def set_as_use_style_of_first_row(self, is_use: bool):
        self._properties[pr_const.TableProperty.FIRST_ROW_STYLE_LOOK.key].value = None if not is_use else '1'

    def is_use_style_of_first_column(self) -> bool:
        if self._properties[pr_const.TableProperty.FIRST_COLUMN_STYLE_LOOK.key].value is None:
            return False
        return self._properties[pr_const.TableProperty.FIRST_COLUMN_STYLE_LOOK.key].value == '1'

    def set_as_use_style_of_first_column(self, is_use: bool):
        self._properties[pr_const.TableProperty.FIRST_COLUMN_STYLE_LOOK.key].value = None if not is_use else '1'

    def is_use_style_of_last_row(self) -> bool:
        if self._properties[pr_const.TableProperty.LAST_ROW_STYLE_LOOK.key].value is None:
            return False
        return self._properties[pr_const.TableProperty.LAST_ROW_STYLE_LOOK.key].value == '1'

    def set_as_use_style_of_last_row(self, is_use: bool):
        self._properties[pr_const.TableProperty.LAST_ROW_STYLE_LOOK.key].value = None if not is_use else '1'

    def is_use_style_of_last_column(self) -> bool:
        if self._properties[pr_const.TableProperty.LAST_COLUMN_STYLE_LOOK.key].value is None:
            return False
        return self._properties[pr_const.TableProperty.LAST_COLUMN_STYLE_LOOK.key].value == '1'

    def set_as_use_style_of_last_column(self, is_use: bool):
        self._properties[pr_const.TableProperty.LAST_COLUMN_STYLE_LOOK.key].value = None if not is_use else '1'

    def is_use_style_of_horizontal_banding(self) -> bool:
        if self._properties[pr_const.TableProperty.NO_HORIZONTAL_BANDING.key].value is None:
            return True
        return self._properties[pr_const.TableProperty.NO_HORIZONTAL_BANDING.key].value == '0'

    def set_as_use_style_of_horizontal_banding(self, is_use: bool):
        self._properties[pr_const.TableProperty.NO_HORIZONTAL_BANDING.key].value = None if is_use else '0'

    def is_use_style_of_vertical_banding(self) -> bool:
        if self._properties[pr_const.TableProperty.NO_VERTICAL_BANDING.key].value is None:
            return True
        return self._properties[pr_const.TableProperty.NO_VERTICAL_BANDING.key].value == '0'

    def set_as_use_style_of_vertical_banding(self, is_use: bool):
        self._properties[pr_const.TableProperty.NO_VERTICAL_BANDING.key].value = None if is_use else '0'


class Body(XMLement):
    element_description = pr_const.Element.BODY

    from docx_microreader.translators.html.html_translators import BodyTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import TranslatorToXML
    translators = {
        TranslateFormat.HTML: BodyTranslatorToHTML(),
        TranslateFormat.XML: TranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Paragraph, Table]


class Document(DocumentParser):
    from docx_microreader.translators.html.html_translators import DocumentTranslatorToHTML
    from docx_microreader.translators.xml.xml_translators import DocumentTranslatorToXML
    translators = {
        TranslateFormat.HTML: DocumentTranslatorToHTML(),
        TranslateFormat.XML: DocumentTranslatorToXML(),
    }

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [Body]

    def __init__(self, path: str, path_for_images: Union[None, str] = None):
        super(Document, self).__init__(path, path_for_images)
        self.body: Body = self.inner_elements[0]

    def translate(self, to_format: Union[TranslateFormat, str], is_recursive_translate: bool = True):
        """
        :param to_format: using translate_format of element if None
        :param is_recursive_translate: pass to_format to inner element if True
        """
        translator = self.translators[TranslateFormat(to_format)]
        return translator.translate(self, [self.body.translate(to_format, is_recursive_translate)])

    def save_as_docx(self, name: str):
        from docx_microreader.translators.xml.xml_translators import DocumentTranslatorToXML
        import os
        import zipfile

        file = open(f'{DocumentTranslatorToXML.template_directory_path()}\\word\\document.xml', 'w', encoding='utf-8')
        document: ET.Element = self.translate(TranslateFormat.XML, is_recursive_translate=True)
        file.write(DocumentTranslatorToXML.document_header() + ET.tostring(document, encoding="unicode"))
        file.close()

        zipf = zipfile.ZipFile(f'{name}.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk(DocumentTranslatorToXML.template_directory_path()):
            for file in files:
                namef = ''
                l = str(os.path.join(root, file)).split('\\')
                is_begin: bool = False
                for i in range(1, len(l)):
                    if is_begin:
                        if i + 1 == len(l):
                            namef += l[i]
                        else:
                            namef += l[i] + '/'
                    if not is_begin and l[i] == DocumentTranslatorToXML.template_directory_name():
                        is_begin = True
                zipf.write(os.path.join(root, file), namef)
        zipf.close()
        if os.path.exists(name):
            os.remove(name)
        os.rename(f'{name}.zip', name)

    def _get_document(self):
        return self
