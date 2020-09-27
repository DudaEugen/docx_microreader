from docx_parser import XMLement
import xml.etree.ElementTree as ET
from mixins.getters_setters import ParagraphPropertiesGetSetMixin, RunPropertiesGetSetMixin, TablePropertiesGetSetMixin
from constants import keys_consts as k_const
from typing import Union, Dict


class Style(XMLement):
    tag = k_const.Style_tag

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        self.id: str = style_id
        self.is_default = is_default
        self.is_custom_style = is_custom_style
        super(Style, self).__init__(element, parent)

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[k_const.StyleBasedOn].value


class ParagraphStyle(Style, ParagraphPropertiesGetSetMixin):
    type = k_const.ParStyle_type


class NumberingStyle(Style):
    type = k_const.NumStyle_type


class CharacterStyle(Style, RunPropertiesGetSetMixin):
    type = k_const.CharStyle_type


class TableStyle(Style, TablePropertiesGetSetMixin):
    type = k_const.TabStyle_type

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        self.first_row: Union[TableStyle.TableAreaStyle, None] = None
        self.last_row: Union[TableStyle.TableAreaStyle, None] = None
        self.first_col: Union[TableStyle.TableAreaStyle, None] = None
        self.last_col: Union[TableStyle.TableAreaStyle, None] = None
        self.odd_col: Union[TableStyle.TableAreaStyle, None] = None
        self.even_col: Union[TableStyle.TableAreaStyle, None] = None
        self.odd_row: Union[TableStyle.TableAreaStyle, None] = None
        self.even_row: Union[TableStyle.TableAreaStyle, None] = None
        self.top_right_cell: Union[TableStyle.TableAreaStyle, None] = None
        self.top_left_cell: Union[TableStyle.TableAreaStyle, None] = None
        self.bottom_right_cell: Union[TableStyle.TableAreaStyle, None] = None
        self.bottom_left_cell: Union[TableStyle.TableAreaStyle, None] = None
        super(TableStyle, self).__init__(element, parent, style_id, is_default, is_custom_style)

    def _init(self):
        self.__table_area_styles: Dict[str, TableStyle.TableAreaStyle] = {
            st.type: st for st in self._get_elements(TableStyle.TableAreaStyle)
        }

    def _parse_element(self, element: ET.Element):
        return TableStyle.TableAreaStyle(element, self)

    def get_table_area_style(self, table_area_style_type: str):
        result = self.__table_area_styles.get(table_area_style_type)
        if result is not None:
            return result
        if self._base_style is not None:
            return self._base_style.get_table_area_style(table_area_style_type)
        return None

    class TableAreaStyle(XMLement, TablePropertiesGetSetMixin):
        tag = k_const.StyleTableArea_tag

        def __init__(self, element: ET.Element, parent):
            self.type: str = ''
            super(TableStyle.TableAreaStyle, self).__init__(element, parent)

        def _init(self):
            self.type = self._element.get(self._check_namespace('w:type'))
