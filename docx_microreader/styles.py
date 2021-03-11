from .docx_parser import XMLement
import xml.etree.ElementTree as ET
from .mixins.getters_setters import ParagraphPropertiesGetSetMixin, RunPropertiesGetSetMixin, TablePropertiesGetSetMixin
from .constants import property_enums as pr_const
from typing import Union, Dict


class Style(XMLement):
    element_description: pr_const.Style

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        self.id: str = style_id
        self.is_default = is_default
        self.is_custom_style = is_custom_style
        super(Style, self).__init__(element, parent)

    def _get_style_id(self) -> Union[str, None]:
        return self._properties[pr_const.StyleProperty.BASE_STYLE.key].value


class ParagraphStyle(Style, ParagraphPropertiesGetSetMixin):
    element_description = pr_const.Style.PARAGRAPH


class NumberingStyle(Style):
    element_description = pr_const.Style.NUMBERING


class CharacterStyle(Style, RunPropertiesGetSetMixin):
    element_description = pr_const.Style.CHARACTER


class TableStyle(Style, TablePropertiesGetSetMixin):
    element_description = pr_const.Style.TABLE

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        super(TableStyle, self).__init__(element, parent, style_id, is_default, is_custom_style)

    def _init(self):
        self.__table_area_styles: Dict[str, TableStyle.TableAreaStyle] = {
            st.area: st for st in self._get_elements(TableStyle.TableAreaStyle)
        }

    def _parse_element(self, element: ET.Element):
        return TableStyle.TableAreaStyle(element, self)

    def get_table_area_style(self, table_area: Union[str, pr_const.TableArea]):
        area = pr_const.convert_to_enum_element(table_area, pr_const.TableArea)
        result = self.__table_area_styles.get(area)
        if result is not None:
            return result
        if self._base_style is not None:
            return self._base_style.get_table_area_style(area)
        return None

    class TableAreaStyle(XMLement, TablePropertiesGetSetMixin):
        element_description = pr_const.SubStyle.TABLE_AREA

        def __init__(self, element: ET.Element, parent):
            self.area: pr_const.TableArea
            super(TableStyle.TableAreaStyle, self).__init__(element, parent)

        def _init(self):
            from .constants.namespaces import check_namespace_of_tag
            self.area = pr_const.convert_to_enum_element(
                self._element.get(check_namespace_of_tag('w:type')),
                pr_const.TableArea
            )
