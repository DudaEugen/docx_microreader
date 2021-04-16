from .xml_element import XMLement
import xml.etree.ElementTree as ET
from .mixins.getters_setters import ParagraphPropertiesGetSetMixin, RunPropertiesGetSetMixin, TablePropertiesGetSetMixin
from .constants import property_enums as pr_const
from typing import Union, Dict, Optional


class Style(XMLement):
    element_description: pr_const.Style

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        self.id: str = style_id
        self.is_default = is_default
        self.is_custom_style = is_custom_style
        super(Style, self).__init__(element, parent)

    def _get_style_id(self) -> Optional[str]:
        return self._properties[pr_const.StyleProperty.BASE_STYLE.key].value

    @classmethod
    def _set_default_style_of_class(cls):
        return


class ParagraphStyle(Style, ParagraphPropertiesGetSetMixin):
    element_description = pr_const.Style.PARAGRAPH


class NumberingStyle(Style):
    element_description = pr_const.Style.NUMBERING


class CharacterStyle(Style, RunPropertiesGetSetMixin):
    element_description = pr_const.Style.CHARACTER


class TableAreaStyle(XMLement, TablePropertiesGetSetMixin):
    element_description = pr_const.SubStyle.TABLE_AREA

    def __init__(self, element: ET.Element, parent):
        from .constants.namespaces import check_namespace_of_tag

        self.area = pr_const.convert_to_enum_element(
            element.get(check_namespace_of_tag('w:type')),
            pr_const.TableArea
        )
        super(TableAreaStyle, self).__init__(element, parent)

    @classmethod
    def _set_default_style_of_class(cls):
        return


class TableStyle(Style, TablePropertiesGetSetMixin):
    element_description = pr_const.Style.TABLE

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [TableAreaStyle]

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        super(TableStyle, self).__init__(element, parent, style_id, is_default, is_custom_style)
        self.__table_area_styles: Dict[str, TableAreaStyle] = {
            st.area: st for st in self.inner_elements if isinstance(st, TableAreaStyle)
        }

    def get_table_area_style(self, table_area: Union[str, pr_const.TableArea]):
        area = pr_const.convert_to_enum_element(table_area, pr_const.TableArea)
        result = self.__table_area_styles.get(area)
        if result is not None:
            return result
        if self._base_style is not None:
            return self._base_style.get_table_area_style(area)
        return None
