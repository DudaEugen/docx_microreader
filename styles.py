from docx_parser import XMLement
import xml.etree.ElementTree as ET
from mixins.getters_setters import ParagraphPropertiesGetSetMixin, RunPropertiesGetSetMixin, TablePropertiesGetSetMixin
from constants import keys_consts as k_const


class Style(XMLement):
    tag = k_const.Style_tag

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        self.id: str = style_id
        self.is_default = is_default
        self.is_custom_style = is_custom_style
        super(Style, self).__init__(element, parent)


class ParagraphStyle(Style, ParagraphPropertiesGetSetMixin):
    type = k_const.ParStyle_type


class NumberingStyle(Style):
    type = k_const.NumStyle_type


class CharacterStyle(Style, RunPropertiesGetSetMixin):
    type = k_const.CharStyle_type


class TableStyle(Style, TablePropertiesGetSetMixin):
    type = k_const.TabStyle_type
