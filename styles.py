from docx_parser import XMLement
import xml.etree.ElementTree as ET
from constants.properties_consts import *
from mixins.getters_setters import *


class Style(XMLement):
    tag = Style_tag

    def __init__(self, element: ET.Element, parent,
                 style_id: str, is_default: bool = False, is_custom_style: bool = False):
        self.id: str = style_id
        self.is_default = is_default
        self.is_custom_style = is_custom_style
        super(Style, self).__init__(element, parent)


class ParagraphStyle(Style, ParagraphPropertiesGetSetMixin):
    type = ParStyle_type


class NumberingStyle(Style):
    type = NumStyle_type


class CharacterStyle(Style, RunPropertiesGetSetMixin):
    type = CharStyle_type


class TableStyle(Style, TablePropertiesGetSetMixin):
    type = TabStyle_type
