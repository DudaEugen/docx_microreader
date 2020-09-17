from docx_parser import XMLement
from constants import *


class Style(XMLement):
    tag = Style_tag


class ParagraphStyle(Style):
    type = ParStyle_type


class NumberingStyle(Style):
    type = NumStyle_type


class CharacterStyle(Style):
    type = CharStyle_type


class TableStyle(Style):
    type = TabStyle_type
