from docx_parser import XMLement


class Style(XMLement):
    tag = 'w:style'


class ParagraphStyle(Style):
    type = 'paragraph'


class NumberingStyle(Style):
    type = 'numbering'


class CharacterStyle(Style):
    type = 'character'


class TableStyle(Style):
    type = 'table'
