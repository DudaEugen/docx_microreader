from typing import Dict
from properties import PropertyDescription
from constants.keys_consts import *

Style_parameters: Dict[str, str] = {
    StyleParam_type: 'w:type',
    StyleParam_id: 'w:styleId',
    StyleParam_is_default: 'w:default',
    StyleParam_is_custom: 'w:customStyle',
}

# description of properties of styles
style_properties: Dict[str, PropertyDescription] = {
    StyleBasedOn: PropertyDescription(None, 'w:basedOn', 'w:val'),
}

# description of properties Paragraph and ParagraphStyle
PARAGRAPH_STYLE_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    PARAGRAPH_ALIGN: PropertyDescription('w:pPr', 'w:jc', 'w:val'),
    PARAGRAPH_INDENT_LEFT: PropertyDescription('w:pPr', 'w:ind', ['w:left', 'w:start']),
    PARAGRAPH_INDENT_RIGHT: PropertyDescription('w:pPr', 'w:ind', ['w:right', 'w:end']),
    PARAGRAPH_HANGING: PropertyDescription('w:pPr', 'w:ind', 'w:hanging'),
    PARAGRAPH_FIRST_LINE: PropertyDescription('w:pPr', 'w:ind', 'w:firstLine'),
    PARAGRAPH_KEEP_LINES: PropertyDescription('w:pPr', 'w:keepLines', None),
    PARAGRAPH_KEEP_NEXT: PropertyDescription('w:pPr', 'w:keepNext', None),
    PARAGRAPH_OUTLINE_LEVEL: PropertyDescription('w:pPr', 'w:outlineLvl', 'w:val'),
    PARAGRAPH_BORDER_TOP_TYPE: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:val'),
    PARAGRAPH_BORDER_TOP_COLOR: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:color'),
    PARAGRAPH_BORDER_TOP_SIZE: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:sz'),
    PARAGRAPH_BORDER_TOP_SPACE: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:space'),
    PARAGRAPH_BORDER_BOTTOM_TYPE: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:val'),
    PARAGRAPH_BORDER_BOTTOM_COLOR: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:color'),
    PARAGRAPH_BORDER_BOTTOM_SIZE: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:sz'),
    PARAGRAPH_BORDER_BOTTOM_SPACE: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:space'),
    PARAGRAPH_BORDER_RIGHT_TYPE: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:val'),
    PARAGRAPH_BORDER_RIGHT_COLOR: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:color'),
    PARAGRAPH_BORDER_RIGHT_SIZE: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:sz'),
    PARAGRAPH_BORDER_RIGHT_SPACE: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:space'),
    PARAGRAPH_BORDER_LEFT_TYPE: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:val'),
    PARAGRAPH_BORDER_LEFT_COLOR: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:color'),
    PARAGRAPH_BORDER_LEFT_SIZE: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:sz'),
    PARAGRAPH_BORDER_LEFT_SPACE: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:space'),
}

# description of properties Paragraph, but not ParagraphStyle
paragraph_property_description: Dict[str, PropertyDescription] = {
    ParStyle: PropertyDescription('w:pPr', 'w:pStyle', 'w:val'),
}

# description of properties Run and RunStyle
RUN_STYLE_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    RUN_SIZE: PropertyDescription('w:rPr', 'w:sz', 'w:val'),
    RUN_IS_BOLD: PropertyDescription('w:rPr', 'w:b', None),
    RUN_IS_ITALIC: PropertyDescription('w:rPr', 'w:i', None),
    RUN_VERTICAL_ALIGN: PropertyDescription('w:rPr', 'w:vertAlign', 'w:val'),
    RUN_LANGUAGE: PropertyDescription('w:rPr', 'w:lang', 'w:val'),
    RUN_COLOR: PropertyDescription('w:rPr', 'w:color', 'w:val'),
    RUN_THEME_COLOR: PropertyDescription('w:rPr', 'w:color', 'w:themeColor'),
    RUN_BACKGROUND_COLOR: PropertyDescription('w:rPr', 'w:highlight', 'w:val'),
    RUN_BACKGROUND_FILL: PropertyDescription('w:rPr', 'w:shd', 'w:fill'),
    RUN_UNDERLINE_TYPE: PropertyDescription('w:rPr', 'w:u', 'w:val'),
    RUN_UNDERLINE_COLOR: PropertyDescription('w:rPr', 'w:u', 'w:color'),
    RUN_IS_STRIKE: PropertyDescription('w:rPr', 'w:strike', None),
    RUN_BORDER_TYPE: PropertyDescription('w:rPr', 'w:bdr', 'w:val'),
    RUN_BORDER_COLOR: PropertyDescription('w:rPr', 'w:bdr', 'w:color'),
    RUN_BORDER_SIZE: PropertyDescription('w:rPr', 'w:bdr', 'w:sz'),
    RUN_BORDER_SPACE: PropertyDescription('w:rPr', 'w:bdr', 'w:space'),
}

# description of properties Run, but not RunStyle
run_property_descriptions: Dict[str, PropertyDescription] = {
    CharStyle: PropertyDescription('w:rPr', 'w:rStyle', 'w:val'),
}

# description of properties Table and TableStyle
TABLE_STYLE_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    TABLE_LAYOUT: PropertyDescription('w:tblPr', 'w:tblLayout', 'w:type'),
    TABLE_WIDTH: PropertyDescription('w:tblPr', 'w:tblW', 'w:w'),
    TABLE_WIDTH_TYPE: PropertyDescription('w:tblPr', 'w:tblW', 'w:type'),
    TABLE_ALIGN: PropertyDescription('w:tblPr', 'w:jc', 'w:val'),
    TABLE_INSIDE_BORDER_HORIZONTAL_TYPE: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                            'w:insideH', 'w:val'),
    TABLE_INSIDE_BORDER_HORIZONTAL_COLOR: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                              'w:insideH', 'w:color'),
    TABLE_INSIDE_BORDER_HORIZONTAL_SIZE: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                             'w:insideH', 'w:sz'),
    TABLE_INSIDE_BORDER_VERTICAL_TYPE: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideV',
                                                           'w:val'),
    TABLE_INSIDE_BORDER_VERTICAL_COLOR: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideV',
                                                            'w:color'),
    TABLE_INSIDE_BORDER_VERTICAL_SIZE: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideV',
                                                           'w:sz'),
    TABLE_BORDER_TOP_TYPE: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:val'),
    TABLE_BORDER_TOP_COLOR: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:color'),
    TABLE_BORDER_TOP_SIZE: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:sz'),
    TABLE_BORDER_BOTTOM_TYPE: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:val'),
    TABLE_BORDER_BOTTOM_COLOR: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:color'),
    TABLE_BORDER_BOTTOM_SIZE: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:sz'),
    TABLE_BORDER_RIGHT_TYPE: PropertyDescription('w:tblPr/w:tblBorders', ['w:right', 'w:end'], 'w:val'),
    TABLE_BORDER_RIGHT_COLOR: PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:color'),
    TABLE_BORDER_RIGHT_SIZE: PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:sz'),
    TABLE_BORDER_LEFT_TYPE: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:val'),
    TABLE_BORDER_LEFT_COLOR: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:color'),
    TABLE_BORDER_LEFT_SIZE: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:sz'),
    TABLE_CELL_MARGIN_TOP_SIZE: PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:w'),
    TABLE_CELL_MARGIN_TOP_TYPE: PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:type'),
    TABLE_CELL_MARGIN_LEFT_SIZE: PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:w'),
    TABLE_CELL_MARGIN_LEFT_TYPE: PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:type'),
    TABLE_CELL_MARGIN_BOTTOM_SIZE: PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:w'),
    TABLE_CELL_MARGIN_BOTTOM_TYPE: PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:type'),
    TABLE_CELL_MARGIN_RIGHT_SIZE: PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:w'),
    TABLE_CELL_MARGIN_RIGHT_TYPE: PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:type'),
    TABLE_INDENTATION: PropertyDescription('w:tblPr', 'w:tblInd', ['w:w', 'w:val']),
    TABLE_INDENTATION_TYPE: PropertyDescription('w:tblPr', 'w:tblInd', 'w:type'),
}

# description of properties Table, but not TableStyle
TABLE_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    TabStyle: PropertyDescription('w:tblPr', 'w:tblStyle', 'w:val'),
    TABLE_FIRST_ROW_STYLE_LOOK: PropertyDescription('w:tblPr', 'w:tblLook', 'w:firstRow'),
    TABLE_FIRST_COLUMN_STYLE_LOOK: PropertyDescription('w:tblPr', 'w:tblLook', 'w:firstColumn'),
    TABLE_LAST_ROW_STYLE_LOOK: PropertyDescription('w:tblPr', 'w:tblLook', 'w:lastRow'),
    TABLE_LAST_COLUMN_STYLE_LOOK: PropertyDescription('w:tblPr', 'w:tblLook', 'w:lastColumn'),
    TABLE_NO_HORIZONTAL_BANDING: PropertyDescription('w:tblPr', 'w:tblLook', 'w:noHBand'),
    TABLE_NO_VERTICAL_BANDING: PropertyDescription('w:tblPr', 'w:tblLook', 'w:noVBand'),
}

ROW_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    ROW_IS_HEADER: PropertyDescription('w:trPr', 'w:tblHeader', None),
    ROW_HEIGHT: PropertyDescription('w:trPr', 'w:trHeight', 'w:val'),
    ROW_HEIGHT_RULE: PropertyDescription('w:trPr', 'w:trHeight', 'w:hRule'),
}

CELL_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    CELL_FILL_COLOR: PropertyDescription('w:tcPr', 'w:shd', 'w:fill'),
    CELL_FILL_THEME: PropertyDescription('w:tcPr', 'w:shd', 'w:themeFill'),
    CELL_BORDER_TOP_TYPE: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:val'),
    CELL_BORDER_TOP_COLOR: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:color'),
    CELL_BORDER_TOP_SIZE: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:sz'),
    CELL_BORDER_BOTTOM_TYPE: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:val'),
    CELL_BORDER_BOTTOM_COLOR: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:color'),
    CELL_BORDER_BOTTOM_SIZE: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:sz'),
    CELL_BORDER_RIGHT_TYPE: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:val'),
    CELL_BORDER_RIGHT_COLOR: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:color'),
    CELL_BORDER_RIGHT_SIZE: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:sz'),
    CELL_BORDER_LEFT_TYPE: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:val'),
    CELL_BORDER_LEFT_COLOR: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:color'),
    CELL_BORDER_LEFT_SIZE: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:sz'),
    CELL_WIDTH: PropertyDescription('w:tcPr', 'w:tcW', 'w:w'),
    CELL_WIDTH_TYPE: PropertyDescription('w:tcPr', 'w:tcW', 'w:type'),
    CELL_COLUMN_SPAN: PropertyDescription('w:tcPr', 'w:gridSpan', 'w:val'),
    CELL_VERTICAL_MARGE: PropertyDescription('w:tcPr', 'w:vMerge', 'w:val'),
    CELL_IS_VERTICAL_MARGE_CONTINUE: PropertyDescription('w:tcPr', 'w:vMerge', None),
    CELL_VERTICAL_ALIGN: PropertyDescription('w:tcPr', 'w:vAlign', 'w:val'),
    CELL_TEXT_DIRECTION: PropertyDescription('w:tcPr', 'w:textDirection', 'w:val'),
    CELL_MARGIN_TOP_SIZE: PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:w'),
    CELL_MARGIN_TOP_TYPE: PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:type'),
    CELL_MARGIN_BOTTOM_SIZE: PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:w'),
    CELL_MARGIN_BOTTOM_TYPE: PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:type'),
    CELL_MARGIN_RIGHT_SIZE: PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:w'),
    CELL_MARGIN_RIGHT_TYPE: PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:type'),
    CELL_MARGIN_LEFT_SIZE: PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:w'),
    CELL_MARGIN_LEFT_TYPE: PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:type'),
}

DRAWING_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    DRAWING_SIZE_HORIZONTAL: PropertyDescription('wp:inline', 'wp:extent', 'cx'),
    DRAWING_SIZE_VERTICAL: PropertyDescription('wp:inline', 'wp:extent', 'cy'),
}

IMAGE_PROPERTY_DESCRIPTIONS: Dict[str, PropertyDescription] = {
    IMAGE_ID: PropertyDescription(None, None, 'r:embed'),
}


def merge_dicts(*args) -> dict:
    result = {}
    for arg in args:
        if len(set(arg.keys()) & set(result.keys())) > 0:
            raise KeyError(rf'keys_consts have same keys: {set(arg.keys()) & set(result.keys())}')
        result.update(arg)
    return result


def get_properties_dict(ob) -> Dict[str, PropertyDescription]:
    from models import Paragraph, Table, Document, Drawing
    from styles import ParagraphStyle, CharacterStyle, TableStyle, NumberingStyle

    if isinstance(ob, Paragraph):
        return merge_dicts(paragraph_property_description, PARAGRAPH_STYLE_PROPERTY_DESCRIPTIONS)
    elif isinstance(ob, Paragraph.Run):
        return merge_dicts(run_property_descriptions, RUN_STYLE_PROPERTY_DESCRIPTIONS)
    elif isinstance(ob, Paragraph.Run.Text):
        return {}
    elif isinstance(ob, Table):
        return merge_dicts(TABLE_PROPERTY_DESCRIPTIONS, TABLE_STYLE_PROPERTY_DESCRIPTIONS)
    elif isinstance(ob, Table.Row):
        return ROW_PROPERTY_DESCRIPTIONS
    elif isinstance(ob, Table.Row.Cell):
        return CELL_PROPERTY_DESCRIPTIONS
    elif isinstance(ob, Document.Body):
        return {}
    elif isinstance(ob, NumberingStyle):
        return style_properties
    elif isinstance(ob, TableStyle):
        return merge_dicts(TABLE_STYLE_PROPERTY_DESCRIPTIONS, style_properties,
                           ROW_PROPERTY_DESCRIPTIONS, CELL_PROPERTY_DESCRIPTIONS,
                           PARAGRAPH_STYLE_PROPERTY_DESCRIPTIONS, RUN_STYLE_PROPERTY_DESCRIPTIONS)
    elif isinstance(ob, TableStyle.TableAreaStyle):
        return merge_dicts(TABLE_STYLE_PROPERTY_DESCRIPTIONS, ROW_PROPERTY_DESCRIPTIONS, CELL_PROPERTY_DESCRIPTIONS,
                           PARAGRAPH_STYLE_PROPERTY_DESCRIPTIONS, RUN_STYLE_PROPERTY_DESCRIPTIONS)
    elif isinstance(ob, CharacterStyle):
        return merge_dicts(RUN_STYLE_PROPERTY_DESCRIPTIONS, style_properties)
    elif isinstance(ob, ParagraphStyle):
        return merge_dicts(PARAGRAPH_STYLE_PROPERTY_DESCRIPTIONS, style_properties, RUN_STYLE_PROPERTY_DESCRIPTIONS)
    elif isinstance(ob, Drawing):
        return DRAWING_PROPERTY_DESCRIPTIONS
    elif isinstance(ob, Drawing.Image):
        return IMAGE_PROPERTY_DESCRIPTIONS

    raise ValueError('argument in create properties dict function is mistake')
