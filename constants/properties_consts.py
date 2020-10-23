from typing import Dict, Tuple
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
paragraph_style_property_description: Dict[str, PropertyDescription] = {
    Par_align: PropertyDescription('w:pPr', 'w:jc', 'w:val'),
    Par_indent_left: PropertyDescription('w:pPr', 'w:ind', ['w:left', 'w:start']),
    Par_indent_right: PropertyDescription('w:pPr', 'w:ind', ['w:right', 'w:end']),
    Par_hanging: PropertyDescription('w:pPr', 'w:ind', 'w:hanging'),
    Par_first_line: PropertyDescription('w:pPr', 'w:ind', 'w:firstLine'),
    Par_keep_lines: PropertyDescription('w:pPr', 'w:keepLines', None),
    Par_keep_next: PropertyDescription('w:pPr', 'w:keepNext', None),
    Par_outline_level: PropertyDescription('w:pPr', 'w:outlineLvl', 'w:val'),
    Par_border_top: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:val'),
    Par_border_top_color: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:color'),
    Par_border_top_size: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:sz'),
    Par_border_top_space: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:space'),
    Par_border_bottom: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:val'),
    Par_border_bottom_color: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:color'),
    Par_border_bottom_size: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:sz'),
    Par_border_bottom_space: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:space'),
    Par_border_right: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:val'),
    Par_border_right_color: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:color'),
    Par_border_right_size: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:sz'),
    Par_border_right_space: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:space'),
    Par_border_left: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:val'),
    Par_border_left_color: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:color'),
    Par_border_left_size: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:sz'),
    Par_border_left_space: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:space'),
}

# description of properties Paragraph, but not ParagraphStyle
paragraph_property_description: Dict[str, PropertyDescription] = {
    ParStyle: PropertyDescription('w:pPr', 'w:pStyle', 'w:val'),
}

# description of properties Run and RunStyle
run_style_property_descriptions: Dict[str, PropertyDescription] = {
    Run_size: PropertyDescription('w:rPr', 'w:sz', 'w:val'),
    Run_is_bold: PropertyDescription('w:rPr', 'w:b', None),
    Run_is_italic: PropertyDescription('w:rPr', 'w:i', None),
    Run_vertical_align: PropertyDescription('w:rPr', 'w:vertAlign', 'w:val'),
    Run_language: PropertyDescription('w:rPr', 'w:lang', 'w:val'),
    Run_color: PropertyDescription('w:rPr', 'w:color', 'w:val'),
    Run_theme_color: PropertyDescription('w:rPr', 'w:color', 'w:themeColor'),
    Run_background_color: PropertyDescription('w:rPr', 'w:highlight', 'w:val'),
    Run_background_fill: PropertyDescription('w:rPr', 'w:shd', 'w:fill'),
    Run_underline: PropertyDescription('w:rPr', 'w:u', 'w:val'),
    Run_underline_color: PropertyDescription('w:rPr', 'w:u', 'w:color'),
    Run_is_strike: PropertyDescription('w:rPr', 'w:strike', None),
    Run_border: PropertyDescription('w:rPr', 'w:bdr', 'w:val'),
    Run_border_color: PropertyDescription('w:rPr', 'w:bdr', 'w:color'),
    Run_border_size: PropertyDescription('w:rPr', 'w:bdr', 'w:sz'),
    Run_border_space: PropertyDescription('w:rPr', 'w:bdr', 'w:space'),
}

# description of properties Run, but not RunStyle
run_property_descriptions: Dict[str, PropertyDescription] = {
    CharStyle: PropertyDescription('w:rPr', 'w:rStyle', 'w:val'),
}

# description of properties Table and TableStyle
table_style_property_descriptions: Dict[str, PropertyDescription] = {
    Tab_layout: PropertyDescription('w:tblPr', 'w:tblLayout', 'w:type'),
    Tab_width: PropertyDescription('w:tblPr', 'w:tblW', 'w:w'),
    Tab_width_type: PropertyDescription('w:tblPr', 'w:tblW', 'w:type'),
    Tab_align: PropertyDescription('w:tblPr', 'w:jc', 'w:val'),
    Tab_borders_inside_horizontal: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                       'w:insideH', 'w:val'),
    Tab_borders_inside_horizontal_color: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                             'w:insideH', 'w:color'),
    Tab_borders_inside_horizontal_size: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                            'w:insideH', 'w:sz'),
    Tab_borders_inside_vertical: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                     'w:insideV', 'w:val'),
    Tab_borders_inside_vertical_color: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                           'w:insideV', 'w:color'),
    Tab_borders_inside_vertical_size: PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'],
                                                          'w:insideV', 'w:sz'),
    Tab_border_top: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:val'),
    Tab_border_top_color: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:color'),
    Tab_border_top_size: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:sz'),
    Tab_border_bottom: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:val'),
    Tab_border_bottom_color: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:color'),
    Tab_border_bottom_size: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:sz'),
    Tab_border_right: PropertyDescription('w:tblPr/w:tblBorders', ['w:right', 'w:end'], 'w:val'),
    Tab_border_right_color: PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:color'),
    Tab_border_right_size: PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:sz'),
    Tab_border_left: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:val'),
    Tab_border_left_color: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:color'),
    Tab_border_left_size: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:sz'),
    Tab_cell_margin_top: PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:w'),
    Tab_cell_margin_top_type: PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:type'),
    Tab_cell_margin_left: PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:w'),
    Tab_cell_margin_left_type: PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:type'),
    Tab_cell_margin_bottom: PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:w'),
    Tab_cell_margin_bottom_type: PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:type'),
    Tab_cell_margin_right: PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:w'),
    Tab_cell_margin_right_type: PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:type'),
    Tab_indentation: PropertyDescription('w:tblPr', 'w:tblInd', ['w:w', 'w:val']),
    Tab_indentation_type: PropertyDescription('w:tblPr', 'w:tblInd', 'w:type'),
}

# description of properties Table, but not TableStyle
table_property_descriptions: Dict[str, PropertyDescription] = {
    TabStyle: PropertyDescription('w:tblPr', 'w:tblStyle', 'w:val'),
    Tab_first_row_style_look: PropertyDescription('w:tblPr', 'w:tblLook', 'w:firstRow'),
    Tab_first_column_style_look: PropertyDescription('w:tblPr', 'w:tblLook', 'w:firstColumn'),
    Tab_last_row_style_look: PropertyDescription('w:tblPr', 'w:tblLook', 'w:lastRow'),
    Tab_last_column_style_look: PropertyDescription('w:tblPr', 'w:tblLook', 'w:lastColumn'),
    Tab_no_horizontal_banding: PropertyDescription('w:tblPr', 'w:tblLook', 'w:noHBand'),
    Tab_no_vertical_banding: PropertyDescription('w:tblPr', 'w:tblLook', 'w:noVBand'),
}

row_property_descriptions: Dict[str, PropertyDescription] = {
    Row_is_header: PropertyDescription('w:trPr', 'w:tblHeader', None),
    Row_height: PropertyDescription('w:trPr', 'w:trHeight', 'w:val'),
    Row_height_rule: PropertyDescription('w:trPr', 'w:trHeight', 'w:hRule'),
}

cell_property_descriptions: Dict[str, PropertyDescription] = {
    Cell_fill_color: PropertyDescription('w:tcPr', 'w:shd', 'w:fill'),
    Cell_fill_theme: PropertyDescription('w:tcPr', 'w:shd', 'w:themeFill'),
    Cell_border_top: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:val'),
    Cell_border_top_color: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:color'),
    Cell_border_top_size: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:sz'),
    Cell_border_bottom: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:val'),
    Cell_border_bottom_color: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:color'),
    Cell_border_bottom_size: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:sz'),
    Cell_border_right: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:val'),
    Cell_border_right_color: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:color'),
    Cell_border_right_size: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:sz'),
    Cell_border_left: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:val'),
    Cell_border_left_color: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:color'),
    Cell_border_left_size: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:sz'),
    Cell_width: PropertyDescription('w:tcPr', 'w:tcW', 'w:w'),
    Cell_width_type: PropertyDescription('w:tcPr', 'w:tcW', 'w:type'),
    Cell_col_span: PropertyDescription('w:tcPr', 'w:gridSpan', 'w:val'),
    Cell_vertical_merge: PropertyDescription('w:tcPr', 'w:vMerge', 'w:val'),
    Cell_is_vertical_merge_continue: PropertyDescription('w:tcPr', 'w:vMerge', None),
    Cell_vertical_align: PropertyDescription('w:tcPr', 'w:vAlign', 'w:val'),
    Cell_text_direction: PropertyDescription('w:tcPr', 'w:textDirection', 'w:val'),
    Cell_margin_top: PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:w'),
    Cell_margin_top_type: PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:type'),
    Cell_margin_bottom: PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:w'),
    Cell_margin_bottom_type: PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:type'),
    Cell_margin_left: PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:w'),
    Cell_margin_left_type: PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:type'),
    Cell_margin_right: PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:w'),
    Cell_margin_right_type: PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:type'),
}

image_property_descriptions: Dict[str, PropertyDescription] = {
    Img_id: PropertyDescription(None, None, 'r:embed'),
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
        return merge_dicts(paragraph_property_description, paragraph_style_property_description)
    elif isinstance(ob, Paragraph.Run):
        return merge_dicts(run_property_descriptions, run_style_property_descriptions)
    elif isinstance(ob, Paragraph.Run.Text):
        return {}
    elif isinstance(ob, Table):
        return merge_dicts(table_property_descriptions, table_style_property_descriptions)
    elif isinstance(ob, Table.Row):
        return row_property_descriptions
    elif isinstance(ob, Table.Row.Cell):
        return cell_property_descriptions
    elif isinstance(ob, Document.Body):
        return {}
    elif isinstance(ob, NumberingStyle):
        return style_properties
    elif isinstance(ob, TableStyle):
        return merge_dicts(table_style_property_descriptions, style_properties,
                           row_property_descriptions, cell_property_descriptions,
                           paragraph_style_property_description, run_style_property_descriptions)
    elif isinstance(ob, TableStyle.TableAreaStyle):
        return merge_dicts(table_style_property_descriptions, row_property_descriptions, cell_property_descriptions,
                           paragraph_style_property_description, run_style_property_descriptions)
    elif isinstance(ob, CharacterStyle):
        return merge_dicts(run_style_property_descriptions, style_properties)
    elif isinstance(ob, ParagraphStyle):
        return merge_dicts(paragraph_style_property_description, style_properties, run_style_property_descriptions)
    elif isinstance(ob, Drawing):
        return {}
    elif isinstance(ob, Drawing.Image):
        return image_property_descriptions

    raise ValueError('argument in create properties dict function is mistake')
