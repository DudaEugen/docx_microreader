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
    StyleBasedOn: PropertyDescription(None, 'w:basedOn', 'w:val', True),
}

# description of properties Paragraph and ParagraphStyle
paragraph_style_property_description: Dict[str, PropertyDescription] = {
    Par_align: PropertyDescription('w:pPr', 'w:jc', 'w:val', True),
    Par_indent_left: PropertyDescription('w:pPr', 'w:ind', ['w:left', 'w:start'], True),
    Par_indent_right: PropertyDescription('w:pPr', 'w:ind', ['w:right', 'w:end'], True),
    Par_hanging: PropertyDescription('w:pPr', 'w:ind', 'w:hanging', True),
    Par_first_line: PropertyDescription('w:pPr', 'w:ind', 'w:firstLine', True),
    Par_keep_lines: PropertyDescription('w:pPr', 'w:keepLines', None, True),
    Par_keep_next: PropertyDescription('w:pPr', 'w:keepNext', None, True),
    Par_outline_level: PropertyDescription('w:pPr', 'w:outlineLvl', 'w:val', True),
    Par_border_top: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:val', True),
    Par_border_top_color: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:color', True),
    Par_border_top_size: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:sz', True),
    Par_border_top_space: PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:space', True),
    Par_border_bottom: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:val', True),
    Par_border_bottom_color: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:color', True),
    Par_border_bottom_size: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:sz', True),
    Par_border_bottom_space: PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:space', True),
    Par_border_right: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:val', True),
    Par_border_right_color: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:color', True),
    Par_border_right_size: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:sz', True),
    Par_border_right_space: PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:space', True),
    Par_border_left: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:val', True),
    Par_border_left_color: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:color', True),
    Par_border_left_size: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:sz', True),
    Par_border_left_space: PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:space', True),
}

# description of properties Paragraph, but not ParagraphStyle
paragraph_property_description: Dict[str, PropertyDescription] = {
    ParStyle: PropertyDescription('w:pPr', 'w:pStyle', 'w:val', True),
}

# description of properties Run and RunStyle
run_style_property_descriptions: Dict[str, PropertyDescription] = {
    Run_size: PropertyDescription('w:rPr', 'w:sz', 'w:val', True),
    Run_is_bold: PropertyDescription('w:rPr', 'w:b', None, True),
    Run_is_italic: PropertyDescription('w:rPr', 'w:i', None, True),
    Run_vertical_align: PropertyDescription('w:rPr', 'w:vertAlign', 'w:val', True),
    Run_language: PropertyDescription('w:rPr', 'w:lang', 'w:val', True),
    Run_color: PropertyDescription('w:rPr', 'w:color', 'w:val', True),
    Run_theme_color: PropertyDescription('w:rPr', 'w:color', 'w:themeColor', True),
    Run_background_color: PropertyDescription('w:rPr', 'w:highlight', 'w:val', True),
    Run_background_fill: PropertyDescription('w:rPr', 'w:shd', 'w:fill', True),
    Run_underline: PropertyDescription('w:rPr', 'w:u', 'w:val', True),
    Run_underline_color: PropertyDescription('w:rPr', 'w:u', 'w:color', True),
    Run_is_strike: PropertyDescription('w:rPr', 'w:strike', None, True),
    Run_border: PropertyDescription('w:rPr', 'w:bdr', 'w:val', True),
    Run_border_color: PropertyDescription('w:rPr', 'w:bdr', 'w:color', True),
    Run_border_size: PropertyDescription('w:rPr', 'w:bdr', 'w:sz', True),
    Run_border_space: PropertyDescription('w:rPr', 'w:bdr', 'w:space', True),
}

# description of properties Run, but not RunStyle
run_property_descriptions: Dict[str, PropertyDescription] = {
    CharStyle: PropertyDescription('w:rPr', 'w:rStyle', 'w:val', True),
}

# description of properties Table and TableStyle
table_style_property_descriptions: Dict[str, PropertyDescription] = {
    Tab_layout: PropertyDescription('w:tblPr', 'w:tblLayout', 'w:type', True),
    Tab_width: PropertyDescription('w:tblPr', 'w:tblW', 'w:w', True),
    Tab_width_type: PropertyDescription('w:tblPr', 'w:tblW', 'w:type', True),
    Tab_align: PropertyDescription('w:tblPr', 'w:jc', 'w:val', True),
    Tab_borders_inside_horizontal: PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:val', True),
    Tab_borders_inside_horizontal_color: PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:color', True),
    Tab_borders_inside_horizontal_size: PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:sz', True),
    Tab_borders_inside_vertical: PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:val', True),
    Tab_borders_inside_vertical_color: PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:color', True),
    Tab_borders_inside_vertical_size: PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:sz', True),
    Tab_border_top: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:val', True),
    Tab_border_top_color: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:color', True),
    Tab_border_top_size: PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:sz', True),
    Tab_border_bottom: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:val', True),
    Tab_border_bottom_color: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:color', True),
    Tab_border_bottom_size: PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:sz', True),
    Tab_border_right: PropertyDescription('w:tblPr/w:tblBorders', ['w:right', 'w:end'], 'w:val', True),
    Tab_border_right_color: PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:color', True),
    Tab_border_right_size: PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:sz', True),
    Tab_border_left: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:val', True),
    Tab_border_left_color: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:color', True),
    Tab_border_left_size: PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:sz', True),
    Tab_cell_margin_top: PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:w', True),
    Tab_cell_margin_top_type: PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:type', True),
    Tab_cell_margin_left: PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:w', True),
    Tab_cell_margin_left_type: PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:type', True),
    Tab_cell_margin_bottom: PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:w', True),
    Tab_cell_margin_bottom_type: PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:type', True),
    Tab_cell_margin_right: PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:w', True),
    Tab_cell_margin_right_type: PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:type', True),
    Tab_indentation: PropertyDescription('w:tblPr', 'w:tblInd', ['w:w', 'w:val'], True),
    Tab_indentation_type: PropertyDescription('w:tblPr', 'w:tblInd', 'w:type', True),
}

# description of properties Table, but not TableStyle
table_property_descriptions: Dict[str, PropertyDescription] = {
    TabStyle: PropertyDescription('w:tblPr', 'w:tblStyle', 'w:val', True),
}

row_property_descriptions: Dict[str, PropertyDescription] = {
    Row_is_header: PropertyDescription('w:trPr', 'w:tblHeader', None, True),
    Row_height: PropertyDescription('w:trPr', 'w:trHeight', 'w:val', True),
    Row_height_rule: PropertyDescription('w:trPr', 'w:trHeight', 'w:hRule', True),
}

cell_property_descriptions: Dict[str, PropertyDescription] = {
    Cell_fill_color: PropertyDescription('w:tcPr', 'w:shd', 'w:fill', True),
    Cell_fill_theme: PropertyDescription('w:tcPr', 'w:shd', 'w:themeFill', True),
    Cell_border_top: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:val', True),
    Cell_border_top_color: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:color', True),
    Cell_border_top_size: PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:sz', True),
    Cell_border_bottom: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:val', True),
    Cell_border_bottom_color: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:color', True),
    Cell_border_bottom_size: PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:sz', True),
    Cell_border_right: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:val', True),
    Cell_border_right_color: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:color', True),
    Cell_border_right_size: PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:sz', True),
    Cell_border_left: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:val', True),
    Cell_border_left_color: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:color', True),
    Cell_border_left_size: PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:sz', True),
    Cell_width: PropertyDescription('w:tcPr', 'w:tcW', 'w:w', True),
    Cell_width_type: PropertyDescription('w:tcPr', 'w:tcW', 'w:type', True),
    Cell_col_span: PropertyDescription('w:tcPr', 'w:gridSpan', 'w:val', True),
    Cell_vertical_merge: PropertyDescription('w:tcPr', 'w:vMerge', 'w:val', True),
    Cell_is_vertical_merge_continue: PropertyDescription('w:tcPr', 'w:vMerge', None, True),
    Cell_vertical_align: PropertyDescription('w:tcPr', 'w:vAlign', 'w:val', True),
    Cell_text_direction: PropertyDescription('w:tcPr', 'w:textDirection', 'w:val', True),
    Cell_margin_top: PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:w', True),
    Cell_margin_top_type: PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:type', True),
    Cell_margin_bottom: PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:w', True),
    Cell_margin_bottom_type: PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:type', True),
    Cell_margin_left: PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:w', True),
    Cell_margin_left_type: PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:type', True),
    Cell_margin_right: PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:w', True),
    Cell_margin_right_type: PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:type', True)
}


def merge_dicts(*args) -> dict:
    result = {}
    for arg in args:
        if len(set(arg.keys()) & set(result.keys())) > 0:
            raise KeyError(rf'keys_consts have same keys: {set(arg.keys()) & set(result.keys())}')
        result.update(arg)
    return result


def get_properties_dict(ob) -> Dict[str, PropertyDescription]:
    from models import Paragraph, Table, Document
    from models import ParagraphStyle, CharacterStyle, TableStyle, NumberingStyle

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

    raise ValueError('argument in create properties dict function is mistake')
