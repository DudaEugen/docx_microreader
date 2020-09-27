from typing import Dict

# tags of classes constants
Par_tag: str = 'w:p'
Run_tag: str = 'w:r'
Text_tag: str = 'w:t'
Tab_tag: str = 'w:tbl'
Row_tag: str = 'w:tr'
Cell_tag: str = 'w:tc'
Body_tag: str = 'w:body'
Style_tag: str = 'w:style'
StyleTableArea_tag: str = 'w:tblStylePr'

# style types constants
ParStyle_type: str = 'paragraph'
NumStyle_type: str = 'numbering'
CharStyle_type: str = 'character'
TabStyle_type: str = 'table'
TabFirsRowStyle_type: str = 'firstRow'
TabLastRowStyle_type: str = 'lastRow'
TabFirstColumnStyle_type: str = 'firstCol'
TabLastColumnStyle_type: str = 'lastCol'
TabOddRowStyle_type: str = 'band1Horz'
TabEvenRowStyle_type: str = 'band2Horz'
TabOddColumnStyle_type: str = 'band1Vert'
TabEvenColumnStyle_type: str = 'band2Vert'
TabTopRightCellStyle_type: str = 'neCell'
TabTopLeftCellStyle_type: str = 'nwCell'
TabBottomRightCellStyle_type: str = 'seCell'
TabBottomLeftCellStyle_type: str = 'swCell'

ParStyle: str = 'paragraph_style'
TabStyle: str = 'table_style'
CharStyle: str = 'character_style'
NumStyle: str = 'numbering_style'

# parameters of style
StyleParam_type: str = 'type'
StyleParam_id: str = 'id'
StyleParam_is_default: str = 'is_default'
StyleParam_is_custom: str = 'is_custom'

# styles properties
StyleBasedOn: str = 'based_on'

# value of bool property
BoolPropertyValue: str = 'w:val'

# properties constants
Const_directions: Dict[str, str] = {
    'top': '_top',
    'bottom': '_bottom',
    'left': '_left',
    'right': '_right',
    'horizontal': '_horizontal',
    'vertical': '_vertical',
}
Const_property_names: Dict[str, str] = {
    'color': '_color',
    'size': '_size',
    'space': '_space',
    'type': '_type',
}
Const_quantities: Dict[str, str] = {
    'margin': 'margin',
    'cell_margin': 'cell_margin',
    'borders_inside': 'borders_inside',
    'border': 'border',
    'paragraph_border': 'paragraph_border',
    'cell_border': 'cell_border',
}


def get_key(quantity: str, direction: str = '', property_name: str = '') -> str:
    """
    :param quantity: key of Const_quantities dict
    :param direction: key of Const_directions dict
    :param property_name: key of Const_property_names dict
    """
    dir: str = Const_directions[direction] if direction in Const_directions else ''
    p_name: str = Const_property_names[property_name] if property_name in Const_property_names else ''
    return rf'{Const_quantities[quantity]}{dir}{p_name}'


Par_align: str = 'paragraph_align'
Par_indent_left: str = 'indent_left'
Par_indent_right: str = 'indent_right'
Par_hanging: str = 'hanging'
Par_first_line: str = 'first_line'
Par_keep_lines: str = 'keep_lines'
Par_keep_next: str = 'keep_next'
Par_outline_level: str = 'outline_level'
Par_border_top: str = get_key('paragraph_border', 'top', 'type')
Par_border_top_color: str = get_key('paragraph_border', 'top', 'color')
Par_border_top_size: str = get_key('paragraph_border', 'top', 'size')
Par_border_top_space: str = get_key('paragraph_border', 'top', 'space')
Par_border_bottom: str = get_key('paragraph_border', 'bottom', 'type')
Par_border_bottom_color: str = get_key('paragraph_border', 'bottom', 'color')
Par_border_bottom_size: str = get_key('paragraph_border', 'bottom', 'size')
Par_border_bottom_space: str = get_key('paragraph_border', 'bottom', 'space')
Par_border_right: str = get_key('paragraph_border', 'right', 'type')
Par_border_right_color: str = get_key('paragraph_border', 'right', 'color')
Par_border_right_size: str = get_key('paragraph_border', 'right', 'size')
Par_border_right_space: str = get_key('paragraph_border', 'right', 'space')
Par_border_left: str = get_key('paragraph_border', 'left', 'type')
Par_border_left_color: str = get_key('paragraph_border', 'left', 'color')
Par_border_left_size: str = get_key('paragraph_border', 'left', 'size')
Par_border_left_space: str = get_key('paragraph_border', 'left', 'space')
Run_size: str = 'size'
Run_is_bold: str = 'is_bold'
Run_is_italic: str = 'is_italic'
Run_vertical_align: str = 'vertical_align'
Run_language: str = 'language'
Run_color: str = 'color'
Run_theme_color: str = 'theme_color'
Run_background_color: str = 'background_color'
Run_background_fill: str = 'background_fill'
Run_underline: str = 'underline'
Run_underline_color: str = 'underline_color'
Run_is_strike: str = 'is_strike'
Run_border: str = get_key('border', property_name='type')
Run_border_color: str = get_key('border', property_name='color')
Run_border_size: str = get_key('border', property_name='size')
Run_border_space: str = get_key('border', property_name='space')
Tab_layout: str = 'layout'
Tab_width: str = 'width'
Tab_width_type: str = 'width_type'
Tab_align: str = 'align'
Tab_borders_inside_horizontal: str = get_key('borders_inside', 'horizontal', 'type')
Tab_borders_inside_horizontal_color: str = get_key('borders_inside', 'horizontal', 'color')
Tab_borders_inside_horizontal_size: str = get_key('borders_inside', 'horizontal', 'size')
Tab_borders_inside_vertical: str = get_key('borders_inside', 'vertical', 'type')
Tab_borders_inside_vertical_color: str = get_key('borders_inside', 'vertical', 'color')
Tab_borders_inside_vertical_size: str = get_key('borders_inside', 'vertical', 'size')
Tab_border_top: str = get_key('border', 'top', 'type')
Tab_border_top_color: str = get_key('border', 'top', 'color')
Tab_border_top_size: str = get_key('border', 'top', 'size')
Tab_border_bottom: str = get_key('border', 'bottom', 'type')
Tab_border_bottom_color: str = get_key('border', 'bottom', 'color')
Tab_border_bottom_size: str = get_key('border', 'bottom', 'size')
Tab_border_right: str = get_key('border', 'right', 'type')
Tab_border_right_color: str = get_key('border', 'right', 'color')
Tab_border_right_size: str = get_key('border', 'right', 'size')
Tab_border_left: str = get_key('border', 'left', 'type')
Tab_border_left_color: str = get_key('border', 'left', 'color')
Tab_border_left_size: str = get_key('border', 'left', 'size')
Tab_cell_margin_top: str = get_key('cell_margin', 'top', 'size')
Tab_cell_margin_top_type: str = get_key('cell_margin', 'top', 'type')
Tab_cell_margin_bottom: str = get_key('cell_margin', 'bottom', 'size')
Tab_cell_margin_bottom_type: str = get_key('cell_margin', 'bottom', 'type')
Tab_cell_margin_left: str = get_key('cell_margin', 'left', 'size')
Tab_cell_margin_left_type: str = get_key('cell_margin', 'left', 'type')
Tab_cell_margin_right: str = get_key('cell_margin', 'right', 'size')
Tab_cell_margin_right_type: str = get_key('cell_margin', 'right', 'type')
Tab_indentation: str = 'indentation'
Tab_indentation_type: str = 'indentation_type'
Row_is_header: str = 'is_header'
Row_height: str = 'height'
Row_height_rule: str = 'height_rule'
Cell_fill_color: str = 'fill_color'
Cell_fill_theme: str = 'fill_theme'
Cell_border_top: str = get_key('cell_border', 'top', 'type')
Cell_border_top_color: str = get_key('cell_border', 'top', 'color')
Cell_border_top_size: str = get_key('cell_border', 'top', 'size')
Cell_border_bottom: str = get_key('cell_border', 'bottom', 'type')
Cell_border_bottom_color: str = get_key('cell_border', 'bottom', 'color')
Cell_border_bottom_size: str = get_key('cell_border', 'bottom', 'size')
Cell_border_right: str = get_key('cell_border', 'right', 'type')
Cell_border_right_color: str = get_key('cell_border', 'right', 'color')
Cell_border_right_size: str = get_key('cell_border', 'right', 'size')
Cell_border_left: str = get_key('cell_border', 'left', 'type')
Cell_border_left_color: str = get_key('cell_border', 'left', 'color')
Cell_border_left_size: str = get_key('cell_border', 'left', 'size')
Cell_width: str = 'cell_width'
Cell_width_type: str = 'cell_width_type'
Cell_col_span: str = 'col_span'
Cell_vertical_merge: str = 'vertical_merge'
Cell_is_vertical_merge_continue: str = 'is_vertical_merge_continue'
Cell_vertical_align: str = 'cell_vertical_align'
Cell_text_direction: str = 'text_direction'
Cell_margin_top: str = get_key('margin', 'top', 'size')
Cell_margin_top_type: str = get_key('margin', 'top', 'type')
Cell_margin_bottom: str = get_key('margin', 'bottom', 'size')
Cell_margin_bottom_type: str = get_key('margin', 'bottom', 'type')
Cell_margin_left: str = get_key('margin', 'left', 'size')
Cell_margin_left_type: str = get_key('margin', 'left', 'type')
Cell_margin_right: str = get_key('margin', 'right', 'size')
Cell_margin_right_type: str = get_key('margin', 'right', 'type')
