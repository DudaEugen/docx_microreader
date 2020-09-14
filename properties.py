from typing import Union, List, Dict


class PropertyDescription:

    def __init__(self, tag_wrap: str, tag: Union[List[str], str], tag_property: Union[List[str], str, None], is_view: bool):
        self.tag_wrap: str = tag_wrap
        self.tag: Union[List[str], str] = tag
        self.tag_property: Union[List[str], str, None] = tag_property
        self.value_type: str = 'str' if self.tag_property is not None else 'bool'
        self.is_view: bool = is_view

    def get_wrapped_tags(self) -> Union[List[str], str]:
        if isinstance(self.tag, list):
            return [self.tag_wrap + '/' + tag for tag in self.tag]
        return self.tag_wrap + '/' + self.tag


class Property:

    def __init__(self, value: Union[str, None, bool], description: PropertyDescription):
        self.description: PropertyDescription = description
        self.value: Union[str, None, bool] = value

    def is_view_and_not_none(self) -> bool:
        return self.description.is_view and (self.value is not None and self.value is not False)


class XMLementPropertyDescriptions:
    Const_directions: Dict[str, str] = {
        'top': '_top',
        'bottom': '_bottom',
        'left': '_left',
        'right': '_right',
    }
    Const_property_names: Dict[str, str] = {
        'color': '_color',
        'size': '_size',
        'space': '_space',
        'type': '',
    }
    Const_prefixes: Dict[str, str] = {
        'margin': 'margin',
        'cell_margin': 'cell_margin',
        'borders_inside': 'borders_inside',
        'border': 'border',
    }

    Par_align: str = 'align'
    Par_indent_left: str = 'indent_left'
    Par_indent_right: str = 'indent_right'
    Par_hanging: str = 'hanging'
    Par_first_line: str = 'first_line'
    Par_keep_lines: str = 'keep_lines'
    Par_keep_next: str = 'keep_next'
    Par_outline_level: str = 'outline_level'
    Par_border_top: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["type"]}'
    Par_border_top_color: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["color"]}'
    Par_border_top_size: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["size"]}'
    Par_border_top_space: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["space"]}'
    Par_border_bottom: str = rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["type"]}'
    Par_border_bottom_color: str = \
        rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["color"]}'
    Par_border_bottom_size: str = \
        rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["size"]}'
    Par_border_bottom_space: str = \
        rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["space"]}'
    Par_border_right: str = rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["type"]}'
    Par_border_right_color: str = \
        rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["color"]}'
    Par_border_right_size: str = rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["size"]}'
    Par_border_right_space: str = \
        rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["space"]}'
    Par_border_left: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["type"]}'
    Par_border_left_color: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["color"]}'
    Par_border_left_size: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["size"]}'
    Par_border_left_space: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["space"]}'
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
    Run_border: str = rf'{Const_prefixes["border"]}{Const_property_names["type"]}'
    Run_border_color: str = rf'{Const_prefixes["border"]}{Const_property_names["color"]}'
    Run_border_size: str = rf'{Const_prefixes["border"]}{Const_property_names["size"]}'
    Run_border_space: str = rf'{Const_prefixes["border"]}{Const_property_names["space"]}'
    Tab_layout: str = 'layout'
    Tab_width: str = 'width'
    Tab_width_type: str = 'width_type'
    Tab_align: str = 'align'
    Tab_borders_inside_horizontal: str = \
        rf'{Const_prefixes["borders_inside"]}_horizontal{Const_property_names["type"]}'
    Tab_borders_inside_horizontal_color: str = \
        rf'{Const_prefixes["borders_inside"]}_horizontal{Const_property_names["color"]}'
    Tab_borders_inside_horizontal_size: str = \
        rf'{Const_prefixes["borders_inside"]}_horizontal{Const_property_names["size"]}'
    Tab_borders_inside_vertical: str = \
        rf'{Const_prefixes["borders_inside"]}_vertical{Const_property_names["type"]}'
    Tab_borders_inside_vertical_color: str = \
        rf'{Const_prefixes["borders_inside"]}_vertical{Const_property_names["color"]}'
    Tab_borders_inside_vertical_size: str = \
        rf'{Const_prefixes["borders_inside"]}_vertical{Const_property_names["size"]}'
    Tab_border_top: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["type"]}'
    Tab_border_top_color: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["color"]}'
    Tab_border_top_size: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["size"]}'
    Tab_border_bottom: str = rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["type"]}'
    Tab_border_bottom_color: str = \
        rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["color"]}'
    Tab_border_bottom_size: str = \
        rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["size"]}'
    Tab_border_right: str = rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["type"]}'
    Tab_border_right_color: str = \
        rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["color"]}'
    Tab_border_right_size: str = rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["size"]}'
    Tab_border_left: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["type"]}'
    Tab_border_left_color: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["color"]}'
    Tab_border_left_size: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["size"]}'
    Tab_cell_margin_top: str = \
        rf'{Const_prefixes["cell_margin"]}{Const_directions["top"]}{Const_property_names["size"]}'
    Tab_cell_margin_top_type: str = \
        rf'{Const_prefixes["cell_margin"]}{Const_directions["top"]}{Const_property_names["type"]}'
    Tab_cell_margin_bottom: str = \
        rf'{Const_prefixes["cell_margin"]}{Const_directions["bottom"]}{Const_property_names["size"]}'
    Tab_cell_margin_bottom_type: str = \
        rf'{Const_prefixes["cell_margin"]}{Const_directions["bottom"]}{Const_property_names["type"]}'
    Tab_cell_margin_left: str =\
        rf'{Const_prefixes["cell_margin"]}{Const_directions["left"]}{Const_property_names["size"]}'
    Tab_cell_margin_left_type: str = \
        rf'{Const_prefixes["cell_margin"]}{Const_directions["left"]}{Const_property_names["type"]}'
    Tab_cell_margin_right: str = \
        rf'{Const_prefixes["cell_margin"]}{Const_directions["right"]}{Const_property_names["size"]}'
    Tab_cell_margin_right_type: str = \
        rf'{Const_prefixes["cell_margin"]}{Const_directions["right"]}{Const_property_names["type"]}'
    Tab_indentation: str = 'indentation'
    Tab_indentation_type: str = 'indentation_type'
    Row_is_header: str = 'is_header'
    Row_height: str = 'height'
    Row_height_rule: str = 'height_rule'
    Cell_fill_color: str = 'fill_color'
    Cell_fill_theme: str = 'fill_theme'
    Cell_border_top: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["type"]}'
    Cell_border_top_color: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["color"]}'
    Cell_border_top_size: str = rf'{Const_prefixes["border"]}{Const_directions["top"]}{Const_property_names["size"]}'
    Cell_border_bottom: str = rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["type"]}'
    Cell_border_bottom_color: str =\
        rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["color"]}'
    Cell_border_bottom_size: str = \
        rf'{Const_prefixes["border"]}{Const_directions["bottom"]}{Const_property_names["size"]}'
    Cell_border_right: str = rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["type"]}'
    Cell_border_right_color: str = \
        rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["color"]}'
    Cell_border_right_size: str = \
        rf'{Const_prefixes["border"]}{Const_directions["right"]}{Const_property_names["size"]}'
    Cell_border_left: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["type"]}'
    Cell_border_left_color: str = \
        rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["color"]}'
    Cell_border_left_size: str = rf'{Const_prefixes["border"]}{Const_directions["left"]}{Const_property_names["size"]}'
    Cell_width: str = 'width'
    Cell_width_type: str = 'width_type'
    Cell_col_span: str = 'col_span'
    Cell_vertical_merge: str = 'vertical_merge'
    Cell_is_vertical_merge_continue: str = 'is_vertical_merge_continue'
    Cell_vertical_align: str = 'vertical_align'
    Cell_text_direction: str = 'text_direction'
    Cell_margin_top: str = rf'{Const_prefixes["margin"]}{Const_directions["top"]}{Const_property_names["size"]}'
    Cell_margin_top_type: str = rf'{Const_prefixes["margin"]}{Const_directions["top"]}{Const_property_names["type"]}'
    Cell_margin_bottom: str = rf'{Const_prefixes["margin"]}{Const_directions["bottom"]}{Const_property_names["size"]}'
    Cell_margin_bottom_type: str = \
        rf'{Const_prefixes["margin"]}{Const_directions["bottom"]}{Const_property_names["type"]}'
    Cell_margin_left: str = rf'{Const_prefixes["margin"]}{Const_directions["left"]}{Const_property_names["size"]}'
    Cell_margin_left_type: str = rf'{Const_prefixes["margin"]}{Const_directions["left"]}{Const_property_names["type"]}'
    Cell_margin_right: str = rf'{Const_prefixes["margin"]}{Const_directions["right"]}{Const_property_names["size"]}'
    Cell_margin_right_type: str = \
        rf'{Const_prefixes["margin"]}{Const_directions["right"]}{Const_property_names["type"]}'

    paragraph_property_descriptions: Dict[str, PropertyDescription] = {
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

    run_property_descriptions: Dict[str, PropertyDescription] = {
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

    table_property_descriptions: Dict[str, PropertyDescription] = {
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

    @staticmethod
    def get_paragraph_properties_dict() -> Dict[str, PropertyDescription]:
        return XMLementPropertyDescriptions.paragraph_property_descriptions

    @staticmethod
    def get_run_properties_dict() -> Dict[str, PropertyDescription]:
        return XMLementPropertyDescriptions.run_property_descriptions

    @staticmethod
    def get_table_properties_dict() -> Dict[str, PropertyDescription]:
        return XMLementPropertyDescriptions.table_property_descriptions

    @staticmethod
    def get_row_properties_dict() -> Dict[str, PropertyDescription]:
        return XMLementPropertyDescriptions.row_property_descriptions

    @staticmethod
    def get_cell_properties_dict() -> Dict[str, PropertyDescription]:
        return XMLementPropertyDescriptions.cell_property_descriptions

    @staticmethod
    def empty_properties_dict() -> Dict[str, PropertyDescription]:
        return {}

    @staticmethod
    def get_properties_dict(ob) -> Dict[str, PropertyDescription]:
        from models import Paragraph, Table, Document
        if isinstance(ob, Paragraph):
            return XMLementPropertyDescriptions.get_paragraph_properties_dict()
        elif isinstance(ob, Paragraph.Run):
            return XMLementPropertyDescriptions.get_run_properties_dict()
        elif isinstance(ob, Paragraph.Run.Text):
            return XMLementPropertyDescriptions.empty_properties_dict()
        elif isinstance(ob, Table):
            return XMLementPropertyDescriptions.get_table_properties_dict()
        elif isinstance(ob, Table.Row):
            return XMLementPropertyDescriptions.get_row_properties_dict()
        elif isinstance(ob, Table.Row.Cell):
            return XMLementPropertyDescriptions.get_cell_properties_dict()
        elif isinstance(ob, Document) or isinstance(ob, Document.Body):
            return XMLementPropertyDescriptions.empty_properties_dict()

        raise ValueError('argument in create properties dict function is mistake')
