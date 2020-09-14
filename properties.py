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
    paragraph_property_descriptions: Dict[str, PropertyDescription] = {
        'align': PropertyDescription('w:pPr', 'w:jc', 'w:val', True),
        'indent_left': PropertyDescription('w:pPr', 'w:ind', ['w:left', 'w:start'], True),
        'indent_right': PropertyDescription('w:pPr', 'w:ind', ['w:right', 'w:end'], True),
        'hanging': PropertyDescription('w:pPr', 'w:ind', 'w:hanging', True),
        'first_line': PropertyDescription('w:pPr', 'w:ind', 'w:firstLine', True),
        'keep_lines': PropertyDescription('w:pPr', 'w:keepLines', None, True),
        'keep_next': PropertyDescription('w:pPr', 'w:keepNext', None, True),
        'outline_level': PropertyDescription('w:pPr', 'w:outlineLvl', 'w:val', True),
        'border_top': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:val', True),
        'border_top_color': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:color', True),
        'border_top_size': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:sz', True),
        'border_top_space': PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:space', True),
        'border_bottom': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:val', True),
        'border_bottom_color': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:color', True),
        'border_bottom_size': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:sz', True),
        'border_bottom_space': PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:space', True),
        'border_right': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:val', True),
        'border_right_color': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:color', True),
        'border_right_size': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:sz', True),
        'border_right_space': PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:space', True),
        'border_left': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:val', True),
        'border_left_color': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:color', True),
        'border_left_size': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:sz', True),
        'border_left_space': PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:space', True),
    }

    run_property_descriptions: Dict[str, PropertyDescription] = {
            'size': PropertyDescription('w:rPr', 'w:sz', 'w:val', True),
            'is_bold': PropertyDescription('w:rPr', 'w:b', None, True),
            'is_italic': PropertyDescription('w:rPr', 'w:i', None, True),
            'vertical_align': PropertyDescription('w:rPr', 'w:vertAlign', 'w:val', True),
            'language': PropertyDescription('w:rPr', 'w:lang', 'w:val', True),
            'color': PropertyDescription('w:rPr', 'w:color', 'w:val', True),
            'theme_color': PropertyDescription('w:rPr', 'w:color', 'w:themeColor', True),
            'background_color': PropertyDescription('w:rPr', 'w:highlight', 'w:val', True),
            'background_fill': PropertyDescription('w:rPr', 'w:shd', 'w:fill', True),
            'underline': PropertyDescription('w:rPr', 'w:u', 'w:val', True),
            'underline_color': PropertyDescription('w:rPr', 'w:u', 'w:color', True),
            'is_strike': PropertyDescription('w:rPr', 'w:strike', None, True),
            'border': PropertyDescription('w:rPr', 'w:bdr', 'w:val', True),
            'border_color': PropertyDescription('w:rPr', 'w:bdr', 'w:color', True),
            'border_size': PropertyDescription('w:rPr', 'w:bdr', 'w:sz', True),
            'border_space': PropertyDescription('w:rPr', 'w:bdr', 'w:space', True),
        }

    table_property_descriptions: Dict[str, PropertyDescription] = {
        'layout': PropertyDescription('w:tblPr', 'w:tblLayout', 'w:type', True),
        'width': PropertyDescription('w:tblPr', 'w:tblW', 'w:w', True),
        'width_type': PropertyDescription('w:tblPr', 'w:tblW', 'w:type', True),
        'align': PropertyDescription('w:tblPr', 'w:jc', 'w:val', True),
        'borders_inside_horizontal': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:val', True),
        'borders_inside_horizontal_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:color', True),
        'borders_inside_horizontal_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideH', 'w:sz', True),
        'borders_inside_vertical': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:val', True),
        'borders_inside_vertical_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:color', True),
        'borders_inside_vertical_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:insideV', 'w:sz', True),
        'border_top': PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:val', True),
        'border_top_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:color', True),
        'border_top_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:sz', True),
        'border_bottom': PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:val', True),
        'border_bottom_color': PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:color', True),
        'border_bottom_size': PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:sz', True),
        'border_right': PropertyDescription('w:tblPr/w:tblBorders', ['w:right', 'w:end'], 'w:val', True),
        'border_right_color': PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:color', True),
        'border_right_size': PropertyDescription('w:tblPr/w:tblBorders',  ['w:right', 'w:end'], 'w:sz', True),
        'border_left': PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:val', True),
        'border_left_color': PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:color', True),
        'border_left_size': PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:sz', True),
        'cell_margin_top': PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:w', True),
        'cell_margin_top_type': PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:type', True),
        'cell_margin_left': PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:w', True),
        'cell_margin_left_type': PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:type', True),
        'cell_margin_bottom': PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:w', True),
        'cell_margin_bottom_type': PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:type', True),
        'cell_margin_right': PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:w', True),
        'cell_margin_right_type': PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:type', True),
        'indentation': PropertyDescription('w:tblPr', 'w:tblInd', ['w:w', 'w:val'], True),
        'indentation_type': PropertyDescription('w:tblPr', 'w:tblInd', 'w:type', True),
    }

    row_property_descriptions: Dict[str, PropertyDescription] = {
        'is_header': PropertyDescription('w:trPr', 'w:tblHeader', None, True),
        'height': PropertyDescription('w:trPr', 'w:trHeight', 'w:val', True),
        'height_rule': PropertyDescription('w:trPr', 'w:trHeight', 'w:hRule', True),
    }

    cell_property_descriptions: Dict[str, PropertyDescription] = {
        'fill_color': PropertyDescription('w:tcPr', 'w:shd', 'w:fill', True),
        'fill_theme': PropertyDescription('w:tcPr', 'w:shd', 'w:themeFill', True),
        'border_top': PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:val', True),
        'border_top_color': PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:color', True),
        'border_top_size': PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:sz', True),
        'border_bottom': PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:val', True),
        'border_bottom_color': PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:color', True),
        'border_bottom_size': PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:sz', True),
        'border_right': PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:val', True),
        'border_right_color': PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:color', True),
        'border_right_size': PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:sz', True),
        'border_left': PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:val', True),
        'border_left_color': PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:color', True),
        'border_left_size': PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:sz', True),
        'width': PropertyDescription('w:tcPr', 'w:tcW', 'w:w', True),
        'width_type': PropertyDescription('w:tcPr', 'w:tcW', 'w:type', True),
        'col_span': PropertyDescription('w:tcPr', 'w:gridSpan', 'w:val', True),
        'vertical_merge': PropertyDescription('w:tcPr', 'w:vMerge', 'w:val', True),
        'is_vertical_merge_continue': PropertyDescription('w:tcPr', 'w:vMerge', None, True),
        'vertical_align': PropertyDescription('w:tcPr', 'w:vAlign', 'w:val', True),
        'text_direction': PropertyDescription('w:tcPr', 'w:textDirection', 'w:val', True),
        'margin_top': PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:w', True),
        'margin_top_type': PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:type', True),
        'margin_bottom': PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:w', True),
        'margin_bottom_type': PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:type', True),
        'margin_left': PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:w', True),
        'margin_left_type': PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:type', True),
        'margin_right': PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:w', True),
        'margin_right_type': PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:type', True)
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
