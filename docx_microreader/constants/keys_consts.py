from enum import Enum, unique
from typing import Union, List


class StrEnum(str, Enum):
    def __eq__(self, other):
        """
        :param other: str or StrEnum
        """
        if isinstance(other, str):
            if self.value == other:
                return True
            return False
        return self == other


@unique
class ElementTag(StrEnum):
    BODY = 'w:body'
    PARAGRAPH = 'w:p'
    RUN = 'w:r'
    TEXT = 'w:t'
    TABLE = 'w:tbl'
    ROW = 'w:tr'
    CELL = 'w:tc'
    DRAWING = 'w:drawing'
    IMAGE = 'wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip'
    STYLE = 'w:style'
    STYLE_TABLE_AREA = 'w:tblStylePr'


@unique
class Element(StrEnum):
    BODY = 'body'
    PARAGRAPH = 'paragraph'
    RUN = 'run'
    TEXT = 'text'
    TABLE = 'table'
    ROW = 'row'
    CELL = 'cell'
    DRAWING = 'drawing'
    IMAGE = 'image'


# Elements for which there are no models
@unique
class SubElement(StrEnum):
    BORDER = 'border'
    MARGIN = 'margin'
    BACKGROUND = 'background'
    UNDERLINE = 'underline'
    COLUMN = 'column'


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


@unique
class Direction(StrEnum):
    TOP = 'top'
    BOTTOM = 'bottom'
    LEFT = 'left'
    RIGHT = 'right'
    HORIZONTAL = 'horizontal'
    VERTICAL = 'vertical'

    @staticmethod
    def horizontal_or_vertical_straight(value) -> Union[HORIZONTAL, VERTICAL]:
        """
        direction of straight. It is HORIZONTAL if place to TOP or BOTTOM and VERTICAL if place to RIGHT or LEFT
        :param value: str or value of Direction enum
        :return: Direction.HORIZONTAL (for horizontal, top, bottom) or Direction.VERTICAL (for vertical, right, left)
        """
        if Direction.TOP == value or Direction.BOTTOM == value:
            return Direction.HORIZONTAL
        if Direction.RIGHT == value or Direction.LEFT == value:
            return Direction.VERTICAL
        raise ValueError(f"Placement of straight can't be equal {value}")


@unique
class PropertyName(StrEnum):
    COLOR = 'color'
    SIZE = 'size'
    TYPE = 'type'
    SPACE = 'space'
    WIDTH = 'width'
    WIDTH_TYPE = 'width_type'
    HEIGHT = 'height'
    HEIGHT_RULE = 'height_rule'
    ALIGN = 'align'
    ID = 'id'
    INDENT = 'indent'
    HANGING = 'hanging'
    FIRST_LINE = 'first_line'
    KEEP_LINES = 'keep_lines'
    KEEP_NEXT = 'keep_next'
    OUTLINE_LEVEL = 'outline_level'
    IS_BOLD = 'is_bold'
    IS_ITALIC = 'is_italic'
    IS_STRIKE = 'is_strike'
    LANGUAGE = 'language'
    THEME_COLOR = 'theme_color'
    FILL = 'fill'
    FILL_COLOR = 'fill_color'
    FILL_THEME = 'fill_theme'
    LAYOUT = 'layout'
    INDENTATION = 'indentation'
    INDENTATION_TYPE = 'indentation_type'
    STYLE_LOOK_FIRST = 'style_look_first'
    STYLE_LOOK_LAST = 'style_look_last'
    NO_BANDING = 'no_banding'
    IS_HEADER = 'is_header'
    SPAN = 'span'
    MERGE = 'merge'
    IS_MERGE_CONTINUE = 'is_merge_continue'
    DIRECTION = 'direction'


def get_property_key(element: str, *args, subelements: Union[str, None, List[str]] = None,
                     direction: Union[str, None] = None, property_name: Union[str, None] = None,
                     separator: str = ' ') -> str:
    """
    :param element: value from Element Enum or string
    :param args: values from Element, SubElement, Direction or PropertyName Enums
    :param separator: separator between elements in result key
    :param subelements: value or list of value from Element Enum or SubElement Enum or string;
                        (i+1)-th element of list is subelement for i-th element of list
    :param direction: value from Direction Enum or string
    :param property_name: value from PropertyName Enum or string
    :result: key for dict of properties
    """
    _direction: Union[str, None] = direction.value if isinstance(direction, Direction) else direction
    _property_name: Union[str, None] = property_name.value if isinstance(property_name, PropertyName) else property_name
    _subelements: Union[str, None]
    if isinstance(subelements, list):
        _subelements: Union[str, None] = ''
        for i, el in enumerate(subelements):
            if isinstance(el, SubElement) or isinstance(el, Element):
                _subelements += (separator + el.value) if not i == 0 else el.value
            else:
                _subelements += separator + el
    else:
        _subelements = subelements

    if isinstance(element, Element):
        element = element.value
    for arg in args:
        if isinstance(arg, Direction):
            if _direction is None:
                _direction = arg.value
        elif isinstance(arg, Element) or isinstance(arg, SubElement):
            if _subelements is None:
                _subelements = arg.value
        elif isinstance(arg, PropertyName):
            if _property_name is None:
                _property_name = arg.value

    _subelements = (separator + _subelements) if _subelements is not None else ''
    _direction = (separator + _direction) if _direction is not None else ''
    _property_name = (separator + _property_name) if _property_name is not None else ''

    return f'{element}{_subelements}{_direction}{_property_name}'


DRAWING_SIZE_HORIZONTAL: str = get_property_key(Element.DRAWING, Direction.HORIZONTAL, PropertyName.SIZE)
DRAWING_SIZE_VERTICAL: str = get_property_key(Element.DRAWING, Direction.VERTICAL, PropertyName.SIZE)
IMAGE_ID: str = get_property_key(Element.IMAGE, PropertyName.ID)
PARAGRAPH_ALIGN: str = get_property_key(Element.PARAGRAPH, PropertyName.ALIGN)
PARAGRAPH_INDENT_LEFT: str = get_property_key(Element.PARAGRAPH, Direction.LEFT, PropertyName.INDENT)
PARAGRAPH_INDENT_RIGHT: str = get_property_key(Element.PARAGRAPH, Direction.RIGHT, PropertyName.INDENT)
PARAGRAPH_HANGING: str = get_property_key(Element.PARAGRAPH, PropertyName.HANGING)
PARAGRAPH_FIRST_LINE: str = get_property_key(Element.PARAGRAPH, PropertyName.FIRST_LINE)
PARAGRAPH_KEEP_LINES: str = get_property_key(Element.PARAGRAPH, PropertyName.KEEP_LINES)
PARAGRAPH_KEEP_NEXT: str = get_property_key(Element.PARAGRAPH, PropertyName.KEEP_NEXT)
PARAGRAPH_OUTLINE_LEVEL: str = get_property_key(Element.PARAGRAPH, PropertyName.OUTLINE_LEVEL)
PARAGRAPH_BORDER_TOP_TYPE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.TOP,
                                                  PropertyName.TYPE)
PARAGRAPH_BORDER_TOP_COLOR: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.TOP,
                                                   PropertyName.COLOR)
PARAGRAPH_BORDER_TOP_SIZE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.TOP,
                                                  PropertyName.SIZE)
PARAGRAPH_BORDER_TOP_SPACE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.TOP,
                                                   PropertyName.SPACE)
PARAGRAPH_BORDER_BOTTOM_TYPE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.BOTTOM,
                                                     PropertyName.TYPE)
PARAGRAPH_BORDER_BOTTOM_COLOR: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.BOTTOM,
                                                      PropertyName.COLOR)
PARAGRAPH_BORDER_BOTTOM_SIZE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.BOTTOM,
                                                     PropertyName.SIZE)
PARAGRAPH_BORDER_BOTTOM_SPACE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.BOTTOM,
                                                      PropertyName.SPACE)
PARAGRAPH_BORDER_RIGHT_TYPE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.RIGHT,
                                                    PropertyName.TYPE)
PARAGRAPH_BORDER_RIGHT_COLOR: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.RIGHT,
                                                     PropertyName.COLOR)
PARAGRAPH_BORDER_RIGHT_SIZE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.RIGHT,
                                                    PropertyName.SIZE)
PARAGRAPH_BORDER_RIGHT_SPACE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.RIGHT,
                                                    PropertyName.SPACE)
PARAGRAPH_BORDER_LEFT_TYPE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.LEFT,
                                                   PropertyName.TYPE)
PARAGRAPH_BORDER_LEFT_COLOR: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.LEFT,
                                                    PropertyName.COLOR)
PARAGRAPH_BORDER_LEFT_SIZE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.LEFT,
                                                   PropertyName.SIZE)
PARAGRAPH_BORDER_LEFT_SPACE: str = get_property_key(Element.PARAGRAPH, SubElement.BORDER, Direction.LEFT,
                                                    PropertyName.SPACE)
RUN_SIZE: str = get_property_key(Element.RUN, PropertyName.SIZE)
RUN_IS_BOLD: str = get_property_key(Element.RUN, Element.TEXT, PropertyName.IS_BOLD)
RUN_IS_ITALIC: str = get_property_key(Element.RUN, Element.TEXT, PropertyName.IS_ITALIC)
RUN_VERTICAL_ALIGN: str = get_property_key(Element.RUN, Direction.VERTICAL, PropertyName.ALIGN)
RUN_LANGUAGE: str = get_property_key(Element.RUN, PropertyName.LANGUAGE)
RUN_COLOR: str = get_property_key(Element.RUN, PropertyName.COLOR)
RUN_THEME_COLOR: str = get_property_key(Element.RUN, PropertyName.THEME_COLOR)
RUN_BACKGROUND_COLOR: str = get_property_key(Element.RUN, SubElement.BACKGROUND, PropertyName.COLOR)
RUN_BACKGROUND_FILL: str = get_property_key(Element.RUN, SubElement.BACKGROUND, PropertyName.FILL)
RUN_UNDERLINE_TYPE: str = get_property_key(Element.RUN, SubElement.UNDERLINE, PropertyName.TYPE)
RUN_UNDERLINE_COLOR: str = get_property_key(Element.RUN, SubElement.UNDERLINE, PropertyName.COLOR)
RUN_IS_STRIKE: str = get_property_key(Element.RUN, Element.TEXT, PropertyName.IS_STRIKE)
RUN_BORDER_TYPE: str = get_property_key(Element.RUN, SubElement.BORDER, PropertyName.TYPE)
RUN_BORDER_COLOR: str = get_property_key(Element.RUN, SubElement.BORDER, PropertyName.COLOR)
RUN_BORDER_SIZE: str = get_property_key(Element.RUN, SubElement.BORDER, PropertyName.SIZE)
RUN_BORDER_SPACE: str = get_property_key(Element.RUN, SubElement.BORDER, PropertyName.SPACE)
TABLE_LAYOUT: str = get_property_key(Element.TABLE, PropertyName.LAYOUT)
TABLE_WIDTH: str = get_property_key(Element.TEXT, PropertyName.WIDTH)
TABLE_WIDTH_TYPE: str = get_property_key(Element.TEXT, PropertyName.WIDTH_TYPE)
TABLE_ALIGN: str = get_property_key(Element.TABLE, PropertyName.ALIGN)
TABLE_INSIDE_BORDER_HORIZONTAL_TYPE: str = get_property_key(Element.TABLE, Direction.HORIZONTAL, PropertyName.TYPE,
                                                            subelements=[Element.CELL, SubElement.BORDER])
TABLE_INSIDE_BORDER_HORIZONTAL_COLOR: str = get_property_key(Element.TABLE, Direction.HORIZONTAL, PropertyName.COLOR,
                                                             subelements=[Element.CELL, SubElement.BORDER])
TABLE_INSIDE_BORDER_HORIZONTAL_SIZE: str = get_property_key(Element.TABLE, Direction.HORIZONTAL, PropertyName.SIZE,
                                                            subelements=[Element.CELL, SubElement.BORDER])
TABLE_INSIDE_BORDER_VERTICAL_TYPE: str = get_property_key(Element.TABLE, Direction.VERTICAL, PropertyName.TYPE,
                                                          subelements=[Element.CELL, SubElement.BORDER])
TABLE_INSIDE_BORDER_VERTICAL_COLOR: str = get_property_key(Element.TABLE, Direction.VERTICAL, PropertyName.COLOR,
                                                           subelements=[Element.CELL, SubElement.BORDER])
TABLE_INSIDE_BORDER_VERTICAL_SIZE: str = get_property_key(Element.TABLE, Direction.VERTICAL, PropertyName.SIZE,
                                                          subelements=[Element.CELL, SubElement.BORDER])
TABLE_BORDER_TOP_TYPE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.TOP, PropertyName.TYPE)
TABLE_BORDER_TOP_COLOR: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.TOP, PropertyName.COLOR)
TABLE_BORDER_TOP_SIZE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.TOP, PropertyName.SIZE)
TABLE_BORDER_BOTTOM_TYPE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.BOTTOM, PropertyName.TYPE)
TABLE_BORDER_BOTTOM_COLOR: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.BOTTOM, PropertyName.COLOR)
TABLE_BORDER_BOTTOM_SIZE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.BOTTOM, PropertyName.SIZE)
TABLE_BORDER_RIGHT_TYPE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.RIGHT, PropertyName.TYPE)
TABLE_BORDER_RIGHT_COLOR: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.RIGHT, PropertyName.COLOR)
TABLE_BORDER_RIGHT_SIZE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.RIGHT, PropertyName.SIZE)
TABLE_BORDER_LEFT_TYPE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.LEFT, PropertyName.TYPE)
TABLE_BORDER_LEFT_COLOR: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.LEFT, PropertyName.COLOR)
TABLE_BORDER_LEFT_SIZE: str = get_property_key(Element.TABLE, SubElement.BORDER, Direction.LEFT, PropertyName.SIZE)
TABLE_CELL_MARGIN_TOP_TYPE: str = get_property_key(Element.TABLE, Direction.TOP, PropertyName.TYPE,
                                                   subelements=[Element.CELL, SubElement.MARGIN])
TABLE_CELL_MARGIN_TOP_SIZE: str = get_property_key(Element.TABLE, Direction.TOP, PropertyName.SIZE,
                                                   subelements=[Element.CELL, SubElement.MARGIN])
TABLE_CELL_MARGIN_BOTTOM_TYPE: str = get_property_key(Element.TABLE, Direction.BOTTOM, PropertyName.TYPE,
                                                      subelements=[Element.CELL, SubElement.MARGIN])
TABLE_CELL_MARGIN_BOTTOM_SIZE: str = get_property_key(Element.TABLE, Direction.BOTTOM, PropertyName.SIZE,
                                                      subelements=[Element.CELL, SubElement.MARGIN])
TABLE_CELL_MARGIN_RIGHT_TYPE: str = get_property_key(Element.TABLE, Direction.RIGHT, PropertyName.TYPE,
                                                     subelements=[Element.CELL, SubElement.MARGIN])
TABLE_CELL_MARGIN_RIGHT_SIZE: str = get_property_key(Element.TABLE, Direction.RIGHT, PropertyName.SIZE,
                                                     subelements=[Element.CELL, SubElement.MARGIN])
TABLE_CELL_MARGIN_LEFT_TYPE: str = get_property_key(Element.TABLE, Direction.LEFT, PropertyName.TYPE,
                                                    subelements=[Element.CELL, SubElement.MARGIN])
TABLE_CELL_MARGIN_LEFT_SIZE: str = get_property_key(Element.TABLE, Direction.LEFT, PropertyName.SIZE,
                                                    subelements=[Element.CELL, SubElement.MARGIN])
TABLE_INDENTATION: str = get_property_key(Element.TABLE, PropertyName.INDENTATION)
TABLE_INDENTATION_TYPE: str = get_property_key(Element.TABLE, PropertyName.INDENTATION_TYPE)
TABLE_FIRST_ROW_STYLE_LOOK: str = get_property_key(Element.TABLE, Element.ROW, PropertyName.STYLE_LOOK_FIRST)
TABLE_FIRST_COLUMN_STYLE_LOOK: str = get_property_key(Element.TABLE, SubElement.COLUMN, PropertyName.STYLE_LOOK_FIRST)
TABLE_LAST_ROW_STYLE_LOOK: str = get_property_key(Element.TABLE, Element.ROW, PropertyName.STYLE_LOOK_LAST)
TABLE_LAST_COLUMN_STYLE_LOOK: str = get_property_key(Element.TABLE, SubElement.COLUMN, PropertyName.STYLE_LOOK_LAST)
TABLE_NO_HORIZONTAL_BANDING: str = get_property_key(Element.TABLE, Direction.HORIZONTAL, PropertyName.NO_BANDING)
TABLE_NO_VERTICAL_BANDING: str = get_property_key(Element.TABLE, Direction.VERTICAL, PropertyName.NO_BANDING)
ROW_IS_HEADER: str = get_property_key(Element.ROW, PropertyName.IS_HEADER)
ROW_HEIGHT: str = get_property_key(Element.ROW, PropertyName.HEIGHT)
ROW_HEIGHT_RULE: str = get_property_key(Element.ROW, PropertyName.HEIGHT_RULE)
CELL_FILL_COLOR: str = get_property_key(Element.CELL, PropertyName.FILL_COLOR)
CELL_FILL_THEME: str = get_property_key(Element.CELL, PropertyName.FILL_THEME)
CELL_BORDER_TOP_TYPE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.TOP, PropertyName.TYPE)
CELL_BORDER_TOP_COLOR: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.TOP, PropertyName.COLOR)
CELL_BORDER_TOP_SIZE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.TOP, PropertyName.SIZE)
CELL_BORDER_BOTTOM_TYPE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.BOTTOM, PropertyName.TYPE)
CELL_BORDER_BOTTOM_COLOR: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.BOTTOM, PropertyName.COLOR)
CELL_BORDER_BOTTOM_SIZE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.BOTTOM, PropertyName.SIZE)
CELL_BORDER_RIGHT_TYPE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.RIGHT, PropertyName.TYPE)
CELL_BORDER_RIGHT_COLOR: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.RIGHT, PropertyName.COLOR)
CELL_BORDER_RIGHT_SIZE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.RIGHT, PropertyName.SIZE)
CELL_BORDER_LEFT_TYPE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.LEFT, PropertyName.TYPE)
CELL_BORDER_LEFT_COLOR: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.LEFT, PropertyName.COLOR)
CELL_BORDER_LEFT_SIZE: str = get_property_key(Element.CELL, SubElement.BORDER, Direction.LEFT, PropertyName.SIZE)
CELL_WIDTH: str = get_property_key(Element.CELL, PropertyName.WIDTH)
CELL_WIDTH_TYPE: str = get_property_key(Element.CELL, PropertyName.WIDTH_TYPE)
CELL_COLUMN_SPAN: str = get_property_key(Element.CELL, SubElement.COLUMN, PropertyName.SPAN)
CELL_VERTICAL_MARGE: str = get_property_key(Element.CELL, Direction.VERTICAL, PropertyName.MERGE)
CELL_IS_VERTICAL_MARGE_CONTINUE: str = get_property_key(Element.CELL, Direction.VERTICAL,
                                                        PropertyName.IS_MERGE_CONTINUE)
CELL_VERTICAL_ALIGN: str = get_property_key(Element.CELL, Direction.VERTICAL, PropertyName.ALIGN)
CELL_TEXT_DIRECTION: str = get_property_key(Element.CELL, Element.TEXT, PropertyName.DIRECTION)
CELL_MARGIN_TOP_SIZE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.TOP, PropertyName.SIZE)
CELL_MARGIN_TOP_TYPE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.TOP, PropertyName.TYPE)
CELL_MARGIN_BOTTOM_SIZE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.BOTTOM, PropertyName.SIZE)
CELL_MARGIN_BOTTOM_TYPE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.BOTTOM, PropertyName.TYPE)
CELL_MARGIN_RIGHT_SIZE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.RIGHT, PropertyName.SIZE)
CELL_MARGIN_RIGHT_TYPE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.RIGHT, PropertyName.TYPE)
CELL_MARGIN_LEFT_SIZE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.LEFT, PropertyName.SIZE)
CELL_MARGIN_LEFT_TYPE: str = get_property_key(Element.CELL, SubElement.MARGIN, Direction.LEFT, PropertyName.TYPE)
