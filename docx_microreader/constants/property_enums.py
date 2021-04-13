from enum import Enum, unique
from ..properties import PropertyDescription
from typing import Dict, Union
from functools import cache
from ..utils.enums import ElementPropertyEnum, convert_to_enum_element
from ..mixins.enums import EnumOfBorderedElementMixin, CellMarginEnumMixin


_BODY_NAME: str = 'body'
_PARAGRAPH_NAME: str = 'paragraph'
_RUN_NAME: str = 'run'
_TEXT_NAME: str = 'text'
_TABLE_NAME: str = 'table'
_ROW_NAME: str = 'row'
_CELL_NAME: str = 'cell'
_DRAWING_NAME: str = 'drawing'
_IMAGE_NAME: str = 'image'
_LINE_BREAK: str = 'break'
_CARRIAGE_RETURN: str = 'carriage return'
_TABULATION: str = 'tabulation'

# name of attribute for bool property values. Properties can be missed only if corresponding to this attribute
MissedPropertyAttribute: str = 'w:val'


@unique
class Direction(str, Enum):
    TOP = 'top'
    BOTTOM = 'bottom'
    LEFT = 'left'
    RIGHT = 'right'
    HORIZONTAL = 'horizontal'
    VERTICAL = 'vertical'

    @staticmethod
    def horizontal_or_vertical_straight(direction):
        """
        :param direction: instance of Direction or corresponding str
        :return: Direction.HORIZONTAL for TOP, BOTTOM, HORIZONTAL else Direction.VERTICAL
        """
        d = convert_to_enum_element(direction, Direction)
        if d == Direction.TOP or d == Direction.BOTTOM or d == Direction.HORIZONTAL:
            return Direction.HORIZONTAL
        return Direction.VERTICAL  # d == Direction.LEFT or d == Direction.RIGHT or d == Direction.VERTICAL


@unique
class SubElement(str, Enum):
    BORDER = 'border'
    MARGIN = 'margin'


@unique
class BorderProperty(str, Enum):
    TYPE = 'type'
    COLOR = 'color'
    SIZE = 'size'
    SPACE = 'space'


@unique
class MarginProperty(str, Enum):
    TYPE = 'type'
    SIZE = 'size'


@cache
def subelement_property_key(subelement: SubElement, direction: Union[str, None, Direction],
                             property_name: Union[str, BorderProperty, MarginProperty]) -> str:
    """
    create key for properties of SubElement instances
    :param subelement: SubElement instance
    :param direction: Direction instance or corresponding value
    :param property_name: instance of BorderProperty (if subelement is Border) or
                          MarginProperty (if subelement is Margin) or corresponding value
    :return: key that using for create keys for dict of XMLelement properties
    """
    pr_name: str
    if isinstance(property_name, str):
        if subelement == SubElement.BORDER:
            pr_name = convert_to_enum_element(property_name, BorderProperty).value
        else:  # subelement == SubElement.MARGIN:
            pr_name = convert_to_enum_element(property_name, MarginProperty).value
    elif isinstance(property_name, BorderProperty) and subelement == SubElement.BORDER or \
            isinstance(property_name, MarginProperty) and subelement == SubElement.MARGIN:
        pr_name = property_name.value
    else:
        raise TypeError(f'''type of property_name must be equal str, BorderProperty or MarginProperty.
                         Type of property_name ({type(property_name)}) must correspond of subelement 
                         ({subelement.value}) value''')

    if direction is not None:
        return f'{convert_to_enum_element(direction, Direction).value} {subelement.value} {pr_name}'
    return f'{subelement.value} {pr_name}'


@unique
class ElementBorderProperty(str, Enum):
    TOP_TYPE = subelement_property_key(SubElement.BORDER, Direction.TOP, BorderProperty.TYPE)
    TOP_COLOR = subelement_property_key(SubElement.BORDER, Direction.TOP, BorderProperty.COLOR)
    TOP_SIZE = subelement_property_key(SubElement.BORDER, Direction.TOP, BorderProperty.SIZE)
    TOP_SPACE = subelement_property_key(SubElement.BORDER, Direction.TOP, BorderProperty.SPACE)
    BOTTOM_TYPE = subelement_property_key(SubElement.BORDER, Direction.BOTTOM, BorderProperty.TYPE)
    BOTTOM_COLOR = subelement_property_key(SubElement.BORDER, Direction.BOTTOM, BorderProperty.COLOR)
    BOTTOM_SIZE = subelement_property_key(SubElement.BORDER, Direction.BOTTOM, BorderProperty.SIZE)
    BOTTOM_SPACE = subelement_property_key(SubElement.BORDER, Direction.BOTTOM, BorderProperty.SPACE)
    LEFT_TYPE = subelement_property_key(SubElement.BORDER, Direction.LEFT, BorderProperty.TYPE)
    LEFT_COLOR = subelement_property_key(SubElement.BORDER, Direction.LEFT, BorderProperty.COLOR)
    LEFT_SIZE = subelement_property_key(SubElement.BORDER, Direction.LEFT, BorderProperty.SIZE)
    LEFT_SPACE = subelement_property_key(SubElement.BORDER, Direction.LEFT, BorderProperty.SPACE)
    RIGHT_TYPE = subelement_property_key(SubElement.BORDER, Direction.RIGHT, BorderProperty.TYPE)
    RIGHT_COLOR = subelement_property_key(SubElement.BORDER, Direction.RIGHT, BorderProperty.COLOR)
    RIGHT_SIZE = subelement_property_key(SubElement.BORDER, Direction.RIGHT, BorderProperty.SIZE)
    RIGHT_SPACE = subelement_property_key(SubElement.BORDER, Direction.RIGHT, BorderProperty.SPACE)
    HORIZONTAL_TYPE = subelement_property_key(SubElement.BORDER, Direction.HORIZONTAL, BorderProperty.TYPE)
    HORIZONTAL_COLOR = subelement_property_key(SubElement.BORDER, Direction.HORIZONTAL, BorderProperty.COLOR)
    HORIZONTAL_SIZE = subelement_property_key(SubElement.BORDER, Direction.HORIZONTAL, BorderProperty.SIZE)
    VERTICAL_TYPE = subelement_property_key(SubElement.BORDER, Direction.VERTICAL, BorderProperty.TYPE)
    VERTICAL_COLOR = subelement_property_key(SubElement.BORDER, Direction.VERTICAL, BorderProperty.COLOR)
    VERTICAL_SIZE = subelement_property_key(SubElement.BORDER, Direction.VERTICAL, BorderProperty.SIZE)


@unique
class ElementMarginProperty(str, Enum):
    TOP_TYPE = subelement_property_key(SubElement.MARGIN, Direction.TOP, BorderProperty.TYPE)
    TOP_SIZE = subelement_property_key(SubElement.MARGIN, Direction.TOP, BorderProperty.SIZE)
    BOTTOM_TYPE = subelement_property_key(SubElement.MARGIN, Direction.BOTTOM, BorderProperty.TYPE)
    BOTTOM_SIZE = subelement_property_key(SubElement.MARGIN, Direction.BOTTOM, BorderProperty.SIZE)
    LEFT_TYPE = subelement_property_key(SubElement.MARGIN, Direction.LEFT, BorderProperty.TYPE)
    LEFT_SIZE = subelement_property_key(SubElement.MARGIN, Direction.LEFT, BorderProperty.SIZE)
    RIGHT_TYPE = subelement_property_key(SubElement.MARGIN, Direction.RIGHT, BorderProperty.TYPE)
    RIGHT_SIZE = subelement_property_key(SubElement.MARGIN, Direction.RIGHT, BorderProperty.SIZE)


@cache
def subelement_property_key_of_element(element_name: str, subelement_property_name: Union[str, ElementBorderProperty,
                                                                                          ElementMarginProperty]) -> str:
    """
    Combines the element name and property name into a key
    :param element_name: name of element
    :param subelement_property_name: instance of ElementBorderProperty, ElementMarginProperty or corresponding value
    :return: key that using for create keys for dict of XMLelement properties
    """
    if isinstance(subelement_property_name, str):
        return f'{subelement_property_name} of {element_name}'
    return f'{subelement_property_name.value} of {element_name}'


@unique
class BodyProperty(ElementPropertyEnum):
    pass


@unique
class ParagraphProperty(EnumOfBorderedElementMixin, ElementPropertyEnum):
    ALIGN = ('align of paragraph', PropertyDescription('w:pPr', 'w:jc', 'w:val'), True)
    INDENT_LEFT = ('indent left of paragraph', PropertyDescription('w:pPr', 'w:ind', ['w:left', 'w:start']), True)
    INDENT_RIGHT = ('indent right of paragraph', PropertyDescription('w:pPr', 'w:ind', ['w:right', 'w:end']), True)
    HANGING = ('hanging of paragraph', PropertyDescription('w:pPr', 'w:ind', 'w:hanging'), True)
    FIRST_LINE = ('first line of paragraph', PropertyDescription('w:pPr', 'w:ind', 'w:firstLine'), True)
    KEEP_LINES = ('keep lines of paragraph', PropertyDescription('w:pPr', 'w:keepLines', is_can_be_miss=True), True)
    KEEP_NEXT = ('keep next of paragraph', PropertyDescription('w:pPr', 'w:keepNext', is_can_be_miss=True), True)
    OUTLINE_LEVEL = ('outline level of paragraph', PropertyDescription('w:pPr', 'w:outlineLvl', 'w:val'), True)
    TOP_BORDER_TYPE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.TOP_TYPE),
                       PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:val'), True)
    TOP_BORDER_COLOR = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.TOP_COLOR),
                        PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:color'), True)
    TOP_BORDER_SIZE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.TOP_SIZE),
                       PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:sz'), True)
    TOP_BORDER_SPACE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.TOP_SPACE),
                        PropertyDescription('w:pPr/w:pBdr', 'w:top', 'w:space'), True)
    BOTTOM_BORDER_TYPE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.BOTTOM_TYPE),
                          PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:val'), True)
    BOTTOM_BORDER_COLOR = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.BOTTOM_COLOR),
                           PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:color'), True)
    BOTTOM_BORDER_SIZE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.BOTTOM_SIZE),
                          PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:sz'), True)
    BOTTOM_BORDER_SPACE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.BOTTOM_SPACE),
                           PropertyDescription('w:pPr/w:pBdr', 'w:bottom', 'w:space'), True)
    RIGHT_BORDER_TYPE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.RIGHT_TYPE),
                         PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:val'), True)
    RIGHT_BORDER_COLOR = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.RIGHT_COLOR),
                          PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:color'), True)
    RIGHT_BORDER_SIZE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.RIGHT_SIZE),
                         PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:sz'), True)
    RIGHT_BORDER_SPACE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.RIGHT_SPACE),
                          PropertyDescription('w:pPr/w:pBdr', 'w:right', 'w:space'), True)
    LEFT_BORDER_TYPE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.LEFT_TYPE),
                        PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:val'), True)
    LEFT_BORDER_COLOR = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.LEFT_COLOR),
                         PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:color'), True)
    LEFT_BORDER_SIZE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.LEFT_SIZE),
                        PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:sz'), True)
    LEFT_BORDER_SPACE = (subelement_property_key_of_element(_PARAGRAPH_NAME, ElementBorderProperty.LEFT_SPACE),
                         PropertyDescription('w:pPr/w:pBdr', 'w:left', 'w:space'), True)
    STYLE = ('style of paragraph', PropertyDescription('w:pPr', 'w:pStyle', 'w:val'))

    @classmethod
    def _element_key(cls) -> str:
        return _PARAGRAPH_NAME


@unique
class RunProperty(EnumOfBorderedElementMixin, ElementPropertyEnum):
    FONT_ASCII = ('font ascii of run', PropertyDescription('w:rPr', 'w:rFonts', 'w:ascii'), True)
    FONT_EAST_ASIA = ('font east asia of run', PropertyDescription('w:rPr', 'w:rFonts', 'w:eastAsia'), True)
    FONT_H_ANSI = ('font h ansi of run', PropertyDescription('w:rPr', 'w:rFonts', 'w:hAnsi'), True)
    FONT_CS = ('font cs of run', PropertyDescription('w:rPr', 'w:rFonts', 'w:cs'), True)
    SIZE = ('size of run', PropertyDescription('w:rPr', 'w:sz', 'w:val'), True)
    BOLD = ('bold of run', PropertyDescription('w:rPr', 'w:b', is_can_be_miss=True), True)
    ITALIC = ('italic of run', PropertyDescription('w:rPr', 'w:i', is_can_be_miss=True), True)
    STRIKE = ('strike of run', PropertyDescription('w:rPr', 'w:strike', is_can_be_miss=True), True)
    VERTICAL_ALIGN = ('vertical align of run', PropertyDescription('w:rPr', 'w:vertAlign', 'w:val'), True)
    LANGUAGE = ('language of run', PropertyDescription('w:rPr', 'w:lang', 'w:val'), True)
    COLOR = ('color of run', PropertyDescription('w:rPr', 'w:color', 'w:val'), True)
    THEME_COLOR = ('theme color of run', PropertyDescription('w:rPr', 'w:color', 'w:themeColor'), True)
    BACKGROUND_COLOR = ('background color of run', PropertyDescription('w:rPr', 'w:highlight', 'w:val'), True)
    BACKGROUND_FILL = ('background fill of run', PropertyDescription('w:rPr', 'w:shd', 'w:fill'), True)
    UNDERLINE_TYPE = ('underline type of run', PropertyDescription('w:rPr', 'w:u', 'w:val'), True)
    UNDERLINE_COLOR = ('underline color of run', PropertyDescription('w:rPr', 'w:u', 'w:color'), True)
    BORDER_TYPE = (
        subelement_property_key_of_element(
            _RUN_NAME, subelement_property_key(SubElement.BORDER, None, BorderProperty.TYPE)
        ),
        PropertyDescription('w:rPr', 'w:bdr', 'w:val'),
        True
    )
    BORDER_COLOR = (
        subelement_property_key_of_element(
            _RUN_NAME, subelement_property_key(SubElement.BORDER, None, BorderProperty.COLOR)
        ),
        PropertyDescription('w:rPr', 'w:bdr', 'w:color'),
        True
    )
    BORDER_SIZE = (
        subelement_property_key_of_element(
            _RUN_NAME, subelement_property_key(SubElement.BORDER, None, BorderProperty.SIZE)
        ),
        PropertyDescription('w:rPr', 'w:bdr', 'w:sz'),
        True
    )
    BORDER_SPACE = (
        subelement_property_key_of_element(
            _RUN_NAME, subelement_property_key(SubElement.BORDER, None, BorderProperty.SPACE)
        ),
        PropertyDescription('w:rPr', 'w:bdr', 'w:space'),
        True
    )
    STYLE = ('char style of run', PropertyDescription('w:rPr', 'w:rStyle', 'w:val'))

    @classmethod
    def _element_key(cls) -> str:
        return _RUN_NAME


@unique
class TextProperty(ElementPropertyEnum):
    SPACE = ('space of text', PropertyDescription(None, None, 'xml:space'))


@unique
class LineBreakProperty(ElementPropertyEnum):
    TYPE = ('type of line break', PropertyDescription(None, None, 'w:type'))
    CLEAR = ('clear of line break', PropertyDescription(None, None, 'w:clear'))


@unique
class CarriageReturnProperty(ElementPropertyEnum):
    pass


@unique
class TabulationProperty(ElementPropertyEnum):
    pass


@unique
class TableProperty(EnumOfBorderedElementMixin, CellMarginEnumMixin, ElementPropertyEnum):
    LAYOUT = ('layout of table', PropertyDescription('w:tblPr', 'w:tblLayout', 'w:type'), True)
    WIDTH = ('width of table', PropertyDescription('w:tblPr', 'w:tblW', 'w:w'), True)
    WIDTH_TYPE = ('width type of table', PropertyDescription('w:tblPr', 'w:tblW', 'w:type'), True)
    ALIGN = ('align of table', PropertyDescription('w:tblPr', 'w:jc', 'w:val'), True)
    INDENTATION = ('indentation of table', PropertyDescription('w:tblPr', 'w:tblInd', ['w:w', 'w:val']), True)
    INDENTATION_TYPE = ('indentation type of table', PropertyDescription('w:tblPr', 'w:tblInd', 'w:type'), True)
    INSIDE_HORIZONTAL_BORDER_TYPE = (
        subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.HORIZONTAL_TYPE),
        PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideH', 'w:val'),
        True
    )
    INSIDE_HORIZONTAL_BORDER_COLOR = (
        subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.HORIZONTAL_COLOR),
        PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideH', 'w:color'),
        True
    )
    INSIDE_HORIZONTAL_BORDER_SIZE = (
        subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.HORIZONTAL_SIZE),
        PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideH', 'w:sz'),
        True
    )
    INSIDE_VERTICAL_BORDER_TYPE = (
        subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.VERTICAL_TYPE),
        PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideV', 'w:val'),
        True
    )
    INSIDE_VERTICAL_BORDER_COLOR = (
        subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.VERTICAL_COLOR),
        PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideV', 'w:color'),
        True
    )
    INSIDE_VERTICAL_BORDER_SIZE = (
        subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.VERTICAL_SIZE),
        PropertyDescription(['w:tblPr/w:tblBorders', 'w:tcPr/w:tcBorders'], 'w:insideV', 'w:sz'),
        True
    )
    TOP_BORDER_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.TOP_TYPE),
                       PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:val'), True)
    TOP_BORDER_COLOR = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.TOP_COLOR),
                        PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:color'), True)
    TOP_BORDER_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.TOP_SIZE),
                       PropertyDescription('w:tblPr/w:tblBorders', 'w:top', 'w:sz'), True)
    BOTTOM_BORDER_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.BOTTOM_TYPE),
                          PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:val'), True)
    BOTTOM_BORDER_COLOR = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.BOTTOM_COLOR),
                           PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:color'), True)
    BOTTOM_BORDER_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.BOTTOM_SIZE),
                          PropertyDescription('w:tblPr/w:tblBorders', 'w:bottom', 'w:sz'), True)
    LEFT_BORDER_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.LEFT_TYPE),
                        PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:val'), True)
    LEFT_BORDER_COLOR = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.LEFT_COLOR),
                         PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:color'), True)
    LEFT_BORDER_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.LEFT_SIZE),
                        PropertyDescription('w:tblPr/w:tblBorders', ['w:left', 'w:start'], 'w:sz'), True)
    RIGHT_BORDER_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.RIGHT_TYPE),
                         PropertyDescription('w:tblPr/w:tblBorders', ['w:right', 'w:end'], 'w:val'), True)
    RIGHT_BORDER_COLOR = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.RIGHT_COLOR),
                          PropertyDescription('w:tblPr/w:tblBorders', ['w:right', 'w:end'], 'w:color'), True)
    RIGHT_BORDER_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementBorderProperty.RIGHT_SIZE),
                         PropertyDescription('w:tblPr/w:tblBorders', ['w:right', 'w:end'], 'w:sz'), True)
    CELL_MARGIN_TOP_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.TOP_SIZE),
                            PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:w'), True)
    CELL_MARGIN_TOP_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.TOP_TYPE),
                            PropertyDescription('w:tblPr/w:tblCellMar', 'w:top', 'w:type'), True)
    CELL_MARGIN_LEFT_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.LEFT_SIZE),
                             PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:w'), True)
    CELL_MARGIN_LEFT_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.LEFT_TYPE),
                             PropertyDescription('w:tblPr/w:tblCellMar', ['w:left', 'w:start'], 'w:type'), True)
    CELL_MARGIN_BOTTOM_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.BOTTOM_SIZE),
                               PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:w'), True)
    CELL_MARGIN_BOTTOM_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.BOTTOM_TYPE),
                               PropertyDescription('w:tblPr/w:tblCellMar', 'w:bottom', 'w:type'), True)
    CELL_MARGIN_RIGHT_SIZE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.RIGHT_SIZE),
                              PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:w'), True)
    CELL_MARGIN_RIGHT_TYPE = (subelement_property_key_of_element(_TABLE_NAME, ElementMarginProperty.RIGHT_TYPE),
                              PropertyDescription('w:tblPr/w:tblCellMar', ['w:right', 'w:end'], 'w:type'), True)
    STYLE = ('style of table', PropertyDescription('w:tblPr', 'w:tblStyle', 'w:val'))
    FIRST_ROW_STYLE_LOOK = ('first row style of table look', PropertyDescription('w:tblPr', 'w:tblLook', 'w:firstRow'))
    FIRST_COLUMN_STYLE_LOOK = ('first column style of table look',
                               PropertyDescription('w:tblPr', 'w:tblLook', 'w:firstColumn'))
    LAST_ROW_STYLE_LOOK = ('last row style of table look', PropertyDescription('w:tblPr', 'w:tblLook', 'w:lastRow'))
    LAST_COLUMN_STYLE_LOOK = ('last column style of table look',
                              PropertyDescription('w:tblPr', 'w:tblLook', 'w:lastColumn'))
    NO_HORIZONTAL_BANDING = ('no horizontal banding of table', PropertyDescription('w:tblPr', 'w:tblLook', 'w:noHBand'))
    NO_VERTICAL_BANDING = ('no vertical banding of table', PropertyDescription('w:tblPr', 'w:tblLook', 'w:noVBand'))

    @classmethod
    def _element_key(cls) -> str:
        return _TABLE_NAME


@unique
class RowProperty(ElementPropertyEnum):
    HEADER = ('is header row', PropertyDescription('w:trPr', 'w:tblHeader', is_can_be_miss=True), True)
    HEIGHT = ('height of row', PropertyDescription('w:trPr', 'w:trHeight', 'w:val'), True)
    HEIGHT_RULE = ('height rule of row', PropertyDescription('w:trPr', 'w:trHeight', 'w:hRule'), True)


@unique
class CellProperty(EnumOfBorderedElementMixin, CellMarginEnumMixin, ElementPropertyEnum):
    FILL_COLOR = ('fill color of cell', PropertyDescription('w:tcPr', 'w:shd', 'w:fill'), True)
    FILL_THEME = ('fill theme of cell', PropertyDescription('w:tcPr', 'w:shd', 'w:themeFill'), True)
    WIDTH = ('width of cell', PropertyDescription('w:tcPr', 'w:tcW', 'w:w'), True)
    WIDTH_TYPE = ('width type of cell', PropertyDescription('w:tcPr', 'w:tcW', 'w:type'), True)
    COLUMN_SPAN = ('column span of cell', PropertyDescription('w:tcPr', 'w:gridSpan', 'w:val'), True)
    VERTICAL_MERGE = ('vertical merge of cell', PropertyDescription('w:tcPr', 'w:vMerge', is_can_be_miss=True), True)
    VERTICAL_ALIGN = ('vertical align of cell', PropertyDescription('w:tcPr', 'w:vAlign', 'w:val'), True)
    TEXT_DIRECTION = ('text direction of cell', PropertyDescription('w:tcPr', 'w:textDirection', 'w:val'), True)
    TOP_BORDER_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.TOP_TYPE),
        PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:val'),
        True
    )
    TOP_BORDER_COLOR = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.TOP_COLOR),
        PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:color'),
        True
    )
    TOP_BORDER_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.TOP_SIZE),
        PropertyDescription('w:tcPr/w:tcBorders', 'w:top', 'w:sz'),
        True
    )
    BOTTOM_BORDER_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.BOTTOM_TYPE),
        PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:val'),
        True
    )
    BOTTOM_BORDER_COLOR = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.BOTTOM_COLOR),
        PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:color'),
        True
    )
    BOTTOM_BORDER_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.BOTTOM_SIZE),
        PropertyDescription('w:tcPr/w:tcBorders', 'w:bottom', 'w:sz'),
        True
    )
    RIGHT_BORDER_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.RIGHT_TYPE),
        PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:val'),
        True
    )
    RIGHT_BORDER_COLOR = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.RIGHT_COLOR),
        PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:color'),
        True
    )
    RIGHT_BORDER_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.RIGHT_SIZE),
        PropertyDescription('w:tcPr/w:tcBorders', ['w:right', 'w:end'], 'w:sz'),
        True
    )
    LEFT_BORDER_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.LEFT_TYPE),
        PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:val'),
        True
    )
    LEFT_BORDER_COLOR = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.LEFT_COLOR),
        PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:color'),
        True
    )
    LEFT_BORDER_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementBorderProperty.LEFT_SIZE),
        PropertyDescription('w:tcPr/w:tcBorders', ['w:left', 'w:start'], 'w:sz'),
        True
    )
    TOP_MARGIN_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.TOP_TYPE),
        PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:type'),
        True
    )
    TOP_MARGIN_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.TOP_SIZE),
        PropertyDescription('w:tcPr/w:tcMar', 'w:top', 'w:w'),
        True
    )
    BOTTOM_MARGIN_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.BOTTOM_TYPE),
        PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:type'),
        True
    )
    BOTTOM_MARGIN_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.BOTTOM_SIZE),
        PropertyDescription('w:tcPr/w:tcMar', 'w:bottom', 'w:w'),
        True
    )
    LEFT_MARGIN_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.LEFT_TYPE),
        PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:type'),
        True
    )
    LEFT_MARGIN_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.LEFT_SIZE),
        PropertyDescription('w:tcPr/w:tcMar', ['w:left', 'w:start'], 'w:w'),
        True
    )
    RIGHT_MARGIN_TYPE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.RIGHT_TYPE),
        PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:type'),
        True
    )
    RIGHT_MARGIN_SIZE = (
        subelement_property_key_of_element(_CELL_NAME, ElementMarginProperty.RIGHT_SIZE),
        PropertyDescription('w:tcPr/w:tcMar', ['w:right', 'w:end'], 'w:w'),
        True
    )

    @classmethod
    def _element_key(cls) -> str:
        return _CELL_NAME


@unique
class DrawingProperty(ElementPropertyEnum):
    HORIZONTAL_SIZE = ('horizontal size of drawing', PropertyDescription('wp:inline', 'wp:extent', 'cx'))
    VERTICAL_SIZE = ('vertical size of drawing', PropertyDescription('wp:inline', 'wp:extent', 'cy'))


@unique
class ImageProperty(ElementPropertyEnum):
    ID = ('id of image', PropertyDescription(None, None, 'r:embed'))


@unique
class Element(Enum):
    BODY = (_BODY_NAME, 'w:body', BodyProperty)
    PARAGRAPH = (_PARAGRAPH_NAME, 'w:p', ParagraphProperty)
    RUN = (_RUN_NAME, 'w:r', RunProperty)
    TEXT = (_TEXT_NAME, 'w:t', TextProperty)
    LINE_BREAK = (_LINE_BREAK, 'w:br', LineBreakProperty)
    CARRIAGE_RETURN = (_CARRIAGE_RETURN, 'w:cr', CarriageReturnProperty)
    Tabulation = (_TABULATION, 'w:tab', TabulationProperty)
    TABLE = (_TABLE_NAME, 'w:tbl', TableProperty)
    ROW = (_ROW_NAME, 'w:tr', RowProperty)
    CELL = (_CELL_NAME, 'w:tc', CellProperty)
    DRAWING = (_DRAWING_NAME, 'w:drawing', DrawingProperty)
    IMAGE = (_IMAGE_NAME, 'wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip', ImageProperty)

    def __init__(self, s: str, tag: str, properties_enum):
        self.key: str = s
        self.tag: str = tag
        self.props = properties_enum

    def get_property_descriptions_dict(self) -> Dict[str, PropertyDescription]:
        """
        create keys for dict of property descriptions of elements (those keys use also for corresponding properties)
        """
        return self.props.get_properties_description_dict()


@unique
class StyleProperty(ElementPropertyEnum):
    TYPE = ('type of style', PropertyDescription(None, None, 'w:type'))
    ID = ('id of style', PropertyDescription(None, None, 'w:styleId'))
    DEFAULT = ('is default style', PropertyDescription(None, None, 'w:default'))
    CUSTOM = ('is custom style', PropertyDescription(None, None, 'w:customStyle'))
    BASE_STYLE = ('style based on', PropertyDescription(None, 'w:basedOn', 'w:val'), True)


# substyle is element of other style like style for area of table
@unique
class SubStyle(Enum):
    TABLE_AREA = ('table area style', [
        TableProperty, RowProperty, CellProperty, ParagraphProperty, RunProperty
    ])

    def __init__(self, s: str, property_enums: list):
        self.tag: str = 'w:tblStylePr'
        self.key: str = s
        self._property_enums: list = property_enums

    def get_property_descriptions_dict(self) -> Dict[str, PropertyDescription]:
        """
        create keys for dict of property descriptions of substyles (those keys use also for corresponding properties)
        """
        from ..utils.dicts import merge_dicts

        return merge_dicts(
            [enum.get_properties_description_dict(True) for enum in self._property_enums]
        )


@unique
class TableArea(Enum):
    FIRST_ROW = 'firstRow'
    LAST_ROW = 'lastRow'
    FIRST_COLUMN = 'firstCol'
    LAST_COLUMN = 'lastCol'
    ODD_ROW = 'band1Horz'
    EVEN_ROW = 'band2Horz'
    ODD_COLUMN = 'band1Vert'
    EVEN_COLUMN = 'band2Vert'
    TOP_RIGHT_CELL = 'neCell'
    TOP_LEFT_CELL = 'nwCell'
    BOTTOM_RIGHT_CELL = 'seCell'
    BOTTOM_LEFT_CELL = 'swCell'


@unique
class Style(Enum):
    NUMBERING = ('numbering style', 'numbering', [StyleProperty])
    CHARACTER = ('character style', 'character', [StyleProperty, RunProperty])
    PARAGRAPH = ('paragraph style', 'paragraph', [StyleProperty, ParagraphProperty, RunProperty])
    TABLE = ('table style', 'table', [
        StyleProperty, TableProperty, RowProperty, CellProperty, ParagraphProperty, RunProperty
    ])

    def __init__(self, s: str, style_type: str, property_enums: list):
        self.key: str = s
        self.type: str = style_type
        self._property_enums: list = property_enums

    def get_property_descriptions_dict(self) -> Dict[str, PropertyDescription]:
        """
        create keys for dict of property descriptions of styles (those keys use also for corresponding properties)
        """
        from ..utils.dicts import merge_dicts

        return merge_dicts(
            [enum.get_properties_description_dict(True) for enum in self._property_enums]
        )

    @staticmethod
    def tag() -> str:
        return 'w:style'


@unique
class DefaultStyle(Enum):
    PARAGRAPH = ('default paragraph', Style.PARAGRAPH, Element.PARAGRAPH, 'w:pPrDefault')
    RUN = ('default run', Style.CHARACTER, Element.RUN, 'w:rPrDefault')

    def __init__(self, s: str, style: Style, element: Element, tag: str):
        self.key: str = s
        self._style: Style = style
        self.element: Element = element
        self._tag: str = tag

    def get_property_descriptions_dict(self) -> Dict[str, PropertyDescription]:
        """
        create keys for dict of property descriptions of styles (those keys use also for corresponding properties)
        """
        return self.element.get_property_descriptions_dict()

    def style_type(self) -> str:
        return self._style.type

    def tag(self) -> str:
        return self._tag
