from enum import Enum, unique, auto


@unique
class ParagraphMark(Enum):
    FIRST_ELEMENT_NUMBERING = auto()
    LAST_ELEMENT_NUMBERING = auto()
