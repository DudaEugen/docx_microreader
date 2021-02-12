from enum import Enum, unique


@unique
class TranslateFormat(Enum):
    HTML = 'html'
    XML = 'xml'
