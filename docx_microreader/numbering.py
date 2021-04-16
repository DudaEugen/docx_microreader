from .constants import property_enums as pr_const
from .xml_element import XMLement
from typing import Union, Optional


class NumberingLevel(XMLement):
    element_description = pr_const.Element.NUMBERING_LEVEL

    def get_id(self) -> str:
        return self._properties[pr_const.NumberingLevelProperty.INDEX.key].value


class AbstractNumbering(XMLement):
    element_description = pr_const.Element.ABSTRACT_NUMBERING

    @classmethod
    def _possible_inner_elements_descriptions(cls) -> list:
        return [NumberingLevel]

    def get_id(self) -> str:
        return self._properties[pr_const.AbstractNumberingProperty.ID.key].value

    def get_level(self, index: Union[str, int]) -> Optional[NumberingLevel]:
        for level in self._inner_elements:
            if level.get_id() == str(index):
                return level


class Numbering(XMLement):
    element_description = pr_const.Element.NUMBERING

    def get_id(self) -> str:
        return self._properties[pr_const.NumberingProperty.ID.key].value

    def get_abstract_numbering_id(self) -> str:
        return self._properties[pr_const.NumberingProperty.ABSTRACT_NUMBERING.key].value
