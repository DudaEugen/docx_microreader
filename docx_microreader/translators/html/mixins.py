from abc import ABC, abstractmethod
from typing import Optional, Dict, Tuple, List


class BorderedElementToHTMLMixin(ABC):
    from .docx_html_correspondings import border_type
    border_types_corresponding: Dict[str, str] = border_type

    @abstractmethod
    def _add_to_many_properties_style(self, t: Optional[Tuple[str, str]]):
        pass

    def _to_css_border(self, element, direction: str) -> Optional[Tuple[str, str]]:
        border = element.get_border(direction, 'type')
        if border is not None:
            return rf'border-{direction}', self.border_types_corresponding.get(border, 'solid')
        return None

    def _to_css_border_color(self, element, direction: str) -> Optional[Tuple[str, str]]:
        from .html_translators import TranslatorToHTML

        border_color = element.get_border(direction, 'color')
        if border_color is not None:
            if border_color == 'auto':
                return rf'border-{direction}', 'black'
            elif border_color is not None:
                return rf'border-{direction}', TranslatorToHTML.translate_color(border_color)
        return None

    def _to_css_border_size(self, element, direction: str) -> Optional[Tuple[str, str]]:
        border_size = element.get_border(direction, 'size')
        if border_size is not None:
            return rf'border-{direction}', str(int(border_size) / 8) + 'pt'
        return None

    def _to_css_border_top(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'top'))
        self._add_to_many_properties_style(self._to_css_border_color(element, 'top'))
        self._add_to_many_properties_style(self._to_css_border_size(element, 'top'))

    def _to_css_border_bottom(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'bottom'))
        self._add_to_many_properties_style(self._to_css_border_color(element, 'bottom'))
        self._add_to_many_properties_style(self._to_css_border_size(element, 'bottom'))

    def _to_css_border_left(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'left'))
        self._add_to_many_properties_style(self._to_css_border_color(element, 'left'))
        self._add_to_many_properties_style(self._to_css_border_size(element, 'left'))

    def _to_css_border_right(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'right'))
        self._add_to_many_properties_style(self._to_css_border_color(element, 'right'))
        self._add_to_many_properties_style(self._to_css_border_size(element, 'right'))

    def _to_css_all_borders(self, element):
        self._to_css_border_top(element)
        self._to_css_border_bottom(element)
        self._to_css_border_left(element)
        self._to_css_border_right(element)


class ParagraphContainerMixin:
    def _mark_numbering_paragraphs(self, element, translation_context: dict):
        from ...models import Paragraph
        from ...numbering import NumberingLevel

        num_levels: List[Tuple[NumberingLevel, Paragraph]] = []  # (level, current last Paragraph of this level)
        previous_paragraph: Optional[Paragraph] = None  # None if previous element is not Paragraph
        for el in element.iterate_by_inner_elements():
            if isinstance(el, Paragraph):
                numbering_level: Optional[NumberingLevel] = el.get_numbering_level()
                if numbering_level is not None:
                    if len(num_levels) > 0:
                        if numbering_level.get_index() > num_levels[-1][0].get_index():
                            ParagraphContainerMixin._numbering_begin(el, num_levels, translation_context)
                        elif numbering_level.get_index() == num_levels[-1][0].get_index():
                            num_levels[-1] = (num_levels[-1][0], el)
                        else:
                            while len(num_levels) > 0 and numbering_level.get_index() < num_levels[-1][0].get_index():
                                ParagraphContainerMixin._numbering_ended(num_levels, translation_context)
                            num_levels[-1] = (num_levels[-1][0], el)
                    else:
                        ParagraphContainerMixin._numbering_begin(el, num_levels, translation_context)
                else:
                    ParagraphContainerMixin._all_numbering_ended(num_levels, translation_context)
                previous_paragraph = el
            else:
                if previous_paragraph is not None and previous_paragraph.get_numbering_level() is not None:
                    ParagraphContainerMixin._all_numbering_ended(num_levels, translation_context)
                previous_paragraph = None
        ParagraphContainerMixin._all_numbering_ended(num_levels, translation_context)

    @staticmethod
    def _add_to_context(paragraph, value, translation_context: dict):
        if paragraph in translation_context:
            translation_context[paragraph].append(value)
        else:
            translation_context[paragraph] = [value]

    @staticmethod
    def _numbering_ended(numbering_levels: List[tuple], translation_context: dict):
        from .marks import ParagraphMark
        ParagraphContainerMixin._add_to_context(numbering_levels[-1][1], ParagraphMark.LAST_ELEMENT_NUMBERING,
                                                translation_context)
        numbering_levels.pop()

    @staticmethod
    def _numbering_begin(first_paragraph, numbering_levels: List[tuple], translation_context: dict):
        from .marks import ParagraphMark
        ParagraphContainerMixin._add_to_context(first_paragraph, ParagraphMark.FIRST_ELEMENT_NUMBERING,
                                                translation_context)
        numbering_levels.append((first_paragraph.get_numbering_level(), first_paragraph))

    @staticmethod
    def _all_numbering_ended(numbering_levels: List[tuple], translation_context: dict):
        while len(numbering_levels) > 0:
            ParagraphContainerMixin._numbering_ended(numbering_levels, translation_context)
