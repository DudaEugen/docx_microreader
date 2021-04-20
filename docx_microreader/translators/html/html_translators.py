from typing import Dict, List, Tuple, Optional
from .mixins import BorderedElementToHTMLMixin, ParagraphContainerMixin


class TranslatorToHTML(ParagraphContainerMixin):
    def __init__(self):
        self.styles: Dict[str, str] = {}
        self.inner_html_wrapper_styles: Dict[str, str] = {}
        self.attributes: Dict[str, str] = {}
        self.ext_tags: List[Tuple[str, bool, bool]] = []
        self.ext_tags_attributes: Dict[str, Dict[str, str]] = {}
        self.ext_tags_styles: Dict[str, Dict[str, str]] = {}
        self.inner_text: str = ''

    def _reset_value(self):
        self.styles = {}
        self.inner_html_wrapper_styles = {}
        self.attributes = {}
        self.ext_tags = []
        self.ext_tags_attributes = {}
        self.ext_tags_styles = {}
        self.inner_text = ''

    def translate(self, element, inner_elements: list, context: dict) -> str:
        self.inner_text = ''.join([str(s) for s in inner_elements])
        self._convert_fields(element)
        self._do_methods(element, context)
        result: str = self._get_html()
        self._reset_value()
        return result

    def _convert_fields(self, element):
        pass

    def _do_methods(self, element, context: dict):
        pass

    def _get_html(self) -> str:
        open_ext_tags, close_ext_tags = self._get_ext_tags()
        styles: str = self._get_styles()
        attrs: str = self._get_attributes()
        tag: str = self._get_html_tag()

        inner_text_styles: str = self._get_inner_html_wrapper_styles()
        if inner_text_styles != '':
            self.inner_text: str = rf'<div{inner_text_styles}>{self.inner_text}</div>'
        if tag == '':
            return rf'{open_ext_tags}{self.inner_text}{close_ext_tags}'
        if not self._is_single_tag():
            return rf'{open_ext_tags}<{tag}{attrs}{styles}>{self.inner_text}</{tag}>{close_ext_tags}'
        return rf'{open_ext_tags}<{tag}{attrs}{styles}>{close_ext_tags}'

    def _get_html_tag(self) -> str:
        pass

    @staticmethod
    def _is_single_tag() -> bool:
        return False

    def _get_styles(self, key: Optional[str] = None) -> str:
        """
        :param key: external tag. Return styles for tag if None
        """
        styles = self.styles if key is None else self.ext_tags_styles[key]
        result: str = ''
        for k in styles:
            result += rf' {k}: {styles[k]};'
        if result != '':
            result = rf' style="{result}"'
        return result

    def _get_attributes(self, key: Optional[str] = None) -> str:
        """
        :param key: external tag. Return styles for tag if None
        """
        attributes = self.attributes if key is None else self.ext_tags_attributes[key]
        result: str = ''
        for k in attributes:
            result += rf' {k}="{attributes[k]}"'
        return result

    def _add_to_ext_tags(self, tag: str, is_open: bool = True, is_close: bool = True):
        if is_open or is_close:
            self.ext_tags.append((tag, is_open, is_close))
            self.ext_tags_attributes[tag] = {}
            self.ext_tags_styles[tag] = {}

    def _ext_tags_list(self) -> List[str]:
        return [tag for tag, b1, b2 in self.ext_tags]

    def _get_ext_tags(self) -> Tuple[str, str]:
        open_tags: str = ''
        close_tags: str = ''
        for tag in self.ext_tags:
            if tag[1]:
                open_tags += rf'<{tag[0]}{self._get_attributes(tag[0])}{self._get_styles(tag[0])}>'
            if tag[2]:
                close_tags = rf'</{tag[0]}>' + close_tags
        return open_tags, close_tags

    def _get_inner_html_wrapper_styles(self) -> str:
        result: str = ''
        for k in self.inner_html_wrapper_styles:
            result += rf' {k}: {self.inner_html_wrapper_styles[k]};'
        if result != '':
            result = rf' style="{result}"'
        return result

    @staticmethod
    def translate_color(color: str) -> str:
        import re
        match = re.match('[0-9A-F]{6}', color)
        if match is not None:
            if color == match.group(0):
                return '#' + color
        return color

    def _add_to_many_properties_style(self, t: Optional[Tuple[str, str]]):
        if t is not None:
            if t[0] not in self.styles:
                self.styles[t[0]] = t[1]
            else:
                self.styles[t[0]] += rf' {t[1]}'

    def preparation_to_translate_inner_elements(self, element, translation_context: dict):
        from ...models import Paragraph
        
        if element.__class__.is_possible_inner_element(Paragraph):
            self._mark_numbering_paragraphs(element, translation_context)


class ContainerTranslatorToHTML(TranslatorToHTML):
    """
    displayed only containing elements
    """
    def _get_html_tag(self) -> str:
        return ''

    @staticmethod
    def _is_single_tag() -> bool:
        return True


class DocumentTranslatorToHTML(TranslatorToHTML):
    def _get_html_tag(self) -> str:
        return 'html'


class BodyTranslatorToHTML(TranslatorToHTML):
    def _get_html_tag(self) -> str:
        return 'body'


class ImageTranslatorToHTML(TranslatorToHTML):

    def _do_methods(self, image, context: dict):
        self._to_attribute_src(image)
        self._to_attribute_width(image)
        self._to_attribute_height(image)

    def _get_html_tag(self) -> str:
        return 'img'

    @staticmethod
    def _is_single_tag() -> bool:
        return True

    def _to_attribute_src(self, image):
        self.attributes['src'] = image.get_path()

    @staticmethod
    def _emu_to_px(value: int) -> int:
        return value // 12700

    def _to_attribute_width(self, image):
        self.attributes['width'] = str(ImageTranslatorToHTML._emu_to_px(image.get_size()[0]))

    def _to_attribute_height(self, image):
        self.attributes['height'] = str(ImageTranslatorToHTML._emu_to_px(image.get_size()[1]))


class ParagraphTranslatorToHTML(TranslatorToHTML, BorderedElementToHTMLMixin):
    from .docx_html_correspondings import align
    aligns: Dict[str, str] = align

    def __init__(self):
        super(ParagraphTranslatorToHTML, self).__init__()
        self.tag: str = 'p'
        self.is_first_paragraph_in_numbering: bool = False

    def _reset_value(self):
        super(ParagraphTranslatorToHTML, self)._reset_value()
        self.tag: str = 'p'
        self.is_first_paragraph_in_numbering = False

    def _do_methods(self, paragraph, context: dict):
        self._to_numbering(paragraph, context)
        if self.is_first_paragraph_in_numbering:
            level = paragraph.get_numbering_level()
            if 'ol' in self._ext_tags_list():
                self._to_ext_css_numbering_format(level)
                self._to_ext_attribute_start(level)
        else:
            self._to_attribute_align(paragraph)
            self._to_css_margin_left(paragraph)
            self._to_css_margin_right(paragraph)
            self._to_css_text_indent(paragraph)
            self._to_css_all_borders(paragraph)
            self._to_numbering(paragraph, context)

    def _get_html_tag(self) -> str:
        return self.tag

    def _to_attribute_align(self, paragraph):
        align = paragraph.get_align()
        if align is not None:
            if align in self.aligns and align != 'left':    # left is default in browsers
                self.attributes['align'] = self.aligns[align]

    def _to_css_margin_left(self, paragraph):
        margin_left = paragraph.get_indent_left()
        if margin_left is not None:
            self.styles['margin-left'] = str(int(margin_left) // 20) + 'px'

    def _to_css_margin_right(self, paragraph):
        margin_right = paragraph.get_indent_right()
        if margin_right is not None:
            self.styles['margin-left'] = str(int(margin_right) // 20) + 'px'

    def _to_css_text_indent(self, paragraph):
        hanging = paragraph.get_hanging()
        first_line = paragraph.get_first_line()
        if hanging is not None:
            self.styles['text-indent'] = str(-int(hanging) // 20) + 'px'
        elif first_line is not None:
            self.styles['text-indent'] = str(int(first_line) // 20) + 'px'

    def _to_numbering(self, paragraph, context: dict):
        from .marks import ParagraphMark

        context_of_paragraph = context.get(paragraph)
        numbering_level = paragraph.get_numbering_level()
        if context_of_paragraph is not None and numbering_level is not None:
            self.is_first_paragraph_in_numbering = ParagraphMark.FIRST_ELEMENT_NUMBERING in context_of_paragraph
            level_num_format = numbering_level.get_numbering_format()
            self._add_to_ext_tags('ul' if level_num_format == 'bullet' or level_num_format == 'chicago' else 'ol',
                                  is_open=self.is_first_paragraph_in_numbering,
                                  is_close=ParagraphMark.LAST_ELEMENT_NUMBERING in context_of_paragraph)
        if numbering_level is not None:
            self.tag = 'li'

    def _to_ext_css_numbering_format(self, level):
        from .docx_html_correspondings import numbering_formats

        level_num_format: Optional[str] = numbering_formats.get(level.get_numbering_format())
        if level_num_format is not None:
            self.ext_tags_styles['ol']['list-style-type'] = level_num_format

    def _to_ext_attribute_start(self, level):
        start: str = level.get_start()
        if start != '1':
            self.ext_tags_attributes['ol']['start'] = start


class RunTranslatorToHTML(TranslatorToHTML, BorderedElementToHTMLMixin):
    from .docx_html_correspondings import text_typeface, underline
    external_tags: Dict[str, str] = text_typeface
    underlines: Dict[str, str] = underline

    def __init__(self):
        super(RunTranslatorToHTML, self).__init__()
        self.border: Dict[str, Optional[str]]
        self.defining_of_border: Dict[str, bool]
        self._set_default_for_border_dicts()

    def _reset_value(self):
        super(RunTranslatorToHTML, self)._reset_value()
        self._set_default_for_border_dicts()

    def _set_default_for_border_dicts(self):
        self.border: Dict[str, Optional[str]] = {
            'type': None,
            'size': None,
            'color': None,
        }
        self.defining_of_border: Dict[str, bool] = {
            'type': False,
            'size': False,
            'color': False,
        }

    def _do_methods(self, run, context: dict):
        self._to_css_font(run)
        self._to_css_size(run)
        self._to_ext_tag_bold(run)
        self._to_ext_tag_italic(run)
        self._to_ext_tags_vertical_align(run)
        self._to_css_background_color(run)
        self._to_css_color(run)
        self._to_css_line_throught(run)
        self._to_css_underline(run)
        self._to_css_all_borders(run)

    def _get_html_tag(self) -> str:
        if self.attributes or self.styles:
            return 'span'
        return ''

    def _to_css_size(self, run):
        size = run.get_size()
        if size is not None:
            self.styles['font-size'] = size

    def _to_css_font(self, run):
        font = run.get_font()
        if font is not None:
            self.styles['font-family'] = font

    def _to_ext_tag_bold(self, run):
        if run.is_bold():
            self._add_to_ext_tags(self.external_tags['bold'])

    def _to_ext_tag_italic(self, run):
        if run.is_italic():
            self._add_to_ext_tags(self.external_tags['italic'])

    def _to_ext_tags_vertical_align(self, run):
        vert_align = run.get_vertical_align()
        if vert_align is not None:
            if vert_align == 'subscript':
                self._add_to_ext_tags('sub')
            elif vert_align == 'superscript':
                self._add_to_ext_tags('sup')

    def _to_css_background_color(self, run):
        background_color = run.get_background_color()
        background_fill = run.get_background_fill()
        if background_color is not None:
            if background_color != 'none':
                self.styles['background-color'] = background_color
        elif background_fill is not None:
            self.styles['background-color'] = TranslatorToHTML.translate_color(background_fill)

    def _to_css_color(self, run):
        color = run.get_color()
        if color is not None:
            self.styles['color'] = TranslatorToHTML.translate_color(color)

    def _to_css_line_throught(self, run):
        if run.is_strike():
            self._add_to_many_properties_style(('text-decoration', 'line-through'))

    def _to_css_underline(self, run):
        underline = run.get_underline()
        if underline is not None:
            self._add_to_many_properties_style(('text-decoration', 'underline'))
            if underline in self.underlines:
                self.styles['text-decoration-style'] = self.underlines[underline]
            self._to_css_underline_color(run)

    def _to_css_underline_color(self, run):
        """
        must called after self.__to_css_underline
        """
        underline_color = run.get_underline_color()
        if underline_color is not None:
            self.styles['text-decoration-color'] = TranslatorToHTML.translate_color(underline_color)

    def _get_border_property_of_run(self, run, pr: str) -> Optional[str]:
        if not self.defining_of_border[pr]:
            self.border[pr] = run.get_border(pr)
            self.defining_of_border[pr] = True
        return self.border[pr]

    def _to_css_border_color(self, run, direction: str) -> Optional[Tuple[str, str]]:
        border_color = self._get_border_property_of_run(run, 'color')
        if border_color is not None:
            if border_color == 'auto':
                return rf'border-{direction}', 'black'
            elif border_color is not None:
                return rf'border-{direction}', TranslatorToHTML.translate_color(border_color)
        return None

    def _to_css_border_size(self, run, direction: str) -> Optional[Tuple[str, str]]:
        border_size = self._get_border_property_of_run(run, 'size')
        if border_size is not None:
            return rf'border-{direction}', str(int(border_size) / 8) + 'pt'
        return None

    def _to_css_border(self, run, direction: str) -> Optional[Tuple[str, str]]:
        border = self._get_border_property_of_run(run, 'type')
        if border is not None:
            return rf'border-{direction}', self.border_types_corresponding.get(border, 'solid')
        return None


class TextTranslatorToHTML(TranslatorToHTML):
    from .docx_html_correspondings import character
    characters_html: Dict[str, str] = character

    def translate(self, text_element, inner_elements: list, context: dict) -> str:
        import re

        text = text_element.content
        for char in self.characters_html:
            text = re.sub(char, self.characters_html[char], text)
        return text


class LineBreakTranslatorToHTML(TranslatorToHTML):
    def translate(self, text_element, inner_elements: list, context: dict) -> str:
        return '<br>'


class CarriageReturnTranslatorToHTML(TranslatorToHTML):
    def translate(self, text_element, inner_elements: list, context: dict) -> str:
        return '<br>'


class TabulationTranslatorToHTML(TranslatorToHTML):
    def translate(self, text_element, inner_elements: list, context: dict) -> str:
        return '&emsp;&emsp;'


class NoBreakHyphenTranslatorToHTML(TranslatorToHTML):
    def translate(self, text_element, inner_elements: list, context: dict) -> str:
        return '&#8209;'


class SoftHyphenTranslatorToHTML(TranslatorToHTML):
    def translate(self, text_element, inner_elements: list, context: dict) -> str:
        return '&shy;'


class SymbolTranslatorToHTML(TranslatorToHTML):
    def _do_methods(self, symbol, context: dict):
        self._to_css_font(symbol)
        self._get_char(symbol)

    def _get_html_tag(self) -> str:
        if self.styles:
            return 'span'
        return ''

    def _to_css_font(self, symbol):
        font = symbol.get_font()
        if font is not None:
            self.styles['font-family'] = font

    def _get_char(self, symbol):
        char = symbol.get_char()
        if char is not None:
            char_code = int(char, 16)
            if char_code > int('F000', 16):
                char_code -= int('F000', 16)
            self.inner_text = f'&#{char_code};'


class TableTranslatorToHTML(TranslatorToHTML, BorderedElementToHTMLMixin):

    def _do_methods(self, table, context: dict):
        self._to_css_border_collapse()
        self._to_attribute_width(table)
        self._to_attribute_align(table)
        self._to_css_all_borders(table)
        self._to_css_margin(table)

    def _get_html_tag(self) -> str:
        return 'table'

    def _to_css_border_collapse(self):
        self.styles['border-collapse'] = 'collapse'

    def _to_attribute_width(self, table):
        width, width_type, layout = table.get_width()
        if width is not None and width_type is not None:
            if (layout is not None and layout != 'autofit') or layout is None:
                if width_type == 'dxa':
                    self.attributes['width'] = str(int(width) // 20) + 'px'
                elif width_type == 'pct':
                    self.attributes['width'] = str(int(width) / 50) + '%'
                elif width_type == 'nil':
                    self.attributes['width'] = '0pt'
                elif width_type == 'auto':
                    return

    def _to_attribute_align(self, table):
        align = table.get_align()
        if align is not None:
            self.attributes['align'] = align

    def _to_css_margin(self, table):
        align = table.get_align()
        if align is None or align != 'center' and align != 'right':
            margin, indentation_type = table.get_indentation()
            if margin is not None and indentation_type is not None:
                if indentation_type == 'dxa':
                    self.styles['margin-left'] = str(int(margin) // 20) + 'px'
                elif indentation_type == 'nil':
                    self.styles['margin-left'] = '0px'
                elif indentation_type == 'pct':
                    return
                elif indentation_type == 'auto':
                    return


class RowTranslatorToHTML(TranslatorToHTML):

    def _do_methods(self, row, context: dict):
        self._to_attribute_or_css_height(row)

    def _convert_fields(self, row):
        self._to_ext_tag_head(row)

    def _get_html_tag(self) -> str:
        return 'tr'

    def _to_ext_tag_head(self, row):
        if row.is_header():
            self._add_to_ext_tags('thead', row.is_first_row_in_header, row.is_last_row_in_header)

    def _to_attribute_or_css_height(self, row):
        height, height_type = row.get_height()
        if height is not None:
            if height_type is None:
                self.attributes['height'] = str(int(height) // 20) + 'px'
            elif height_type == 'exact':
                self.attributes['height'] = str(int(height) // 20) + 'px'
            elif height_type == 'atLeast':
                self.styles['min-height'] = str(int(height) // 20) + 'px'
            elif height_type == 'auto':
                return


class CellTranslatorToHTML(TranslatorToHTML, BorderedElementToHTMLMixin):
    from .docx_html_correspondings import text_direction
    text_directions: Dict[str, str] = text_direction

    def __init__(self):
        super(CellTranslatorToHTML, self).__init__()
        self._is_header: bool = False
        self._tag_for_header = 'td'

    def translate(self, element, inner_elements: list, context: dict) -> str:
        from ...constants.property_enums import CellProperty
        from ...properties import Property
        if isinstance(element.get_property(CellProperty.VERTICAL_MERGE, False), Property.Missed):
            return ''
        return super(CellTranslatorToHTML, self).translate(element, inner_elements, context)

    def _do_methods(self, cell, context: dict):
        self._to_css_fill_color(cell)
        self._to_css_all_borders(cell)
        self._to_attribute_width(cell)
        self._to_attribute_col_span(cell)
        self._to_css_text_direction(cell)
        self._to_attribute_vertical_align(cell)
        self._to_css_all_padding(cell)

    def _convert_fields(self, cell):
        self._to_attribute_row_span(cell)
        self._is_header = cell.is_header

    def _get_html_tag(self) -> str:
        return 'td' if not self._is_header else self._tag_for_header

    def is_th_for_header(self, value: bool = True):
        self._tag_for_header = 'th' if value is True else 'td'

    def _to_css_fill_color(self, cell):
        color = cell.get_fill_color()
        if color is not None:
            if color != 'auto':
                self.styles['background-color'] = TranslatorToHTML.translate_color(color)

    def _to_css_all_padding(self, cell):
        self._to_css_padding_top(cell)
        self._to_css_padding_bottom(cell)
        self._to_css_padding_left(cell)
        self._to_css_padding_right(cell)

    def _to_css_padding_top(self, cell):
        padding = cell.get_margin('top')[0]
        if padding is not None:
            self.styles['padding-top'] = padding
            self._to_css_padding_type(cell, 'top')

    def _to_css_padding_bottom(self, cell):
        padding = cell.get_margin('bottom')[0]
        if padding is not None:
            self.styles['padding-bottom'] = padding
            self._to_css_padding_type(cell, 'bottom')

    def _to_css_padding_left(self, cell):
        padding = cell.get_margin('left')[0]
        if padding is not None:
            self.styles['padding-left'] = padding
            self._to_css_padding_type(cell, 'left')

    def _to_css_padding_right(self, cell):
        padding = cell.get_margin('right')[0]
        if padding is not None:
            self.styles['padding-right'] = padding
            self._to_css_padding_type(cell, 'right')

    def _to_css_padding_type(self, cell, direction: str):
        """
        self.__to_css_padding_... need run before this method
        """
        padding_type = cell.get_margin(direction)[1]
        if padding_type is not None:
            if padding_type == 'dxa':
                self.styles[rf'padding-{direction}'] = \
                    str(int(self.styles[rf'padding-{direction}']) // 20) + 'px'
            elif padding_type == 'nil':
                self.styles[rf'padding-{direction}'] = '0px'
            else:
                if rf'padding-{direction}' in self.attributes:
                    self.styles.pop(rf'padding-{direction}')
        else:
            if rf'padding-{direction}' in self.attributes:
                self.styles.pop(rf'padding-{direction}')

    def _to_attribute_width(self, cell):
        width, width_type = cell.get_width()
        if width is not None and width_type is not None:
            if width_type == 'dxa':
                self.attributes['width'] = str(int(width) // 20) + 'px'
            elif width_type == 'pct':
                self.attributes['width'] = str(int(width) / 50) + '%'
            elif width_type == 'nil':
                self.attributes['width'] = '0px'
            elif width_type == 'auto':
                return

    def _to_attribute_col_span(self, cell):
        col_span = cell.get_col_span()
        if col_span is not None:
            if int(col_span) > 1:
                self.attributes['colspan'] = col_span

    def _to_attribute_row_span(self, cell):
        if cell.row_span > 1:
            self.attributes['rowspan'] = str(cell.row_span)

    def _to_css_text_direction(self, cell):
        text_direction = cell.get_text_direction()
        if text_direction is not None:
            if text_direction in self.text_directions:
                self.inner_html_wrapper_styles['writing-mode'] = self.text_directions[text_direction]

    def _to_attribute_vertical_align(self, cell):
        vertical_align = cell.get_vertical_align()
        if vertical_align is not None:
            self.attributes['valign'] = vertical_align
