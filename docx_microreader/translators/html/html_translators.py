from typing import Dict, List, Tuple, Union
from abc import ABC


class TranslatorToHTML:
    def __init__(self):
        self.styles: Dict[str, str] = {}
        self.inner_html_wrapper_styles: Dict[str, str] = {}
        self.attributes: Dict[str, str] = {}
        self.ext_tags: List[Tuple[str, bool, bool]] = []

    def _reset_value(self):
        self.styles = {}
        self.inner_html_wrapper_styles = {}
        self.attributes = {}
        self.ext_tags = []

    def translate(self, element, inner_elements: list) -> str:
        self._reset_value()     # it is need, because one translator can using for many objects
        # it is need convert inner element before tacking styles, attributes etc
        # because one translator can using for many objects, but this objects can containing each other
        # In this algorithm styles, attributes etc. takes from inner elements to outher elements
        # and inner elements don't reset this parameters for outher

        inner_text: str = ''.join([str(s) for s in inner_elements])
        self._convert_fields(element)
        self._do_methods(element)
        return self._get_html(inner_text)

    def _convert_fields(self, element):
        pass

    def _do_methods(self, element):
        pass

    def _get_html(self, inner_text: str) -> str:
        open_ext_tags, close_ext_tags = self._get_ext_tags()
        styles: str = self._get_styles()
        attrs: str = self._get_attributes()
        tag: str = self._get_html_tag()

        inner_text_styles: str = self._get_inner_html_wrapper_styles()
        if inner_text_styles != '':
            inner_text: str = rf'<div{inner_text_styles}>{inner_text}</div>'
        if tag == '':
            return rf'{open_ext_tags}{inner_text}{close_ext_tags}'
        if not self._is_single_tag():
            return rf'{open_ext_tags}<{tag}{attrs}{styles}>{inner_text}</{tag}>{close_ext_tags}'
        return rf'{open_ext_tags}<{tag}{attrs}{styles}>{close_ext_tags}'

    def _get_html_tag(self) -> str:
        pass

    @staticmethod
    def _is_single_tag() -> bool:
        return False

    def _get_styles(self) -> str:
        result: str = ''
        for k in self.styles:
            result += rf' {k}: {self.styles[k]};'
        if result != '':
            result = rf' style="{result}"'
        return result

    def _get_attributes(self) -> str:
        result: str = ''
        for k in self.attributes:
            result += rf' {k}="{self.attributes[k]}"'
        return result

    def _add_to_ext_tags(self, tag: str, is_open: bool = True, is_close: bool = True):
        if is_open or is_close:
            self.ext_tags.append((tag, is_open, is_close))

    def _get_ext_tags(self) -> Tuple[str, str]:
        open_tags: str = ''
        close_tags: str = ''
        for tag in self.ext_tags:
            if tag[1]:
                open_tags += rf'<{tag[0]}>'
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
    def _translate_color(color: str) -> str:
        import re
        match = re.match('[0-9A-F]{6}', color)
        if match is not None:
            if color == match.group(0):
                return '#' + color
        return color

    def _add_to_many_properties_style(self, t: Union[Tuple[str, str], None]):
        if t is not None:
            if t[0] not in self.styles:
                self.styles[t[0]] = t[1]
            else:
                self.styles[t[0]] += rf' {t[1]}'


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


class TranslatorBorderedElementToHTML(ABC):
    border_types_corresponding: Dict[str, str] = {
        'nil': 'none',
        'none': 'none',
        'thick': 'solid',
        'single': 'solid',
        'dotted': 'dotted',
        'dashed': 'dashed',
        'dashSmallGap': 'dashed',
        'dotDash': 'dashed',
        'dotDotDash': 'dashed',
        'double': 'double',
        'wave': 'dashed',
        'doubleWave': 'double',
        'triple': 'dashed',
        'thinThickSmallGap': 'double',
        'thickThinSmallGap': 'double',
        'thinThickThinSmallGap': 'double',
        'thinThickMediumGap': 'double',
        'thickThinMediumGap': 'double',
        'thinThickThinLargeGap': 'double',
        'dashDotStroked': 'dotted',
        'threeDEmboss': 'double',
        'threeDEngrave': 'double',
        'inset': 'inset',
        'outset': 'outset',
        'thickThinLargeGap': 'double',
        'thinThickLargeGap': 'double',
        'thinThickThinMediumGap': 'double',
    }

    def _add_to_many_properties_style(self, t: Union[Tuple[str, str], None]):
        raise NotImplementedError('method _add_to_many_properties_style is not implemented. Your TranslatorToHTML class'
                                  'must implement this method or TranslatorBorderedElementToHTML must be inherited'
                                  'after the class that implements this method')

    def _to_css_border(self, element, direction: str) -> Union[Tuple[str, str], None]:
        border = element.get_border(direction, 'type')
        if border is not None:
            return rf'border-{direction}', self.border_types_corresponding.get(border, 'solid')
        return None

    def _to_css_border_color(self, element, direction: str) -> Union[Tuple[str, str], None]:
        border_color = element.get_border(direction, 'color')
        if border_color is not None:
            if border_color == 'auto':
                return rf'border-{direction}', 'black'
            elif border_color is not None:
                return rf'border-{direction}', TranslatorToHTML._translate_color(border_color)
        return None

    def _to_css_border_size(self, element, direction: str) -> Union[Tuple[str, str], None]:
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


class ImageTranslatorToHTML(TranslatorToHTML):

    def _do_methods(self, image):
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


class ParagraphTranslatorToHTML(TranslatorToHTML, TranslatorBorderedElementToHTML):
    aligns: Dict[str, str] = {
        'both': 'justify',
        'right': 'right',
        'center': 'center',
        'left': 'left',
    }

    def _do_methods(self, paragraph):
        self._to_attribute_align(paragraph)
        self._to_css_margin_left(paragraph)
        self._to_css_margin_right(paragraph)
        self._to_css_text_indent(paragraph)
        self._to_css_all_borders(paragraph)

    def _get_html_tag(self) -> str:
        return 'p'

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


class RunTranslatorToHTML(TranslatorToHTML, TranslatorBorderedElementToHTML):
    external_tags: Dict[str, str] = {
        'bold': 'b',
        'italic': 'i',
    }
    underlines: Dict[str, str] = {
        'wave': 'wavy',
        'dash': 'dashed',
        'single': 'solid',
        'thick': 'solid',
        'double': 'double',
        'dotted': 'dotted',
    }

    def __init__(self):
        super(RunTranslatorToHTML, self).__init__()
        self.border: Dict[str, Union[None, str]]
        self.defining_of_border: Dict[str, bool]
        self._set_default_for_border_dicts()

    def _reset_value(self):
        super(RunTranslatorToHTML, self)._reset_value()
        self._set_default_for_border_dicts()

    def _set_default_for_border_dicts(self):
        self.border: Dict[str, Union[None, str]] = {
            'type': None,
            'size': None,
            'color': None,
        }
        self.defining_of_border: Dict[str, bool] = {
            'type': False,
            'size': False,
            'color': False,
        }

    def _do_methods(self, run):
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
            self.styles['background-color'] = TranslatorToHTML._translate_color(background_fill)

    def _to_css_color(self, run):
        color = run.get_color()
        if color is not None:
            self.styles['color'] = TranslatorToHTML._translate_color(color)

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
            self.styles['text-decoration-color'] = TranslatorToHTML._translate_color(underline_color)

    def _get_border_property_of_run(self, run, pr: str) -> Union[str, None]:
        if not self.defining_of_border[pr]:
            self.border[pr] = run.get_border(pr)
            self.defining_of_border[pr] = True
        return self.border[pr]

    def _to_css_border_color(self, run, direction: str) -> Union[Tuple[str, str], None]:
        border_color = self._get_border_property_of_run(run, 'color')
        if border_color is not None:
            if border_color == 'auto':
                return rf'border-{direction}', 'black'
            elif border_color is not None:
                return rf'border-{direction}', TranslatorToHTML._translate_color(border_color)
        return None

    def _to_css_border_size(self, run, direction: str) -> Union[Tuple[str, str], None]:
        border_size = self._get_border_property_of_run(run, 'size')
        if border_size is not None:
            return rf'border-{direction}', str(int(border_size) / 8) + 'pt'
        return None

    def _to_css_border(self, run, direction: str) -> Union[Tuple[str, str], None]:
        border = self._get_border_property_of_run(run, 'type')
        if border is not None:
            return rf'border-{direction}', self.border_types_corresponding.get(border, 'solid')
        return None


class TextTranslatorToHTML:
    characters_html: Dict[str, str] = {
        'Α': '&Alpha;',
        'Β': '&Beta;',
        'Γ': '&Gamma;',
        'Δ': '&Delta;',
        'Ε': '&Epsilon;',
        'Ζ': '&Zeta;',
        'Η': '&Eta;',
        'Θ': '&Theta;',
        'Ι': '&Iota;',
        'Κ': '&Kappa;',
        'Λ': '&Lambda;',
        'Μ': '&Mu;',
        'Ν': '&Nu;',
        'Ξ': '&Xi;',
        'Ο': '&Omicron;',
        'Π': '&Pi;',
        'Ρ': '&Rho;',
        'Σ': '&Sigma;',
        'Τ': '&Tau;',
        'Υ': '&Upsilon;',
        'Φ': '&Phi;',
        'Χ': '&Chi;',
        'Ψ': '&Psi;',
        'Ω': '&Omega;',
        'α': '&alpha;',
        'β': '&beta;',
        'γ': '&gamma;',
        'δ': '&delta;',
        'ε': '&epsilon;',
        'ζ': '&zeta;',
        'η': '&eta;',
        'θ': '&theta;',
        'ι': '&iota;',
        'κ': '&kappa;',
        'λ': '&lambda;',
        'μ': '&mu;',
        'ν': '&nu;',
        'ξ': '&xi;',
        'ο': '&omicron;',
        'π': '&pi;',
        'ρ': '&rho;',
        'ς': '&sigmaf;',
        'σ': '&sigma;',
        'τ': '&tau;',
        'υ': '&upsilon;',
        'φ': '&phi;',
        'χ': '&chi;',
        'ψ': '&psi;',
        'ω': '&omega;',
        'Ω': '&Omega;',
        'µ': '&mu;',
        '©': '&copy;',
        '®': '&reg;',
        '™': '&trade;',
        'º': '&ordm;',
        'ª': '&ordf;',
        '‰': '&permil;',
        '¦': '&brvbar;',
        '§': '&sect;',
        '°': '&deg;',
        '¶': '&para;',
        '…': '&hellip;',
        '‾': '&oline;',
        '´': '&acute;',
        '№': '&#8470;',
        '☎': '&#9742;',
        '×': '&times;',
        '÷': '&divide;',
        '<': '&lt;',
        '>': '&gt;',
        '±': '&plusmn;',
        '¹': '&sup1;',
        '²': '&sup2;',
        '³': '&sup3;',
        '¬': '&not;',
        '¼': '&frac14;',
        '½': '&frac12;',
        '¾': '&frac34;',
        '≤': '&le;',
        '≥': '&ge;',
        '≈': '&asymp;',
        '≠': '&ne;',
        '≡': '&equiv;',
        '√': '&radic;',
        '∞': '&infin;',
        '∑': '&sum;',
        '∏': '&prod;',
        '∂': '&part;',
        '∫': '&int;',
        '∀': '&forall;',
        '∃': '&exist;',
        '∅': '&empty;',
        'Ø': '&Oslash;',
        '∈': '&isin;',
        '∉': '&notin;',
        '∋': '&ni;',
        '⊂': '&sub;',
        '⊃': '&sup;',
        '⊄': '&nsub;',
        '⊆': '&sube;',
        '⊇': '&supe;',
        '⊕': '&oplus;',
        '⊗': '&otimes;',
        '⊥': '&perp;',
        '∠': '&ang;',
        '∧': '&and;',
        '∨': '&or;',
        '∩': '&cap;',
        '∪': '&cup;',
        '←': '&larr;',
        '↑': '&uarr;',
        '→': '&rarr;',
        '↓': '&darr;',
        '↔': '&harr;',
        '↕': '&#8597;',
        '↵': '&crarr;',
        '⇐': '&lArr;',
        '⇑': '&uArr;',
        '⇒': '&rArr;',
        '⇓': '&dArr;',
        '⇔': '&hArr;',
        '⇕': '&#8661;',
        '▲': '&#9650;',
        '▼': '&#9660;',
        '►': '&#9658;',
        '◄': '&#9668;',
        '"': '&quot;',
        '«': '&laquo;',
        '»': '&raquo;',
        '‹': '&#8249;',
        '›': '&#8250;',
        '′': '&prime;',
        '″': '&Prime;',
        '‘': '&lsquo;',
        '’': '&rsquo;',
        '‚': '&sbquo;',
        '“': '&ldquo;',
        '”': '&rdquo;',
        '„': '&bdquo;',
        '❛': '&#10075;',
        '❜': '&#10076;',
        '❝': '&#10077;',
        '❞': '&#10078;',
        '•': '&bull;',
        '○': '&#9675;',
        '·': '&middot;',
        '−': '-',
        'ß': '&beta;',
        '\'': '&#39;',
        '£': '&pound;',
        '¯': '&macr;',
        '¨': '&uml;',
        '¥': '&yen;',
        '¢': '&cent;',
        '¿': '&iquest;',
        'Ґ': '&#1168',
        'ґ': '&#1169',
    }

    def translate(self, text_element, inner_elements: list) -> str:
        import re

        text = text_element.content
        for char in self.characters_html:
            text = re.sub(char, self.characters_html[char], text)
        return text


class LineBreakTranslatorToHTML:
    def translate(self, text_element, inner_elements: list) -> str:
        return '<br>'


class CarriageReturnTranslatorToHTML:
    def translate(self, text_element, inner_elements: list) -> str:
        return '<br>'


class TabulationTranslatorToHTML:
    def translate(self, text_element, inner_elements: list) -> str:
        return '&emsp;&emsp;'


class TableTranslatorToHTML(TranslatorToHTML, TranslatorBorderedElementToHTML):

    def _do_methods(self, table):
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

    def _do_methods(self, row):
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


class CellTranslatorToHTML(TranslatorToHTML, TranslatorBorderedElementToHTML):
    text_directions: Dict[str, str] = {
        'btLr': 'vertical-rl',
        'tbRl': 'vertical-rl',
    }

    def __init__(self):
        super(CellTranslatorToHTML, self).__init__()
        self._is_header: bool = False
        self._tag_for_header = 'td'

    def translate(self, element, inner_elements: list) -> str:
        from ...constants.property_enums import CellProperty
        from ...properties import Property
        if isinstance(element.get_property(CellProperty.VERTICAL_MERGE, False), Property.Missed):
            return ''
        return super(CellTranslatorToHTML, self).translate(element, inner_elements)

    def _do_methods(self, cell):
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
                self.styles['background-color'] = TranslatorToHTML._translate_color(color)

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
