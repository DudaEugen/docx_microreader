from typing import Dict, Callable, List, Tuple, Union


class TranslatorToHTML:
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

    def __init__(self):
        self.methods: Dict[str, Callable] = {}
        self.styles: Dict[str, str] = {}
        self.inner_html_wrapper_styles: Dict[str, str] = {}
        self.attributes: Dict[str, str] = {}
        self.ext_tags: List[Tuple[str, bool, bool]] = []

    def _reset_value(self):
        self.styles = {}
        self.inner_html_wrapper_styles = {}
        self.attributes = {}
        self.ext_tags = []

    def translate(self, element) -> str:
        self._reset_value()     # it is need, because one translator can using for many objects
        inner_text: str = element.get_inner_text()
        # it is need convert inner element before tacking styles, attributes etc
        # because one translator can using for many objects, but this objects can containing each other
        # In this algorithm styles, attributes etc. takes from inner elements to outher elements
        # and inner elements don't reset this parameters for outher

        self._convert_fields(element)
        self._do_methods(element)
        return self._get_html(inner_text)

    def _convert_fields(self, element):
        pass

    def _do_methods(self, element):
        for k in self.methods:
            self.methods[k](element)

    def _get_html(self, inner_text: str) -> str:
        open_ext_tags, close_ext_tags = self._get_ext_tags()
        styles: str = self._get_styles()
        attrs: str = self._get_attributes()
        tag: str = self._get_html_tag()

        inner_text_styles: str = self._get_inner_html_wrapper_styles()
        if inner_text_styles != '':
            inner_text: str = rf'<div{inner_text_styles}>{inner_text}</div>'
        return rf'{open_ext_tags}<{tag}{attrs}{styles}>{inner_text}</{tag}>{close_ext_tags}' \
            if tag != '' else \
               rf'{open_ext_tags}{inner_text}{close_ext_tags}'

    def _get_html_tag(self) -> str:
        pass

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

    def _pass(self, pr):
        pass

    def _add_to_many_properties_style(self, t: Union[Tuple[str, str], None]):
        if t is not None:
            if t[0] not in self.styles:
                self.styles[t[0]] = t[1]
            else:
                self.styles[t[0]] += rf' {t[1]}'

    def _to_css_border(self, element, direction: str) -> Union[Tuple[str, str], None]:
        border = element.get_border(direction, 'type')
        if border is not None:
            return rf'border-{direction}', self.border_types_corresponding.get(border, 'solid')
        return None

    def _to_css_border_top(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'top'))

    def _to_css_border_bottom(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'bottom'))

    def _to_css_border_left(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'left'))

    def _to_css_border_right(self, element):
        self._add_to_many_properties_style(self._to_css_border(element, 'right'))

    @staticmethod
    def _to_css_border_color(element, direction: str) -> Union[Tuple[str, str], None]:
        border_color = element.get_border(direction, 'color')
        if border_color is not None:
            if border_color == 'auto':
                return rf'border-{direction}', 'black'
            return rf'border-{direction}', CellTranslatorToHTML._translate_color(border_color)
        return None

    def _to_css_border_top_color(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_color(element, 'top'))

    def _to_css_border_bottom_color(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_color(element, 'bottom'))

    def _to_css_border_left_color(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_color(element, 'left'))

    def _to_css_border_right_color(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_color(element, 'right'))

    @staticmethod
    def _to_css_border_size(cell, direction: str) -> Union[Tuple[str, str], None]:
        border_size = cell.get_border(direction, 'size')
        if border_size is not None:
            return rf'border-{direction}', str(int(border_size) / 8) + 'pt'
        return None

    def _to_css_border_top_size(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_size(element, 'top'))

    def _to_css_border_bottom_size(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_size(element, 'bottom'))

    def _to_css_border_left_size(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_size(element, 'left'))

    def _to_css_border_right_size(self, element):
        self._add_to_many_properties_style(TranslatorToHTML._to_css_border_size(element, 'right'))


class ParagraphTranslatorToHTML(TranslatorToHTML):
    aligns: Dict[str, str] = {
        'both': 'justify',
        'right': 'right',
        'center': 'center',
        'left': 'left',
    }

    def __init__(self):
        super(ParagraphTranslatorToHTML, self).__init__()
        self.methods: Dict[str, Callable] = {
            'w:pPr/w:jc/w:val': self.__to_attribute_align,
            'w:pPr/w:ind/w:left': self.__to_css_margin_left,
            'w:pPr/w:ind/w:right': self.__to_css_margin_right,
            'w:pPr/w:ind/w:hanging': self.__to_css_text_indent,
            'w:tcPr/w:pBdr/w:top/w:val': self._to_css_border_top,
            'w:tcPr/w:pBdr/w:bottom/w:val': self._to_css_border_bottom,
            'w:tcPr/w:pBdr/w:left/w:val': self._to_css_border_left,
            'w:tcPr/w:pBdr/w:right/w:val': self._to_css_border_right,
            'w:tcPr/w:pBdr/w:top/w:color': self._to_css_border_top_color,
            'w:tcPr/w:pBdr/w:bottom/w:color': self._to_css_border_bottom_color,
            'w:tcPr/w:pBdr/w:left/w:color': self._to_css_border_left_color,
            'w:tcPr/w:pBdr/w:right/w:color': self._to_css_border_right_color,
            'w:tcPr/w:pBdr/w:top/w:sz': self._to_css_border_top_size,
            'w:tcPr/w:pBdr/w:bottom/w:sz': self._to_css_border_bottom_size,
            'w:tcPr/w:pBdr/w:left/w:sz': self._to_css_border_left_size,
            'w:tcPr/w:pBdr/w:right/w:sz': self._to_css_border_right_size,
        }

    def _get_html_tag(self) -> str:
        return 'p'

    def __to_attribute_align(self, paragraph):
        align = paragraph.get_align()
        if align is not None:
            if align in self.aligns and align != 'left':    # left is default in browsers
                self.attributes['align'] = self.aligns[align]

    def __to_css_margin_left(self, paragraph):
        margin_left = paragraph.get_indent_left()
        if margin_left is not None:
            self.styles['margin-left'] = str(int(margin_left) // 20) + 'px'

    def __to_css_margin_right(self, paragraph):
        margin_right = paragraph.get_indent_right()
        if margin_right is not None:
            self.styles['margin-left'] = str(int(margin_right) // 20) + 'px'

    def __to_css_text_indent(self, paragraph):
        hanging = paragraph.get_hanging()
        first_line = paragraph.get_first_line()
        if hanging is not None:
            self.styles['text-indent'] = str(-int(hanging) // 20) + 'px'
        elif first_line is not None:
            self.styles['text-indent'] = str(int(first_line) // 20) + 'px'


class RunTranslatorToHTML(TranslatorToHTML):
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
        self.methods: Dict[str, Callable] = {
            'w:rPr/w:sz/w:val': self.__to_css_size,
            'w:rPr/w:b': self.__to_ext_tag_bold,
            'w:rPr/w:i': self.__to_ext_tag_italic,
            'w:rPr/w:vertAlign/w:val': self.__to_ext_tags_vertical_align,
            'w:rPr/w:highlight/w:val': self.__to_css_background_color,
            'w:rPr/w:color/w:val': self.__to_css_color,
            'w:rPr/w:strike': self.__to_css_line_throught,
            'w:rPr/w:u/w:val': self.__to_css_underline,
            'w:rPr/w:u/w:color': self.__to_css_underline_color,     # must called after self.__to_css_underline
            'w:rPr/w:bdr/w:val': self._to_css_border_all,
            'w:rPr/w:bdr/w:color': self._to_css_border_all_color,
            'w:rPr/w:bdr/w:sz': self._to_css_border_all_size,
        }

    def _get_html_tag(self) -> str:
        if self.attributes or self.styles:
            return 'span'
        return ''

    def __to_css_size(self, run):
        size = run.get_size()
        if size is not None:
            self.styles['font-size'] = size

    def __to_ext_tag_bold(self, run):
        if run.is_bold():
            self._add_to_ext_tags(self.external_tags['bold'])

    def __to_ext_tag_italic(self, run):
        if run.is_italic():
            self._add_to_ext_tags(self.external_tags['italic'])

    def __to_ext_tags_vertical_align(self, run):
        vert_align = run.get_vertical_align()
        if vert_align is not None:
            if vert_align == 'subscript':
                self._add_to_ext_tags('sub')
            elif vert_align == 'superscript':
                self._add_to_ext_tags('sup')

    def __to_css_background_color(self, run):
        background_color = run.get_background_color()
        background_fill = run.get_background_fill()
        if background_color is not None:
            if background_color != 'none':
                self.styles['background-color'] = background_color
        elif background_fill is not None:
            self.styles['background-color'] = TranslatorToHTML._translate_color(background_fill)

    def __to_css_color(self, run):
        color = run.get_color()
        if color is not None:
            self.styles['color'] = TranslatorToHTML._translate_color(color)

    def __to_css_line_throught(self, run):
        if run.is_strike():
            self._add_to_many_properties_style(('text-decoration', 'line-through'))

    def __to_css_underline(self, run):
        underline = run.get_underline()
        if underline is not None:
            self._add_to_many_properties_style(('text-decoration', 'underline'))
            if underline in self.underlines:
                self.styles['text-decoration-style'] = self.underlines[underline]

    def __to_css_underline_color(self, run):
        """
        must called after self.__to_css_underline
        """
        underline_color = run.get_underline_color()
        if underline_color is not None:
            self.styles['text-decoration-color'] = TranslatorToHTML._translate_color(underline_color)

    def _to_css_border(self, run, direction: str) -> Union[Tuple[str, str], None]:
        border = run.get_border('type')
        if border is not None:
            return rf'border-{direction}', self.border_types_corresponding.get(border, 'solid')
        return None

    def _to_css_border_all(self, run):
        self._add_to_many_properties_style(self._to_css_border(run, 'top'))
        self._add_to_many_properties_style(self._to_css_border(run, 'bottom'))
        self._add_to_many_properties_style(self._to_css_border(run, 'left'))
        self._add_to_many_properties_style(self._to_css_border(run, 'right'))

    @staticmethod
    def _to_css_border_color(run, direction: str) -> Union[Tuple[str, str], None]:
        border_color = run.get_border('color')
        if border_color is not None:
            if border_color == 'auto':
                return rf'border-{direction}', 'black'
            elif border_color is not None:
                return rf'border-{direction}', CellTranslatorToHTML._translate_color(border_color)
        return None

    def _to_css_border_all_color(self, element):
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_color(element, 'top'))
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_color(element, 'bottom'))
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_color(element, 'left'))
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_color(element, 'right'))

    @staticmethod
    def _to_css_border_size(run, direction: str) -> Union[Tuple[str, str], None]:
        border_size = run.get_border('size')
        if border_size is not None:
            return rf'border-{direction}', str(int(border_size) / 8) + 'pt'
        return None

    def _to_css_border_all_size(self, element):
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_size(element, 'top'))
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_size(element, 'bottom'))
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_size(element, 'left'))
        self._add_to_many_properties_style(RunTranslatorToHTML._to_css_border_size(element, 'right'))


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
    }

    def translate(self, text_element) -> str:
        import re

        text = text_element.content
        for char in self.characters_html:
            text = re.sub(char, self.characters_html[char], text)
        return text


class TableTranslatorToHTML(TranslatorToHTML):

    def __init__(self):
        super(TableTranslatorToHTML, self).__init__()
        self.methods: Dict[str, Callable] = {
            'w:tblPr/w:tblW/w:w': self.__to_attribute_width,
            'w:tblPr/w:jc/w:val': self.__to_attribute_align,
            'w:tblPr/w:tblBorders/w:top/w:val': self._to_css_border_top,
            'w:tblPr/w:tblBorders/w:bottom/w:val': self._to_css_border_bottom,
            'w:tblPr/w:tblBorders/w:left/w:val': self._to_css_border_left,
            'w:tblPr/w:tblBorders/w:right/w:val': self._to_css_border_right,
            'w:tblPr/w:tblBorders/w:top/w:color': self._to_css_border_top_color,
            'w:tblPr/w:tblBorders/w:bottom/w:color': self._to_css_border_bottom_color,
            'w:tblPr/w:tblBorders/w:left/w:color': self._to_css_border_left_color,
            'w:tblPr/w:tblBorders/w:right/w:color': self._to_css_border_right_color,
            'w:tblPr/w:tblBorders/w:top/w:sz': self._to_css_border_top_size,
            'w:tblPr/w:tblBorders/w:bottom/w:sz': self._to_css_border_bottom_size,
            'w:tblPr/w:tblBorders/w:left/w:sz': self._to_css_border_left_size,
            'w:tblPr/w:tblBorders/w:right/w:sz': self._to_css_border_right_size,
            'w:tblPr/w:tblInd/w:w': self.__to_css_margin,
        }

    def _do_methods(self, table):
        self.__to_css_border_collapse()
        super(TableTranslatorToHTML, self)._do_methods(table)

    def _get_html_tag(self) -> str:
        return 'table'

    def __to_css_border_collapse(self):
        self.styles['border-collapse'] = 'collapse'

    def __to_attribute_width(self, table):
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

    def __to_attribute_align(self, table):
        align = table.get_align()
        if align is not None:
            self.attributes['align'] = align

    def __to_css_margin(self, table):
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

    def __init__(self):
        super(RowTranslatorToHTML, self).__init__()
        self.methods: Dict[str, Callable] = {
            'w:trPr/w:trHeight/w:val': self.__to_attribute_or_css_height,
        }

    def _convert_fields(self, row):
        self.__to_ext_tag_head(row)

    def _get_html_tag(self) -> str:
        return 'tr'

    def __to_ext_tag_head(self, row):
        if row.is_header():
            self._add_to_ext_tags('thead', row.is_first_row_in_header, row.is_last_row_in_header)

    def __to_attribute_or_css_height(self, row):
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


class CellTranslatorToHTML(TranslatorToHTML):
    text_directions: Dict[str, str] = {
        'btLr': 'vertical-rl',
        'tbRl': 'vertical-rl',
    }

    def __init__(self):
        super(CellTranslatorToHTML, self).__init__()
        self.methods: Dict[str, Callable] = {
            'w:tcPr/w:shd/w:fill': self.__to_css_fill_color,
            'w:tcPr/w:shd/w:themeFill': self._pass,
            'w:tcPr/w:vMerge/w:val': self._pass,
            'w:tcPr/w:vMerge': self._pass,
            'w:tcPr/w:tcBorders/w:top/w:val': self._to_css_border_top,
            'w:tcPr/w:tcBorders/w:bottom/w:val': self._to_css_border_bottom,
            'w:tcPr/w:tcBorders/w:left/w:val': self._to_css_border_left,
            'w:tcPr/w:tcBorders/w:right/w:val': self._to_css_border_right,
            'w:tcPr/w:tcBorders/w:top/w:color': self._to_css_border_top_color,
            'w:tcPr/w:tcBorders/w:bottom/w:color': self._to_css_border_bottom_color,
            'w:tcPr/w:tcBorders/w:left/w:color': self._to_css_border_left_color,
            'w:tcPr/w:tcBorders/w:right/w:color': self._to_css_border_right_color,
            'w:tcPr/w:tcBorders/w:top/w:sz': self._to_css_border_top_size,
            'w:tcPr/w:tcBorders/w:bottom/w:sz': self._to_css_border_bottom_size,
            'w:tcPr/w:tcBorders/w:left/w:sz': self._to_css_border_left_size,
            'w:tcPr/w:tcBorders/w:right/w:sz': self._to_css_border_right_size,
            'w:tcPr/w:tcW/w:w': self.__to_attribute_width,
            'w:tcPr/w:gridSpan/w:val': self.__to_attribute_col_span,
            'w:tcPr/w:textDirection/w:val': self.__to_css_text_direction,
            'w:tcPr/w:vAlign/w:val': self.__to_attribute_vertical_align,
            'w:tcPr/w:tcMar/w:top/w:w': self.__to_css_padding_top,
            'w:tcPr/w:tcMar/w:bottom/w:w': self.__to_css_padding_bottom,
            'w:tcPr/w:tcMar/w:left/w:w': self.__to_css_padding_left,
            'w:tcPr/w:tcMar/w:right/w:w': self.__to_css_padding_right,
            # self.__to_css_padding_..._type must called after self.__to_css_padding_...
            'w:tcPr/w:tcMar/w:top/w:type': self.__to_css_padding_top_type,
            'w:tcPr/w:tcMar/w:bottom/w:type': self.__to_css_padding_bottom_type,
            'w:tcPr/w:tcMar/w:left/w:type': self.__to_css_padding_left_type,
            'w:tcPr/w:tcMar/w:right/w:type': self.__to_css_padding_right_type,
        }
        self.__is_header: bool = False
        self.__tag_for_header = 'td'

    def _convert_fields(self, element):
        self.__to_attribute_row_span(element)
        self.__is_header = element.is_header

    def _get_html_tag(self) -> str:
        return 'td' if not self.__is_header else self.__tag_for_header

    def is_th_for_header(self, value: bool = True):
        self.__tag_for_header = 'th' if value is True else 'td'

    def __to_css_fill_color(self, cell):
        color = cell.get_fill_color()
        if color is not None:
            if color != 'auto':
                self.styles['background-color'] = TranslatorToHTML._translate_color(color)

    def __to_css_padding_top(self, cell):
        padding = cell.get_margin('top')[0]
        if padding is not None:
            self.styles['padding-top'] = padding

    def __to_css_padding_bottom(self, cell):
        padding = cell.get_margin('bottom')[0]
        if padding is not None:
            self.styles['padding-bottom'] = padding

    def __to_css_padding_left(self, cell):
        padding = cell.get_margin('left')[0]
        if padding is not None:
            self.styles['padding-left'] = padding

    def __to_css_padding_right(self, cell):
        padding = cell.get_margin('right')[0]
        if padding is not None:
            self.styles['padding-right'] = padding

    def __to_css_padding_type(self, cell, direction: str):
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

    def __to_css_padding_top_type(self, cell):
        self.__to_css_padding_type(cell, 'top')

    def __to_css_padding_bottom_type(self, cell):
        self.__to_css_padding_type(cell, 'bottom')

    def __to_css_padding_left_type(self, cell):
        self.__to_css_padding_type(cell, 'left')

    def __to_css_padding_right_type(self, cell):
        self.__to_css_padding_type(cell, 'right')

    def __to_attribute_width(self, cell):
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

    def __to_attribute_col_span(self, cell):
        col_span = cell.get_col_span()
        if col_span is not None:
            if int(col_span) > 1:
                self.attributes['colspan'] = col_span

    def __to_attribute_row_span(self, cell):
        if cell.row_span > 1:
            self.attributes['rowspan'] = str(cell.row_span)

    def __to_css_text_direction(self, cell):
        text_direction = cell.get_text_direction()
        if text_direction is not None:
            if text_direction in self.text_directions:
                self.inner_html_wrapper_styles['writing-mode'] = self.text_directions[text_direction]

    def __to_attribute_vertical_align(self, cell):
        vertical_align = cell.get_vertical_align()
        if vertical_align is not None:
            self.attributes['valign'] = vertical_align
