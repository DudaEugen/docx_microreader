from .docx_parser import Parser
import xml.etree.ElementTree as ET
from typing import Union, List, Dict, Tuple, Optional
from .properties import Property
from .constants import property_enums as pr_const
from .utils.functions import execute_if_not_none


class XMLement(Parser):
    element_description: Union[pr_const.Element, pr_const.Style, pr_const.SubStyle]
    from .constants.translate_formats import TranslateFormat

    # {TranslateFormat: translator} tarnslator must have method:
    # def translate(self, xml_element, translated_inner_elements: str)
    translators = {}
    translate_format: TranslateFormat = TranslateFormat.HTML

    # first element of Tuple is correct variant of property value; second element is variants of this value
    # _properties_validate method set correct variant if find value equal of one of variant
    # if value of property not equal one of variants or correct variant set None
    _properties_unificators: Dict[str, List[Tuple[str, List[str]]]] = {}

    _default_style: Optional[pr_const.DefaultStyle] = None

    def __init__(self, element: ET.Element, parent):
        self.parent: Optional[XMLement] = parent
        super(XMLement, self).__init__(element)
        self._properties_unificate()
        self._base_style = self._get_style_from_document()
        self._set_default_style_of_class()

    @classmethod
    def create(cls):
        return cls(ET.Element(cls.element_description.tag), None)

    def get_inner_element(self, index: int):
        return self._inner_elements[index]

    def count_inner_elements(self) -> int:
        return len(self._inner_elements)

    def insert_inner_element(self, index: int, element):
        if not self.__class__.is_possible_inner_element(element.__class__):
            raise TypeError(f"{element} can't be inner element of {self.__class__}")
        self._inner_elements.insert(index, element)
        element.parent = self

    def append_inner_element(self, element):
        if not self.__class__.is_possible_inner_element(element.__class__):
            raise TypeError(f"{element} can't be inner element of {self.__class__}")
        self._inner_elements.append(element)
        element.parent = self

    def pop_inner_element(self, index: int = -1):
        self._inner_elements[index].parent = None
        return self._inner_elements.pop(index)

    def iterate_by_inner_elements(self):
        for el in self._inner_elements:
            yield el

    def translate(self, to_format: Union[TranslateFormat, str, None] = None, is_recursive_translate: bool = True,
                  context: Optional[dict] = None):
        """
        :param to_format: using translate_format of element if None
        :param is_recursive_translate: pass to_format to inner element if True
        :param context: Translators can use this variable for create context of translation
        """
        from docx_microreader.constants.translate_formats import TranslateFormat
        translator = self.translators[TranslateFormat(to_format)] if to_format is not None else \
                     self.translators[self.translate_format]

        if context is None:
            context = {}
        translator.preparation_to_translate_inner_elements(self, context)

        translated_inner_elements = []
        for el in self._inner_elements:
            if is_recursive_translate:
                translated_inner_elements.append(el.translate(to_format, is_recursive_translate, context))
            else:
                translated_inner_elements.append(el.translate(context=context))
        return translator.translate(self, translated_inner_elements, context)

    def _get_document(self):
        return execute_if_not_none(self.parent, lambda x: x._get_document())

    def _get_style_id(self) -> Optional[str]:
        return None

    def _properties_unificate(self):
        """
        first element of _properties_validators[key] is correct variant of property value;
        second element is variants of this value
        set correct variant if find value equal of one of variant
        if value of property not equal one of variants or correct variant set None
        """
        for key in self._properties_unificators:
            if not (key in self._properties):
                raise KeyError(f'not found key "{key}" from _properties_validators in _all_properties')
            is_finding_value: bool = False
            for correct_value, variants in self._properties_unificators[key]:
                if self._properties[key].value == correct_value:
                    is_finding_value = True
                    break
                else:
                    for variant in variants:
                        if self._properties[key].value == variant:
                            self._properties[key].value = correct_value
                            is_finding_value = True
                            break
            if not is_finding_value:
                self._properties[key].value = None                

    @staticmethod
    def _key_of_property(property_name) -> str:
        """
        :param property_name: str or instance of Enum from constants.property_enums
        :return: key of property
        """
        return property_name if isinstance(property_name, str) else property_name.key

    def get_parent(self):
        return self.parent

    def get_property(self, property_name, is_find_missed_or_true: bool = True):
        """
        :param property_name: key of property (str or instance of Enum from constants.property_enums)
        :param is_find_missed_or_true: recursive find value if result equal Property.Missed. Return True if not find
        :return: Property.value or None
        """
        result = None
        key: str = XMLement._key_of_property(property_name)

        if key in self._properties:
            result = self._properties[key].value
            if result is not None and (not isinstance(result, Property.Missed) or not is_find_missed_or_true):
                return result

        if self._base_style is not None:
            base_style_result = self._base_style.get_property(key, is_find_missed_or_true)
            if base_style_result is not None:
                result = base_style_result
                if not isinstance(base_style_result, Property.Missed) or not is_find_missed_or_true:
                    return base_style_result

        if isinstance(result, Property.Missed) and is_find_missed_or_true:
            return True
        return result

    def get_base_style(self):
        return self._base_style

    def set_style(self, style_id: str):
        self._base_style = execute_if_not_none(self._get_document(), lambda x: x.get_style(style_id))

    def set_property_value(self, property_name, value: Union[str, bool, None]):
        """
        :param property_name: key of property (str or instance of Enum from constants.property_enums)
        :param value: new value
        """
        self._properties[XMLement._key_of_property(property_name)].value = value

    def _get_style_from_document(self):
        style_id = self._get_style_id()
        if style_id is not None:
            return execute_if_not_none(self._get_document(), lambda x: x.get_style(style_id))
        return None

    def _get_default_style_from_document(self, style: pr_const.DefaultStyle):
        return execute_if_not_none(self._get_document(), lambda x: x.get_default_style(style))

    @classmethod
    def _set_default_style_of_class(cls):
        if cls._default_style is None:
            for default_style in pr_const.DefaultStyle:
                if cls.element_description == default_style.element:
                    cls._default_style = default_style
