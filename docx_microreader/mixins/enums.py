class EnumOfBorderedElementMixin:
    @classmethod
    def _element_key(cls) -> str:
        raise NotImplementedError('method _element_key is not implemented. Your class must implement this method or'
                                  'GetSetMixin must be inherited after the class that implements this method')

    @classmethod
    def get_border_property_enum_value(cls, direction, prop_name):
        """
        :param direction: instance of property_enums.Direction or corresponding value
        :param prop_name: instance of property_enums.BorderProperty or corresponding value
        :return: instance of border property of cls
        """
        from ..constants.property_enums import subelement_property_key, subelement_property_key_of_element, SubElement
        from ..utils.enums import convert_to_enum_element

        return convert_to_enum_element(
            subelement_property_key_of_element(
                cls._element_key(), subelement_property_key(
                    SubElement.BORDER, direction, prop_name
                )
            ), cls
        )


class CellMarginEnumMixin:
    @classmethod
    def _element_key(cls) -> str:
        raise NotImplementedError('method _element_key is not implemented. Your class must implement this method or'
                                  'GetSetMixin must be inherited after the class that implements this method')

    @classmethod
    def get_cell_margin_property_enum_value(cls, direction, prop_name):
        """
        :param direction: instance of property_enums.Direction or corresponding value
        :param prop_name: instance of property_enums.MarginProperty or corresponding value
        :return: instance of margin property of cls
        """
        from ..constants.property_enums import subelement_property_key, subelement_property_key_of_element, SubElement
        from ..utils.enums import convert_to_enum_element

        return convert_to_enum_element(
            subelement_property_key_of_element(
                cls._element_key(), subelement_property_key(
                    SubElement.MARGIN, direction, prop_name
                )
            ), cls
        )
