from typing import Union, List


class PropertyDescription:

    def __init__(self, tag_wrap: Union[List[str], str, None] = None, tag: Union[List[str], str, None] = None,
                 tag_property: Union[List[str], str, None] = None, is_can_be_miss: bool = False):
        """
        :param tag_wrap: wrapped tags of property tag (example: tag1/tag2/tag2/...)
        :param tag: tag of property
        :param tag_property: tag attribute that corresponding to property
        :param is_can_be_miss: value of Property can be equal Property.Missed
        """
        from .constants.property_enums import MissedPropertyAttribute
        
        self.tag_wrap: Union[List[str], str, None] = tag_wrap
        self.tag: Union[List[str], str] = tag
        self.tag_property: Union[List[str], str, None] = tag_property if not is_can_be_miss else MissedPropertyAttribute
        self.is_can_be_miss: bool = is_can_be_miss

    def get_wrapped_tags(self) -> Union[List[str], str]:
        tag_wrap: Union[List[str], str] = self.__wrap_for_tag()
        if isinstance(self.tag, list) and isinstance(tag_wrap, list):
            result: List[str] = []
            for wrap in tag_wrap:
                result.extend([(wrap + tag) for tag in self.tag])
            return result
        if isinstance(self.tag, list):
            return [(tag_wrap + tag) for tag in self.tag]
        if isinstance(tag_wrap, list):
            return [(wrap + self.tag) for wrap in tag_wrap]
        return tag_wrap + self.tag

    def __wrap_for_tag(self) -> Union[List[str], str]:
        if self.tag_wrap is not None:
            if isinstance(self.tag_wrap, list):
                return [(tag_wrap + '/') for tag_wrap in self.tag_wrap]
            return self.tag_wrap + '/'
        return ''


class Property:

    def __init__(self, value):
        self.value: Union[str, bool, Property.Missed, None] = value

    class Missed:
        """
        if Property.value equal Missed, tag don't have attribute for this Property but it is correct value of Property.
        This is mean using default value for corresponding attribute
        """
        pass
