def merge_dicts(dict_list: list) -> dict:
    """
    merging dicts from dict_list to one dict. If dicts from dict_list have same keys pass KeyError
    """
    result: dict = {}
    for d in dict_list:
        if len(set(d.keys()) & set(result.keys())) > 0:
            raise KeyError(rf'keys_consts have same keys: {set(d.keys()) & set(result.keys())}')
        result.update(d)
    return result
