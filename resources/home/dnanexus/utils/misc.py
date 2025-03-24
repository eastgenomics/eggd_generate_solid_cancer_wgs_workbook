import importlib
from pathlib import Path
from types import ModuleType
from typing import Optional

# conditional path depending on whether the script is run on DNAnexus or not
if Path("/home/dnanexus").exists():
    CONFIG_PATH = Path("/home/dnanexus/configs")
else:
    CONFIG_PATH = Path("resources/home/dnanexus/configs")


def select_config(name_config: str) -> Optional[ModuleType]:
    """Given a config name, import the appropriate module for writing the sheet

    Parameters
    ----------
    name_config : str
        Name of the config module to import

    Returns
    -------
    Optional[ModuleType]
        Module to use for config information for writing the sheet or None if
        the names do not matches
    """

    for config in CONFIG_PATH.glob("*.py"):
        module_name = config.name.replace(".py", "")

        if name_config.lower() == module_name:
            return importlib.import_module(
                f".{module_name}", f"{config.parts[-2]}"
            )

    return None


def merge_dicts(left_dict: dict, right_dict: dict, sheet: str) -> dict:
    """Recursive function to merge 2 dicts:
    - Get unique keys from both dicts
    - Get common keys:
        - if values for those keys are lists, concatenate them
        - if values for those keys are dicts, get the values and run them
        through the process from the top

    Parameters
    ----------
    left_dict : dict
        First dict to merge
    right_dict : dict
        Second dict to merge
    sheet : str
        Name of the sheet for capturing the right data

    Returns
    -------
    dict
        Dict containing merged data from both dicts
    """

    # get the unique keys from both dicts
    if right_dict.get(sheet):
        right_dict = right_dict.get(sheet)

    unique_left_dict_keys = [
        left_key
        for left_key in left_dict.keys()
        if left_key not in right_dict.keys()
    ]

    unique_right_dict_keys = [
        right_key
        for right_key in right_dict.keys()
        if right_key not in left_dict.keys()
    ]

    new_dict = {}

    # add all unique keys from both dicts
    for key in unique_left_dict_keys:
        new_dict[key] = left_dict[key]

    for key in unique_right_dict_keys:
        new_dict[key] = right_dict[key]

    # get the common keys as it will require some processing
    common_keys = set(left_dict.keys()).intersection(right_dict)

    for key in common_keys:
        left_value = left_dict[key]
        right_value = right_dict[key]

        # if the types of the values for the same key are not the same, there's
        # a problem
        assert type(left_value) is type(
            right_value
        ), f"Types are not identical {left_value} | {right_value}"

        # if the type of the values is a list, just concatenate them
        if type(left_value) is list:
            new_dict[key] = left_value + right_value

        # if the type of the values is a dict, run the function recursively
        # until we reach keys that can be simply added or lists that we can
        # concatenate
        elif type(left_value) is dict:
            new_dict[key] = merge_dicts(left_value, right_value, sheet)

    return new_dict
