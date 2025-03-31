import importlib
from pathlib import Path
import string
from types import ModuleType
from typing import Optional

import pandas as pd

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


def lookup_value_in_other_df(
    target_df: pd.DataFrame,
    col_to_map: str,
    reference_df: pd.DataFrame,
    col_to_index: str,
    col_to_look_up: str,
) -> pd.Series:
    """Map a column from a reference dataframe column to another dataframe's
    column

    Parameters
    ----------
    target_df : pd.DataFrame
        Dataframe to map data to
    col_to_map : str
        Name of the column to map data to
    reference_df : pd.DataFrame
        Dataframe containing data to map from
    col_to_index : str
        Name of the column to index with
    col_to_look_up : str
        Name of the column to map from

    Returns
    -------
    pd.Series
        Series containing the data from the reference dataframe
    """

    return target_df[col_to_map].map(
        reference_df.set_index(col_to_index)[col_to_look_up]
    )


def split_confidence_support(value: str) -> list:
    """Split a value for paired and single read information (used in a Pandas
    context)

    Parameters
    ----------
    value : str
        String value to process

    Returns
    -------
    list
        2 element list containing information for the paired reads and single
        reads (in that order)
    """
    returned_value = []

    if "PR-" in value and "SR-" in value:
        value = value.split(";")

        for v in value:
            if "PR-" in v:
                cleaned_value = v.replace("PR-", "")
            elif "SR-" in v:
                cleaned_value = v.replace("SR-", "")
            else:
                cleaned_value = ""

            returned_value.append(cleaned_value)

    else:
        if "PR-" in value:
            returned_value.append(value.replace("PR-", ""))
            returned_value.append("")
        elif "SR-" in value:
            returned_value.append("")
            returned_value.append(value.replace("SR-", ""))

    return returned_value


def get_column_letter(df: pd.DataFrame, column_name: str = None) -> str:
    """Given a column name get the corresponding letter in a hypothetical
    excel worksheet. If not given, get the last column letter

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe on which columns to loop on
    column_name : str, optional
        Column name to look for, by default None

    Returns
    -------
    str
        Corresponding letter column
    """

    for i, column in enumerate(df.columns):
        if column_name and column_name == column:
            return string.ascii_uppercase[i]

        # reached the last column
        if len(df.columns) - 1 == i:
            return string.ascii_uppercase[i]


def get_lookup_groups(df: pd.DataFrame) -> list:
    """Get the position of the lookup groups i.e. position of the columns
    corresponding to the columns used for lookups

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe in which to look for column names corresponding to lookup
        columns

    Returns
    -------
    list
        List of tuple for the start and end of each lookup group
    """

    lookup_groups_position = []

    for i, column in enumerate(df.columns):
        if column == "COSMIC":
            lookup_groups_position.append([j for j in range(i, i + 6)])

    return lookup_groups_position
