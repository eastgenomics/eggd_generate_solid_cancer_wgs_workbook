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


def merge_dicts(original_dict: dict, new_dict: dict) -> dict:
    """Recursive function to merge 2 dicts:
    - Get unique keys from both dicts
    - Get common keys:
        - if values for those keys are lists, concatenate them
        - if values for those keys are dicts, get the values and run them
        through the process from the top

    Parameters
    ----------
    original_dict : dict
        First dict to merge
    new_dict : dict
        Second dict to merge

    Returns
    -------
    dict
        Dict containing merged data from both dicts
    """

    unique_original_dict_keys = [
        left_key
        for left_key in original_dict.keys()
        if left_key not in new_dict.keys()
    ]

    unique_new_dict_keys = [
        right_key
        for right_key in new_dict.keys()
        if right_key not in original_dict.keys()
    ]

    new_dict = {}

    # add all unique keys from both dicts
    for key in unique_original_dict_keys:
        new_dict[key] = original_dict[key]

    for key in unique_new_dict_keys:
        new_dict[key] = new_dict[key]

    # get the common keys as it will require some processing
    common_keys = set(original_dict.keys()).intersection(new_dict)

    for key in common_keys:
        original_value = original_dict[key]
        new_value = new_dict[key]

        # if the types of the values for the same key are not the same, there's
        # a problem
        assert type(original_value) is type(
            new_value
        ), f"Types are not identical {original_value} | {new_value}"

        # if the type of the values is a list, just concatenate them
        if type(original_value) is list:
            new_dict[key] = original_value + new_value

        # if the type of the values is a dict, run the function recursively
        # until we reach keys that can be simply added or lists that we can
        # concatenate
        elif type(original_value) is dict:
            new_dict[key] = merge_dicts(original_value, new_value)

        else:
            new_dict[key] = new_value

    return new_dict


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


def get_column_letter_using_column_name(
    df: pd.DataFrame, column_name: str = None
) -> str:
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
        # if the number of columns is over 26, i need to add the additional
        # letter i.e. 27th column would be AA in Excel
        if i >= 26:
            nb_alphabet_passes = int(i / 26)
            i -= nb_alphabet_passes * 26
            additional_letter = string.ascii_uppercase[nb_alphabet_passes - 1]
        else:
            nb_alphabet_passes = 0
            additional_letter = ""

        if column_name and column_name == column:
            return f"{additional_letter}{string.ascii_uppercase[i]}"

        # reached the last column
        if len(df.columns) - 1 == i + nb_alphabet_passes * 26:
            return f"{additional_letter}{string.ascii_uppercase[i]}"


def convert_letter_column_to_index(letters: str) -> int:
    """Convert a letter string into its position in the alphabet

    Parameters
    ----------
    letters : str
        Letter string

    Returns
    -------
    int
        Position in the alphabet
    """

    if len(letters) == 1:
        return string.ascii_uppercase.index(letters)
    else:
        return string.ascii_uppercase.index(
            letters[0]
        ) * 26 + string.ascii_uppercase.index(letters[1])


def convert_index_to_letters(index: int) -> str:
    """Convert a alphabet position into a letter

    Parameters
    ----------
    index : int
        Integer representing the position in the alphabet

    Returns
    -------
    str
        Equivalent letter of the position in the alphabet
    """

    if index >= 26:
        nb_alphabet_passes = int(index / 26)
        index -= nb_alphabet_passes * 26
        additional_letter = string.ascii_uppercase[nb_alphabet_passes - 1]
    else:
        additional_letter = ""

    return f"{additional_letter}{string.ascii_uppercase[index]}"


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


def convert_3_letter_protein_to_1(string_element: str) -> str:
    """Convert the 3 letter protein to a 1 letter protein

    Parameters
    ----------
    string_element : str
        String element to convert

    Returns
    -------
    str
        Converted string
    """

    if type(string_element) is not str:
        return string_element

    mapping = {
        "Ala": "A",
        "Arg": "R",
        "Asn": "N",
        "Asp": "D",
        "Cys": "C",
        "Gln": "Q",
        "Glu": "E",
        "Gly": "G",
        "His": "H",
        "Ile": "I",
        "Leu": "L",
        "Lys": "K",
        "Met": "M",
        "Phe": "F",
        "Pro": "P",
        "Ser": "S",
        "Thr": "T",
        "Trp": "W",
        "Tyr": "Y",
        "Val": "V",
    }

    for three_letter_protein, single_letter_protein in mapping.items():
        string_element = string_element.replace(
            three_letter_protein, single_letter_protein
        )

    return string_element
