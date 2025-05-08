from openpyxl.styles import Border, Side
from openpyxl.styles.fills import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

from utils import misc

THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "cells_to_colour": [
        (
            f"{col}1",
            PatternFill(patternType="solid", start_color="fcba03"),
        )
        for col in ["B", "C", "D"]
    ]
    + [
        (
            f"{col}1",
            PatternFill(patternType="solid", start_color="db5a14"),
        )
        for col in ["E", "F", "G", "H"]
    ]
    + [
        (
            f"{col}1",
            PatternFill(patternType="solid", start_color="2c8500"),
        )
        for col in ["I", "J", "K"]
    ]
    + [
        (
            f"{col}1",
            PatternFill(patternType="solid", start_color="008581"),
        )
        for col in ["L", "M", "N", "O"]
    ]
    + [
        (
            f"{col}1",
            PatternFill(patternType="solid", start_color="024dc7"),
        )
        for col in ["P", "Q", "R", "S"]
    ]
    + [
        (
            f"{col}1",
            PatternFill(patternType="solid", start_color="ad2323"),
        )
        for col in ["T", "U", "V", "W"]
    ]
    + [
        (
            f"{col}1",
            PatternFill(patternType="solid", start_color="b686da"),
        )
        for col in ["X", "Y", "Z", "AA"]
    ],
    "to_bold": [f"{misc.convert_index_to_letters(i)}1" for i in range(0, 27)],
    "auto_filter": "A:AA",
    "borders": {
        "cell_rows": [
            ("A1:AA1", THIN_BORDER),
        ],
    },
}


SHEETS2COLUMNS = {
    "somatic_db": {
        "Gene": "Gene",
        "Role in Cancer": "Comments",
        "Driver_SV": "Alteration",
        "Entities": "Entities",
    },
    "haem": {
        "Gene": "Gene",
        "Driver": "Haem_Alteration",
        "Entities": "Haem_Entities",
        "Comments": "Haem_Comments",
        "Reference": "Haem_Reference",
    },
    "paed": {
        "Gene": "Gene",
        "Driver": "Paed_Alteration",
        "Entities": "Paed_Entities",
        "Comments": "Paed_Comments",
    },
    "ovarian": {
        "Gene": "Gene",
        "Driver": "Ovarian_Alteration",
        "Entities": "Ovarian_Entities",
        "Comments": "Ovarian_Comments",
        "Reference": "Ovarian_Reference",
    },
    "sarc": {
        "Gene": "Gene",
        "Driver": "Sarcoma_Alteration",
        "Entities": "Sarcoma_Entites",
        "Comments": "Sarcoma_Comments",
        "Reference": "Sarcoma_Reference",
    },
    "neuro": {
        "Gene": "Gene",
        "Driver": "Neuro_Alteration",
        "Entities": "Neuro_Entities",
        "Comments": "Neuro_Comments",
        "Reference": "Neuro_Reference",
    },
}

RESCUE_COLUMNS = {"somatic_db": ["cosmic"]}


def add_dynamic_values(df: pd.DataFrame) -> dict:
    """Add dynamic values for the refgene sheet

    Parameters
    ----------
    data : pd.DataFrame
        Dataframe containing the data for refgene

    Returns
    -------
    dict
        Dict containing data that needs to be merged to the CONFIG variable
    """

    config_with_dynamic_values = {
        "cells_to_write": {
            (1, i): column for i, column in enumerate(df.columns, 1)
        }
        | {
            # remove the col and row index from the writing?
            (r_idx - 1, c_idx - 1): value
            for r_idx, row in enumerate(dataframe_to_rows(df), 1)
            for c_idx, value in enumerate(row, 1)
            if c_idx != 1 and r_idx != 1
        },
    }

    return config_with_dynamic_values
