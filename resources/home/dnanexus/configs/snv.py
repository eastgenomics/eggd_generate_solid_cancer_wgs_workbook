import string

from openpyxl.styles import Border, Side
from openpyxl.styles.fills import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

from utils import misc

THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "cells_to_write": {
        (1, i): value
        for i, value in enumerate(
            [
                "Domain",
                "Gene",
                "GRCh38 coordinates",
                "Cyto",
                "Variant",
                "Predicted consequences",
                "VAF",
                "LOH",
                "Error flag",
                "Alt allele/total read depth",
                "Gene mode of action",
                "Variant class",
                "TSG_NMD",
                "TSG_LOH",
                "Splice fs?",
                "SpliceAI",
                "REVEL",
                "OG_3' Ter",
                "Recurrence somatic database",
                "HS_Total",
                "HS_Sample",
                "HS_Tumour",
                "COSMIC Driver",
                "COSMIC Alterations",
                "Paed Driver",
                "Paed Entities",
                "Sarc Driver",
                "Sarc Entities",
                "Neuro Driver",
                "Neuro Entities",
                "Ovary Driver",
                "Ovary Entities",
                "Haem Driver",
                "Haem Entities",
                "MTBP c.",
                "MTBP p.",
            ],
            1,
        )
    },
    "to_bold": [f"{misc.convert_index_to_letters(i)}1" for i in range(0, 42)],
    "col_width": [
        ("B", 12),
        ("C", 28),
        ("D", 14),
        ("E", 18),
        ("F", 14),
        ("J", 20),
        ("K", 20),
        ("L", 20),
        ("M", 20),
        ("N", 20),
        ("O", 14),
        ("P", 22),
        ("Q", 26),
        ("R", 18),
        ("S", 18),
        ("T", 16),
        ("U", 16),
        ("V", 18),
        ("W", 18),
    ],
    "borders": {
        "cell_rows": [
            ("A1:AJ1", THIN_BORDER),
        ],
    },
    "auto_filter": "E:AJ",
    "freeze_panes": "F1",
}


def add_dynamic_values(data: pd.DataFrame) -> dict:
    """Add dynamic values for the SNV sheet

    Parameters
    ----------
    data : pd.DataFrame
        Dataframe containing the data for somatic variants and appropriate
        additional data from inputs

    Returns
    -------
    dict
        Dict containing data that needs to be merged to the CONFIG variable
    """

    nb_somatic_variants = data.shape[0]

    config_with_dynamic_values = {
        "cells_to_write": {
            # remove the col and row index from the writing?
            (r_idx - 1, c_idx - 1): value
            for r_idx, row in enumerate(dataframe_to_rows(data), 1)
            for c_idx, value in enumerate(row, 1)
            if c_idx != 1 and r_idx != 1
        },
        "cells_to_colour": [
            # letters K to R
            (
                f"{string.ascii_uppercase[i]}{j}",
                PatternFill(patternType="solid", start_color="FFDBBB"),
            )
            for i in range(11, 19)
            for j in range(1, nb_somatic_variants + 2)
        ]
        + [
            (
                f"{letter}{j}",
                PatternFill(patternType="solid", start_color="c4d9ef"),
            )
            for letter in ["T", "U", "V"]
            for j in range(1, nb_somatic_variants + 2)
        ]
        + [
            # letters V to AG
            (
                f"{misc.convert_index_to_letters(i)}{j}",
                PatternFill(patternType="solid", start_color="00FFFF"),
            )
            for i in range(22, 36)
            for j in range(1, nb_somatic_variants + 2)
        ]
        + [
            (
                f"{col}{i}",
                PatternFill(patternType="solid", start_color="dabcff"),
            )
            for col in ["AI", "AJ"]
            for i in range(1, nb_somatic_variants + 2)
        ],
        "dropdowns": [
            {
                "cells": {
                    (f"L{i}" for i in range(2, nb_somatic_variants + 2)): (
                        '"Pathogenic, Likely pathogenic,'
                        "Uncertain, Likely passenger,"
                        'Likely artefact"'
                    ),
                },
                "title": "Variant class",
            },
        ],
        "data_bar": f"G2:G{nb_somatic_variants + 1}",
    }

    return config_with_dynamic_values
