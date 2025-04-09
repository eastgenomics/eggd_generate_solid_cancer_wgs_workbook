import string

from openpyxl.styles import Border, Side
from openpyxl.styles.fills import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


CONFIG = {
    "cells_to_write": {
        (1, i): value
        for i, value in enumerate(
            [
                "Event domain",
                "Impacted transcript region",
                "Gene",
                "GRCh38 coordinates",
                "Chromosomal bands",
                "Type",
                "Copy Number",
                "Size",
                "Gene mode of action",
                "Variant class",
                "TSG_Hom",
                "SNV_LOH",
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
            ],
            1,
        )
    },
    "to_bold": [f"{string.ascii_uppercase[i]}1" for i in range(0, 24)],
    "col_width": [
        ("B", 12),
        ("C", 16),
        ("D", 22),
        ("E", 20),
        ("G", 16),
        ("H", 14),
        ("I", 22),
        ("J", 20),
        ("K", 20),
        ("L", 20),
        ("M", 22),
        ("N", 20),
        ("O", 16),
        ("P", 16),
        ("Q", 16),
        ("R", 16),
        ("S", 16),
        ("T", 16),
        ("U", 16),
        ("V", 16),
        ("W", 16),
        ("X", 16),
    ],
    "borders": {
        "cell_rows": [
            ("A1:X1", THIN_BORDER),
        ],
    },
    "auto_filter": "A:X",
    "freeze_panes": "F1",
}


def add_dynamic_values(data: pd.DataFrame) -> dict:
    """Add dynamic values for the Loss config

    Parameters
    ----------
    data : pd.DataFrame
        Dataframe for the Loss structural variants

    Returns
    -------
    dict
        Dict populated with the dynamic values processed for the Loss
        structural variants
    """

    nb_sv_variants = data.shape[0]

    config_with_dynamic_values = {
        "cells_to_write": {
            # remove the col and row index from the writing?
            (r_idx - 1, c_idx - 1): value
            for r_idx, row in enumerate(dataframe_to_rows(data), 1)
            for c_idx, value in enumerate(row, 1)
            if c_idx != 1 and r_idx != 1
        },
        "cells_to_colour": [
            (
                f"{col}{i}",
                PatternFill(patternType="solid", start_color="FFDBBB"),
            )
            for col in ["J", "K", "L"]
            for i in range(1, nb_sv_variants + 2)
        ]
        + [
            (
                # letters M to R
                f"{string.ascii_uppercase[i]}{j}",
                PatternFill(patternType="solid", start_color="c4d9ef"),
            )
            for i in range(12, 24)
            for j in range(1, nb_sv_variants + 2)
        ],
        "to_align": [f"G{i}" for i in range(2, nb_sv_variants + 2)],
        "dropdowns": [
            {
                "cells": {
                    (f"J{i}" for i in range(2, nb_sv_variants + 2)): (
                        '"Pathogenic, Likely pathogenic,'
                        "Uncertain, Likely passenger,"
                        'Likely artefact"'
                    ),
                },
                "title": "Variant class",
            },
        ],
    }

    return config_with_dynamic_values
