from openpyxl.styles import Border, Side
from openpyxl.styles.fills import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "cells_to_write": {
        (1, 1): "=SOC!A2",
        (2, 1): "=SOC!A3",
        (1, 3): "=SOC!A5",
        (2, 3): "=SOC!A6",
        (1, 5): "=SOC!A9",
        (4, 1): "Gene",
        (4, 2): "GRCh38 Coordinates",
        (4, 3): "Variant",
        (4, 4): "Consequence",
        (4, 5): "Genotype",
        (4, 6): "gnomAD",
        (4, 7): "Role in Cancer",
        (4, 8): "ClinVar",
        (4, 9): "Tumour VAF",
        (4, 10): "Panelapp Adult v2.2",
        (4, 11): "Panelapp Childhood v4.0",
    },
    "to_bold": ["A1"] + [f"{col}4" for col in list("ABCDEFGHIJK")],
    "col_width": [
        ("A", 12),
        ("B", 20),
        ("C", 16),
        ("D", 18),
        ("E", 12),
        ("F", 18),
        ("G", 24),
        ("H", 25),
        ("I", 16),
        ("J", 40),
        ("K", 40),
    ],
    "cells_to_colour": [
        (f"{column}4", PatternFill(patternType="solid", start_color="F2F2F2"))
        for column in list("ABCDEFGHIJK")
    ],
    "row_height": [(4, 40)],
}


def add_dynamic_values(data: pd.DataFrame) -> dict:
    """Add the parsed data to the CONFIG variable

    Parameters
    ----------
    data : pd.DataFrame
        Dataframe containing data parsed from the inputs

    Returns
    -------
    dict
        Dict with the parsed data and processed to have the correct position
    """

    if data.empty:
        return None

    nb_germline_variants = data.shape[0]

    config_with_dynamic_values = {
        # merge 2 dicts with parsed data and hard coded values
        "cells_to_write": {
            (r_idx + 2, c_idx - 1): value
            for r_idx, row in enumerate(dataframe_to_rows(data), 1)
            for c_idx, value in enumerate(row, 1)
            if c_idx != 1 and r_idx != 1
        }
        | {
            (nb_germline_variants + 6, 1): "Pertinent variants/feedback",
            (nb_germline_variants + 7, 1): "None",
        },
        "to_bold": [f"A{nb_germline_variants + 6}"],
        "alignment_info": [
            (
                f"{col}{row}",
                {
                    "vertical": "center",
                    "horizontal": "center",
                    "wrapText": True,
                },
            )
            for col in list("ABCDEFGHIJK")
            for row in range(4, nb_germline_variants + 6)
        ],
        "row_height": [(i, 40) for i in range(5, nb_germline_variants + 5)],
        "borders": {
            "single_cells": [(f"A{nb_germline_variants + 6}", LOWER_BORDER)],
            "cell_rows": [
                (f"A{i}:K{i}", THIN_BORDER)
                for i in range(4, nb_germline_variants + 5)
            ],
        },
    }

    return config_with_dynamic_values
