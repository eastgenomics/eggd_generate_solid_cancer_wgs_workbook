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
        (4, 4): "Genotype",
        (4, 5): "gnomAD",
        (4, 6): "Role in Cancer",
        (4, 7): "ClinVar",
        (4, 8): "Tumour VAF",
        (4, 9): "Panelapp Adult v2.2",
        (4, 10): "Panelapp Chilhood v4.0",
    },
    "to_bold": [
        "A1",
        "A4",
        "B4",
        "C4",
        "D4",
        "E4",
        "F4",
        "G4",
        "H4",
        "I4",
        "J4",
    ],
    "col_width": [
        ("A", 20),
        ("B", 18),
        ("C", 22),
        ("D", 14),
        ("F", 22),
        ("G", 18),
        ("H", 12),
        ("I", 44),
        ("J", 44),
    ],
    "cells_to_colour": [
        (f"{column}4", PatternFill(patternType="solid", start_color="ADD8E6"))
        for column in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
    ],
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
        "to_align": [f"I{i}" for i in range(4, nb_germline_variants + 5)]
        + [f"J{i}" for i in range(4, nb_germline_variants + 5)],
        "row_height": [(i, 30) for i in range(5, nb_germline_variants + 6)],
        "borders": {
            "single_cells": [(f"A{nb_germline_variants + 6}", LOWER_BORDER)],
            "cell_rows": [
                (f"A{i}:J{i}", THIN_BORDER)
                for i in range(4, nb_germline_variants + 5)
            ],
        },
    }

    return config_with_dynamic_values
