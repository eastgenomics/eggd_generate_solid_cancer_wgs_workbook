import string

from openpyxl.styles import Border, Side
from openpyxl.styles.fills import PatternFill
import pandas as pd

from utils import misc

# prepare formatting
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
        (23, 1): "Germline SNV",
        (24, 1): "Gene",
        (24, 2): "GRCh38 Coordinates",
        (24, 3): "Variant",
        (24, 4): "Consequence",
        (24, 5): "Zygosity",
        (24, 6): "Variant Class",
        (24, 7): "Actionability",
        (24, 8): "Comments",
        (31, 1): "Somatic SNV",
        (32, 1): "Gene",
        (32, 2): "GRCh38 Coordinates",
        (32, 3): "Variant",
        (32, 4): "Consequence",
        (32, 5): "Zygosity",
        (32, 6): "Variant Class",
        (32, 7): "Actionability",
        (44, 1): "Somatic CNV_SV",
        (45, 1): "Gene/Locus",
        (45, 2): "GRCh38 Coordinates",
        (45, 3): "Cytological Bands",
        (45, 4): "Variant Type",
        (45, 5): "Consequence",
        (45, 6): "Variant Class",
        (45, 7): "Actionability",
        (45, 8): "Comments",
        (58, 1): "Germline_SNV",
        (67, 1): "Somatic_SNV",
        (78, 1): "CNV",
        (89, 1): "SV",
        (59, 1): "Gene",
        (59, 2): "GRCh38 Coordinates",
        (59, 3): "Variant",
        (59, 4): "Consequence",
        (59, 5): "Genotype",
        (59, 6): "Variant Class",
        (59, 7): "Actionability",
        (59, 8): "Role in Cancer",
        (59, 9): "ClinVar",
        (59, 10): "gnomAD",
        (59, 11): "Tumour VAF",
    }
    | {
        (row, col_index + 1): f"={col}{row+35}"
        for row in range(25, 30)
        for col_index, col in enumerate(["A", "B", "C", "D", "E", "F", "G"])
    }
    | {
        (row, col_index): f"={letter}{row+34}"
        for row in range(46, 52)
        for col_index, letter in [
            (1, "C"),
            (2, "D"),
            (3, "E"),
            (4, "F"),
            (6, "J"),
            (7, "K"),
            (8, "L"),
        ]
    },
    "to_bold": [
        "A1",
        "A23",
        "A24",
        "A31",
        "A32",
        "A44",
        "A45",
        "B24",
        "B32",
        "B45",
        "C24",
        "C32",
        "C45",
        "D24",
        "D32",
        "D45",
        "E24",
        "E32",
        "E45",
        "F24",
        "F32",
        "F45",
        "G24",
        "G32",
        "G45",
        "H24",
        "H32",
        "H45",
        "A58",
        "A67",
        "A78",
        "A89",
    ],
    "col_width": [
        ("A", 26),
        ("B", 20),
        ("C", 22),
        ("D", 24),
        ("F", 24),
        ("G", 24),
        ("H", 24),
    ],
    "cells_to_colour": [
        (
            f"{column}{row}",
            PatternFill(patternType="solid", start_color="ADD8E6"),
        )
        for row in [24, 32, 45]
        for column in ["A", "B", "C", "D", "E", "F", "G", "H"]
    ]
    + [
        (
            f"{string.ascii_uppercase[col_index-1]}{row}",
            PatternFill(patternType="solid", start_color="E7B8C8"),
        )
        for row in range(46, 52)
        for col_index in [1, 2, 3, 4, 5, 6, 7, 8]
    ]
    + [
        (
            "E46",
            PatternFill(patternType="solid", start_color="E7B8C8"),
        ),
        (
            "E47",
            PatternFill(patternType="solid", start_color="E7B8C8"),
        ),
        (
            "E48",
            PatternFill(patternType="solid", start_color="E7B8C8"),
        ),
        (
            "E49",
            PatternFill(patternType="solid", start_color="E7B8C8"),
        ),
        (
            "E50",
            PatternFill(patternType="solid", start_color="E7B8C8"),
        ),
        (
            "E51",
            PatternFill(patternType="solid", start_color="E7B8C8"),
        ),
    ],
    "borders": {
        "cell_rows": [(f"A{row}:H{row}", THIN_BORDER) for row in range(24, 30)]
        + [(f"A{row}:H{row}", THIN_BORDER) for row in range(32, 43)]
        + [(f"A{row}:H{row}", THIN_BORDER) for row in range(45, 57)],
    },
    "images": [
        {"cell": "A4", "img_index": 2, "size": (350, 700)},
        {"cell": "G4", "img_index": 1, "size": (350, 350)},
    ],
}


def add_dynamic_values(
    SNV_df: pd.DataFrame,
    fusion_count: int,
    SNV_df_columns: list,
    gain_df_columns: list,
    SV_df_columns: list,
):
    variant_class_column_letter = misc.get_column_letter(
        SNV_df, "Variant class"
    )
    actionability_column_letter = misc.get_column_letter(
        SNV_df, "Actionability"
    )
    comments_column_letter = misc.get_column_letter(SNV_df, "Comments")

    sv_pair = [
        (1, "C"),
        (2, "D"),
        (3, "E"),
        (4, "F"),
        (5, "G"),
        (6, variant_class_column_letter),
        (7, actionability_column_letter),
        (8, comments_column_letter),
    ]

    if fusion_count == 0:
        sv_pair.pop(4)

    config_with_dynamic_values = {
        "cells_to_write": {
            (68, col_index): col_name
            for col_index, col_name in enumerate(SNV_df_columns, 1)
        }
        | {
            (79, col_index): col_name
            for col_index, col_name in enumerate(gain_df_columns, 1)
        }
        | {
            (90, col_index): col_name
            for col_index, col_name in enumerate(SV_df_columns, 1)
        }
        | {
            (row, col_index): f"={col_letter}{row+36}"
            for row in range(33, 43)
            for col_index, col_letter in [
                (1, "B"),
                (2, "C"),
                (3, "D"),
                (4, "E"),
                (5, "F"),
                (6, variant_class_column_letter),
                (7, actionability_column_letter),
                (8, comments_column_letter),
            ]
        }
        | {
            (row, col_index): f"={col_letter}{row+39}"
            for row in range(52, 57)
            for col_index, col_letter in sv_pair
        },
        "cells_to_colour": [
            (
                f"{string.ascii_uppercase[col_index-1]}{row}",
                PatternFill(patternType="solid", start_color="C8F7E3"),
            )
            for row in range(52, 57)
            for col_index, letter in sv_pair
        ],
    }

    return config_with_dynamic_values
