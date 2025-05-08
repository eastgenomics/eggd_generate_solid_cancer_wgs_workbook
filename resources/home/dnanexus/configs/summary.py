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
        # hardcoded table headers and other elements
        (1, 1): "=SOC!A2",
        (2, 1): "=SOC!A3",
        (1, 3): "=SOC!A5",
        (2, 3): "=SOC!A6",
        (1, 5): "=SOC!A9",
        (23, 1): "Somatic SNV",
        (24, 1): "Gene",
        (24, 2): "GRCh38 Coordinates",
        (24, 3): "Variant",
        (24, 4): "Consequence",
        (24, 5): "Zygosity",
        (24, 6): "Variant Class",
        (24, 7): "Actionability",
        (24, 8): "Comments",
        (35, 1): "Somatic CNV_SV",
        (36, 1): "Gene/Locus",
        (36, 2): "GRCh38 Coordinates",
        (36, 3): "Cytological Bands",
        (36, 4): "Variant Type",
        (36, 5): "Consequence",
        (36, 6): "Variant Class",
        (36, 7): "Actionability",
        (36, 8): "Comments",
        (48, 1): "Germline SNV",
        (49, 1): "Gene",
        (49, 2): "GRCh38 Coordinates",
        (49, 3): "Variant",
        (49, 4): "Consequence",
        (49, 5): "Zygosity",
        (49, 6): "Variant Class",
        (49, 7): "Actionability",
        (55, 1): "Germline CNV",
        (56, 1): "Gene",
        (56, 2): "GRCh38 Coordinates",
        (56, 3): "Variant",
        (56, 4): "Consequence",
        (56, 5): "Zygosity",
        (56, 6): "Variant Class",
        (56, 7): "Actionability",
        (56, 8): "Comments",
        (62, 1): "Somatic_SNV",
        (74, 1): "Somatic_CNV",
        (82, 1): "Somatic_SV",
        (90, 1): "Germline_SNV",
        (97, 1): "Germline_CNV",
    }
    ####
    # somatic snv gene lookup
    | {(row, 1): f"=B{row+39}" for row in range(25, 34)}
    # somatic snv coordinates
    | {
        (row, 2): f'=SUBSTITUTE(C{row+39},";",CHAR(10))'
        for row in range(25, 34)
    }
    # somatic snv variant
    | {
        (row, 3): f'=SUBSTITUTE(F{row+39},";",CHAR(10))'
        for row in range(25, 34)
    }
    # somatic snv consequences
    | {(row, 4): f"=G{row+39}" for row in range(25, 34)}
    | {
        (row, 5): f"=CONCATENATE(J{row+39},CHAR(10),K{row+39})"
        for row in range(25, 34)
    }
    ####
    # somatic cnv gene lookup
    | {(row, 1): f"=B{row+39}" for row in range(37, 42)}
    # somatic cnv coordinates
    | {
        (row, 2): f'=SUBSTITUTE(E{row+39},";",CHAR(10))'
        for row in range(37, 42)
    }
    # somatic cnv cytological bands
    | {
        (row, 3): f"=CONCATENATE(I{row+39},CHAR(10),J{row+39})"
        for row in range(37, 42)
    }
    # somatic cnv variant type
    | {
        (row, 4): f'=CONCATENATE(F{row+39}," (",G{row+39},")")'
        for row in range(37, 42)
    }
    ####
    # somatic fusion gene lookup
    | {(row, 1): f"=C{row+42}" for row in range(42, 47)}
    # somatic fusion coordinates
    | {
        (row, 2): f'=SUBSTITUTE(E{row+42},";",CHAR(10))'
        for row in range(42, 47)
    }
    # somatic fusion cytological bands
    | {
        (row, 3): f"=CONCATENATE(I{row+42},CHAR(10),J{row+42})"
        for row in range(42, 47)
    }
    # somatic fusion variant type
    | {
        (row, 4): f'=CONCATENATE(F{row+42}," (",G{row+42},")")'
        for row in range(42, 47)
    }
    ####
    # germline snv gene lookup
    | {(row, 1): f"=A{row+41}" for row in range(50, 54)}
    ####
    # germline cnv gene lookup
    | {(row, 1): f"=A{row+41}" for row in range(57, 61)},
    "to_bold": [
        # table names to be bolded
        "A1",
        "A23",
        "A35",
        "A48",
        "A55",
        "A62",
        "A74",
        "A82",
        "A90",
        "A97",
    ]
    # table headers to be bolded
    + [f"{col}24" for col in list("ABCDEFGH")]
    + [f"{col}36" for col in list("ABCDEFGH")]
    + [f"{col}49" for col in list("ABCDEFGH")]
    + [f"{col}56" for col in list("ABCDEFGH")],
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
        for row in [24, 36, 49, 56]
        for column in list("ABCDEFGH")
    ]
    + [
        (
            f"{column}{row}",
            PatternFill(patternType="solid", start_color="f2746b"),
        )
        for row in range(37, 42)
        for column in list("ABCDEFGH")
    ]
    + [
        (
            f"{column}{row}",
            PatternFill(patternType="solid", start_color="7beda3"),
        )
        for row in range(42, 47)
        for column in list("ABCDEFGH")
    ],
    "borders": {
        "cell_rows": [(f"A{row}:H{row}", THIN_BORDER) for row in range(24, 34)]
        + [(f"A{row}:H{row}", THIN_BORDER) for row in range(36, 47)]
        + [(f"A{row}:H{row}", THIN_BORDER) for row in range(49, 54)]
        + [(f"A{row}:H{row}", THIN_BORDER) for row in range(56, 61)],
    },
    "images": [
        {"cell": "A4", "img_index": 2, "size": (350, 700)},
        {"cell": "G4", "img_index": 1, "size": (350, 350)},
    ],
    "wrap_text": [
        f"{col}{row}" for col in list("ABCDE") for row in range(25, 34)
    ]
    + [f"{col}{row}" for col in list("ABCDE") for row in range(37, 47)]
    + [f"{col}{row}" for col in list("ABCDE") for row in range(50, 54)]
    + [f"{col}{row}" for col in list("ABCDE") for row in range(57, 61)],
    "row_height": [
        (row, 30)
        for start, end in [(25, 34), (37, 47), (50, 54), (57, 61)]
        for row in range(start, end)
    ],
    "dropdowns": [
        {
            "cells": {
                (
                    f"F{row}"
                    for start, end in [(25, 34), (37, 47), (50, 54), (57, 61)]
                    for row in range(start, end)
                ): (
                    '"Oncogenic, Likely oncogenic,'
                    "Uncertain, Likely passenger,"
                    'Likely artefact"'
                ),
            },
            "title": "Variant class",
        },
        {
            "cells": {
                (
                    f"G{row}"
                    for start, end in [(25, 34), (37, 47), (50, 54), (57, 61)]
                    for row in range(start, end)
                ): (
                    '"Predicts therapeutic response,'
                    "Prognostic"
                    "Defines diagnosis group,"
                    "Eligibility for trial,"
                    'Other"'
                ),
            },
            "title": "Actionability",
        },
    ],
}


def add_dynamic_values(
    SV_df: pd.DataFrame,
    fusion_count: int,
    SNV_df_columns: list,
    gain_df_columns: list,
    SV_df_columns: list,
    germline_df_columns: list,
) -> dict:
    """Add dynamic values for the Summary sheet

    Parameters
    ----------
    SV_df : pd.DataFrame
        Dataframe containing the data for SV fusion variants and appropriate
        additional data from inputs
    fusion_count : int
        Integer for the maximum number of fusion for a variant
    SNV_df_columns : list
        List of columns for the SNV dataframe
    gain_df_columns : list
        List of columns for the gain dataframe
    SV_df_columns : list
        List of columns for the SV dataframe
    germline_df_columns : list
        List of columns for the germline dataframe

    Returns
    -------
    dict
        Dict containing data that needs to be merged to the CONFIG variable
    """

    variant_class_column_letter = misc.get_column_letter_using_column_name(
        SV_df, "Variant class"
    )
    actionability_column_letter = misc.get_column_letter_using_column_name(
        SV_df, "Actionability"
    )
    comments_column_letter = misc.get_column_letter_using_column_name(
        SV_df, "Comments"
    )

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
            (63, col_index): col_name
            for col_index, col_name in enumerate(SNV_df_columns, 1)
        }
        | {
            (75, col_index): col_name
            for col_index, col_name in enumerate(gain_df_columns, 1)
        }
        | {
            (83, col_index): col_name
            for col_index, col_name in enumerate(SV_df_columns, 1)
        }
        | {
            (91, col_index): col_name
            for col_index, col_name in enumerate(germline_df_columns, 1)
        },
    }

    return config_with_dynamic_values
