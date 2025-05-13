from openpyxl.styles import Border, Side
from openpyxl.styles.fills import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

from utils import misc

# prepare formatting
THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


CONFIG = {
    "col_width": [
        ("B", 18),
        ("C", 22),
        ("D", 22),
        ("E", 20),
        ("H", 16),
        ("I", 12),
        ("J", 14),
        ("M", 18),
    ],
    "freeze_panes": "F1",
    "expected_columns": [
        "Event domain",
        "Gene",
        "RefSeq IDs",
        "Impacted transcript region",
        "GRCh38 coordinates",
        "Chromosomal bands",
        "Size",
        (
            "Population germline allele frequency (GESG | GECG for somatic "
            "SVs or AF | AUC for germline CNVs)"
        ),
        "Paired reads",
        "Split reads",
        "Gene mode of action",
        "Variant class",
        "OG_Fusion",
        "OG_IntDup",
        "OG_IntDel",
        "Disruptive",
    ],
    "alternative_columns": [
        [
            (
                "Population germline allele frequency (GESG | GECG for "
                "somatic SVs or AF | AUC for germline CNVs)"
            ),
            (
                "Population germline allele frequency (AF | AUC for germline "
                "CNVs)"
            ),
        ]
    ],
    "row_height": [(1, 120)],
}


def add_dynamic_values(data: pd.DataFrame) -> dict:
    """Add dynamic values for the SV sheet

    Parameters
    ----------
    data : pd.DataFrame
        Dataframe containing the data for SV fusion variants and appropriate
        additional data from inputs

    Returns
    -------
    dict
        Dict containing data that needs to be merged to the CONFIG variable
    """

    nb_structural_variants = data.shape[0]

    last_column_letter = misc.get_column_letter_using_column_name(data)
    variant_class_column_letter = misc.get_column_letter_using_column_name(
        data, "Variant class"
    )
    variant_class_column_index = misc.convert_letter_column_to_index(
        variant_class_column_letter
    )

    first_letter_lookup_groups = variant_class_column_index + 4

    cells_to_color = []

    lookup_start, lookup_end = (
        first_letter_lookup_groups,
        misc.convert_letter_column_to_index(last_column_letter),
    )

    # there are 12 look up groups
    number_genes = (lookup_end - lookup_start + 1) / 12
    group_number = 1

    # build the cells to color data
    for i, index in enumerate(range(lookup_start, lookup_end + 1)):
        while i >= number_genes * group_number:
            group_number += 1

        if group_number % 2 == 0:
            pattern = PatternFill(patternType="solid", start_color="c4d9ef")
        else:
            pattern = PatternFill(patternType="solid", start_color="B8E7E0")

        for j in range(1, nb_structural_variants + 2):
            cells_to_color.append(
                (f"{misc.convert_index_to_letters(index)}{j}", pattern)
            )

    config_with_dynamic_values = {
        "cells_to_write": {
            (1, i): column for i, column in enumerate(data.columns, 1)
        }
        | {
            # remove the col and row index from the writing?
            (r_idx - 1, c_idx - 1): value
            for r_idx, row in enumerate(dataframe_to_rows(data), 1)
            for c_idx, value in enumerate(row, 1)
            if c_idx != 1 and r_idx != 1
        },
        "cells_to_colour": [
            (
                f"{misc.convert_index_to_letters(i)}{j}",
                PatternFill(patternType="solid", start_color="FFDBBB"),
            )
            for i in range(
                variant_class_column_index - 1, variant_class_column_index + 4
            )
            for j in range(1, nb_structural_variants + 2)
        ]
        + cells_to_color,
        "to_bold": [
            f"{misc.convert_index_to_letters(i)}1"
            for i in range(
                misc.convert_letter_column_to_index(last_column_letter) + 1
            )
        ],
        "borders": {"cell_rows": [(f"A1:{last_column_letter}1", THIN_BORDER)]},
        "text_orientation": [
            (f"{misc.convert_index_to_letters(i)}1", 90)
            for i in range(lookup_start, lookup_end + 1)
        ],
        "dropdowns": [
            {
                "cells": {
                    (
                        f"{variant_class_column_letter}{i}"
                        for i in range(2, nb_structural_variants + 2)
                    ): (
                        '"Oncogenic, Likely oncogenic,'
                        "Uncertain, Likely passenger,"
                        'Likely artefact"'
                    ),
                },
                "title": "Variant class",
            },
        ],
        "auto_filter": f"F:{last_column_letter}",
    }

    return config_with_dynamic_values
