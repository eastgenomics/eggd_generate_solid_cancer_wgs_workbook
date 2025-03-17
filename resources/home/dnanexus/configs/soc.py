import string

from openpyxl.styles import Border, Side
from openpyxl.styles.fills import PatternFill

# prepare formatting
THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "tables": [
        {
            "headers": {
                (1, 1): "Patient Details (Epic demographics)",
                (1, 3): "Previous testing",
                (2, 1): "NAME",
                (2, 3): "Alteration",
                (2, 4): "Assay",
                (2, 5): "Result",
                (2, 6): "WGS concordance",
                (3, 1): "Sex, Age, DOB",
                (4, 1): "Phone number",
                (5, 1): "MRN",
                (6, 1): "NHS Number",
                (8, 1): "Histology",
                (12, 1): "Comments",
            }
        }
    ],
    "to_merge": {
        "start_row": 1,
        "end_row": 1,
        "start_column": 3,
        "end_column": 6,
    },
    "to_align": ["C1", "C2", "D2", "E2", "F2"],
    "to_bold": ["A1", "A8", "A12", "A16", "C1"],
    "col_width": (
        ("A", 32),
        ("C", 16),
        ("E", 16),
        ("D", 26),
        ("F", 26),
    ),
    "cells_to_colour": [
        ("C3", PatternFill(patternType="solid", start_color="90EE90")),
        ("D3", PatternFill(patternType="solid", start_color="90EE90")),
        ("E3", PatternFill(patternType="solid", start_color="90EE90")),
        ("F3", PatternFill(patternType="solid", start_color="90EE90")),
        ("C4", PatternFill(patternType="solid", start_color="90EE90")),
        ("D4", PatternFill(patternType="solid", start_color="90EE90")),
        ("E4", PatternFill(patternType="solid", start_color="90EE90")),
        ("F4", PatternFill(patternType="solid", start_color="90EE90")),
    ],
    "borders": {
        "single_cells": [
            # generate list of letter and numbers from C-F and 1-8
            # i.e. C1, C2, C3 ..
            (f"{string.ascii_uppercase[i]}{j}", THIN_BORDER)
            for i in range(2, 6)
            for j in range(1, 9)
        ]
        + [("A1", LOWER_BORDER), ("A8", LOWER_BORDER), ("A12", LOWER_BORDER)],
    },
    "dropdowns": {
        "cells": {
            (f"D{i}" for i in range(3, 9)): (
                '"FISH,IHC,NGS,Sanger,NGS multi-gene panel,'
                "RNA fusion panel,SNP array, Methylation array,"
                "MALDI-TOF, MLPA, MS-MLPA, Chromosome breakage,"
                'Digital droplet PCR, RT-PCR, LR-PCR"'
            ),
            (f"E{i}" for i in range(3, 9)): '"Detected, Not detected"',
            (f"F{i}" for i in range(3, 9)): (
                '"Novel,Concordant (detected),'
                "Concordant (undetected),"
                "Disconcordant (detected),"
                'Disconcordant (undetected),N/A"'
            ),
        },
        "title": "",
    },
}
