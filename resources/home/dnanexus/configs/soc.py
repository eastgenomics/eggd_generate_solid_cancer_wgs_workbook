import string

from openpyxl.styles import Border, Side

# prepare formatting
THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "cells_to_write": {
        (1, 1): "Patient Details (Epic demographics)",
        (1, 3): "Previous testing",
        (2, 1): "NAME",
        (3, 1): "Sex, Age, DOB",
        (4, 1): "Phone number",
        (5, 1): "MRN",
        (6, 1): "NHS Number",
        (8, 1): "Histological diagnosis",
        (12, 1): "Comments",
    },
    "to_merge": {
        "start_row": 1,
        "end_row": 1,
        "start_column": 3,
        "end_column": 6,
    },
    "alignment_info": [("C1", {"horizontal": "center", "wrapText": True})],
    "to_bold": ["A1", "A8", "A12", "C1"],
    "col_width": [
        ("A", 32),
        ("C", 16),
        ("E", 16),
        ("D", 26),
        ("F", 26),
    ],
    "borders": {
        "single_cells": [
            # generate list of letter and numbers from C-F and 1-8
            # i.e. C1, C2, C3 ..
            (f"{string.ascii_uppercase[i]}1", THIN_BORDER)
            for i in range(2, 6)
        ]
        + [
            ("A1", LOWER_BORDER),
            ("A8", LOWER_BORDER),
            ("A12", LOWER_BORDER),
        ],
    },
}
