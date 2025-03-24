from openpyxl.styles import Border, Side

THIN = Side(border_style="thin", color="000000")
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "cells_to_write": {
        (1, 1): "=SOC!A2",
        (2, 1): "=SOC!A3",
        (1, 3): "=SOC!A5",
        (2, 3): "=SOC!A6",
        (1, 5): "=SOC!A9",
        (38, 1): "Pertinent chromosomal CNVs",
        (39, 1): "None",
    },
    "to_bold": [
        "A1",
        "A38",
    ],
    "col_width": [
        ("A", 18),
        ("B", 22),
        ("C", 18),
        ("D", 22),
        ("E", 22),
    ],
    "borders": {
        "single_cells": [
            ("A38", LOWER_BORDER),
        ],
    },
    "images": [
        {"cell": "A4", "img_index": 2, "size": (600, 1000)},
        {"cell": "K4", "img_index": 1, "size": (500, 500)},
    ],
}
