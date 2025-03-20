from openpyxl.styles import Border, Side

THIN = Side(border_style="thin", color="000000")
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "tables": [
        {
            "headers": {
                (1, 1): "=SOC!A2",
                (2, 1): "=SOC!A3",
                (1, 3): "=SOC!A5",
                (2, 3): "=SOC!A6",
                (1, 5): "=SOC!A9",
            }
        },
        {
            "headers": {
                (4, 1): "Gene",
                (4, 2): "GRCh38 Coordinates",
                (4, 3): "Variant",
                (4, 4): "Consequence",
                (4, 5): "Genotype",
                (4, 6): "Variant Class",
                (4, 7): "Actionability",
                (4, 8): "Role in Cancer",
                (4, 9): "ClinVar",
                (4, 10): "gnomAD",
                (4, 11): "Tumour VAF",
            }
        },
    ],
    "to_bold": [
        "A1",
        "A38",
    ],
    "col_width": (
        ("A", 18),
        ("B", 22),
        ("C", 18),
        ("D", 22),
        ("E", 22),
    ),
    "borders": {
        "single_cells": [
            ("A38", LOWER_BORDER),
        ],
    },
    "images": [
        {"cell": "A4", "img_index": 2, "size": (600, 1000)},
        {"cell": "K4", "img_index": 1, "size": (500, 500)},
    ],
    "dropdowns": {
        "cells": {
            ("A16",): (
                '"None,<30% tumour purity,SNVs low VAF (<6%),TINC (<5%),'
                'Somatic CNV, Germline CNV"'
            ),
        },
        "title": "QC alerts",
    },
}
