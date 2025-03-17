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
                (4, 1): "Diagnosis Date",
                (4, 2): "Tumour Received",
                (4, 3): "Tumour ID",
                (4, 4): "Presentation",
                (4, 5): "Diagnosis",
                (4, 6): "Tumour Site",
                (4, 7): "Tumour Type",
                (4, 8): "Germline Sample",
            }
        },
        {
            "headers": {
                (7, 1): "Purity (Histo)",
                (7, 2): "Purity (Calculated)",
                (7, 3): "Ploidy",
                (7, 4): "Total SNVs",
                (7, 5): "Total Indels",
                (7, 6): "Total SVs",
                (7, 7): "TMB",
            }
        },
        {
            "headers": {
                (10, 1): "Sample type",
                (10, 2): "Mean depth, x",
                (10, 3): "Mapped reads, %",
                (10, 4): "Chimeric DNA frag, %",
                (10, 5): "Insert size, bp",
                (10, 6): "Unevenness, x",
            }
        },
        {
            "headers": {
                (1, 1): "=SOC!A2",
                (2, 1): "=SOC!A3",
                (1, 3): "=SOC!A5",
                (2, 3): "=SOC!A6",
                (1, 5): "=SOC!A9",
                (15, 1): "QC alerts",
                (16, 1): "None",
            }
        },
    ],
    "to_bold": [
        "A1",
        "A4",
        "A7",
        "A10",
        "A15",
        "B4",
        "B7",
        "B10",
        "C4",
        "C7",
        "C10",
        "D4",
        "D7",
        "D10",
        "E4",
        "E7",
        "E10",
        "F4",
        "F7",
        "F10",
        "G4",
        "G7",
        "H4",
    ],
    "col_width": (
        ("A", 22),
        ("B", 22),
        ("C", 22),
        ("D", 22),
        ("E", 22),
        ("F", 22),
        ("G", 22),
        ("H", 22),
        ("I", 22),
        ("J", 22),
    ),
    "cells_to_colour": [
        ("A4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("B4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("C4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("D4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("E4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("F4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("G4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("H4", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("A7", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("B7", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("C7", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("D7", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("E7", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("F7", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("G7", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("A10", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("B10", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("C10", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("D10", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("E10", PatternFill(patternType="solid", start_color="ADD8E6")),
        ("F10", PatternFill(patternType="solid", start_color="ADD8E6")),
    ],
    "borders": {
        "single_cells": [
            ("A15", LOWER_BORDER),
        ],
        "cell_rows": [
            ("A4:H4", THIN_BORDER),
            ("A5:H5", THIN_BORDER),
            ("A7:G7", THIN_BORDER),
            ("A8:G8", THIN_BORDER),
            ("A10:F10", THIN_BORDER),
            ("A11:F11", THIN_BORDER),
            ("A12:F12", THIN_BORDER),
        ],
    },
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
