import os
from datetime import datetime, timezone, timedelta

from openpyxl.styles import Border, Side

from utils import dnanexus

THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)

CONFIG = {
    "cells_to_write": {
        (1, 1): "Project id",
        (2, 1): (
            os.environ["DX_PROJECT_CONTEXT_ID"]
            if "DX_PROJECT_CONTEXT_ID" in os.environ
            else "Id not retrievable"
        ),
        (4, 1): "Job id",
        (5, 1): (
            os.environ["DX_JOB_ID"]
            if "DX_JOB_ID" in os.environ
            else "Id not retrievable"
        ),
        (7, 1): "Job datetime",
        (8, 1): datetime.now()
        .replace(tzinfo=timezone(timedelta(hours=1)))
        .strftime("%a %d %b %Y, %H:%M"),
        (1, 3): "Refgene file used",
        (2, 3): dnanexus.get_refgene_input_file_info(),
        (4, 3): "App version",
        (5, 3): dnanexus.get_app_version(),
    },
    "to_bold": ["A1", "A4", "A7", "C1", "C4"],
    "borders": {
        "single_cells": [
            ("A1", LOWER_BORDER),
            ("A4", LOWER_BORDER),
            ("A7", LOWER_BORDER),
            ("C1", LOWER_BORDER),
            ("C4", LOWER_BORDER),
        ]
    },
    "col_width": [
        ("A", 36),
        ("C", 36),
    ],
}
