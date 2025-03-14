import openpyxl
from openpyxl.styles import Alignment, DEFAULT_FONT, Font
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd

from utils import misc


def open_file(file: str, file_type: str) -> pd.DataFrame:
    """Read in CSV or XLS files using pandas

    Parameters
    ----------
    file : str
        File path
    file_type : str
        File type with the file path

    Returns
    -------
    pd.DataFrame
        Dataframe created by pandas
    """

    if file_type == "csv":
        return pd.read_csv(file)
    elif file_type == "xls":
        return pd.read_excel(file)


def write_sheet(
    excel_writer: pd.ExcelWriter, sheet_name: str
) -> openpyxl.worksheet.worksheet.Worksheet:
    """Using a config file, write in the appropriate data

    Parameters
    ----------
    excel_writer : pd.ExcelWriter
        ExcelWriter object
    sheet_name : str
        Name of the sheet used to match the config

    Returns
    -------
    openpyxl.worksheet.worksheet.Worksheet
        Worksheet object
    """

    sheet = excel_writer.book.create_sheet(sheet_name)

    type_config = misc.select_config(sheet_name)
    assert type_config, "Config file couldn't be imported"

    for table in type_config.CONFIG["tables"]:
        headers = table["headers"]

        for cell_x, cell_y in headers:
            value = headers[cell_x, cell_y]
            sheet.cell(cell_x, cell_y).value = value

    if type_config.CONFIG.get("to_merge"):
        # merge columns that have longer text
        sheet.merge_cells(**type_config.CONFIG.get("to_merge"))

    # align cells
    for cell in type_config.CONFIG["to_align"]:
        sheet[cell].alignment = Alignment(wrapText=True, horizontal="center")
        if cell != "C1":
            sheet[cell].font = Font(italic=True)

    # titles to set to bold
    for cell in type_config.CONFIG["to_bold"]:
        sheet[cell].font = Font(bold=True, name=DEFAULT_FONT.name)

    # set column widths for readability
    for cell, width in type_config.CONFIG["col_width"]:
        sheet.column_dimensions[cell].width = width

    # colour cells
    for cell, color in type_config.CONFIG["cells_to_colour"]:
        sheet[cell].fill = color

    if type_config.CONFIG.get("borders"):
        if type_config.CONFIG.get("borders").get("single_cells"):
            # set borders around table areas
            for cells, type_border in type_config.CONFIG["borders"][
                "single_cells"
            ]:
                sheet[cell].border = type_border

        if type_config.CONFIG.get("borders").get("cell_rows"):
            for cell_range, type_border in type_config.CONFIG.get(
                "borders"
            ).get("cell_rows"):
                for cells in sheet[cell_range]:
                    for cell in cells:
                        cell.border = type_border

    for cells, options in type_config.CONFIG["dropdowns"].items():
        dropdown = DataValidation(
            type="list", formula1=options, allow_blank=True
        )
        dropdown.prompt = "Select from the list"
        dropdown.promptTitle = ""
        dropdown.showInputMessage = True
        dropdown.showErrorMessage = True
        sheet.add_data_validation(dropdown)

        for cell in cells:
            dropdown.add(sheet[cell])

    return sheet
