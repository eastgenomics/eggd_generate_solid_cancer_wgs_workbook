from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, DEFAULT_FONT, Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

from configs.tables import get_table_value
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
    excel_writer: pd.ExcelWriter,
    sheet_name: str,
    data_tables: list = None,
    soup: BeautifulSoup = None,
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

    if type_config.CONFIG.get("tables"):
        write_tables(sheet, type_config.CONFIG["tables"], data_tables, soup)

    if type_config.CONFIG.get("to_merge"):
        # merge columns that have longer text
        sheet.merge_cells(**type_config.CONFIG["to_merge"])

    if type_config.CONFIG.get("to_align"):
        align_cells(sheet, type_config.CONFIG["to_align"])

    if type_config.CONFIG.get("to_bold"):
        bold_cells(sheet, type_config.CONFIG["to_bold"])

    if type_config.CONFIG.get("col_width"):
        set_col_width(sheet, type_config.CONFIG["col_width"])

    if type_config.CONFIG.get("cells_to_colour"):
        color_cells(sheet, type_config.CONFIG["cells_to_colour"])

    if type_config.CONFIG.get("borders"):
        draw_borders(sheet, type_config.CONFIG["borders"])

    if type_config.CONFIG.get("dropdowns"):
        generate_dropdowns(sheet, type_config.CONFIG["dropdowns"])

    return sheet


def write_tables(
    sheet: Worksheet, config_data: list, data_tables: list, soup: BeautifulSoup
):
    """Write the tables from the config

    Parameters
    ----------
    sheet : Worksheet
        Worksheet to write the tables into
    config_data : list
        List of tables to write
    data_tables: list
        List of dict for table configuration
    soup: BeautifulSoup
        HTML page
    """

    for table in config_data:
        headers = table["headers"]

        for cell_x, cell_y in headers:
            value_to_write = headers[cell_x, cell_y]
            sheet.cell(cell_x, cell_y).value = value_to_write

        if table.get("values"):
            values = table.get("values")

            for cell_x, cell_y in values:
                # if the value is a list, it means that concatenation is
                # required
                if isinstance(values[cell_x, cell_y], list):
                    value_to_write = []

                    for table_name, row, column, formatting in values[
                        cell_x, cell_y
                    ]:
                        subvalue = get_table_value(
                            table_name,
                            row,
                            column,
                            data_tables,
                            formatting,
                        )
                        value_to_write.append(subvalue)

                    value_to_write = " ".join(value_to_write)

                # single value to add in the table
                elif isinstance(values[cell_x, cell_y], tuple):
                    table_name, row, column = values[cell_x, cell_y]
                    value_to_write = get_table_value(
                        table_name, row, column, data_tables
                    )
                else:
                    # special hardcoded case, haven't found a way to make that
                    # better for now (which means it'll probably stay that way
                    # forever)
                    value_to_write = values[cell_x, cell_y](
                        soup,
                        "b",
                        (
                            "Total number of somatic non-synonymous small "
                            "variants per megabase"
                        ),
                    )

                sheet.cell(cell_x, cell_y).value = value_to_write


def align_cells(sheet: Worksheet, config_data: list):
    """For given list of cells, align the cells

    Parameters
    ----------
    sheet : Worksheet
        Worksheet in which to align the cells
    config_data : list
        List of cells to align
    """

    for cell in config_data:
        sheet[cell].alignment = Alignment(wrapText=True, horizontal="center")


def bold_cells(sheet: Worksheet, config_data: list):
    """Given a list of cells, bold them

    Parameters
    ----------
    sheet : Worksheet
        Worksheet in which to bold the cells
    config_data : list
        List of cells to bold
    """

    for cell in config_data:
        sheet[cell].font = Font(bold=True, name=DEFAULT_FONT.name)


def set_col_width(sheet: Worksheet, config_data: list):
    """Given a list of columns, set their width

    Parameters
    ----------
    sheet : Worksheet
        Worksheet in which to set the width
    config_data : list
        List of tuple with the column and its width to set
    """

    for cell, width in config_data:
        sheet.column_dimensions[cell].width = width


def color_cells(sheet: Worksheet, config_data: list):
    """Given a list of cells and their color, color the cells appropriately

    Parameters
    ----------
    sheet : Worksheet
        Worksheet to color the cells in
    config_data : list
        List of tuples with the cells and their color
    """

    for cell, color in config_data:
        sheet[cell].fill = color


def draw_borders(sheet: Worksheet, config_data: dict):
    """Draw borders around the cells

    Parameters
    ----------
    sheet : Worksheet
        Worksheet in which to draw borders
    config_data : dict
        Dict containing info for the single cells to draw borders around and
        the rows of cells
    """

    if config_data.get("single_cells"):
        for cell, type_border in config_data["single_cells"]:
            sheet[cell].border = type_border

    if config_data.get("cell_rows"):
        for cell_range, type_border in config_data["cell_rows"]:
            for cells in sheet[cell_range]:
                for cell in cells:
                    cell.border = type_border


def generate_dropdowns(sheet: Worksheet, config_data: dict):
    """Write in the dropdown menus

    Parameters
    ----------
    sheet : Worksheet
        Worksheet in which to write the dropdown menus
    config_data : dict
        Dict of data for the dropdown menus
    """

    for cells, options in config_data["cells"].items():
        dropdown = DataValidation(
            type="list", formula1=options, allow_blank=True
        )
        dropdown.prompt = "Select from the list"
        dropdown.promptTitle = config_data["title"]
        dropdown.showInputMessage = True
        dropdown.showErrorMessage = True
        sheet.add_data_validation(dropdown)

        for cell in cells:
            dropdown.add(sheet[cell])
