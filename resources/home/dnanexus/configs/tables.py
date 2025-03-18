import pandas as pd


# Config list for the tables present in the HTML file i.e.
# - First table is Patient info and has Clinical indication as headers
CONFIG = [
    {
        "name": "Patient info",
        "expected_headers": ["Clinical Indication"],
        "alternatives": [],
    },
    {
        "name": "Tumor info",
        "expected_headers": [
            "Tumour Diagnosis Date",
            "Histopathology or SIHMDS LAB ID",
            "Presentation",
            "Primary or Metastatic",
            "Tumour Topography",
        ],
        "alternatives": [],
    },
    {
        "name": "Sample info",
        "expected_headers": [
            "Clinical Sample Date Time",
            "Storage Medium",
            "Source",
            "Tumour Content",
            "Calculated Tumour Content",
            "Calculated Overall Ploidy",
        ],
        "alternatives": [],
    },
    {
        "name": "Germline info",
        "expected_headers": [
            "Storage Medium",
            "Source",
        ],
        "alternatives": [],
    },
    {
        "name": "Sequencing info",
        "expected_headers": [
            "Total somatic SNVs",
            "Total somatic indels",
            "Total somatic SVs",
            "Sample type",
            "Genome-wide coverage mean, x",
            "Mapped reads, %",
            "Chimeric DNA fragments, %",
            "Insert size median, bp",
            "Unevenness of local genome coverage, x",
        ],
        "alternatives": [
            [
                "Unevenness of local genome coverage, x",
                "Genome coverage evenness",
            ]
        ],
    },
]


def get_table_value(
    config_name: str,
    row: int,
    column: str,
    tables: list,
    formatting: str = None,
) -> str:
    """Get the table value in the matched df for writing in the worksheet

    Parameters
    ----------
    config_name : str
        Name of the table in the table config file
    row : int
        Number of the row to look data in
    column : str
        Column name to look data in
    tables : list
        List of the tables stored in the html
    formatting : str, optional
        String describing how to modify the table value, by default None

    Returns
    -------
    str
        Value from the dataframe
    """

    for table_name, table_info in tables.items():
        if config_name == table_name:
            # check if we have an alternative for the given column name i.e.
            # that column name is already not present
            if column in table_info["alternatives"]:
                column = table_info["alternatives"][column]

            value_to_return = table_info["data"].loc[row, column]

            if formatting:
                # hardcoded way to reformat the df value to extract
                if formatting == "split":
                    value_to_return = value_to_return.split("_")[0]
                elif formatting == "parentheses":
                    value_to_return = f"({value_to_return})"

    return value_to_return


def find_headers(
    table: pd.DataFrame, expected_headers: list, alternatives: list
):
    """Validation step for checking that the appropriate tables have the
    appropriate headers

    Parameters
    ----------
    table : pd.DataFrame
        Dataframe to check headers
    expected_headers : list
        List of headers to check
    alternatives: list
        List of potential alternatives if the header is not found

    Returns
    -------
    dict
        dict for mapping for the header and its alternative
    """

    for header in expected_headers:
        if header not in table.columns:
            # can't find the header so try the alternatives
            for alternative_set in alternatives:
                if header in alternative_set:
                    for alternative in alternative_set:
                        if alternative in table.columns:
                            return {header: alternative}

            # Could find an alternative
            raise Exception(
                (
                    f"{header} is not present in the expected column "
                    "and no alternatives is present to match "
                    "potential other versions"
                )
            )

    return {}
