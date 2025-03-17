import argparse

import pandas as pd

from utils import excel, html, vcf


def main(**kwargs):
    # prepare inputs and link type with the args
    inputs = {
        "hotspots": {"id": kwargs["hotspots"], "type": "csv"},
        "reference_gene_groups": {
            "id": kwargs["reference_gene_groups"],
            "type": "xls",
        },
        "clinvar": {"id": kwargs["clinvar"], "type": "vcf"},
        "clinvar_index": {"id": kwargs["clinvar_index"], "type": "index"},
        "supplementary_html": {
            "id": kwargs["supplementary_html"],
            "type": "html",
        },
        "reported_variants": {
            "id": kwargs["reported_variants"],
            "type": "csv",
        },
        "reported_structural_variants": {
            "id": kwargs["reported_structural_variants"],
            "type": "csv",
        },
    }

    # loop through the inputs to parse the files
    for name, info_dict in inputs.items():
        file = info_dict["id"]
        file_type = info_dict["type"]

        if file_type == "vcf":
            data = vcf.open_vcf(file)
        elif file_type == "xls" or file_type == "csv":
            data = excel.open_file(file, file_type)
        elif file_type == "html":
            data = html.open_html(file)

        inputs[name]["data"] = data

    # get images and tables from the html file
    html_images = html.get_images(inputs["supplementary_html"]["data"])
    html_tables = html.get_tables(inputs["supplementary_html"]["id"])

    with pd.ExcelWriter("output.xlsx", engine="openpyxl") as output_excel:
        excel.write_sheet(output_excel, "SOC")
        output_excel.book.save("output.xlsx")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-hs",
        "--hotspots",
        required=True,
        help="CSV file containing information about the cancer hotspots",
    )
    parser.add_argument(
        "-r",
        "--reference_gene_groups",
        required=True,
        help="Excel file obtained from the Solid cancer team",
    )
    parser.add_argument(
        "-c", "--clinvar", required=True, help="Clinvar asset VCF file"
    )
    parser.add_argument(
        "-i",
        "--clinvar_index",
        required=True,
        help="Clinvar asset VCF index file",
    )
    parser.add_argument(
        "-html",
        "--supplementary_html",
        required=True,
        help="HTML file from GEL",
    )
    parser.add_argument(
        "-rv",
        "--reported_variants",
        required=True,
        help="CSV/excel file from GEL containing info on reported variants",
    )
    parser.add_argument(
        "-rsv",
        "--reported_structural_variants",
        required=True,
        help="CSV/excel file from GEL containing info on reported structural variants",
    )

    main(**vars(parser.parse_args()))
