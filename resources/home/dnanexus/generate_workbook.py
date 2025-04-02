import argparse

import pandas as pd

from configs import tables, germline, snv, gain, loss, sv, summary
from utils import excel_parsing, excel_writing, html, vcf


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
            data = excel_parsing.open_file(file, file_type)
        elif file_type == "html":
            data = html.open_html(file)

        inputs[name]["data"] = data

    refgene_dfs = excel_parsing.process_refgene(
        inputs["reference_gene_groups"]["data"]
    )

    # list of tuple allowing:
    # - the writing of the column (1st element)
    # - by mapping 2nd element to 4th element
    # - using the 3rd element as a reference df
    # - and getting the 5th element from the reference column
    lookup_refgene_data = (
        ("COSMIC", "Gene", refgene_dfs["cosmic"], "Gene", "Entities"),
        ("Paed", "Gene", refgene_dfs["paed"], "Gene", "Driver"),
        ("Sarc", "Gene", refgene_dfs["sarc"], "Gene", "Driver"),
        ("Neuro", "Gene", refgene_dfs["neuro"], "Gene", "Driver"),
        ("Ovary", "Gene", refgene_dfs["ovarian"], "Gene", "Driver"),
        ("Haem", "Gene", refgene_dfs["haem"], "Gene", "Driver"),
    )

    germline_df = excel_parsing.process_reported_variants_germline(
        inputs["reported_variants"]["data"],
        inputs["clinvar"]["data"],
    )
    somatic_df = excel_parsing.process_reported_variants_somatic(
        inputs["reported_variants"]["data"],
        lookup_refgene_data,
        inputs["hotspots"]["data"],
    )
    gain_df = excel_parsing.process_reported_SV(
        inputs["reported_structural_variants"]["data"],
        lookup_refgene_data,
        "gain",
    )
    loss_df = excel_parsing.process_reported_SV(
        inputs["reported_structural_variants"]["data"],
        lookup_refgene_data,
        "loss|loh",
    )
    fusion_df, fusion_count = excel_parsing.process_fusion_SV(
        inputs["reported_structural_variants"]["data"], lookup_refgene_data
    )

    dynamic_values_per_sheet = {
        "Germline": germline.add_dynamic_values(germline_df),
        "SNV": snv.add_dynamic_values(somatic_df),
        "Gain": gain.add_dynamic_values(gain_df),
        "Loss": loss.add_dynamic_values(loss_df),
        "SV": sv.add_dynamic_values(fusion_df),
        "Summary": summary.add_dynamic_values(
            fusion_df,
            fusion_count,
            list(somatic_df.columns),
            list(gain_df.columns),
            list(fusion_df.columns),
        ),
    }

    # get images and tables from the html file
    html_images = html.download_images(inputs["supplementary_html"]["data"])
    html_tables = html.get_tables(inputs["supplementary_html"]["id"])

    data_tables = {}

    # validate the tables as the order in the html and the config file should
    # be the same
    for i in range(len(tables.CONFIG)):
        alternative_headers = tables.find_alternative_headers(
            html_tables[i],
            tables.CONFIG[i]["expected_headers"],
            tables.CONFIG[i]["alternatives"],
        )

        data_tables[tables.CONFIG[i]["name"]] = {
            "data": html_tables[i],
            "alternatives": alternative_headers,
        }

    with pd.ExcelWriter("output.xlsx", engine="openpyxl") as output_excel:
        excel_writing.write_sheet(output_excel, "SOC")
        excel_writing.write_sheet(
            output_excel,
            "QC",
            html_tables=data_tables,
            soup=inputs["supplementary_html"]["data"],
        )
        excel_writing.write_sheet(
            output_excel,
            "Plot",
            html_images=html_images,
        )
        excel_writing.write_sheet(
            output_excel,
            "Signatures",
            html_images=html_images,
        )
        excel_writing.write_sheet(
            output_excel,
            "Germline",
            dynamic_data=dynamic_values_per_sheet,
        )
        excel_writing.write_sheet(
            output_excel,
            "SNV",
            dynamic_data=dynamic_values_per_sheet,
        )
        excel_writing.write_sheet(
            output_excel,
            "Gain",
            dynamic_data=dynamic_values_per_sheet,
        )
        excel_writing.write_sheet(
            output_excel,
            "Loss",
            dynamic_data=dynamic_values_per_sheet,
        )
        excel_writing.write_sheet(
            output_excel,
            "SV",
            dynamic_data=dynamic_values_per_sheet,
        )
        excel_writing.write_sheet(
            output_excel,
            "Summary",
            dynamic_data=dynamic_values_per_sheet,
            html_images=html_images,
        )


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
        help=(
            "CSV/excel file from GEL containing info on reported structural "
            "variants"
        ),
    )

    main(**vars(parser.parse_args()))
