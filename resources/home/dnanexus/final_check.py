import argparse
from pathlib import Path
import re

import pandas as pd


INPUT_PATTERNS = [
    r"[-_]reported_structural_variants\..*\.csv",
    r"[-_]reported_variants\..*\.csv",
    r"\..*\.supplementary\.html",
    r"\..*\.xlsx",
]


def get_sample_id_from_files(files: list, patterns: list) -> dict:
    """Get the sample id from the new files detected and link sample ids to
    their files

    Parameters
    ----------
    files : list
        List of DNAnexus file objects
    patterns : list
        Expected patterns for the file name

    Returns
    -------
    Dict
        Dict containing the sample id and their files
    """

    detected_sample_ids = set()

    for file in files:
        file_name = file.name
        # we have some .html files that we don't want to match
        matched_pattern = None

        for pattern in patterns:
            match = re.search(pattern, file_name)

            if match:
                matched_pattern = match

        if matched_pattern:
            detected_sample_ids.add(file_name[: matched_pattern.start()])

    file_dict = {}

    for sample_id in detected_sample_ids:
        file_dict.setdefault(sample_id, [])

        for file in files:
            if sample_id in file.name:
                file_dict[sample_id].append(file)

    return file_dict


def compare_somatic_snvs(
    output_somatic_snv: pd.DataFrame, rv: pd.DataFrame
) -> tuple:
    """Compare the somatic variants from the input csv and the output xlsx

    Parameters
    ----------
    output_somatic_snv : pd.DataFrame
        Dataframe with the somatic variants from the output xlsx
    rv : pd.DataFrame
        Dataframe with the variants from the input reported variants csv

    Returns
    -------
    tuple
        Tuple of lists containing the unique variants from the input csv and
        the output xlsx
    """

    rename_dict = {
        "GRCh38 coordinates;ref/alt allele": "GRCh38 coordinates",
        "CDS change and protein change": "Variant",
        "Predicted consequences": "Predicted consequences",
    }

    # check somatic sheet
    coor_SNV_output = {
        tuple(row)
        for i, row in output_somatic_snv[
            ["GRCh38 coordinates", "Variant", "Predicted consequences"]
        ].iterrows()
    }

    coor_SNV_input = rv[rv["Origin"].str.lower().str.contains("somatic")][
        [
            "GRCh38 coordinates;ref/alt allele",
            "CDS change and protein change",
            "Predicted consequences",
        ]
    ]

    coor_SNV_input["Predicted consequences"] = coor_SNV_input[
        "Predicted consequences"
    ].str.split(";", expand=True)[0]

    coor_SNV_input = {
        tuple(row)
        for i, row in coor_SNV_input.rename(columns=rename_dict).iterrows()
    }

    value_input = [
        " - ".join(ele) for ele in coor_SNV_input.difference(coor_SNV_output)
    ]
    value_output = [
        " - ".join(ele) for ele in coor_SNV_output.difference(coor_SNV_input)
    ]

    return value_input, value_output


def compare_germline_snvs(
    output_germline_snv: pd.DataFrame, rv: pd.DataFrame
) -> tuple:
    """Compare the germline variants from the input csv and the output xlsx

    Parameters
    ----------
    output_germline_snv : pd.DataFrame
        Dataframe with the germline variants from the output xlsx
    rv : pd.DataFrame
        Dataframe with the variants from the input reported variants csv

    Returns
    -------
    tuple
        Tuple of lists containing the unique variants from the input csv and
        the output xlsx
    """

    rename_dict = {
        "Unnamed: 1": "GRCh38 coordinates;ref/alt allele",
        "Unnamed: 2": "CDS change and protein change",
    }

    # remove the first 3 lines because they never should have variants
    coor_germline_output = output_germline_snv.iloc[3:, 1:3]
    coor_germline_output.iloc[:, 0] = coor_germline_output.iloc[
        :, 0
    ].str.replace("\n", ";")
    coor_germline_output.iloc[:, 1] = coor_germline_output.iloc[
        :, 1
    ].str.replace("\n", ";")
    coor_germline_output = coor_germline_output.dropna()

    coor_germline_input = {
        tuple(row)
        for i, row in rv[rv["Origin"].str.lower().str.contains("germline")][
            [
                "GRCh38 coordinates;ref/alt allele",
                "CDS change and protein change",
            ]
        ].iterrows()
    }

    coor_germline_output = {
        tuple(row)
        for i, row in coor_germline_output.rename(
            columns=rename_dict
        ).iterrows()
    }

    value_input = [
        " - ".join(ele)
        for ele in coor_germline_input.difference(coor_germline_output)
    ]
    value_output = [
        " - ".join(ele)
        for ele in coor_germline_output.difference(coor_germline_input)
    ]

    return value_input, value_output


def compare_gain_cnvs(
    output_gain_cnvs: pd.DataFrame, rsv: pd.DataFrame
) -> tuple:
    """Compare the gain variants from the input csv and the output xlsx

    Parameters
    ----------
    output_gain_cnvs : pd.DataFrame
        Dataframe with the gain variants from the output xlsx
    rsv : pd.DataFrame
        Dataframe with the variants from the input structural variants csv

    Returns
    -------
    tuple
        Tuple of lists containing the unique variants from the input csv and
        the output xlsx
    """

    coor_gain_output = {
        tuple(row)
        for i, row in output_gain_cnvs[
            ["Gene", "GRCh38 coordinates"]
        ].iterrows()
    }

    coor_gain_input = {
        tuple(row)
        for i, row in rsv[rsv["Type"].str.lower().str.contains("gain")][
            ["Gene", "GRCh38 coordinates"]
        ].iterrows()
    }

    value_input = [
        " - ".join(ele) for ele in coor_gain_input.difference(coor_gain_output)
    ]
    value_output = [
        " - ".join(ele) for ele in coor_gain_output.difference(coor_gain_input)
    ]

    return value_input, value_output


def compare_loss_cnvs(
    output_loss_cnvs: pd.DataFrame, rsv: pd.DataFrame
) -> tuple:
    """Compare the loss variants from the input csv and the output xlsx

    Parameters
    ----------
    output_loss_cnvs : pd.DataFrame
        Dataframe with the loss variants from the output xlsx
    rsv : pd.DataFrame
        Dataframe with the variants from the input structural variant csv

    Returns
    -------
    tuple
        Tuple of lists containing the unique variants from the input csv and
        the output xlsx
    """

    # check loss sheet
    coor_loss_output = {
        tuple(row)
        for i, row in output_loss_cnvs[
            ["Gene", "GRCh38 coordinates"]
        ].iterrows()
    }

    coor_loss_input = {
        tuple(row)
        for i, row in rsv[rsv["Type"].str.lower().str.match("loss|loh")][
            ["Gene", "GRCh38 coordinates"]
        ].iterrows()
    }

    value_input = [
        " - ".join(ele) for ele in coor_loss_input.difference(coor_loss_output)
    ]
    value_output = [
        " - ".join(ele) for ele in coor_loss_output.difference(coor_loss_input)
    ]

    return value_input, value_output


def compare_fusion_cnvs(
    output_sv_cnvs: pd.DataFrame, rsv: pd.DataFrame
) -> tuple:
    """Compare the fusion variants from the input csv and the output xlsx

    Parameters
    ----------
    output_sv_cnvs : pd.DataFrame
        Dataframe with the fusion variants from the output xlsx
    rsv : pd.DataFrame
        Dataframe with the variants from the input structural variants csv

    Returns
    -------
    tuple
        Tuple of lists containing the unique variants from the input csv and
        the output xlsx
    """

    # check SV sheet
    coor_sv_output = {
        tuple(row)
        for i, row in output_sv_cnvs[["Gene", "GRCh38 coordinates"]].iterrows()
    }

    coor_sv_input = {
        tuple(row)
        for i, row in rsv[
            ~rsv["Type"].str.lower().str.contains(r"loss|loh|gain")
        ][["Gene", "GRCh38 coordinates"]].iterrows()
    }

    value_input = [
        " - ".join(ele) for ele in coor_sv_input.difference(coor_sv_output)
    ]
    value_output = [
        " - ".join(ele) for ele in coor_sv_output.difference(coor_sv_input)
    ]

    return value_input, value_output


def parse_files(files: list) -> dict:
    """Parse the files depending on the suffix of the files

    Parameters
    ----------
    files : list
        List of files to parse

    Returns
    -------
    dict
        Dict of dataframes from parsing the files
    """

    dfs = {}

    for file in files:
        if file.name.endswith(".xlsx"):
            dfs["xlsx"] = pd.read_excel(file, sheet_name=None)
        elif file.name.endswith(".csv"):
            df = pd.read_csv(file)

            if "reported_structural_variants" in file.name:
                dfs["rsv"] = df
            elif "reported_variants" in file.name:
                dfs["rv"] = df

    return dfs


def main(folder):
    file_dict = get_sample_id_from_files(
        [Path(file) for file in Path(folder).glob("*")],
        INPUT_PATTERNS,
    )

    errors = []

    for sample, files in file_dict.items():
        file_dfs = parse_files(files)

        somatic_differences = compare_somatic_snvs(
            file_dfs["xlsx"]["SNV"], file_dfs["rv"]
        )
        germline_differences = compare_germline_snvs(
            file_dfs["xlsx"]["Germline"], file_dfs["rv"]
        )
        gain_differences = compare_gain_cnvs(
            file_dfs["xlsx"]["Gain"], file_dfs["rsv"]
        )
        loss_differences = compare_loss_cnvs(
            file_dfs["xlsx"]["Loss"], file_dfs["rsv"]
        )
        sv_differences = compare_fusion_cnvs(
            file_dfs["xlsx"]["SV"], file_dfs["rsv"]
        )

        for sheet, differences in {
            "somatic_snv": somatic_differences,
            "germline_snv": germline_differences,
            "gain_cnv": gain_differences,
            "loss_cnv": loss_differences,
            "fusion_cnv": sv_differences,
        }.items():
            input_diff, output_diff = differences

            if input_diff or output_diff:
                errors.append((sheet, input_diff, output_diff))

        if errors:
            msg = f"Unequal variants found in {sample}:\n"

            for sheet, input_diffs, output_diffs in errors:
                msg += f"- {sheet}\n"
                msg += f"  - Input : {' | '.join(sorted(input_diffs))}\n"
                msg += f"  - Output : {' | '.join(sorted(output_diffs))}\n"

            raise AssertionError(msg)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("folder")
    args = parser.parse_args()
    main(args.folder)
