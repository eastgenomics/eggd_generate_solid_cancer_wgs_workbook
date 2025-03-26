import re

import pandas as pd
import vcfpy

from utils import vcf


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
        df = pd.read_csv(file)
    elif file_type == "xls":
        df = pd.read_excel(file, sheet_name=None)

    # convert the clinvar id column as a string and remove the trailing .0 that
    # the automatic conversion that pandas applies added
    if df is pd.DataFrame and "ClinVar ID" in df.columns:
        df["ClinVar ID"] = df["ClinVar ID"].astype(str)
        df["ClinVar ID"] = df["ClinVar ID"].str.removesuffix(".0")

    return df


def process_reported_variants_germline(
    df: pd.DataFrame, clinvar_resource: vcfpy.Reader
) -> pd.DataFrame:
    """Process the data from the reported variants excel file

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe from parsing the reported variants excel file
    clinvar_resource : vcfpy.Reader
        vcfpy.Reader object from the Clinvar resource

    Returns
    -------
    pd.DataFrame
        Dataframe containing clinical significance info for germline variants
    """

    df = df[df["Origin"].str.lower() == "germline"]

    if df.empty:
        return None

    df.reset_index(drop=True, inplace=True)

    clinvar_ids_to_find = [
        value for value in df.loc[:, "ClinVar ID"].to_numpy()
    ]
    clinvar_info = vcf.find_clinvar_info(
        clinvar_resource, *clinvar_ids_to_find
    )

    # add the clinvar info by merging the clinvar dataframe
    df = df.merge(clinvar_info, on="ClinVar ID", how="left")

    # split the col to get gnomAD
    df[["GE", "gnomAD"]] = df[
        "Population germline allele frequency (GE | gnomAD)"
    ].str.split("|", expand=True)

    df.drop(
        ["GE", "Population germline allele frequency (GE | gnomAD)"],
        axis=1,
        inplace=True,
    )
    df.loc[:, "Variant Class"] = ""
    df.loc[:, "Actionability"] = ""
    df = df[
        [
            "Gene",
            "GRCh38 coordinates;ref/alt allele",
            "CDS change and protein change",
            "Predicted consequences",
            "Genotype",
            "Variant Class",
            "Actionability",
            "Gene mode of action",
            "clnsigconf",
            "gnomAD",
        ]
    ]

    df.fillna("", inplace=True)

    return df


def process_reported_variants_somatic(
    df: pd.DataFrame, refgene_dfs: dict, hotspots_df: pd.DataFrame
) -> pd.DataFrame:
    """Get the somatic variants and format the data for them

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe from parsing the reported variants excel file

    Returns
    -------
    pd.DataFrame
        Dataframe with additional formatting for c. and p. annotation
    """

    # select only somatic rows
    df = df[df["Origin"].str.lower().str.contains("somatic")]
    df.reset_index(drop=True, inplace=True)
    df[["c_dot", "p_dot"]] = df["CDS change and protein change"].str.split(
        r"(?=;p)", n=1, expand=True
    )
    df["p_dot"] = df["p_dot"].str.slice(1)

    df["MTBP c."] = df["Gene"] + ":" + df["c_dot"]
    df["MTBP p."] = df["Gene"] + ":" + df["p_dot"]
    df.fillna({"MTBP p.": ""}, inplace=True)

    # convert string like: NRAS:p.Gln61Arg to NRAS:p.Gln61 for lookup in the
    # hotspots excel
    df["HS p."] = df["MTBP p."].apply(
        lambda x: (
            x[: re.search(r":p.[A-Za-z]+[0-9]+", x).end()]
            if re.search(r":p.[A-Za-z]+[0-9]+", x)
            else x
        )
    )

    # populate the somatic variant dataframe with data from the refgene excel
    # file
    lookup_refgene = (
        ("COSMIC", "Gene", refgene_dfs["cosmic"], "Gene", "Entities"),
        ("Paed", "Gene", refgene_dfs["paed"], "Gene", "Driver"),
        ("Sarc", "Gene", refgene_dfs["sarc"], "Gene", "Driver"),
        ("Neuro", "Gene", refgene_dfs["neuro"], "Gene", "Driver"),
        ("Ovary", "Gene", refgene_dfs["ovarian"], "Gene", "Driver"),
        ("Haem", "Gene", refgene_dfs["haem"], "Gene", "Driver"),
        ("HS_Sample", "HS p.", hotspots_df, "HS_PROTEIN_ID", "HS_Samples"),
        (
            "HS_Tumour",
            "HS p.",
            hotspots_df,
            "HS_PROTEIN_ID",
            "HS_Tumor Type Composition",
        ),
    )

    for (
        new_column,
        col_to_map,
        reference_df,
        col_to_index,
        col_to_look_up,
    ) in lookup_refgene:
        df[new_column] = df[col_to_map].map(
            reference_df.set_index(col_to_index)[col_to_look_up]
        )
        df[new_column] = df[new_column].fillna("-")

    df.loc[:, "Error flag"] = ""

    df["con_count"] = df["Predicted consequences"].str.count(r"\;")

    if df["con_count"].max() > 0:
        df[["Predicted consequences", "Error flag"]] = df[
            "Predicted consequences"
        ].str.split(";", expand=True)

    df.loc[:, "LOH"] = ""

    df["VAF"] = df["VAF"].astype("str")
    df["VAF_count"] = df["VAF"].str.count(r"\;")

    if df["VAF_count"].max() > 0:
        df[["VAF", "LOH"]] = df["VAF"].str.split(";", expand=True)

    df.loc[:, "Variant class"] = ""
    df.loc[:, "Actionability"] = ""
    df.loc[:, "Comments"] = ""
    df = df[
        [
            "Domain",
            "Gene",
            "GRCh38 coordinates;ref/alt allele",
            "CDS change and protein change",
            "Predicted consequences",
            "VAF",
            "LOH",
            "Error flag",
            "Alt allele/total read depth",
            "Gene mode of action",
            "Variant class",
            "Actionability",
            "Comments",
            "COSMIC",
            "Paed",
            "Sarc",
            "Neuro",
            "Ovary",
            "Haem",
            "HS_Sample",
            "HS_Tumour",
            "MTBP c.",
            "MTBP p.",
        ]
    ]
    df.rename(
        columns={
            "GRCh38 coordinates;ref/alt allele": "GRCh38 coordinates",
            "CDS change and protein change": "Variant",
        },
        inplace=True,
    )
    df.sort_values(["Domain", "VAF"], ascending=[True, False], inplace=True)
    df = df.replace([None], [""], regex=True)
    df["VAF"] = df["VAF"].astype(float)

    return df


def process_reported_SV(
    df: pd.DataFrame, refgene_dfs: dict, type_sv: str
) -> tuple:
    """Process the reported structural variants excel

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe containing data from the structural variants excel
    refgene_dfs : dict
        Dict of dataframes from the refgene excel

    Returns
    -------
    tuple
        Tuple of the dataframes for Gain and Loss structural variants
    """

    sv_df = df[df["Type"].str.lower().str.contains(type_sv)]
    sv_df.reset_index(drop=True, inplace=True)

    # populate the structural variant dataframe with data from the refgene
    # excel file
    lookup_refgene = (
        ("COSMIC", "Gene", refgene_dfs["cosmic"], "Gene", "Entities"),
        ("Paed", "Gene", refgene_dfs["paed"], "Gene", "Driver"),
        ("Sarc", "Gene", refgene_dfs["sarc"], "Gene", "Driver"),
        ("Neuro", "Gene", refgene_dfs["neuro"], "Gene", "Driver"),
        ("Ovary", "Gene", refgene_dfs["ovarian"], "Gene", "Driver"),
        ("Haem", "Gene", refgene_dfs["haem"], "Gene", "Driver"),
    )

    for (
        new_column,
        col_to_map,
        reference_df,
        col_to_index,
        col_to_look_up,
    ) in lookup_refgene:
        sv_df[new_column] = sv_df[col_to_map].map(
            reference_df.set_index(col_to_index)[col_to_look_up]
        )
        sv_df[new_column] = sv_df[new_column].fillna("-")

    sv_df.loc[:, "Variant class"] = ""
    sv_df.loc[:, "Actionability"] = ""
    sv_df.loc[:, "Comments"] = ""

    sv_df[["Type", "Copy Number"]] = sv_df.Type.str.split(
        r"\(|\)", expand=True
    ).iloc[:, [0, 1]]
    sv_df["Copy Number"] = sv_df["Copy Number"].astype(int)
    sv_df["Size"] = sv_df.apply(lambda x: "{:,.0f}".format(x["Size"]), axis=1)

    if list(sv_df["Type"].unique()) == ["GAIN"]:
        sv_df.sort_values(
            ["Event domain", "Copy Number"],
            ascending=[True, False],
            inplace=True,
        )
    else:
        sv_df.sort_values(
            ["Event domain", "Copy Number"],
            ascending=[True, True],
            inplace=True,
        )

    selected_col = [
        "Event domain",
        "Impacted transcript region",
        "Gene",
        "GRCh38 coordinates",
        "Chromosomal bands",
        "Type",
        "Copy Number",
        "Size",
        "Gene mode of action",
        "Variant class",
        "Actionability",
        "Comments",
        "COSMIC",
        "Paed",
        "Sarc",
        "Neuro",
        "Ovary",
        "Haem",
    ]

    return sv_df[selected_col]


def process_refgene(dfs: dict) -> dict:
    """Process the refgene group excel by replacing the NA by * in select
    columns

    Parameters
    ----------
    dfs : dict
        Dict of dataframes corresponding to the data in the sheets in the
        refgene group excel

    Returns
    -------
    dict
        Dict of processed dataframes
    """

    for df in [
        dfs["cosmic"],
        dfs["paed"],
        dfs["sarc"],
        dfs["neuro"],
        dfs["ovarian"],
        dfs["haem"],
    ]:
        if "Entities" in list(df.columns):
            df["Entities"].astype(str)
            df.fillna({"Entities": "*"}, inplace=True)
        if "Driver" in list(df.columns):
            df["Driver"].astype(str)
            df.fillna({"Driver": "*"}, inplace=True)

    return dfs
