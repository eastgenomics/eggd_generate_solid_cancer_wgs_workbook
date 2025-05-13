import re

import pandas as pd
import vcfpy

from configs import tables, sv, refgene
from utils import misc, vcf


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

    return df


def process_reported_variants_germline(
    df: pd.DataFrame, clinvar_resource: vcfpy.Reader, panelapp_dfs: dict
) -> pd.DataFrame:
    """Process the data from the reported variants excel file

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe from parsing the reported variants excel file
    clinvar_resource : vcfpy.Reader
        vcfpy.Reader object from the Clinvar resource
    panelapp_dfs : dict
        Dict containing dfs to Panelapp adult and childhood data

    Returns
    -------
    pd.DataFrame
        Dataframe containing clinical significance info for germline variants
    """

    if "Origin" not in df:
        return None

    df = df[df["Origin"].str.lower() == "germline"]

    if df.empty:
        return None

    # convert the clinvar id column as a string and remove the trailing .0 that
    # the automatic conversion that pandas applies added
    df.loc[:, "ClinVar ID"] = df["ClinVar ID"].astype(str)
    df.loc[:, "ClinVar ID"] = df["ClinVar ID"].str.removesuffix(".0")

    df.reset_index(drop=True, inplace=True)

    clinvar_ids_to_find = [
        value for value in df.loc[:, "ClinVar ID"].to_numpy()
    ]
    clinvar_info = vcf.find_clinvar_info(
        clinvar_resource, *clinvar_ids_to_find
    )

    # add the clinvar info by merging the clinvar dataframe
    df = df.merge(clinvar_info, on="ClinVar ID", how="left")

    df.loc[:, "Tumour VAF"] = ""

    lookup_panelapp_data = (
        (
            "PanelApp Adult_v2.2",
            "Gene",
            panelapp_dfs["Adult v2.2"],
            "Gene Symbol",
            "Formatted mode",
        ),
        (
            "PanelApp Childhood_v4.0",
            "Gene",
            panelapp_dfs["Childhood v4.0"],
            "Gene Symbol",
            "Formatted mode",
        ),
    )

    for (
        new_column,
        mapping_column_target_df,
        reference_df,
        mapping_column_ref_df,
        col_to_look_up,
    ) in lookup_panelapp_data:
        # link the mapping column to the column of data in the ref df
        reference_dict = dict(
            zip(
                reference_df[mapping_column_ref_df],
                reference_df[col_to_look_up],
            )
        )
        df[new_column] = df[mapping_column_target_df].map(reference_dict)
        df[new_column] = df[new_column].fillna("-")

    df = df[
        [
            "Gene",
            "GRCh38 coordinates;ref/alt allele",
            "CDS change and protein change",
            "Genotype",
            "Population germline allele frequency (GE | gnomAD)",
            "Gene mode of action",
            "clnsigconf",
            "Tumour VAF",
            "PanelApp Adult_v2.2",
            "PanelApp Childhood_v4.0",
        ]
    ]

    df.fillna("", inplace=True)

    return df


def process_reported_variants_somatic(
    df: pd.DataFrame,
    lookup_refgene: tuple,
    hotspots_df: pd.DataFrame,
    cyto_df: dict,
) -> pd.DataFrame:
    """Get the somatic variants and format the data for them

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe from parsing the reported variants excel file
    lookup_refgene : tuple
        Tuple of data allowing lookup in the refgene dataframes
    hotspots_df : pd.DataFrame
        Dataframe containing data from the parsed hotspots excel
    cyto_df : dict
        Dict containing dataframe of data per sheet for cytological bands

    Returns
    -------
    pd.DataFrame
        Dataframe with additional formatting for c. and p. annotation
    """

    if "Origin" not in df:
        return None

    # select only somatic rows
    df = df[df["Origin"].str.lower().str.contains("somatic")]

    if df.empty:
        return None

    df.reset_index(drop=True, inplace=True)
    df[["c_dot", "p_dot"]] = df["CDS change and protein change"].str.split(
        r"(?=;p)", n=1, expand=True
    )
    df["p_dot"] = df["p_dot"].str.slice(1)

    df["MTBP c."] = df["Gene"] + ":" + df["c_dot"]
    df["MTBP p."] = (
        df["Gene"]
        + ":"
        + df["p_dot"]
        .apply(misc.convert_3_letter_protein_to_1)
        .str.replace("p.", "")
    )
    df.fillna({"MTBP p.": ""}, inplace=True)

    df["HS mutation lookup"] = df["MTBP p."].apply(
        lambda x: re.sub(r"[A-Z]+$", "", x)
    )

    # populate the somatic variant dataframe with data from the refgene excel
    # file
    lookup_refgene = lookup_refgene + (
        (
            "HS_Total",
            "HS mutation lookup",
            hotspots_df["HS_Samples"],
            "Gene_AA",
            "Total",
        ),
        (
            "HS_Mut",
            "HS mutation lookup",
            hotspots_df["HS_Samples"],
            "Gene_AA",
            "Mutations",
        ),
        (
            "HS_Tissue",
            "MTBP p.",
            hotspots_df["HS_Tissue"],
            "Gene_Mut",
            "Tissue",
        ),
        ("Cyto", "Gene", cyto_df["Sheet1"], "Gene", "Cyto"),
    )

    for (
        new_column,
        mapping_column_target_df,
        reference_df,
        mapping_column_ref_df,
        col_to_look_up,
    ) in lookup_refgene:
        # link the mapping column to the column of data in the ref df
        reference_dict = dict(
            zip(
                reference_df[mapping_column_ref_df],
                reference_df[col_to_look_up],
            )
        )
        df[new_column] = df[mapping_column_target_df].map(reference_dict)
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
    df.loc[:, "TSG_NMD"] = ""
    df.loc[:, "TSG_LOH"] = ""
    df.loc[:, "Splice fs?"] = ""
    df.loc[:, "SpliceAI"] = ""
    df.loc[:, "REVEL"] = ""
    df.loc[:, "OG_3' Ter"] = ""
    df.loc[:, "Recurrence somatic database"] = ""

    df = df[
        [
            "Domain",
            "Gene",
            "GRCh38 coordinates;ref/alt allele",
            "Cyto",
            "RefSeq IDs",
            "CDS change and protein change",
            "Predicted consequences",
            "Error flag",
            "Population germline allele frequency (GE | gnomAD)",
            "VAF",
            "LOH",
            "Alt allele/total read depth",
            "Gene mode of action",
            "Variant class",
            "TSG_NMD",
            "TSG_LOH",
            "Splice fs?",
            "SpliceAI",
            "REVEL",
            "OG_3' Ter",
            "Recurrence somatic database",
            "HS_Total",
            "HS_Mut",
            "HS_Tissue",
            "COSMIC Driver",
            "COSMIC Entities",
            "Paed Driver",
            "Paed Entities",
            "Sarc Driver",
            "Sarc Entities",
            "Neuro Driver",
            "Neuro Entities",
            "Ovary Driver",
            "Ovary Entities",
            "Haem Driver",
            "Haem Entities",
            "MTBP c.",
            "MTBP p.",
        ]
    ]
    df.rename(
        columns={
            "GRCh38 coordinates;ref/alt allele": "GRCh38 coordinates",
        },
        inplace=True,
    )
    df.sort_values(["Domain", "VAF"], ascending=[True, False], inplace=True)
    df = df.replace([None], [""], regex=True)
    df["VAF"] = df["VAF"].astype(float)

    return df


def process_reported_SV(
    df: pd.DataFrame, lookup_refgene: tuple, type_sv: str, *check_columns
) -> pd.DataFrame:
    """Process the reported structural variants excel

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe containing data from the structural variants excel
    lookup_refgene : tuple
        Tuple of data allowing lookup in the refgene dataframes
    type_sv: str
        Type of structural variant to look at in the function

    Returns
    -------
    pd.DataFrame
        Dataframe for variants with the given SV type
    """

    if "Type" not in df.columns:
        return None

    sv_df = df[df["Type"].str.lower().str.match(type_sv)]

    if sv_df.empty:
        return None

    sv_df.reset_index(drop=True, inplace=True)

    # populate the structural variant dataframe with data from the refgene
    # excel file
    for (
        new_column,
        mapping_column_target_df,
        reference_df,
        mapping_column_ref_df,
        col_to_look_up,
    ) in lookup_refgene:
        # link the mapping column to the column of data in the ref df
        reference_dict = dict(
            zip(
                reference_df[mapping_column_ref_df],
                reference_df[col_to_look_up],
            )
        )
        sv_df[new_column] = sv_df[mapping_column_target_df].map(reference_dict)
        sv_df[new_column] = sv_df[new_column].fillna("-")

    sv_df.loc[:, "Variant class"] = ""

    for column in check_columns:
        sv_df.loc[:, column] = ""

    sv_df[["Type", "Copy Number"]] = sv_df.Type.str.split(
        r"\(|\)", expand=True
    ).iloc[:, [0, 1]]
    sv_df["Copy Number"] = sv_df["Copy Number"].astype(int)
    sv_df["Size"] = sv_df.apply(lambda x: "{:,.0f}".format(x["Size"]), axis=1)
    sv_df[["Cyto 1", "Cyto 2"]] = sv_df["Chromosomal bands"].str.split(
        ";", expand=True
    )

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

    selected_col = (
        [
            "Event domain",
            "Gene",
            "RefSeq IDs",
            "Impacted transcript region",
            "GRCh38 coordinates",
            "Type",
            "Copy Number",
            "Size",
            "Cyto 1",
            "Cyto 2",
            "Gene mode of action",
            "Variant class",
        ]
        + [column for column in check_columns]
        + [
            "COSMIC Driver",
            "COSMIC Entities",
            "Paed Driver",
            "Paed Entities",
            "Sarc Driver",
            "Sarc Entities",
            "Neuro Driver",
            "Neuro Entities",
            "Ovary Driver",
            "Ovary Entities",
            "Haem Driver",
            "Haem Entities",
        ]
    )

    return sv_df[selected_col]


def process_fusion_SV(
    df: pd.DataFrame, lookup_refgene: tuple, cyto_df: dict
) -> pd.DataFrame:
    """Process the fusions from the structural variants excel

    Parameters
    ----------
    df : pd.DataFrame
        Dataframe containing the data from the structural variant excel
    lookup_refgene : tuple
        Tuple of data allowing lookup in the refgene dataframes
    cyto_df : dict
        Dict containing dataframe of data per sheet for cytological bands

    Returns
    -------
    tuple
        - Dataframe containing data for the fusion structural variants
        - Max number of fusion
    """

    if "Type" not in df.columns:
        return None

    df_SV = df[~df["Type"].str.lower().str.contains("loss|loh|gain")]

    if df_SV.empty:
        return None, 0

    df_SV.reset_index(drop=True, inplace=True)

    # split fusion columns
    df_SV["fusion_count"] = df_SV["Type"].str.count(r"\;")
    fusion_count = df_SV["fusion_count"].max()

    fusion_col = ["Type"]

    for i in range(fusion_count):
        fusion_col.append(f"Fusion_{i+1}")

    # create intermediate dataframe to concatenate the fusion information with
    # the main dataframe
    inter_df = pd.DataFrame({}, columns=fusion_col)
    inter_df[fusion_col] = df_SV.Type.str.split(";", expand=True)
    df_SV.drop("Type", inplace=True, axis=1)
    df_SV = pd.concat([df_SV, inter_df], axis=1)

    # remove prefixes for single reads and paired reads and store in separate
    # columns
    df_SV[["Paired reads", "Split reads"]] = (
        df_SV["Confidence/support"]
        .apply(misc.split_confidence_support)
        .to_list()
    )

    # get thousands separator
    df_SV["Size"] = df_SV.apply(lambda x: "{:,.0f}".format(x["Size"]), axis=1)

    # replace nan in size with empty string
    df_SV.fillna({"Size": ""}, inplace=True)
    df_SV.replace({"Size": "nan"}, {"Size": ""}, inplace=True)

    # get gene counts and look up for each gene
    max_num_gene = df_SV["Gene"].str.count(r"\;").max() + 1

    gene_col = []

    for i in range(max_num_gene):
        gene_col.append(f"Gene_{i+1}")

    df_SV[gene_col] = df_SV["Gene"].str.split(";", expand=True)

    lookup_cols = []
    cyto_cols = []

    lookup_refgene = lookup_refgene + (
        ("Cyto", "Gene", cyto_df["Sheet1"], "Gene", "Cyto"),
    )

    for (
        new_column,
        mapping_column_target_df,
        reference_df,
        mapping_column_ref_df,
        col_to_look_up,
    ) in lookup_refgene:
        for gene_col_name in gene_col:
            column_to_write = f"{gene_col_name} | {new_column}"
            mapping_column_target_df = gene_col_name
            # link the mapping column to the column of data in the ref df
            reference_dict = dict(
                zip(
                    reference_df[mapping_column_ref_df],
                    reference_df[col_to_look_up],
                )
            )
            df_SV[column_to_write] = df_SV[mapping_column_target_df].map(
                reference_dict
            )
            df_SV.fillna({column_to_write: "-"}, inplace=True)

            # store the cyto columns apart from the other lookup groups to
            # reorder
            if "Cyto" in new_column:
                cyto_cols.append(column_to_write)
            else:
                lookup_cols.append(column_to_write)

    df_SV.loc[:, "Variant class"] = ""
    df_SV.loc[:, "OG_Fusion"] = ""
    df_SV.loc[:, "OG_IntDup"] = ""
    df_SV.loc[:, "OG_IntDel"] = ""
    df_SV.loc[:, "Disruptive"] = ""

    expected_columns = sv.CONFIG["expected_columns"]
    alternatives = sv.CONFIG["alternative_columns"]

    alternative_columns = tables.find_alternative_headers(
        df_SV, expected_columns, alternatives
    )

    subset_column = [
        (
            column
            if column not in alternative_columns
            else alternative_columns[column]
        )
        for column in expected_columns
    ]

    if cyto_cols:
        for col in cyto_cols[::-1]:
            subset_column.insert(10, col)

    for col in fusion_col[::-1]:
        subset_column.insert(6, col)

    selected_col = subset_column + lookup_cols

    return df_SV[selected_col], fusion_count


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

    output_dataframe = None

    for sheet_name in refgene.SHEETS2COLUMNS:
        if sheet_name in dfs:
            df = dfs[sheet_name]
            df.rename(columns=refgene.SHEETS2COLUMNS[sheet_name], inplace=True)
            df = df[list(refgene.SHEETS2COLUMNS[sheet_name].values())]
        else:
            found_alternative = False

            for key, alternatives in refgene.RESCUE_COLUMNS.items():
                if key == sheet_name:
                    for alternative in alternatives:
                        if alternative in dfs:
                            df = dfs[alternative]
                            df.rename(
                                columns=refgene.SHEETS2COLUMNS[sheet_name],
                                inplace=True,
                            )
                            df = df[
                                list(
                                    refgene.SHEETS2COLUMNS[sheet_name].values()
                                )
                            ]
                            found_alternative = True
                            break

            assert (
                found_alternative
            ), f"Couldn't find an alternative to sheet name: {sheet_name}"

        # in the first loop just assign the output dataframe to the processed
        # df
        if output_dataframe is None:
            output_dataframe = df
        else:
            # on the other loops, merge dataframes using the Gene column
            output_dataframe = output_dataframe.merge(
                df, how="outer", on="Gene"
            )

    return output_dataframe


def process_panelapp(dfs: dict) -> dict:
    """Process and format panelapp reference dataframes for use in the
    Germline sheet

    Parameters
    ----------
    dfs : dict
        Dict of sheets of the panelapp reference file and the corresponding
        dataframes parsed from the sheets

    Returns
    -------
    dict
        Dict of sheets with the corresponding reformatted dataframes
    """

    data = {}

    for type_df, df in dfs.items():
        df.fillna({"Gene Symbol": ""}, inplace=True)
        df.fillna({"Mode": ""}, inplace=True)
        df.fillna({"Phenotypes": ""}, inplace=True)
        df["Mode"] = df["Mode"].astype(str)
        df["Phenotypes"] = df["Phenotypes"].astype(str)

        df["Formatted mode"] = (
            df["Mode"] + " " + df["Phenotypes"].apply(lambda x: f"[{x}]")
        )
        df = df[["Gene Symbol", "Formatted mode"]]
        data[type_df] = df

    return data


def lookup_data_from_variants(
    refgene_df: pd.DataFrame, **kwargs
) -> pd.DataFrame:
    """Lookup data from other variant dataframes and add it to the refgene df

    Parameters
    ----------
    refgene_df : pd.DataFrame
        Dataframe containing the refgene data

    Returns
    -------
    pd.DataFrame
        Refgene data dataframe with data from variant dataframes
    """

    lookup_variant_data = (
        (
            "SNV",
            "Gene",
            kwargs["somatic"],
            "Gene",
            "CDS change and protein change",
        ),
        ("CN", "Gene", kwargs["gain"], "Gene", "Copy Number"),
    )

    for (
        new_column,
        mapping_column_target_df,
        reference_df,
        mapping_column_ref_df,
        col_to_look_up,
    ) in lookup_variant_data:
        # link the mapping column to the column of data in the ref df
        reference_dict = dict(
            zip(
                reference_df[mapping_column_ref_df],
                reference_df[col_to_look_up],
            )
        )
        refgene_df[new_column] = refgene_df[mapping_column_target_df].map(
            reference_dict
        )
        refgene_df[new_column] = refgene_df[new_column].fillna("-")

    df_fusion = kwargs["fusion"]

    gene_col = []

    for i in range(df_fusion["Gene"].str.count(r"\;").max() + 1):
        gene_col.append(f"Gene_{i+1}")

    df_fusion[gene_col] = df_fusion["Gene"].str.split(";", expand=True)

    # dynamic number of columns to be generated out of fusion partners
    for (
        new_column,
        mapping_column_target_df,
        reference_df,
        col_to_look_up,
    ) in [("SV_{}", "Gene", df_fusion, "Type")]:
        for gene in gene_col:
            column_to_write = new_column.format(gene.lower())
            # link the mapping column to the column of data in the ref df
            reference_dict = dict(
                zip(
                    reference_df[gene],
                    reference_df[col_to_look_up],
                )
            )
            refgene_df[column_to_write] = refgene_df[
                mapping_column_target_df
            ].map(reference_dict)
            refgene_df[column_to_write] = refgene_df[column_to_write].fillna(
                "-"
            )

    return refgene_df
