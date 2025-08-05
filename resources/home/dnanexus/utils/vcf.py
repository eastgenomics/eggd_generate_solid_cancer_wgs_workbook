import pandas as pd
import vcfpy


def open_vcf(file: str) -> vcfpy.Reader:
    """Open VCF file

    Parameters
    ----------
    file : str
        File path

    Returns
    -------
    vcfpy.Reader
        Reader object for the VCF
    """

    return vcfpy.Reader.from_path(file)


def get_clinvar_info(vcf_reader: vcfpy.Reader) -> dict:
    """Parse the clinvar data

    Parameters
    ----------
    vcf_reader : vcfpy.Reader
        Vcfpy reader object containing the clinvar VCF resource data

    Returns
    -------
    dict
        Dict containing data of interest from the clinvar resource
    """

    data = {}

    for record in vcf_reader:
        assert len(record.ID) == 1, f"Multiple IDs for {record.ID}"
        record_id = str(record.ID[0])
        clnsigconf = None
        clnsig = None
        alt = None

        if record.INFO.get("CLNSIGCONF"):
            clnsigconf = record.INFO.get("CLNSIGCONF")

        if record.INFO.get("CLNSIG"):
            clnsig = record.INFO.get("CLNSIG")

        # the ALT value is a list of SUBSTITUTION object, so extraction is
        # required
        if record.ALT:
            alt = record.ALT[0].value

        data.setdefault(record_id, {})
        data[record_id]["change"] = f"{record.REF}>{alt}"
        data[record_id].setdefault("clnsigconf", []).append(clnsigconf)
        data[record_id].setdefault("clnsig", []).append(clnsig)

    return data


def find_clinvar_info(vcf_dict: dict, data: pd.DataFrame) -> pd.DataFrame:
    """Find the clinvar CLNSIGCONF at best, CLNSIG if not or returns an empty
    string for the clinvar id at worst

    Parameters
    ----------
    vcf_dict : dict
        Dict containing the clinvar data
    data : pd.DataFrame
        Dataframe with the germline variant data

    Returns
    -------
    pd.DataFrame
        Dataframe for the clinvar ids and their clinical significance
    """

    new_data = []

    for index, row in data.iterrows():
        significance = None

        # loop through the ids
        for clinvar_id in row["ClinVar ID"]:
            # check if the id is in the clinvar resource data
            if clinvar_id in vcf_dict:
                clinvar_data = vcf_dict[clinvar_id]

                # check if a significance has already been assigned
                if significance:
                    # check the nucleotide change
                    if (
                        clinvar_data["change"]
                        in row["CDS change and protein change"]
                    ):
                        significance = (
                            clinvar_data["clnsigconf"]
                            if clinvar_data["clnsigconf"]
                            else clinvar_data["clnsig"]
                        )
                else:
                    significance = (
                        clinvar_data["clnsigconf"]
                        if clinvar_data["clnsigconf"]
                        else clinvar_data["clnsig"]
                    )

        if significance:
            significance = "; ".join([ele for s in significance for ele in s])

        new_data.append(
            (
                row["Gene"],
                row["GRCh38 coordinates;ref/alt allele"],
                row["CDS change and protein change"],
                row["Predicted consequences"],
                row["Genotype"],
                row["Population germline allele frequency (GE | gnomAD)"],
                row["Gene mode of action"],
                row["ClinVar ID"],
                significance,
            )
        )

    return pd.DataFrame(
        new_data,
        columns=[
            "Gene",
            "GRCh38 coordinates;ref/alt allele",
            "CDS change and protein change",
            "Predicted consequences",
            "Genotype",
            "Population germline allele frequency (GE | gnomAD)",
            "Gene mode of action",
            "ClinVar ID",
            "clnsigconf",
        ],
    )
