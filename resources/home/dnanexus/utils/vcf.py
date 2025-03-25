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


def find_clinvar_info(vcf_file: vcfpy.Reader, *clinvar_ids) -> pd.DataFrame:
    """Find the clinvar CLNSIGCONF at best, CLNSIG if not or returns an empty
    string for the clinvar id at worst

    Parameters
    ----------
    vcf_file : vcfpy.Reader
        vcfpy.Reader object

    Returns
    -------
    pd.DataFrame
        Dataframe for the clinvar ids and their clinical significance
    """

    data = {"ClinVar ID": [], "clnsigconf": []}

    for record in vcf_file:
        for clinvar_id in clinvar_ids:
            assert len(record.ID) == 1, f"Multiple IDs for {record.ID}"
            record_id = record.ID[0]

            if record_id == clinvar_id:
                if record.INFO.get("CLNSIGCONF"):
                    clnsigconf = record.INFO.get("CLNSIGCONF")
                elif record.INFO.get("CLNSIG"):
                    clnsigconf = record.INFO.get("CLNSIG")
                else:
                    clnsigconf = [""]

                assert (
                    len(clnsigconf) == 1
                ), f"Multiple clinical significance found for {record.ID}"

                data["ClinVar ID"].append(clinvar_id)
                data["clnsigconf"].append(clnsigconf[0])

    data = pd.DataFrame(data).astype(str)
    return data
