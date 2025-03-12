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
