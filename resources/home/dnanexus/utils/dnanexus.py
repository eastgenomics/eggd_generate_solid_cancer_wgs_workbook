import os

import dxpy


def get_refgene_input_file_info():
    """Get info from the refgene input file

    Returns
    -------
    str
        String with metadata info for the refgene file
    """

    job_id = os.environ.get("DX_JOB_ID", None)

    if job_id:
        job = dxpy.DXJob(job_id)
        refgene_file = dxpy.DXFile(
            job.describe()["input"]["reference_gene_groups"]["$dnanexus_link"]
        )
        refgene_info = f"{refgene_file.name} - {refgene_file.id}"
    else:
        refgene_info = "File not retrievable"

    return refgene_info


def get_app_version():
    """Get the version of the app used to create the workbook

    Returns
    -------
    str
        String of the version
    """

    job_id = os.environ.get("DX_JOB_ID", None)

    if job_id:
        job = dxpy.DXJob(job_id)
        app_version = dxpy.DXApp(job.describe()["executable"]).version
    else:
        app_version = "App not retrievable"

    return app_version
