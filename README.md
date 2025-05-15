# eggd_generate_solid_cancer_wgs_workbook (DNAnexus Platform App)

## What does this app do?

Generates an Excel workbook from various files from Genomics England Solid cancer case.

## What data are required for this app to run?

**Packages**

* Python packages (specified in requirements.txt)

**Inputs**

* `hotspots`: CSV file containing information about the cancer hotspots. Provided by the Solid Cancer
* `reference_gene_groups`: Excel file containing information for genes and their impact on somatic variation. Provided by the Solid cancer team
* `panelapp`: Panelapp reference excel file obtained from the Solid cancer team
* `cytological_bands`: Excel file obtained from the Solid cancer team containing cytological reference data
* `clinvar`: Clinvar asset VCF file
* `clinvar_index`: Clinvar asset VCF index file
* `supplementary_html`: Supplementary HTML file from GEL
* `reported_variants`: CSV file from GEL containing info on reported variants
* `reported_structural_variants`: CSV/excel file from GEL containing info on reported structural variants

## How to run

```bash
# in dnanexus
dx run ${app_id} \
-ihotspots= \
-ireference_gene_groups= \
-ipanelapp= \
-icytological_bands= \
-iclinvar= \
-iclinvar_index= \
-isupplementary_html= \
-ireported_variants= \
-ireported_structural_variants= \
-y

# locally
python3 -m venv ${environment_name}
source ${environment_name}/bin/activate
pip install requirements.txt
python resources/home/dnanexus/generate_workbook.py \
-hs ${hotspots_file} \
-r ${reference_gene_groups} \
-c ${clinvar} \
-i ${clinvar_index} \
-html ${supplementary_html} \
-rv ${reported_variants} \
-rsv ${reported_structural_variants} \
-p ${panelapp} \
-cb ${cytological_bands} \
```

```bash
# Unittesting
source ${environment_name}/bin/activate
pytest -s --disable-warnings
```

## What does this app output?

This app outputs an Excel workbook.
