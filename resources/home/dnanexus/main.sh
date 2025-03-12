set -exo pipefail

main() {
    python3 -m pip install -q --no-index --no-deps  packages/*

    dx-download-all-inputs --parallel

    python3 /home/dnanexus/generate_workbook.py \
        -hs in/hotspots/* \
        -r in/reference_gene_groups/* \
        -c in/clinvar/* \
        -i in/clinvar_index/* \
        -html in/supplementary_html/* \
        -rv in/reported_variants/* \
        -rsv in/reported_structural_variants/*

    file_id=$(dx upload "output.xlsx" --brief)
    dx-jobutil-add-output workbook $file_id
}
