set -exo pipefail

main() {
    python3 -m pip install -q --no-index --no-deps  packages/*

    dx-download-all-inputs --parallel

    python3 /home/dnanexus/generate_workbook.py \
        -hs in/hotspots/* \
        -r in/reference_gene_groups/* \
        -p in/panelapp/* \
        -cb in/cytological_bands/* \
        -c in/clinvar/* \
        -i in/clinvar_index/* \
        -html in/supplementary_html/* \
        -rv in/reported_variants/* \
        -rsv in/reported_structural_variants/*

    mv in/supplementary_html/* output/ 
    mv in/reported_variants/* output/
    mv in/reported_structural_variants/* output/

    python3 /home/dnanexus/final_check.py output/

    file_id=$(dx upload "output/$(ls output/*xlsx)" --brief)
    dx-jobutil-add-output workbook $file_id
}
