nextflow.enable.dsl=2

// process to run python script to generate the wgs solid cancer workbook
process wgs_solidCA_workbook {

   debug true
   publishDir params.outdir, mode:'copy'
   tag "${reads[0]}, ${reads[1]}, ${reads[2]}"
    
   input:
      path(reads)
      path hotspots
      path refgenegp
      path clinvar
      path clinvar_index 
        
   output:
      path "*.xlsx"

   script:
      """
      echo "Running ${reads[0]} ${reads[1]} ${reads[2]}"
      python nextflow-bin/create_spreadsheet.py -v ${reads[0]} -sv ${reads[1]} -html ${reads[2]} -hs $hotspots -rgg $refgenegp -c $clinvar 
      """
}

workflow {
    // create channels for input files
    v_ch = Channel.fromPath(params.variant) 
    sv_ch = Channel.fromPath(params.structural_variant)
    html_ch = Channel.fromPath(params.html)

    // split the file paths to map with same sample
   v_ch
      .map { [it.toString().split("-")[0],
         it] }
      .set { v1_ch }
   
   sv_ch
      .map { [it.toString().split("-")[0],
         it] }
      .set { sv1_ch }
   
   html_ch
      .map { [it.toString().split("-")[0],
         it] }
      .set { html1_ch }   

// run the process   
wgs_solidCA_workbook(v1_ch
  .combine(sv1_ch, by: 0)
  .combine(html1_ch, by: 0)    
  .map { id, v, sv,html -> [v, sv,html] }, params.hotspots, params.refgenegp, params.clinvar, params.clinvar_index)    
}

