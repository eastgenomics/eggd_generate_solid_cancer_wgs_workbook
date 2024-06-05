# eggd_generate_wgs_solid_cancer_workbook
This is nexflow applet to generate the workbook for WGS solid cancer variants processed by GEL

### Inputs
- variants csv file
- structural variants csv file
- html file
- b38 clinvar and index files
- RefGene_Groups csv file
- Hotspots csv file
- cancer group

### Outputs
- excel spreadsheet for variants

### How to develop applet on DNAnexus
- `git clone` the repo
- `cd` into `eggd_generate_wgs_solid_cancer_workbook`
- `dx build --nextflow` (Recommend to use dxpy version 0.376)

### How to run the applet
Example command
```
dx run applet-GkGyvGj4Yj3135yV5jqVJgbq \
-ihotspots="project-GkG4Zf84Yj359Q9JYbbqbFpy:file-GkG4q6j4Yj36y2VXZzqB09J5" \
-irefgenegp="project-GkG4Zf84Yj359Q9JYbbqbFpy:file-GkG4q784Yj3J3pjKpq3110gJ" \
-iclinvar="project-GkG4Zf84Yj359Q9JYbbqbFpy:file-GkG4q484Yj300qjXX1vg8f6J" \
-iclinvar_index="project-GkG4Zf84Yj359Q9JYbbqbFpy:file-GkG4q6Q4Yj354zfXJBXg408z" \
-icancer_gp="xxxx" \
-inextflow_pipeline_params="--file_path=dx://project-xxx:/xx/xxx" # file path where csv and html are located
```

:triangular_flag_on_post: DNAnexus told me that the next released dxpy version will allow to specify the input URI with file ID in the config file. So, recommended to update the nextflow.config file with fileID URL once the new dxpy version is released so that it is not require to to specify them in the command line :triangular_flag_on_post:
