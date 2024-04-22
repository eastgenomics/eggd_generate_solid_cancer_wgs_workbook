import argparse
import openpyxl
from openpyxl import load_workbook, drawing
from openpyxl.styles import Alignment, Border, DEFAULT_FONT, Font, Side
from openpyxl.styles.fills import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import urllib.request
from bs4 import BeautifulSoup


# openpyxl style settings
THIN = Side(border_style="thin", color="000000")
MEDIUM = Side(border_style="medium", color="000001")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)

DEFAULT_FONT.name = "Calibri"

cyto_ref = "./resource/CytoRef.txt"
fusion_ref = "./resource/fusions.txt"
hotspots_ref = "./resource/Hotspots.txt"
refgene_ref = "./resource/Ref_Gene.txt"
refgenegp_ref = "./resource/RefGene_groups.txt"

class excel:
    """
    Functions for wrangling input csv files and html files and 
    writing output xlsm file

    Attributes
    ----------
    args : argparse.Namespace
        arguments passed from command line
    vcfs : list of pd.DataFrame
        list of dataframes formatted to write to file from vcf() methods
    additional_files : dict
        (optional) if addition files have been passed, dict will be populated
        with worksheet name : df of file data
    refs : list
        list of reference names parsed from vcf headers
    writer : pandas.io.excel._openpyxl.OpenpyxlWriter
        writer object for writing Excel data to file
    workbook : openpyxl.workbook.workbook.Workbook
        openpyxl workbook object for interacting with per-sheet writing and
        formatting of output Excel file

    Outputs
    -------
    {args.output}.xlsm : file
        Excel file with variants and structural variants with ref sheets
    """

    def __init__(self) -> None:
        self.args = self.parse_args()
        self.writer = pd.ExcelWriter(self.args.output, engine="openpyxl")
        self.workbook = self.writer.book

    def parse_args(self) -> argparse.Namespace:
        """
        Parse command line arguments

        Returns
        -------
        args : Namespace
            Namespace of passed command line argument inputs
        """
        parser = argparse.ArgumentParser()

        parser.add_argument("--output", "--o", help="output xlsm file name")
        parser.add_argument("--html", help="html input")
        parser.add_argument("--variant", "--v", help="variant csv file")
        parser.add_argument("--SV", "--sv", help="structural variant csv file")

        return parser.parse_args()

    def generate(self) -> None:
        """
        Calls all methods in excel() to generate output file
        """
        self.download_html_img()
        self.write_sheets()
        self.workbook.save(self.args.output)
        wb = load_workbook(self.args.output, keep_vba=True)
        wb.save(self.args.output)

        print("Done!")

    def download_html_img(self) -> None:
        """
        get the image links from html input file
        """
        url = self.args.html
        # "file:///home/winmintun/Desktop/workspace/solid_cancer_GEL/2024_04_02_1042440297_p28131153889_LP5101091-DNA_D05_LP5101090-DNA_B05.v3_6_2.supplementary.html"

        f = urllib.request.urlopen(url)
        page = f.read()
        f.close()
        soup = BeautifulSoup(page, features="lxml")
        n = 1
        for link in soup.findAll("img"):
            img_link = link.get("src")
            self.download_image(img_link, "./", "testing_" + str(n))
            n = n + 1

    def download_image(self, url, file_path, file_name):
        """
        donwnload the img from html file
        """
        full_path = file_path + file_name + ".jpg"
        urllib.request.urlretrieve(url, full_path)

    def read_html_tables(self, table_num) -> dict:
        """
        get the tables from html file

        Parameters
        ----------
        table number

        Returns
        -------
        dict of each html table
        """
        url = self.args.html
        f = urllib.request.urlopen(url)
        page = f.read()
        f.close()
        soup = BeautifulSoup(page, features="lxml")
        tables = soup.findAll("table")
        info = tables[table_num]
        headings = [th.get_text() for th in info.find("tr").find_all("th")]
        for row in info.find_all("tr")[1:]:
            dataset = dict(
                zip(headings, (td.get_text() for td in row.find_all("td")))
            )

        return dataset

    def write_sheets(self) -> None:
        """
        Write summary sheet to xlsm file
        """
        print("Writing sheets")

        self.soc = self.workbook.create_sheet("SOC")
        self.create_soc()
        self.QC = self.workbook.create_sheet("QC")
        self.create_QC()
        self.plot = self.workbook.create_sheet("Plot")
        self.create_plot()
        self.signatures = self.workbook.create_sheet("Signatures")
        self.create_signatures()
        self.germline = self.workbook.create_sheet("Germline")
        self.create_germline()
        self.summary = self.workbook.create_sheet("Summary")
        self.create_summary()
        self.write_fusion()
        self.write_refgene()
        self.write_refgene_groups()
        self.write_cytoref()
        self.write_hotspots()
        self.write_SNV()
        self.write_SV()

    def create_soc(self) -> None:
        """
        Write soc sheet
        """
        patient_info = self.read_html_tables(0)
        # write titles for summary values
        self.soc.cell(1, 1).value = "Patient Details (Epic demographics)"
        self.soc.cell(1, 3).value = "Previous testing"
        self.soc.cell(2, 1).value = "Name"
        self.soc.cell(2, 3).value = "Alteration"
        self.soc.cell(2, 4).value = "Assay"
        self.soc.cell(2, 5).value = "Result"
        self.soc.cell(2, 6).value = "WGS concordance"
        self.soc.cell(3, 1).value = patient_info["Gender"]
        self.soc.cell(4, 1).value = patient_info["Patient ID"]
        self.soc.cell(5, 1).value = "MRN"
        self.soc.cell(6, 1).value = "NHS Number"
        self.soc.cell(8, 1).value = "Histology"
        self.soc.cell(12, 1).value = "Comments"
        self.soc.cell(16, 1).value = "WGS in-house gene panel applied"
        self.soc.cell(17, 1).value = "COSMIC_Cancer_Genes"

        # merge some title columns that have longer text
        self.soc.merge_cells(
            start_row=1, end_row=1, start_column=3, end_column=6
        )

        cell_to_align = ["C1", "C2", "D2", "E2", "F2"]
        # make the coverage tile centre of merged rows
        for cell in cell_to_align:
            self.soc[cell].alignment = Alignment(
                wrapText=True, horizontal="center"
            )
            if cell != "C1":
                self.soc[cell].font = Font(italic=True)

        # titles to set to bold
        to_bold = ["A1", "A8", "A12", "A16", "C1"]

        for cell in to_bold:
            self.soc[cell].font = Font(bold=True, name=DEFAULT_FONT.name)

        # set column widths for readability
        self.soc.column_dimensions["A"].width = 32
        self.soc.column_dimensions["C"].width = 16
        self.soc.column_dimensions["D"].width = 26
        self.soc.column_dimensions["F"].width = 26

        # colour title cells
        blueFill = PatternFill(patternType="solid", start_color="90EE90")

        colour_cells = ["C3", "D3", "E3", "F3", "C4", "D4", "E4", "F4"]
        for cell in colour_cells:
            self.soc[cell].fill = blueFill

        # set borders around table areas
        row_ranges = [
            "C1:F1", "C2:F2", "C3:F3", "C4:F4",
            "C5:F5", "C6:F6", "C7:F7", "C8:F8",
        ]
        for row in row_ranges:
            for cells in self.soc[row]:
                for cell in cells:
                    cell.border = THIN_BORDER

        cell_lower_border = ["A1", "A8", "A12", "A16"]
        for cell in cell_lower_border:
            self.soc[cell].border = LOWER_BORDER

        # add dropdowns
        cells_for_concordance = []
        for i in range(3, 17):
            cells_for_concordance.append(f"F{i}")
        concordance_options = '"Novel,Concordant (detected),Concordant \
                              (undetected),Disconcordant (detected) \
                              ,Disconcordant (undetected),N/A"'
        self.get_drop_down(
            dropdown_options=concordance_options,
            prompt="Select from the list",
            title="WGS concordance",
            sheet=self.soc,
            cells=cells_for_concordance,
        )

        cells_for_result = []
        for i in range(3, 17):
            cells_for_result.append(f"E{i}")
        result_options = '"Detected, Not detected"'
        self.get_drop_down(
            dropdown_options=result_options,
            prompt="Select from the list",
            title="Result",
            sheet=self.soc,
            cells=cells_for_result,
        )

        cells_for_assay = []
        for i in range(3, 17):
            cells_for_assay.append(f"D{i}")
        assay_options = '"FISH,IHC,NGS,Sanger,NGS multi-gene panel, \
                        RNA fusion panel,SNP array, Methylation array, \
                        MALDI-TOF, MLPA, MS-MLPA, Chromosome breakage, \
                        Digital droplet PCR, RT-PCR, LR-PCR"'
        self.get_drop_down(
            dropdown_options=assay_options,
            prompt="Select from the list",
            title="Assay",
            sheet=self.soc,
            cells=cells_for_assay,
        )

    def create_QC(self) -> None:
        """
        write QC sheet
        """
        tumor_info = self.read_html_tables(1)
        sample_info = self.read_html_tables(2)
        germlline_info = self.read_html_tables(3)
        self.QC.cell(1, 1).value = "=SOC!A2"
        self.QC.cell(1, 3).value = "Diagnosis Date"
        self.QC.cell(1, 4).value = "Tumour Received"
        self.QC.cell(1, 5).value = "Tumour ID"
        self.QC.cell(1, 6).value = "Presentation"
        self.QC.cell(1, 7).value = "Diagnosis"
        self.QC.cell(1, 8).value = "Tumour Site"
        self.QC.cell(1, 9).value = "Tumour Type"
        self.QC.cell(1, 10).value = "Germline Sample"
        self.QC.cell(2, 1).value = "=SOC!A3"
        self.QC.cell(3, 1).value = "=SOC!A5"
        self.QC.cell(4, 1).value = "=SOC!A6"
        self.QC.cell(4, 3).value = "Purity (Histo)"
        self.QC.cell(4, 4).value = "Purity (Calculated)"
        self.QC.cell(4, 5).value = "Ploidy"
        self.QC.cell(4, 6).value = "Total SNVs"
        self.QC.cell(4, 7).value = "Total Indels"
        self.QC.cell(4, 8).value = "Total SVs"
        self.QC.cell(4, 9).value = "TMB"
        self.QC.cell(6, 1).value = "=SOC!A9"
        self.QC.cell(7, 3).value = "Sample type"
        self.QC.cell(7, 4).value = "Mean depth, x"
        self.QC.cell(7, 5).value = "Mapped reads, %"
        self.QC.cell(7, 6).value = "Chimeric DNA frag, %"
        self.QC.cell(7, 7).value = "Insert size, bp"
        self.QC.cell(7, 8).value = "Unevenness, x"
        self.QC.cell(8, 1).value = "QC alerts"
        self.QC.cell(9, 1).value = "None"  # need dropdown
        self.QC.cell(2, 3).value = tumor_info["Tumour Diagnosis Date"]
        self.QC.cell(2, 5).value = tumor_info[
            "Histopathology or SIHMDS LAB ID"
        ]
        self.QC.cell(2, 6).value = tumor_info["Presentation"]
        # self.QC.cell(7, 2).value = tumor_info['Primary or Metastatic']
        self.QC.cell(2, 8).value = tumor_info["Primary or Metastatic"]
        self.QC.cell(2, 9).value = tumor_info["Tumour Type"]
        self.QC.cell(2, 10).value = germlline_info["Sample ID"]

        # titles to set to bold
        to_bold = [
            "A1", "A8", "C1", "D1", "E1", "F1", "G1",
            "H1", "I1", "J1", "C4", "D4", "E4", "F4",
            "G4", "H4", "I4", "C7", "D7", "E7", "F7",
            "G7", "H7",
        ]

        for cell in to_bold:
            self.QC[cell].font = Font(bold=True, name=DEFAULT_FONT.name)

        # set column widths for readability
        self.QC.column_dimensions["A"].width = 32
        self.QC.column_dimensions["B"].width = 8
        for col in ["C", "D", "E", "F", "G", "H", "I", "J"]:
            self.QC.column_dimensions[col].width = 22

        # set borders around table areas
        row_ranges = [
            "C1:J1", "C2:J2", "C4:I4", "C5:I5",
            "C7:H7", "C8:H8", "C9:H9"
        ]
        for row in row_ranges:
            for cells in self.QC[row]:
                for cell in cells:
                    cell.border = THIN_BORDER
        self.QC["A8"].border = LOWER_BORDER

        # add dropdowns
        cells_for_QC = ["A9"]
        QC_options = '"None,<30% tumour purity,SNVs low VAF (<6%),TINC (<5%)"'
        self.get_drop_down(
            dropdown_options=QC_options,
            prompt="Select from the list",
            title="QC alerts",
            sheet=self.QC,
            cells=cells_for_QC,
        )

        # insert img from html
        self.insert_img(self.QC, "testing_9.jpg", "C12", 400, 600)

    def create_plot(self) -> None:
        """
        write plot sheet
        """
        self.plot.cell(1, 1).value = "=SOC!A2"
        self.plot.cell(2, 1).value = "=SOC!A3"
        self.plot.cell(3, 1).value = "=SOC!A5"
        self.plot.cell(4, 1).value = "=SOC!A6"
        self.plot.cell(6, 1).value = "=SOC!A9"
        self.plot.cell(8, 1).value = "Pertinent chromosomal CNVs"
        self.plot.cell(9, 1).value = "None"  
        self.plot.cell(5, 4).value = "Insert"

        # titles to set to bold
        to_bold = ["A1", "A8"]

        for cell in to_bold:
            self.plot[cell].font = Font(bold=True, name=DEFAULT_FONT.name)

        # set column widths for readability
        self.plot.column_dimensions["A"].width = 32

    def create_signatures(self) -> None:
        """
        write signatures sheet
        """
        self.signatures.cell(1, 1).value = "=SOC!A2"
        self.signatures.cell(2, 1).value = "=SOC!A3"
        self.signatures.cell(3, 1).value = "=SOC!A5"
        self.signatures.cell(4, 1).value = "=SOC!A6"
        self.signatures.cell(6, 1).value = "=SOC!A9"
        self.signatures.cell(8, 1).value = "Signature version"
        self.signatures.cell(9, 1).value = "v2 (March 2015)"
        self.signatures.cell(13, 1).value = "Pertinent signatures"
        self.signatures.cell(14, 1).value = "None"
        self.signatures.cell(14, 5).value = "Insert"

        # titles to set to bold
        to_bold = ["A1", "A8", "A13"]

        for cell in to_bold:
            self.signatures[cell].font = Font(
                bold=True, name=DEFAULT_FONT.name
            )

        self.signatures["A8"].border = LOWER_BORDER

        # set column widths for readability
        self.signatures.column_dimensions["A"].width = 32
        self.insert_img(self.signatures, "testing_6.jpg", "C3", 700, 1000)
        self.insert_img(self.signatures, "testing_7.jpg", "P3", 400, 600)

    def create_germline(self) -> None:
        """
        write germline sheet
        """
        self.germline.cell(1, 1).value = "=SOC!A2"
        self.germline.cell(2, 1).value = "=SOC!A3"
        self.germline.cell(3, 1).value = "=SOC!A5"
        self.germline.cell(4, 1).value = "=SOC!A6"
        self.germline.cell(6, 1).value = "=SOC!A9"
        self.germline.cell(8, 1).value = "Pertinent germline variants"
        self.germline.cell(9, 1).value = "None"

        # Germline SNV table
        self.germline.cell(1, 3).value = "Germline SNV"
        self.germline.cell(2, 3).value = "Gene"
        self.germline.cell(2, 4).value = "GRCh38 Coordinates"
        self.germline.cell(2, 5).value = "Variant"
        self.germline.cell(2, 6).value = "Genotype"
        self.germline.cell(2, 7).value = "Role in Cacer"
        self.germline.cell(2, 8).value = "ClinVar"
        self.germline.cell(2, 9).value = "gnomAD"
        self.germline.cell(2, 10).value = "Tumour VAF"
        for row in range(3, 12):
            ref_row = row + 22
            for col in ["C", "D", "E", "F", "G", "H"]:
                self.germline[f"{col}{row}"] = f"=germline!{col}{ref_row}"

        self.germline.cell(13, 3).value = "Clinical genetics feedback"
        self.germline.cell(24, 1).value = "Domain"
        self.germline.cell(24, 2).value = "Origin"
        self.germline.cell(24, 3).value = "Gene"
        self.germline.cell(24, 4).value = "GRCh38 coordinates;ref/alt allele"
        self.germline.cell(24, 5).value = "Transcript"
        self.germline.cell(24, 6).value = "CDS change and protein change"
        self.germline.cell(24, 7).value = "Predicted consequences"
        self.germline.cell(
            24, 8
        ).value = "Population germline allele frequency (GE | gnomAD)"
        self.germline.cell(24, 9).value = "VAF"
        self.germline.cell(24, 10).value = "Alt allele/total read depth"
        self.germline.cell(24, 11).value = "Genotype"
        self.germline.cell(24, 12).value = "COSMIC ID"
        self.germline.cell(24, 13).value = "ClinVar ID"
        self.germline.cell(24, 14).value = "ClinVar review status"
        self.germline.cell(24, 15).value = "ClinVar clinical significance"
        self.germline.cell(24, 16).value = "Gene mode of action"
        self.germline.cell(
            24, 17
        ).value = "Recruiting Clinical Trials 30 Jan 2023"
        self.germline.cell(24, 18).value = "PharmGKB_ID"

        # ACMG classification table
        self.germline.cell(40, 1).value = "ACMG classification table"
        self.germline.cell(41, 1).value = "Theme"
        self.germline.cell(42, 1).value = "Population"
        self.germline.cell(46, 1).value = "Computational"
        self.germline.cell(55, 1).value = "Functional"
        self.germline.cell(59, 1).value = "Segregation"
        self.germline.cell(61, 1).value = "De novo"
        self.germline.cell(63, 1).value = "Allelic"
        self.germline.cell(66, 1).value = "Other"
        self.germline.cell(69, 1).value = "Guidelines:"
        self.germline.cell(70, 1).value = "E-mail:"
        self.germline.cell(71, 1).value = "Confirmations:"
        self.germline.cell(41, 2).value = "Epic"
        self.germline.cell(
            41, 3
        ).value = "Description (CanVIG-UK guidelines, v2.17)"
        self.germline.cell(41, 4).value = "Criteria"
        self.germline.cell(41, 5).value = "Strength"
        self.germline.cell(41, 6).value = "Points"
        self.germline.cell(41, 7).value = "Prelim"
        self.germline.cell(41, 8).value = "Check"
        self.germline.cell(41, 9).value = "Comments"
        self.germline.cell(68, 4).value = "Classification"
        self.germline.cell(
            69, 3
        ).hyperlink = "https://www.cangene-canvaruk.org/canvig-uk "
        self.germline.cell(
            70, 3
        ).hyperlink = "cuh.eastglh-rarediseases-inheritedcancer@nhs.net"
        self.germline.cell(
            71, 3
        ).value = "TSO pan-cancer add-on (via specimen update)"

        self.germline.cell(42, 3).value = "4sxb"
        self.germline.cell(
            43, 3
        ).value = "Absent from controls (gnomADv2.1.1 non-cancer)"
        self.germline.cell(
            44, 3
        ).value = "Allele frequency is higher than expected for disorder (gnomADv2.1.1 non-cancer)"
        self.germline.cell(
            45, 3
        ).value = "Allele frequency is >5% (gnomADv2.1.1 non-cancer)"
        self.germline.cell(
            46, 3
        ).value = (
            "Null variant in a gene where LOF is a known mechanism of disease"
        )
        self.germline.cell(
            47, 3
        ).value = "Protein length changes as a result of in-frame deletions/ insertions in a non-repeat region or stop-loss variants"
        self.germline.cell(
            48, 3
        ).value = "Same amino acid change as a previously established pathogenic variant, regardless of nucleotide change"
        self.germline.cell(
            49, 3
        ).value = "Missense change at an amino acid residue where a different missense change determined to be pathogenic has been seen before"
        self.germline.cell(
            50, 3
        ).value = "Multiple lines of computational evidence support a deleterious effect on the gene or gene product (conservation, evolutionary, splicing impact)"
        self.germline.cell(
            51, 3
        ).value = "Multiple lines of computational evidence suggest no impact on gene or gene product (conservation, evolutionary, splicing impact, etc.)"
        self.germline.cell(
            52, 3
        ).value = "Missense variant in a gene for which primarily truncating variants are known to cause disease"
        self.germline.cell(
            53, 3
        ).value = "In-frame deletions in a repetitive region without a known function"
        self.germline.cell(
            54, 3
        ).value = "A synonymous (silent) variant for which splicing prediction algorithms predict no impact to the splice consensus sequence nor the creation of a new splice site AND the nucleotide is not highly conserved"
        self.germline.cell(
            55, 3
        ).value = "Well-established in vitro or in vivo functional studies supportive of a damaging effect on the gene or gene product"
        self.germline.cell(
            56, 3
        ).value = "Located in a mutational hot spot and/or critical and well-established functional domain (e.g. active site of an enzyme) without benign variation"
        self.germline.cell(
            57, 3
        ).value = "Missense variant in a gene that has a low rate of benign missense variation and in which missense variants are a common mechanism of disease"
        self.germline.cell(
            58, 3
        ).value = "Well-established in vitro or in vivo functional studies show no damaging effect on protein function or splicing "
        self.germline.cell(
            59, 3
        ).value = "Co-segregation with disease in multiple affected family members in a gene definitively known to cause the disease"
        self.germline.cell(60, 3).value = "Non segregation with disease"
        self.germline.cell(
            61, 3
        ).value = "De novo (both maternity and paternity confirmed) in a patient with the disease and no family history"
        self.germline.cell(
            62, 3
        ).value = "Assumed de novo, but without confirmation of paternity and maternity"
        self.germline.cell(
            63, 3
        ).value = "For recessive disorders, detected in trans with a pathogenic variant"
        self.germline.cell(
            64, 3
        ).value = "Observed in trans with a pathogenic variant for a fully penetrant dominant gene/disorder or observed in cis with a pathogenic variant in any inheritance pattern"
        self.germline.cell(
            65, 3
        ).value = "Observation in controls inconsistent with disease penetrance. Observed in a healthy adult individual for a recessive (homozygous), dominant (heterozygous), or Xlinked (hemizygous) disorder, with full penetrance expected at an early age"
        self.germline.cell(
            66, 3
        ).value = "Patient’s phenotype or family history is highly specific for a disease with a single genetic aetiology"
        self.germline.cell(
            67, 3
        ).value = "Variant found in case with an alternate molecular basis"
        self.germline.cell(42, 4).value = "PS4"
        self.germline.cell(43, 4).value = "PM2"
        self.germline.cell(44, 4).value = "BS1"
        self.germline.cell(45, 4).value = "BA1"
        self.germline.cell(46, 4).value = "PVS1"
        self.germline.cell(47, 4).value = "PM4"
        self.germline.cell(48, 4).value = "PS1"
        self.germline.cell(49, 4).value = "PM5"
        self.germline.cell(50, 4).value = "PP3"
        self.germline.cell(51, 4).value = "BP4:"
        self.germline.cell(52, 4).value = "BP1"
        self.germline.cell(53, 4).value = "BP3"
        self.germline.cell(54, 4).value = "BP7"
        self.germline.cell(55, 4).value = "PS3"
        self.germline.cell(56, 4).value = "PM1"
        self.germline.cell(57, 4).value = "PP2"
        self.germline.cell(58, 4).value = "BS3"
        self.germline.cell(59, 4).value = "PP1"
        self.germline.cell(60, 4).value = "BS4"
        self.germline.cell(61, 4).value = "PS2"
        self.germline.cell(62, 4).value = "PM6"
        self.germline.cell(63, 4).value = "PM3"
        self.germline.cell(64, 4).value = "BP2"
        self.germline.cell(65, 4).value = "BS2"
        self.germline.cell(66, 4).value = "PP4"
        self.germline.cell(67, 4).value = "BP5"
        # add formula
        for row in range(42, 68):
            self.germline[
                f"B{row}"
            ] = f'=CONCATENATE("[", D{row}, "_", E{row}, ":] ", C{row})'

        # titles to set to bold
        to_bold = [
            "A1", "A8", "C2", "D2", "E2", "F2", "G2",
            "H2", "I2", "J2", "C13", "A40", "A41", "A42",
            "A46", "A55", "A59", "A61", "A63", "A66",
            "A69", "A70", "A71", "B41", "C41", "D41",
            "E41", "F41", "G41", "H41", "I41", "D68",
        ]

        for cell in to_bold:
            self.germline[cell].font = Font(bold=True)

        cell_lower_border = [
            "A8", "C13", "A41", "A45", "A54", "A58",
            "A60", "A62", "A65", "A67", "C45", "C54",
            "C58", "C60", "C62", "C65", "C67", "D67",
            "E67", "F67", "G67", "H67", "I67", "D68",
            "E68", "F68", "G68", "H68", "I68"
        ]
        for cell in cell_lower_border:
            self.germline[cell].border = LOWER_BORDER

        # set column widths for readability
        self.germline.column_dimensions["A"].width = 32

        # set borders around table areas
        row_ranges = [
            "C1:J1", "C2:J2", "C3:J3", "C4:J4",
            "C5:J5", "C6:J6", "C7:J7", "C8:J8",
            "C9:J9", "C10:J10", "C11:J11", "D42:I42",
            "D43:I43", "D44:I44", "D45:I45", "D46:I46",
            "D47:I47", "D48:I48", "D49:I49", "D50:I50",
            "D51:I51", "D52:I52", "D53:I53", "D54:I54",
            "D55:I55", "D56:I56", "D57:I57", "D58:I58",
            "D59:I59", "D60:I60", "D61:I61", "D62:I62",
            "D63:I63", "D64:I64", "D65:I65", "D66:I66",
            "D67:I67"
        ]
        for row in row_ranges:
            for cells in self.germline[row]:
                for cell in cells:
                    cell.border = THIN_BORDER

        # colour title cells
        blueFill = PatternFill(patternType="solid", start_color="ADD8E6")
        greenFill = PatternFill(patternType="solid", start_color="90EE90")
        pinkFill = PatternFill(patternType="solid", start_color="ffb6c1")

        blue_colour_cells = ["C2", "D2", "E2", "F2", "G2", "H2", "I2", "J2"]

        green_colour_cells = [
            "D44",
            "E44",
            "F44",
            "D45",
            "E45",
            "F45",
            "D51",
            "E51",
            "F51",
            "D52",
            "E52",
            "F52",
            "D53",
            "E53",
            "F53",
            "D54",
            "E54",
            "F54",
            "D58",
            "E58",
            "F58",
            "D60",
            "E60",
            "F60",
            "D64",
            "E64",
            "F64",
            "D65",
            "E65",
            "F65",
            "D67",
            "E67",
            "F67",
        ]

        pink_colour_cells = [
            "D42",
            "E42",
            "F42",
            "D43",
            "E43",
            "F43",
            "D46",
            "E46",
            "F46",
            "D47",
            "E47",
            "F47",
            "D48",
            "E48",
            "F48",
            "D49",
            "E49",
            "F49",
            "D50",
            "E50",
            "F50",
            "D55",
            "E55",
            "F55",
            "D56",
            "E56",
            "F56",
            "D57",
            "E57",
            "F57",
            "D59",
            "E59",
            "F59",
            "D61",
            "E61",
            "F61",
            "D62",
            "E62",
            "F62",
            "D63",
            "E63",
            "F63",
            "D66",
            "E66",
            "F66",
        ]
        for cell in blue_colour_cells:
            self.germline[cell].fill = blueFill
        for cell in green_colour_cells:
            self.germline[cell].fill = greenFill
        for cell in pink_colour_cells:
            self.germline[cell].fill = pinkFill

        # set column widths for readability
        self.germline.column_dimensions["A"].width = 36
        self.germline.column_dimensions["B"].width = 10
        for col in ["C", "G", "J"]:
            self.germline.column_dimensions[col].width = 28
        for col in ["D", "E"]:
            self.germline.column_dimensions[col].width = 32
        for col in ["F", "I"]:
            self.germline.column_dimensions[col].width = 20
        self.germline.column_dimensions["H"].width = 40

        smaller_font = Font(size=8)
        for i in range(41, 72):
            for cell in self.germline[f"{i}:{i}"]:
                cell.font = smaller_font

    def create_summary(self) -> None:
        """
        Write summary sheet
        """
        self.summary.cell(1, 1).value = "=SOC!A2"
        self.summary.cell(2, 1).value = "=SOC!A3"
        self.summary.cell(3, 1).value = "=SOC!A5"
        self.summary.cell(4, 1).value = "=SOC!A6"
        self.summary.cell(6, 1).value = "=SOC!A9"
        self.summary.cell(9, 1).value = "Reportable genes"
        self.summary.cell(
            10, 1
        ).value = '= _xlfn.TEXTJOIN(", ",TRUE,C21:C28,C33:C40)'
        self.summary.cell(12, 1).value = "Comments"

        self.summary.cell(19, 3).value = "Somatic SNV"
        self.summary.cell(20, 3).value = "Gene"
        self.summary.cell(20, 4).value = "GRCh38 Coordinates"
        self.summary.cell(20, 5).value = "Mutation"
        self.summary.cell(20, 6).value = "VAF"
        self.summary.cell(20, 7).value = "Variant Class"
        self.summary.cell(20, 8).value = "Validation"
        self.summary.cell(20, 9).value = "Actionability"
        # add formula
        for row in range(21, 29):
            ref_row = row + 22
            for col in ["C", "D", "E", "F"]:
                self.summary[f"{col}{row}"] = f"=summary!{col}{ref_row}"

        self.summary.cell(30, 3).value = "Somatic CNV_SV"
        self.summary.cell(31, 3).value = "Gene/Locus"
        self.summary.cell(31, 4).value = "GRCh38 Coordinates"
        self.summary.cell(31, 5).value = "Cytological Bands"
        self.summary.cell(31, 6).value = "Variant Type"
        self.summary.cell(31, 7).value = "Variant Class"
        self.summary.cell(31, 8).value = "Validation"
        self.summary.cell(31, 9).value = "Actionability"
        # add formula
        for row in range(32, 40):
            ref_row = row + 21
            for col in ["C", "D", "E", "F"]:
                self.summary[f"{col}{row}"] = f"=summary!{col}{ref_row}"

        self.summary.cell(41, 1).value = "SNV"
        self.summary.cell(42, 1).value = "Domain"
        self.summary.cell(42, 2).value = "Origin"
        self.summary.cell(42, 3).value = "Gene"
        self.summary.cell(42, 4).value = "GRCh38 coordinates;ref/alt allele"
        self.summary.cell(42, 5).value = "Transcript"
        self.summary.cell(42, 6).value = "CDS change and protein change"
        self.summary.cell(42, 7).value = "Predicted consequences"
        self.summary.cell(
            42, 8
        ).value = "Population germline allele frequency (GE | gnomAD)"
        self.summary.cell(42, 9).value = "VAF"
        self.summary.cell(42, 10).value = "Alt allele/total read depth"
        self.summary.cell(42, 11).value = "Genotype"
        self.summary.cell(42, 12).value = "COSMIC ID"
        self.summary.cell(42, 13).value = "ClinVar ID"
        self.summary.cell(42, 14).value = "ClinVar review status"
        self.summary.cell(42, 15).value = "ClinVar clinical significance"
        self.summary.cell(42, 16).value = "Gene mode of action"
        self.summary.cell(
            42, 17
        ).value = "Recruiting Clinical Trials 30 Jan 2023"
        self.summary.cell(42, 18).value = "PharmGKB_ID"

        self.summary.cell(51, 1).value = "CNV_SV"
        self.summary.cell(52, 2).value = "Origin"
        self.summary.cell(52, 3).value = "Variant domain"
        self.summary.cell(52, 4).value = "Event domain"
        self.summary.cell(52, 5).value = "Gene"
        self.summary.cell(52, 6).value = "Transcript"
        self.summary.cell(52, 7).value = "Impacted transcript region"
        self.summary.cell(52, 8).value = "GRCh38 coordinates"
        self.summary.cell(52, 9).value = "Type"
        self.summary.cell(52, 10).value = "Size"
        self.summary.cell(
            52, 11
        ).value = "Population germline allele frequency (GESG | GECG for somatic SVs or AF | AUC for germline CNVs)"
        self.summary.cell(52, 12).value = "Confidence/support"
        self.summary.cell(52, 13).value = "Chromosomal bands"
        self.summary.cell(
            52, 14
        ).value = "Recruiting Clinical Trials 30 Jan 2023"
        self.summary.cell(52, 15).value = "ClinVar clinical significance"
        self.summary.cell(52, 16).value = "Gene mode of action"

        # titles to set to bold
        to_bold = [
            "A1",
            "A9",
            "A12",
            "A16",
            "C20",
            "D20",
            "E20",
            "F20",
            "G20",
            "H20",
            "I20",
            "C31",
            "D31",
            "E31",
            "F31",
            "G31",
            "H31",
            "I31",
            "A41",
            "A51",
        ]

        for cell in to_bold:
            self.summary[cell].font = Font(bold=True, name=DEFAULT_FONT.name)

        # set column widths for readability
        for col in ["A", "H"]:
            self.summary.column_dimensions[col].width = 32
        self.summary.column_dimensions["C"].width = 22
        for col in ["D", "F", "G", "H"]:
            self.summary.column_dimensions[col].width = 26

        # colour title cells
        blueFill = PatternFill(patternType="solid", start_color="90EE90")

        colour_cells = [
            "C20",
            "D20",
            "E20",
            "F20",
            "G20",
            "H20",
            "I20",
            "C31",
            "D31",
            "E31",
            "F31",
            "G31",
            "H31",
            "I31",
        ]
        for cell in colour_cells:
            self.summary[cell].fill = blueFill

        # set borders around table areas
        row_ranges = [
            "C20:I20",
            "C21:I21",
            "C22:I22",
            "C23:I23",
            "C24:I24",
            "C25:I25",
            "C26:I26",
            "C27:I27",
            "C28:I28",
            "C31:I31",
            "C32:I32",
            "C33:I33",
            "C34:I34",
            "C35:I35",
            "C36:I36",
            "C37:I37",
            "C38:I38",
            "C39:I39",
        ]
        for row in row_ranges:
            for cells in self.summary[row]:
                for cell in cells:
                    cell.border = THIN_BORDER

        cell_lower_border = ["A9", "A12", "A41", "A51"]
        for cell in cell_lower_border:
            self.summary[cell].border = LOWER_BORDER

        smaller_font = Font(size=8)
        for i in range(41, 72):
            for cell in self.summary[f"{i}:{i}"]:
                cell.font = smaller_font

        cells_for_class = []
        for i in range(21, 40):
            if not i in [29, 30, 31]:
                cells_for_class.append(f"G{i}")

        class_options = (
            '"Pathogenic, Likely pathogenic, Uncertain, Likely benign, Benign"'
        )
        self.get_drop_down(
            dropdown_options=class_options,
            prompt="Select from the list",
            title="Variant class",
            sheet=self.summary,
            cells=cells_for_class,
        )

        cells_for_validation = []
        for i in range(21, 40):
            if not i in [29, 30, 31]:
                cells_for_validation.append(f"H{i}")

        validation_options = (
            '"Not indicated, Previously detected, {In progress}, VALIDATED"'
        )
        self.get_drop_down(
            dropdown_options=validation_options,
            prompt="Select from the list",
            title="Validation",
            sheet=self.summary,
            cells=cells_for_validation,
        )

        cells_for_action = []
        for i in range(21, 40):
            if not i in [29, 30, 31]:
                cells_for_action.append(f"I{i}")

        action_options = '"1. Predicts therapeutic response, 2. Prognostic, 3. Defines diagnosis group, 4. Eligibility for trial, 5. Other"'
        self.get_drop_down(
            dropdown_options=action_options,
            prompt="Select from the list",
            title="Actionability",
            sheet=self.summary,
            cells=cells_for_action,
        )

    def write_fusion(self) -> None:
        """
        write fusion sheet
        """
        df = pd.read_csv(fusion_ref, sep="\t")
        df.to_excel(self.writer, sheet_name="fusion", index=False)

    def write_refgene(self) -> None:
        """
        write RefGene sheet
        """
        df = pd.read_csv(refgene_ref, sep="\t")
        df["Reference"] = "COSMIC [Somatic]; [Germline]"
        df["RefGene Group"] = "COSMIC_Cancer_Genes"
        df.to_excel(self.writer, sheet_name="RefGene", index=False)
        ref_gene = self.writer.sheets["RefGene"]
        for col in ["D", "E", "F"]:
            ref_gene.column_dimensions[col].width = 32
        ref_gene.column_dimensions["G"].width = 28
        filters = ref_gene.auto_filter
        filters.ref = "A:G"

    def write_refgene_groups(self) -> None:
        """
        write RefGene_Groups sheet
        """
        df = pd.read_csv(refgenegp_ref, sep="\t")
        df.to_excel(self.writer, sheet_name="RefGene_Groups", index=False)
        ref_gene_gp = self.writer.sheets["RefGene_Groups"]
        for col in ["D", "E", "F", "G"]:
            ref_gene_gp.column_dimensions[col].width = 32
        filters = ref_gene_gp.auto_filter
        filters.ref = "A:G"

    def write_cytoref(self) -> None:
        """
        write CytoRef sheet
        """
        df = pd.read_csv(cyto_ref, sep="\t")
        df.to_excel(self.writer, sheet_name="CytoRef", index=False)
        cytoref = self.writer.sheets["CytoRef"]
        for col in ["D", "F", "G"]:
            cytoref.column_dimensions[col].width = 28
        cytoref.column_dimensions["E"].width = 42
        filters = cytoref.auto_filter
        filters.ref = "A:G"

    def write_hotspots(self) -> None:
        """
        write Hotspots sheet
        """
        df = pd.read_csv(hotspots_ref, sep="\t")
        df.to_excel(self.writer, sheet_name="Hotspots", index=False)
        hotspots = self.writer.sheets["Hotspots"]
        hotspots.column_dimensions["A"].width = 28
        hotspots.column_dimensions["B"].width = 24
        hotspots.column_dimensions["C"].width = 52
        filters = hotspots.auto_filter
        filters.ref = "A:E"

    def write_SNV(self) -> None:
        """
        write SNV sheet
        """
        df = pd.read_csv(self.args.variant, sep=",")
        num_variant = df.shape[0]
        df[["A_variant", "B_variant"]] = df[
            "CDS change and protein change"
        ].str.split(";", expand=True)
        df.to_excel(self.writer, sheet_name="SNV", index=False)
        SNV = self.writer.sheets["SNV"]
        SNV.column_dimensions["D"].width = 28
        for col in ["E", "F", "G", "H"]:
            SNV.column_dimensions[col].width = 24

        SNV["U1"] = "Report (Y/N)"
        SNV["V1"] = "Comments"
        SNV["W1"] = "Alteration_RefGene"
        SNV["X1"] = "Origin_RefGene"
        SNV["Y1"] = "Entities_RefGene"
        SNV["Z1"] = "Comments_RefGene"
        SNV["AA1"] = "HS_Sample"
        SNV["AB1"] = "HS_Tumour"
        SNV["AC1"] = "MTBP c."
        SNV["AD1"] = "MTBP p."
        SNV["AE1"] = "HS p."
        SNV["AF1"] = '";" count_Transcript.'

        filters = SNV.auto_filter
        filters.ref = "A:AF"

        # add vlookup
        for row in range(2, num_variant):
            SNV[f"W{row}"] = f"=VLOOKUP(C{row}, 'RefGene'!A:E, 2, FALSE())"
            SNV[f"X{row}"] = f"=VLOOKUP(C{row}, 'RefGene'!A:E, 3, FALSE())"
            SNV[f"Y{row}"] = f"=VLOOKUP(C{row}, 'RefGene'!A:E, 4, FALSE())"
            SNV[f"Z{row}"] = f"=VLOOKUP(C{row}, 'RefGene'!A:E, 5, FALSE())"
            SNV[f"AA{row}"] = f"=VLOOKUP(AE{row}, 'Hotspots'!A:C, 2, FALSE())"
            SNV[f"AB{row}"] = f"=VLOOKUP(AE{row}, 'Hotspots'!A:C, 3, FALSE())"
            SNV[f"AC{row}"] = f'=CONCATENATE(C{row}, ":", S{row})'
            SNV[f"AD{row}"] = f'=CONCATENATE(C{row}, ":", T{row})'
            num = "{1,2,3,4,5,6,7,8,9,0}"
            SNV[
                f"AE{row}"
            ] = f'=LEFT(AD{row},MAX(IFERROR(FIND({num},AD{row},ROW(INDIRECT("1:"&LEN(AD{row})))),0)))'
            SNV[f"AF{row}"] = f'=LEN(E{row})-LEN(SUBSTITUTE(E{row}, ";", ""))'

    def write_SV(self) -> None:
        """
        write SV, SV_loss and SV_gain sheets
        """
        df_SV = pd.read_csv(self.args.SV, sep=",")
        num_variant = df_SV.shape[0]
        df_loss = df_SV[df_SV["Type"].str.lower().str.contains("loss")]
        df_gain = df_SV[df_SV["Type"].str.lower().str.contains("gain")]

        for sheet, df in [
            ["SV", df_SV],
            ["SV_loss", df_loss],
            ["SV_gain", df_gain],
        ]:
            df.to_excel(self.writer, sheet_name=sheet, index=False)
            SV = self.writer.sheets[sheet]
            for col in ["D", "E", "F", "G"]:
                SV.column_dimensions[col].width = 24
            SV.column_dimensions["H"].width = 14

            SV["O1"] = "Report (Y/N)"
            SV["P1"] = "Comments"
            SV["Q1"] = "Alteration_RefGene"
            SV["R1"] = "Origin_RefGene"
            SV["S1"] = "Entities_RefGene"
            SV["T1"] = "Comments_RefGene"
            SV["U1"] = "B_LOOKUP"
            SV["V1"] = "C_LOOKUP"
            SV["W1"] = "D_LOOKUP"
            SV["X1"] = "A_Gene"
            SV["Y1"] = "B_Gene"
            SV["Z1"] = "C_Gene"
            SV["AA1"] = "D_Gene"

            # add vlookup
            if sheet == "SV":
                for row in range(2, num_variant):
                    SV[
                        f"Q{row}"
                    ] = f"=VLOOKUP(D{row}, 'RefGene'!A:E, 2, FALSE())"
                    SV[
                        f"R{row}"
                    ] = f"=VLOOKUP(D{row}, 'RefGene'!A:E, 3, FALSE())"
                    SV[
                        f"S{row}"
                    ] = f"=VLOOKUP(D{row}, 'RefGene'!A:E, 4, FALSE())"
                    SV[
                        f"T{row}"
                    ] = f"=VLOOKUP(D{row}, 'RefGene'!A:E, 5, FALSE())"

                    SV[
                        f"U{row}"
                    ] = f"=VLOOKUP(Y{row}, 'RefGene'!A:A, 1, FALSE())"
                    SV[
                        f"V{row}"
                    ] = f"=VLOOKUP(Z{row}, 'RefGene'!A:A, 1, FALSE())"
                    SV[
                        f"W{row}"
                    ] = f"=VLOOKUP(AA{row}, 'RefGene'!A:A, 1, FALSE())"
                    SV[f"X{row}"] = f"={sheet}!D{row}"
            elif sheet == "SV_loss":
                df_loss_index = [x + 1 for x in df_loss.index.tolist()]
                for row in range(2, (len(df_loss.index.tolist()) + 2)):
                    SV[
                        f"Q{row}"
                    ] = f"=VLOOKUP(D{df_loss_index[row-2]}, 'RefGene'!A:E, 2, FALSE())"
                    SV[
                        f"R{row}"
                    ] = f"=VLOOKUP(D{df_loss_index[row-2]}, 'RefGene'!A:E, 3, FALSE())"
                    SV[
                        f"S{row}"
                    ] = f"=VLOOKUP(D{df_loss_index[row-2]}, 'RefGene'!A:E, 4, FALSE())"
                    SV[
                        f"T{row}"
                    ] = f"=VLOOKUP(D{df_loss_index[row-2]}, 'RefGene'!A:E, 5, FALSE())"
                    SV[
                        f"U{row}"
                    ] = f"=VLOOKUP(Y{df_loss_index[row-2]}, 'RefGene'!A:A, 1, FALSE())"
                    SV[
                        f"V{row}"
                    ] = f"=VLOOKUP(Z{df_loss_index[row-2]}, 'RefGene'!A:A, 1, FALSE())"
                    SV[
                        f"W{row}"
                    ] = f"=VLOOKUP(AA{df_loss_index[row-2]}, 'RefGene'!A:A, 1, FALSE())"
                    SV[f"X{row}"] = f"={sheet}!D{row}"

            elif sheet == "SV_gain":
                df_gain_index = [x + 1 for x in df_gain.index.tolist()]
                for row in range(2, (len(df_gain.index.tolist()) + 2)):
                    SV[
                        f"Q{row}"
                    ] = f"=VLOOKUP(D{df_gain_index[row-2]}, 'RefGene'!A:E, 2, FALSE())"
                    SV[
                        f"R{row}"
                    ] = f"=VLOOKUP(D{df_gain_index[row-2]}, 'RefGene'!A:E, 3, FALSE())"
                    SV[
                        f"S{row}"
                    ] = f"=VLOOKUP(D{df_gain_index[row-2]}, 'RefGene'!A:E, 4, FALSE())"
                    SV[
                        f"T{row}"
                    ] = f"=VLOOKUP(D{df_gain_index[row-2]}, 'RefGene'!A:E, 5, FALSE())"
                    SV[
                        f"U{row}"
                    ] = f"=VLOOKUP(Y{df_gain_index[row-2]}, 'RefGene'!A:A, 1, FALSE())"
                    SV[
                        f"V{row}"
                    ] = f"=VLOOKUP(Z{df_gain_index[row-2]}, 'RefGene'!A:A, 1, FALSE())"
                    SV[
                        f"W{row}"
                    ] = f"=VLOOKUP(AA{df_gain_index[row-2]}, 'RefGene'!A:A, 1, FALSE())"
                    SV[f"X{row}"] = f"={sheet}!D{row}"

            filters = SV.auto_filter
            filters.ref = "A:AA"

    def get_drop_down(
        self, dropdown_options, prompt, title, sheet, cells
    ) -> None:
        """
        create the drop-downs items for designated cells

        Parameters
        ----------
        dropdown_options: str
            str containing drop-down items
        prompt: str
            prompt message for drop-down
        title: str
            title message for drop-down
        sheet: openpyxl.Writer writer object
            current worksheet
        cells: list
            list of cells to add drop-down
        """
        options = dropdown_options
        val = DataValidation(type="list", formula1=options, allow_blank=True)
        val.prompt = prompt
        val.promptTitle = title
        sheet.add_data_validation(val)
        for cell in cells:
            val.add(sheet[cell])
        val.showInputMessage = True
        val.showErrorMessage = True

    def insert_img(self, sheet, img_to_insert, cell_to_insert, h, w) -> None:
        """
        insert the img downloaed from html into spreadsheet
        """
        ws = sheet
        img = drawing.image.Image(img_to_insert)
        img.height = h
        img.width = w
        img.anchor = cell_to_insert
        ws.add_image(img)


def main():
    # generate output Excel file
    excel_handler = excel()
    excel_handler.generate()


if __name__ == "__main__":
    main()
