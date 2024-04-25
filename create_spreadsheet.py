import argparse
import numpy as np
from openpyxl import load_workbook, drawing
from openpyxl.styles import Alignment, Border, DEFAULT_FONT, Font, Side
from openpyxl.styles.fills import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import urllib.request
from bs4 import BeautifulSoup

# ref files
CYTO_REF = "./resources/CytoRef.txt"
FUSION_REF = "./resources/fusions.txt"
HTOSPOTS_REF = "./resources/Hotspots.txt"
REFGENE_REF = "./resources/Ref_Gene.txt"
REFGENEGP_REF = "./resources/RefGene_groups.txt"


# openpyxl style settings
THIN = Side(border_style="thin", color="000000")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LOWER_BORDER = Border(bottom=THIN)
DEFAULT_FONT.name = "Calibri"


class excel:
    """
    Functions for wrangling input csv files, ref files and html files and
    writing output xlsm file

    Attributes
    ----------
    args : argparse.Namespace
        arguments passed from command line
    writer : pandas.io.excel._openpyxl.OpenpyxlWriter
        writer object for writing Excel data to file
    workbook : openpyxl.workbook.workbook.Workbook
        openpyxl workbook object for interacting with per-sheet writing and
        formatting of output Excel file

    Outputs
    -------
    {args.output}.xlsm : file
        Excel file with variants, structural variants and ref sheets
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

        parser.add_argument(
            "--output", "--o", required=True, help="output xlsm file name"
        )
        parser.add_argument("--html", required=True, help="html input")
        parser.add_argument(
            "--variant", "--v", required=True, help="variant csv file"
        )
        parser.add_argument(
            "--SV", "--sv", required=True, help="structural variant csv file"
        )

        return parser.parse_args()

    def generate(self) -> None:
        """
        Calls all methods in excel() to generate output xlsm
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
        soup = self.get_soup()
        n = 1
        for link in soup.findAll("img"):
            img_link = link.get("src")
            self.download_image(img_link, "./", f"figure_{n}")
            n = n + 1

    def download_image(self, url, file_path, file_name):
        """
        Download the img from html file
        """
        full_path = file_path + file_name + ".jpg"
        urllib.request.urlretrieve(url, full_path)

    def read_html_tables(self, table_num) -> list:
        """
        get the tables from html file

        Parameters
        ----------
        table number

        Returns
        -------
        list of each html table
        """
        soup = self.get_soup()
        tables = soup.findAll("table")
        info = tables[table_num]
        headings = [th.get_text() for th in info.find("tr").find_all("th")]
        datasets = []
        for row in info.find_all("tr")[1:]:
            dataset = dict(
                zip(headings, (td.get_text() for td in row.find_all("td")))
            )
            datasets.append(dataset)

        return datasets

    def get_soup(self) -> BeautifulSoup:
        """
        get Beautiful soup obj from url

        Returns
        -------
        Beautiful soup object
        """
        url = self.args.html
        f = urllib.request.urlopen(url)
        page = f.read()
        f.close()
        soup = BeautifulSoup(page, features="lxml")

        return soup

    def write_sheets(self) -> None:
        """
        Write sheets to xlsm file
        """
        print("Writing sheets")
        self.write_refgene()
        self.soc = self.workbook.create_sheet("SOC")
        self.write_soc()
        self.QC = self.workbook.create_sheet("QC")
        self.write_QC()
        self.plot = self.workbook.create_sheet("Plot")
        self.write_plot()
        self.signatures = self.workbook.create_sheet("Signatures")
        self.write_signatures()
        self.germline = self.workbook.create_sheet("Germline")
        self.write_germline()
        self.summary = self.workbook.create_sheet("Summary")
        self.write_summary()
        self.write_fusion()
        self.write_refgene_groups()
        self.write_cytoref()
        self.write_hotspots()
        self.write_SNV()
        self.write_SV()

    def set_col_width(self, cell_width, sheet):
        """
        set the column width for given col in given sheet
        """
        for cell, width in cell_width:
            sheet.column_dimensions[cell].width = width

    def bold_cell(self, cells_to_bold, sheet):
        """
        bold the given cells in given sheet
        """
        for cell in cells_to_bold:
            sheet[cell].font = Font(bold=True, name=DEFAULT_FONT.name)

    def colour_cell(self, cells_to_colour, sheet, fill):
        """
        colour the given cells in given sheet
        """
        for cell in cells_to_colour:
            sheet[cell].fill = fill

    def all_border(self, row_ranges, sheet):
        """
        create all borders for given cells in given sheet
        """
        for row in row_ranges:
            for cells in sheet[row]:
                for cell in cells:
                    cell.border = THIN_BORDER

    def lower_border(self, cells_lower_border, sheet):
        """
        create lower cell border for given cells in given sheet
        """
        for cell in cells_lower_border:
            sheet[cell].border = LOWER_BORDER

    def write_soc(self) -> None:
        """
        Write soc sheet
        """
        self.patient_info = self.read_html_tables(0)
        # write titles for summary values
        self.soc.cell(1, 1).value = "Patient Details (Epic demographics)"
        self.soc.cell(1, 3).value = "Previous testing"
        self.soc.cell(2, 1).value = "Name"
        self.soc.cell(2, 3).value = "Alteration"
        self.soc.cell(2, 4).value = "Assay"
        self.soc.cell(2, 5).value = "Result"
        self.soc.cell(2, 6).value = "WGS concordance"
        self.soc.cell(3, 1).value = self.patient_info[0]["Gender"]
        self.soc.cell(4, 1).value = self.patient_info[0]["Patient ID"]
        self.soc.cell(5, 1).value = "MRN"
        self.soc.cell(6, 1).value = "NHS Number"
        self.soc.cell(8, 1).value = "Histology"
        self.soc.cell(12, 1).value = "Comments"
        self.soc.cell(16, 1).value = "WGS in-house gene panel applied"
        self.soc.cell(17, 1).value = self.df_refgene["RefGene Group"][0]

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
        self.bold_cell(to_bold, self.soc)

        # set column widths for readability
        cell_col_width = (
            ("A", 32),
            ("C", 16),
            ("E", 16),
            ("D", 26),
            ("F", 26),
        )
        self.set_col_width(cell_col_width, self.soc)

        # colour title cells
        greenFill = PatternFill(patternType="solid", start_color="90EE90")
        colour_cells = ["C3", "D3", "E3", "F3", "C4", "D4", "E4", "F4"]
        self.colour_cell(colour_cells, self.soc, greenFill)

        # set borders around table areas
        row_ranges = [
            "C1:F1",
            "C2:F2",
            "C3:F3",
            "C4:F4",
            "C5:F5",
            "C6:F6",
            "C7:F7",
            "C8:F8",
        ]
        self.all_border(row_ranges, self.soc)

        cells_lower_border = ["A1", "A8", "A12", "A16"]
        self.lower_border(cells_lower_border, self.soc)

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

    def write_QC(self) -> None:
        """
        write QC sheet
        """
        tumor_info = self.read_html_tables(1)
        sample_info = self.read_html_tables(2)
        germline_info = self.read_html_tables(3)
        seq_info = self.read_html_tables(4)

        # PID table
        self.QC.cell(1, 1).value = "=SOC!A2"
        self.QC.cell(2, 1).value = "=SOC!A3"
        self.QC.cell(3, 1).value = "=SOC!A5"
        self.QC.cell(4, 1).value = "=SOC!A6"
        self.QC.cell(6, 1).value = "=SOC!A9"
        self.QC.cell(8, 1).value = "QC alerts"
        self.QC.cell(9, 1).value = "None"

        # table 1
        table1_keys = (
            (3, "Diagnosis Date"),
            (4, "Tumour Received"),
            (5, "Tumour ID"),
            (6, "Presentation"),
            (7, "Diagnosis"),
            (8, "Tumour Site"),
            (9, "Tumour Type"),
            (10, "Germline Sample"),
        )
        for cell, key in table1_keys:
            self.QC.cell(1, cell).value = key

        table1_values = (
            (3, tumor_info[0]["Tumour Diagnosis Date"]),
            (4, germline_info[0]["Clinical Sample Date Time"]),
            (5, tumor_info[0]["Histopathology or SIHMDS LAB ID"]),
            (
                6,
                tumor_info[0]["Presentation"]
                + " "
                + tumor_info[0]["Primary or Metastatic"],
            ),
            (7, self.patient_info[0]["Clinical Indication"]),
            (8, tumor_info[0]["Tumour Topography"]),
            (
                9,
                sample_info[0]["Storage Medium"]
                + " "
                + sample_info[0]["Source"],
            ),
            (
                10,
                germline_info[0]["Storage Medium"]
                + " ("
                + germline_info[0]["Source"]
                + ")",
            ),
        )
        for cell, value in table1_values:
            self.QC.cell(2, cell).value = value

        # table 2
        table2_keys = (
            (3, "Purity (Histo)"),
            (4, "Purity (Calculated)"),
            (5, "Ploidy"),
            (6, "Total SNVs"),
            (7, "Total Indels"),
            (8, "Total SVs"),
            (9, "TMB"),
        )
        for cell, key in table2_keys:
            self.QC.cell(4, cell).value = key
        table2_values = (
            (3, sample_info[0]["Tumour Content"]),
            (4, sample_info[0]["Tumour Content"]),
            (5, sample_info[0]["Calculated Overall Ploidy"]),
            (6, seq_info[1]["Total somatic SNVs"]),
            (7, seq_info[1]["Total somatic indels"]),
            (8, seq_info[1]["Total somatic SVs"]),
        )
        for cell, value in table2_values:
            self.QC.cell(5, cell).value = value

        # table 3
        table3_keys = (
            (3, "Sample type"),
            (4, "Mean depth, x"),
            (5, "Mapped reads, %"),
            (6, "Chimeric DNA frag, %"),
            (7, "Insert size, bp"),
            (8, "Unevenness, x"),
        )
        for cell, key in table3_keys:
            self.QC.cell(7, cell).value = key

        seq_info_title = [
            "Sample type",
            "Genome-wide coverage mean, x",
            "Mapped reads, %",
            "Chimeric DNA fragments, %",
            "Insert size median, bp",
            "Unevenness of local genome coverage, x",
        ]
        for title in seq_info_title:
            for i in range(3, 9):
                self.QC.cell(8, i).value = seq_info[0][title]
                self.QC.cell(9, i).value = seq_info[1][title]

        # titles to set to bold
        to_bold = [
            "A1",
            "A8",
            "C1",
            "D1",
            "E1",
            "F1",
            "G1",
            "H1",
            "I1",
            "J1",
            "C4",
            "D4",
            "E4",
            "F4",
            "G4",
            "H4",
            "I4",
            "C7",
            "D7",
            "E7",
            "F7",
            "G7",
            "H7",
        ]
        self.bold_cell(to_bold, self.QC)

        # set column widths for readability
        cell_col_width = (
            ("A", 32),
            ("B", 8),
            ("C", 22),
            ("D", 22),
            ("E", 22),
            ("F", 22),
            ("G", 22),
            ("H", 22),
            ("I", 22),
            ("J", 22),
        )
        self.set_col_width(cell_col_width, self.QC)

        # set borders around table areas
        row_ranges = [
            "C1:J1",
            "C2:J2",
            "C4:I4",
            "C5:I5",
            "C7:H7",
            "C8:H8",
            "C9:H9",
        ]
        self.all_border(row_ranges, self.QC)
        self.lower_border(["A8"], self.QC)

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
        self.insert_img(self.QC, "figure_9.jpg", "C12", 400, 600)

    def write_plot(self) -> None:
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
        self.bold_cell(to_bold, self.plot)

        # set column widths for readability
        self.plot.column_dimensions["A"].width = 32

    def write_signatures(self) -> None:
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

        # titles to set to bold
        to_bold = ["A1", "A8", "A13"]
        self.bold_cell(to_bold, self.signatures)
        self.lower_border(["A8"], self.signatures)

        # set column widths for readability
        self.signatures.column_dimensions["A"].width = 32
        self.insert_img(self.signatures, "figure_6.jpg", "C3", 700, 1000)
        self.insert_img(self.signatures, "figure_7.jpg", "P3", 400, 600)

    def write_germline(self) -> None:
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
        snv_table_keys = (
            (3, "Gene"),
            (4, "GRCh38 Coordinates"),
            (5, "Variant"),
            (6, "Genotype"),
            (7, "Role in Cacer"),
            (8, "ClinVar"),
            (9, "gnomAD"),
            (10, "Tumour VAF"),
        )
        for cell, key in snv_table_keys:
            self.germline.cell(2, cell).value = key

        for row in range(3, 12):
            ref_row = row + 22
            for col in ["C", "D", "E", "F", "G", "H"]:
                self.germline[f"{col}{row}"] = f"=germline!{col}{ref_row}"

        self.germline.cell(13, 3).value = "Clinical genetics feedback"

        domain_talbe_keys = (
            (1, "Domain"),
            (2, "Origin"),
            (3, "Gene"),
            (4, "GRCh38 coordinates;ref/alt allele"),
            (5, "Transcript"),
            (6, "CDS change and protein change"),
            (7, "Predicted consequences"),
            (8, "Population germline allele frequency (GE | gnomAD)"),
            (9, "VAF"),
            (10, "Alt allele/total read depth"),
            (11, "Genotype"),
            (12, "COSMIC ID"),
            (13, "ClinVar ID"),
            (14, "ClinVar review status"),
            (15, "ClinVar clinical significance"),
            (16, "Gene mode of action"),
            (17, "Recruiting Clinical Trials 30 Jan 2023"),
            (18, "PharmGKB_ID"),
        )
        for cell, key in domain_talbe_keys:
            self.germline.cell(24, cell).value = key

        # ACMG classification table
        self.germline.cell(40, 1).value = "ACMG classification table"
        ACMG_table_keys = (
            (1, "Theme"),
            (2, "Epic"),
            (3, "Description (CanVIG-UK guidelines, v2.17)"),
            (4, "Criteria"),
            (5, "Strength"),
            (6, "Points"),
            (7, "Prelim"),
            (8, "Check"),
            (9, "Comments"),
        )
        for cell, key in ACMG_table_keys:
            self.germline.cell(41, cell).value = key
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

        # Description column
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

        # criteria column
        criteria_col_values = (
            (42, "PS4"),
            (43, "PM2"),
            (44, "BS1"),
            (45, "BA1"),
            (46, "PVS1"),
            (47, "PM4"),
            (48, "PS1"),
            (49, "PM5"),
            (50, "PP3"),
            (51, "BP4"),
            (52, "BP1"),
            (53, "BP3"),
            (54, "BP7"),
            (55, "PS3"),
            (56, "PM1"),
            (57, "PP2"),
            (58, "BS3"),
            (59, "PP1"),
            (60, "BS4"),
            (61, "PS2"),
            (62, "PM6"),
            (63, "PM3"),
            (64, "BP2"),
            (65, "BS2"),
            (66, "PP4"),
            (67, "BP5"),
        )
        for cell, value in criteria_col_values:
            self.germline.cell(cell, 4).value = value
        # add formula
        for row in range(42, 68):
            self.germline[
                f"B{row}"
            ] = f'=CONCATENATE("[", D{row}, "_", E{row}, ":] ", C{row})'

        # titles to set to bold
        to_bold = [
            "A1",
            "A8",
            "C2",
            "D2",
            "E2",
            "F2",
            "G2",
            "H2",
            "I2",
            "J2",
            "C13",
            "A40",
            "A41",
            "A42",
            "A46",
            "A55",
            "A59",
            "A61",
            "A63",
            "A66",
            "A69",
            "A70",
            "A71",
            "B41",
            "C41",
            "D41",
            "E41",
            "F41",
            "G41",
            "H41",
            "I41",
            "D68",
        ]
        self.bold_cell(to_bold, self.germline)

        cells_lower_border = [
            "A8",
            "C13",
            "A41",
            "A45",
            "A54",
            "A58",
            "A60",
            "A62",
            "A65",
            "A67",
            "C45",
            "C54",
            "C58",
            "C60",
            "C62",
            "C65",
            "C67",
            "D67",
            "E67",
            "F67",
            "G67",
            "H67",
            "I67",
            "D68",
            "E68",
            "F68",
            "G68",
            "H68",
            "I68",
        ]
        self.lower_border(cells_lower_border, self.germline)

        # set column widths for readability
        self.germline.column_dimensions["A"].width = 32

        # set borders around table areas
        row_ranges = [
            "C1:J1",
            "C2:J2",
            "C3:J3",
            "C4:J4",
            "C5:J5",
            "C6:J6",
            "C7:J7",
            "C8:J8",
            "C9:J9",
            "C10:J10",
            "C11:J11",
            "D42:I42",
            "D43:I43",
            "D44:I44",
            "D45:I45",
            "D46:I46",
            "D47:I47",
            "D48:I48",
            "D49:I49",
            "D50:I50",
            "D51:I51",
            "D52:I52",
            "D53:I53",
            "D54:I54",
            "D55:I55",
            "D56:I56",
            "D57:I57",
            "D58:I58",
            "D59:I59",
            "D60:I60",
            "D61:I61",
            "D62:I62",
            "D63:I63",
            "D64:I64",
            "D65:I65",
            "D66:I66",
            "D67:I67",
        ]
        self.all_border(row_ranges, self.germline)

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
        self.colour_cell(blue_colour_cells, self.germline, blueFill)
        self.colour_cell(green_colour_cells, self.germline, greenFill)
        self.colour_cell(pink_colour_cells, self.germline, pinkFill)

        # set column widths for readability
        cell_col_width = (
            ("A", 36),
            ("B", 10),
            ("C", 28),
            ("D", 32),
            ("E", 32),
            ("F", 20),
            ("G", 28),
            ("H", 40),
            ("I", 20),
            ("J", 28),
        )
        self.set_col_width(cell_col_width, self.germline)
        smaller_font = Font(size=8)
        for i in range(41, 72):
            for cell in self.germline[f"{i}:{i}"]:
                cell.font = smaller_font

    def write_summary(self) -> None:
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

        # snv table
        self.summary.cell(19, 3).value = "Somatic SNV"
        snv_table_keys = (
            (3, "Gene"),
            (4, "GRCh38 Coordinates"),
            (5, "Mutation"),
            (6, "VAF"),
            (7, "Variant Class"),
            (8, "Validation"),
            (9, "Actionability"),
        )
        for cell, key in snv_table_keys:
            self.summary.cell(20, cell).value = key
        # add formula
        for row in range(21, 29):
            ref_row = row + 22
            for col in ["C", "D", "E", "F"]:
                self.summary[f"{col}{row}"] = f"=summary!{col}{ref_row}"

        # cnv sv table
        self.summary.cell(30, 3).value = "Somatic CNV_SV"
        cnv_sv_table_key = (
            (3, "Gene/Locus"),
            (4, "GRCh38 Coordinates"),
            (5, "Cytological Bands"),
            (6, "Variant Type"),
            (7, "Variant Class"),
            (8, "Validation"),
            (9, "Actionability"),
        )
        for cell, key in cnv_sv_table_key:
            self.summary.cell(31, cell).value = key
        # add formula
        for row in range(32, 40):
            ref_row = row + 21
            for col in ["C", "D", "E", "F"]:
                self.summary[f"{col}{row}"] = f"=summary!{col}{ref_row}"

        self.summary.cell(41, 1).value = "SNV"
        # snv title
        snv_title_keys = (
            (1, "Domain"),
            (2, "Origin"),
            (3, "Gene"),
            (4, "GRCh38 coordinates;ref/alt allele"),
            (5, "Transcript"),
            (6, "CDS change and protein change"),
            (7, "Predicted consequences"),
            (8, "Population germline allele frequency (GE | gnomAD)"),
            (9, "VAF"),
            (10, "Alt allele/total read depth"),
            (11, "Genotype"),
            (12, "COSMIC ID"),
            (13, "ClinVar ID"),
            (14, "ClinVar review status"),
            (15, "ClinVar clinical significance"),
            (16, "Gene mode of action"),
            (17, "Recruiting Clinical Trials 30 Jan 2023"),
            (18, "PharmGKB_ID"),
        )
        for cell, key in snv_title_keys:
            self.summary.cell(42, cell).value = key

        self.summary.cell(51, 1).value = "CNV_SV"

        # cnv_sv title
        cnv_sv_title_keys = (
            (1, "Origin"),
            (2, "Variant domain"),
            (3, "Event domain"),
            (4, "Gene"),
            (5, "Transcript"),
            (6, "Impacted transcript region"),
            (7, "GRCh38 coordinates"),
            (8, "Type"),
            (9, "Size"),
            (
                10,
                "Population germline allele frequency (GESG | GECG for somatic SVs or AF | AUC for germline CNVs)",
            ),
            (11, "Confidence/support"),
            (12, "Chromosomal bands"),
            (13, "Recruiting Clinical Trials 30 Jan 2023"),
            (14, "ClinVar clinical significance"),
            (15, "Gene mode of action"),
        )
        for cell, key in cnv_sv_title_keys:
            self.summary.cell(52, cell).value = key

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
        self.bold_cell(to_bold, self.summary)

        # set column widths for readability
        cell_col_width = (
            ("A", 32),
            ("C", 22),
            ("D", 26),
            ("F", 26),
            ("G", 26),
            ("H", 26),
        )
        self.set_col_width(cell_col_width, self.summary)

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
        self.colour_cell(colour_cells, self.summary, blueFill)

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
        self.all_border(row_ranges, self.summary)
        cells_lower_border = ["A9", "A12", "A41", "A51"]
        self.lower_border(cells_lower_border, self.summary)

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
            if i not in [29, 30, 31]:
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
            if i not in [29, 30, 31]:
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
        df = pd.read_csv(FUSION_REF, sep="\t")
        df.to_excel(
            self.writer, sheet_name="fusion", index=False, header=False
        )

    def write_refgene(self) -> None:
        """
        write RefGene sheet
        """
        self.df_refgene = pd.read_csv(REFGENE_REF, sep="\t")
        self.df_refgene.to_excel(
            self.writer, sheet_name="RefGene", index=False
        )
        ref_gene = self.writer.sheets["RefGene"]
        cell_col_width = (("D", 32), ("E", 32), ("F", 32), ("G", 28))
        self.set_col_width(cell_col_width, ref_gene)
        filters = ref_gene.auto_filter
        filters.ref = "A:G"

    def write_refgene_groups(self) -> None:
        """
        write RefGene_Groups sheet
        """
        df = pd.read_csv(REFGENEGP_REF, sep="\t")
        df.to_excel(self.writer, sheet_name="RefGene_Groups", index=False)
        ref_gene_gp = self.writer.sheets["RefGene_Groups"]
        cell_col_width = (("D", 32), ("E", 32), ("F", 32), ("G", 32))
        self.set_col_width(cell_col_width, ref_gene_gp)
        filters = ref_gene_gp.auto_filter
        filters.ref = "A:G"

    def write_cytoref(self) -> None:
        """
        write CytoRef sheet
        """
        df = pd.read_csv(CYTO_REF, sep="\t")
        df.to_excel(self.writer, sheet_name="CytoRef", index=False)
        cytoref = self.writer.sheets["CytoRef"]
        cell_col_width = (("D", 28), ("E", 42), ("F", 28), ("G", 28))
        self.set_col_width(cell_col_width, cytoref)
        filters = cytoref.auto_filter
        filters.ref = "A:G"

    def write_hotspots(self) -> None:
        """
        write Hotspots sheet
        """
        self.df_hotspots = pd.read_csv(HTOSPOTS_REF, sep="\t")
        self.df_hotspots.to_excel(
            self.writer, sheet_name="Hotspots", index=False
        )
        hotspots = self.writer.sheets["Hotspots"]
        cell_col_width = (("A", 28), ("B", 24), ("C", 52))
        self.set_col_width(cell_col_width, hotspots)
        filters = hotspots.auto_filter
        filters.ref = "A:E"

    def write_SNV(self) -> None:
        """
        write SNV sheet
        """
        df = pd.read_csv(self.args.variant, sep=",")
        df[["A_variant", "B_variant"]] = df[
            "CDS change and protein change"
        ].str.split(";", expand=True)
        df["Report (Y/N)"] = ""
        df["Comments"] = ""
        df["Alteration_RefGene"] = df["Gene"].map(
            self.df_refgene.set_index("Gene")["Alteration"]
        )
        df["Origin_RefGene"] = df["Gene"].map(
            self.df_refgene.set_index("Gene")["Origin"]
        )
        df["Entities_RefGene"] = df["Gene"].map(
            self.df_refgene.set_index("Gene")["Entities"]
        )
        df["Comments_RefGene"] = df["Gene"].map(
            self.df_refgene.set_index("Gene")["Comments"]
        )
        df = df.replace([None], [""], regex=True)
        df["MTBP c."] = df["Gene"] + ":" + df["A_variant"]
        df["MTBP p."] = df["Gene"] + ":" + df["B_variant"]
        df[["HS p.", "col1", "col2"]] = df["MTBP p."].str.split(
            r"([^\d]+)$", expand=True
        )
        df.drop(["col1", "col2"], axis=1, inplace=True)
        df["';' count_Transcript."] = df["Transcript"].str.count(r"\;")
        df["HS_Sample"] = df["HS p."].map(
            self.df_hotspots.set_index("HS_PROTEIN_ID")["HS_Samples"]
        )
        df["HS_Tumour"] = df["HS p."].map(
            self.df_hotspots.set_index("HS_PROTEIN_ID")[
                "HS_Tumor Type Composition"
            ]
        )
        df.to_excel(self.writer, sheet_name="SNV", index=False)
        self.SNV = self.writer.sheets["SNV"]
        cell_col_width = (
            ("D", 28),
            ("E", 24),
            ("F", 24),
            ("G", 24),
            ("H", 24),
        )
        self.set_col_width(cell_col_width, self.SNV)
        filters = self.SNV.auto_filter
        filters.ref = "A:AF"

    def write_SV(self) -> None:
        """
        write SV, SV_loss and SV_gain sheets
        """
        df_SV = pd.read_csv(self.args.SV, sep=",")
        df_SV["Report (Y/N)"] = ""
        df_SV["Comments"] = ""
        df_SV["Alteration_RefGene"] = df_SV["Gene"].map(
            self.df_refgene.set_index("Gene")["Alteration"]
        )
        df_SV["Origin_RefGene"] = df_SV["Gene"].map(
            self.df_refgene.set_index("Gene")["Origin"]
        )
        df_SV["Entities_RefGene"] = df_SV["Gene"].map(
            self.df_refgene.set_index("Gene")["Entities"]
        )
        df_SV["Comments_RefGene"] = df_SV["Gene"].map(
            self.df_refgene.set_index("Gene")["Comments"]
        )
        df_SV["Comments_RefGene"] = df_SV["Gene"].map(
            self.df_refgene.set_index("Gene")["Comments"]
        )
        df_SV["gene_count"] = df_SV["Gene"].str.count(r"\;")
        max_num_gene = df_SV["gene_count"].max() + 1
        print(max_num_gene)
        if max_num_gene == 1:
            df_SV["A_Gene"] = df_SV["Gene"]
            df_SV[["B_Gene", "C_Gene", "D_Gene"]] = ""
        elif max_num_gene == 2:
            df_SV[["A_Gene", "B_Gene"]] = df_SV["Gene"].str.split(";", expand=True)
            df_SV[["C_Gene", "D_Gene"]] = ""
            df_SV["B_LOOKUP"] = np.where(df_SV["B_Gene"].isin(list(self.df_refgene['Gene'])), df_SV['B_Gene'], "")
            
        elif max_num_gene == 3:
            df_SV[["A_Gene", "B_Gene", "C_Gene"]] = df_SV["Gene"].str.split(";", expand=True)
            df_SV[["D_Gene"]] = ""
            df_SV["B_LOOKUP"] = np.where(df_SV["B_Gene"].isin(list(self.df_refgene['Gene'])), df_SV['B_Gene'], "")
            df_SV["C_LOOKUP"] = np.where(df_SV["C_Gene"].isin(list(self.df_refgene['Gene'])), df_SV['C_Gene'], "")
        elif max_num_gene == 4:
            df_SV[["A_Gene", "B_Gene", "C_gene", "D_gene"]] = df_SV["Gene"].str.split(";", expand=True)
            df_SV["B_LOOKUP"] = np.where(df_SV["B_Gene"].isin(list(self.df_refgene['Gene'])), df_SV['B_Gene'], "")
            df_SV["C_LOOKUP"] = np.where(df_SV["C_Gene"].isin(list(self.df_refgene['Gene'])), df_SV['C_Gene'], "")
            df_SV["D_LOOKUP"] = np.where(df_SV["D_Gene"].isin(list(self.df_refgene['Gene'])), df_SV['D_Gene'], "")

        df_loss = df_SV[df_SV["Type"].str.lower().str.contains("loss")]
        df_gain = df_SV[df_SV["Type"].str.lower().str.contains("gain")]

        df_to_write = (
            (df_SV, "SV"),
            (df_loss, "SV_loss"),
            (df_gain, "SV_gain"),
        )
        for df, sheet_name in df_to_write:

            df.to_excel(self.writer, sheet_name=sheet_name, index=False)
            sheet = self.writer.sheets[sheet_name]
            cell_col_width = (
                ("D", 24),
                ("E", 24),
                ("F", 24),
                ("G", 24),
                ("H", 14),
            )
            self.set_col_width(cell_col_width, sheet)
            filters = sheet.auto_filter
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
