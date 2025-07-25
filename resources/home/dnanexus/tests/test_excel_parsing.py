from unittest.mock import patch

import numpy as np
import pandas as pd
import pytest

from utils import excel_parsing


@pytest.fixture()
def germline_variant_data():
    df = pd.DataFrame(
        {
            "Origin": ["germline", "somatic", "germline"],
            "Gene": ["gene1", "gene2", "gene3"],
            "GRCh38 coordinates;ref/alt allele": [
                "coor1",
                "coor2",
                "coor3",
            ],
            "ClinVar ID": ["1", "", "2"],
            "Population germline allele frequency (GE | gnomAD)": [
                "ge1|freq1",
                "",
                "ge2|freq2",
            ],
            "Predicted consequences": [
                "consequence1",
                "consequence2",
                "consequence3",
            ],
            "CDS change and protein change": ["c.1", "c.2", "c.3"],
            "Genotype": ["0/1", "0|0", "1|1"],
            "Gene mode of action": ["deletion1", "missense1", "deletion2"],
        }
    )

    yield df
    del df


@pytest.fixture()
def panelapp_dfs():
    panelapp_dfs = {
        "Adult_v2.2": pd.DataFrame(
            {
                "Gene Symbol": ["gene1", "gene3"],
                "Formatted mode": ["mode1", "mode3"],
            }
        ),
        "Childhood_v4.0": pd.DataFrame(
            {
                "Gene Symbol": ["gene1", "gene3"],
                "Formatted mode": ["mode2", "mode4"],
            }
        ),
    }

    yield panelapp_dfs
    del panelapp_dfs


@pytest.fixture()
def somatic_variant_data():
    test_input = pd.DataFrame(
        {
            "Origin": ["somatic", "germline", "somatic"],
            "Domain": ["domain1", "domain2", "domain3"],
            "GRCh38 coordinates;ref/alt allele": [
                "coor1",
                "coor2",
                "coor3",
            ],
            "RefSeq IDs": ["refseq_id1", "refseq_id2", "refseq_id3"],
            "Gene": ["gene1", "gene2", "gene3"],
            "CDS change and protein change": [
                "c.1;p.Leu40Arg",
                "c.2;p.Gln61Arg",
                "c.3;p.Leu2135Val",
            ],
            "Predicted consequences": [
                "consequence1",
                "consequence2",
                "consequence3;consequence4",
            ],
            "Population germline allele frequency (GE | gnomAD)": [
                "0.1 | 0.2",
                "0.2 | -",
                "- | -",
            ],
            "Alt allele/total read depth": ["depth1", "depth2", "depth3"],
            "Gene mode of action": ["mode1", "mode2", "mode3"],
            "VAF": ["0.3;0.1", 0.6, 0.5],
            "COSMIC Driver": ["", "", ""],
            "COSMIC Entities": ["", "", ""],
            "Paed Driver": ["", "", ""],
            "Paed Entities": ["", "", ""],
            "Sarc Driver": ["", "", ""],
            "Sarc Entities": ["", "", ""],
            "Neuro Driver": ["", "", ""],
            "Neuro Entities": ["", "", ""],
            "Ovary Driver": ["", "", ""],
            "Ovary Entities": ["", "", ""],
            "Haem Driver": ["", "", ""],
            "Haem Entities": ["", "", ""],
        }
    )

    yield test_input
    del test_input


@pytest.fixture()
def hotspots():
    hotspots = {
        "HS_Samples": pd.DataFrame(
            {
                "Gene_AA": ["gene1:L40", "gene3:L2135", "gene4:D23"],
                "Total": ["total1", "total2", "total3"],
                "Mutations": ["mutation1", "mutation2", "mutation3"],
            }
        ),
        "HS_Tissue": pd.DataFrame(
            {
                "Gene_Mut": ["gene1:L40R", "gene3:L2135V", "gene4:D23E"],
                "Tissue": ["tissue1", "tissue2", "tissue3"],
            }
        ),
    }

    yield hotspots
    del hotspots


@pytest.fixture()
def cyto():
    cyto = {
        "Cyto": pd.DataFrame(
            {
                "Gene": ["gene1", "gene3", "gene4"],
                "Cyto": ["cyto1", "cyto2", "cyto3"],
            }
        )
    }

    yield cyto
    del cyto


@pytest.fixture()
def sv_variant_data():
    data = pd.DataFrame(
        {
            "Event domain": ["domain1", "domain2", "domain3"],
            "Impacted transcript region": ["region1", "region2", "region3"],
            "Gene": ["gene1", "gene2", "gene3"],
            "GRCh38 coordinates": ["coor1", "coor2", "coor3"],
            "RefSeq IDs": ["refseq_id1", "refseq_id2", "refseq_id3"],
            "Type": ["GAIN(1)", "GAIN(3)", "LOSS(2)"],
            "Size": [10, 20, 30],
            "Chromosomal bands": ["cyto1;cyto2", "cyto3;cyto4", "cyto5;cyto6"],
            "Gene mode of action": ["mode1", "mode2", "mode3"],
            "COSMIC Driver": ["", "", ""],
            "COSMIC Entities": ["", "", ""],
            "Paed Driver": ["", "", ""],
            "Paed Entities": ["", "", ""],
            "Sarc Driver": ["", "", ""],
            "Sarc Entities": ["", "", ""],
            "Neuro Driver": ["", "", ""],
            "Neuro Entities": ["", "", ""],
            "Ovary Driver": ["", "", ""],
            "Ovary Entities": ["", "", ""],
            "Haem Driver": ["", "", ""],
            "Haem Entities": ["", "", ""],
        }
    )

    yield data
    del data


@pytest.fixture()
def fusion_data():
    fusion_data = pd.DataFrame(
        {
            "Event domain": ["domain1", "domain2", "domain3"],
            "Impacted transcript region": [
                "region1",
                "region2",
                "region3",
            ],
            "GRCh38 coordinates": ["coor1", "coor2", "coor3"],
            "Chromosomal bands": ["cyto1", "cyto2", "cyto3"],
            "RefSeq IDs": ["refseq_id1", "refseq_id2", "refseq_id3"],
            (
                "Population germline allele frequency (GESG | GECG for "
                "somatic SVs or AF | AUC for germline CNVs)"
            ): ["freq1", "freq2", "freq3"],
            "Gene mode of action": ["mode1", "mode2", "mode3"],
            "Gene": ["gene1", "gene2;gene3", "gene4;gene5;gene6"],
            "Type": ["GAIN(1)", "type1;type2", "type3;type4;type5"],
            "Confidence/support": [
                "PR-0/219;SR-16/216",
                "PR-7/133",
                "PR-1/69;SR-18/95",
            ],
            "Size": [10000, 200000, 300000],
        }
    )

    yield fusion_data
    del fusion_data


@pytest.fixture()
def refgene_data():
    refgene_df = pd.DataFrame(
        {
            "Gene": ["gene1", "gene2", "gene3"],
            "Alteration": ["alt1", "alt2", ""],
            "Entities": ["ent1", "ent2", ""],
            "Paed_Alteration": ["paed_alt1", "", "paed_alt2"],
            "Paed_Entities": ["paed_ent1", "", "paed_alt2"],
            "Sarcoma_Alteration": ["", "sarc_alt1", "sarc_alt2"],
            "Sarcoma_Entities": ["", "sarc_ent1", "sarc_alt2"],
            "Neuro_Alteration": ["", "", "neuro_alt1"],
            "Neuro_Entities": ["", "", "neuro_ent1"],
            "Ovarian_Alteration": ["", "", ""],
            "Ovarian_Entities": ["", "", ""],
            "Haem_Alteration": ["haem_alt1", "haem_alt2", "haem_alt3"],
            "Haem_Entities": ["haem_ent1", "haem_ent2", "haem_ent3"],
        }
    )

    yield refgene_df
    del refgene_df


class TestProcessReportedVariantsGermline:
    @pytest.mark.parametrize(
        "test_input", [{}, {"Origin": ["somatic"], "Data": ["data1"]}]
    )
    def test_process_no_germline(self, test_input):
        test_inputs = [pd.DataFrame(test_input), None, None]

        assert (
            excel_parsing.process_reported_variants_germline(*test_inputs)
            is None
        )

    @patch("utils.excel_parsing.vcf.find_clinvar_info")
    def test_process_single_row(
        self, mock_vcf_data, germline_variant_data, panelapp_dfs
    ):
        mock_vcf_data.return_value = pd.DataFrame(
            {
                "ClinVar ID": ["1"],
                "clnsigconf": ["sig1"],
            }
        )

        test_output = excel_parsing.process_reported_variants_germline(
            germline_variant_data.iloc[[0], :], "", panelapp_dfs
        )

        expected_output = pd.DataFrame(
            {
                "Gene": ["gene1"],
                "GRCh38 coordinates;ref/alt allele": ["coor1"],
                "CDS change and protein change": ["c.1"],
                "Predicted consequences": ["consequence1"],
                "Genotype": ["0/1"],
                "Population germline allele frequency (GE | gnomAD)": [
                    "ge1|freq1"
                ],
                "Gene mode of action": ["deletion1"],
                "clnsigconf": ["sig1"],
                "Tumour VAF": [""],
                "PanelApp Adult_v2.2": ["mode1"],
                "PanelApp Childhood_v4.0": ["mode2"],
            }
        )

        assert test_output.equals(expected_output)

    @patch("utils.excel_parsing.vcf.find_clinvar_info")
    def test_process_multiple_rows(
        self, mock_vcf_data, germline_variant_data, panelapp_dfs
    ):
        mock_vcf_data.return_value = pd.DataFrame(
            {
                "ClinVar ID": ["1", "2"],
                "clnsigconf": ["sig1", "sig2"],
            }
        )

        test_output = excel_parsing.process_reported_variants_germline(
            germline_variant_data, "", panelapp_dfs
        )

        expected_output = pd.DataFrame(
            {
                "Gene": ["gene1", "gene3"],
                "GRCh38 coordinates;ref/alt allele": ["coor1", "coor3"],
                "CDS change and protein change": ["c.1", "c.3"],
                "Predicted consequences": ["consequence1", "consequence3"],
                "Genotype": ["0/1", "1|1"],
                "Population germline allele frequency (GE | gnomAD)": [
                    "ge1|freq1",
                    "ge2|freq2",
                ],
                "Gene mode of action": ["deletion1", "deletion2"],
                "clnsigconf": ["sig1", "sig2"],
                "Tumour VAF": ["", ""],
                "PanelApp Adult_v2.2": ["mode1", "mode3"],
                "PanelApp Childhood_v4.0": ["mode2", "mode4"],
            }
        )

        assert test_output.equals(expected_output)


class TestProcessReportedVariantsSomatic:
    @pytest.mark.parametrize(
        "test_input", [{}, {"Origin": ["germline"], "Data": ["data1"]}]
    )
    def test_process_no_somatic(self, test_input, hotspots, cyto):
        test_inputs = [
            pd.DataFrame(test_input),
            tuple(),
            hotspots,
            cyto,
        ]

        assert (
            excel_parsing.process_reported_variants_somatic(*test_inputs)
            is None
        )

    def test_process_single_row(self, somatic_variant_data, hotspots, cyto):
        test_output = excel_parsing.process_reported_variants_somatic(
            somatic_variant_data.iloc[[0], :],
            (),
            hotspots,
            cyto,
        )

        expected_output = pd.DataFrame(
            {
                "Domain": ["domain1"],
                "Gene": ["gene1"],
                "GRCh38 coordinates": ["coor1"],
                "Cyto": ["cyto1"],
                "RefSeq IDs": ["refseq_id1"],
                "CDS change and protein change": ["c.1;p.Leu40Arg"],
                "Predicted consequences": ["consequence1"],
                "Error flag": [""],
                "Population germline allele frequency (GE | gnomAD)": [
                    "0.1 | 0.2"
                ],
                "VAF": [0.3],
                "LOH": ["0.1"],
                "Alt allele/total read depth": ["depth1"],
                "Gene mode of action": ["mode1"],
                "Variant class": [""],
                "TSG_NMD": [""],
                "TSG_LOH": [""],
                "Splice fs?": [""],
                "SpliceAI": [""],
                "REVEL": [""],
                "OG_3' Ter": [""],
                "Recurrence somatic database": [""],
                "HS_Total": ["total1"],
                "HS_Mut": ["mutation1"],
                "HS_Tissue": ["tissue1"],
                "COSMIC Driver": [""],
                "COSMIC Entities": [""],
                "Paed Driver": [""],
                "Paed Entities": [""],
                "Sarc Driver": [""],
                "Sarc Entities": [""],
                "Neuro Driver": [""],
                "Neuro Entities": [""],
                "Ovary Driver": [""],
                "Ovary Entities": [""],
                "Haem Driver": [""],
                "Haem Entities": [""],
                "MTBP c.": ["gene1:c.1"],
                "MTBP p.": ["gene1:L40R"],
            }
        )

        assert test_output.equals(expected_output)

    def test_process_multiple_rows(self, somatic_variant_data, hotspots, cyto):
        test_output = excel_parsing.process_reported_variants_somatic(
            somatic_variant_data, (), hotspots, cyto
        )

        expected_output = pd.DataFrame(
            {
                "Domain": ["domain1", "domain3"],
                "Gene": ["gene1", "gene3"],
                "GRCh38 coordinates": ["coor1", "coor3"],
                "Cyto": ["cyto1", "cyto2"],
                "RefSeq IDs": ["refseq_id1", "refseq_id3"],
                "CDS change and protein change": [
                    "c.1;p.Leu40Arg",
                    "c.3;p.Leu2135Val",
                ],
                "Predicted consequences": ["consequence1", "consequence3"],
                "Error flag": ["", "consequence4"],
                "Population germline allele frequency (GE | gnomAD)": [
                    "0.1 | 0.2",
                    "- | -",
                ],
                "VAF": [0.3, 0.5],
                "LOH": ["0.1", ""],
                "Alt allele/total read depth": ["depth1", "depth3"],
                "Gene mode of action": ["mode1", "mode3"],
                "Variant class": ["", ""],
                "TSG_NMD": ["", ""],
                "TSG_LOH": ["", ""],
                "Splice fs?": ["", ""],
                "SpliceAI": ["", ""],
                "REVEL": ["", ""],
                "OG_3' Ter": ["", ""],
                "Recurrence somatic database": ["", ""],
                "HS_Total": ["total1", "total2"],
                "HS_Mut": ["mutation1", "mutation2"],
                "HS_Tissue": ["tissue1", "tissue2"],
                "COSMIC Driver": ["", ""],
                "COSMIC Entities": ["", ""],
                "Paed Driver": ["", ""],
                "Paed Entities": ["", ""],
                "Sarc Driver": ["", ""],
                "Sarc Entities": ["", ""],
                "Neuro Driver": ["", ""],
                "Neuro Entities": ["", ""],
                "Ovary Driver": ["", ""],
                "Ovary Entities": ["", ""],
                "Haem Driver": ["", ""],
                "Haem Entities": ["", ""],
                "MTBP c.": ["gene1:c.1", "gene3:c.3"],
                "MTBP p.": ["gene1:L40R", "gene3:L2135V"],
            }
        )

        assert test_output.equals(expected_output)


class TestProcessReportedSV:
    @pytest.mark.parametrize(
        "test_input",
        [
            {},
            {"Type": ["not_right_type"], "Data": ["data1"]},
            {"Data": ["data1"]},
        ],
    )
    def test_no_data(self, test_input):
        test_inputs = [pd.DataFrame(test_input), (), "right_type"]
        assert excel_parsing.process_reported_SV(*test_inputs) is None

    def test_process_single_row_gain(self, sv_variant_data):
        test_output = excel_parsing.process_reported_SV(
            sv_variant_data.iloc[[0], :],
            (),
            "gain",
            "new_column1",
            "new_column2",
        )

        expected_output = pd.DataFrame(
            {
                "Event domain": ["domain1"],
                "Gene": ["gene1"],
                "RefSeq IDs": ["refseq_id1"],
                "Impacted transcript region": ["region1"],
                "GRCh38 coordinates": ["coor1"],
                "Type": ["GAIN"],
                "Copy Number": [1],
                "Size": ["10"],
                "Cyto 1": ["cyto1"],
                "Cyto 2": ["cyto2"],
                "Gene mode of action": ["mode1"],
                "Variant class": [""],
                "new_column1": [""],
                "new_column2": [""],
                "COSMIC Driver": [""],
                "COSMIC Entities": [""],
                "Paed Driver": [""],
                "Paed Entities": [""],
                "Sarc Driver": [""],
                "Sarc Entities": [""],
                "Neuro Driver": [""],
                "Neuro Entities": [""],
                "Ovary Driver": [""],
                "Ovary Entities": [""],
                "Haem Driver": [""],
                "Haem Entities": [""],
            }
        )

        assert test_output.equals(expected_output)

    def test_process_single_row_loss(self, sv_variant_data):
        test_output = excel_parsing.process_reported_SV(
            sv_variant_data.iloc[[2], :],
            (),
            "loss",
            "new_column1",
        )

        expected_output = pd.DataFrame(
            {
                "Event domain": ["domain3"],
                "Gene": ["gene3"],
                "RefSeq IDs": ["refseq_id3"],
                "Impacted transcript region": ["region3"],
                "GRCh38 coordinates": ["coor3"],
                "Type": ["LOSS"],
                "Copy Number": [2],
                "Size": ["30"],
                "Cyto 1": ["cyto5"],
                "Cyto 2": ["cyto6"],
                "Gene mode of action": ["mode3"],
                "Variant class": [""],
                "new_column1": [""],
                "COSMIC Driver": [""],
                "COSMIC Entities": [""],
                "Paed Driver": [""],
                "Paed Entities": [""],
                "Sarc Driver": [""],
                "Sarc Entities": [""],
                "Neuro Driver": [""],
                "Neuro Entities": [""],
                "Ovary Driver": [""],
                "Ovary Entities": [""],
                "Haem Driver": [""],
                "Haem Entities": [""],
            }
        )

        assert test_output.equals(expected_output)

    def test_process_multiple_rows_gain(self, sv_variant_data):
        test_output = excel_parsing.process_reported_SV(
            sv_variant_data,
            (),
            "gain",
            "new_column1",
            "new_column2",
            "new_column3",
        )

        expected_output = pd.DataFrame(
            {
                "Event domain": ["domain1", "domain2"],
                "Gene": ["gene1", "gene2"],
                "RefSeq IDs": ["refseq_id1", "refseq_id2"],
                "Impacted transcript region": ["region1", "region2"],
                "GRCh38 coordinates": ["coor1", "coor2"],
                "Type": ["GAIN", "GAIN"],
                "Copy Number": [1, 3],
                "Size": ["10", "20"],
                "Cyto 1": ["cyto1", "cyto3"],
                "Cyto 2": ["cyto2", "cyto4"],
                "Gene mode of action": ["mode1", "mode2"],
                "Variant class": ["", ""],
                "new_column1": ["", ""],
                "new_column2": ["", ""],
                "new_column3": ["", ""],
                "COSMIC Driver": ["", ""],
                "COSMIC Entities": ["", ""],
                "Paed Driver": ["", ""],
                "Paed Entities": ["", ""],
                "Sarc Driver": ["", ""],
                "Sarc Entities": ["", ""],
                "Neuro Driver": ["", ""],
                "Neuro Entities": ["", ""],
                "Ovary Driver": ["", ""],
                "Ovary Entities": ["", ""],
                "Haem Driver": ["", ""],
                "Haem Entities": ["", ""],
            }
        )

        assert test_output.equals(expected_output)


class TestProcessFusion:
    @pytest.mark.parametrize("test_input", [{}, {"Data": ["data1"]}])
    def test_no_data(self, test_input, cyto):
        test_inputs = [pd.DataFrame(test_input), (), cyto]
        assert excel_parsing.process_fusion_SV(*test_inputs) is None

    def test_single_row(self, fusion_data, cyto):
        test_df_output, test_fusion_output, test_alternative_columns = (
            excel_parsing.process_fusion_SV(fusion_data.iloc[[1], :], (), cyto)
        )

        expected_df = pd.DataFrame(
            {
                "Event domain": ["domain2"],
                "Gene": ["gene2;gene3"],
                "RefSeq IDs": ["refseq_id2"],
                "Impacted transcript region": ["region2"],
                "GRCh38 coordinates": ["coor2"],
                "Type": ["type1"],
                "Fusion_1": ["type2"],
                "Size": ["200,000"],
                (
                    "Population germline allele frequency (GESG | GECG for "
                    "somatic SVs or AF | AUC for germline CNVs)"
                ): ["freq2"],
                "Paired reads": ["7/133"],
                "Split reads": [""],
                "Cyto\nGene_1": ["-"],
                "Cyto\nGene_2": ["cyto2"],
                "Gene mode of action": ["mode2"],
                "Variant class": [""],
                "OG_Fusion": [""],
                "OG_IntDup": [""],
                "OG_IntDel": [""],
                "Disruptive": [""],
            }
        )

        assert (
            test_df_output.equals(expected_df)
            and test_fusion_output == 1
            and test_alternative_columns == {}
        )

    def test_multiple_rows(self, fusion_data, cyto):
        test_df_output, test_fusion_output, test_alternative_columns = (
            excel_parsing.process_fusion_SV(fusion_data, (), cyto)
        )

        expected_df = pd.DataFrame(
            {
                "Event domain": ["domain2", "domain3"],
                "Gene": ["gene2;gene3", "gene4;gene5;gene6"],
                "RefSeq IDs": ["refseq_id2", "refseq_id3"],
                "Impacted transcript region": ["region2", "region3"],
                "GRCh38 coordinates": ["coor2", "coor3"],
                "Type": ["type1", "type3"],
                "Fusion_1": ["type2", "type4"],
                "Fusion_2": [None, "type5"],
                "Size": ["200,000", "300,000"],
                (
                    "Population germline allele frequency (GESG | GECG for "
                    "somatic SVs or AF | AUC for germline CNVs)"
                ): ["freq2", "freq3"],
                "Paired reads": ["7/133", "1/69"],
                "Split reads": ["", "18/95"],
                "Cyto\nGene_1": ["-", "cyto3"],
                "Cyto\nGene_2": ["cyto2", "-"],
                "Cyto\nGene_3": ["-", "-"],
                "Gene mode of action": ["mode2", "mode3"],
                "Variant class": ["", ""],
                "OG_Fusion": ["", ""],
                "OG_IntDup": ["", ""],
                "OG_IntDel": ["", ""],
                "Disruptive": ["", ""],
            }
        )

        assert (
            test_df_output.equals(expected_df)
            and test_fusion_output == 2
            and test_alternative_columns == {}
        )


class TestProcessRefgene:
    def test_process_data(self):
        test_output = excel_parsing.process_refgene(
            {
                "cosmic": pd.DataFrame(
                    {
                        "Gene": ["gene1", "gene3"],
                        "Role in Cancer": ["somatic_data1", "somatic_data4"],
                        "Driver_SV": ["somatic_data2", "somatic_data5"],
                        "Entities": ["somatic_data3", "somatic_data6"],
                    }
                ),
                "haem": pd.DataFrame(
                    {
                        "Gene": ["gene2"],
                        "Driver": ["haem_data1"],
                        "Entities": ["haem_data2"],
                        "Comments": ["haem_data3"],
                    }
                ),
                "paed": pd.DataFrame(
                    {
                        "Gene": ["gene1"],
                        "Driver": ["paed_data1"],
                        "Entities": ["paed_data2"],
                        "Comments": ["paed_data3"],
                    }
                ),
                "ovarian": pd.DataFrame(
                    {
                        "Gene": ["gene3"],
                        "Driver": ["ovarian_data1"],
                        "Entities": ["ovarian_data2"],
                        "Comments": ["ovarian_data3"],
                    }
                ),
                "sarc": pd.DataFrame(
                    {
                        "Gene": ["gene4"],
                        "Driver": ["sarc_data1"],
                        "Entities": ["sarc_data2"],
                        "Comments": ["sarc_data3"],
                    }
                ),
                "neuro": pd.DataFrame(
                    {
                        "Gene": ["gene2", "gene1"],
                        "Driver": ["neuro_data1", "neuro_data5"],
                        "Entities": ["neuro_data2", "neuro_data6"],
                        "Comments": ["neuro_data3", "neuro_data7"],
                    }
                ),
            }
        )

        expected_output = pd.DataFrame(
            {
                "Gene": ["gene1", "gene2", "gene3", "gene4"],
                "Comments": ["somatic_data1", np.nan, "somatic_data4", np.nan],
                "COSMIC_Alteration": [
                    "somatic_data2",
                    np.nan,
                    "somatic_data5",
                    np.nan,
                ],
                "COSMIC_Entities": [
                    "somatic_data3",
                    np.nan,
                    "somatic_data6",
                    np.nan,
                ],
                "Haem_Alteration": [np.nan, "haem_data1", np.nan, np.nan],
                "Haem_Entities": [np.nan, "haem_data2", np.nan, np.nan],
                "Haem_Comments": [np.nan, "haem_data3", np.nan, np.nan],
                "Paed_Alteration": ["paed_data1", np.nan, np.nan, np.nan],
                "Paed_Entities": ["paed_data2", np.nan, np.nan, np.nan],
                "Paed_Comments": ["paed_data3", np.nan, np.nan, np.nan],
                "Ovarian_Alteration": [
                    np.nan,
                    np.nan,
                    "ovarian_data1",
                    np.nan,
                ],
                "Ovarian_Entities": [np.nan, np.nan, "ovarian_data2", np.nan],
                "Ovarian_Comments": [np.nan, np.nan, "ovarian_data3", np.nan],
                "Sarcoma_Alteration": [np.nan, np.nan, np.nan, "sarc_data1"],
                "Sarcoma_Entites": [np.nan, np.nan, np.nan, "sarc_data2"],
                "Sarcoma_Comments": [np.nan, np.nan, np.nan, "sarc_data3"],
                "Neuro_Alteration": [
                    "neuro_data5",
                    "neuro_data1",
                    np.nan,
                    np.nan,
                ],
                "Neuro_Entities": [
                    "neuro_data6",
                    "neuro_data2",
                    np.nan,
                    np.nan,
                ],
                "Neuro_Comments": [
                    "neuro_data7",
                    "neuro_data3",
                    np.nan,
                    np.nan,
                ],
            }
        )

        assert test_output.equals(expected_output)


class TestLookupDataFromVariants:
    def test_process_data(self, refgene_data):
        test_output = excel_parsing.lookup_data_from_variants(
            refgene_data,
            **{
                "somatic": pd.DataFrame(
                    {
                        "Gene": ["gene1", "gene4"],
                        "CDS change and protein change": ["data1", "data4"],
                    }
                ),
                "gain": pd.DataFrame(
                    {"Gene": ["gene3", "gene4"], "Copy Number": [3, 4]}
                ),
                "loss": pd.DataFrame(
                    {"Gene": ["gene2", "gene4"], "Copy Number": [1, 3]}
                ),
                "fusion": pd.DataFrame(
                    {"Gene": ["gene1;gene2"], "Type": ["type1"]}
                ),
            },
        )

        expected_output = pd.DataFrame(
            {
                "Gene": ["gene1", "gene2", "gene3"],
                "Alteration": ["alt1", "alt2", ""],
                "Entities": ["ent1", "ent2", ""],
                "Paed_Alteration": ["paed_alt1", "", "paed_alt2"],
                "Paed_Entities": ["paed_ent1", "", "paed_alt2"],
                "Sarcoma_Alteration": ["", "sarc_alt1", "sarc_alt2"],
                "Sarcoma_Entities": ["", "sarc_ent1", "sarc_alt2"],
                "Neuro_Alteration": ["", "", "neuro_alt1"],
                "Neuro_Entities": ["", "", "neuro_ent1"],
                "Ovarian_Alteration": ["", "", ""],
                "Ovarian_Entities": ["", "", ""],
                "Haem_Alteration": ["haem_alt1", "haem_alt2", "haem_alt3"],
                "Haem_Entities": ["haem_ent1", "haem_ent2", "haem_ent3"],
                "SNV": ["data1", "-", "-"],
                "CN": ["-", "1", "3"],
                "SV_gene_1": ["type1", "-", "-"],
                "SV_gene_2": ["-", "type1", "-"],
            }
        )

        assert test_output.equals(expected_output)
