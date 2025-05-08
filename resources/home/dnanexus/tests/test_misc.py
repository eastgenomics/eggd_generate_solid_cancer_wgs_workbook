from types import ModuleType

import pandas as pd
import pytest

from utils import misc


class TestSelectConfig:
    def test_config_not_found(self):
        test_output = misc.select_config("non existent config")
        assert test_output is None

    def test_config_found(self):
        test_output = misc.select_config("soc")
        assert type(test_output) is ModuleType


class TestMergeDicts:
    def test_both_dicts_empty(self):
        test_output = misc.merge_dicts({}, {})

        assert test_output == {}

    def test_one_dict_empty(self):
        test_output = misc.merge_dicts({"not_empty": "value"}, {})

        assert test_output == {"not_empty": "value"}

    def test_add_new_values_to_common_key(self):
        test_dict1 = {"cells_to_write": {(1, 1): "cell1"}}
        test_dict2 = {"cells_to_write": {(1, 2): "cell2"}}

        test_output = misc.merge_dicts(test_dict1, test_dict2)

        expected_output = {
            "cells_to_write": {(1, 1): "cell1", (1, 2): "cell2"}
        }

        assert test_output == expected_output

    def test_merge_lists_to_common_key(self):
        test_dict1 = {"to_align": ["cell1", "cell2"]}
        test_dict2 = {"to_align": ["cell3", "cell4"]}

        test_output = misc.merge_dicts(test_dict1, test_dict2)

        expected_output = {"to_align": ["cell1", "cell2", "cell3", "cell4"]}

        assert test_output == expected_output

    def test_not_common_keys(self):
        test_dict1 = {"to_align": ["cell1", "cell2"]}
        test_dict2 = {"cells_to_write": {(1, 1): "cell1", (1, 2): "cell2"}}

        test_output = misc.merge_dicts(test_dict1, test_dict2)

        expected_output = {
            "to_align": ["cell1", "cell2"],
            "cells_to_write": {(1, 1): "cell1", (1, 2): "cell2"},
        }

        assert test_output == expected_output

    def test_merge_nested_dicts(self):
        test_dict1 = {"borders": {"single_cells": ["cell1", "cell2"]}}
        test_dict2 = {"borders": {"single_cells": ["cell3", "cell4"]}}

        test_output = misc.merge_dicts(test_dict1, test_dict2)

        expected_output = {
            "borders": {"single_cells": ["cell1", "cell2", "cell3", "cell4"]}
        }

        assert test_output == expected_output

    def test_merge_not_list_or_dict(self):
        test_dict1 = {"freeze_panes": "cell1"}
        test_dict2 = {"freeze_panes": "cell2"}

        test_output = misc.merge_dicts(test_dict1, test_dict2)

        expected_output = {"freeze_panes": "cell2"}

        assert test_output == expected_output


class TestSplitConfidenceSupport:
    def test_PR_only(self):
        test_input = "PR-1"

        test_output = misc.split_confidence_support(test_input)
        expected_output = ["1", ""]

        assert test_output == expected_output

    def test_SR_only(self):
        test_input = "SR-1"

        test_output = misc.split_confidence_support(test_input)
        expected_output = ["", "1"]

        assert test_output == expected_output

    def test_PR_and_SR(self):
        test_input = "PR-1;SR-2"

        test_output = misc.split_confidence_support(test_input)
        expected_output = ["1", "2"]

        assert test_output == expected_output

    def test_PR_and_SR_wrong_order(self):
        test_input = "SR-2;PR-1"

        test_output = misc.split_confidence_support(test_input)
        expected_output = ["1", "2"]

        assert test_output == expected_output


@pytest.fixture()
def df_with_many_columns():
    test_input = pd.DataFrame(
        columns=[f"col{col_nb}" for col_nb in range(1, 500)]
    )
    yield test_input
    del test_input


class TestGetColumnLetterUsingColumnName:
    @pytest.mark.parametrize(
        "test_input, expected",
        [("col7", "G"), ("col29", "AC"), ("col400", "OJ"), ("col26", "Z")],
    )
    def test_given_column_name(
        self, test_input, expected, df_with_many_columns
    ):
        test_output = misc.get_column_letter_using_column_name(
            df_with_many_columns, test_input
        )

        assert test_output == expected

    def test_no_given_column(self, df_with_many_columns):
        test_output = misc.get_column_letter_using_column_name(
            df_with_many_columns
        )

        assert test_output == "SE"

    def test_outside_of_possibility(self):
        df_with_too_many_columns = pd.DataFrame(
            columns=[f"col{col_nb}" for col_nb in range(1, 730)]
        )

        with pytest.raises(ValueError):
            misc.get_column_letter_using_column_name(df_with_too_many_columns)


class TestConvertLetterColumnToIndex:
    @pytest.mark.parametrize(
        "test_input, expected",
        [("G", 6), ("AC", 28), ("Z", 25), ("OJ", 399)],
    )
    def test_normal_cases(self, test_input, expected):
        assert misc.convert_letter_column_to_index(test_input) == expected

    def test_3_letter_column(self):
        with pytest.raises(ValueError):
            misc.convert_letter_column_to_index("AAA")


class TestConvertIndexToLetters:
    @pytest.mark.parametrize(
        "test_input, expected",
        [(6, "G"), (28, "AC"), (25, "Z"), (399, "OJ")],
    )
    def test_normal_cases(self, test_input, expected):
        assert misc.convert_index_to_letters(test_input) == expected

    def test_outside_range(self):
        with pytest.raises(ValueError):
            misc.convert_index_to_letters(1000)


class TestConvert3LetterProteinTo1:
    @pytest.mark.parametrize(
        "test_input, expected",
        [("Glu", "E"), ("Pro", "P"), ("Trp", "W"), ("Asn", "N")],
    )
    def test_string_in_mapping(self, test_input, expected):
        assert misc.convert_3_letter_protein_to_1(test_input) == expected

    @pytest.mark.parametrize(
        "test_input, expected",
        [("GLU", "GLU"), ("Blarg", "Blarg"), ("", ""), ("1", "1")],
    )
    def test_string_not_in_mapping(self, test_input, expected):
        assert misc.convert_3_letter_protein_to_1(test_input) == expected

    @pytest.mark.parametrize(
        "test_input, expected",
        [
            (["Glu"], ["Glu"]),
            ({"Glu": "Glu"}, {"Glu": "Glu"}),
            (None, None),
            (True, True),
        ],
    )
    def test_input_is_not_string(self, test_input, expected):
        assert misc.convert_3_letter_protein_to_1(test_input) == expected
