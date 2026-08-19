import os
import sys

import pandas as pd
import pytest
from unittest.mock import MagicMock, patch

sys.path.insert(0, "../src")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

# Import the module to test
from dataprocess import (
    ReturnCodes,
    get_available_periods,
    prepare_department_fte_trend_report,
    prepare_department_headcount_trend_report,
    prepare_department_fte_costcentre_report,
    PDF_Generator_report,
    report_css_style,
    MAX_NUMBER_MONTH_IN_REPORT,
)


class TestReturnCodes:
    """Test cases for ReturnCodes enum"""

    def test_return_codes_values(self):
        """Test that all return codes have expected values"""
        assert ReturnCodes.ERROR_PROGRAM.value == -10
        assert ReturnCodes.ERROR_INPUT.value == -4
        assert ReturnCodes.ERROR_FILE_DATA.value == -4
        assert ReturnCodes.ERROR_FILE_LOADING.value == -2
        assert ReturnCodes.ERROR_FILE.value == -1
        assert ReturnCodes.ERROR.value == 0
        assert ReturnCodes.OK.value == 1
        assert ReturnCodes.OK_GEN_NEW_DATABASE.value == 2
        assert ReturnCodes.OK_UPDATE_DATABASE.value == 3

    def test_return_codes_membership(self):
        """Test that return codes are enum members"""
        assert isinstance(ReturnCodes.ERROR_PROGRAM, ReturnCodes)
        assert isinstance(ReturnCodes.OK, ReturnCodes)


class TestGetAvailablePeriods:
    """Test cases for get_available_periods function.

    The function now returns a dict:
        {"available_periods": list, "return_code": int}
    """

    def test_basic_functionality(self):
        """Test basic period retrieval"""
        data_available = ["202301", "202302", "202303", "202304"]
        result = get_available_periods(data_available, 2023, 1, 4)
        assert result["return_code"] == ReturnCodes.OK.value
        assert result["available_periods"] == ["202301", "202302", "202303", "202304"]

    def test_partial_data_available(self):
        """Test when only some periods are available"""
        data_available = ["202301", "202303"]
        result = get_available_periods(data_available, 2023, 1, 4)
        assert result["available_periods"] == ["202301", "202303"]
        assert len(result["available_periods"]) == 2

    def test_no_data_available(self):
        """Test when no data is available"""
        result = get_available_periods([], 2023, 1, 3)
        assert result["available_periods"] == []
        assert result["return_code"] == ReturnCodes.OK.value

    def test_year_rollover(self):
        """Test period generation across year boundary"""
        data_available = ["202311", "202312", "202401", "202402"]
        result = get_available_periods(data_available, 2023, 11, 4)
        assert result["available_periods"] == ["202311", "202312", "202401", "202402"]

    def test_december_to_january_transition(self):
        """Test month 12 to month 1 transition"""
        data_available = ["202312", "202401"]
        result = get_available_periods(data_available, 2023, 12, 2)
        assert result["available_periods"] == ["202312", "202401"]

    def test_invalid_max_number_of_months_zero(self):
        """Test error when max_number_of_month is 0"""
        result = get_available_periods(["202301"], 2023, 1, 0)
        assert result["return_code"] == ReturnCodes.ERROR_PROGRAM.value
        assert result["available_periods"] == []

    def test_valid_max_number_of_months_one(self):
        """Test that max_number_of_month of 1 is valid"""
        result = get_available_periods(["202301"], 2023, 1, 1)
        assert result["return_code"] == ReturnCodes.OK.value
        assert result["available_periods"] == ["202301"]

    def test_invalid_start_month_zero(self):
        """Test error when start_month is 0"""
        result = get_available_periods(["202301"], 2023, 0, 3)
        assert result["return_code"] == ReturnCodes.ERROR_PROGRAM.value

    def test_invalid_start_month_thirteen(self):
        """Test error when start_month is 13"""
        result = get_available_periods(["202301"], 2023, 13, 3)
        assert result["return_code"] == ReturnCodes.ERROR_PROGRAM.value

    def test_invalid_start_month_negative(self):
        """Test error when start_month is negative"""
        result = get_available_periods(["202301"], 2023, -1, 3)
        assert result["return_code"] == ReturnCodes.ERROR_PROGRAM.value

    def test_invalid_max_number_of_months_over_twelve(self):
        """Test error when max_number_of_month exceeds 12"""
        result = get_available_periods(["202301"], 2023, 1, 24)
        assert result["return_code"] == ReturnCodes.ERROR_PROGRAM.value

    def test_invalid_start_year(self):
        """Test error when start_year is out of range"""
        assert (
            get_available_periods(["202301"], 1999, 1, 3)["return_code"]
            == ReturnCodes.ERROR_PROGRAM.value
        )
        assert (
            get_available_periods(["202301"], 3001, 1, 3)["return_code"]
            == ReturnCodes.ERROR_PROGRAM.value
        )

    def test_max_periods(self):
        """Test with maximum number of months"""
        data_available = [f"2023{str(i).zfill(2)}" for i in range(1, 13)]
        result = get_available_periods(data_available, 2023, 1, 12)
        assert len(result["available_periods"]) == 12
        assert result["available_periods"][0] == "202301"
        assert result["available_periods"][-1] == "202312"

    def test_period_format_single_digit_month(self):
        """Test that single-digit months are zero-padded"""
        data_available = ["202301", "202302"]
        result = get_available_periods(data_available, 2023, 1, 2)
        assert result["available_periods"][0] == "202301"  # Not '20231'

    def test_data_not_in_sequence(self):
        """Test when available data is not in sequence"""
        data_available = ["202301", "202305", "202309"]
        result = get_available_periods(data_available, 2023, 1, 12)
        assert result["available_periods"] == ["202301", "202305", "202309"]


class TestPrepareDepartmentFTETrendReport:
    """Test cases for prepare_department_fte_trend_report function.

    On success returns:
        {"md": [...], "excel_df": {...}, "return_code": ReturnCodes.OK.value}
    On error returns:
        {"return_code": <error code>}
    """

    @patch("dataprocess.pd.read_excel")
    def test_successful_fte_report_generation(self, mock_read_excel):
        """Test successful FTE report generation"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior", "Senior", "Junior"],
                "allocation": ["1.0", "0.5", "1.0", "0.5"],
                "staff category order": [1, 2, 1, 2],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_fte_trend_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.OK.value
        assert isinstance(result["md"], list)
        assert len(result["md"]) == 1
        assert "content" in result["md"][0]
        assert "css" in result["md"][0]
        assert isinstance(result["md"][0]["content"], str)
        assert isinstance(result["md"][0]["css"], str)
        assert "excel_df" in result

    @patch("dataprocess.pd.read_excel")
    def test_file_loading_error(self, mock_read_excel):
        """Test error handling when file cannot be loaded"""
        mock_read_excel.side_effect = Exception("File not found")

        result = prepare_department_fte_trend_report("nonexistent.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.ERROR_FILE_LOADING.value

    @patch("dataprocess.pd.read_excel")
    def test_no_available_periods(self, mock_read_excel):
        """Test when no periods are available in data"""
        mock_read_excel.return_value = {"202305": pd.DataFrame()}

        result = prepare_department_fte_trend_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.ERROR_FILE_DATA.value

    @patch("dataprocess.pd.read_excel")
    def test_multiple_periods(self, mock_read_excel):
        """Test with multiple periods"""
        mock_df1 = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior"],
                "allocation": ["1.0", "0.5"],
                "staff category order": [1, 2],
            }
        )
        mock_df2 = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior"],
                "allocation": ["1.5", "0.8"],
                "staff category order": [1, 2],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df1, "202302": mock_df2}

        result = prepare_department_fte_trend_report("test.xlsx", 2023, 1, 3)

        assert result["return_code"] == ReturnCodes.OK.value
        content = result["md"][0]["content"]
        assert "202301" in content or "202302" in content

    @patch("dataprocess.pd.read_excel")
    def test_css_present(self, mock_read_excel):
        """Test that css is generated for the report"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior"],
                "allocation": ["1.0"],
                "staff category order": [1],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_fte_trend_report("test.xlsx", 2023, 1, 2)

        assert result["md"][0]["css"] == report_css_style()
        assert "table" in result["md"][0]["css"]


class TestPrepareDepartmentHeadcountTrendReport:
    """Test cases for prepare_department_headcount_trend_report function"""

    @patch("dataprocess.pd.read_excel")
    def test_successful_headcount_report_generation(self, mock_read_excel):
        """Test successful headcount report generation"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior", "Senior", "Junior"],
                "staff_number": ["001", "002", "003", "004"],
                "staff category order": [1, 2, 1, 2],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_headcount_trend_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.OK.value
        assert isinstance(result["md"], list)
        assert len(result["md"]) == 1
        assert "content" in result["md"][0]
        assert "css" in result["md"][0]

    @patch("dataprocess.pd.read_excel")
    def test_headcount_file_loading_error(self, mock_read_excel):
        """Test error handling when file cannot be loaded"""
        mock_read_excel.side_effect = Exception("File not found")

        result = prepare_department_headcount_trend_report(
            "nonexistent.xlsx", 2023, 1, 2
        )

        assert result["return_code"] == ReturnCodes.ERROR_FILE_LOADING.value

    @patch("dataprocess.pd.read_excel")
    def test_headcount_no_available_periods(self, mock_read_excel):
        """Test when no periods are available"""
        mock_read_excel.return_value = {"202305": pd.DataFrame()}

        result = prepare_department_headcount_trend_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.ERROR_FILE_DATA.value

    @patch("dataprocess.pd.read_excel")
    def test_duplicate_staff_numbers(self, mock_read_excel):
        """Test that duplicate staff numbers are handled correctly"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Senior", "Senior"],
                "staff_number": ["001", "001", "002"],  # Duplicate staff number
                "staff category order": [1, 1, 1],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_headcount_trend_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.OK.value
        assert "content" in result["md"][0]

    @patch("dataprocess.pd.read_excel")
    def test_headcount_css_present(self, mock_read_excel):
        """Test that the headcount report includes css"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior"],
                "staff_number": ["001"],
                "staff category order": [1],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_headcount_trend_report("test.xlsx", 2023, 1, 2)

        assert result["md"][0]["css"] == report_css_style()


class TestPrepareDepartmentFTECostcentreReport:
    """Test cases for prepare_department_fte_costcentre_report function"""

    @patch("dataprocess.pd.read_excel")
    def test_successful_costcentre_report(self, mock_read_excel):
        """Test successful cost centre report generation"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior"],
                "Rank": ["Manager", "Staff"],
                "allocation": ["1.0", "0.5"],
                "cost centre name": ["IT", "IT"],
                "cost centre code": ["001", "001"],
                "staff category order": [1, 2],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_fte_costcentre_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.OK.value
        assert isinstance(result["md"], list)
        assert len(result["md"]) > 0
        assert "content" in result["md"][0]
        assert "css" in result["md"][0]
        assert "Cost Centre" in result["md"][0]["content"]

    @patch("dataprocess.pd.read_excel")
    def test_multiple_cost_centres(self, mock_read_excel):
        """Test with multiple cost centres"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior", "Senior", "Junior"],
                "Rank": ["Manager", "Staff", "Manager", "Staff"],
                "allocation": ["1.0", "0.5", "1.5", "0.8"],
                "cost centre name": ["IT", "IT", "HR", "HR"],
                "cost centre code": ["001", "001", "002", "002"],
                "staff category order": [1, 2, 1, 2],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_fte_costcentre_report("test.xlsx", 2023, 1, 2)

        # Should have entries for both IT and HR
        assert result["return_code"] == ReturnCodes.OK.value
        assert len(result["md"]) == 2
        assert any("IT" in r["content"] for r in result["md"])
        assert any("HR" in r["content"] for r in result["md"])

    @patch("dataprocess.pd.read_excel")
    def test_costcentre_file_loading_error(self, mock_read_excel):
        """Test file loading error"""
        mock_read_excel.side_effect = Exception("File error")

        result = prepare_department_fte_costcentre_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.ERROR_FILE_LOADING.value

    @patch("dataprocess.pd.read_excel")
    def test_costcentre_no_available_periods(self, mock_read_excel):
        """Test when no periods match"""
        mock_read_excel.return_value = {"202305": pd.DataFrame()}

        result = prepare_department_fte_costcentre_report("test.xlsx", 2023, 1, 2)

        assert result["return_code"] == ReturnCodes.ERROR_FILE_DATA.value

    @patch("dataprocess.pd.read_excel")
    def test_costcentre_css_present(self, mock_read_excel):
        """Test that css is included in cost centre report"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior"],
                "Rank": ["Manager"],
                "allocation": ["1.0"],
                "cost centre name": ["IT"],
                "cost centre code": ["001"],
                "staff category order": [1],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_fte_costcentre_report("test.xlsx", 2023, 1, 2)

        assert result["md"][0]["css"] == report_css_style()


class TestPDFGeneratorReport:
    """Test cases for PDF_Generator_report function"""

    @patch("dataprocess.MarkdownPdf")
    def test_generate_pdf_basic(self, mock_markdown_pdf):
        """Test basic PDF generation"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance

        content = [{"content": "# Test", "css": "table {}"}]
        PDF_Generator_report("test_report", content, "Test Report")

        mock_markdown_pdf.assert_called_once()
        mock_pdf_instance.add_section.assert_called_once()
        mock_pdf_instance.save.assert_called_once_with("test_report.pdf")

    @patch("dataprocess.MarkdownPdf")
    def test_generate_pdf_with_title(self, mock_markdown_pdf):
        """Test PDF generation with custom title"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance

        content = [{"content": "## Content", "css": "h2 {}"}]
        PDF_Generator_report("report", content, "Custom Title")

        mock_pdf_instance.save.assert_called_once_with("report.pdf")

    @patch("dataprocess.MarkdownPdf")
    def test_generate_pdf_default_title(self, mock_markdown_pdf):
        """Test PDF generation using the default title"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance

        PDF_Generator_report("report", [{"content": "x", "css": "y"}])

        mock_pdf_instance.save.assert_called_once_with("report.pdf")

    @patch("dataprocess.MarkdownPdf")
    def test_generate_pdf_empty_content(self, mock_markdown_pdf):
        """Test PDF generation with empty content list"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance

        PDF_Generator_report("empty", [])

        mock_pdf_instance.add_section.assert_not_called()
        mock_pdf_instance.save.assert_called_once_with("empty.pdf")

    @patch("dataprocess.MarkdownPdf")
    def test_generate_pdf_multiple_sections(self, mock_markdown_pdf):
        """Test PDF generation with multiple content sections"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance

        content = [
            {"content": "# Section 1", "css": "h1 {}"},
            {"content": "# Section 2", "css": "h1 {}"},
            {"content": "# Section 3", "css": "h1 {}"},
        ]
        PDF_Generator_report("multi_section", content, "Multi Section Report")

        assert mock_pdf_instance.add_section.call_count == 3
        mock_pdf_instance.save.assert_called_once_with("multi_section.pdf")


class TestIntegration:
    """Integration tests for the module"""

    def test_max_number_month_constant(self):
        """Test that the constant is properly defined"""
        assert MAX_NUMBER_MONTH_IN_REPORT == 12

    @patch("dataprocess.pd.read_excel")
    def test_full_workflow_fte_report(self, mock_read_excel):
        """Test complete workflow for FTE report"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior"],
                "allocation": ["1.0", "0.5"],
                "staff category order": [1, 2],
            }
        )
        mock_read_excel.return_value = {
            "202301": mock_df,
            "202302": mock_df.copy(),
        }

        result = prepare_department_fte_trend_report(
            "test.xlsx", 2023, 1, MAX_NUMBER_MONTH_IN_REPORT
        )

        assert result["return_code"] == ReturnCodes.OK.value
        assert len(result["md"]) > 0

    @patch("dataprocess.pd.read_excel")
    def test_data_consistency_across_functions(self, mock_read_excel):
        """Test that different report functions handle same data consistently"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior"],
                "allocation": ["1.0", "0.5"],
                "staff_number": ["001", "002"],
                "Rank": ["Manager", "Staff"],
                "cost centre name": ["IT", "IT"],
                "cost centre code": ["001", "001"],
                "staff category order": [1, 2],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        fte_result = prepare_department_fte_trend_report("test.xlsx", 2023, 1, 2)
        headcount_result = prepare_department_headcount_trend_report(
            "test.xlsx", 2023, 1, 2
        )
        costcentre_result = prepare_department_fte_costcentre_report(
            "test.xlsx", 2023, 1, 2
        )

        # All should succeed
        assert fte_result["return_code"] == ReturnCodes.OK.value
        assert headcount_result["return_code"] == ReturnCodes.OK.value
        assert costcentre_result["return_code"] == ReturnCodes.OK.value


class TestEdgeCases:
    """Test edge cases and boundary conditions"""

    @patch("dataprocess.pd.read_excel")
    def test_nan_values_in_allocation(self, mock_read_excel):
        """Test handling of NaN values in allocation"""
        mock_df = pd.DataFrame(
            {
                "Staff Category": ["Senior", "Junior"],
                "allocation": ["1.0", None],
                "staff category order": [1, 2],
            }
        )
        mock_read_excel.return_value = {"202301": mock_df}

        result = prepare_department_fte_trend_report("test.xlsx", 2023, 1, 2)

        # Should succeed and replace NaN with '-' in markdown
        assert result["return_code"] == ReturnCodes.OK.value
        assert "-" in result["md"][0]["content"]

    def test_get_available_periods_large_range(self):
        """Test with a range larger than 12 months"""
        data_available = [f"2023{str(i).zfill(2)}" for i in range(1, 13)]
        result = get_available_periods(data_available, 2023, 1, 24)

        assert result["return_code"] == ReturnCodes.ERROR_PROGRAM.value

    @patch("dataprocess.pd.read_excel")
    def test_empty_dataframe(self, mock_read_excel):
        """Test handling of empty dataframes (missing columns)"""
        mock_read_excel.return_value = {"202301": pd.DataFrame()}

        result = prepare_department_fte_trend_report("test.xlsx", 2023, 1, 2)

        # Missing columns are handled and reported as a data error
        assert result["return_code"] == ReturnCodes.ERROR_FILE_DATA.value


# Pytest fixtures
@pytest.fixture
def sample_dataframe():
    """Fixture providing a sample dataframe"""
    return pd.DataFrame(
        {
            "Staff Category": ["Senior", "Junior", "Manager"],
            "allocation": ["1.0", "0.5", "1.0"],
            "staff_number": ["001", "002", "003"],
            "Rank": ["L5", "L3", "L6"],
            "cost centre name": ["IT", "IT", "HR"],
            "cost centre code": ["001", "001", "002"],
            "staff category order": [1, 2, 3],
        }
    )


@pytest.fixture
def mock_excel_file(tmp_path, sample_dataframe):
    """Fixture providing a mock Excel file"""
    file_path = tmp_path / "test_data.xlsx"
    with pd.ExcelWriter(file_path) as writer:
        sample_dataframe.to_excel(writer, sheet_name="202301", index=False)
    return str(file_path)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
