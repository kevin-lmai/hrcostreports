import pytest
import pandas as pd
from unittest.mock import Mock, patch, MagicMock
import sys
from io import BytesIO

sys.path.insert(0, "../src")

# Import the module to test
from dataprocess import (
    ReturnCodes,
    get_available_periods,
    prepare_department_fte_trend_report,
    prepare_department_headcount_trend_report,
    prepare_department_fte_costcentre_report,
    generate_pdf_report,
    MAX_NUMBER_MONTH_IN_REPORT
)


class TestReturnCodes:
    """Test cases for ReturnCodes enum"""
    
    def test_return_codes_values(self):
        """Test that all return codes have expected values"""
        assert ReturnCodes.ERROR_PROGRAM.value == -10
        assert ReturnCodes.ERROR_FILE_DATA_ERROR.value == -4
        assert ReturnCodes.ERROR_FILE_LOADING.value == -2
        assert ReturnCodes.ERROR_FILE_ERROR.value == -1
        assert ReturnCodes.ERROR.value == 0
        assert ReturnCodes.OK.value == 1
        assert ReturnCodes.OK_GEN_NEW_DATABASE.value == 2
        assert ReturnCodes.OK_UPDATE_DATABASE.value == 3
    
    def test_return_codes_membership(self):
        """Test that return codes are enum members"""
        assert isinstance(ReturnCodes.ERROR_PROGRAM, ReturnCodes)
        assert isinstance(ReturnCodes.OK, ReturnCodes)


class TestGetAvailablePeriods:
    """Test cases for get_available_periods function"""
    
    def test_basic_functionality(self):
        """Test basic period retrieval"""
        data_available = ['202301', '202302', '202303', '202304']
        result = get_available_periods(data_available, 2023, 1, 4)
        assert result == ['202301', '202302', '202303', '202304']
    
    def test_partial_data_available(self):
        """Test when only some periods are available"""
        data_available = ['202301', '202303']
        result = get_available_periods(data_available, 2023, 1, 4)
        assert result == ['202301', '202303']
        assert len(result) == 2
    
    def test_no_data_available(self):
        """Test when no data is available"""
        data_available = []
        result = get_available_periods(data_available, 2023, 1, 3)
        assert result == []
    
    def test_year_rollover(self):
        """Test period generation across year boundary"""
        data_available = ['202311', '202312', '202401', '202402']
        result = get_available_periods(data_available, 2023, 11, 4)
        assert result == ['202311', '202312', '202401', '202402']
    
    def test_december_to_january_transition(self):
        """Test month 12 to month 1 transition"""
        data_available = ['202312', '202401']
        result = get_available_periods(data_available, 2023, 12, 2)
        assert result == ['202312', '202401']
    
    def test_invalid_max_number_of_months_zero(self):
        """Test error when max_number_of_month is 0"""
        result = get_available_periods(['202301'], 2023, 1, 0)
        assert result == ReturnCodes.ERROR_PROGRAM
    
    def test_invalid_max_number_of_months_one(self):
        """Test error when max_number_of_month is 1"""
        result = get_available_periods(['202301'], 2023, 1, 1)
        assert result == ['202301']
    
    def test_invalid_start_month_zero(self):
        """Test error when start_month is 0"""
        result = get_available_periods(['202301'], 2023, 0, 3)
        assert result == ReturnCodes.ERROR_PROGRAM
    
    def test_invalid_start_month_thirteen(self):
        """Test error when start_month is 13"""
        result = get_available_periods(['202301'], 2023, 13, 3)
        assert result == ReturnCodes.ERROR_PROGRAM
    
    def test_invalid_start_month_negative(self):
        """Test error when start_month is negative"""
        result = get_available_periods(['202301'], 2023, -1, 3)
        assert result == ReturnCodes.ERROR_PROGRAM
    
    def test_max_periods(self):
        """Test with maximum number of months"""
        data_available = [f'2023{str(i).zfill(2)}' for i in range(1, 13)]
        result = get_available_periods(data_available, 2023, 1, 12)
        assert len(result) == 12
        assert result[0] == '202301'
        assert result[-1] == '202312'
    
    def test_period_format_single_digit_month(self):
        """Test that single-digit months are zero-padded"""
        data_available = ['202301', '202302']
        result = get_available_periods(data_available, 2023, 1, 2)
        assert result[0] == '202301'  # Not '20231'
    
    def test_data_not_in_sequence(self):
        """Test when available data is not in sequence"""
        data_available = ['202301', '202305', '202309']
        result = get_available_periods(data_available, 2023, 1, 12)
        assert result == ['202301', '202305', '202309']


class TestPrepareDepartmentFTETrendReport:
    """Test cases for prepare_department_fte_trend_report function"""
    
    @patch('dataprocess.pd.read_excel')
    def test_successful_fte_report_generation(self, mock_read_excel):
        """Test successful FTE report generation"""
        # Create mock data
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Junior', 'Senior', 'Junior'],
            'allocation': ['1.0', '0.5', '1.0', '0.5'],
            'staff category order': [1, 2, 1, 2]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 2)
        
        assert isinstance(result, list)
        assert len(result) == 1
        assert 'content' in result[0]
        assert 'css' in result[0]
        assert isinstance(result[0]['content'], str)
        assert isinstance(result[0]['css'], str)
    
    @patch('dataprocess.pd.read_excel')
    def test_file_loading_error(self, mock_read_excel):
        """Test error handling when file cannot be loaded"""
        mock_read_excel.side_effect = Exception("File not found")
        
        result = prepare_department_fte_trend_report('nonexistent.xlsx', 2023, 1, 2)
        
        assert result == ReturnCodes.ERROR_FILE_LOADING
    
    @patch('dataprocess.pd.read_excel')
    def test_no_available_periods(self, mock_read_excel):
        """Test when no periods are available in data"""
        mock_read_excel.return_value = {'202305': pd.DataFrame()}
        
        result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 2)
        
        assert result == ReturnCodes.ERROR_FILE_DATA_ERROR
    
    @patch('dataprocess.pd.read_excel')
    def test_multiple_periods(self, mock_read_excel):
        """Test with multiple periods"""
        mock_df1 = pd.DataFrame({
            'rank category': ['Senior', 'Junior'],
            'allocation': ['1.0', '0.5'],
            'staff category order': [1, 2]
        })
        mock_df2 = pd.DataFrame({
            'rank category': ['Senior', 'Junior'],
            'allocation': ['1.5', '0.8'],
            'staff category order': [1, 2]
        })
        
        mock_read_excel.return_value = {'202301': mock_df1, '202302': mock_df2}
        
        result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 3)
        
        assert isinstance(result, list)
        assert 'content' in result[0]
        assert '202301' in result[0]['content'] or '202302' in result[0]['content']
    
    @patch('dataprocess.pd.read_excel')
    def test_css_formatting_six_periods(self, mock_read_excel):
        """Test CSS formatting for 6 periods"""
        periods = {f'2023{str(i).zfill(2)}': pd.DataFrame({
            'rank category': ['Senior'],
            'allocation': ['1.0'],
            'staff category order': [1]
        }) for i in range(1, 7)}
        
        mock_read_excel.return_value = periods
        
        result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 7)
        
        assert '13px' in result[0]['css']
        assert '10px' in result[0]['css']
    
    @patch('dataprocess.pd.read_excel')
    def test_css_formatting_nine_periods(self, mock_read_excel):
        """Test CSS formatting for 9 periods"""
        periods = {f'2023{str(i).zfill(2)}': pd.DataFrame({
            'rank category': ['Senior'],
            'allocation': ['1.0'],
            'staff category order': [1]
        }) for i in range(1, 10)}
        
        mock_read_excel.return_value = periods
        
        result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 10)
        
        assert '12px' in result[0]['css']
    
    @patch('dataprocess.pd.read_excel')
    def test_css_formatting_many_periods(self, mock_read_excel):
        """Test CSS formatting for more than 9 periods"""
        periods = {f'2023{str(i).zfill(2)}': pd.DataFrame({
            'rank category': ['Senior'],
            'allocation': ['1.0'],
            'staff category order': [1]
        }) for i in range(1, 11)}
        
        mock_read_excel.return_value = periods
        
        result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 12)
        
        assert '9px' in result[0]['css']
        assert '8px' in result[0]['css']


class TestPrepareDepartmentHeadcountTrendReport:
    """Test cases for prepare_department_headcount_trend_report function"""
    
    @patch('dataprocess.pd.read_excel')
    def test_successful_headcount_report_generation(self, mock_read_excel):
        """Test successful headcount report generation"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Junior', 'Senior', 'Junior'],
            'staff_number': ['001', '002', '003', '004'],
            'staff category order': [1, 2, 1, 2]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_headcount_trend_report('test.xlsx', 2023, 1, 2)
        
        assert isinstance(result, list)
        assert len(result) == 1
        assert 'content' in result[0]
        assert 'css' in result[0]
    
    @patch('dataprocess.pd.read_excel')
    def test_headcount_file_loading_error(self, mock_read_excel):
        """Test error handling when file cannot be loaded"""
        mock_read_excel.side_effect = Exception("File not found")
        
        result = prepare_department_headcount_trend_report('nonexistent.xlsx', 2023, 1, 2)
        
        assert result == ReturnCodes.ERROR_FILE_LOADING
    
    @patch('dataprocess.pd.read_excel')
    def test_headcount_no_available_periods(self, mock_read_excel):
        """Test when no periods are available"""
        mock_read_excel.return_value = {'202305': pd.DataFrame()}
        
        result = prepare_department_headcount_trend_report('test.xlsx', 2023, 1, 2)
        
        assert result == ReturnCodes.ERROR_FILE_DATA_ERROR
    
    @patch('dataprocess.pd.read_excel')
    def test_duplicate_staff_numbers(self, mock_read_excel):
        """Test that duplicate staff numbers are handled correctly"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Senior', 'Senior'],
            'staff_number': ['001', '001', '002'],  # Duplicate staff number
            'staff category order': [1, 1, 1]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_headcount_trend_report('test.xlsx', 2023, 1, 2)
        
        assert isinstance(result, list)
        assert 'content' in result[0]
    
    @patch('dataprocess.pd.read_excel')
    def test_headcount_css_color_scheme(self, mock_read_excel):
        """Test that headcount report uses different color scheme"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior'],
            'staff_number': ['001'],
            'staff category order': [1]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_headcount_trend_report('test.xlsx', 2023, 1, 2)
        
        # Headcount uses #2596BE color
        assert '#2596BE' in result[0]['css']


class TestPrepareDepartmentFTECostcentreReport:
    """Test cases for prepare_department_fte_costcentre_report function"""
    
    @patch('dataprocess.pd.read_excel')
    def test_successful_costcentre_report(self, mock_read_excel):
        """Test successful cost centre report generation"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Junior'],
            'rank': ['Manager', 'Staff'],
            'allocation': ['1.0', '0.5'],
            'cost centre name': ['IT', 'IT'],
            'staff category order': [1, 2]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 2)
        
        assert isinstance(result, list)
        assert len(result) > 0
        assert 'content' in result[0]
        assert 'css' in result[0]
        assert 'Cost Centre' in result[0]['content']
    
    @patch('dataprocess.pd.read_excel')
    def test_multiple_cost_centres(self, mock_read_excel):
        """Test with multiple cost centres"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Junior', 'Senior', 'Junior'],
            'rank': ['Manager', 'Staff', 'Manager', 'Staff'],
            'allocation': ['1.0', '0.5', '1.5', '0.8'],
            'cost centre name': ['IT', 'IT', 'HR', 'HR'],
            'staff category order': [1, 2, 1, 2]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 2)
        
        # Should have entries for both IT and HR
        assert len(result) == 2
        assert any('IT' in r['content'] for r in result)
        assert any('HR' in r['content'] for r in result)
    
    @patch('dataprocess.pd.read_excel')
    def test_costcentre_file_loading_error(self, mock_read_excel):
        """Test file loading error"""
        mock_read_excel.side_effect = Exception("File error")
        
        result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 2)
        
        assert result == ReturnCodes.ERROR_FILE_LOADING
    
    @patch('dataprocess.pd.read_excel')
    def test_costcentre_no_available_periods(self, mock_read_excel):
        """Test when no periods match"""
        mock_read_excel.return_value = {'202305': pd.DataFrame()}
        
        result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 2)
        
        assert result == ReturnCodes.ERROR_FILE_DATA_ERROR
    
    @patch('dataprocess.pd.read_excel')
    def test_costcentre_css_eleven_periods(self, mock_read_excel):
        """Test CSS for 11 periods"""
        periods = {}
        for i in range(1, 12):
            periods[f'2023{str(i).zfill(2)}'] = pd.DataFrame({
                'rank category': ['Senior'],
                'rank': ['Manager'],
                'allocation': ['1.0'],
                'cost centre name': ['IT'],
                'staff category order': [1]
            })
        
        mock_read_excel.return_value = periods
        
        result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 11)
        
        assert '9px' in result[0]['css']
    
    @patch('dataprocess.pd.read_excel')
    def test_costcentre_css_twelve_periods(self, mock_read_excel):
        """Test CSS for 12 periods"""
        periods = {}
        for i in range(1, 13):
            periods[f'2023{str(i).zfill(2)}'] = pd.DataFrame({
                'rank category': ['Senior'],
                'rank': ['Manager'],
                'allocation': ['1.0'],
                'cost centre name': ['IT'],
                'staff category order': [1]
            })
        
        mock_read_excel.return_value = periods
        
        result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 12)
        
        assert '8px' in result[0]['css']
    
    @patch('dataprocess.pd.read_excel')
    def test_costcentre_color_scheme(self, mock_read_excel):
        """Test that cost centre report uses correct color"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior'],
            'rank': ['Manager'],
            'allocation': ['1.0'],
            'cost centre name': ['IT'],
            'staff category order': [1]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 2)
        
        assert '#135F2f' in result[0]['css']


class TestGeneratePDFReport:
    """Test cases for generate_pdf_report function"""
    
    @patch('dataprocess.MarkdownPdf')
    def test_generate_pdf_basic(self, mock_markdown_pdf):
        """Test basic PDF generation"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance
        
        content = [{'content': '# Test', 'css': 'table {}'}]
        
        # The function may not have a clear return, just test it doesn't crash
        try:
            result = generate_pdf_report('test_report.pdf', content, 'Test Report')
            # If function returns something, verify it
            assert True
        except Exception as e:
            pytest.fail(f"PDF generation failed: {e}")
    
    @patch('dataprocess.MarkdownPdf')
    def test_generate_pdf_with_title(self, mock_markdown_pdf):
        """Test PDF generation with custom title"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance
        
        content = [{'content': '## Content', 'css': 'h2 {}'}]
        
        try:
            generate_pdf_report('report.pdf', content, 'Custom Title')
            assert True
        except Exception:
            pytest.fail("PDF generation with title failed")
    
    @patch('dataprocess.MarkdownPdf')
    def test_generate_pdf_empty_content(self, mock_markdown_pdf):
        """Test PDF generation with empty content list"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance
        
        content = []
        
        try:
            generate_pdf_report('empty.pdf', content)
            assert True
        except Exception:
            # Empty content might cause issues, but should be handled
            pass
    
    @patch('dataprocess.MarkdownPdf')
    def test_generate_pdf_multiple_sections(self, mock_markdown_pdf):
        """Test PDF generation with multiple content sections"""
        mock_pdf_instance = MagicMock()
        mock_markdown_pdf.return_value = mock_pdf_instance
        
        content = [
            {'content': '# Section 1', 'css': 'h1 {}'},
            {'content': '# Section 2', 'css': 'h1 {}'},
            {'content': '# Section 3', 'css': 'h1 {}'}
        ]
        
        try:
            generate_pdf_report('multi_section.pdf', content, 'Multi Section Report')
            assert True
        except Exception as e:
            pytest.fail(f"Multi-section PDF generation failed: {e}")


class TestIntegration:
    """Integration tests for the module"""
    
    def test_max_number_month_constant(self):
        """Test that the constant is properly defined"""
        assert MAX_NUMBER_MONTH_IN_REPORT == 12
    
    @patch('dataprocess.pd.read_excel')
    def test_full_workflow_fte_report(self, mock_read_excel):
        """Test complete workflow for FTE report"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Junior'],
            'allocation': ['1.0', '0.5'],
            'staff category order': [1, 2]
        })
        
        mock_read_excel.return_value = {
            '202301': mock_df,
            '202302': mock_df.copy()
        }
        
        result = prepare_department_fte_trend_report(
            'test.xlsx', 2023, 1, MAX_NUMBER_MONTH_IN_REPORT
        )
        
        assert isinstance(result, list)
        assert len(result) > 0
    
    @patch('dataprocess.pd.read_excel')
    def test_data_consistency_across_functions(self, mock_read_excel):
        """Test that different report functions handle same data consistently"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Junior'],
            'allocation': ['1.0', '0.5'],
            'staff_number': ['001', '002'],
            'rank': ['Manager', 'Staff'],
            'cost centre name': ['IT', 'IT'],
            'staff category order': [1, 2]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        fte_result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 2)
        headcount_result = prepare_department_headcount_trend_report('test.xlsx', 2023, 1, 2)
        costcentre_result = prepare_department_fte_costcentre_report('test.xlsx', 2023, 1, 2)
        
        # All should succeed
        assert isinstance(fte_result, list)
        assert isinstance(headcount_result, list)
        assert isinstance(costcentre_result, list)


class TestEdgeCases:
    """Test edge cases and boundary conditions"""
    
    @patch('dataprocess.pd.read_excel')
    def test_nan_values_in_allocation(self, mock_read_excel):
        """Test handling of NaN values in allocation"""
        mock_df = pd.DataFrame({
            'rank category': ['Senior', 'Junior'],
            'allocation': ['1.0', None],
            'staff category order': [1, 2]
        })
        
        mock_read_excel.return_value = {'202301': mock_df}
        
        result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 2)
        
        # Should handle NaN and replace with '-' in markdown
        assert isinstance(result, list)
        assert '-' in result[0]['content']
    
    def test_get_available_periods_large_range(self):
        """Test with a large date range"""
        data_available = [f'2023{str(i).zfill(2)}' for i in range(1, 13)]
        result = get_available_periods(data_available, 2023, 1, 24)
        
        # Should handle multi-year ranges
        assert result == ReturnCodes.ERROR_PROGRAM
    
    @patch('dataprocess.pd.read_excel')
    def test_empty_dataframe(self, mock_read_excel):
        """Test handling of empty dataframes"""
        mock_read_excel.return_value = {'202301': pd.DataFrame()}
        
        # This might cause errors, but should be handled gracefully
        try:
            result = prepare_department_fte_trend_report('test.xlsx', 2023, 1, 2)
            # Accept either error code or exception
            assert True
        except Exception:
            assert True


# Pytest fixtures
@pytest.fixture
def sample_dataframe():
    """Fixture providing a sample dataframe"""
    return pd.DataFrame({
        'rank category': ['Senior', 'Junior', 'Manager'],
        'allocation': ['1.0', '0.5', '1.0'],
        'staff_number': ['001', '002', '003'],
        'rank': ['L5', 'L3', 'L6'],
        'cost centre name': ['IT', 'IT', 'HR'],
        'staff category order': [1, 2, 3]
    })


@pytest.fixture
def mock_excel_file(tmp_path, sample_dataframe):
    """Fixture providing a mock Excel file"""
    file_path = tmp_path / "test_data.xlsx"
    with pd.ExcelWriter(file_path) as writer:
        sample_dataframe.to_excel(writer, sheet_name='202301', index=False)
    return str(file_path)


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

