import pandas as pd
import os
import calendar
from markdown_pdf import MarkdownPdf, Section
from py_markdown_table.markdown_table import markdown_table
from enum import Enum

DEBUG = True
MAX_NUMBER_MONTH_IN_REPORT = 12

class ReturnCodes(Enum):
    ERROR_PROGRAM = -10
    ERROR_FILE_DATA_STAFF = -4
    ERROR_FILE_LOADING = -2
    ERROR_FILE_ERROR = -1
    ERROR = 0
    OK = 1
    OK_GEN_NEW_DATABASE = 2
    OK_UPDATE_DATABASE = 3
    
RETURN_CODE = Enum('Color', [('RED', 1), ('GREEN', 2), ('BLUE', 3)])

pd.options.display.float_format = '${:,.2f}'.format

def get_available_periods(data_available: list, start_year: int, start_month: int, max_number_of_month: int):
    if max_number_of_month <= 1:
        return ReturnCodes.ERROR_PROGRAM
        #raise "Number of month should be greater than 1"
    if start_month < 1 or start_month > 12:
        return ReturnCodes.ERROR_PROGRAM
        raise "Start month should be between 1 and 12"
    
    report_periods = []
    report_year = start_year
    report_month = start_month
    for i in range(max_number_of_month):        
        report_periods.append(f"{str(report_year)}{str(report_month).zfill(2)}")
        if report_month == 12:
            report_month = 1
            report_year += 1
        else:
            report_month += 1
    
    available_periods = []
    for period in report_periods:
        if period in data_available:
            available_periods.append(period)
    
    return available_periods

    
def prepare_department_fte_trend_report(data_file_name: str, start_year: int, start_month: int, max_number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT):
    
    
    try:
        data_df_dict = pd.read_excel(data_file_name,sheet_name=None,header=0,dtype=object)
    except Exception as e:
        return ReturnCodes.ERROR_FILE_LOADING
        #raise f"Error loading file {data_file_name}: {e}"

    available_periods = get_available_periods(data_df_dict.keys(), start_year, start_month, max_number_of_month)

    if type(available_periods) is int:
        return ReturnCodes.ERROR_PROGRAM
    
    result_dict = {}
    results_order_dict = {}

    
    for period in available_periods:
        data_df = data_df_dict[period]
        
        
        period_df = data_df.groupby(['rank category'])['allocation'].sum().astype(float)
        
        result_dict[period] = period_df
        result_order_df = data_df.drop_duplicates(subset=['rank category']).loc[:, ['rank category', 'staff category order']]
        dict_from_zipped = dict(zip(result_order_df['rank category'], result_order_df['staff category order']))
        if len(results_order_dict.keys()) == 0:
            results_order_dict = dict_from_zipped
        else:
            results_order_dict.update(dict_from_zipped)
                
    result_order_to_df = {}
    result_order_to_df['rank category'] = []
    result_order_to_df['staff category order'] = []
    order=1
    for k, v in sorted(results_order_dict.items(), key=lambda item: (item, item)):
        result_order_to_df['rank category'].append(k)
        result_order_to_df['staff category order'].append(order)
        order += 1


    result = pd.DataFrame(result_dict)
    results_order_df = pd.DataFrame.from_dict(result_order_to_df)

    
    sorted_result_df = pd.merge(result, results_order_df, on='rank category', how='inner')
    sorted_result_df.set_index('staff category order', inplace=True)


    sorted_result_df.sort_index(inplace=True)
    sorted_result_df.set_index('rank category',inplace=True)

    sorted_result_df.loc['Total'] = sorted_result_df.sum(numeric_only=True)

    sorted_result_df.reset_index(inplace=True)
    

    sorted_result_dict = sorted_result_df.round(2).astype(str).to_dict(orient='index')
    
    markdown_table_data = []
    
    for k,v in sorted_result_dict.items():
        markdown_table_data.append(v)
    
    markdown = markdown_table(markdown_table_data).set_params(row_sep = 'markdown', quote = False).get_markdown()
    
    
    '''
    css = """
        table {width: 100%; border-collapse: collapse;}
        thead th {background-color: #4CAF50;color: white;padding: 5px;text-align: center;border: 1px solid #ddd;border-collapse: collapse;}
        tbody td:first-child { text-align: left; font-weight: bold;border-collapse: collapse;}
        tbody td { padding: 10px 20px; text-align: center; border-collapse: collapse;}
        h2 {text-align: center; background-color: #4CAF50;}
    """
    '''
    
    if len(available_periods) <= 6:
        cell_y_padding = "10px"
        header_cell_y_padding = "25px"
        font_size = "13px"
        padding = "2px"
    elif len(available_periods) <= 9:
        cell_y_padding = "10px"
        header_cell_y_padding = "10px"
        font_size = "12px"
        padding = "3px"
    else:
        cell_y_padding = "8px"
        header_cell_y_padding = "10px"
        font_size = "9px"
        padding = "2px"
    table_header_bg_color = '#4CAF50'
    title_bg_color = '#4CAF50'
    css = f"table {{width: 100%; border-collapse: collapse; font-size: {font_size} ; padding: {padding} {cell_y_padding} }} thead th {{background-color: {table_header_bg_color};color: white;text-align: center;border: 1px solid #ddd;border-collapse: collapse; padding: {padding} {header_cell_y_padding}}} tbody td:first-child {{ text-align: left; font-weight: bold;border-collapse: collapse;}} tbody td {{ text-align: center; border-collapse: collapse;}} h2 {{text-align: center; background-color: {title_bg_color};}}"
    
    
    return_md = {}
    return_md['content'] = markdown
    return_md['css'] = css
    return([return_md])

    
    
def prepare_department_headcount_trend_report(data_file_name: str, start_year: int, start_month: int, max_number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT):
    
    try:
        data_df_dict = pd.read_excel(data_file_name,sheet_name=None,header=0,dtype=object)
    except Exception as e:
        return ReturnCodes.ERROR_FILE_LOADING

    available_periods = get_available_periods(data_df_dict.keys(), start_year, start_month, max_number_of_month)
    
    if type(available_periods) is int:
        return ReturnCodes.ERROR_PROGRAM
    
    result_dict = {}
    results_order_dict = {}
            
    for period in available_periods:
        data_df = data_df_dict[period]
    
        period_df = data_df.drop_duplicates(subset=['staff_number']).groupby(['rank category']).size()
        result_dict[period] = period_df
        
        result_order_df = data_df.drop_duplicates(subset=['rank category']).loc[:, ['rank category', 'staff category order']]
        dict_from_zipped = dict(zip(result_order_df['rank category'], result_order_df['staff category order']))
        if len(results_order_dict.keys()) == 0:
            results_order_dict = dict_from_zipped
        else:
            results_order_dict.update(dict_from_zipped)
        
    
    result_order_to_df = {}
    result_order_to_df['rank category'] = []
    result_order_to_df['staff category order'] = []
    order=1
    for k, v in sorted(results_order_dict.items(), key=lambda item: (item, item)):
        result_order_to_df['rank category'].append(k)
        result_order_to_df['staff category order'].append(order)
        order += 1

    result = pd.DataFrame(result_dict)    
    results_order_df = pd.DataFrame.from_dict(result_order_to_df)
        
    sorted_result_df = pd.merge(result, results_order_df, on='rank category', how='inner')
    sorted_result_df.set_index('staff category order', inplace=True)


    sorted_result_df.sort_index(inplace=True)
    sorted_result_df.set_index('rank category',inplace=True)
    
    sorted_result_df.loc['Total'] = sorted_result_df.sum(numeric_only=True)

    sorted_result_df.reset_index(inplace=True)
    

    sorted_result_dict = sorted_result_df.round(2).astype(str).to_dict(orient='index')
    
    markdown_table_data = []
    
    
    
    for k,v in sorted_result_dict.items():
        markdown_table_data.append(v)
    
    markdown = markdown_table(markdown_table_data).set_params(row_sep = 'markdown', quote = False).get_markdown()
    
    
    '''
    css = """
        table {width: 100%; border-collapse: collapse;}
        thead th {background-color: #4CAF50;color: white;padding: 5px;text-align: center;border: 1px solid #ddd;border-collapse: collapse;}
        tbody td:first-child { text-align: left; font-weight: bold;border-collapse: collapse;}
        tbody td { padding: 10px 20px; text-align: center; border-collapse: collapse;}
        h2 {text-align: center; background-color: #4CAF50;}
    """
    '''
    
    if len(available_periods) <= 6:
        cell_y_padding = "10px"
        header_cell_y_padding = "25px"
        font_size = "13px"
        padding = "2px"
    elif len(available_periods) <= 9:
        cell_y_padding = "10px"
        header_cell_y_padding = "10px"
        font_size = "12px"
        padding = "3px"
    else:
        cell_y_padding = "8px"
        header_cell_y_padding = "10px"
        font_size = "9px"
        padding = "2px"
    table_header_bg_color = '#2596BE'
    title_bg_color = '#2596BE'
    css = f"table {{width: 100%; border-collapse: collapse; font-size: {font_size} ; padding: {padding} {cell_y_padding} }} thead th {{background-color: {table_header_bg_color} ;color: white;text-align: center;border: 1px solid #ddd;border-collapse: collapse; padding: {padding} {header_cell_y_padding}}} tbody td:first-child {{ text-align: left; font-weight: bold;border-collapse: collapse;}} tbody td {{ text-align: center; border-collapse: collapse;}} h2 {{text-align: center; background-color: {title_bg_color};}}"
    
    
    return_md = {}
    return_md['content'] = markdown
    return_md['css'] = css
    return([return_md])


    
def prepare_department_fte_costcentre_report(data_file_name: str, start_year: int, start_month: int, max_number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT):
    
    try:
        data_df_dict = pd.read_excel(data_file_name,sheet_name=None,header=0,dtype=object)
    except Exception as e:
        return ReturnCodes.ERROR_FILE_LOADING

    available_periods = get_available_periods(data_df_dict.keys(), start_year, start_month, max_number_of_month)
    
    if type(available_periods) is int:
        return ReturnCodes.ERROR_PROGRAM
    
    return_md = []
    all_costcentre_result_dict = {}
        
    for period in available_periods:
        data_df = data_df_dict[period]

        cost_centres = data_df['cost centre name'].copy().drop_duplicates().to_list()
        
        data_df['rank category duplicate'] = data_df['rank category']
        
        for c in cost_centres:
            if c not in all_costcentre_result_dict.keys():
                all_costcentre_result_dict[c] = {period : data_df[data_df['cost centre name'] == c]}
            else:
                all_costcentre_result_dict[c][period] = data_df[data_df['cost centre name'] == c]
                
                

    for cost_centre, v in sorted(all_costcentre_result_dict.items()):
        result_dict = {}
        results_order_dict = {}
        for period, target_df in v.items():
            period_df = target_df.groupby(['rank category', 'rank'])['allocation'].sum().astype(float)
            result_dict[period] = period_df
            
            result_order_df = data_df.drop_duplicates(subset=['rank category']).loc[:, ['rank category', 'staff category order']]
            dict_from_zipped = dict(zip(result_order_df['rank category'], result_order_df['staff category order']))
            if len(results_order_dict.keys()) == 0:
                results_order_dict = dict_from_zipped
            else:
                results_order_dict.update(dict_from_zipped)
            
        result_order_to_df = {}
        result_order_to_df['rank category'] = []
        result_order_to_df['staff category order'] = []
        order=1
        for k, v in sorted(results_order_dict.items(), key=lambda item: (item, item)):
            result_order_to_df['rank category'].append(k)
            result_order_to_df['staff category order'].append(order)
            order += 1

        result = pd.DataFrame(result_dict)
        results_order_df = pd.DataFrame.from_dict(result_order_to_df)

        
        sorted_result_df = pd.merge(result, results_order_df, on='rank category', how='inner')
        sorted_result_df.set_index('staff category order', inplace=True)


        sorted_result_df.sort_index(inplace=True)
        sorted_result_df.set_index('rank category',inplace=True)

        sorted_result_df.loc['Total'] = sorted_result_df.sum(numeric_only=True)

        sorted_result_df.reset_index(inplace=True)
        
        sorted_result_dict = sorted_result_df.round(2).astype(str).to_dict(orient='index')
    
        markdown_table_data = []
    
        for k,v in sorted_result_dict.items():
            markdown_table_data.append(v)
        
        markdown = markdown_table(markdown_table_data).set_params(row_sep = 'markdown', quote = False).get_markdown()
        markdown_with_costcentre_name = f'#### Cost Centre : {cost_centre}\n\n{markdown}'

        if len(available_periods) <= 6:
            cell_y_padding = "10px"
            header_cell_y_padding = "25px"
            font_size = "13px"
            padding = "2px"
        elif len(available_periods) <= 9:
            cell_y_padding = "10px"
            header_cell_y_padding = "10px"
            font_size = "12px"
            padding = "3px"
        else:
            cell_y_padding = "8px"
            header_cell_y_padding = "10px"
            font_size = "9px"
            padding = "2px"
        table_header_bg_color = '#135F2f'
        title_bg_color = '#135F2f'
        css = f"table {{width: 100%; border-collapse: collapse; font-size: {font_size} ; padding: {padding} {cell_y_padding} }} thead th {{background-color: {table_header_bg_color};color: white;text-align: center;border: 1px solid #ddd;border-collapse: collapse; padding: {padding} {header_cell_y_padding}}} tbody td:first-child {{ text-align: left; font-weight: bold;border-collapse: collapse;}} tbody td {{ text-align: center; border-collapse: collapse;}} h2 {{text-align: center; background-color: {title_bg_color};}} h3 h4 {{text-align: left;}}"
        
        
        result_md = {}
        result_md['content'] = markdown_with_costcentre_name
        result_md['css'] = css
        return_md.append(result_md)
    return(return_md)

def generate_pdf_report(report_name: str, content : list, title: str = "Report"):

    '''
    text = """# Section with Table

    |TableHeader1|TableHeader2|
    |--|--|
    |Text1|Text2|
    |ListCell|<ul><li>FirstBullet</li><li>SecondBullet</li></ul>|
    """
    '''
    #css = "table, th, td {border: 1px solid black;}"

    #header= f'<h1 style="text-align: center;">{title}</h1>'
    header= f'## {title}'

    #header2 = f'### Cost Centre : {xxx}'

    #header = f"##{title} XXX\n###Hello World\n\n"
    #header = "##Head2\n\n### <a id='head3'></a>Head3\n\n"
    #header = "##Head2\n"
    #header = 'Head'
    
    pdf = MarkdownPdf()
    for c in content:
        pdf.add_section(Section(header+'\n\n'+c['content'], paper_size="A4-L", toc=False), user_css=c['css'])
    pdf.save(report_name)

def check_file_header(df: pd.DataFrame, expected_headers: list) -> list:
    """
        check header are available in dataframe
    """
    missing_headers = []
    for h in expected_headers:
        if h not in df.columns:
            missing_headers.append(h)
    
    return missing_headers

'''
def generate_report(data_file_name: str):
    try:
        data_df = pd.read_excel(data_file_name,sheet_name=None,header=0,dtype=object)
    except Exception as e:
        return f"Error loading base sheet: {e}"

    #header = ["staff number","rank","rank category","staff category order","cost centre code","cost centre name","allocation"]
          
    pass
'''

def process_update_database(excelfile: str, month_of_data_str:str, reportname:str) -> int:
    """
    This function is used for CUHK Hospital Data Processing of HR / Cost Centre Allocation.
    
    Process the input Excel file which has two sheets by expanding the data in sheet1 with sheet2
    
    Parameters:
    excelfile (str): Path to the Excel file containing the data, with 2 sheets

    Returns:
    error message or success message which is "OK"
    """
    
    # read whole file:
    '''
    ### want to read all sheet but in vains as different sheets has differ header row number
    
    try:
        file_dict = pd.read_excel(excelfile,sheet_name=None,header=0,dtype=object)
    except Exception as e:
        raise f"Error loading file {excelfile}: {e}"
    
    
    file_dict_keys_list = []
    for key in file_dict.keys():
        file_dict_keys_list.append(key)
        
    if len(file_dict_keys_list) < 3:
        raise f"Data file should have at least 3 sheets"
    
    '''
    
    # read sheet 1
    try:
        file_base_data_df = pd.read_excel(excelfile,sheet_name=0,header=0,dtype=object)
    except Exception as e:
        return ReturnCodes.ERROR_FILE_ERROR
    
    header = ['StaffNo', 'Rank', 'Section', 'Staff Category', 'FTE', 'Default Cost Centre']
    missing_headers = check_file_header(file_base_data_df, header)
    if len(missing_headers) > 0:
        return ReturnCodes.ERROR_FILE_ERROR
    
    base_data_df = file_base_data_df[header]
    clean_base_data_df = base_data_df.dropna(how='all')
    
    # Remove last row , if Rank is empty
    last_row_df = clean_base_data_df.tail(1)
    if len(last_row_df.dropna(subset=['Rank'])) == 0:
        clean_base_data_df = clean_base_data_df.head(len(clean_base_data_df) - 1)

    file_base_records_count : int = len(base_data_df.index)
    clean_base_records_count : int = len(clean_base_data_df.index)

    if file_base_records_count != clean_base_records_count:
        pass
        #print(f"Base data had {file_base_records_count - clean_base_records_count} empty rows removed.")


    # check for duplicate StaffNo in base data
    if len(pd.unique(clean_base_data_df['StaffNo'])) != len(clean_base_data_df):
        #print("Warning: Duplicate Staff Numbers found in Base Data!")
        return ReturnCodes.ERROR_FILE_DATA_STAFF
    
    # set the right data types for data Series
    clean_base_data_df['FTE'] = clean_base_data_df['FTE'].astype(float)

    rank_cat = pd.DataFrame(clean_base_data_df['Rank']+"\t"+clean_base_data_df['Staff Category'])
    rank_cat.columns = ['cat_info']
    unique_rank_cat = pd.unique(rank_cat['cat_info']).tolist()
    unique_rank_cat_dict = {}
    for i in unique_rank_cat:
        cats = i.split("\t")
        unique_rank_cat_dict[cats[0]] = cats[1]
    if DEBUG:
        print(unique_rank_cat_dict)
        print(type(unique_rank_cat_dict))

    # read sheet 2
    try:
        file_expand_data_df = pd.read_excel(excelfile,sheet_name=1,header=1,dtype=object)
    except Exception as e:

        return ReturnCodes.ERROR_FILE_ERROR
        
    first_row_df = file_expand_data_df.head(1)
    if len(first_row_df.dropna(subset=['Rank'])) == 0:
        clean_base_data_df = clean_base_data_df.head(len(clean_base_data_df) - 1)

    header = ['StaffNo', 'Rank', 'CCode', 'CostCentre', 'Allocated Percentage']
    missing_headers = check_file_header(file_expand_data_df, header)
    if len(missing_headers) > 0:

        return ReturnCodes.ERROR_FILE_ERROR
    
    expand_data_df = file_expand_data_df[header]
    clean_expand_data_df = expand_data_df.dropna(how='all')

    file_expand_records_count : int = len(file_expand_data_df.index)        
    clean_expand_records_count : int = len(clean_expand_data_df.index)
        
    new_clean_expand_data_df = clean_expand_data_df.copy()
    new_clean_expand_data_df['Allocated Percentage'] = new_clean_expand_data_df['Allocated Percentage'].astype(float)
    new_clean_expand_data_df['Allocated Percentage'] = new_clean_expand_data_df['Allocated Percentage']/100.0
    clean_expand_data_df = new_clean_expand_data_df

    # read sheet 3, cost center information
    try:
        file_cost_centre_data_df = pd.read_excel(excelfile,sheet_name=2,header=0,dtype=object)
    except Exception as e:
        #print(f"Error loading base sheet 3: {e}")
        return ReturnCodes.ERROR_FILE_ERROR
    
    
    header = ['Value', 'Description','Enabled/ Disabled']
    missing_headers = check_file_header(file_cost_centre_data_df, header)
    if len(missing_headers) > 0:
        #print(f"Error: sheet 3 Missing expected column '{", ".join(missing_headers)}' in cost centre  data sheet.")
        return ReturnCodes.ERROR_FILE_ERROR

    
    cost_centre_data_df = file_cost_centre_data_df[header]
    clean_cost_centre_data_df = cost_centre_data_df.dropna(how='all')
    clean_cost_centre_data_df = clean_cost_centre_data_df[clean_cost_centre_data_df['Enabled/ Disabled'] == 'Enabled']

    file_cost_centre_records_count : int = len(file_cost_centre_data_df.index)        
    clean_cost_centre_records_count : int = len(clean_cost_centre_data_df.index)
    
    if DEBUG:    
        if file_cost_centre_records_count != clean_cost_centre_records_count:
            print(f"Expand data had {file_expand_records_count - clean_expand_records_count} empty rows removed.")


    # get the cost centre information
    clean_cost_centre_dict = clean_cost_centre_data_df.to_dict(orient='index')
    
    cost_centre_info = {}
    for k,v in clean_cost_centre_dict.items():
        cost_centre_info[v['Value']] = v['Description']
            
    
    # read sheet 4 Staff Category Order
    try:
        file_staff_category_order_data_df = pd.read_excel(excelfile,sheet_name=3,header=0,dtype=object)
        has_staff_category_order_data = True
    except Exception as e:
        has_staff_category_order_data = False


    if has_staff_category_order_data:
        header = ['Staff Category', 'Order']
        missing_headers = check_file_header(file_staff_category_order_data_df, header)
        if len(missing_headers) > 0:
            return ReturnCodes.ERROR_FILE_ERROR
            #return f"Error: Missing expected column '{", ".join(missing_headers)}' in staff category data sheet."
        
        staff_category_order_data_df = file_staff_category_order_data_df[header]
        clean_staff_category_order_data_df = staff_category_order_data_df.dropna(how='all')
        
        file_staff_category_order_records_count : int = len(file_staff_category_order_data_df.index)        
        clean_staff_category_order_records_count : int = len(clean_staff_category_order_data_df.index)
        
        if DEBUG:
            if file_staff_category_order_records_count != clean_staff_category_order_records_count:
                print(f"Expand data had {file_expand_records_count - clean_expand_records_count} empty rows removed.")
    
        clean_staff_category_order_data_dict = clean_staff_category_order_data_df.to_dict(orient='index')

        
        sorted_staff_category_order_dict = dict(sorted(clean_staff_category_order_data_dict.items(), key=lambda item: item))
        
        count = 1
        staff_category_order_dict = {}
        for k,v in sorted_staff_category_order_dict.items():
            staff_category_order_dict[v['Staff Category']] = count
            count += 1
            

    else:

        sorted_staff_category_order_list = sorted(pd.unique(clean_base_data_df['Staff Category']).tolist())
        #sorted_staff_category_order_list = sorted(staff_category_order_list)
        staff_category_order_dict = {}
        for i in range(len(sorted_staff_category_order_list)):
            staff_category_order_dict[sorted_staff_category_order_list[i]] = i+1
        
    # expand the list
    clean_base_data_df = clean_base_data_df.set_index('StaffNo')
    clean_base_dict = clean_base_data_df.to_dict(orient='index')
    # clean_base_dict has key as StaffNo, value as dict of other columns
    clean_expand_dict = clean_expand_data_df.to_dict(orient='index')
    # clean_expand_dict has key as index, value as dict of other columns, including StaffNo as StaffNo is not unique here
        
    expanded_entries = []
    staff_rank_category = {}
    unique_staff_in_base = set()
    for k,v in clean_expand_dict.items():
        staff_number = v['StaffNo']
        if staff_number in clean_base_dict.keys():
            staff_rank_category[staff_number] = clean_base_dict[staff_number]['Staff Category']
            unique_staff_in_base.add(staff_number)
            del clean_base_dict[staff_number]
        if staff_number in unique_staff_in_base:
            expanded_entries.append({'staff_number' : staff_number, 'rank' : v['Rank'], 'rank category' : unique_rank_cat_dict[v['Rank']], 'staff category order' : staff_category_order_dict[unique_rank_cat_dict[v['Rank']]], 'cost centre code' : v['CCode'], 'cost centre name' : cost_centre_info[v['CCode']], 'allocation' : v['Allocated Percentage']})
        else:
            pass
            # found record in expand data but not in base data. It is not counted as error, just skip it.
            # print(f"Error : Rank {v['Rank']} not found in base data for staff number {staff_number}.")
         
    for k,v in clean_base_dict.items():
        expanded_entries.append({'staff_number' : k, 'rank' : v['Rank'], 'rank category' : v['Staff Category'], 'staff category order' : staff_category_order_dict[v['Staff Category']],'cost centre code' : v['Default Cost Centre'], 'cost centre name' : cost_centre_info[v['Default Cost Centre']], 'allocation' : v['FTE']})
    
    result_df = pd.DataFrame(expanded_entries)
    
    if os.path.exists(reportname) == True:
        with pd.ExcelWriter(f"{reportname}", mode='a') as writer:  
                #df1.to_excel(writer, sheet_name='Sheet_name_3')
            workBook = writer.book
            try:
                workBook.remove(workBook[f"{month_of_data_str}"])
            except:  # noqa: E722
                pass
                #print(f"Error: removing existing sheet {month_of_data_str} in {reportname}")
                #return ReturnCodes.ERROR_FILE_ERROR
            finally:
                result_df.to_excel(writer,index=False,sheet_name=f"{month_of_data_str}")
        return ReturnCodes.OK_UPDATE_DATABASE
    else:
        with pd.ExcelWriter(f"{reportname}", mode='w') as writer:  
                #df1.to_excel(writer, sheet_name='Sheet_name_3')
            result_df.to_excel(writer,index=False,sheet_name=f"{month_of_data_str}")
        return ReturnCodes.OK_GEN_NEW_DATABASE
        

def generate_department_fte_summary_report(fte_data_file_name: str, summary_report_file_name: str, report_title: str, start_year: int, start_month: int, number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT):
    department_fte_trend_content = prepare_department_fte_trend_report(fte_data_file_name,start_year, start_month,number_of_month)
    if type(department_fte_trend_content) is int:
        return department_fte_trend_content
    elif type(department_fte_trend_content) is list:
        period = f"{str(start_year)}" if start_month == 1 else f"{str(start_year)}/{str(start_year+1)}"
        report_title = f"{report_title} {period}"
        generate_pdf_report(summary_report_file_name, department_fte_trend_content, report_title)
    else:
        if DEBUG:
            print(f"Error: department_headcount_trend_content is '{department_fte_trend_content}'")
        return ReturnCodes.ERROR_PROGRAM

    return ReturnCodes.OK


def generate_department_headcount_summary_report(fte_data_file_name: str, summary_report_file_name: str, report_title: str, start_year: int, start_month: int, number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT):
    department_headcount_trend_content = prepare_department_headcount_trend_report(fte_data_file_name,start_year, start_month,number_of_month)
    if type(department_headcount_trend_content) is int:
        return department_headcount_trend_content
    elif type(department_headcount_trend_content) is list:
        period = f"{str(start_year)}" if start_month == 1 else f"{str(start_year)}/{str(start_year+1)}"
        report_title = f"{report_title} {period}"
        generate_pdf_report(summary_report_file_name, department_headcount_trend_content, report_title)
    else:
        if DEBUG:
            print(f"Error: department_headcount_trend_content is '{department_headcount_trend_content}'")
        return ReturnCodes.ERROR_PROGRAM
    
    return ReturnCodes.OK


def generate_department_fte_costcentre_report(fte_data_file_name: str, costcentre_report_file_name: str, report_title: str, start_year: int, start_month: int, number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT):
    department_fte_costcentre_content = prepare_department_fte_costcentre_report(fte_data_file_name,start_year, start_month,number_of_month)
    #print(department_fte_trend_content)
    if type(department_fte_costcentre_content) is int:
        return department_fte_costcentre_content
    elif type(department_fte_costcentre_content) is list:
        period = f"{str(start_year)}" if start_month == 1 else f"{str(start_year)}/{str(start_year+1)}"
        report_title = f"{report_title} {period}"
        generate_pdf_report(costcentre_report_file_name, department_fte_costcentre_content, report_title)
    else:
        if DEBUG:
            print(f"Error: department_headcount_trend_content is '{department_fte_costcentre_content}'")
        return ReturnCodes.ERROR_PROGRAM
    
    return ReturnCodes.OK



if __name__ == "__main__":
    # Load the dataset
    
    '''
    report_year = 2025
    report_month = 7
    report_period = f"{str(report_year)}{str(report_month).zfill(2)}"
    report_data_file = "HR_headcount_all.xlsx"
    #filename = 'Headcount to Finance Dept - Aug 2025.xlsx'
    #filename = 'Headcount to Finance Dept - Sept 2025.xlsx'
    #filename = 'Headcount to Finance Dept - Oct 2025.xlsx'
    
    filename = 'Headcount to Finance Dept - July 2025.xlsx'
    
    result = process_update_database(filename,report_period,report_data_file)
    if type(process_update_database) is int and process_update_database <= 0:
        print(f"Error: Error {ReturnCodes(process_update_database).name} ")
    
    '''
    
    database_file = "HR_FTE_Database.xlsx"
    start_year = 2025
    start_month = 7
    
    '''
    department_fte_summary_report_file_name = "HR_department_fte_summary_report.pdf"
    department_fte_summary_report_title = "Yearly Department FTE Trend.pdf"
    
    generate_department_fte_summary_report(database_file, department_fte_summary_report_file_name, department_fte_summary_report_title, start_year, start_month, 12)
    '''
    
        
    department_headcount_summary_report_file_name = "HR_department_headcount_summary_report.pdf"
    department_headcount_summary_report_title = "Yearly Department Headcount Trend.pdf"
    
    generate_department_headcount_summary_report(database_file, department_headcount_summary_report_file_name, department_headcount_summary_report_title, start_year, start_month, 12)

    '''
    department_fte_costcentre_report_file_name = "HR_department_fte_summary_report.pdf"
    department_fte_costcentre_report_title = "Yearly Department FTE Trend.pdf"
    
    generate_department_fte_costcentre_report(database_file, department_fte_costcentre_report_file_name, department_fte_costcentre_report_title, start_year, start_month, 12)

    '''

    '''
    period = f"{str(start_year)}" if start_month == 1 else f"{str(start_year)}/{str(start_year+1)}"
    
    summary_report_file_name = "HR_headcount_summary_report_main.pdf"

    department_fte_trend_content = prepare_department_fte_trend_report(report_data_file,start_year, start_month,number_of_month=12)
    
    if type(department_fte_trend_content) is int:
        print(f"Error: Error {ReturnCodes(department_fte_trend_content).name} ")
    else:
        report_title = "Yearly Department FTE Trend Report " + period   
        generate_pdf_report(summary_report_file_name,department_fte_trend_content, report_title)

    '''

