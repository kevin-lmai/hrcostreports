"""
Module provide all the necessary functions for data processing which are consumed by main program

Functions:
exported class:
- ReturnCodes

exported functions:
- process_source_data
- ReturnCodes
- generate_department_fte_summary_report
- generate_department_headcount_summary_report
- generate_department_fte_costcentre_report
- generate_excel_fr_df

local functions:
- get_available_periods
- prepare_department_fte_trend_report
- prepare_department_headcount_trend_report
- prepare_department_fte_costcentre_report
- generate_pdf_report
- check_file_header
- report_css_style
- clean_sheet_name

"""

import pandas as pd
import os
from markdown_pdf import MarkdownPdf, Section
from py_markdown_table.markdown_table import markdown_table
from enum import Enum
from textwrap import shorten

# set DEBUG True to display verbose debug information, programmer use only
DEBUG = True

HEADER_SEPARATOR = "!"

# set the maximum number of months in report
MAX_NUMBER_MONTH_IN_REPORT = 12
STAFF_CATEGORY_LENGTH = 30
SHEET_NAME_MAX_LENGTH = 31


def generate_markdown_padding(
    orgin_text: str, length: int = STAFF_CATEGORY_LENGTH
) -> str:
    """Generate markdown padding for text to a specified length using &nbsp; for spaces"""
    space_size = length - len(orgin_text)
    if space_size <= 0:
        return orgin_text
    else:
        padding_text = "$" + "".ljust(space_size, "~") + "$"

    return orgin_text + " " + padding_text


def report_css_style():

    cell_y_padding = "8px"
    header_cell_y_padding = "10px"
    font_size = "9px"
    padding = "8px 8px 8px 8px"
    line_height = "1.2"
    margin_bottom = "0"
    margin_top = "0"
    font_family = "arial, sans-serif"
    table_header_bg_color = "#FFFFFF"
    title_bg_color = "#FFFFFF"
    # table_css = f"table {{width: 100%; border-collapse: collapse; font-size: {font_size} ; padding: {padding} {cell_y_padding}; }}"
    table_css = f"table {{width: 100%; font-size: {font_size} ; text-align:center; font-family: {font_family} ; border-spacing: 4px; border-collapse: collapse; padding: {padding} ; }}"
    # table_last_row = "table tr:last-child { font-weight: bold;color: black;}"
    # table_2th_css = "table th + th { text-align: center; }"
    table_2td_css = "table td + td { text-align: center; }"
    table_td_css = "table td {text-align: left}"
    # table_last_row2_css = "table tbody tr:last-child { font-weight: bold;color: black; padding: 10px; }"
    # table_last_row_css = "table tr:last-child { font-weight: bold;color: black; padding: 10px; }"
    # thead_th_css = f"thead th {{background-color: {table_header_bg_color} ;color: black;border: 0px solid #ddd;border-collapse: collapse; padding: {padding} ; }}"
    # tbody_td_first_child_css = f"tbody td:first-child {{ text-align: left; font-weight: bold;border-collapse: collapse; padding: {padding};}}"
    # tbody_td_css = f"tbody td {{ border-collapse: collapse;padding: {padding}; }}"
    h2_css = f"h2 {{text-align: left; background-color: {title_bg_color}; line-height: {line_height}; margin-bottom: {margin_bottom}; margin-top: {margin_top};}}"
    h3_css = f"h3 {{text-align: left; background-color: {title_bg_color}; line-height: {line_height}; margin-bottom: {margin_bottom}; margin-top: {margin_top};}}"
    h4_css = f"h4 {{text-align: left; background-color: {title_bg_color}; line-height: {line_height}; margin-bottom: {margin_bottom}; margin-top: {margin_top};}}"
    h5_css = f"h5 {{text-align: left; background-color: {title_bg_color}; line-height: {line_height}; margin-bottom: {margin_bottom}; margin-top: {margin_top};}}"
    css = (
        table_css
        + " "
        + table_td_css
        + " "
        + table_2td_css
        + " "
        + h2_css
        + " "
        + h3_css
        + " "
        + h4_css
        + " "
        + h5_css
    )

    return css


class ReturnCodes(Enum):
    """Enumeration for return codes"""

    ERROR_PROGRAM = -10
    ERROR_FILE_DATA_ERROR = -4
    ERROR_FILE_LOADING = -2
    ERROR_FILE_ERROR = -1
    ERROR = 0
    OK = 1
    OK_GEN_NEW_DATABASE = 2
    OK_UPDATE_DATABASE = 3


# set dataframe display format for float to 2 decimal places with $ sign
pd.options.display.float_format = "${:,.2f}".format


def header_processing_excel(header_text: str) -> list:
    return header_text.split(HEADER_SEPARATOR)


def header_processing_pdf(header_text: str, header_mark="##### ") -> str:

    header_substrings = header_text.split(HEADER_SEPARATOR)
    processed_header_strings = ""

    for s in header_substrings:
        processed_header_strings += header_mark + s.strip() + "\n"

    return processed_header_strings


def get_available_periods(
    data_available: list, start_year: int, start_month: int, max_number_of_month: int
):
    """Return data_available list based on start year/month and max number of months"""

    if max_number_of_month < 1 or max_number_of_month > 12:
        return ReturnCodes.ERROR_PROGRAM
        # raise "Number of months should be from 1 to 12"
    if start_year < 2000 or start_year > 3000:
        return ReturnCodes.ERROR_PROGRAM
        # raise "Start month should be between 2000 and 3000"
    if start_month < 1 or start_month > 12:
        return ReturnCodes.ERROR_PROGRAM
        # raise "Start month should be between 1 and 12"

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


def prepare_department_fte_trend_report(
    data_file_name: str,
    start_year: int,
    start_month: int,
    max_number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT,
):
    """Return markdown report content and css for fte trend report generation from database file"""

    try:
        # data_df_dict = pd.read_excel(data_file_name,sheet_name=None,header=0,dtype=object)
        data_df_dict = pd.read_excel(data_file_name, sheet_name=None, header=0)
    except Exception:
        return ReturnCodes.ERROR_FILE_LOADING
        # raise f"Error loading file {data_file_name}: {e}"

    available_periods = get_available_periods(
        data_df_dict.keys(), start_year, start_month, max_number_of_month
    )

    if len(available_periods) == 0:
        return ReturnCodes.ERROR_FILE_DATA_ERROR

    result_dict = {}
    results_order_dict = {}

    excel_df_dict = {}
    for period in available_periods:
        data_df = data_df_dict[period]

        data_df["allocation"] = data_df["allocation"].astype(float)
        period_df = data_df.groupby(["Staff Category"])["allocation"].sum()
        result_dict[period] = period_df

        result_order_df = data_df.drop_duplicates(subset=["Staff Category"]).loc[
            :, ["Staff Category", "staff category order"]
        ]
        dict_from_zipped = dict(
            zip(
                result_order_df["Staff Category"],
                result_order_df["staff category order"],
            )
        )

        if len(results_order_dict.keys()) == 0:
            results_order_dict = dict_from_zipped
        else:
            results_order_dict.update(dict_from_zipped)

    result_order_to_df = {}
    result_order_to_df["Staff Category"] = []
    result_order_to_df["staff category order"] = []
    order = 1
    for k, v in sorted(results_order_dict.items(), key=lambda x: (x[1], x[0])):
        result_order_to_df["Staff Category"].append(k)
        result_order_to_df["staff category order"].append(order)
        order += 1

    result = pd.DataFrame(result_dict)
    results_order_df = pd.DataFrame.from_dict(result_order_to_df)

    sorted_result_df = result.join(
        results_order_df.set_index(["Staff Category", "staff category order"]),
        how="inner",
    )

    sorted_result_df.reset_index(inplace=True)
    sorted_result_df.set_index("staff category order", inplace=True)

    sorted_result_df.sort_index(inplace=True)

    sorted_result_df.set_index("Staff Category", inplace=True)

    sorted_result_df.loc["Total"] = sorted_result_df.sum(numeric_only=True)

    sorted_result_df.reset_index(inplace=True)

    sorted_result_df = sorted_result_df.round(2).astype(str)

    sorted_result_dict = sorted_result_df.to_dict(orient="index")

    # sorted_result_dict = sorted_result_df.round(2).astype(str).to_dict(orient="index")

    excel_df_dict["fte"] = {"data": sorted_result_df}

    markdown_table_data = []

    empty_v = {}
    for k, v in sorted_result_dict.items():
        for key in v.keys():
            empty_v[key] = ""
        break

    for k, v in sorted_result_dict.items():
        for key in v.keys():
            if key != "Staff Category":
                v[key] = f"{float(v[key]):,.2f}"
        if v["Staff Category"] == "Total":
            markdown_table_data.append(empty_v)
            markdown_table_data.append(empty_v)
            for key in v.keys():
                v[key] = "**" + v[key] + "**"
        markdown_table_data.append(v)

    markdown = (
        markdown_table(markdown_table_data)
        .set_params(row_sep="markdown", quote=False)
        .get_markdown()
    )
    markdown = markdown.replace("nan", "-")

    css = report_css_style()
    md = {}
    md["content"] = markdown
    md["css"] = css
    return_md = [md]
    return {"md": return_md, "excel_df": excel_df_dict}


def prepare_department_headcount_trend_report(
    data_file_name: str,
    start_year: int,
    start_month: int,
    max_number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT,
):
    """Return markdown report content and css for department headcount trend report generation from database file"""

    try:
        # data_df_dict = pd.read_excel(data_file_name,sheet_name=None,header=0,dtype=object)
        data_df_dict = pd.read_excel(data_file_name, sheet_name=None, header=0)
    except Exception:
        return ReturnCodes.ERROR_FILE_LOADING

    available_periods = get_available_periods(
        data_df_dict.keys(), start_year, start_month, max_number_of_month
    )

    if len(available_periods) == 0:
        return ReturnCodes.ERROR_FILE_DATA_ERROR

    result_dict = {}
    results_order_dict = {}
    excel_df_dict = {}

    for period in available_periods:
        data_df = data_df_dict[period]

        period_df = (
            data_df.drop_duplicates(subset=["staff_number"])
            .groupby(["Staff Category"])
            .size()
        )
        result_dict[period] = period_df

        result_order_df = data_df.drop_duplicates(subset=["Staff Category"]).loc[
            :, ["Staff Category", "staff category order"]
        ]
        dict_from_zipped = dict(
            zip(
                result_order_df["Staff Category"],
                result_order_df["staff category order"],
            )
        )
        if len(results_order_dict.keys()) == 0:
            results_order_dict = dict_from_zipped
        else:
            results_order_dict.update(dict_from_zipped)

    result_order_to_df = {}
    result_order_to_df["Staff Category"] = []
    result_order_to_df["staff category order"] = []
    order = 1
    for k, v in sorted(results_order_dict.items(), key=lambda x: (x[1], x[0])):
        result_order_to_df["Staff Category"].append(k)
        result_order_to_df["staff category order"].append(order)
        order += 1

    result = pd.DataFrame(result_dict)
    results_order_df = pd.DataFrame.from_dict(result_order_to_df)

    sorted_result_df = pd.merge(
        result, results_order_df, on="Staff Category", how="inner"
    )
    sorted_result_df.set_index("staff category order", inplace=True)

    sorted_result_df.sort_index(inplace=True)
    sorted_result_df.set_index("Staff Category", inplace=True)

    sorted_result_df.loc["Total"] = sorted_result_df.sum(numeric_only=True)

    sorted_result_df.reset_index(inplace=True)

    sorted_result_df = sorted_result_df.round(2).astype(str)

    sorted_result_dict = sorted_result_df.to_dict(orient="index")

    excel_df_dict["headcount"] = {"data": sorted_result_df}

    markdown_table_data = []

    empty_v = {}
    for k, v in sorted_result_dict.items():
        for key in v.keys():
            empty_v[key] = ""
        break

    for k, v in sorted_result_dict.items():
        for key in v.keys():
            if key != "Staff Category":
                v[key] = f"{float(v[key]):,.0f}"
        if v["Staff Category"] == "Total":
            markdown_table_data.append(empty_v)
            markdown_table_data.append(empty_v)
            for key in v.keys():
                v[key] = "**" + v[key] + "**"
        markdown_table_data.append(v)

    markdown = (
        markdown_table(markdown_table_data)
        .set_params(row_sep="markdown", quote=False)
        .get_markdown()
    )
    markdown = markdown.replace("nan", "-")

    css = report_css_style()
    md = {}
    md["content"] = markdown
    md["css"] = css
    return_md = [md]
    return {"md": return_md, "excel_df": excel_df_dict}


def prepare_department_fte_costcentre_report(
    data_file_name: str,
    start_year: int,
    start_month: int,
    max_number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT,
):
    """Return markdown report content and css for department fte report generation from database file"""

    try:
        # data_df_dict = pd.read_excel(data_file_name,sheet_name=None,header=0,dtype=object)
        data_df_dict = pd.read_excel(data_file_name, sheet_name=None, header=0)
    except Exception:
        return ReturnCodes.ERROR_FILE_LOADING

    available_periods = get_available_periods(
        data_df_dict.keys(), start_year, start_month, max_number_of_month
    )

    if type(available_periods) is ReturnCodes:
        return available_periods
    if len(available_periods) == 0:
        return ReturnCodes.ERROR_FILE_DATA_ERROR

    return_md = []
    all_costcentre_result_dict = {}
    cost_centre_code_dict = {}
    for period in available_periods:
        data_df = data_df_dict[period]

        cost_centres = data_df["cost centre name"].copy().drop_duplicates().to_list()

        # data_df["Staff Category duplicate"] = data_df["Staff Category"]

        for c in cost_centres:
            if c not in all_costcentre_result_dict.keys():
                all_costcentre_result_dict[c] = {
                    period: data_df[data_df["cost centre name"] == c]
                }
            else:
                all_costcentre_result_dict[c][period] = data_df[
                    data_df["cost centre name"] == c
                ]
            cost_centre_code_dict[c] = all_costcentre_result_dict[c][period][
                "cost centre code"
            ].iloc[0]

    excel_df_dict = {}
    for cost_centre, v in sorted(all_costcentre_result_dict.items()):
        result_dict = {}
        results_order_dict = {}
        for period, target_df in v.items():
            new_target_df = target_df.copy()

            new_target_df["allocation"] = new_target_df["allocation"].astype(float)
            period_df = new_target_df.groupby(["Staff Category", "Rank"])[
                "allocation"
            ].sum()
            result_dict[period] = period_df

            result_order_df = data_df.drop_duplicates(subset=["Staff Category"]).loc[
                :, ["Staff Category", "staff category order"]
            ]
            dict_from_zipped = dict(
                zip(
                    result_order_df["Staff Category"],
                    result_order_df["staff category order"],
                )
            )
            if len(results_order_dict.keys()) == 0:
                results_order_dict = dict_from_zipped
            else:
                results_order_dict.update(dict_from_zipped)

        result_order_to_df = {}
        result_order_to_df["Staff Category"] = []
        result_order_to_df["staff category order"] = []
        order = 1
        for k, v in sorted(results_order_dict.items(), key=lambda x: (x[1], x[0])):
            result_order_to_df["Staff Category"].append(k)
            result_order_to_df["staff category order"].append(order)
            order += 1

        result = pd.DataFrame(result_dict)
        results_order_df = pd.DataFrame.from_dict(result_order_to_df)

        sorted_result_df = result.join(
            results_order_df.set_index(["Staff Category", "staff category order"]),
            how="inner",
        )

        sorted_result_df.reset_index(inplace=True)
        sorted_result_df.set_index("staff category order", inplace=True)

        sorted_result_df.sort_index(inplace=True)

        sorted_result_df.set_index("Staff Category", inplace=True)

        sorted_result_df["Rank"] = sorted_result_df["Rank"].astype(str)

        sorted_result_df.loc["Total"] = sorted_result_df.sum(numeric_only=True)

        sorted_result_df.reset_index(inplace=True)

        sorted_result_df = sorted_result_df.round(2).astype(str)

        sorted_result_dict = (
            # sorted_result_df.round(2).astype(str).to_dict(orient="index")
            sorted_result_df.to_dict(orient="index")
        )

        # excel_df_list.append({'data' : sorted_result_df})
        excel_df_dict[cost_centre] = {"data": sorted_result_df}

        markdown_table_data = []

        empty_v = {}
        for k, v in sorted_result_dict.items():
            for key in v.keys():
                empty_v[key] = ""
            break

        last_staff_category = ""
        for k, v in sorted_result_dict.items():
            for key in v.keys():
                if key != "Staff Category" and key != "Rank":
                    v[key] = f"{float(v[key]):,.1f}"
            if v["Staff Category"] == "Total":
                last_staff_category = ""
                markdown_table_data.append(empty_v)
                markdown_table_data.append(empty_v)
                for key in v.keys():
                    v[key] = "**" + v[key] + "**"
                v["Rank"] = ""
            elif v["Staff Category"] == last_staff_category:
                v["Staff Category"] = ""
            else:
                last_staff_category = v["Staff Category"]

            markdown_table_data.append(v)

        markdown = (
            markdown_table(markdown_table_data)
            .set_params(row_sep="markdown", quote=False)
            .get_markdown()
        )
        markdown = markdown.replace("nan", "-")

        markdown_with_costcentre_name = f"##### Cost Centre : {cost_centre} ({cost_centre_code_dict[cost_centre]})<p>\n\n{markdown}"

        css = report_css_style()
        result_md = {}
        result_md["content"] = markdown_with_costcentre_name
        result_md["css"] = css
        return_md.append(result_md)
    return {"md": return_md, "excel_df": excel_df_dict}


def generate_pdf_report(report_name: str, content: list, title: str = "Report"):
    """Generate PDF report from markdown content and css list input from prepare report functions"""

    # header = f"## {title}"
    header = header_processing_pdf(title, header_mark="##### ")

    pdf = MarkdownPdf()
    for c in content:
        pdf.add_section(
            Section(header + "\n\n\n" + c["content"], paper_size="A4-L", toc=False),
            user_css=c["css"],
        )
    pdf.save(report_name + ".pdf")


def check_file_header(df: pd.DataFrame, expected_headers: list) -> list:
    """check header are available in dataframe"""

    missing_headers = []
    for h in expected_headers:
        if h not in df.columns:
            missing_headers.append(h)

    return missing_headers


def process_source_data(excelfile: str) -> int:
    """Process the source excel file and return data dictionary or error code"""

    # read sheet 1
    try:
        file_base_data_df = pd.read_excel(
            excelfile, sheet_name=0, header=0, dtype=object
        )
        # file_base_data_df = pd.read_excel(excelfile,sheet_name=0,header=0)
    except Exception:
        return ReturnCodes.ERROR_FILE_ERROR

    header = [
        "StaffNo",
        "Rank",
        "Section",
        "Staff Category",
        "FTE",
        "Default Cost Centre",
    ]
    missing_headers = check_file_header(file_base_data_df, header)
    if len(missing_headers) > 0:
        return ReturnCodes.ERROR_FILE_ERROR

    base_data_df = file_base_data_df[header]
    clean_base_data_df = base_data_df.dropna(how="all")

    # Remove last row , if Rank is empty
    last_row_df = clean_base_data_df.tail(1)
    if len(last_row_df.dropna(subset=["Rank"])) == 0:
        clean_base_data_df = clean_base_data_df.head(len(clean_base_data_df) - 1)

    file_base_records_count: int = len(base_data_df.index)
    clean_base_records_count: int = len(clean_base_data_df.index)

    if file_base_records_count != clean_base_records_count:
        pass
        # print(f"Base data had {file_base_records_count - clean_base_records_count} empty rows removed.")

    # check for duplicate StaffNo in base data
    if len(pd.unique(clean_base_data_df["StaffNo"])) != len(clean_base_data_df):
        # print("Warning: Duplicate Staff Numbers found in Base Data!")
        return ReturnCodes.ERROR_FILE_DATA_ERROR

    # set the right data types for data Series
    clean_base_data_df["FTE"] = clean_base_data_df["FTE"].astype(float)
    clean_base_data_df["StaffNo"] = (
        clean_base_data_df["StaffNo"].astype(int).astype(str)
    )

    rank_cat = pd.DataFrame(
        clean_base_data_df["Rank"] + "\t" + clean_base_data_df["Staff Category"]
    )
    rank_cat.columns = ["cat_info"]
    unique_rank_cat = pd.unique(rank_cat["cat_info"]).tolist()
    unique_rank_cat_dict = {}
    for i in unique_rank_cat:
        cats = i.split("\t")
        unique_rank_cat_dict[cats[0]] = cats[1]
    if DEBUG:
        print(unique_rank_cat_dict)
        print(type(unique_rank_cat_dict))

    # read sheet 2
    try:
        # file_expand_data_df = pd.read_excel(excelfile,sheet_name=1,header=1,dtype=object)
        file_expand_data_df = pd.read_excel(excelfile, sheet_name=1, header=1)
    except Exception:

        return ReturnCodes.ERROR_FILE_ERROR

    first_row_df = file_expand_data_df.head(1)
    if len(first_row_df.dropna(subset=["Rank"])) == 0:
        clean_base_data_df = clean_base_data_df.head(len(clean_base_data_df) - 1)

    header = ["StaffNo", "Rank", "CCode", "CostCentre", "Allocated Percentage"]
    missing_headers = check_file_header(file_expand_data_df, header)
    if len(missing_headers) > 0:

        return ReturnCodes.ERROR_FILE_ERROR

    expand_data_df = file_expand_data_df[header]
    clean_expand_data_df = expand_data_df.dropna(how="all")

    file_expand_records_count: int = len(file_expand_data_df.index)
    clean_expand_records_count: int = len(clean_expand_data_df.index)

    new_clean_expand_data_df = clean_expand_data_df.copy()
    new_clean_expand_data_df["Allocated Percentage"] = new_clean_expand_data_df[
        "Allocated Percentage"
    ].astype(float)
    new_clean_expand_data_df["StaffNo"] = (
        new_clean_expand_data_df["StaffNo"].astype(int).astype(str)
    )
    new_clean_expand_data_df["Allocated Percentage"] = (
        new_clean_expand_data_df["Allocated Percentage"] / 100.0
    )
    new_clean_expand_data_df["CCode"] = (
        new_clean_expand_data_df["CCode"].astype(int).astype(str)
    )
    clean_expand_data_df = new_clean_expand_data_df

    if DEBUG:
        print("clean_expand_data_df ------ ")
        print(clean_expand_data_df.head(5))

    # read sheet 3, cost center information
    try:
        file_cost_centre_data_df = pd.read_excel(
            excelfile, sheet_name=2, header=0, dtype=object
        )
        # file_cost_centre_data_df = pd.read_excel(excelfile,sheet_name=2,header=0)
    except Exception:
        # print(f"Error loading base sheet 3: {e}")
        return ReturnCodes.ERROR_FILE_ERROR

    header = ["Value", "Description", "Enabled/ Disabled"]
    missing_headers = check_file_header(file_cost_centre_data_df, header)
    if len(missing_headers) > 0:
        # print(f"Error: sheet 3 Missing expected column '{", ".join(missing_headers)}' in cost centre  data sheet.")
        return ReturnCodes.ERROR_FILE_ERROR

    cost_centre_data_df = file_cost_centre_data_df[header]
    clean_cost_centre_data_df = cost_centre_data_df.dropna(how="all")
    clean_cost_centre_data_df = clean_cost_centre_data_df[
        clean_cost_centre_data_df["Enabled/ Disabled"] == "Enabled"
    ]

    file_cost_centre_records_count: int = len(file_cost_centre_data_df.index)
    clean_cost_centre_records_count: int = len(clean_cost_centre_data_df.index)

    if DEBUG:
        if file_cost_centre_records_count != clean_cost_centre_records_count:
            print(
                f"Cost centres data had {file_cost_centre_records_count - clean_cost_centre_records_count } 'Disabled' rows removed."
            )

    # get the cost centre information
    clean_cost_centre_dict = clean_cost_centre_data_df.to_dict(orient="index")

    cost_centre_info = {}
    for k, v in clean_cost_centre_dict.items():
        cost_centre_info[str(v["Value"])] = v["Description"]

    if DEBUG:
        first_line = True
        for k, v in cost_centre_info.items():
            if first_line:
                print(f"{k} : {v}", end="")
                first_line = False
            else:
                print(", ")
                print(f"{k} : {v}", end="")
        print()
    # read sheet 4 Staff Category Order
    try:
        # file_staff_category_order_data_df = pd.read_excel(excelfile,sheet_name=3,header=0,dtype=object)
        file_staff_category_order_data_df = pd.read_excel(
            excelfile, sheet_name=3, header=0
        )
        has_staff_category_order_data = True
    except Exception:
        has_staff_category_order_data = False

    if has_staff_category_order_data:
        header = ["Staff Category", "Order"]
        missing_headers = check_file_header(file_staff_category_order_data_df, header)
        if len(missing_headers) > 0:
            return ReturnCodes.ERROR_FILE_ERROR
            # return f"Error: Missing expected column '{", ".join(missing_headers)}' in staff category data sheet."

        staff_category_order_data_df = file_staff_category_order_data_df[header]
        clean_staff_category_order_data_df = staff_category_order_data_df.dropna(
            how="all"
        )

        file_staff_category_order_records_count: int = len(
            file_staff_category_order_data_df.index
        )
        clean_staff_category_order_records_count: int = len(
            clean_staff_category_order_data_df.index
        )

        if DEBUG:
            if (
                file_staff_category_order_records_count
                != clean_staff_category_order_records_count
            ):
                print(
                    f"Expand data had {file_expand_records_count - clean_expand_records_count} empty rows removed."
                )

        clean_staff_category_order_data_dict = (
            clean_staff_category_order_data_df.to_dict(orient="index")
        )

        staff_category_order = {}
        for v in clean_staff_category_order_data_dict.values():
            staff_category_order[v["Staff Category"]] = v["Order"]

        count = 1
        staff_category_order_dict = {}
        for k, v in sorted(staff_category_order.items(), key=lambda x: (x[1], x[0])):
            staff_category_order_dict[k] = count
            count += 1

    else:

        sorted_staff_category_order_list = sorted(
            pd.unique(clean_base_data_df["Staff Category"]).tolist()
        )
        # sorted_staff_category_order_list = sorted(staff_category_order_list)
        staff_category_order_dict = {}
        for i in range(len(sorted_staff_category_order_list)):
            staff_category_order_dict[sorted_staff_category_order_list[i]] = i + 1

    # expand the list
    clean_base_data_df = clean_base_data_df.set_index("StaffNo")
    clean_base_dict = clean_base_data_df.to_dict(orient="index")
    if DEBUG:
        print("clean_base_dict ------ ")
        print(clean_base_dict)
    # clean_base_dict has key as StaffNo, value as dict of other columns
    clean_expand_dict = clean_expand_data_df.to_dict(orient="index")
    # clean_expand_dict has key as index, value as dict of other columns, including StaffNo as StaffNo is not unique here

    expanded_entries = []
    staff_rank_category = {}
    issue_staff_numbers_not_in_base = set()
    issue_staff_numbers_fte_not_100_in_expand = []

    unique_staff_in_base = {}
    unique_staff_in_expand = {}
    for k, v in clean_base_dict.items():
        unique_staff_in_base[str(k)] = v["FTE"]

    for k, v in clean_expand_dict.items():

        staff_number = str(v["StaffNo"])
        if staff_number not in unique_staff_in_expand.keys():
            unique_staff_in_expand[staff_number] = v["Allocated Percentage"]
        else:
            unique_staff_in_expand[staff_number] += v["Allocated Percentage"]

        if DEBUG:
            print(f"Processing expand record for staff number {staff_number}")
            print(v)
        if staff_number in clean_base_dict.keys():
            if DEBUG:
                print(f"Found staff number {staff_number} in base data.")
            staff_rank_category[staff_number] = clean_base_dict[staff_number][
                "Staff Category"
            ]
            del clean_base_dict[staff_number]

        if staff_number in unique_staff_in_base.keys():
            if (
                v["Rank"] in unique_rank_cat_dict.keys()
                and unique_rank_cat_dict[v["Rank"]] in staff_category_order_dict.keys()
            ):
                expanded_item = {
                    "staff_number": staff_number,
                    "Rank": v["Rank"],
                    "Staff Category": unique_rank_cat_dict[v["Rank"]],
                    "staff category order": staff_category_order_dict[
                        unique_rank_cat_dict[v["Rank"]]
                    ],
                    "cost centre code": str(v["CCode"]).zfill(3),
                    "cost centre name": cost_centre_info[str(v["CCode"]).zfill(3)],
                    "allocation": v["Allocated Percentage"]
                    * unique_staff_in_base[staff_number],
                }
                expanded_entries.append(expanded_item)

                if DEBUG:
                    print(f"Adding expanded record for staff number {staff_number}")
            else:
                return ReturnCodes.ERROR_FILE_DATA_ERROR
        else:
            issue_staff_numbers_not_in_base.add(staff_number)
            # found record in expand data but not in base data. It is not counted as error, just skip it.
            if DEBUG:
                print(
                    f"Error : Rank {v['Rank']} not found in base data for staff number {staff_number}."
                )

    for k, v in clean_base_dict.items():
        clean_item = {
            "staff_number": k,
            "Rank": v["Rank"],
            "Staff Category": v["Staff Category"],
            "staff category order": staff_category_order_dict[v["Staff Category"]],
            "cost centre code": str(v["Default Cost Centre"]).zfill(3),
            "cost centre name": cost_centre_info[
                str(v["Default Cost Centre"]).zfill(3)
            ],
            "allocation": v["FTE"],
        }
        expanded_entries.append(clean_item)

        # if DEBUG:
        #    print(f"Adding base only record for staff number {k}")
        #    print(clean_item)

        # expanded_entries.append({'staff_number' : k, 'rank' : v['Rank'], 'Staff Category' : v['Staff Category'], 'staff category order' : staff_category_order_dict[v['Staff Category']],'cost centre code' : str(v['Default Cost Centre']).zfill(3), 'cost centre name' : cost_centre_info[str(v['Default Cost Centre']).zfill(3)], 'allocation' : v['FTE']})

    result_df = pd.DataFrame(expanded_entries)
    if DEBUG:
        print(f"Total records processed: {len(result_df.index)}")
        print(result_df)

    result_dict = {"hr_fte_df": result_df}
    for k, v in unique_staff_in_expand.items():
        if v != 1.0:
            issue_staff_numbers_fte_not_100_in_expand.append(f"{k}({v})")
            if DEBUG:
                print(
                    f"Staff number {k} has total FTE allocation of {v} in expand data."
                )
    result_dict["issue_staff_numbers_not_in_base"] = sorted(
        list(issue_staff_numbers_not_in_base)
    )
    result_dict["issue_expand_staff_fte_not_1"] = (
        issue_staff_numbers_fte_not_100_in_expand
    )
    if DEBUG:
        print(
            f"Staff numbers with FTE not equal to 1 in expand data: {issue_staff_numbers_fte_not_100_in_expand}"
        )
        print(
            f"Staff numbers not in base: {result_dict['issue_staff_numbers_not_in_base']}"
        )

    return result_dict


def clean_sheet_name(sheet_name: str) -> str:
    mytable = str.maketrans("\\/*?:[].", "________")

    return sheet_name.translate(mytable)[:SHEET_NAME_MAX_LENGTH]


def generate_excel_fr_df(
    # reportname: str, sheet_names: list[str], result_df: pd.DataFrame
    reportname: str,
    input_data_dict: dict,
):
    reportname = reportname + ".xlsx"

    if not os.path.exists(reportname):
        with pd.ExcelWriter(f"{reportname}", mode="w") as writer:
            # df1.to_excel(writer, sheet_name='Sheet_name_3')
            for sheet_name, data_df_dict in input_data_dict.items():
                
                clean_name = clean_sheet_name(sheet_name)
                
                if "header" in data_df_dict.keys():
                    data_df_dict["header"].style.set_properties(**{'text-align': 'center', 'vertical-align': 'middle'}).to_excel(
                        writer, sheet_name=f"{clean_name}", index=False, header=False,
                    )
                    

                    if "data" in data_df_dict.keys():
                        data_df_dict["data"].style.set_properties(**{'text-align': 'center', 'vertical-align': 'middle'}).to_excel(
                            writer,
                            sheet_name=f"{clean_name}",
                            index=False,
                            startrow=writer.sheets[clean_name].max_row, header=True,
                        )
                elif "data" in data_df_dict.keys():
                    data_df_dict["data"].style.set_properties(**{'text-align': 'center', 'vertical-align': 'middle'}).to_excel(
                        writer, index=False, sheet_name=f"{clean_name}"
                    )
        
        return ReturnCodes.OK_GEN_NEW_DATABASE
    else:
        if DEBUG:
            print(f"report file {reportname} existed")
        return ReturnCodes.ERROR_FILE_ERROR


def generate_department_fte_summary_report(
    fte_data_file_name: str,
    summary_report_file_name: str,
    report_title: str,
    start_year: int,
    start_month: int,
    number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT,
):
    """Generate department FTE summary report from database file"""

    department_fte_trend_content = prepare_department_fte_trend_report(
        fte_data_file_name, start_year, start_month, number_of_month
    )
    if type(department_fte_trend_content) is ReturnCodes:
        return department_fte_trend_content
    elif type(department_fte_trend_content) is dict:
        period = (
            f"{str(start_year)}"
            if start_month == 1
            else f"{str(start_year)}/{str(start_year+1)}"
        )
        report_title = f"{report_title} {period}"

        if "md" in department_fte_trend_content.keys():
            generate_pdf_report(
                summary_report_file_name,
                department_fte_trend_content["md"],
                report_title,
            )
        if "excel_df" in department_fte_trend_content.keys():
            title_lines = header_processing_excel(f"{report_title}")
            sheet_header = {"title": title_lines}

            header_df = pd.DataFrame(sheet_header)

            for k, v in department_fte_trend_content["excel_df"].items():
                department_fte_trend_content["excel_df"][k]["header"] = header_df

            generate_excel_fr_df(
                summary_report_file_name, department_fte_trend_content["excel_df"]
            )
    else:
        if DEBUG:
            print(
                f"Error: department_headcount_trend_content is '{department_fte_trend_content}'"
            )
        return ReturnCodes.ERROR_PROGRAM

    return ReturnCodes.OK


def generate_department_headcount_summary_report(
    fte_data_file_name: str,
    summary_report_file_name: str,
    report_title: str,
    start_year: int,
    start_month: int,
    number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT,
):
    """Generate department headcount summary report from database file"""

    department_headcount_trend_content = prepare_department_headcount_trend_report(
        fte_data_file_name, start_year, start_month, number_of_month
    )
    if type(department_headcount_trend_content) is ReturnCodes:
        return department_headcount_trend_content
    elif type(department_headcount_trend_content) is dict:
        period = (
            f"{str(start_year)}"
            if start_month == 1
            else f"{str(start_year)}/{str(start_year+1)}"
        )
        report_title = f"{report_title} {period}"

        if "md" in department_headcount_trend_content.keys():

            generate_pdf_report(
                summary_report_file_name,
                department_headcount_trend_content["md"],
                report_title,
            )
        if "excel_df" in department_headcount_trend_content.keys():
            title_lines = header_processing_excel(f"{report_title}")
            sheet_header = {"title": title_lines}
            
            header_df = pd.DataFrame(sheet_header)

            for k, v in department_headcount_trend_content["excel_df"].items():
                department_headcount_trend_content["excel_df"][k]["header"] = header_df

            generate_excel_fr_df(
                summary_report_file_name, department_headcount_trend_content["excel_df"]
            )

    else:
        if DEBUG:
            print(
                f"Error: department_headcount_trend_content is '{department_headcount_trend_content}'"
            )
        return ReturnCodes.ERROR_PROGRAM

    return ReturnCodes.OK


def generate_department_fte_costcentre_report(
    fte_data_file_name: str,
    costcentre_report_file_name: str,
    report_title: str,
    start_year: int,
    start_month: int,
    number_of_month: int = MAX_NUMBER_MONTH_IN_REPORT,
):
    """Generate department fte report with costcentre breakdown from database file"""

    department_fte_costcentre_content = prepare_department_fte_costcentre_report(
        fte_data_file_name, start_year, start_month, number_of_month
    )
    # print(department_fte_trend_content)
    if type(department_fte_costcentre_content) is ReturnCodes:
        return department_fte_costcentre_content
    elif type(department_fte_costcentre_content) is dict:
        period = (
            f"{str(start_year)}"
            if start_month == 1
            else f"{str(start_year)}/{str(start_year+1)}"
        )
        report_title = f"{report_title} {period}"

        if "md" in department_fte_costcentre_content.keys():
            generate_pdf_report(
                costcentre_report_file_name,
                department_fte_costcentre_content["md"],
                report_title,
            )
        if "excel_df" in department_fte_costcentre_content.keys():
            title_lines = header_processing_excel(f"{report_title}")
            sheet_header = {"title": title_lines}
            
            header_df = pd.DataFrame(sheet_header)

            for k, v in department_fte_costcentre_content["excel_df"].items():
                department_fte_costcentre_content["excel_df"][k]["header"] = header_df

            generate_excel_fr_df(
                costcentre_report_file_name,
                department_fte_costcentre_content["excel_df"]
            )
    else:
        if DEBUG:
            print(
                f"Error: department_headcount_trend_content is '{department_fte_costcentre_content}'"
            )
        return ReturnCodes.ERROR_PROGRAM

    return ReturnCodes.OK


if __name__ == "__main__":
    # Load the dataset

    database_file = "HR_FTE_Database.xlsx"
    start_year = 2025
    start_month = 7

    department_headcount_summary_report_file_name = (
        "HR_department_headcount_summary_report.pdf"
    )
    department_headcount_summary_report_title = "Yearly Department Headcount Trend.pdf"

    generate_department_headcount_summary_report(
        database_file,
        department_headcount_summary_report_file_name,
        department_headcount_summary_report_title,
        start_year,
        start_month,
        12,
    )
