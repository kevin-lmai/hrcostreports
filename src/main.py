"""The main module of the HR Cost Reporting Application using Flet framework."""

import os
from datetime import datetime

import flet
from dateutil.relativedelta import relativedelta
from flet import (
    Button,
    ButtonStyle,
    Card,
    Colors,
    Column,
    Container,
    ControlState,
    CrossAxisAlignment,
    FilePicker,
    FontWeight,
    IconButton,
    Icons,
    MainAxisAlignment,
    NavigationRail,
    NavigationRailDestination,
    NavigationRailLabelType,
    Page,
    Row,
    Text,
    VerticalDivider,
)

from dataprocess import (
    HEADER_SEPARATOR,
    ReturnCodes,
    generate_department_fte_costcentre_report,
    generate_department_fte_summary_report,
    generate_department_headcount_summary_report,
    generate_excel_fr_df,
    process_source_data,
)

APP_VERSION = 'v0.9.8'
# FTE Data Upload related global variables
data_name = None
data_directory = None
database_file_directory = None
database_file_name = None
fte_data_date = None

# Reports Generation related global variables
database_file_saved = False
saved_database_file_directory = None
saved_database_name = None

report_start_date = None

company_name = "CUHK Medical Centre Limited"
financial_year_header = "Financial Year: "

department_fte_summary_report_file_name = "HR_department_fte_summary_report"
department_fte_summary_report_title = "Full Time Equivalent (FTE) - Total"

department_headcount_summary_report_file_name = "HR_department_headcount_summary_report"
department_headcount_summary_report_title = "Headcount - Total"

department_fte_costcentre_report_file_name = "HR_department_fte_costcentres_report"
department_fte_costcentre_report_title = "Full Time Equivalent (FTE) by Department"


def set_button_text(button: Button, value: str):
    """Set text shown inside a Flet Button."""
    if isinstance(button.content, Text):
        button.content.value = value
    else:
        button.content = Text(value)


def init_data_upload_setup():
    """Initialize the data upload, database setup parameters."""

    global data_name
    global data_directory
    global database_file_directory
    global database_file_name
    global fte_data_date

    data_name = None
    data_directory = None
    database_file_directory = None
    database_file_name = "HR_FTE_Database.xlsx"
    fte_data_date = datetime.now() - relativedelta(months=1)


init_data_upload_setup()


class DesktopAppLayout(Row):
    """A desktop app layout with a menu on the left."""

    def __init__(self, title, page, pages, *args, window_size=(800, 600), **kwargs):
        super().__init__(*args, **kwargs)

        self.app_page = page
        self.pages = pages
        self.expand = True

        self.navigation_items = [navigation_item for navigation_item, _ in pages]
        self.navigation_rail = self.build_navigation_rail()

        self.update_destinations()
        self._menu_extended = True
        self.navigation_rail.extended = True

        self.menu_panel = Row(controls=[self.navigation_rail], spacing=0, tight=True)

        page_contents = [page_content for _, page_content in pages]
        self.content_area = Column(page_contents, expand=True)

        self._was_portrait = self.is_portrait()
        self._panel_visible = self.is_landscape()

        self.set_content()
        self._change_displayed_page()
        self.app_page.on_resize = self.handle_resize

        self.window_size = window_size
        self.app_page.window.width, self.app_page.window.height = self.window_size
        self.app_page.title = title

    def select_page(self, page_number):
        """Select the page to be displayed by page number index."""
        self.navigation_rail.selected_index = page_number
        self._change_displayed_page()

    def _navigation_change(self, e):
        """Handle navigation rail index change event."""
        self._change_displayed_page()
        self.app_page.update()

    def _change_displayed_page(self):
        """Change the displayed page based on selected navigation rail index."""
        page_number = self.navigation_rail.selected_index
        for i, content_page in enumerate(self.content_area.controls):
            content_page.visible = page_number == i

    def build_navigation_rail(self):
        """Build the navigation rail for the desktop app layout."""
        return NavigationRail(
            selected_index=0,
            label_type=NavigationRailLabelType.NONE,
            on_change=self._navigation_change,
            bgcolor=Colors.WHITE,
            indicator_color=Colors.BLUE,
            selected_label_text_style=flet.TextStyle(color=Colors.BLUE),
            unselected_label_text_style=flet.TextStyle(color=Colors.BLACK),
            group_alignment=0.0,
        )

    def update_destinations(self):
        """Update the navigation rail destinations."""
        self.navigation_rail.destinations = self.navigation_items
        self.navigation_rail.label_type = NavigationRailLabelType.ALL

    def handle_resize(self, e):
        """Handle the page resize event."""
        pass

    def set_content(self):
        """Set the content layout of the desktop app."""
        self.controls = [
            self.menu_panel,
            VerticalDivider(width=1, color=Colors.RED),
            self.content_area,
        ]
        self.update_destinations()
        self.navigation_rail.extended = self._menu_extended
        self.menu_panel.visible = self._panel_visible

    def is_portrait(self) -> bool:
        """Determine if the window/display is in portrait orientation."""
        return self.app_page.height >= self.app_page.width

    def is_landscape(self) -> bool:
        """Determine if the window/display is in landscape orientation."""
        return self.app_page.width > self.app_page.height


def main(page: Page):
    """The main function to setup Flet application page and layout."""

    page.window.width = 900
    page.window.height = 450
    page.update()

    disabled_button_style = ButtonStyle(
        bgcolor={ControlState.DISABLED: Colors.GREY_300},
        color={ControlState.DISABLED: Colors.GREEN},
    )

    init_fte_upload_status_content = "Data not uploaded"
    status_text_fte_upload = Text(
        init_fte_upload_status_content, bgcolor=Colors.YELLOW, color=Colors.BLACK
    )

    datafile_upload_button_prompt = (
        f"Please Select {fte_data_date.strftime('%Y / %m')}  FTE Data File"
    )
    optional_report_upload_button_prompt = "Optionally Set Database File"
    update_database_button_prompt = "Update Database File"
    restart_button_fte_upload_prompt = "Restart Upload Process"

    fte_data_month_text = Text(fte_data_date.strftime("%Y / %m"), color=Colors.BLUE)

    fte_data_upload_button = Button(
        content=Text(datafile_upload_button_prompt),
        icon=Icons.UPLOAD_FILE,
        on_click=lambda _: page.run_task(pick_data_files_click),
        style=disabled_button_style,
    )

    optional_report_upload_button = Button(
        content=Text(optional_report_upload_button_prompt),
        icon=Icons.ADD_BOX,
        on_click=lambda _: page.run_task(pick_optional_report_files_click),
        disabled=True,
        style=disabled_button_style,
    )

    def update_database(e):
        """Update or create database file from uploaded data file."""

        global database_file_saved
        global saved_database_file_directory
        global saved_database_name

        data_period = f"{str(fte_data_date.year)}{str(fte_data_date.month).zfill(2)}"
        datafile = os.path.join(data_directory, data_name)
        report_file = os.path.join(database_file_directory, database_file_name)
        result_dict = process_source_data(datafile)
        if result_dict["return_code"] <= 0:
            result = result_dict
        else:
            input_excel_data_dict = {data_period: {"data": result_dict["hr_fte_df"]}}
            result = generate_excel_fr_df(report_file, input_excel_data_dict)

        if result["return_code"] == ReturnCodes.OK_UPDATE_DATABASE.value:
            status_text_fte_upload.value = f"Congratulation!!\nDatabase file {os.path.join(database_file_directory, database_file_name)} was updated."
            database_file_saved = True
            saved_database_file_directory = database_file_directory
            saved_database_name = database_file_name
            generate_reports_button.disabled = False
            status_text_generate_reports.value = generate_report_status_content()
        elif result["return_code"] == ReturnCodes.OK_GEN_NEW_DATABASE.value:
            status_text_fte_upload.value = f"Congratulation!!\nDatabase file {os.path.join(database_file_directory, database_file_name)} was created."
            database_file_saved = True
            saved_database_file_directory = database_file_directory
            saved_database_name = database_file_name
            generate_reports_button.disabled = False
            status_text_generate_reports.value = generate_report_status_content()
        elif result["return_code"] == ReturnCodes.ERROR_FILE.value:
            status_text_fte_upload.value = "Oops!!\nInput file has error. Please check Headers and Sheets"
        elif result["return_code"] == ReturnCodes.ERROR_FILE_DATA.value:
            status_text_fte_upload.value = "Oops!!\nInput file has duplicated staff ID or Error in Category Order"
        elif result["return_code"] == ReturnCodes.ERROR_FILE_LOADING.value:
            status_text_fte_upload.value = "Oops!!\nInput file cannot be loaded"
        elif result["return_code"] == ReturnCodes.ERROR_PROGRAM.value:
            status_text_fte_upload.value = "Oops!!\nPossible program error occurred"
        elif result["return_code"] == ReturnCodes.ERROR.value:
            status_text_fte_upload.value = "Oops!!\nSome error occurred"
        else:
            status_text_fte_upload.value = "Oops!!\nUnknown error occurred"

        if "issue_staff_numbers_not_in_base" in result:
            status_text_fte_upload.value += (
                "\n"
                + f"Staff Numbers not found in Base Data: {', '.join(result_dict['issue_staff_numbers_not_in_base'])}"
            )

        if "issue_expand_staff_fte_not_1" in result:
            status_text_fte_upload.value += (
                "\n"
                + f"Staff Numbers with FTE not 100% in Expand Data: {', '.join(result_dict['issue_expand_staff_fte_not_1'])}"
            )

        page.update()

    update_database_button = Button(
        content=Text(update_database_button_prompt),
        icon=Icons.FORWARD,
        on_click=update_database,
        disabled=True,
        style=disabled_button_style,
    )

    def reset_upload_fte(e):
        """Reset FTE data upload process to initial state."""
        init_data_upload_setup()
        set_button_text(fte_data_upload_button, datafile_upload_button_prompt)
        fte_data_month_text.value = fte_data_date.strftime("%Y / %m")
        status_text_fte_upload.value = init_fte_upload_status_content
        update_database_button.disabled = True
        optional_report_upload_button.disabled = True
        page.update()

    restart_button_fte_upload = Button(
        content=Text(restart_button_fte_upload_prompt),
        icon=Icons.RESET_TV,
        on_click=reset_upload_fte,
        style=disabled_button_style,
    )

    def generate_report_status_content():
        """Prepare the status content for reports generation page."""
        if database_file_saved:
            return f"Database file at {saved_database_file_directory}\nnamed {saved_database_name} is set."
        return "Please set the database file and reports start month"

    status_text_generate_reports = Text(
        generate_report_status_content(), bgcolor=Colors.YELLOW, color=Colors.BLACK
    )

    def init_generate_report_setup():
        """Initialize the reports generation setup parameters."""

        global report_start_date
        global saved_database_file_directory
        global saved_database_name
        global database_file_saved

        if database_file_saved is False:
            saved_database_name = None
            saved_database_file_directory = None

        status_text_generate_reports.value = generate_report_status_content()
        report_start_date = datetime.now() - relativedelta(months=1)

    init_generate_report_setup()

    database_file_upload_button_prompt = "Set Database File for Reports Generation"

    generate_report_start_month_text = Text(
        report_start_date.strftime("%Y / %m"), color=Colors.BLUE
    )

    database_file_upload_button = Button(
        content=Text(database_file_upload_button_prompt),
        icon=Icons.ADD_BOX,
        on_click=lambda _: page.run_task(pick_report_files_click),
        style=disabled_button_style,
    )

    def generate_reports(e):
        """Generate reports from saved database file."""

        database_file_name = os.path.join(
            saved_database_file_directory, saved_database_name
        )
        timestamp = (
            str(report_start_date.year)
            + str(report_start_date.month).zfill(2)
            + "_"
            + str(report_start_date.hour).zfill(2)
            + str(report_start_date.minute).zfill(2)
        )

        adj_department_fte_summary_report_file_name = os.path.join(
            saved_database_file_directory,
            department_fte_summary_report_file_name + "_" + timestamp,
        )
        report_header = f"{company_name}{HEADER_SEPARATOR}{department_fte_summary_report_title}{HEADER_SEPARATOR}{financial_year_header}"

        if (
            generate_department_fte_summary_report(
                database_file_name,
                adj_department_fte_summary_report_file_name,
                report_header,
                report_start_date.year,
                report_start_date.month,
            )
            == ReturnCodes.OK.value
        ):
            pdf = adj_department_fte_summary_report_file_name + ".pdf"
            excel = adj_department_fte_summary_report_file_name + ".xlsx"
            if os.path.exists(pdf) and os.path.exists(excel):
                status_text_generate_reports.value = f"Congratulation!!\nReport {adj_department_fte_summary_report_file_name} (pdf / xlsx) was generated."
            else:
                status_text_generate_reports.value = f"Oops\nGenerating report named {adj_department_fte_summary_report_file_name} (pdf / xlsx) was not successful."
        else:
            status_text_generate_reports.value = f"Oops\nDatabase file has problems. Report named {adj_department_fte_summary_report_file_name} (pdf / xlsx) not generated"

        adj_department_headcount_summary_report_file_name = os.path.join(
            saved_database_file_directory,
            department_headcount_summary_report_file_name + "_" + timestamp,
        )
        report_header = f"{company_name}{HEADER_SEPARATOR}{department_headcount_summary_report_title}{HEADER_SEPARATOR}{financial_year_header}"

        if (
            generate_department_headcount_summary_report(
                database_file_name,
                adj_department_headcount_summary_report_file_name,
                report_header,
                report_start_date.year,
                report_start_date.month,
            )
            == ReturnCodes.OK.value
        ):
            pdf = adj_department_headcount_summary_report_file_name + ".pdf"
            excel = adj_department_headcount_summary_report_file_name + ".xlsx"
            if os.path.exists(pdf) and os.path.exists(excel):
                status_text_generate_reports.value += "\n" + f"Congratulation!!\nReport {adj_department_headcount_summary_report_file_name} (pdf / xlsx) was generated."
            else:
                status_text_generate_reports.value += "\n" + f"Oops\nGenerating report named {adj_department_headcount_summary_report_file_name} (pdf / xlsx) was not successful."
        else:
            status_text_generate_reports.value += "\n" + f"Oops\nDatabase file has problems. Report named {adj_department_headcount_summary_report_file_name} (pdf / xlsx) not generated"

        adj_department_fte_costcentre_report_file_name = os.path.join(
            saved_database_file_directory,
            department_fte_costcentre_report_file_name + "_" + timestamp,
        )
        report_header = f"{company_name}{HEADER_SEPARATOR}{department_fte_costcentre_report_title}{HEADER_SEPARATOR}{financial_year_header}"

        if (
            generate_department_fte_costcentre_report(
                database_file_name,
                adj_department_fte_costcentre_report_file_name,
                report_header,
                report_start_date.year,
                report_start_date.month,
            )
            == ReturnCodes.OK.value
        ):
            pdf = adj_department_fte_costcentre_report_file_name + ".pdf"
            excel = adj_department_fte_costcentre_report_file_name + ".xlsx"
            if os.path.exists(pdf) and os.path.exists(excel):
                status_text_generate_reports.value += "\n" + f"Congratulation!!\nReport {adj_department_fte_costcentre_report_file_name} (pdf / xlsx) was generated."
            else:
                status_text_generate_reports.value += "\n" + f"Oops\nGenerating report named {adj_department_fte_costcentre_report_file_name} (pdf / xlsx) was not successful."
        else:
            status_text_generate_reports.value += "\n" + f"Oops\nDatabase file has problems. Report named {adj_department_fte_costcentre_report_file_name} (pdf / xlsx) not generated"

        page.update()

    generate_reports_button = Button(
        content=Text("Generate Reports"),
        icon=Icons.FORWARD,
        on_click=generate_reports,
        disabled=not database_file_saved,
        style=disabled_button_style,
    )

    def reset_generate_reports(e):
        """Reset report generation process to initial state."""

        global database_file_saved
        global saved_database_file_directory
        global saved_database_name

        database_file_saved = False
        saved_database_file_directory = None
        saved_database_name = None

        init_generate_report_setup()
        generate_report_start_month_text.value = report_start_date.strftime("%Y / %m")
        status_text_generate_reports.value = generate_report_status_content()
        generate_reports_button.disabled = True
        page.update()

    restart_button_generate_reports = Button(
        content=Text("Restart Reports Generation"),
        icon=Icons.RESET_TV,
        on_click=reset_generate_reports,
        style=disabled_button_style,
    )

    async def pick_data_files_click():
        """Handle the selected data file from file picker dialog."""

        global data_name
        global data_directory
        global database_file_directory
        global database_file_name

        files = await pick_data_files_dialog.pick_files(
            initial_directory=data_directory,
            allow_multiple=False,
        )
        if not files:
            status_text_fte_upload.value = (
                "Wrong file or no file. Please Select FTE Monthly Data File"
            )
        else:
            result = files[0]
            name = result.name
            path = result.path
            directory = os.path.dirname(path)

            data_name = name
            data_directory = directory
            database_file_directory = directory

            datafile_upload_button_prompt = f"Data Uploaded: {data_name}"
            set_button_text(fte_data_upload_button, datafile_upload_button_prompt)
            status_text_fte_upload.value = f"Data file at {data_directory}\nnamed {data_name} loaded.\n\nDatabase file at {database_file_directory}\nnamed {database_file_name} will be generated/updated."
            update_database_button.disabled = False
            optional_report_upload_button.disabled = False

        page.update()

    async def pick_optional_report_files_click():
        """Handle the selected optional database file from file picker dialog."""

        global database_file_directory
        global database_file_name

        files = await pick_optional_report_files_dialog.pick_files(
            initial_directory=data_directory,
            allow_multiple=False,
        )
        if not files:
            status_text_fte_upload.value = (
                "Wrong Upload or no file. Please Select Database File"
            )
        else:
            result = files[0]
            name = result.name
            path = result.path
            directory = os.path.dirname(path)

            database_file_name = name
            database_file_directory = directory

            status_text_fte_upload.value = f"Data file at {data_directory}\nnamed {data_name} loaded.\n\nDatabase file at {database_file_directory}\nnamed {database_file_name} will be generated/updated."
        page.update()

    async def pick_report_files_click():
        """Handle the selected database file from file picker dialog."""

        global database_file_saved
        global saved_database_file_directory
        global saved_database_name

        files = await pick_report_files_dialog.pick_files(
            initial_directory=data_directory,
            allow_multiple=False,
        )
        if not files:
            database_file_saved = False
            saved_database_name = None
            saved_database_file_directory = None
            generate_reports_button.disabled = True
            status_text_generate_reports.value = generate_report_status_content()
        else:
            result = files[0]
            name = result.name
            path = result.path
            directory = os.path.dirname(path)

            database_file_saved = True
            saved_database_name = name
            saved_database_file_directory = directory
            generate_reports_button.disabled = False
            status_text_generate_reports.value = generate_report_status_content()
        page.update()

    pick_data_files_dialog = FilePicker()
    pick_optional_report_files_dialog = FilePicker()
    pick_report_files_dialog = FilePicker()

    page.services.append(pick_data_files_dialog)
    page.services.append(pick_optional_report_files_dialog)
    page.services.append(pick_report_files_dialog)

    def minus_fte_data_date_click(e):
        """Decrease FTE data date by one month when minus button clicked."""
        global fte_data_date
        fte_data_date = fte_data_date - relativedelta(months=1)
        fte_data_month_text.value = fte_data_date.strftime("%Y / %m")
        set_button_text(
            fte_data_upload_button,
            f"Please Select {fte_data_date.strftime('%Y / %m')} FTE Data File",
        )
        page.update()

    def plus_fte_data_date_click(e):
        """Increase FTE data date by one month when plus button clicked."""
        global fte_data_date
        fte_data_date = fte_data_date + relativedelta(months=1)
        fte_data_month_text.value = fte_data_date.strftime("%Y / %m")
        set_button_text(
            fte_data_upload_button,
            f"Please Select {fte_data_date.strftime('%Y / %m')} FTE Data File",
        )
        page.update()

    def minus_report_start_month_click(e):
        """Decrease report start month by one month when minus button clicked."""
        global report_start_date
        report_start_date = report_start_date - relativedelta(months=1)
        generate_report_start_month_text.value = report_start_date.strftime("%Y / %m")
        page.update()

    def plus_report_start_month_click(e):
        """Increase report start month by one month when plus button clicked."""
        global report_start_date
        report_start_date = report_start_date + relativedelta(months=1)
        generate_report_start_month_text.value = report_start_date.strftime("%Y / %m")
        page.update()

    pages = [
        (
            NavigationRailDestination(
                icon=Icons.CLOUD_UPLOAD_OUTLINED,
                selected_icon=Icons.CLOUD_UPLOAD,
                label="Upload Monthy FTE Data",
            ),
            Row(
                controls=[
                    Column(
                        horizontal_alignment=CrossAxisAlignment.STRETCH,
                        controls=[
                            Card(
                                content=Container(
                                    Text("Upload FTE monthly data", weight=FontWeight.BOLD),
                                    padding=20,
                                    bgcolor=Colors.BLUE,
                                )
                            ),
                            status_text_fte_upload,
                            Row(
                                [
                                    IconButton(Icons.REMOVE, on_click=minus_fte_data_date_click),
                                    fte_data_month_text,
                                    IconButton(Icons.ADD, on_click=plus_fte_data_date_click),
                                ],
                                alignment=MainAxisAlignment.CENTER,
                            ),
                            fte_data_upload_button,
                            optional_report_upload_button,
                            update_database_button,
                            restart_button_fte_upload,
                        ],
                        expand=True,
                    ),
                ]
            ),
        ),
        (
            NavigationRailDestination(
                icon=Icons.DATA_EXPLORATION_OUTLINED,
                selected_icon=Icons.DATA_EXPLORATION,
                label="Generate Reports",
            ),
            Row(
                controls=[
                    Column(
                        horizontal_alignment=CrossAxisAlignment.STRETCH,
                        controls=[
                            Card(
                                content=Container(
                                    Text("Generate FTE Reports", weight=FontWeight.BOLD),
                                    padding=20,
                                    bgcolor=Colors.BLUE,
                                )
                            ),
                            status_text_generate_reports,
                            Row(
                                [
                                    IconButton(Icons.REMOVE, on_click=minus_report_start_month_click),
                                    generate_report_start_month_text,
                                    IconButton(Icons.ADD, on_click=plus_report_start_month_click),
                                ],
                                alignment=MainAxisAlignment.CENTER,
                            ),
                            database_file_upload_button,
                            generate_reports_button,
                            restart_button_generate_reports,
                        ],
                        expand=True,
                    ),
                ]
            ),
        ),
    ]

    menu_layout = DesktopAppLayout(page=page, pages=pages, title=f"HR Cost Reporting ({APP_VERSION})")

    page.bgcolor = Colors.WHITE
    page.add(menu_layout)


if __name__ == "__main__":
    flet.run(main)
