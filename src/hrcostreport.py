from dataprocess import process_update_database, ReturnCodes, generate_department_fte_summary_report, generate_department_headcount_summary_report, generate_department_fte_costcentre_report
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os

import flet
from flet import (
    AppBar,
    Column,
    Row,
    Container,
    IconButton,
    Icons,
    NavigationRail,
    NavigationRailDestination,
    Page,
    Text,
    Card,
    Colors,
    Divider,
    PopupMenuButton,
    PopupMenuItem,
    ElevatedButton,
    IconButton,
    TextField,
    TextAlign,
    VerticalDivider,
    Padding,
    FilePicker,
    FilePickerResultEvent,
    Theme,
    ElevatedButtonTheme,
    MainAxisAlignment,
)
#from flet import colors, icons

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
department_fte_summary_report_file_name = "HR_department_fte_summary_report"
department_fte_summary_report_title = "Yearly Department FTE Trend"

department_headcount_summary_report_file_name = "HR_department_headcount_summary_report"
department_headcount_summary_report_title = "Yearly Department Headcount Trend"

department_fte_costcentre_report_file_name = "HR_department_fte_costcentres_report"
department_fte_costcentre_report_title = "Yearly Department FTE (Cost Centres) Trend"



def init_data_upload_setup():
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

    def __init__(
        self,
        title,
        page,
        pages,
        *args,
        window_size=(800, 600),
        **kwargs,
    ):
        super().__init__(*args, **kwargs)


        self.page = page
        self.pages = pages

        self.expand = True

        self.navigation_items = [navigation_item for navigation_item, _ in pages]
        self.navigation_rail = self.build_navigation_rail()
        
        self.update_destinations()
        self._menu_extended = True
        self.navigation_rail.extended = True

        self.menu_panel = Row(
            controls=[
                self.navigation_rail,
            ],
            spacing=0,
            tight=True,
        )

        page_contents = [page_content for _, page_content in pages]
        self.content_area = Column(page_contents, expand=True)

        self._was_portrait = self.is_portrait()
        self._panel_visible = self.is_landscape()

        self.set_content()

        self._change_displayed_page()

        self.page.on_resize = self.handle_resize

        #self.page.appbar = self.create_appbar()

        self.window_size = window_size
        self.page.window_width, self.page.window_height = self.window_size
        
        self.page.title = title

    def select_page(self, page_number):
        self.navigation_rail.selected_index = page_number
        self._change_displayed_page()

    def _navigation_change(self, e):
        self._change_displayed_page()
        self.page.update()

    def _change_displayed_page(self):
        page_number = self.navigation_rail.selected_index
        for i, content_page in enumerate(self.content_area.controls):
            # update selected page
            content_page.visible = page_number == i

    def build_navigation_rail(self):
        return NavigationRail(
            selected_index=0,
            label_type="none",
            on_change=self._navigation_change,
            bgcolor=Colors.WHITE,
            indicator_color=Colors.BLUE,
            selected_label_text_style=flet.TextStyle(color=Colors.BLUE),
            unselected_label_text_style=flet.TextStyle(color=Colors.BLACK),
            group_alignment = 0.0,
        )
        
    
    def update_destinations(self):
        self.navigation_rail.destinations = self.navigation_items
        self.navigation_rail.label_type = "all"

    def handle_resize(self, e):
        pass

    def set_content(self):
        self.controls = [self.menu_panel, VerticalDivider(width=1, color=Colors.RED)  , self.content_area]
        self.update_destinations()
        self.navigation_rail.extended = self._menu_extended
        self.menu_panel.visible = self._panel_visible

    def is_portrait(self) -> bool:
        # Return true if window/display is narrow
        # return self.page.window_height >= self.page.window_width
        return self.page.height >= self.page.width

    def is_landscape(self) -> bool:
        # Return true if window/display is wide
        return self.page.width > self.page.height

def create_page(title: str, body: str, page: Page):
        
    return Row(
        controls=[
            Column(
                horizontal_alignment="stretch",
                controls=[
                    Card(content=Container(Text(title, weight="bold"), padding=8, bgcolor=Colors.RED)),
                    Text(body,bgcolor=Colors.BLUE,),
                ],
                expand=True,
            ),
        ],
        expand=True,
    )


def main(page: Page):
    
    page.window.width = 900        # window's width is 200 px
    page.window.height = 450       # window's height is 200 px
    page.update()
    
    page.theme = Theme(
        elevated_button_theme=ElevatedButtonTheme(
            #bgcolor=Colors.ERROR,
            #foreground_color=Colors.ERROR_CONTAINER,
            #fixed_size=Size(200, 50),
            disabled_bgcolor=Colors.GREY_300,
            disabled_foreground_color=Colors.GREEN,
        ),
        
    )

### setup page 1 - FTE data upload
  
    init_fte_upload_status_content = "Data not uploaded"
    status_text_fte_upload = Text(init_fte_upload_status_content,bgcolor=Colors.YELLOW,color=Colors.BLACK)    
    
    datafile_upload_button_prompt = f"Please Select {fte_data_date.strftime('%Y / %m')}  FTE Data File"
    optional_report_upload_button_prompt = "Optionally Set Database File"
    update_database_button_prompt = "Update Database File"
    restart_button_fte_upload_prompt = "Restart Upload Process"

    fte_data_month_text = Text(fte_data_date.strftime("%Y / %m"),color=Colors.BLUE)

    fte_data_upload_button = ElevatedButton(
                        datafile_upload_button_prompt,
                        icon=Icons.UPLOAD_FILE,
                        on_click=lambda _: pick_data_files_dialog.pick_files(
                            allow_multiple=False
                        ),                        
                    )
    
    optional_report_upload_button = ElevatedButton(
                        optional_report_upload_button_prompt,
                        icon=Icons.ADD_BOX,
                        on_click=lambda _: pick_optional_report_files_dialog.pick_files(
                            allow_multiple=False
                        ),
                        disabled=True,
                    )
                    
    def update_database(e):
        global database_file_saved
        global saved_database_file_directory
        global saved_database_name

        data_period = f"{str(fte_data_date.year)}{str(fte_data_date.month).zfill(2)}"
        datafile = data_directory + data_name
        report_file = database_file_directory + database_file_name
        result = process_update_database(datafile,data_period,report_file)
        if result == ReturnCodes.OK_UPDATE_DATABASE:
            status_text_fte_upload.value = f"Congratulation!!\nDatabase file {database_file_directory}{database_file_name} was updated."
            database_file_saved = True
            saved_database_file_directory = database_file_directory
            saved_database_name = database_file_name
            generate_reports_button.disabled = False
            status_text_generate_reports.value = generate_report_status_content()
        elif result == ReturnCodes.OK_GEN_NEW_DATABASE:
            status_text_fte_upload.value = f"Congratulation!!\nDatabase file {database_file_directory}]{database_file_name} was created."
            database_file_saved = True
            saved_database_file_directory = database_file_directory
            saved_database_name = database_file_name
            generate_reports_button.disabled = False
            status_text_generate_reports.value = generate_report_status_content()
        elif result == ReturnCodes.ERROR_FILE_ERROR:
            status_text_fte_upload.value = f"Oops!!\nInput file has error. Please check Headers and Sheets"
        elif result == ReturnCodes.ERROR_FILE_DATA_ERROR:
            status_text_fte_upload.value = f"Oops!!\nInput file has duplicated staff ID or Error in Category Order"
        elif result == ReturnCodes.ERROR_FILE_LOADING:
            status_text_fte_upload.value = f"Oops!!\nInput file cannot be loaded"
        elif result == ReturnCodes.ERROR_PROGRAM:
            status_text_fte_upload.value = f"Oops!!\nPossible program error occurred"
        elif result == ReturnCodes.ERROR:
            status_text_fte_upload.value = f"Oops!!\nSome error occurred"
        else:
            status_text_fte_upload.value = f"Oops!!\nUnknown error occurred"        
            
        page.update()
    
    update_database_button = ElevatedButton(
                        update_database_button_prompt,
                        icon=Icons.FORWARD,
                        on_click=update_database,
                        disabled=True,
                    )

    def reset_upload_fte(e):
        init_data_upload_setup()
        fte_data_upload_button.text = datafile_upload_button_prompt        
        fte_data_month_text.value = fte_data_date.strftime("%Y / %m")
        status_text_fte_upload.value = init_fte_upload_status_content
        update_database_button.disabled = True
        optional_report_upload_button.disabled = True
        page.update()
        
    restart_button_fte_upload = ElevatedButton(
                    restart_button_fte_upload_prompt,
                    icon=Icons.RESET_TV,
                    on_click=reset_upload_fte,
                    )

### setup page 2 - reports generation

    def generate_report_status_content():
        #global database_file_saved
        if database_file_saved:
            return f"Database file at {saved_database_file_directory}\nnamed {saved_database_name} is set."
        else:
            return "Please set the database file and reports start month"    


    status_text_generate_reports = Text(generate_report_status_content(),bgcolor=Colors.YELLOW,color=Colors.BLACK)


    def init_generate_report_setup():

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

    generate_report_start_month_text = Text(report_start_date.strftime("%Y / %m"),color=Colors.BLUE)
    
    database_file_upload_button = ElevatedButton(
                        database_file_upload_button_prompt,
                        icon=Icons.ADD_BOX,
                        on_click=lambda _: pick_report_files_dialog.pick_files(
                            allow_multiple=False
                        ),
                    )

# def generate_summary_report(fte_data_file_name: str, summary_report_file_name: str, report_title: str, start_year: int, start_month: int, number_of_month: i):

    def generate_reports(e):
        database_file_name = saved_database_file_directory + saved_database_name
        timestamp = str(report_start_date.year) + str(report_start_date.month).zfill(2) + "_" + str(report_start_date.hour).zfill(2) + str(report_start_date.minute).zfill(2)
        
        adj_department_fte_summary_report_file_name = department_fte_summary_report_file_name + "_" + timestamp + ".pdf"
        adj_department_fte_summary_report_file_name = saved_database_file_directory + adj_department_fte_summary_report_file_name        
        
        if generate_department_fte_summary_report(database_file_name,
                                adj_department_fte_summary_report_file_name,
                                department_fte_summary_report_title,
                                report_start_date.year,
                                report_start_date.month) == ReturnCodes.OK:

            if os.path.exists(adj_department_fte_summary_report_file_name):
                status_text_generate_reports.value = f"Congratulation!!\nReport {adj_department_fte_summary_report_file_name} was generated."
            else:
                status_text_generate_reports.value = f"Oops\nGenerating report named {adj_department_fte_summary_report_file_name} was not successful."
        else:
            status_text_generate_reports.value = f"Oops\nDatabase file has problem. Report named {adj_department_fte_summary_report_file_name} not generated"

        adj_department_headcount_summary_report_file_name = department_headcount_summary_report_file_name + "_" + timestamp + ".pdf"
        adj_department_headcount_summary_report_file_name = saved_database_file_directory + adj_department_headcount_summary_report_file_name
        
        if generate_department_headcount_summary_report(database_file_name,
                                adj_department_headcount_summary_report_file_name,
                                department_headcount_summary_report_title,
                                report_start_date.year,
                                report_start_date.month) == ReturnCodes.OK:
            if os.path.exists(adj_department_headcount_summary_report_file_name):
                status_text_generate_reports.value = status_text_generate_reports.value + "\n" + f"Congratulation!!\nReport {adj_department_headcount_summary_report_file_name} was generated."
            else:
                status_text_generate_reports.value = status_text_generate_reports.value + "\n" + f"Oops\nGenerating report named {adj_department_headcount_summary_report_file_name} was not successful."
        else:
            status_text_generate_reports.value = status_text_generate_reports.value + "\n" + f"Oops\nDatabase file has problem. Report named {adj_department_headcount_summary_report_file_name} not generated"

            
        adj_department_fte_costcentre_report_file_name = department_fte_costcentre_report_file_name + "_" + timestamp + ".pdf"
        adj_department_fte_costcentre_report_file_name = saved_database_file_directory + adj_department_fte_costcentre_report_file_name
        
        if generate_department_fte_costcentre_report(database_file_name,
                                adj_department_fte_costcentre_report_file_name,
                                department_fte_costcentre_report_title,
                                report_start_date.year,
                                report_start_date.month) == ReturnCodes.OK:
            if os.path.exists(adj_department_fte_costcentre_report_file_name):
                status_text_generate_reports.value = status_text_generate_reports.value + "\n" + f"Congratulation!!\nReport {adj_department_fte_costcentre_report_file_name} was generated."
            else:
                status_text_generate_reports.value = status_text_generate_reports.value + "\n" + f"Oops\nGenerating report named {adj_department_fte_costcentre_report_file_name} was not successful."
        else:
            status_text_generate_reports.value = status_text_generate_reports.value + "\n" + f"Oops\nDatabase file has problem. Report named {adj_department_fte_costcentre_report_file_name} not generated"

        page.update()
    
    generate_reports_button = ElevatedButton(
                        "Generate Reports",
                        icon=Icons.FORWARD,
                        on_click=generate_reports,
                        disabled=database_file_saved == False, 
                    )    

    def reset_generate_reports(e):
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
    
    restart_button_generate_reports = ElevatedButton(
                    "Restart Reports Generation",
                    icon=Icons.RESET_TV,
                    on_click=reset_generate_reports,
                    )

        
    def pick_data_files_result(e: FilePickerResultEvent):
        
        global data_name
        global data_directory
        global database_file_directory
        global database_file_name
        global datafile_upload_button_prompt
        
        selected_file_result = e
        if selected_file_result.files == None or len(selected_file_result.files) == 0:
            status_text_fte_upload.value = f"Wrong file or no file. Please Select FTE Monthly Data File" 
        else:
            result = selected_file_result.files.pop()
            name = result.name
            path = result.path
            directory = path.replace(name, "")
            
            data_name = name
            data_directory = directory
            database_file_directory = directory
            
            datafile_upload_button_prompt = f"Data Uploaded: {data_name}"
            status_text_fte_upload.value = f"Data file at {data_directory}\nnamed {data_name} loaded.\n\nDatabase file at {database_file_directory}\nnamed {database_file_name} will be generated/updated."
            update_database_button.disabled = False
            optional_report_upload_button.disabled = False
            
        page.update()
                
    def optional_report_files_result(e: FilePickerResultEvent):
        
        selected_file_result = e
        global data_name
        global data_directory
        global database_file_directory
        global database_file_name
        if selected_file_result.files == None or len(selected_file_result.files) == 0:
            status_text_fte_upload.value = f"Wrong Upload or no file. Please Select Database File" 
        else:
            result = selected_file_result.files.pop()
            name = result.name
            path = result.path
            directory = path.replace(name, "")
            
            database_file_name = name
            database_file_directory = directory
            
            status_text_fte_upload.value = f"Data file at {data_directory}\nnamed {data_name} loaded.\n\nDatabase file at {database_file_directory}\nnamed {database_file_name} will be generated/updated."
        page.update() 

    def report_files_result(e: FilePickerResultEvent):
            
            global database_file_saved
            global saved_database_file_directory
            global saved_database_name
            
            selected_file_result = e

            if selected_file_result.files == None or len(selected_file_result.files) == 0:

                database_file_saved = False
                saved_database_name = None
                saved_database_file_directory = None                
                generate_reports_button.disabled = True
                status_text_generate_reports.value = generate_report_status_content()
                #status_text_generate_reports.value = f"Wrong Upload or no file. Please Select Database File" 
            else:
                result = selected_file_result.files.pop()
                name = result.name
                path = result.path
                directory = path.replace(name, "")
                
                database_file_saved = True
                saved_database_name = name
                saved_database_file_directory = directory
                generate_reports_button.disabled = False
                status_text_generate_reports.value = generate_report_status_content()
            page.update() 



    pick_data_files_dialog = FilePicker(on_result=pick_data_files_result)
    pick_optional_report_files_dialog = FilePicker(on_result=optional_report_files_result)
    pick_report_files_dialog = FilePicker(on_result=report_files_result)

    page.overlay.append(pick_data_files_dialog)
    page.overlay.append(pick_optional_report_files_dialog)
    page.overlay.append(pick_report_files_dialog)

    
    def minus_fte_data_date_click(e):
        global fte_data_date
        fte_data_date = fte_data_date - relativedelta(months=1)
        fte_data_month_text.value = fte_data_date.strftime("%Y / %m")
        fte_data_upload_button.text = f"Please Select {fte_data_date.strftime('%Y / %m')} FTE Data File"
        page.update()
        
    def plus_fte_data_date_click(e):
        global fte_data_date
        fte_data_date = fte_data_date + relativedelta(months=1)
        fte_data_month_text.value = fte_data_date.strftime("%Y / %m")
        fte_data_upload_button.text = f"Please Select {fte_data_date.strftime('%Y / %m')} FTE Data File"
        page.update()
    
    def minus_report_start_month_click(e):
        global report_start_date
        report_start_date = report_start_date - relativedelta(months=1)
        generate_report_start_month_text.value = report_start_date.strftime("%Y / %m")
        page.update()
        
    def plus_report_start_month_click(e):
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
                        horizontal_alignment="stretch",
                        controls=[
                            Card(content=Container(Text("Upload FTE monthly data", weight="bold"), padding=20, bgcolor=Colors.BLUE)),
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
            )
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
                        horizontal_alignment="stretch",
                        controls=[
                            Card(content=Container(Text("Generate FTE Reports", weight="bold"), padding=20, bgcolor=Colors.BLUE)),
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
            )
        ),
    ]

    menu_layout = DesktopAppLayout(
        page=page,
        pages=pages,
        title="HR Cost Reporting",
        #window_size=(320, 120),
    )

    page.bgcolor = Colors.WHITE
    page.add(menu_layout)


if __name__ == "__main__":
    flet.app(
        target=main,
    )
