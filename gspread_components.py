from xai_components.base import InArg, OutArg, InCompArg, Component, xai_component

@xai_component
class GspreadAuth(Component):
    """A component to authenticate the user with Gspread and generate a client object.

    ##### inPorts:
    - json_path: The path to the JSON key file for OAuth2 authentication.

    ##### outPorts:
    - gc: A Gspread client object.
    """

    json_path: InArg[str]
    gc: OutArg[any]

    def execute(self, ctx) -> None:
        import gspread

        json_path = self.json_path.value
        if json_path:
            self.gc.value = gspread.service_account(filename=json_path)
        else:
            # will try to fetch from default gspread locations
            self.gc.value = gspread.service_account()

        ctx.update({'gc': self.gc.value})

        
@xai_component
class OpenSpreadsheet(Component):
    """A component to open a Google Spreadsheet by its title.
    If a worksheet title is not provided, it will open the first worksheet by default.

    ##### inPorts:
    - title: The title of the Google Spreadsheet.
    - gc: A Gspread client object.
    - worksheet_title: The title of the worksheet.

    ##### outPorts:
    - sh: The opened Spreadsheet object.
    - worksheet: The selected worksheet object.
    """

    title: InCompArg[str]
    gc: InArg[any]
    worksheet_title: InArg[str]
    sh: OutArg[any]
    worksheet: OutArg[any]

    def execute(self, ctx) -> None:
        gc = self.gc.value if self.gc.value else ctx["gc"]

        # Open the spreadsheet
        sh = gc.open(self.title.value)
        ctx.update({'sh': sh})
        if self.worksheet_title.value:
            self.worksheet.value = sh.worksheet(self.worksheet_title.value)
        else:
            self.worksheet.value = sh.sheet1
        
        ctx.update({'worksheet': self.worksheet.value})



@xai_component
class OpenSpreadsheetFromUrl(Component):
    """A component to open a Google Spreadsheet by its URL.
    If a worksheet title is not provided, it will open the first worksheet by default.

    ##### inPorts:
    - url: The URL of the Google Spreadsheet.
    - gc: A Gspread client object.
    - worksheet_title: The title of the worksheet.

    ##### outPorts:
    - sh: The opened Spreadsheet object.
    - worksheet: The selected worksheet object.
    """

    url: InCompArg[str]
    gc: InArg[any]
    worksheet_title: InArg[str]
    sh: OutArg[any]
    worksheet: OutArg[any]

    def execute(self, ctx) -> None:
        gc = self.gc.value if self.gc.value else ctx["gc"]

        # Open the spreadsheet
        sh = gc.open_by_url(self.url.value)
        ctx.update({'sh': sh})

        if self.worksheet_title.value:
            self.worksheet.value = sh.worksheet(self.worksheet_title.value)
        else:
            self.worksheet.value = sh.sheet1
        
        ctx.update({'worksheet': self.worksheet.value})


@xai_component
class OpenWorksheet(Component):
    """A component to open a Google worksheet.
    If a title is not provided, it will return the first worksheet.

    ##### inPorts:
    - sh: The Spreadsheet object.
    - worksheet_title: The title of the worksheet.

    ##### outPorts:
    - worksheet: The selected worksheet object.
    """

    sh: InArg[any]
    worksheet_title: InArg[str]

    worksheet: OutArg[any]

    def execute(self, ctx) -> None:
        sh = self.sh.value if self.sh.value else ctx["sh"]

        if self.worksheet_title.value:
            self.worksheet.value = sh.worksheet(self.worksheet_title.value)
        else:
            self.worksheet.value = sh.sheet1
        
        ctx.update({'worksheet': self.worksheet.value})

@xai_component
class ReadCell(Component):
    """A component to read a cell from a worksheet.

    ##### inPorts:
    - worksheet: The worksheet object.
    - cell_address: The cell address in the 'A1' format.

    ##### outPorts:
    - cell: The cell content.
    """

    worksheet: InArg[any]
    cell_address: InCompArg[str]
    cell: OutArg[list]

    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']
        cell_address = self.cell_address.value

        # Select a range
        self.cell.value = worksheet.get(cell_address)
        print(self.cell.value)
        

@xai_component
class UpdateCell(Component):
    """A component to update a cell in the spreadsheet.

    ##### inPorts:
    - worksheet: The worksheet object.
    - cell_address: The cell address in the 'A1' format.
    - value: The new value for the cell.
    """

    worksheet: InArg[any]
    cell_address: InCompArg[str]
    value: InCompArg[str]

    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']

        # Update a range
        worksheet.update(self.cell_address.value, [[self.value.value]])


@xai_component
class CreateSpreadsheet(Component):
    """A component to create a new spreadsheet.

    ##### inPorts:
    - gc: A Gspread client object.
    - spreadsheet_title: The title for the new spreadsheet.
    - share_email: The email to share the new spreadsheet with.

    ##### outPorts:
    - sh: The created Spreadsheet object.
    - worksheet: The first worksheet of the new spreadsheet.
    """

    gc: InArg[any]
    spreadsheet_title: InCompArg[str]
    share_email: InArg[str]
    sh: OutArg[any]
    worksheet: OutArg[any]

    def execute(self, ctx) -> None:
        gc = self.gc.value if self.gc.value else ctx['gc']

        self.sh.value = gc.create(self.spreadsheet_title.value)

        if self.share_email.value:
            self.sh.value.share('self.share_email.value', perm_type='user', role='writer')

        ctx.update({'sh': self.sh.value})
        self.worksheet.value = self.sh.value.sheet1
        ctx.update({'worksheet': self.worksheet.value})

@xai_component
class CreateWorksheet(Component):
    """A component to create a new worksheet in the spreadsheet.

    ##### inPorts:
    - sh: The Spreadsheet object.
    - worksheet_title: The title for the new worksheet.
    - rows: The number of rows for the new worksheet. Default 1000.
    - cols: The number of columns for the new worksheet. Default 26.

    ##### outPorts:
    - worksheet: The created worksheet object.
    """

    sh: InArg[any]
    worksheet_title: InCompArg[str]
    rows: InArg[int]
    cols: InArg[int]
    worksheet: OutArg[any]

    def __init__(self):
        super().__init__()
        self.rows.value = 1000
        self.cols.value = 26

    def execute(self, ctx) -> None:
        sh = self.sh.value if self.sh.value else ctx['sh']
        worksheet_title = self.worksheet_title.value

        self.worksheet.value = sh.add_worksheet(title=worksheet_title, rows=self.rows.value, cols=self.cols.value)
        ctx.update({'worksheet': self.worksheet.value})


@xai_component
class DeleteWorksheet(Component):
    """A component to delete a worksheet in the spreadsheet.

    ##### inPorts:
    - sh: The Spreadsheet object.
    - worksheet_title: The title of the worksheet to be deleted.
    """
    sh: InArg[any]
    worksheet_title: InCompArg[str]

    def execute(self, ctx) -> None:
        sh = self.sh.value if self.sh.value else ctx['sh']
        worksheet_title = self.worksheet_title.value

        # Get the worksheet to delete
        worksheet = sh.worksheet(worksheet_title)

        # Delete the worksheet
        sh.del_worksheet(worksheet)

@xai_component
class AppendRow(Component):
    """A component that adds a row of values at the bottom of the worksheet.

    ##### inPorts:
    - worksheet: the worksheet object.
    - values: a list of values to be added as a new row.
    """

    worksheet: InArg[any]
    values: InCompArg[list]

    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']
        values = self.values.value

        # Append a new row with values
        worksheet.append_row(values)
        
@xai_component
class ReadRow(Component):
    """A component that retrieves all values from a specific row in the worksheet.

    ##### inPorts:
    - worksheet: The worksheet object.
    - row_index: The index of the row from which the values need to be retrieved.

    ##### outPorts:
    - row_values: The values obtained from the specified row.
    """

    worksheet: InArg[any]
    row_index: InCompArg[int]
    row_values: OutArg[any]

    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']

        # Read all values from the row
        self.row_values.value = worksheet.row_values(self.row_index.value)
        print(self.row_values.value)

@xai_component
class ReadColumn(Component):
    """A component that reads all values from a specific column in the worksheet.

    ##### inPorts:
    - worksheet: The worksheet object.
    - col_index: The index of the column from which the values need to be read.

    ##### outPorts:
    - col_values: The values obtained from the specified column.
    """

    worksheet: InArg[any]
    col_index: InCompArg[int]
    col_values: OutArg[any]

    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']

        # Read all values from the column
        self.col_values.value = worksheet.col_values(self.col_index.value)
        print(self.col_values.value)


@xai_component
class InsertRow(Component):
    """A component that inserts a row at a specific index in the worksheet.

    ##### inPorts:
    - values: The values to be inserted as a new row.
    - index: The index at which the new row needs to be inserted.
    """

    values: InCompArg[list]
    index: InCompArg[int]

    def execute(self, ctx) -> None:
        worksheet = ctx['worksheet']
        values = self.values.value
        index = self.index.value

        # Insert a new row with values at a specific index
        worksheet.insert_row(values, index)

@xai_component
class UpdateRange(Component):
    """A component that updates a range of cells in the worksheet with new values.

    ##### inPorts:
    - worksheet: The worksheet object.
    - cell_range: The range of cells to be updated, specified in the 'A1:B2' format.
    - values: The new values for the cells in the range.
    """

    worksheet: InArg[any]
    cell_range: InCompArg[str]
    values: InCompArg[list]

    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']
        cell_range = self.cell_range.value
        values = self.values.value

        # Update a range of cells with new values
        worksheet.update(cell_range, values)


@xai_component
class GetAllValues(Component):
    """A component that retrieves all values from a worksheet as a list of lists.

    ##### inPorts:
    - worksheet: The worksheet object.

    ##### outPorts:
    - list_of_values: A list of all values retrieved from the worksheet.
    """
    worksheet: InArg[any]
    list_of_values: OutArg[list]
    
    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']

        # Get all records from the worksheet
        self.list_of_values.value = worksheet.get_values()
        print(self.list_of_values.value)


@xai_component
class GetAllRecords(Component):
    """A component that retrieves all records from the worksheet as a list of dictionaries.

    ##### inPorts:
    - worksheet: The worksheet object.

    ##### outPorts:
    - records: A list of all records retrieved from the worksheet. Each record is represented as a dictionary.
    """

    worksheet: InArg[any]
    records: OutArg[any]
    
    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']

        # Get all records from the worksheet
        self.records.value = worksheet.get_all_records()
        print(self.records.value)

@xai_component
class ClearWorksheet(Component):
    """A component that removes all values from the worksheet.

    ##### inPorts:
    - sh: The worksheet object.
    - worksheet: The worksheet object.
    """

    sh: InArg[any]
    worksheet: InArg[any]
    
    def execute(self, ctx) -> None:

        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']
        
        # Clear all values from the worksheet
        worksheet.clear()


@xai_component
class FindAllStringMatches(Component):
    """A component that finds all cells in the worksheet that contain a specific string.

    ##### inPorts:
    - worksheet: The worksheet object.
    - value: The string value to search for in the cells.

    ##### outPorts:
    - cell_list: A list of all cells that contain the specified string value.
    """

    worksheet: InArg[any]
    value: InCompArg[str]

    cell_list: OutArg[list]
    
    def execute(self, ctx) -> None:

        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']
        value = self.value.value

        # Find all cells with string value
        self.cell_list.value = worksheet.findall(value)
        print(self.cell_list.value)


@xai_component
class FindAllRegexMatches(Component):
    """A component that finds all cells in the worksheet that match a specified regular expression.

    ##### inPorts:
    - worksheet: The worksheet object.
    - regex: The regular expression to search for in the cells.

    ##### outPorts:
    - cell_list: A list of all cells that match the specified regular expression.
    """
    worksheet: InArg[any]
    regex: InCompArg[str]

    cell_list: OutArg[list]
    
    def execute(self, ctx) -> None:

        import re

        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']
        regex = self.regex.value

        # Compile the regular expression
        criteria_re = re.compile(regex)

        # Find all cells with regex
        self.cell_list.value = worksheet.findall(criteria_re)
        print(self.cell_list.value)
