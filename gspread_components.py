from xai_components.base import InArg, OutArg, InCompArg, Component, xai_component

@xai_component
class GspreadAuth(Component):
    """A component that authenticates the user with Gspread and creates a client object.

    ##### inPorts:
    - json_path: a path to the JSON key file for OAuth2 authentication.
    """
    json_path: InArg[str]
    gc = OutArg[any]

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
    """A component that opens a Google Spreadsheet by its title. 
    By default will open the first worksheet if worksheet title is not provided.

    ##### inPorts:
    - title: the title of the Google Spreadsheet.
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
    """A component that opens a Google Spreadsheet by its url.
    By default will open the first worksheet if worksheet title is not provided.

    ##### inPorts:
    - url: the url of the Google Spreadsheet.
    """
    url: InCompArg[str]
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
class ReadCell(Component):
    """A component that reads a cell from a worksheet.

    ##### inPorts:
    - cell_address: the cell address in the format 'A1'.
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
    """A component that updates a cell in the spreadsheet.

    ##### inPorts:
    - cell_address: the cell address in the format 'A1'.
    - value: the new value for the cell.
    """
    worksheet: InArg[any]
    cell_address: InCompArg[str]
    value: InCompArg[str]

    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']

        # Update a range
        worksheet.update(self.cell_address.value, [[self.value.value]])


@xai_component
class CreateWorksheet(Component):
    """A component that creates a new worksheet in the spreadsheet.

    ##### inPorts:
    - worksheet_title: the title of the new worksheet.
    - rows: the number of rows in the new worksheet.
    - cols: the number of columns in the new worksheet.
    """
    worksheet_title: InCompArg[str]
    rows: InArg[int]
    cols: InArg[int]
    worksheet: OutArg[any]

    def __init__(self):
        super().__init__()
        self.rows.value = 10
        self.cols.value = 10

    def execute(self, ctx) -> None:
        sh = ctx['sh']
        worksheet_title = self.worksheet_title.value

        self.worksheet.value = sh.add_worksheet(title=worksheet_title, rows=self.rows.value, cols=self.cols.value)
        ctx.update({'worksheet': self.worksheet.value})


@xai_component
class DeleteWorksheet(Component):
    """A component that deletes a worksheet in the spreadsheet.

    ##### inPorts:
    - worksheet_title: the title of the worksheet to be deleted.
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
    """A component that appends a row to the worksheet.

    ##### inPorts:
    - values: the values to append as a new row.
    """
    values: InCompArg[list]

    def execute(self, ctx) -> None:
        worksheet = ctx['worksheet']
        values = self.values.value

        # Append a new row with values
        worksheet.append_row(values)
        
@xai_component
class ReadRow(Component):
    """A component that reads all values from a row in the worksheet.

    ##### inPorts:
    - row_values: the number of the row to read.
    """
    row_values: InCompArg[int]

    def execute(self, ctx) -> None:
        worksheet = ctx['worksheet']
        row_values = self.row_values.value

        # Read all values from the column
        print(worksheet.row_values(row_values))

@xai_component
class ReadColumn(Component):
    """A component that reads all values from a column in the worksheet.

    ##### inPorts:
    - col_number: the number of the column to read.
    """
    col_number: InCompArg[int]

    def execute(self, ctx) -> None:
        worksheet = ctx['worksheet']
        col_number = self.col_number.value

        # Read all values from the column
        print(worksheet.col_values(col_number))


@xai_component
class InsertRow(Component):
    """A component that inserts a row at a specific index in the worksheet.

    ##### inPorts:
    - values: the values to insert as a new row.
    - index: the index at which to insert the new row.
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
    """A component that updates a range of cells in the worksheet.

    ##### inPorts:
    - cell_range: the range of cells to update in the 'A1:B2' format.
    - values: the new values for the cells in the range.
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
class GetAllRecords(Component):
    """A component that gets all records from the worksheet."""
    worksheet: InArg[any]
    records: OutArg[any]
    
    def execute(self, ctx) -> None:
        worksheet = self.worksheet.value if self.worksheet.value else ctx['worksheet']

        # Get all records from the worksheet
        self.records.value = worksheet.get_all_self.records.value()
        print(self.records.value)

@xai_component
class ClearWorksheet(Component):
    """A component that clears all values from the worksheet."""
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
    - value: the string value to search for.
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
    """A component that finds all cells in the worksheet that match a regular expression.

    ##### inPorts:
    - regex: the regular expression to search for.
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
