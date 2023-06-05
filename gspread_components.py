from xai_components.base import InArg, OutArg, InCompArg, Component, xai_component

@xai_component(color="green")
class GspreadAuth(Component):
    """A component that authenticates the user with Gspread and creates a client object.

    ##### inPorts:
    - json_path: a path to the JSON key file for OAuth2 authentication.
    """
    json_path: InCompArg[str]

    def execute(self, ctx) -> None:
        import gspread

        json_path = self.json_path.value
        gc = gspread.service_account(filename=json_path)
        ctx.update({'gc': gc})

        
@xai_component(color="blue")
class OpenSpreadsheet(Component):
    """A component that opens a Google Spreadsheet by its title.

    ##### inPorts:
    - title: the title of the Google Spreadsheet.
    """
    title: InCompArg[str]

    def execute(self, ctx) -> None:
        gc = ctx['gc']
        title = self.title.value

        # Open the spreadsheet
        sh = gc.open(title)
        ctx.update({'sh': sh})


@xai_component(color="orange")
class ReadCell(Component):
    """A component that reads a cell from the spreadsheet.

    ##### inPorts:
    - cell_address: the cell address in the format 'A1'.
    """
    cell_address: InCompArg[str]

    def execute(self, ctx) -> None:
        sh = ctx['sh']
        cell_address = self.cell_address.value

        # Select a range
        print(sh.sheet1.get(cell_address))


@xai_component(color="purple")
class UpdateCell(Component):
    """A component that updates a cell in the spreadsheet.

    ##### inPorts:
    - cell_address: the cell address in the format 'A1'.
    - value: the new value for the cell.
    """
    cell_address: InCompArg[str]
    value: InCompArg[str]

    def execute(self, ctx) -> None:
        sh = ctx['sh']
        cell_address = self.cell_address.value
        value = self.value.value

        # Update a range
        sh.sheet1.update(cell_address, [[value]])


@xai_component(color="brown")
class CreateWorksheet(Component):
    """A component that creates a new worksheet in the spreadsheet.

    ##### inPorts:
    - worksheet_title: the title of the new worksheet.
    - rows: the number of rows in the new worksheet.
    - cols: the number of columns in the new worksheet.
    """
    worksheet_title: InCompArg[str]
    rows: InCompArg[int]
    cols: InCompArg[int]

    def execute(self, ctx) -> None:
        sh = ctx['sh']
        worksheet_title = self.worksheet_title.value
        rows = self.rows.value
        cols = self.cols.value

        # Add a new worksheet with 10 rows and 10 columns
        worksheet = sh.add_worksheet(title=worksheet_title, rows=rows, cols=cols)
        ctx.update({'worksheet': worksheet})

@xai_component(color="pink")
class DeleteWorksheet(Component):
    """A component that deletes a worksheet in the spreadsheet.

    ##### inPorts:
    - worksheet_title: the title of the worksheet to be deleted.
    """
    worksheet_title: InCompArg[str]

    def execute(self, ctx) -> None:
        sh = ctx['sh']
        worksheet_title = self.worksheet_title.value

        # Get the worksheet to delete
        worksheet = sh.worksheet(worksheet_title)

        # Delete the worksheet
        sh.del_worksheet(worksheet)

@xai_component(color="yellow")
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

@xai_component(color="grey")
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


@xai_component(color="teal")
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

@xai_component(color="lime")
class UpdateRange(Component):
    """A component that updates a range of cells in the worksheet.

    ##### inPorts:
    - cell_range: the range of cells to update in the 'A1:B2' format.
    - values: the new values for the cells in the range.
    """
    cell_range: InCompArg[str]
    values: InCompArg[list]

    def execute(self, ctx) -> None:
        worksheet = ctx['worksheet']
        cell_range = self.cell_range.value
        values = self.values.value

        # Update a range of cells with new values
        worksheet.update(cell_range, values)

@xai_component(color="violet")
class GetAllRecords(Component):
    """A component that gets all records from the worksheet."""
    
    def execute(self, ctx) -> None:
        worksheet = ctx['worksheet']

        # Get all records from the worksheet
        records = worksheet.get_all_records()
        print(records)

@xai_component(color="maroon")
class ClearWorksheet(Component):
    """A component that clears all values from the worksheet."""
    
    def execute(self, ctx) -> None:
        worksheet = ctx['worksheet']

        # Clear all values from the worksheet
        worksheet.clear()
