import smartsheet
from datetime import datetime

def find_column_by_name(columns, name):
    for column in columns:
        if (column.title == name):
            return column.id

def apply_formula_to_column_sequential(sheet, column_id, formula):
    rows_to_update = []
    for row in sheet.rows:
        new_row = smartsheet.models.Row()
        new_row.id = row.id
        rows_to_update.append(new_row)
        for cell in row.cells:
            if(cell.column_id == column_id):
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = column_id
                new_cell.formula = formula.format(row.row_number)
                new_row.cells.append(new_cell)
    return rows_to_update

def apply_formula_to_column_all_but(sheet, column_id, formula):
    rows_to_update = []
    for row in sheet.rows:
        new_row = smartsheet.models.Row()
        new_row.id = row.id
        rows_to_update.append(new_row)
        for cell in row.cells:
            if(cell.column_id == column_id):
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = column_id
                new_cell.formula = formula.format(row.row_number)
                new_row.cells.append(new_cell)
    return rows_to_update


# Initialize client
smartsheet_client = smartsheet.Smartsheet()

# Make sure we don't miss any errors
smartsheet_client.errors_as_exceptions(True)

SHEET_RAMI_ROADMAP = 6659608434501508

sheet_id = SHEET_RAMI_ROADMAP

sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

column_id_to_update = 0
rows_to_update = []

columns = smartsheet_client.Sheets.get_columns(
  sheet_id,       # sheet_id
  include_all=True).data

ORIGINAL_TIME = "Original Time Estimated"
PROGRESS = "Progress"
ISSUE_TYPE = "Issue type"

original_time_formula = "=Duration{0}*1440"
progress_formula = "=IF(Completion{0} < 0.25, \"Empty\", IF(Completion{0} < 0.5, \"Quarter\", IF(Completion{0} < 0.75, \"Half\", IF(Completion{0} < 1, \"Three Quarter\", IF(Completion{0} >= 1, \"Full\", \"Full\")))))"
remaining_formula="=SUM(CHILDREN())"

# Find column to update
original_time_column_id = find_column_by_name(columns, ORIGINAL_TIME)
progress_column_id = find_column_by_name(columns, PROGRESS)
issue_type_column_id = find_column_by_name(columns, ISSUE_TYPE)

rows = apply_formula_to_column_sequential(sheet, progress_column_id, progress_formula)
# rows = apply_formula_to_column(sheet, original_time_column_id, original_time_formula)

updated_row = smartsheet_client.Sheets.update_rows(
    sheet_id,      # sheet_id
    rows)