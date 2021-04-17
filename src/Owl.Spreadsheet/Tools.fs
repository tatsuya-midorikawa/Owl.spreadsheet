namespace Owl.Spreadsheet

open ClosedXML.Excel

module Tools =
  [<Literal>]
  let TRUE = true
  [<Literal>]
  let FALSE = false

  let load filepath = XlWorkbook.open_workbook filepath
  let create name = XlWorkbook.new_workbook_with name
  let create_from template = XlWorkbook.new_workbook_from template
  let workbook (sheet: XlWorksheet) = sheet.workbook

  let fst (workbook: XlWorkbook) = workbook.worksheet(1)
  let snd (workbook: XlWorkbook) = workbook.worksheet(2)
  let thd (workbook: XlWorkbook) = workbook.worksheet(3)
  let last (workbook: XlWorkbook) = workbook.worksheet(workbook.worksheets.count)
  let at (n: int) (workbook: XlWorkbook) = workbook.worksheet(n)
  let by (name: string) (workbook: XlWorkbook) = workbook.worksheet(name)

  let add (name: string) (workbook: XlWorkbook) = workbook.worksheets.add(name)
  let del (name: string) (workbook: XlWorkbook) = workbook.worksheets.delete(name)
  let del_at (n: int) (workbook: XlWorkbook) = workbook.worksheets.delete(n)
  let del_by (n: int) (worksheet: XlWorksheet) = worksheet.workbook.worksheets.delete(worksheet.raw.Name)

  let save (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.save |> ignore
  let save_as (name: string) (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.save_as name |> ignore
  let save_and_close (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.save_and_close |> ignore
  let close (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.close |> ignore

  let cell (row: int, column: int) (sheet: XlWorksheet) = XlWorksheet.get_cell sheet (row, column)
  let cells (range: string) (sheet: XlWorksheet) = XlWorksheet.get_cells_at sheet range
  let range (range: string) (sheet: XlWorksheet) = XlWorksheet.get_range_at sheet range
  
  let row (row: int) (sheet: XlWorksheet) = XlWorksheet.get_row sheet row
  let rows (first': int, last': int) (sheet: XlWorksheet) = XlWorksheet.get_rows sheet (first', last')
  let rows_for (rows: string) (sheet: XlWorksheet) = XlWorksheet.get_rows_at sheet rows
  let all_rows (sheet: XlWorksheet) = XlWorksheet.get_all_rows sheet

  let column (column: int) (sheet: XlWorksheet) = XlWorksheet.get_column sheet column
  let column_at (column: string) (sheet: XlWorksheet) = XlWorksheet.get_column_at sheet column
  let columns (from': int, to': int) (sheet: XlWorksheet) = XlWorksheet.get_columns sheet (from', to')
  let columns_for (columns: string) (sheet: XlWorksheet) = XlWorksheet.get_columns_at sheet columns
  let all_columns (sheet: XlWorksheet) = XlWorksheet.get_all_columns sheet
