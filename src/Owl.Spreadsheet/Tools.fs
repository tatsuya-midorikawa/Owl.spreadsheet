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
  let workbook (sheet: XlWorksheet) = sheet.raw.Workbook

  let fst (workbook: XLWorkbook) = workbook.Worksheet(1) |> XlWorksheet
  let snd (workbook: XLWorkbook) = workbook.Worksheet(2) |> XlWorksheet
  let thd (workbook: XLWorkbook) = workbook.Worksheet(3) |> XlWorksheet
  let last (workbook: XLWorkbook) = workbook.Worksheet(workbook.Worksheets.Count) |> XlWorksheet
  let at (n: int) (workbook: XLWorkbook) = workbook.Worksheet(n) |> XlWorksheet
  let by (name: string) (workbook: XLWorkbook) = workbook.Worksheet(name) |> XlWorksheet

  let add (name: string) (workbook: XLWorkbook) = workbook.Worksheets.Add(name) |> XlWorksheet
  let del (name: string) (workbook: XLWorkbook) = workbook.Worksheets.Delete(name)
  let del_at (n: int) (workbook: XLWorkbook) = workbook.Worksheets.Delete(n)
  let del_by (n: int) (worksheet: XlWorksheet) = worksheet.raw.Workbook.Worksheets.Delete(worksheet.raw.Name)

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
