namespace Owl.Spreadsheet

open ClosedXML.Excel

module Tools =
  [<Literal>]
  let TRUE = true
  [<Literal>]
  let FALSE = false

  let load (filepath: string) = new XLWorkbook(filepath) |> XlWorkbook
  let create name = XlWorkbook.create name
  let create_from template = XlWorkbook.create_from template
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
  let del_by (name: string) (worksheet: XlWorksheet) = 
    let target = if System.String.IsNullOrEmpty name then worksheet.raw.Name else name
    worksheet.workbook.worksheets.delete(target)

  let save (sheet: XlWorksheet) = sheet.save()
  let save_as (name: string) (sheet: XlWorksheet) = sheet.save_as name
  let save_and_close (sheet: XlWorksheet) = sheet.save_and_close()
  let close (sheet: XlWorksheet) = sheet.close()

  let cell (row: int, column: int) (sheet: XlWorksheet) = sheet.cell(row, column)
  let cells (range: string) (sheet: XlWorksheet) = sheet.cells range
  let range (range: string) (sheet: XlWorksheet) = sheet.range range
  
  let row (row: int) (sheet: XlWorksheet) = sheet.row row
  let rows (first': int, last': int) (sheet: XlWorksheet) = sheet.rows(first', last')
  let rows_for (rows: string) (sheet: XlWorksheet) = sheet.rows rows
  let all_rows (sheet: XlWorksheet) = sheet.rows()

  let column (column: int) (sheet: XlWorksheet) = sheet.column column
  let column_at (column: string) (sheet: XlWorksheet) = sheet.column column
  let columns (from': int, to': int) (sheet: XlWorksheet) = sheet.columns(from', to')
  let columns_for (columns: string) (sheet: XlWorksheet) = sheet.columns columns
  let all_columns (sheet: XlWorksheet) = sheet.columns()

  let set value (cell: XlCell) = cell.set value
  let set_to_cells value (cells: XlCells) = cells.set value
  let set_r value (range: XlRange) = range.set value

  let fx formula (cell: XlCell) = cell.fx formula
  let fx_s formula (cells: XlCells) = cells.fx formula
  let fx_r formula (range: XlRange) = range.fx formula
