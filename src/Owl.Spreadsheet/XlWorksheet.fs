namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

type XlWorksheet (sheet: IXLWorksheet) =
  member internal __.raw with get() = sheet
  member __.Item 
    with get(row: int, column: int) = sheet.Cell(row, column) |> XlCell
    and set(row: int, column: int) (value: obj) = sheet.Cell(row, column).Value <- value
  member __.Item 
    with get(start': string, end': string) = sheet.Range($"{start'}:{end'}") |> XlRange
    and set(start': string, end': string) (value: obj) = sheet.Range($"{start'}:{end'}").Value <- value
  member __.Item
    with get(address: string) = sheet.Cells(address) |> XlCells
    and set(address: string) (value: obj) = sheet.Cells(address).Value <- value

module XlWorksheet =
  let public get_cell_at (sheet: XlWorksheet) (address: string) =
    sheet.raw.Cell address |> XlCell
  let public get_cell (sheet: XlWorksheet) (row:int, colmun:int) =
    sheet.raw.Cell(row, colmun) |> XlCell
    
  let public get_cells_at (sheet: XlWorksheet) (range: string)  =
    sheet.raw.Cells range |> XlCells
  let public get_cells (sheet: XlWorksheet) (from': int * int) (to': int * int) =
    get_cells_at sheet $"%s{from'.to_address()}:%s{to'.to_address()}"

  let public get_range_at (sheet: XlWorksheet) (range: string) =
    sheet.raw.Range range |> XlRange
  let public get_range (sheet: XlWorksheet) (from': int * int) (to': int * int) =
    get_range_at sheet $"%s{from'.to_address()}:%s{to'.to_address()}"
  let public get_range_by (sheet: XlWorksheet) (from': string) (to': string) =
    get_range_at sheet $"{from'}:{to'}"
    
  let public get_column (sheet: XlWorksheet) (column: int) =
    sheet.raw.Column(column) |> XlColumn
  let public get_column_at (sheet: XlWorksheet) (column: string) =
    sheet.raw.Column(column) |> XlColumn
    
  let public get_columns (sheet: XlWorksheet) (from': int, to': int) =
    sheet.raw.Columns(from', to') |> XlColumns
  let public get_columns_at (sheet: XlWorksheet) (columns: string) =
    sheet.raw.Columns(columns) |> XlColumns
  let public get_columns_by (sheet: XlWorksheet) (from': string, to': string) =
    sheet.raw.Columns(from', to') |> XlColumns
    
  let public get_row (sheet: XlWorksheet) (row: int) =
    sheet.raw.Row(row) |> XlRow
  
  let public get_all_rows (sheet: XlWorksheet) =
    sheet.raw.Rows() |> XlRows
  let public get_rows (sheet: XlWorksheet) (first': int, last': int) =
    sheet.raw.Rows(first', last') |> XlRows
  let public get_rows_at (sheet: XlWorksheet) (rows: string) =
    sheet.raw.Rows(rows) |> XlRows

