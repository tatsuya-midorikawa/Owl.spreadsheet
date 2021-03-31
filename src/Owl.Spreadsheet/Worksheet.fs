namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

type Worksheet (sheet: IXLWorksheet) =
  member __.Item with get(row: int, column: int) = sheet.Cell(row, column) |> XlCell
  member __.Item with get(start': string, end': string) = sheet.Range($"{start'}:{end'}") |> XlRange
  member __.Item with get(address: string) = sheet.Cells(address) |> XlCells

module Worksheet =
  let public get_cell_at (sheet: IXLWorksheet) (address: string) =
    XlCell(sheet.Cell address)
  let public get_cell (sheet: IXLWorksheet) (row:int, colmun:int) =
    XlCell(sheet.Cell(row, colmun))
    
  let public get_cells_at (sheet: IXLWorksheet) (range: string)  =
    XlCells(sheet.Cells range)
  let public get_cells (sheet: IXLWorksheet) (from': int * int) (to': int * int) =
    get_cells_at sheet $"%s{from'.to_address()}:%s{to'.to_address()}"

  let public get_range_at (sheet: IXLWorksheet) (range: string) =
    XlRange(sheet.Range range)
  let public get_range (sheet: IXLWorksheet) (from': int * int) (to': int * int) =
    get_range_at sheet $"%s{from'.to_address()}:%s{to'.to_address()}"
  let public get_range_by (sheet: IXLWorksheet) (from': string) (to': string) =
    get_range_at sheet $"{from'}:{to'}"
    
  let public get_column (sheet: IXLWorksheet) (column: int) =
    XlColumn(sheet.Column(column))
  let public get_column_at (sheet: IXLWorksheet) (column: string) =
    XlColumn(sheet.Column(column))
    
  // TODO
  let public get_columns (sheet: IXLWorksheet) (from': int, to': int) =
    sheet.Columns(from', to')
  // TODO
  let public get_columns_at (sheet: IXLWorksheet) (columns: string) =
    sheet.Columns(columns)
  // TODO
  let public get_columns_by (sheet: IXLWorksheet) (from': string, to': string) =
    sheet.Columns(from', to')
    
  let public get_row (sheet: IXLWorksheet) (row: int) =
    XlRow(sheet.Row(row))
  
  // TODO
  let public get_all_rows (sheet: IXLWorksheet) =
    sheet.Rows()
  // TODO
  let public get_rows (sheet: IXLWorksheet) (first': int, last': int) =
    sheet.Rows(first', last')
  // TODO
  let public get_rows_at (sheet: IXLWorksheet) (rows: string) =
    sheet.Rows(rows)

