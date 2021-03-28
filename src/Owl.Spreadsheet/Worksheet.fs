namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

module Worksheet =
  let public get_cell (sheet: IXLWorksheet) ((row, colmun): int * int) =
    XlCell(sheet.Cell(row, colmun))

  let public get_cell_at (sheet: IXLWorksheet) (address: string) =
    XlCell(sheet.Cell address)

  let public get_cells (sheet: IXLWorksheet) (from': int * int) (to': int * int) =
    let address = $"%s{from'.to_address()}:%s{to'.to_address()}"
    XlCells(sheet.Cells address)
    
  let public get_cells_at (sheet: IXLWorksheet) (range: string)  =
    XlCells(sheet.Cells range)

  let public get_range (sheet: IXLWorksheet) (from': int * int) (to': int * int)=
    let address = $"%s{from'.to_address()}:%s{to'.to_address()}"
    XlRange(sheet.Range address)

  let public get_range_at (sheet: IXLWorksheet) (range: string)=
    XlRange(sheet.Range range)
    
  let public get_column (sheet: IXLWorksheet) (column: int) =
    sheet.Column(column)
    
  let public get_column_at (sheet: IXLWorksheet) (column: string) =
    sheet.Column(column)

  let public get_columns (sheet: IXLWorksheet) (from': int) (to': int) =
    sheet.Columns(from', to')

  let public get_columns_at (sheet: IXLWorksheet) (columns: string)=
    sheet.Columns(columns)
