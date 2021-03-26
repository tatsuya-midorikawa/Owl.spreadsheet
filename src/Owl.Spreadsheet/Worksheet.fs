namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

module Worksheet =
  let public get_cell (sheet: IXLWorksheet) ((row, colmun): int * int) =
    Cell(sheet.Cell(row, colmun))

  let public get_cell_at (sheet: IXLWorksheet) (address: string) =
    Cell(sheet.Cell address)

  let public get_cells (sheet: IXLWorksheet) (from': int * int) (to': int * int) =
    let address = $"%s{from'.to_address()}:%s{to'.to_address()}"
    Cells(sheet.Cells address)
    
  let public get_cells_at (sheet: IXLWorksheet) (range: string)  =
    Cells(sheet.Cells range)

  let public get_range (sheet: IXLWorksheet) (from': int * int) (to': int * int)=
    let address = $"%s{from'.to_address()}:%s{to'.to_address()}"
    Owl.Spreadsheet.Range(sheet.Range address)

  let public get_range_at (sheet: IXLWorksheet) (range: string)=
    Owl.Spreadsheet.Range(sheet.Range range)

