namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

module Worksheet =
  let public get_cell (sheet: IXLWorksheet) ((row, colmun): int * int) =
    Cell(sheet.Cell(row, colmun))

  let public get_cell_at (sheet: IXLWorksheet) (address: Address) =
    Cell(sheet.Cell(address.row, address.column))

  let public get_cells (sheet: IXLWorksheet) (from': int * int) (to':  int * int)  =
    let address = $"%s{from'.to_address()}:%s{to'.to_address()}"
    Cells(sheet.Cells address)
    
  let public get_cells_at (sheet: IXLWorksheet) (from': Address) (to': Address)  =
    let address = $"%s{from'.to_tuple().to_address()}:%s{to'.to_tuple().to_address()}"
    Cells(sheet.Cells address)

  let public get_range (sheet: IXLWorksheet) (from': Address) (to': Address) =
    let address = $"%s{from'.to_tuple().to_address()}:%s{to'.to_tuple().to_address()}"
    Owl.Spreadsheet.Range(sheet.Range address)
