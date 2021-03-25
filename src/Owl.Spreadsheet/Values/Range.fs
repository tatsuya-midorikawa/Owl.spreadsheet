namespace Owl.Spreadsheet

open System
open ClosedXML.Excel
open System.Collections
open System.Collections.Generic

type Range internal (range: IXLRange) =
  member __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.set(value) = __.value <- box value
  member __.set_formula(value: string) = range.FormulaA1  <- value
  member __.set_formula_r1c1(value: string) = range.FormulaR1C1  <- value
  member __.cell(row: int, column: int) = Cell(range.Cell(row, column))
  member __.cell(address: string) = Cell(range.Cell address)
  member __.cells(from': Address, to': Address) = 
    let address = $"%s{from'.to_string()}:%s{to'.to_string()}"
    Cells(range.Cells address)
  member __.cells(from': int * int, to':  int * int) = 
    let address = $"%s{from'.to_address()}:%s{to'.to_address()}"
    Cells(range.Cells address)
  member __.cells(address: string) = Cells(range.Cells address)
