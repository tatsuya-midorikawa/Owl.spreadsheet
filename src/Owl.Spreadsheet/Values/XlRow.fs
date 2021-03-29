namespace Owl.Spreadsheet

open System
open ClosedXML.Excel
open System.Collections
open System.Collections.Generic

type XlRow internal (row: IXLRow) =
  member internal __.raw with get() = row
  member __.value with set(value) = row.Value <- value
  member __.cell_count with get() = row.CellCount
  member __.worksheet with get() = row.Worksheet
  member __.height with get() = row.Height
  member __.first_cell with get() = XlCell(row.FirstCell())
  member __.first_cell_used with get() = XlCell(row.FirstCellUsed())
  member __.last_cell with get() = XlCell(row.LastCell())
  member __.last_cell_used with get() = XlCell(row.LastCellUsed())

  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = row.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = row.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = row.FormulaR1C1 <- value
  member __.cell(row': int) = XlCell(row.Cell row')
  member __.cells() = XlCells(row.Cells())
  member __.cells(row_range: string) = XlCells(row.Cells row_range)
  member __.cells(row': int) = __.cells(row'.ToString())
  member __.cells(from': int, to': int) = __.cells($"%d{from'}:%d{to'}")
  
  member __.clear(?options: ClearOption) = match options with Some opt -> row.Clear(opt) | None -> row.Clear()

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = row.Cells() |> Seq.map(fun cell -> XlCell(cell))
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = row.Cells() |> Seq.map(fun cell -> XlCell(cell))
      cells.GetEnumerator()
    
