namespace Owl.Spreadsheet

open System
open ClosedXML.Excel
open System.Collections
open System.Collections.Generic

type XlColumn internal (column: IXLColumn) =
  member internal __.raw with get() = column
  member __.value with set(value) = column.Value <- value
  member __.cell_count with get() = column.CellCount
  member __.worksheet with get() = column.Worksheet
  member __.width with get() = column.Width
  member __.first_cell with get() = XlCell(column.FirstCell())
  member __.first_cell_used with get() = XlCell(column.FirstCellUsed())
  member __.last_cell with get() = XlCell(column.LastCell())
  member __.last_cell_used with get() = XlCell(column.LastCellUsed())

  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = column.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = column.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = column.FormulaR1C1 <- value
  member __.cell(row': int) = XlCell(column.Cell row')
  member __.cells() = XlCells(column.Cells())
  member __.cells(row_range: string) = XlCells(column.Cells row_range)
  member __.cells(row': int) = __.cells(row'.ToString())
  member __.cells(from': int, to': int) = __.cells($"%d{from'}:%d{to'}")
  member __.left() = XlColumn(column.ColumnLeft())
  member __.left(step: int) = XlColumn(column.ColumnLeft(step))
  member __.right() = XlColumn(column.ColumnRight())
  member __.right(step: int) = XlColumn(column.ColumnRight(step))

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = column.Cells() |> Seq.map(fun cell -> XlCell(cell))
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = column.Cells() |> Seq.map(fun cell -> XlCell(cell))
      cells.GetEnumerator()
    
