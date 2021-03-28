namespace Owl.Spreadsheet

open System
open ClosedXML.Excel
open System.Collections
open System.Collections.Generic

type Column internal (column: IXLColumn) =
  member __.raw with get() = column
  member __.value with set(value) = column.Value <- value
  member __.cell_count with get() = column.CellCount
  member __.worksheet with get() = column.Worksheet
  member __.width with get() = column.Width
  member __.first_cell with get() = Cell(column.FirstCell())
  member __.first_cell_used with get() = Cell(column.FirstCellUsed())
  member __.last_cell with get() = Cell(column.LastCell())
  member __.last_cell_used with get() = Cell(column.LastCellUsed())

  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = column.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = column.FormulaA1  <- value
  member __.set_formula_r1c1(value: string) = column.FormulaR1C1  <- value
  member __.cell(row': int) = Cell(column.Cell row')
  member __.cells(row': int) = Cells(column.Cells(row'.ToString()))
  member __.cells(from': int, to': int) = Cells(column.Cells $"%d{from'}:%d{to'}")
  member __.cells(row_range: string) = Cells(column.Cells row_range)
  member __.left() = Column(column.ColumnLeft())
  member __.left(step: int) = Column(column.ColumnLeft(step))
  member __.right() = Column(column.ColumnRight())
  member __.right(step: int) = Column(column.ColumnRight(step))

