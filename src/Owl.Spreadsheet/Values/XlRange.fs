namespace Owl.Spreadsheet

open System
open ClosedXML.Excel
open System.Collections
open System.Collections.Generic

type XlRange internal (range: IXLRange) as self =
  member internal __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.first_cell with get() = XlCell(range.FirstCell())
  member __.first_cell_used with get() = XlCell(range.FirstCellUsed())
  member __.last_cell with get() = XlCell(range.LastCell())
  member __.last_cell_used with get() = XlCell(range.LastCellUsed())

  member __.fx(value: obj) = range.FormulaA1 <- value.ToString(); self
  member __.set(value) = __.value <- box value; self
  member __.set_formula(value: string) = range.FormulaA1 <- value; self
  member __.set_formula_r1c1(value: string) = range.FormulaR1C1 <- value; self
  member __.cell(row: int, column: int) = XlCell(range.Cell(row, column))
  member __.cell(address: string) = XlCell(range.Cell address)
  member __.cells(from': Address, to': Address) = XlCells(range.Cells $"%s{from'.to_string()}:%s{to'.to_string()}")
  member __.cells(from': int * int, to':  int * int) = XlCells(range.Cells $"%s{from'.to_address()}:%s{to'.to_address()}")
  member __.cells(address: string) = XlCells(range.Cells address)
  // TODO
  member __.insert_column_after(number_of_columns: int) = range.InsertColumnsAfter(number_of_columns)
  // TODO
  member __.insert_column_before(number_of_columns: int) = range.InsertColumnsBefore(number_of_columns)
  // TODO
  member __.insert_row_above(number_of_rows: int) = range.InsertRowsAbove(number_of_rows)
  // TODO
  member __.insert_row_below(number_of_rows: int) = range.InsertRowsBelow(number_of_rows)
