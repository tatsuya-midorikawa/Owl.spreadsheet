namespace Owl.Spreadsheet

open System
open ClosedXML.Excel
open System.Collections
open System.Collections.Generic

type Range internal (range: IXLRange) =
  member __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.first_cell with get() = Cell(range.FirstCell())
  member __.first_cell_used with get() = Cell(range.FirstCellUsed())
  member __.last_cell with get() = Cell(range.LastCell())
  member __.last_cell_used with get() = Cell(range.LastCellUsed())

  member __.fx(value: obj) = range.FormulaA1 <- value.ToString()
  member __.set(value) = __.value <- box value
  member __.set_formula(value: string) = range.FormulaA1  <- value
  member __.set_formula_r1c1(value: string) = range.FormulaR1C1  <- value
  member __.cell(row: int, column: int) = Cell(range.Cell(row, column))
  member __.cell(address: string) = Cell(range.Cell address)
  member __.cells(from': Address, to': Address) = Cells(range.Cells $"%s{from'.to_string()}:%s{to'.to_string()}")
  member __.cells(from': int * int, to':  int * int) = Cells(range.Cells $"%s{from'.to_address()}:%s{to'.to_address()}")
  member __.cells(address: string) = Cells(range.Cells address)
  // TODO
  member __.insert_column_after(number_of_columns: int) = range.InsertColumnsAfter(number_of_columns)
  // TODO
  member __.insert_column_before(number_of_columns: int) = range.InsertColumnsBefore(number_of_columns)
  // TODO
  member __.insert_row_above(number_of_rows: int) = range.InsertRowsAbove(number_of_rows)
  // TODO
  member __.insert_row_below(number_of_rows: int) = range.InsertRowsBelow(number_of_rows)
