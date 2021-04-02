namespace Owl.Spreadsheet

open System
open System.Collections
open System.Collections.Generic
open System.Data
open ClosedXML.Excel
open Owl.Spreadsheet.Convert

type XlCell internal (cell: IXLCell) =    
  member internal __.raw with get() = cell
  member __.value with get() = cell.Value and set(value) = cell.Value <- value
  member __.worksheet with get() = cell.Worksheet
  member __.get() = __.value
  member __.get<'T>() = 
    match typeof<'T> with
       | t when t = typeof<int16> -> cell.GetDouble() |> to_int16 |> unbox<'T>
       | t when t = typeof<int> -> cell.GetDouble() |> to_int |> unbox<'T>
       | t when t = typeof<int64> -> cell.GetDouble() |> to_int64 |> unbox<'T>
       | t when t = typeof<float32> -> cell.GetDouble() |> to_single |> unbox<'T>
       | t when t = typeof<float> -> cell.GetDouble() |> to_double |> unbox<'T>
       | t when t = typeof<decimal> -> cell.GetDouble() |> to_decimal |> unbox<'T>
       | t when t = typeof<string> -> cell.GetString() |> unbox<'T>
       | t when t = typeof<DateTime> -> cell.GetDateTime() |> unbox<'T>
       | t when t = typeof<TimeSpan> -> cell.GetTimeSpan() |> unbox<'T>
       | t when t = typeof<obj> -> __.value |> unbox<'T>
       | _ -> raise(exn "")

  member __.set<'T>(value: 'T) = __.value <- box value
  member __.fx(value: obj) = cell.FormulaA1 <- value.ToString()
  member __.get_formula() = cell.FormulaA1 
  member __.set_formula(value: string) = cell.FormulaA1 <- value
  member __.get_formula_r1c1() = cell.FormulaR1C1
  member __.set_formula_r1c1(value: string) = cell.FormulaR1C1 <- value
  member __.row_number with get() = cell.Address.RowNumber
  member __.column_number with get() = cell.Address.ColumnNumber
  member __.style with get() = cell.Style |> XlStyle

  member __.column with get() = cell.WorksheetColumn() |> XlColumn
  member __.row with get() = cell.WorksheetRow() |> XlRow

  member __.delete(option: ShiftDeleted) = cell.Delete(option)
  member __.clear(?options: ClearOption) = match options with Some opt -> cell.Clear(opt) | None -> cell.Clear()
  member __.left() = cell.CellLeft() |> XlCell
  member __.left(step: int) = cell.CellLeft(step) |> XlCell
  member __.right() = cell.CellRight() |> XlCell
  member __.right(step: int) =cell.CellRight(step) |> XlCell
  member __.above() = cell.CellAbove() |> XlCell
  member __.above(step: int) = cell.CellAbove(step) |> XlCell
  member __.below() = cell.CellBelow() |> XlCell
  member __.below(step: int) = cell.CellBelow(step) |> XlCell

  member __.copy_from(other_cell: IXLCell) = cell.CopyFrom(other_cell) |> XlCell
  member __.copy_from(other_cell: XlCell) = cell.CopyFrom(other_cell.raw) |> XlCell
  member __.copy_from(other_cell: string) = cell.CopyFrom(other_cell) |> XlCell
  member __.copy_to(target: IXLCell) = cell.CopyTo(target) |> XlCell
  member __.copy_to(target: XlCell) = cell.CopyTo(target.raw) |> XlCell
  member __.copy_to(target: string) = cell.CopyTo(target) |> XlCell

  member __.insert_cells_above(number_of_rows: int) = cell.InsertCellsAbove number_of_rows |> XlCells
  member __.insert_cells_after(number_of_columns: int) = cell.InsertCellsAfter number_of_columns |> XlCells
  member __.insert_cells_before(number_of_columns: int) = cell.InsertCellsBefore number_of_columns |> XlCells
  member __.insert_cells_below(number_of_rows: int) = cell.InsertCellsBelow number_of_rows |> XlCells

  member __.insert_table(data: DataTable) = cell.InsertTable(data) |> XlTable
  member __.insert_table(data: DataTable, create_table: bool) = cell.InsertTable(data, create_table) |> XlTable
  member __.insert_table(data: DataTable, table_name: string) = cell.InsertTable(data, table_name) |> XlTable
  member __.insert_table(data: DataTable, table_name: string, create_table: bool) = cell.InsertTable(data, table_name, create_table) |> XlTable



and XlCells internal (cells: IXLCells) =
  member internal __.raw with get() = cells
  member __.value with set(value) = cells.Value <- value
  member __.get() = cells |> Seq.map XlCell
  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = cells.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = cells.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = cells.FormulaR1C1 <- value
  member __.style with get() = cells.Style |> XlStyle
  
  member __.clear(?options: ClearOption) = match options with Some opt -> cells.Clear(opt) | None -> cells.Clear()
  member __.delete_comments() = cells.DeleteComments()
  member __.delete_sparklines() = cells.DeleteSparklines()

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = cells |> Seq.map XlCell
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = cells |> Seq.map XlCell
      cells.GetEnumerator()


    
and XlRow internal (row: IXLRow) =
  member internal __.raw with get() = row
  member __.value with set(value) = row.Value <- value
  member __.cell_count with get() = row.CellCount
  member __.worksheet with get() = row.Worksheet
  member __.height with get() = row.Height
  member __.first_cell with get() = row.FirstCell() |> XlCell
  member __.first_cell_used with get() = row.FirstCellUsed() |> XlCell
  member __.last_cell with get() = row.LastCell() |> XlCell
  member __.last_cell_used with get() = row.LastCellUsed() |> XlCell
  member __.row_number with get() = row.RowNumber()
  member __.style with get() = row.Style |> XlStyle

  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = row.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = row.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = row.FormulaR1C1 <- value
  member __.cell(row': int) = row.Cell row' |> XlCell
  member __.cells() = row.Cells() |> XlCells
  member __.cells(row_range: string) = row.Cells row_range |> XlCells
  member __.cells(row': int) = __.cells(row'.ToString())
  member __.cells(from': int, to': int) = __.cells($"%d{from'}:%d{to'}")

  member __.above() = row.RowAbove() |> XlRow
  member __.above(step: int) = row.RowAbove(step) |> XlRow
  member __.below() = row.RowBelow() |> XlRow
  member __.below(step: int) = row.RowBelow(step) |> XlRow
  
  member __.adjust() = row.AdjustToContents() |> XlRow
  member __.adjust(start_column: int) = row.AdjustToContents(start_column) |> XlRow
  member __.adjust(start_column: int, end_column: int) = row.AdjustToContents(start_column, end_column) |> XlRow
  member __.adjust(min_height: float, max_height: float) = row.AdjustToContents(min_height, max_height) |> XlRow
  member __.adjust(start_column: int, min_height: float, max_height: float) = row.AdjustToContents(start_column, min_height, max_height) |> XlRow
  member __.adjust(start_column: int, end_column: int, min_height: float, max_height: float) = row.AdjustToContents(start_column, end_column, min_height, max_height) |> XlRow

  member __.delete() = row.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> row.Clear(opt) | None -> row.Clear()
  member __.hide() = row.Hide()
  member __.unhide() = row.Unhide()

  member __.group() = row.Group() |> XlRow
  member __.group(outline_level: int) = row.Group(outline_level) |> XlRow
  member __.group(collapse: bool) = row.Group(collapse) |> XlRow
  member __.group(outline_level: int, collapse: bool) = row.Group(outline_level, collapse) |> XlRow
  member __.ungroup() = row.Ungroup() |> XlRow
  member __.ungroup(from_all: bool) = row.Ungroup(from_all) |> XlRow
  member __.expand() = row.Expand() |> XlRow
  member __.collapse() = row.Collapse() |> XlRow

  member __.add_horizontal_pagebreak() = row.AddHorizontalPageBreak() |> XlRow
  member __.insert_above(number_of_rows: int) = row.InsertRowsAbove(number_of_rows) |> XlRows
  member __.insert_below(number_of_rows: int) = row.InsertRowsBelow(number_of_rows) |> XlRows

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = row.Cells() |> Seq.map XlCell
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = row.Cells() |> Seq.map XlCell
      cells.GetEnumerator()
    

    
and XlRows internal (rows: IXLRows) =
  member internal __.raw with get() = rows
  member __.cells with get() = rows.Cells() |> XlCells
  member __.used_cells with get() = rows.CellsUsed() |> XlCells
  member __.style with get() = rows.Style |> XlStyle

  member __.adjust() = rows.AdjustToContents() |> XlRows
  member __.adjust(start_column: int) = rows.AdjustToContents(start_column) |> XlRows
  member __.adjust(start_column: int, end_column: int) = rows.AdjustToContents(start_column, end_column) |> XlRows
  member __.adjust(min_height: float, max_height: float) = rows.AdjustToContents(min_height, max_height) |> XlRows
  member __.adjust(start_column: int, min_height: float, max_height: float) = rows.AdjustToContents(start_column, min_height, max_height) |> XlRows
  member __.adjust(start_column: int, end_column: int, min_height: float, max_height: float) = rows.AdjustToContents(start_column, end_column, min_height, max_height) |> XlRows

  member __.delete() = rows.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> rows.Clear(opt) | None -> rows.Clear()
  member __.hide() = rows.Hide()
  member __.unhide() = rows.Unhide()
  
  member __.group() = rows.Group()
  member __.group(outline_level: int) = rows.Group(outline_level)
  member __.group(collapse: bool) = rows.Group(collapse)
  member __.group(outline_level: int, collapse: bool) = rows.Group(outline_level, collapse)
  member __.ungroup() = rows.Ungroup()
  member __.ungroup(from_all: bool) = rows.Ungroup(from_all)
  member __.expand() = rows.Expand()
  member __.collapse() = rows.Collapse()
  
  member __.add_horizontal_pagebreak() = rows.AddHorizontalPageBreaks() |> XlRows

  interface IEnumerable<XlRow> with
    member __.GetEnumerator(): IEnumerator = 
      let rs = rows |> Seq.map XlRow
      (rs :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlRow> =
      let rs = rows |> Seq.map XlRow
      rs.GetEnumerator()
      


and XlRangeRow internal (range: IXLRangeRow) =
  member internal __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.first_cell with get() = range.FirstCell() |> XlCell
  member __.first_used_cell with get() = range.FirstCellUsed() |> XlCell
  member __.last_cell with get() = range.LastCell() |> XlCell
  member __.last_used_cell with get() = range.LastCellUsed() |> XlCell
  member __.row_number with get() = range.RowNumber()
  member __.row_span with get() = range.RangeAddress.RowSpan
  member __.column_span with get() = range.RangeAddress.ColumnSpan
  member __.cell_count with get() = range.CellCount()
  member __.style with get() = range.Style |> XlStyle

  member __.fx(value: obj) = range.FormulaA1 <- value.ToString()
  member __.set(value) = __.value <- box value
  member __.set_formula(value: string) = range.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = range.FormulaR1C1 <- value
  member __.cell(column_number: int) = range.Cell(column_number) |> XlCell
  member __.cell(column_number: string) = range.Cell(column_number) |> XlCell
  member __.cells(first_column: int, last_column:  int) = range.Cells $"%d{first_column}:%d{last_column}" |> XlCells
  member __.cells(cells_in_row: string) = range.Cells cells_in_row |> XlCells
  member __.as_range() = range.AsRange() |> XlRange

  member __.row(start': int, end': int) = range.Row(start', end') |> XlRangeRow
  member __.row(start': IXLCell, end': IXLCell) = range.Row(start', end') |> XlRangeRow
  member __.row(start': XlCell, end': XlCell) = range.Row(start'.raw, end'.raw) |> XlRangeRow
  member __.used_row(?options: CellsUsedOptions) = (match options with Some opt -> range.RowUsed(opt) | None -> range.RowUsed()) |> XlRangeRow
  member __.above() = range.RowAbove() |> XlRangeRow
  member __.above(step: int) = range.RowAbove(step) |> XlRangeRow
  member __.below() = range.RowBelow() |> XlRangeRow
  member __.below(step: int) = range.RowBelow(step) |> XlRangeRow

  member __.insert_cells_after(number_of_rows: int) = range.InsertCellsAfter(number_of_rows) |> XlCells
  member __.insert_cells_after(number_of_rows: int, expand_range: bool) = range.InsertCellsAfter(number_of_rows, expand_range) |> XlCells
  member __.insert_cells_before(number_of_rows: int) = range.InsertCellsBefore(number_of_rows) |> XlCells
  member __.insert_cells_before(number_of_rows: int, expand_range: bool) = range.InsertCellsBefore(number_of_rows, expand_range) |> XlCells
  member __.insert_row_above(number_of_rows: int) = range.InsertRowsAbove(number_of_rows) |> XlRangeRows
  member __.insert_row_above(number_of_rows: int, expand_range: bool) = range.InsertRowsAbove(number_of_rows, expand_range) |> XlRangeRows
  member __.insert_row_below(number_of_rows: int) = range.InsertRowsBelow(number_of_rows) |> XlRangeRows
  member __.insert_row_below(number_of_rows: int, expand_range: bool) = range.InsertRowsBelow(number_of_rows, expand_range) |> XlRangeRows
  
  member __.create_pivot(target: XlCell, name: string) = range.CreatePivotTable(target.raw, name) |> XlPivotTable
  
  member __.copy_to(target: IXLCell) = range.CopyTo(target) |> XlRangeRow
  member __.copy_to(target: XlRangeRow) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeRow
  member __.copy_to(target: XlRange) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeRow
  member __.copy_to(target: XlColumn) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeRow

  member __.delete(?option: ShiftDeleted) = match option with Some opt -> range.Delete(opt) | None -> range.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> range.Clear(opt) | None -> range.Clear()
  


and XlRangeRows internal (range: IXLRangeRows) =
  member internal __.raw with get() = range
  member __.cells() = range.Cells() |> XlCells
  member __.style with get() = range.Style |> XlStyle

  member __.delete() = range.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> range.Clear(opt) | None -> range.Clear()
  
  interface IEnumerable<XlRangeRow> with
    member __.GetEnumerator(): IEnumerator = 
      let rs = range |> Seq.map XlRangeRow
      (rs :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlRangeRow> =
      let rs = range |> Seq.map XlRangeRow
      rs.GetEnumerator()
  


and XlColumn internal (column: IXLColumn) =
  member internal __.raw with get() = column
  member __.value with set(value) = column.Value <- value
  member __.cell_count with get() = column.CellCount
  member __.worksheet with get() = column.Worksheet
  member __.width with get() = column.Width
  member __.first_cell with get() = column.FirstCell() |> XlCell
  member __.first_cell_used with get() = column.FirstCellUsed() |> XlCell
  member __.last_cell with get() = column.LastCell() |> XlCell
  member __.last_cell_used with get() = column.LastCellUsed() |> XlCell
  member __.style with get() = column.Style |> XlStyle

  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = column.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = column.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = column.FormulaR1C1 <- value
  member __.cell(row': int) = column.Cell row' |> XlCell
  member __.cells() = column.Cells() |> XlCells
  member __.cells(row_range: string) = column.Cells row_range |> XlCells
  member __.cells(row': int) = __.cells(row'.ToString())
  member __.cells(from': int, to': int) = __.cells($"%d{from'}:%d{to'}")

  member __.left() = column.ColumnLeft() |> XlColumn
  member __.left(step: int) = column.ColumnLeft(step) |> XlColumn
  member __.right() = column.ColumnRight() |> XlColumn
  member __.right(step: int) = column.ColumnRight(step) |> XlColumn
  
  member __.adjust() = column.AdjustToContents() |> XlColumn
  member __.adjust(start_row: int) = column.AdjustToContents(start_row) |> XlColumn
  member __.adjust(start_row: int, end_row: int) = column.AdjustToContents(start_row, end_row) |> XlColumn
  member __.adjust(min_width: float, max_width: float) = column.AdjustToContents(min_width, max_width) |> XlColumn
  member __.adjust(start_row: int, min_width: float, max_width: float) = column.AdjustToContents(start_row, min_width, max_width) |> XlColumn
  member __.adjust(start_row: int, end_row: int, min_width: float, max_width: float) = column.AdjustToContents(start_row, end_row, min_width, max_width) |> XlColumn
  
  member __.clear(?options: ClearOption) = match options with Some opt -> column.Clear(opt) | None -> column.Clear()
  member __.hide() = column.Hide()
  member __.unhide() = column.Unhide()

  member __.group() = column.Group() |> ignore
  member __.group(outline_level: int) = column.Group(outline_level) |> ignore
  member __.group(collapse: bool) = column.Group(collapse) |> ignore
  member __.group(outline_level: int, collapse: bool) = column.Group(outline_level, collapse) |> ignore
  member __.ungroup() = column.Ungroup() |> ignore
  member __.ungroup(from_all: bool) = column.Ungroup(from_all) |> ignore
  member __.expand() = column.Expand() |> ignore
  member __.collapse() = column.Collapse() |> ignore

  member __.add_vertical_pagebreak() = column.AddVerticalPageBreak() |> XlColumn
  member __.insert_after(number_of_columns: int) = column.InsertColumnsAfter(number_of_columns) |> XlColumns
  member __.insert_before(number_of_columns: int) = column.InsertColumnsBefore(number_of_columns) |> XlColumns

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = column.Cells() |> Seq.map XlCell
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = column.Cells() |> Seq.map XlCell
      cells.GetEnumerator()
      

      
and XlColumns internal (columns: IXLColumns) =
  member internal __.raw with get() = columns
  member __.cells() = columns.Cells() |> XlCells
  member __.used_cells() = columns.CellsUsed() |> XlCells
  member __.style with get() = columns.Style |> XlStyle
  member __.set_width(width: float) = columns.Width <- width
  
  member __.adjust() = columns.AdjustToContents() |> XlColumns
  member __.adjust(start_row: int) = columns.AdjustToContents(start_row) |> XlColumns
  member __.adjust(start_row: int, end_row: int) = columns.AdjustToContents(start_row, end_row) |> XlColumns
  member __.adjust(min_width: float, max_width: float) = columns.AdjustToContents(min_width, max_width) |> XlColumns
  member __.adjust(start_row: int, min_width: float, max_width: float) = columns.AdjustToContents(start_row, min_width, max_width) |> XlColumns
  member __.adjust(start_row: int, end_row: int, min_width: float, max_width: float) = columns.AdjustToContents(start_row, end_row, min_width, max_width) |> XlColumns
  
  member __.delete() = columns.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> columns.Clear(opt) | None -> columns.Clear()
  member __.hide() = columns.Hide()
  member __.unhide() = columns.Unhide()
  
  member __.group() = columns.Group()
  member __.group(outline_level: int) = columns.Group(outline_level)
  member __.group(collapse: bool) = columns.Group(collapse)
  member __.group(outline_level: int, collapse: bool) = columns.Group(outline_level, collapse)
  member __.ungroup() = columns.Ungroup()
  member __.ungroup(from_all: bool) = columns.Ungroup(from_all)
  member __.expand() = columns.Expand()
  member __.collapse() = columns.Collapse()

  interface IEnumerable<XlColumn> with
    member __.GetEnumerator(): IEnumerator = 
      let columns = columns |> Seq.map XlColumn
      (columns :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlColumn> = 
      let columns = columns |> Seq.map XlColumn
      columns.GetEnumerator()
  
  

and XlRangeColumn internal (range: IXLRangeColumn) =
  member internal __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.first_cell with get() = range.FirstCell() |> XlCell
  member __.first_used_cell with get() = range.FirstCellUsed() |> XlCell
  member __.last_cell with get() = range.LastCell() |> XlCell
  member __.last_used_cell with get() = range.LastCellUsed() |> XlCell
  member __.column_number with get() = range.ColumnNumber()
  member __.row_span with get() = range.RangeAddress.RowSpan
  member __.column_span with get() = range.RangeAddress.ColumnSpan
  member __.cell_count with get() = range.CellCount()
  member __.style with get() = range.Style |> XlStyle
    
  member __.as_table() = range.AsTable() |> XlTable
  member __.as_table(name: string) = range.AsTable(name) |> XlTable
  member __.cell(row_number: int) = range.Cell(row_number) |> XlCell
  member __.cells(first_row: int, last_row: int) = range.Cells(first_row, last_row) |> XlCells
  member __.cells(cell_in_column: string) = range.Cells(cell_in_column) |> XlCells
  
  member __.column(start': int, end': int) = range.Column(start', end') |> XlRangeColumn
  member __.column(start': XlCell, end': XlCell) = range.Column(start'.raw, end'.raw) |> XlRangeColumn
  member __.used_column(?options: CellsUsedOptions) =
    match options with Some opt -> range.ColumnUsed(opt) | None -> range.ColumnUsed()
    |> XlRangeColumn
  member __.left() = range.ColumnLeft() |> XlRangeColumn
  member __.left(step: int) = range.ColumnLeft(step) |> XlRangeColumn
  member __.right() = range.ColumnRight() |> XlRangeColumn
  member __.right(step: int) = range.ColumnRight(step) |> XlRangeColumn
  
  member __.insert_cells_above(number_of_rows: int) = range.InsertCellsAbove(number_of_rows) |> XlCells
  member __.insert_cells_above(number_of_rows: int, expand_range: bool) = range.InsertCellsAbove(number_of_rows, expand_range) |> XlCells
  member __.insert_cells_below(number_of_rows: int) = range.InsertCellsBelow(number_of_rows) |> XlCells
  member __.insert_cells_below(number_of_rows: int, expand_range: bool) = range.InsertCellsBelow(number_of_rows, expand_range) |> XlCells
  member __.insert_column_after(number_of_columns: int) = range.InsertColumnsAfter(number_of_columns) |> XlRangeColumns
  member __.insert_column_after(number_of_columns: int, expand_range: bool) = range.InsertColumnsAfter(number_of_columns, expand_range) |> XlRangeColumns
  member __.insert_column_before(number_of_columns: int) = range.InsertColumnsBefore(number_of_columns) |> XlRangeColumns
  member __.insert_column_before(number_of_columns: int, expand_range: bool) = range.InsertColumnsBefore(number_of_columns, expand_range) |> XlRangeColumns

  member __.copy_to(target: XlRangeColumn) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeColumn
  member __.copy_to(target: XlRange) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeColumn
  member __.copy_to(target: XlCell) = range.CopyTo(target.raw) |> XlRangeColumn
  member __.copy_to(target: IXLCell) = range.CopyTo(target) |> XlRangeColumn

  member __.create_table() = range.CreateTable() |> XlTable
  member __.create_table(name: string) = range.CreateTable(name) |> XlTable
  member __.create_pivot(target: XlCell, name: string) = range.CreatePivotTable(target.raw, name) |> XlPivotTable

  member __.clear() = range.Clear()
  member __.delete() = range.Delete()
  member __.delete_comments() = range.DeleteComments()



and XlRangeColumns internal (range: IXLRangeColumns) =
  member internal __.raw with get() = range
  member __.cells() = range.Cells() |> XlCells
  member __.used_cells(?options: CellsUsedOptions) = 
    match options with Some opt -> range.CellsUsed(opt) | None -> range.CellsUsed()
    |> XlCells
  member __.style with get() = range.Style |> XlStyle
  
  member __.delete() = range.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> range.Clear(opt) | None -> range.Clear()

  interface IEnumerable<XlRangeColumn> with
    member __.GetEnumerator(): IEnumerator = 
      let cs = range |> Seq.map XlRangeColumn
      (cs :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlRangeColumn> =
      let cs = range |> Seq.map XlRangeColumn
      cs.GetEnumerator()
    


and XlRange internal (range: IXLRange) =
  member internal __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.first_cell with get() = range.FirstCell() |> XlCell
  member __.first_used_cell with get() = range.FirstCellUsed() |> XlCell
  member __.last_cell with get() = range.LastCell() |> XlCell
  member __.last_used_cell with get() = range.LastCellUsed() |> XlCell
  member __.style with get() = range.Style |> XlStyle
  member __.column_count() = range.ColumnCount()
  member __.row_count() = range.RowCount()

  member __.column(number: int) = range.Column(number) |> XlRangeColumn
  member __.column(letter: string) = range.Column(letter) |> XlRangeColumn
  member __.columns(first: int, last: int) = range.Columns(first, last) |> XlRangeColumns
  member __.columns(columns: string) = range.Columns(columns) |> XlRangeColumns
  member __.used_columns(?predicate: XlRangeColumn -> bool) =
    match predicate with
      | Some predicate' -> range.ColumnsUsed(fun column -> predicate'(XlRangeColumn column))
      | None -> range.ColumnsUsed()
    |> XlRangeColumns
  member __.row(row: int) = range.Row(row) |> XlRangeRow
  member __.rows(first: int, last: int) = range.Rows(first, last) |> XlRangeRows
  member __.rows(rows: string) = range.Rows(rows) |> XlRangeRows
  member __.used_rows(?predicate: XlRangeRow -> bool) =
    match predicate with
      | Some predicate' -> range.RowsUsed(fun row -> predicate'(XlRangeRow row))
      | None -> range.RowsUsed()
    |> XlRangeRows

  member __.fx(value: obj) = range.FormulaA1 <- value.ToString()
  member __.set(value) = __.value <- box value
  member __.set_formula(value: string) = range.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = range.FormulaR1C1 <- value
  member __.cell(row: int, column: int) = range.Cell(row, column) |> XlCell
  member __.cell(address: string) = range.Cell address |> XlCell
  member __.cells(from': Address, to': Address) = range.Cells $"%s{from'.to_string()}:%s{to'.to_string()}" |> XlCells
  member __.cells(from': int * int, to':  int * int) = range.Cells $"%s{from'.to_address()}:%s{to'.to_address()}" |> XlCells
  member __.cells(address: string) = range.Cells address |> XlCells
  member __.insert_after(number_of_columns: int) = range.InsertColumnsAfter(number_of_columns) |> XlRangeColumns
  member __.insert_before(number_of_columns: int) = range.InsertColumnsBefore(number_of_columns) |> XlRangeColumns
  member __.insert_above(number_of_rows: int) = range.InsertRowsAbove(number_of_rows) |> XlRangeRows
  member __.insert_below(number_of_rows: int) = range.InsertRowsBelow(number_of_rows) |> XlRangeRows
  
  member __.find_column(predicate: XlRangeColumn -> bool) = range.FindColumn(fun column -> predicate(XlRangeColumn column)) |> XlRangeColumn
  member __.find_row(predicate: XlRangeRow -> bool) = range.FindRow(fun row -> predicate(XlRangeRow row)) |> XlRangeRow
  member __.first_column(?predicate: XlRangeColumn -> bool) =
    match predicate with
      | Some predicate' -> range.FirstColumn(fun column -> predicate'(XlRangeColumn column))
      | None -> range.FirstColumn()
    |> XlRangeColumn
  member __.first_used_column(?predicate: XlRangeColumn -> bool) =
    match predicate with
      | Some predicate' -> range.FirstColumnUsed(fun column -> predicate'(XlRangeColumn column))
      | None -> range.FirstColumnUsed()
    |> XlRangeColumn
  member __.first_row(?predicate: XlRangeRow -> bool) =
    match predicate with
      | Some predicate' -> range.FirstRow(fun row -> predicate'(XlRangeRow row))
      | None -> range.FirstRow()
    |> XlRangeRow
  member __.first_used_row(?predicate: XlRangeRow -> bool) =
    match predicate with
      | Some predicate' -> range.FirstRowUsed(fun row -> predicate'(XlRangeRow row))
      | None -> range.FirstRowUsed()
    |> XlRangeRow
  member __.last_column(?predicate: XlRangeColumn -> bool) =
    match predicate with
      | Some predicate' -> range.LastColumn(fun column -> predicate'(XlRangeColumn column))
      | None -> range.LastColumn()
    |> XlRangeColumn
  member __.last_used_column(?predicate: XlRangeColumn -> bool) =
    match predicate with
      | Some predicate' -> range.LastColumnUsed(fun column -> predicate'(XlRangeColumn column))
      | None -> range.LastColumnUsed()
    |> XlRangeColumn
  member __.last_row(?predicate: XlRangeRow -> bool) =
    match predicate with
      | Some predicate' -> range.LastRow(fun row -> predicate'(XlRangeRow row))
      | None -> range.LastRow()
    |> XlRangeRow
  member __.last_used_row(?predicate: XlRangeRow -> bool) =
    match predicate with
      | Some predicate' -> range.LastRowUsed(fun row -> predicate'(XlRangeRow row))
      | None -> range.LastRowUsed()
    |> XlRangeRow


  member __.create_table(?name: string) =
    match name with Some name' -> range.CreateTable(name') | None -> range.CreateTable()
    |> XlTable

  member __.clear(?options: ClearOption) = match options with Some opt -> range.Clear(opt) | None -> range.Clear()
  member __.delete(option: ShiftDeleted) = range.Delete(option)

  

// TODO
and XlTable internal (table: IXLTable) =
  member internal __.raw with get() = table
  member __.name with get() = table.Name
  member __.fields() = table.Fields |> Seq.map XlTableField
  member __.header() = table.HeadersRow() |> XlRangeRow

  member __.append(data: DataTable, ?propagate_extra_columns: bool) =
    let flag = match propagate_extra_columns with Some propagate_extra_columns' -> propagate_extra_columns' | None -> false
    table.AppendData(data, flag) |> ignore
  member __.append(data: seq<'T>, ?propagate_extra_columns: bool) =
    let flag = match propagate_extra_columns with Some propagate_extra_columns' -> propagate_extra_columns' | None -> false
    table.AppendData(data, flag) |> ignore
  member __.replace(data: DataTable, ?propagate_extra_columns: bool) =
    let flag = match propagate_extra_columns with Some propagate_extra_columns' -> propagate_extra_columns' | None -> false
    table.ReplaceData(data, flag) |> ignore
  member __.replace(data: seq<'T>, ?propagate_extra_columns: bool) =
    let flag = match propagate_extra_columns with Some propagate_extra_columns' -> propagate_extra_columns' | None -> false
    table.ReplaceData(data, flag) |> ignore
  member __.as_datatable() = table.AsNativeDataTable()

  // TODO
  member __.style with get() = table.Style
  
  member __.clear(?options: ClearOption) = 
    match options with Some opt -> table.Clear(opt) | None -> table.Clear()
    |> ignore
  member __.delete(option: ShiftDeleted) = table.Delete(option)
  member __.delete_comments() = table.DeleteComments()
  


// TODO
and XlTableField internal (field: IXLTableField) =
  member internal __.raw with get() = field



// TODO
and XlPivotTable internal (table: IXLPivotTable) =
  member internal __.raw with get() = table
  

  
// TODO
and XlStyle internal (style: IXLStyle) =
  member internal __.raw with get() = style
