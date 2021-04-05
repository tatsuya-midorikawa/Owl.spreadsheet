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

  

and XlTable internal (table: IXLTable) =
  member internal __.raw with get() = table
  member __.auto_filter with get() = table.AutoFilter |> XlAutoFilter
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
  member __.as_range() = table :> IXLRange |> XlRange

  member __.style with get() = table.Style |> XlStyle
  member __.resize(first: XlCell, last: XlCell) = table.Resize(first.raw, last.raw) |> ignore
  member __.resize(first: IXLCell, last: IXLCell) = table.Resize(first, last) |> ignore
  member __.resize(first_address: string, last_address: string) = table.Resize(first_address, last_address) |> ignore
  member __.resize(first_row: int, first_column: int, last_row: int, last_column: int) = table.Resize(first_row, first_column, last_row, last_column) |> ignore

  member __.set_auto_filter() = table.SetAutoFilter() |> ignore

  member __.clear(?options: ClearOption) = 
    match options with Some opt -> table.Clear(opt) | None -> table.Clear()
    |> ignore
  member __.delete(option: ShiftDeleted) = table.Delete(option)
  member __.delete_comments() = table.DeleteComments()
  


and XlTableField internal (field: IXLTableField) =
  member internal __.raw with get() = field
  member __.index with get() = field.Index
  member __.name with get() = field.Name
  member __.table with get() = field.Table |> XlTable
  member __.cells() = field.DataCells |> XlCells
  member __.totals_cell() = field.TotalsCell |> XlCell
  member __.header() = field.HeaderCell |> XlCell
  member __.get_formula() = field.TotalsRowFormulaA1
  member __.fx(value: obj) =
    field.TotalsRowFormulaA1 <- match value.GetType() with t when t = typeof<string> -> value |> unbox<string> | _ -> $"{value}"
  member __.delete() = field.Delete()


                                
// TODO
and XlPivotTable internal (table: IXLPivotTable) =
  member internal __.raw with get() = table
  

  
and XlStyle internal (style: IXLStyle) =
  member internal __.raw with get() = style
  member __.alignment with get() = style.Alignment |> XlAlignment
  member __.border with get() = style.Border |> XlBorder
  member __.date_format with get() = style.DateFormat |> XlNumberFormat
  member __.fill with get() = style.Fill |> XlFill
  member __.font with get() = style.Font |> XlFont
  member __.include_quote_prefix with get() = style.IncludeQuotePrefix
  member __.number_format with get() = style.NumberFormat |> XlNumberFormat
  member __.protection with get() = style.Protection |> XlProtection

  member __.set_alignment(alignment: XlAlignment) = style.Alignment <- alignment.raw
  member __.set_border(border: XlBorder) = style.Border <- border.raw
  member __.set_fill(fill: XlFill) = style.Fill <- fill.raw
  member __.set_font(font: XlFont) = style.Font <- font.raw
  member __.set_include_quote_prefix(value: bool) = style.IncludeQuotePrefix <- value
  member __.set_number_format(format: XlNumberFormat) = style.NumberFormat <- format.raw
  member __.set_protection(protection: XlProtection) = style.Protection <- protection.raw
  

  
and XlAlignment internal (align: IXLAlignment) =
  member internal __.raw with get() = align
  member __.top_to_bottom with get() = align.TopToBottom
  member __.text_rotation with get() = align.TextRotation
  member __.shrink_to_fit with get() = align.ShrinkToFit
  member __.relative_indent with get() = align.RelativeIndent
  member __.justify_lastline with get() = align.JustifyLastLine
  member __.indent with get() = align.Indent
  member __.wrap_text with get() = align.WrapText
  member __.reading_order with get() : AlignmentReadingOrderValues = align.ReadingOrder
  member __.vertical with get() : AlignmentVerticalValues = align.Vertical
  member __.horizontal with get() : AlignmentHorizontalValues = align.Horizontal

  member __.set_top_to_bottom(value: bool) = align.TopToBottom <- value
  member __.set_text_rotation(value: int) = align.TextRotation <- value
  member __.set_shrink_to_fit(value: bool) = align.ShrinkToFit <- value
  member __.set_relative_indent(value: int) = align.RelativeIndent <- value
  member __.set_justify_lastline(value: bool) = align.JustifyLastLine <- value
  member __.set_indent(value: int) = align.Indent <- value
  member __.set_wrap_text(value: bool) = align.WrapText <- value
  member __.set_reading_order(value: AlignmentReadingOrderValues) = align.ReadingOrder <- value
  member __.set_vertical(vertical: AlignmentVerticalValues) = align.Vertical <- vertical
  member __.set_horizontal(horizontal: AlignmentHorizontalValues) = align.Horizontal <- horizontal



and XlBorder internal (border: IXLBorder) =
  member internal __.raw with get() = border
  member __.diagonal_color with get() = border.DiagonalBorderColor |> XlColor
  member __.top_color with get() = border.TopBorderColor |> XlColor
  member __.bottom_color with get() = border.BottomBorderColor |> XlColor
  member __.left_color with get() = border.LeftBorderColor |> XlColor
  member __.right_color with get() = border.RightBorderColor |> XlColor
  member __.diagonal with get() : BorderStyleValues = border.DiagonalBorder
  member __.top with get() : BorderStyleValues = border.TopBorder
  member __.bottom with get() : BorderStyleValues = border.BottomBorder
  member __.left with get() : BorderStyleValues = border.LeftBorder
  member __.right with get() : BorderStyleValues = border.RightBorder
  member __.diagonal_up with get() = border.DiagonalUp
  member __.diagonal_down with get() = border.DiagonalDown

  member __.set_diagonal_color(color: XlColor) = border.DiagonalBorderColor <- color.raw
  member __.set_outside_color(color: XlColor) = border.OutsideBorderColor <- color.raw
  member __.set_inside_color(color: XlColor) = border.InsideBorderColor <- color.raw
  member __.set_top_color(color: XlColor) = border.TopBorderColor <- color.raw
  member __.set_bottom_color(color: XlColor) = border.BottomBorderColor <- color.raw
  member __.set_left_color(color: XlColor) = border.LeftBorderColor <- color.raw
  member __.set_right_color(color: XlColor) = border.RightBorderColor <- color.raw
  member __.set_diagonal(value: BorderStyleValues) = border.DiagonalBorder <- value
  member __.set_top(value: BorderStyleValues) = border.TopBorder <- value
  member __.set_bottom(value: BorderStyleValues) = border.BottomBorder <- value
  member __.set_left(value: BorderStyleValues) = border.LeftBorder <- value
  member __.set_right(value: BorderStyleValues) = border.RightBorder <- value
  member __.set_diagonal_up(value: bool) = border.DiagonalUp <- value
  member __.set_diagonal_down(value: bool) = border.DiagonalDown <- value



and XlColor internal (color: XLColor) =
  member internal __.raw with get() = color
  member __.value with get() = color.Color
  member __.color_type with get() : ColorType = color.ColorType
  member __.tint with get() = color.ThemeTint

  static member from_argb(argb: int) = XLColor.FromArgb(argb) |> XlColor
  static member from_argb(a: int, r: int, g: int, b: int) = XLColor.FromArgb(a, r, g, b) |> XlColor
  static member from_rgb(r: int, g: int, b: int) = XLColor.FromArgb(r, g, b) |> XlColor
  static member from_color(color: System.Drawing.Color) = XLColor.FromColor(color) |> XlColor
  static member from_theme(theme: ThemeColor) = XLColor.FromTheme(theme) |> XlColor

  static member forest_green_traditional with get() = XLColor.ForestGreenTraditional |> XlColor
  static member myrtle with get() = XLColor.Myrtle |> XlColor
  static member nadeshiko_pink with get() = XLColor.NadeshikoPink |> XlColor
  static member napier_green with get() = XLColor.NapierGreen |> XlColor
  static member naples_yellow with get() = XLColor.NaplesYellow |> XlColor
  static member neon_carrot with get() = XLColor.NeonCarrot |> XlColor
  static member neon_fuchsia with get() = XLColor.NeonFuchsia |> XlColor
  static member neon_green with get() = XLColor.NeonGreen |> XlColor
  static member non_photo_blue with get() = XLColor.NonPhotoBlue |> XlColor
  static member ocean_boat_blue with get() = XLColor.OceanBoatBlue |> XlColor
  static member ochre with get() = XLColor.Ochre |> XlColor
  static member old_gold with get() = XLColor.OldGold |> XlColor
  static member old_lavender with get() = XLColor.OldLavender |> XlColor
  static member old_mauve with get() = XLColor.OldMauve |> XlColor
  static member old_rose with get() = XLColor.OldRose |> XlColor
  static member olive_drab7 with get() = XLColor.OliveDrab7 |> XlColor
  static member olivine with get() = XLColor.Olivine |> XlColor
  static member onyx with get() = XLColor.Onyx |> XlColor
  static member opera_mauve with get() = XLColor.OperaMauve |> XlColor
  static member orange_color_wheel with get() = XLColor.OrangeColorWheel |> XlColor
  static member mustard with get() = XLColor.Mustard |> XlColor
  static member mulberry with get() = XLColor.Mulberry |> XlColor
  static member msu_green with get() = XLColor.MsuGreen |> XlColor
  static member mountbatten_pink with get() = XLColor.MountbattenPink |> XlColor
  static member medium_aquamarine1 with get() = XLColor.MediumAquamarine1 |> XlColor
  static member medium_candy_apple_red with get() = XLColor.MediumCandyAppleRed |> XlColor
  static member medium_carmine with get() = XLColor.MediumCarmine |> XlColor
  static member medium_champagne with get() = XLColor.MediumChampagne |> XlColor
  static member medium_electric_blue with get() = XLColor.MediumElectricBlue |> XlColor
  static member medium_jungle_green with get() = XLColor.MediumJungleGreen |> XlColor
  static member medium_persian_blue with get() = XLColor.MediumPersianBlue |> XlColor
  static member medium_redViolet with get() = XLColor.MediumRedViolet |> XlColor
  static member medium_spring_bud with get() = XLColor.MediumSpringBud |> XlColor
  static member orange_peel with get() = XLColor.OrangePeel |> XlColor
  static member medium_taupe with get() = XLColor.MediumTaupe |> XlColor
  static member midnight_green_eagle_green with get() = XLColor.MidnightGreenEagleGreen |> XlColor
  static member mikado_yellow with get() = XLColor.MikadoYellow |> XlColor
  static member mint with get() = XLColor.Mint |> XlColor
  static member mint_green with get() = XLColor.MintGreen |> XlColor
  static member mode_beige with get() = XLColor.ModeBeige |> XlColor
  static member moonstone_blue with get() = XLColor.MoonstoneBlue |> XlColor
  static member mordant_red19 with get() = XLColor.MordantRed19 |> XlColor
  static member moss_green with get() = XLColor.MossGreen |> XlColor
  static member mountain_meadow with get() = XLColor.MountainMeadow |> XlColor
  static member melon with get() = XLColor.Melon |> XlColor
  static member meat_brown with get() = XLColor.MeatBrown |> XlColor
  static member orange_ryb with get() = XLColor.OrangeRyb |> XlColor
  static member ou_crimson_red with get() = XLColor.OuCrimsonRed |> XlColor
  static member pastel_gray with get() = XLColor.PastelGray |> XlColor
  static member pastel_green with get() = XLColor.PastelGreen |> XlColor
  static member pastel_magenta with get() = XLColor.PastelMagenta |> XlColor
  static member pastel_orange with get() = XLColor.PastelOrange |> XlColor
  static member pastel_pink with get() = XLColor.PastelPink |> XlColor
  static member pastel_purple with get() = XLColor.PastelPurple |> XlColor
  static member pastel_red with get() = XLColor.PastelRed |> XlColor
  static member pastel_violet with get() = XLColor.PastelViolet |> XlColor
  static member pastel_yellow with get() = XLColor.PastelYellow |> XlColor
  static member paynes_grey with get() = XLColor.PaynesGrey |> XlColor
  static member peach with get() = XLColor.Peach |> XlColor
  static member peach_orange with get() = XLColor.PeachOrange |> XlColor
  static member peach_yellow with get() = XLColor.PeachYellow |> XlColor
  static member pear with get() = XLColor.Pear |> XlColor
  static member pearl with get() = XLColor.Pearl |> XlColor
  static member peridot with get() = XLColor.Peridot |> XlColor
  static member periwinkle with get() = XLColor.Periwinkle |> XlColor
  static member persian_blue with get() = XLColor.PersianBlue |> XlColor
  static member persian_green with get() = XLColor.PersianGreen |> XlColor
  static member pastel_brown with get() = XLColor.PastelBrown |> XlColor
  static member pastel_blue with get() = XLColor.PastelBlue |> XlColor
  static member paris_green with get() = XLColor.ParisGreen |> XlColor
  static member pansy_purple with get() = XLColor.PansyPurple |> XlColor
  static member outer_space with get() = XLColor.OuterSpace |> XlColor
  static member outrageous_orange with get() = XLColor.OutrageousOrange |> XlColor
  static member oxford_blue with get() = XLColor.OxfordBlue |> XlColor
  static member pakistan_green with get() = XLColor.PakistanGreen |> XlColor
  static member palatinate_blue with get() = XLColor.PalatinateBlue |> XlColor
  static member palatinate_purple with get() = XLColor.PalatinatePurple |> XlColor
  static member pale_aqua with get() = XLColor.PaleAqua |> XlColor
  static member pale_brown with get() = XLColor.PaleBrown |> XlColor
  static member pale_carmine with get() = XLColor.PaleCarmine |> XlColor
  static member otter_brown with get() = XLColor.OtterBrown |> XlColor
  static member pale_cerulean with get() = XLColor.PaleCerulean |> XlColor
  static member pale_copper with get() = XLColor.PaleCopper |> XlColor
  static member pale_cornflower_blue with get() = XLColor.PaleCornflowerBlue |> XlColor
  static member pale_gold with get() = XLColor.PaleGold |> XlColor
  static member pale_magenta with get() = XLColor.PaleMagenta |> XlColor
  static member pale_pink with get() = XLColor.PalePink |> XlColor
  static member pale_robin_egg_blue with get() = XLColor.PaleRobinEggBlue |> XlColor
  static member pale_silver with get() = XLColor.PaleSilver |> XlColor
  static member pale_spring_bud with get() = XLColor.PaleSpringBud |> XlColor
  static member pale_taupe with get() = XLColor.PaleTaupe |> XlColor
  static member pale_chestnut with get() = XLColor.PaleChestnut |> XlColor
  static member maya_blue with get() = XLColor.MayaBlue |> XlColor
  static member mauve_taupe with get() = XLColor.MauveTaupe |> XlColor
  static member mauvelous with get() = XLColor.Mauvelous |> XlColor
  static member harvest_gold with get() = XLColor.HarvestGold |> XlColor
  static member heliotrope with get() = XLColor.Heliotrope |> XlColor
  static member hollywood_cerise with get() = XLColor.HollywoodCerise |> XlColor
  static member hookers_green with get() = XLColor.HookersGreen |> XlColor
  static member hotMagenta with get() = XLColor.HotMagenta |> XlColor
  static member hunter_green with get() = XLColor.HunterGreen |> XlColor
  static member iceberg with get() = XLColor.Iceberg |> XlColor
  static member icterine with get() = XLColor.Icterine |> XlColor
  static member inchworm with get() = XLColor.Inchworm |> XlColor
  static member india_green with get() = XLColor.IndiaGreen |> XlColor
  static member indian_yellow with get() = XLColor.IndianYellow |> XlColor
  static member indigo_dye with get() = XLColor.IndigoDye |> XlColor
  static member international_kleinBlue with get() = XLColor.InternationalKleinBlue |> XlColor
  static member international_orange with get() = XLColor.InternationalOrange |> XlColor
  static member iris with get() = XLColor.Iris |> XlColor
  static member isabelline with get() = XLColor.Isabelline |> XlColor
  static member islamic_green with get() = XLColor.IslamicGreen |> XlColor
  static member jade with get() = XLColor.Jade |> XlColor
  static member jasper with get() = XLColor.Jasper |> XlColor
  static member harvard_crimson with get() = XLColor.HarvardCrimson |> XlColor
  static member harlequin with get() = XLColor.Harlequin |> XlColor
  static member hansa_yellow with get() = XLColor.HansaYellow |> XlColor
  static member han_purple with get() = XLColor.HanPurple |> XlColor
  static member french_blue with get() = XLColor.FrenchBlue |> XlColor
  static member french_lilac with get() = XLColor.FrenchLilac |> XlColor
  static member french_rose with get() = XLColor.FrenchRose |> XlColor
  static member fuchsia_pink with get() = XLColor.FuchsiaPink |> XlColor
  static member fulvous with get() = XLColor.Fulvous |> XlColor
  static member fuzzy_wuzzy with get() = XLColor.FuzzyWuzzy |> XlColor
  static member gamboge with get() = XLColor.Gamboge |> XlColor
  static member ginger with get() = XLColor.Ginger |> XlColor
  static member glaucous with get() = XLColor.Glaucous |> XlColor
  static member jazzberry_jam with get() = XLColor.JazzberryJam |> XlColor
  static member golden_brown with get() = XLColor.GoldenBrown |> XlColor
  static member golden_yellow with get() = XLColor.GoldenYellow |> XlColor
  static member gold_metallic with get() = XLColor.GoldMetallic |> XlColor
  static member granny_smith_apple with get() = XLColor.GrannySmithApple |> XlColor
  static member gray_asparagus with get() = XLColor.GrayAsparagus |> XlColor
  static member green_pigment with get() = XLColor.GreenPigment |> XlColor
  static member green_ryb with get() = XLColor.GreenRyb |> XlColor
  static member grullo with get() = XLColor.Grullo |> XlColor
  static member halaya_ube with get() = XLColor.HalayaUbe |> XlColor
  static member han_blue with get() = XLColor.HanBlue |> XlColor
  static member golden_poppy with get() = XLColor.GoldenPoppy |> XlColor
  static member jonquil with get() = XLColor.Jonquil |> XlColor
  static member june_bud with get() = XLColor.JuneBud |> XlColor
  static member jungle_green with get() = XLColor.JungleGreen |> XlColor
  static member light_thulian_pink with get() = XLColor.LightThulianPink |> XlColor
  static member light_yellow1 with get() = XLColor.LightYellow1 |> XlColor
  static member lilac with get() = XLColor.Lilac |> XlColor
  static member lime_color_wheel with get() = XLColor.LimeColorWheel |> XlColor
  static member lincoln_green with get() = XLColor.LincolnGreen |> XlColor
  static member liver with get() = XLColor.Liver |> XlColor
  static member lust with get() = XLColor.Lust |> XlColor
  static member macaroni_and_cheese with get() = XLColor.MacaroniAndCheese |> XlColor
  static member magenta_dye with get() = XLColor.MagentaDye |> XlColor
  static member light_taupe with get() = XLColor.LightTaupe |> XlColor
  static member magenta_process with get() = XLColor.MagentaProcess |> XlColor
  static member magnolia with get() = XLColor.Magnolia |> XlColor
  static member mahogany with get() = XLColor.Mahogany |> XlColor
  static member maize with get() = XLColor.Maize |> XlColor
  static member majorelle_blue with get() = XLColor.MajorelleBlue |> XlColor
  static member malachite with get() = XLColor.Malachite |> XlColor
  static member manatee with get() = XLColor.Manatee |> XlColor
  static member mango_tango with get() = XLColor.MangoTango |> XlColor
  static member maroon_x11 with get() = XLColor.MaroonX11 |> XlColor
  static member mauve with get() = XLColor.Mauve |> XlColor
  static member magic_mint with get() = XLColor.MagicMint |> XlColor
  static member persian_indigo with get() = XLColor.PersianIndigo |> XlColor
  static member light_salmon_pink with get() = XLColor.LightSalmonPink |> XlColor
  static member light_mauve with get() = XLColor.LightMauve |> XlColor
  static member kelly_green with get() = XLColor.KellyGreen |> XlColor
  static member khaki_html_css_khaki with get() = XLColor.KhakiHtmlCssKhaki |> XlColor
  static member languid_lavender with get() = XLColor.LanguidLavender |> XlColor
  static member lapis_lazuli with get() = XLColor.LapisLazuli |> XlColor
  static member la_salle_green with get() = XLColor.LaSalleGreen |> XlColor
  static member laser_lemon with get() = XLColor.LaserLemon |> XlColor
  static member lava with get() = XLColor.Lava |> XlColor
  static member lavender_blue with get() = XLColor.LavenderBlue |> XlColor
  static member lavender_floral with get() = XLColor.LavenderFloral |> XlColor
  static member light_pastel_purple with get() = XLColor.LightPastelPurple |> XlColor
  static member lavender_gray with get() = XLColor.LavenderGray |> XlColor
  static member lavender_pink with get() = XLColor.LavenderPink |> XlColor
  static member lavender_purple with get() = XLColor.LavenderPurple |> XlColor
  static member lavender_rose with get() = XLColor.LavenderRose |> XlColor
  static member lemon with get() = XLColor.Lemon |> XlColor
  static member light_apricot with get() = XLColor.LightApricot |> XlColor
  static member light_brown with get() = XLColor.LightBrown |> XlColor
  static member light_carmine_pink with get() = XLColor.LightCarminePink |> XlColor
  static member light_cornflower_blue with get() = XLColor.LightCornflowerBlue |> XlColor
  static member light_fuchsia_pink with get() = XLColor.LightFuchsiaPink |> XlColor
  static member lavender_indigo with get() = XLColor.LavenderIndigo |> XlColor
  static member persian_orange with get() = XLColor.PersianOrange |> XlColor
  static member persian_pink with get() = XLColor.PersianPink |> XlColor
  static member persian_plum with get() = XLColor.PersianPlum |> XlColor
  static member tenné_tawny with get() = XLColor.TennéTawny |> XlColor
  static member terra_cotta with get() = XLColor.TerraCotta |> XlColor
  static member thulian_pink with get() = XLColor.ThulianPink |> XlColor
  static member tickle_me_pink with get() = XLColor.TickleMePink |> XlColor
  static member tiffany_blue with get() = XLColor.TiffanyBlue |> XlColor
  static member tigers_eye with get() = XLColor.TigersEye |> XlColor
  static member timberwolf with get() = XLColor.Timberwolf |> XlColor
  static member titanium_yellow with get() = XLColor.TitaniumYellow |> XlColor
  static member toolbox with get() = XLColor.Toolbox |> XlColor
  static member tractor_red with get() = XLColor.TractorRed |> XlColor
  static member tropical_rain_forest with get() = XLColor.TropicalRainForest |> XlColor
  static member tufts_blue with get() = XLColor.TuftsBlue |> XlColor
  static member tumbleweed with get() = XLColor.Tumbleweed |> XlColor
  static member turkish_rose with get() = XLColor.TurkishRose |> XlColor
  static member turquoise1 with get() = XLColor.Turquoise1 |> XlColor
  static member turquoise_blue with get() = XLColor.TurquoiseBlue |> XlColor
  static member turquoise_green with get() = XLColor.TurquoiseGreen |> XlColor
  static member tuscan_red with get() = XLColor.TuscanRed |> XlColor
  static member twilightLavender with get() = XLColor.TwilightLavender |> XlColor
  static member tea_rose_rose with get() = XLColor.TeaRoseRose |> XlColor
  static member tea_rose_orange with get() = XLColor.TeaRoseOrange |> XlColor
  static member teal_green with get() = XLColor.TealGreen |> XlColor
  static member teal_blue with get() = XLColor.TealBlue |> XlColor
  static member sinopia with get() = XLColor.Sinopia |> XlColor
  static member skobeloff with get() = XLColor.Skobeloff |> XlColor
  static member sky_magenta with get() = XLColor.SkyMagenta |> XlColor
  static member smalt_dark_powder_blue with get() = XLColor.SmaltDarkPowderBlue |> XlColor
  static member smokey_topaz with get() = XLColor.SmokeyTopaz |> XlColor
  static member smoky_black with get() = XLColor.SmokyBlack |> XlColor
  static member spiro_disco_ball with get() = XLColor.SpiroDiscoBall |> XlColor
  static member splashed_white with get() = XLColor.SplashedWhite |> XlColor
  static member spring_bud with get() = XLColor.SpringBud |> XlColor
  static member tyrian_purple with get() = XLColor.TyrianPurple |> XlColor
  static member st_patricks_blue with get() = XLColor.StPatricksBlue |> XlColor
  static member straw with get() = XLColor.Straw |> XlColor
  static member sunglow with get() = XLColor.Sunglow |> XlColor
  static member sunset with get() = XLColor.Sunset |> XlColor
  static member tangelo with get() = XLColor.Tangelo |> XlColor
  static member tangerine with get() = XLColor.Tangerine |> XlColor
  static member tangerine_yellow with get() = XLColor.TangerineYellow |> XlColor
  static member taupe with get() = XLColor.Taupe |> XlColor
  static member taupe_gray with get() = XLColor.TaupeGray |> XlColor
  static member tea_green with get() = XLColor.TeaGreen |> XlColor
  static member stil_de_grain_yellow with get() = XLColor.StilDeGrainYellow |> XlColor
  static member ua_blue with get() = XLColor.UaBlue |> XlColor
  static member ua_red with get() = XLColor.UaRed |> XlColor
  static member ube with get() = XLColor.Ube |> XlColor
  static member violet_ryb with get() = XLColor.VioletRyb |> XlColor
  static member viridian with get() = XLColor.Viridian |> XlColor
  static member vivid_auburn with get() = XLColor.VividAuburn |> XlColor
  static member vivid_burgundy with get() = XLColor.VividBurgundy |> XlColor
  static member vivid_cerise with get() = XLColor.VividCerise |> XlColor
  static member vivid_tangerine with get() = XLColor.VividTangerine |> XlColor
  static member vivid_violet with get() = XLColor.VividViolet |> XlColor
  static member warm_black with get() = XLColor.WarmBlack |> XlColor
  static member wenge with get() = XLColor.Wenge |> XlColor
  static member violet_color_wheel with get() = XLColor.VioletColorWheel |> XlColor
  static member wild_blue_yonder with get() = XLColor.WildBlueYonder |> XlColor
  static member wild_watermelon with get() = XLColor.WildWatermelon |> XlColor
  static member wisteria with get() = XLColor.Wisteria |> XlColor
  static member xanadu with get() = XLColor.Xanadu |> XlColor
  static member yale_blue with get() = XLColor.YaleBlue |> XlColor
  static member yellow_munsell with get() = XLColor.YellowMunsell |> XlColor
  static member yellow_ncs with get() = XLColor.YellowNcs |> XlColor
  static member yellow_process with get() = XLColor.YellowProcess |> XlColor
  static member yellow_ryb with get() = XLColor.YellowRyb |> XlColor
  static member zaffre with get() = XLColor.Zaffre |> XlColor
  static member wild_strawberry with get() = XLColor.WildStrawberry |> XlColor
  static member sienna1 with get() = XLColor.Sienna1 |> XlColor
  static member violet1 with get() = XLColor.Violet1 |> XlColor
  static member vermilion with get() = XLColor.Vermilion |> XlColor
  static member ucla_blue with get() = XLColor.UclaBlue |> XlColor
  static member ucla_gold with get() = XLColor.UclaGold |> XlColor
  static member ufo_green with get() = XLColor.UfoGreen |> XlColor
  static member ultramarine with get() = XLColor.Ultramarine |> XlColor
  static member ultramarine_blue with get() = XLColor.UltramarineBlue |> XlColor
  static member ultra_pink with get() = XLColor.UltraPink |> XlColor
  static member umber with get() = XLColor.Umber |> XlColor
  static member united_nations_blue with get() = XLColor.UnitedNationsBlue |> XlColor
  static member unmellow_yellow with get() = XLColor.UnmellowYellow |> XlColor
  static member veronica with get() = XLColor.Veronica |> XlColor
  static member up_forest_green with get() = XLColor.UpForestGreen |> XlColor
  static member upsdell_red with get() = XLColor.UpsdellRed |> XlColor
  static member urobilin with get() = XLColor.Urobilin |> XlColor
  static member usc_cardinal with get() = XLColor.UscCardinal |> XlColor
  static member usc_gold with get() = XLColor.UscGold |> XlColor
  static member utah_crimson with get() = XLColor.UtahCrimson |> XlColor
  static member vanilla with get() = XLColor.Vanilla |> XlColor
  static member vegas_gold with get() = XLColor.VegasGold |> XlColor
  static member venetian_red with get() = XLColor.VenetianRed |> XlColor
  static member verdigris with get() = XLColor.Verdigris |> XlColor
  static member upMaroon with get() = XLColor.UpMaroon |> XlColor
  static member french_beige with get() = XLColor.FrenchBeige |> XlColor
  static member shocking_pink with get() = XLColor.ShockingPink |> XlColor
  static member shadow with get() = XLColor.Shadow |> XlColor
  static member purple_pizzazz with get() = XLColor.PurplePizzazz |> XlColor
  static member purple_taupe with get() = XLColor.PurpleTaupe |> XlColor
  static member purple_x11 with get() = XLColor.PurpleX11 |> XlColor
  static member radicalRed with get() = XLColor.RadicalRed |> XlColor
  static member raspberry with get() = XLColor.Raspberry |> XlColor
  static member raspberryGlace with get() = XLColor.RaspberryGlace |> XlColor
  static member raspberryPink with get() = XLColor.RaspberryPink |> XlColor
  static member raspberryRose with get() = XLColor.RaspberryRose |> XlColor
  static member rawUmber with get() = XLColor.RawUmber |> XlColor
  static member razzle_dazzle_rose with get() = XLColor.RazzleDazzleRose |> XlColor
  static member razzmatazz with get() = XLColor.Razzmatazz |> XlColor
  static member red_munsell with get() = XLColor.RedMunsell |> XlColor
  static member red_ncs with get() = XLColor.RedNcs |> XlColor
  static member red_pigment with get() = XLColor.RedPigment |> XlColor
  static member red_ryb with get() = XLColor.RedRyb |> XlColor
  static member redwood with get() = XLColor.Redwood |> XlColor
  static member regalia with get() = XLColor.Regalia |> XlColor
  static member rich_black with get() = XLColor.RichBlack |> XlColor
  static member rich_brilliant_lavender with get() = XLColor.RichBrilliantLavender |> XlColor
  static member purple_munsell with get() = XLColor.PurpleMunsell |> XlColor
  static member purple_mountain_majesty with get() = XLColor.PurpleMountainMajesty |> XlColor
  static member purple_heart with get() = XLColor.PurpleHeart |> XlColor
  static member pumpkin with get() = XLColor.Pumpkin |> XlColor
  static member persian_red with get() = XLColor.PersianRed |> XlColor
  static member persian_rose with get() = XLColor.PersianRose |> XlColor
  static member persimmon with get() = XLColor.Persimmon |> XlColor
  static member phlox with get() = XLColor.Phlox |> XlColor
  static member phthalo_blue with get() = XLColor.PhthaloBlue |> XlColor
  static member phthalo_green with get() = XLColor.PhthaloGreen |> XlColor
  static member piggy_pink with get() = XLColor.PiggyPink |> XlColor
  static member pine_green with get() = XLColor.PineGreen |> XlColor
  static member pink_orange with get() = XLColor.PinkOrange |> XlColor
  static member rich_carmine with get() = XLColor.RichCarmine |> XlColor
  static member pink_pearl with get() = XLColor.PinkPearl |> XlColor
  static member pistachio with get() = XLColor.Pistachio |> XlColor
  static member platinum with get() = XLColor.Platinum |> XlColor
  static member plum_traditional with get() = XLColor.PlumTraditional |> XlColor
  static member portland_orange with get() = XLColor.PortlandOrange |> XlColor
  static member princeton_orange with get() = XLColor.PrincetonOrange |> XlColor
  static member prune with get() = XLColor.Prune |> XlColor
  static member prussian_blue with get() = XLColor.PrussianBlue |> XlColor
  static member psychedelic_purple with get() = XLColor.PsychedelicPurple |> XlColor
  static member puce with get() = XLColor.Puce |> XlColor
  static member pink_sherbet with get() = XLColor.PinkSherbet |> XlColor
  static member rich_electric_blue with get() = XLColor.RichElectricBlue |> XlColor
  static member rich_lavender with get() = XLColor.RichLavender |> XlColor
  static member rich_lilac with get() = XLColor.RichLilac |> XlColor
  static member rust with get() = XLColor.Rust |> XlColor
  static member sacramento_state_green with get() = XLColor.SacramentoStateGreen |> XlColor
  static member safety_orange_blaze_orange with get() = XLColor.SafetyOrangeBlazeOrange |> XlColor
  static member saffron with get() = XLColor.Saffron |> XlColor
  static member salmon1 with get() = XLColor.Salmon1 |> XlColor
  static member salmon_pink with get() = XLColor.SalmonPink |> XlColor
  static member sand with get() = XLColor.Sand |> XlColor
  static member sand_dune with get() = XLColor.SandDune |> XlColor
  static member sandstorm with get() = XLColor.Sandstorm |> XlColor
  static member russet with get() = XLColor.Russet |> XlColor
  static member sandy_taupe with get() = XLColor.SandyTaupe |> XlColor
  static member sap_green with get() = XLColor.SapGreen |> XlColor
  static member sapphire with get() = XLColor.Sapphire |> XlColor
  static member satin_sheen_gold with get() = XLColor.SatinSheenGold |> XlColor
  static member scarlet with get() = XLColor.Scarlet |> XlColor
  static member school_bus_yellow with get() = XLColor.SchoolBusYellow |> XlColor
  static member screamin_green with get() = XLColor.ScreaminGreen |> XlColor
  static member seal_brown with get() = XLColor.SealBrown |> XlColor
  static member selective_yellow with get() = XLColor.SelectiveYellow |> XlColor
  static member sepia with get() = XLColor.Sepia |> XlColor
  static member sangria with get() = XLColor.Sangria |> XlColor
  static member shamrock_green with get() = XLColor.ShamrockGreen |> XlColor
  static member rufous with get() = XLColor.Rufous |> XlColor
  static member ruddy_brown with get() = XLColor.RuddyBrown |> XlColor
  static member rich_maroon with get() = XLColor.RichMaroon |> XlColor
  static member rifle_green with get() = XLColor.RifleGreen |> XlColor
  static member robin_egg_blue with get() = XLColor.RobinEggBlue |> XlColor
  static member rose with get() = XLColor.Rose |> XlColor
  static member rose_bonbon with get() = XLColor.RoseBonbon |> XlColor
  static member rose_ebony with get() = XLColor.RoseEbony |> XlColor
  static member rose_gold with get() = XLColor.RoseGold |> XlColor
  static member rose_madder with get() = XLColor.RoseMadder |> XlColor
  static member rose_pink with get() = XLColor.RosePink |> XlColor
  static member ruddy_pink with get() = XLColor.RuddyPink |> XlColor
  static member rose_quartz with get() = XLColor.RoseQuartz |> XlColor
  static member rose_vale with get() = XLColor.RoseVale |> XlColor
  static member rosewood with get() = XLColor.Rosewood |> XlColor
  static member rosso_corsa with get() = XLColor.RossoCorsa |> XlColor
  static member royal_azure with get() = XLColor.RoyalAzure |> XlColor
  static member royal_blue_traditional with get() = XLColor.RoyalBlueTraditional |> XlColor
  static member royal_fuchsia with get() = XLColor.RoyalFuchsia |> XlColor
  static member royal_purple with get() = XLColor.RoyalPurple |> XlColor
  static member ruby with get() = XLColor.Ruby |> XlColor
  static member ruddy with get() = XLColor.Ruddy |> XlColor
  static member rose_Uaupe with get() = XLColor.RoseTaupe |> XlColor
  static member zinnwaldite_brown with get() = XLColor.ZinnwalditeBrown |> XlColor
  static member transparent with get() = XLColor.Transparent |> XlColor
  static member fluorescent_yellow with get() = XLColor.FluorescentYellow |> XlColor
  static member plum with get() = XLColor.Plum |> XlColor
  static member powder_blue with get() = XLColor.PowderBlue |> XlColor
  static member purple with get() = XLColor.Purple |> XlColor
  static member red with get() = XLColor.Red |> XlColor
  static member rosy_brown with get() = XLColor.RosyBrown |> XlColor
  static member royal_blue with get() = XLColor.RoyalBlue |> XlColor
  static member saddle_brown with get() = XLColor.SaddleBrown |> XlColor
  static member salmon with get() = XLColor.Salmon |> XlColor
  static member pink with get() = XLColor.Pink |> XlColor
  static member sandy_brown with get() = XLColor.SandyBrown |> XlColor
  static member sea_shell with get() = XLColor.SeaShell |> XlColor
  static member sienna with get() = XLColor.Sienna |> XlColor
  static member silver with get() = XLColor.Silver |> XlColor
  static member sky_blue with get() = XLColor.SkyBlue |> XlColor
  static member slate_blue with get() = XLColor.SlateBlue |> XlColor
  static member slate_gray with get() = XLColor.SlateGray |> XlColor
  static member snow with get() = XLColor.Snow |> XlColor
  static member spring_green with get() = XLColor.SpringGreen |> XlColor
  static member sea_green with get() = XLColor.SeaGreen |> XlColor
  static member steel_blue with get() = XLColor.SteelBlue |> XlColor
  static member peru with get() = XLColor.Peru |> XlColor
  static member papaya_whip with get() = XLColor.PapayaWhip |> XlColor
  static member medium_turquoise with get() = XLColor.MediumTurquoise |> XlColor
  static member medium_violet_red with get() = XLColor.MediumVioletRed |> XlColor
  static member midnight_blue with get() = XLColor.MidnightBlue |> XlColor
  static member mint_cream with get() = XLColor.MintCream |> XlColor
  static member misty_rose with get() = XLColor.MistyRose |> XlColor
  static member moccasin with get() = XLColor.Moccasin |> XlColor
  static member navajo_white with get() = XLColor.NavajoWhite |> XlColor
  static member navy with get() = XLColor.Navy |> XlColor
  static member peach_puff with get() = XLColor.PeachPuff |> XlColor
  static member old_lace with get() = XLColor.OldLace |> XlColor
  static member olive_drab with get() = XLColor.OliveDrab |> XlColor
  static member orange with get() = XLColor.Orange |> XlColor
  static member orange_red with get() = XLColor.OrangeRed |> XlColor
  static member orchid with get() = XLColor.Orchid |> XlColor
  static member pale_goldenrod with get() = XLColor.PaleGoldenrod |> XlColor
  static member pale_green with get() = XLColor.PaleGreen |> XlColor
  static member pale_turquoise with get() = XLColor.PaleTurquoise |> XlColor
  static member pale_violet_red with get() = XLColor.PaleVioletRed |> XlColor
  static member olive with get() = XLColor.Olive |> XlColor
  static member tan with get() = XLColor.Tan |> XlColor
  static member teal with get() = XLColor.Teal |> XlColor
  static member thistle with get() = XLColor.Thistle |> XlColor
  static member arsenic with get() = XLColor.Arsenic |> XlColor
  static member arylide_yellow with get() = XLColor.ArylideYellow |> XlColor
  static member ash_grey with get() = XLColor.AshGrey |> XlColor
  static member asparagus with get() = XLColor.Asparagus |> XlColor
  static member atomic_tangerine with get() = XLColor.AtomicTangerine |> XlColor
  static member auburn with get() = XLColor.Auburn |> XlColor
  static member aureolin with get() = XLColor.Aureolin |> XlColor
  static member aurometalsaurus with get() = XLColor.Aurometalsaurus |> XlColor
  static member army_green with get() = XLColor.ArmyGreen |> XlColor
  static member awesome with get() = XLColor.Awesome |> XlColor
  static member folly with get() = XLColor.Folly |> XlColor
  static member baby_blue_eyes with get() = XLColor.BabyBlueEyes |> XlColor
  static member baby_pink with get() = XLColor.BabyPink |> XlColor
  static member ball_blue with get() = XLColor.BallBlue |> XlColor
  static member banana_mania with get() = XLColor.BananaMania |> XlColor
  static member battleship_grey with get() = XLColor.BattleshipGrey |> XlColor
  static member bazaar with get() = XLColor.Bazaar |> XlColor
  static member beau_blue with get() = XLColor.BeauBlue |> XlColor
  static member azure_color_wheel with get() = XLColor.AzureColorWheel |> XlColor
  static member aquamarine1 with get() = XLColor.Aquamarine1 |> XlColor
  static member apricot with get() = XLColor.Apricot |> XlColor
  static member apple_green with get() = XLColor.AppleGreen |> XlColor
  static member tomato with get() = XLColor.Tomato |> XlColor
  static member turquoise with get() = XLColor.Turquoise |> XlColor
  static member violet with get() = XLColor.Violet |> XlColor
  static member wheat with get() = XLColor.Wheat |> XlColor
  static member white with get() = XLColor.White |> XlColor
  static member white_smoke with get() = XLColor.WhiteSmoke |> XlColor
  static member yellow with get() = XLColor.Yellow |> XlColor
  static member yellow_green with get() = XLColor.YellowGreen |> XlColor
  static member airForce_blue with get() = XLColor.AirForceBlue |> XlColor
  static member alizarin with get() = XLColor.Alizarin |> XlColor
  static member almond with get() = XLColor.Almond |> XlColor
  static member amaranth with get() = XLColor.Amaranth |> XlColor
  static member amber with get() = XLColor.Amber |> XlColor
  static member amber_sae_ece with get() = XLColor.AmberSaeEce |> XlColor
  static member american_rose with get() = XLColor.AmericanRose |> XlColor
  static member amethyst with get() = XLColor.Amethyst |> XlColor
  static member anti_flash_white with get() = XLColor.AntiFlashWhite |> XlColor
  static member antique_brass with get() = XLColor.AntiqueBrass |> XlColor
  static member antique_fuchsia with get() = XLColor.AntiqueFuchsia |> XlColor
  static member medium_spring_green with get() = XLColor.MediumSpringGreen |> XlColor
  static member medium_slate_blue with get() = XLColor.MediumSlateBlue |> XlColor
  static member medium_sea_green with get() = XLColor.MediumSeaGreen |> XlColor
  static member medium_purple with get() = XLColor.MediumPurple |> XlColor
  static member dark_blue with get() = XLColor.DarkBlue |> XlColor
  static member dark_cyan with get() = XLColor.DarkCyan |> XlColor
  static member dark_goldenrod with get() = XLColor.DarkGoldenrod |> XlColor
  static member dark_gray with get() = XLColor.DarkGray |> XlColor
  static member dark_green with get() = XLColor.DarkGreen |> XlColor
  static member dark_khaki with get() = XLColor.DarkKhaki |> XlColor
  static member dark_magenta with get() = XLColor.DarkMagenta |> XlColor
  static member dark_olive_green with get() = XLColor.DarkOliveGreen |> XlColor
  static member cyan with get() = XLColor.Cyan |> XlColor
  static member dark_orange with get() = XLColor.DarkOrange |> XlColor
  static member dark_red with get() = XLColor.DarkRed |> XlColor
  static member dark_salmon with get() = XLColor.DarkSalmon |> XlColor
  static member dark_sea_green with get() = XLColor.DarkSeaGreen |> XlColor
  static member dark_slate_blue with get() = XLColor.DarkSlateBlue |> XlColor
  static member dark_slate_gray with get() = XLColor.DarkSlateGray |> XlColor
  static member dark_turquoise with get() = XLColor.DarkTurquoise |> XlColor
  static member dark_violet with get() = XLColor.DarkViolet |> XlColor
  static member deep_pink with get() = XLColor.DeepPink |> XlColor
  static member dark_orchid with get() = XLColor.DarkOrchid |> XlColor
  static member crimson with get() = XLColor.Crimson |> XlColor
  static member cornsilk with get() = XLColor.Cornsilk |> XlColor
  static member cornflower_blue with get() = XLColor.CornflowerBlue |> XlColor
  static member no_color with get() = XLColor.NoColor |> XlColor
  static member alice_blue with get() = XLColor.AliceBlue |> XlColor
  static member antique_white with get() = XLColor.AntiqueWhite |> XlColor
  static member aqua with get() = XLColor.Aqua |> XlColor
  static member aquamarine with get() = XLColor.Aquamarine |> XlColor
  static member azure with get() = XLColor.Azure |> XlColor
  static member beige with get() = XLColor.Beige |> XlColor
  static member bisque with get() = XLColor.Bisque |> XlColor
  static member black with get() = XLColor.Black |> XlColor
  static member blanched_almond with get() = XLColor.BlanchedAlmond |> XlColor
  static member blue with get() = XLColor.Blue |> XlColor
  static member blue_violet with get() = XLColor.BlueViolet |> XlColor
  static member brown with get() = XLColor.Brown |> XlColor
  static member burly_wood with get() = XLColor.BurlyWood |> XlColor
  static member cadet_blue with get() = XLColor.CadetBlue |> XlColor
  static member chartreuse with get() = XLColor.Chartreuse |> XlColor
  static member chocolate with get() = XLColor.Chocolate |> XlColor
  static member coral with get() = XLColor.Coral |> XlColor
  static member deep_sky_blue with get() = XLColor.DeepSkyBlue |> XlColor
  static member beaver with get() = XLColor.Beaver |> XlColor
  static member dim_gray with get() = XLColor.DimGray |> XlColor
  static member firebrick with get() = XLColor.Firebrick |> XlColor
  static member light_goldenrod_yellow with get() = XLColor.LightGoldenrodYellow |> XlColor
  static member light_gray with get() = XLColor.LightGray |> XlColor
  static member light_green with get() = XLColor.LightGreen |> XlColor
  static member light_pink with get() = XLColor.LightPink |> XlColor
  static member light_salmon with get() = XLColor.LightSalmon |> XlColor
  static member light_sea_green with get() = XLColor.LightSeaGreen |> XlColor
  static member light_sky_blue with get() = XLColor.LightSkyBlue |> XlColor
  static member light_slate_gray with get() = XLColor.LightSlateGray |> XlColor
  static member light_cyan with get() = XLColor.LightCyan |> XlColor
  static member light_steel_blue with get() = XLColor.LightSteelBlue |> XlColor
  static member lime with get() = XLColor.Lime |> XlColor
  static member lime_green with get() = XLColor.LimeGreen |> XlColor
  static member linen with get() = XLColor.Linen |> XlColor
  static member magenta with get() = XLColor.Magenta |> XlColor
  static member maroon with get() = XLColor.Maroon |> XlColor
  static member medium_aquamarine with get() = XLColor.MediumAquamarine |> XlColor
  static member medium_blue with get() = XLColor.MediumBlue |> XlColor
  static member medium_orchid with get() = XLColor.MediumOrchid |> XlColor
  static member light_yellow with get() = XLColor.LightYellow |> XlColor
  static member light_coral with get() = XLColor.LightCoral |> XlColor
  static member light_blue with get() = XLColor.LightBlue |> XlColor
  static member lemon_chiffon with get() = XLColor.LemonChiffon |> XlColor
  static member floral_white with get() = XLColor.FloralWhite |> XlColor
  static member forest_green with get() = XLColor.ForestGreen |> XlColor
  static member fuchsia with get() = XLColor.Fuchsia |> XlColor
  static member gainsboro with get() = XLColor.Gainsboro |> XlColor
  static member ghost_white with get() = XLColor.GhostWhite |> XlColor
  static member gold with get() = XLColor.Gold |> XlColor
  static member goldenrod with get() = XLColor.Goldenrod |> XlColor
  static member gray with get() = XLColor.Gray |> XlColor
  static member green with get() = XLColor.Green |> XlColor
  static member green_yellow with get() = XLColor.GreenYellow |> XlColor
  static member honeydew with get() = XLColor.Honeydew |> XlColor
  static member hot_pink with get() = XLColor.HotPink |> XlColor
  static member indian_red with get() = XLColor.IndianRed |> XlColor
  static member indigo with get() = XLColor.Indigo |> XlColor
  static member ivory with get() = XLColor.Ivory |> XlColor
  static member khaki with get() = XLColor.Khaki |> XlColor
  static member lavender with get() = XLColor.Lavender |> XlColor
  static member lavender_blush with get() = XLColor.LavenderBlush |> XlColor
  static member lawn_green with get() = XLColor.LawnGreen |> XlColor
  static member dodger_blue with get() = XLColor.DodgerBlue |> XlColor
  static member bistre with get() = XLColor.Bistre |> XlColor
  static member baby_blue with get() = XLColor.BabyBlue |> XlColor
  static member bleu_de_france with get() = XLColor.BleuDeFrance |> XlColor
  static member dark_pastel_red with get() = XLColor.DarkPastelRed |> XlColor
  static member dark_pink with get() = XLColor.DarkPink |> XlColor
  static member dark_powder_blue with get() = XLColor.DarkPowderBlue |> XlColor
  static member dark_raspberry with get() = XLColor.DarkRaspberry |> XlColor
  static member dark_scarlet with get() = XLColor.DarkScarlet |> XlColor
  static member dark_sienna with get() = XLColor.DarkSienna |> XlColor
  static member dark_spring_green with get() = XLColor.DarkSpringGreen |> XlColor
  static member dark_tan with get() = XLColor.DarkTan |> XlColor
  static member dark_pastel_purple with get() = XLColor.DarkPastelPurple |> XlColor
  static member dark_tangerine with get() = XLColor.DarkTangerine |> XlColor
  static member dark_terra_cotta with get() = XLColor.DarkTerraCotta |> XlColor
  static member dartmouth_green with get() = XLColor.DartmouthGreen |> XlColor
  static member davys_grey with get() = XLColor.DavysGrey |> XlColor
  static member debian_red with get() = XLColor.DebianRed |> XlColor
  static member deep_carmine with get() = XLColor.DeepCarmine |> XlColor
  static member deep_carmine_pink with get() = XLColor.DeepCarminePink |> XlColor
  static member deep_carrot_orange with get() = XLColor.DeepCarrotOrange |> XlColor
  static member deep_cerise with get() = XLColor.DeepCerise |> XlColor
  static member dark_taupe with get() = XLColor.DarkTaupe |> XlColor
  static member dark_pastel_green with get() = XLColor.DarkPastelGreen |> XlColor
  static member dark_pastel_blue with get() = XLColor.DarkPastelBlue |> XlColor
  static member dark_midnight_blue with get() = XLColor.DarkMidnightBlue |> XlColor
  static member cosmic_latte with get() = XLColor.CosmicLatte |> XlColor
  static member cotton_candy with get() = XLColor.CottonCandy |> XlColor
  static member cream with get() = XLColor.Cream |> XlColor
  static member crimson_glory with get() = XLColor.CrimsonGlory |> XlColor
  static member cyan_process with get() = XLColor.CyanProcess |> XlColor
  static member daffodil with get() = XLColor.Daffodil |> XlColor
  static member dandelion with get() = XLColor.Dandelion |> XlColor
  static member dark_brown with get() = XLColor.DarkBrown |> XlColor
  static member bittersweet with get() = XLColor.Bittersweet |> XlColor
  static member dark_candy_apple_red with get() = XLColor.DarkCandyAppleRed |> XlColor
  static member dark_cerulean with get() = XLColor.DarkCerulean |> XlColor
  static member dark_champagne with get() = XLColor.DarkChampagne |> XlColor
  static member dark_chestnut with get() = XLColor.DarkChestnut |> XlColor
  static member dark_coral with get() = XLColor.DarkCoral |> XlColor
  static member dark_electric_blue with get() = XLColor.DarkElectricBlue |> XlColor
  static member dark_green1 with get() = XLColor.DarkGreen1 |> XlColor
  static member dark_jungle_green with get() = XLColor.DarkJungleGreen |> XlColor
  static member dark_lava with get() = XLColor.DarkLava |> XlColor
  static member dark_lavender with get() = XLColor.DarkLavender |> XlColor
  static member deep_champagne with get() = XLColor.DeepChampagne |> XlColor
  static member cornell_red with get() = XLColor.CornellRed |> XlColor
  static member deep_chestnut with get() = XLColor.DeepChestnut |> XlColor
  static member deep_jungle_green with get() = XLColor.DeepJungleGreen |> XlColor
  static member electric_violet with get() = XLColor.ElectricViolet |> XlColor
  static member emerald with get() = XLColor.Emerald |> XlColor
  static member eton_blue with get() = XLColor.EtonBlue |> XlColor
  static member fallow with get() = XLColor.Fallow |> XlColor
  static member falu_red with get() = XLColor.FaluRed |> XlColor
  static member fandango with get() = XLColor.Fandango |> XlColor
  static member fashion_fuchsia with get() = XLColor.FashionFuchsia |> XlColor
  static member fawn with get() = XLColor.Fawn |> XlColor
  static member electric_ultramarine with get() = XLColor.ElectricUltramarine |> XlColor
  static member feldgrau with get() = XLColor.Feldgrau |> XlColor
  static member ferrari_red with get() = XLColor.FerrariRed |> XlColor
  static member field_drab with get() = XLColor.FieldDrab |> XlColor
  static member fire_engine_red with get() = XLColor.FireEngineRed |> XlColor
  static member flame with get() = XLColor.Flame |> XlColor
  static member flamingo_pink with get() = XLColor.FlamingoPink |> XlColor
  static member flavescent with get() = XLColor.Flavescent |> XlColor
  static member flax with get() = XLColor.Flax |> XlColor
  static member fluorescent_orange with get() = XLColor.FluorescentOrange |> XlColor
  static member fern_green with get() = XLColor.FernGreen |> XlColor
  static member electric_purple with get() = XLColor.ElectricPurple |> XlColor
  static member electric_lime with get() = XLColor.ElectricLime |> XlColor
  static member electric_lavender with get() = XLColor.ElectricLavender |> XlColor
  static member deep_lilac with get() = XLColor.DeepLilac |> XlColor
  static member deep_magenta with get() = XLColor.DeepMagenta |> XlColor
  static member deep_peach with get() = XLColor.DeepPeach |> XlColor
  static member deep_saffron with get() = XLColor.DeepSaffron |> XlColor
  static member denim with get() = XLColor.Denim |> XlColor
  static member desert with get() = XLColor.Desert |> XlColor
  static member desert_sand with get() = XLColor.DesertSand |> XlColor
  static member dogwood_rose with get() = XLColor.DogwoodRose |> XlColor
  static member dollar_bill with get() = XLColor.DollarBill |> XlColor
  static member drab with get() = XLColor.Drab |> XlColor
  static member duke_blue with get() = XLColor.DukeBlue |> XlColor
  static member earth_yellow with get() = XLColor.EarthYellow |> XlColor
  static member ecru with get() = XLColor.Ecru |> XlColor
  static member eggplant with get() = XLColor.Eggplant |> XlColor
  static member eggshell with get() = XLColor.Eggshell |> XlColor
  static member egyptian_blue with get() = XLColor.EgyptianBlue |> XlColor
  static member electric_blue with get() = XLColor.ElectricBlue |> XlColor
  static member electric_crimson with get() = XLColor.ElectricCrimson |> XlColor
  static member electric_indigo with get() = XLColor.ElectricIndigo |> XlColor
  static member deep_fuchsia with get() = XLColor.DeepFuchsia |> XlColor
  static member corn with get() = XLColor.Corn |> XlColor
  static member dark_byzantium with get() = XLColor.DarkByzantium |> XlColor
  static member coral_red with get() = XLColor.CoralRed |> XlColor
  static member brink_pink with get() = XLColor.BrinkPink |> XlColor
  static member british_racing_green with get() = XLColor.BritishRacingGreen |> XlColor
  static member bronze with get() = XLColor.Bronze |> XlColor
  static member brown_traditional with get() = XLColor.BrownTraditional |> XlColor
  static member bubble_gum with get() = XLColor.BubbleGum |> XlColor
  static member bubbles with get() = XLColor.Bubbles |> XlColor
  static member buff with get() = XLColor.Buff |> XlColor
  static member bulgarian_rose with get() = XLColor.BulgarianRose |> XlColor
  static member brilliant_rose with get() = XLColor.BrilliantRose |> XlColor
  static member burgundy with get() = XLColor.Burgundy |> XlColor
  static member burnt_sienna with get() = XLColor.BurntSienna |> XlColor
  static member burnt_umber with get() = XLColor.BurntUmber |> XlColor
  static member byzantine with get() = XLColor.Byzantine |> XlColor
  static member byzantium with get() = XLColor.Byzantium |> XlColor
  static member cadet with get() = XLColor.Cadet |> XlColor
  static member cadet_grey with get() = XLColor.CadetGrey |> XlColor
  static member cadmium_green with get() = XLColor.CadmiumGreen |> XlColor
  static member cadmium_orange with get() = XLColor.CadmiumOrange |> XlColor
  static member cordovan with get() = XLColor.Cordovan |> XlColor
  static member brilliant_lavender with get() = XLColor.BrilliantLavender |> XlColor
  static member bright_ube with get() = XLColor.BrightUbe |> XlColor
  static member bright_turquoise with get() = XLColor.BrightTurquoise |> XlColor
  static member blizzard_blue with get() = XLColor.BlizzardBlue |> XlColor
  static member blond with get() = XLColor.Blond |> XlColor
  static member blue_bell with get() = XLColor.BlueBell |> XlColor
  static member blue_gray with get() = XLColor.BlueGray |> XlColor
  static member blue_green with get() = XLColor.BlueGreen |> XlColor
  static member blue_pigment with get() = XLColor.BluePigment |> XlColor
  static member blue_ryb with get() = XLColor.BlueRyb |> XlColor
  static member blush with get() = XLColor.Blush |> XlColor
  static member bole with get() = XLColor.Bole |> XlColor
  static member bondi_blue with get() = XLColor.BondiBlue |> XlColor
  static member boston_university_red with get() = XLColor.BostonUniversityRed |> XlColor
  static member brandeis_blue with get() = XLColor.BrandeisBlue |> XlColor
  static member brass with get() = XLColor.Brass |> XlColor
  static member brick_red with get() = XLColor.BrickRed |> XlColor
  static member bright_cerulean with get() = XLColor.BrightCerulean |> XlColor
  static member bright_green with get() = XLColor.BrightGreen |> XlColor
  static member bright_lavender with get() = XLColor.BrightLavender |> XlColor
  static member bright_maroon with get() = XLColor.BrightMaroon |> XlColor
  static member bright_pink with get() = XLColor.BrightPink |> XlColor
  static member cadmium_red with get() = XLColor.CadmiumRed |> XlColor
  static member cadmium_yellow with get() = XLColor.CadmiumYellow |> XlColor
  static member burnt_orange with get() = XLColor.BurntOrange |> XlColor
  static member cambridge_blue with get() = XLColor.CambridgeBlue |> XlColor
  static member champagne with get() = XLColor.Champagne |> XlColor
  static member charcoal with get() = XLColor.Charcoal |> XlColor
  static member chartreuse_traditional with get() = XLColor.ChartreuseTraditional |> XlColor
  static member cherry_blossom_pink with get() = XLColor.CherryBlossomPink |> XlColor
  static member chocolate1 with get() = XLColor.Chocolate1 |> XlColor
  static member chrome_yellow with get() = XLColor.ChromeYellow |> XlColor
  static member cinereous with get() = XLColor.Cinereous |> XlColor
  static member cal_poly_pomona_green with get() = XLColor.CalPolyPomonaGreen |> XlColor
  static member chamoisee with get() = XLColor.Chamoisee |> XlColor
  static member citrine with get() = XLColor.Citrine |> XlColor
  static member cobalt with get() = XLColor.Cobalt |> XlColor
  static member columbia_blue with get() = XLColor.ColumbiaBlue |> XlColor
  static member cool_black with get() = XLColor.CoolBlack |> XlColor
  static member cool_grey with get() = XLColor.CoolGrey |> XlColor
  static member copper with get() = XLColor.Copper |> XlColor
  static member copper_rose with get() = XLColor.CopperRose |> XlColor
  static member coquelicot with get() = XLColor.Coquelicot |> XlColor
  static member coral_pink with get() = XLColor.CoralPink |> XlColor
  static member classic_rose with get() = XLColor.ClassicRose |> XlColor
  static member cerulean_blue with get() = XLColor.CeruleanBlue |> XlColor
  static member cinnabar with get() = XLColor.Cinnabar |> XlColor
  static member cerise_pink with get() = XLColor.CerisePink |> XlColor
  static member cerulean with get() = XLColor.Cerulean |> XlColor
  static member camel with get() = XLColor.Camel |> XlColor
  static member camouflage_green with get() = XLColor.CamouflageGreen |> XlColor
  static member canary_yellow with get() = XLColor.CanaryYellow |> XlColor
  static member candy_apple_red with get() = XLColor.CandyAppleRed |> XlColor
  static member candy_pink with get() = XLColor.CandyPink |> XlColor
  static member caput_mortuum with get() = XLColor.CaputMortuum |> XlColor
  static member cardinal with get() = XLColor.Cardinal |> XlColor
  static member caribbean_green with get() = XLColor.CaribbeanGreen |> XlColor
  static member carmine with get() = XLColor.Carmine |> XlColor
  static member carmine_red with get() = XLColor.CarmineRed |> XlColor
  static member carnation_pink with get() = XLColor.CarnationPink |> XlColor
  static member carnelian with get() = XLColor.Carnelian |> XlColor
  static member carolina_blue with get() = XLColor.CarolinaBlue |> XlColor
  static member carrot_orange with get() = XLColor.CarrotOrange |> XlColor
  static member ceil with get() = XLColor.Ceil |> XlColor
  static member celadon with get() = XLColor.Celadon |> XlColor
  static member celestial_blue with get() = XLColor.CelestialBlue |> XlColor
  static member cerise with get() = XLColor.Cerise |> XlColor
  static member carmine_pink with get() = XLColor.CarminePink |> XlColor


  
and XlNumberFormat internal (format: IXLNumberFormat) =
  member internal __.raw with get() = format
  member __.format with get() = format.Format
  member __.set_format(format': string) = format.SetFormat(format') |> ignore
  


and XlFill internal (fill: IXLFill) =
  member internal __.raw with get() = fill
  member __.background_color with get() = fill.BackgroundColor |> XlColor
  member __.pattern_color with get() = fill.PatternColor |> XlColor
  member __.pattern_type with get() : FillPatternValues = fill.PatternType

  member __.set_background_color(color: XlColor) = fill.SetBackgroundColor(color.raw)
  member __.set_pattern_color(color: XlColor) = fill.SetPatternColor(color.raw)
  member __.set_pattern_type(pattern: FillPatternValues) = fill.SetPatternType(pattern)



and XlFont internal (font: IXLFont) =
  member internal __.raw with get() = font
  member __.bold with get() = font.Bold
  member __.italic with get() = font.Italic
  member __.strikethrough with get() = font.Strikethrough
  member __.shadow with get() = font.Shadow
  member __.size with get() = font.FontSize
  member __.name with get() = font.FontName
  member __.color with get() = font.FontColor |> XlColor
  member __.underline with get() : FontUnderlineValues = font.Underline
  member __.vertical_alignment with get() : FontVerticalTextAlignmentValues = font.VerticalAlignment
  member __.familiy_numbering with get() : FontFamilyNumberingValues = font.FontFamilyNumbering
  member __.charset with get() : FontCharSet = font.FontCharSet

  member __.set_bold(?value: bool) = 
    match value with Some v -> font.SetBold(v) | None -> font.SetBold()
    |> ignore
  member __.set_italic(?value: bool) =
    match value with Some v -> font.SetItalic(v) | None -> font.SetItalic()
    |> ignore
  member __.set_strikethrough(?value: bool) =
    match value with Some v -> font.SetStrikethrough(v) | None -> font.SetStrikethrough()
    |> ignore
  member __.set_shadow(?value: bool) =
    match value with Some v -> font.SetShadow(v) | None -> font.SetShadow()
    |> ignore
  member __.set_siza(value: float) = font.SetFontSize(value) |> ignore
  member __.set_name(value: string) = font.SetFontName(value) |> ignore
  member __.set_color(value: XlColor) = font.SetFontColor(value.raw) |> ignore
  member __.set_underline(?value: FontUnderlineValues) =
    match value with Some v -> font.SetUnderline(v) | None -> font.SetUnderline()
    |> ignore
  member __.set_vertical_alignment(value: FontVerticalTextAlignmentValues) = font.SetVerticalAlignment(value) |> ignore
  member __.set_familiy_numbering(value: FontFamilyNumberingValues) = font.SetFontFamilyNumbering(value) |> ignore
  member __.set_charset(value: FontCharSet) = font.SetFontCharSet(value) |> ignore



and XlProtection internal (protection: IXLProtection) =
  member internal __.raw with get() = protection
  member __.locked with get() = protection.Locked
  member __.hidden with get() = protection.Hidden

  member __.set_locked(value: bool) = protection.SetLocked(value) |> ignore
  member __.set_hidden(value: bool) = protection.SetHidden(value) |> ignore



// TODO
and XlAutoFilter internal (filter: IXLAutoFilter) =
  member internal __.raw with get() = filter
