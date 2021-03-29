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
       | t when t = typeof<int16> -> to_int16(cell.GetDouble()) |> unbox<'T>
       | t when t = typeof<int> -> to_int(cell.GetDouble()) |> unbox<'T>
       | t when t = typeof<int64> -> to_int64(cell.GetDouble()) |> unbox<'T>
       | t when t = typeof<float32> -> to_single(cell.GetDouble()) |> unbox<'T>
       | t when t = typeof<float> -> to_double(cell.GetDouble()) |> unbox<'T>
       | t when t = typeof<decimal> -> to_decimal(cell.GetDouble()) |> unbox<'T>
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
  // TODO
  member __.style with get() = cell.Style

  member __.column with get() = XlColumn(cell.WorksheetColumn())
  member __.row with get() = XlRow(cell.WorksheetRow())

  member __.delete(option: ShiftDeleted) = cell.Delete(option)
  member __.clear(?options: ClearOption) = match options with Some opt -> cell.Clear(opt) | None -> cell.Clear()
  member __.left() = XlCell(cell.CellLeft())
  member __.left(step: int) = XlCell(cell.CellLeft(step))
  member __.right() = XlCell(cell.CellRight())
  member __.right(step: int) = XlCell(cell.CellRight(step))
  member __.above() = XlCell(cell.CellAbove())
  member __.above(step: int) = XlCell(cell.CellAbove(step))
  member __.below() = XlCell(cell.CellBelow())
  member __.below(step: int) = XlCell(cell.CellBelow(step))

  member __.copy_from(other_cell: IXLCell) = XlCell(cell.CopyFrom(other_cell))
  member __.copy_from(other_cell: XlCell) = XlCell(cell.CopyFrom(other_cell.raw))
  member __.copy_from(other_cell: string) = XlCell(cell.CopyFrom(other_cell))
  member __.copy_to(target: IXLCell) = XlCell(cell.CopyTo(target))
  member __.copy_to(target: XlCell) = XlCell(cell.CopyTo(target.raw))
  member __.copy_to(target: string) = XlCell(cell.CopyTo(target))

  member __.insert_cells_above(number_of_rows: int) = XlCells(cell.InsertCellsAbove number_of_rows)
  member __.insert_cells_after(number_of_columns: int) = XlCells(cell.InsertCellsAfter number_of_columns)
  member __.insert_cells_before(number_of_columns: int) = XlCells(cell.InsertCellsBefore number_of_columns)
  member __.insert_cells_below(number_of_rows: int) = XlCells(cell.InsertCellsBelow number_of_rows)

  member __.insert_table(data: DataTable) = XlTable(cell.InsertTable(data))
  member __.insert_table(data: DataTable, create_table: bool) = XlTable(cell.InsertTable(data, create_table))
  member __.insert_table(data: DataTable, table_name: string) = XlTable(cell.InsertTable(data, table_name))
  member __.insert_table(data: DataTable, table_name: string, create_table: bool) = XlTable(cell.InsertTable(data, table_name, create_table))



and XlCells internal (cells: IXLCells) =
  member internal __.raw with get() = cells
  member __.value with set(value) = cells.Value <- value
  member __.get() = cells |> Seq.map(fun cell -> XlCell(cell))
  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = cells.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = cells.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = cells.FormulaR1C1 <- value
  // TODO
  member __.style with get() = cells.Style
  
  member __.clear(?options: ClearOption) = match options with Some opt -> cells.Clear(opt) | None -> cells.Clear()
  member __.delete_comments() = cells.DeleteComments()
  member __.delete_sparklines() = cells.DeleteSparklines()

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = cells |> Seq.map(fun cell -> XlCell(cell))
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = cells |> Seq.map(fun cell -> XlCell(cell))
      cells.GetEnumerator()


    
and XlRow internal (row: IXLRow) =
  member internal __.raw with get() = row
  member __.value with set(value) = row.Value <- value
  member __.cell_count with get() = row.CellCount
  member __.worksheet with get() = row.Worksheet
  member __.height with get() = row.Height
  member __.first_cell with get() = XlCell(row.FirstCell())
  member __.first_cell_used with get() = XlCell(row.FirstCellUsed())
  member __.last_cell with get() = XlCell(row.LastCell())
  member __.last_cell_used with get() = XlCell(row.LastCellUsed())
  member __.row_number with get() = row.RowNumber()
  // TODO
  member __.style with get() = row.Style

  member __.set(value) = __.value <- box value
  member __.fx(value: obj) = row.FormulaA1 <- value.ToString()
  member __.set_formula(value: string) = row.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = row.FormulaR1C1 <- value
  member __.cell(row': int) = XlCell(row.Cell row')
  member __.cells() = XlCells(row.Cells())
  member __.cells(row_range: string) = XlCells(row.Cells row_range)
  member __.cells(row': int) = __.cells(row'.ToString())
  member __.cells(from': int, to': int) = __.cells($"%d{from'}:%d{to'}")

  member __.above() = XlRow(row.RowAbove())
  member __.above(step: int) = XlRow(row.RowAbove(step))
  member __.below() = XlRow(row.RowBelow())
  member __.below(step: int) = XlRow(row.RowBelow(step))
  
  member __.adjust() = XlRow(row.AdjustToContents())
  member __.adjust(start_column: int) = XlRow(row.AdjustToContents(start_column))
  member __.adjust(start_column: int, end_column: int) = XlRow(row.AdjustToContents(start_column, end_column))
  member __.adjust(min_height: float, max_height: float) = XlRow(row.AdjustToContents(min_height, max_height))
  member __.adjust(start_column: int, min_height: float, max_height: float) = XlRow(row.AdjustToContents(start_column, min_height, max_height))
  member __.adjust(start_column: int, end_column: int, min_height: float, max_height: float) = XlRow(row.AdjustToContents(start_column, end_column, min_height, max_height))

  member __.delete() = row.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> row.Clear(opt) | None -> row.Clear()
  member __.hide() = row.Hide()
  member __.unhide() = row.Unhide()

  member __.group() = XlRow(row.Group())
  member __.group(outline_level: int) = XlRow(row.Group(outline_level))
  member __.group(collapse: bool) = XlRow(row.Group(collapse))
  member __.group(outline_level: int, collapse: bool) = XlRow(row.Group(outline_level, collapse))
  member __.ungroup() = XlRow(row.Ungroup())
  member __.ungroup(from_all: bool) = XlRow(row.Ungroup(from_all))
  member __.expand() = XlRow(row.Expand())
  member __.collapse() = XlRow(row.Collapse())

  member __.add_horizontal_pagebreak() = XlRow(row.AddHorizontalPageBreak())
  member __.insert_above(number_of_rows: int) = XlRows(row.InsertRowsAbove(number_of_rows))
  member __.insert_below(number_of_rows: int) = XlRows(row.InsertRowsBelow(number_of_rows))

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = row.Cells() |> Seq.map(fun cell -> XlCell(cell))
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = row.Cells() |> Seq.map(fun cell -> XlCell(cell))
      cells.GetEnumerator()
    

    
and XlRows internal (rows: IXLRows) =
  member internal __.raw with get() = rows
  member __.cells with get() = XlCells(rows.Cells())
  member __.used_cells with get() = XlCells(rows.CellsUsed())
  // TODO
  member __.style with get() = rows.Style

  member __.adjust() = XlRows(rows.AdjustToContents())
  member __.adjust(start_column: int) = XlRows(rows.AdjustToContents(start_column))
  member __.adjust(start_column: int, end_column: int) = XlRows(rows.AdjustToContents(start_column, end_column))
  member __.adjust(min_height: float, max_height: float) = XlRows(rows.AdjustToContents(min_height, max_height))
  member __.adjust(start_column: int, min_height: float, max_height: float) = XlRows(rows.AdjustToContents(start_column, min_height, max_height))
  member __.adjust(start_column: int, end_column: int, min_height: float, max_height: float) = XlRows(rows.AdjustToContents(start_column, end_column, min_height, max_height))

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
  
  member __.add_horizontal_pagebreak() = XlRows(rows.AddHorizontalPageBreaks())

  interface IEnumerable<XlRow> with
    member __.GetEnumerator(): IEnumerator = 
      let rs = rows |> Seq.map(fun row -> XlRow(row))
      (rs :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlRow> =
      let rs = rows |> Seq.map(fun row -> XlRow(row))
      rs.GetEnumerator()
      


and XlRangeRow internal (range: IXLRangeRow) =
  member internal __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.first_cell with get() = range.FirstCell() |> XlCell
  member __.first_cell_used with get() = range.FirstCellUsed() |> XlCell
  member __.last_cell with get() = range.LastCell() |> XlCell
  member __.last_cell_used with get() = range.LastCellUsed() |> XlCell
  member __.row_number with get() = range.RowNumber()
  member __.row_span with get() = range.RangeAddress.RowSpan
  member __.column_span with get() = range.RangeAddress.ColumnSpan
  // TODO
  member __.style with get() = range.Style

  member __.fx(value: obj) = range.FormulaA1 <- value.ToString()
  member __.set(value) = __.value <- box value
  member __.set_formula(value: string) = range.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = range.FormulaR1C1 <- value
  member __.cell(column_number: int) = range.Cell(column_number) |> XlCell
  member __.cell(column_number: string) = range.Cell(column_number) |> XlCell
  member __.cells(first_column: int, last_column:  int) = range.Cells $"%d{first_column}:%d{last_column}" |> XlCells
  member __.cells(cells_in_row: string) = range.Cells cells_in_row |> XlCells

  member __.row(start': int, end': int) = range.Row(start', end') |> XlRangeRow
  member __.row(start': IXLCell, end': IXLCell) = range.Row(start', end') |> XlRangeRow
  member __.row(start': XlCell, end': XlCell) = range.Row(start'.raw, end'.raw) |> XlRangeRow
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
  
  member __.copy_to(target: IXLCell) = range.CopyTo(target) |> XlRangeRow
  member __.copy_to(target: XlRangeRow) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeRow
  member __.copy_to(target: XlRange) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeRow
  member __.copy_to(target: XlColumn) = range.CopyTo(target.raw :> IXLRangeBase) |> XlRangeRow

  member __.delete(?option: ShiftDeleted) = match option with Some opt -> range.Delete(opt) | None -> range.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> range.Clear(opt) | None -> range.Clear()
  


and XlRangeRows internal (range: IXLRangeRows) =
  member internal __.raw with get() = range
  member __.cells with get() = range.Cells() |> XlCells
  // TODO
  member __.style with get() = range.Style

  member __.delete() = range.Delete()
  member __.clear(?options: ClearOption) = match options with Some opt -> range.Clear(opt) | None -> range.Clear()
  
  interface IEnumerable<XlRangeRow> with
    member __.GetEnumerator(): IEnumerator = 
      let rs = range |> Seq.map(fun row -> XlRangeRow(row))
      (rs :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlRangeRow> =
      let rs = range |> Seq.map(fun row -> XlRangeRow(row))
      rs.GetEnumerator()
  


and XlColumn internal (column: IXLColumn) =
  member internal __.raw with get() = column
  member __.value with set(value) = column.Value <- value
  member __.cell_count with get() = column.CellCount
  member __.worksheet with get() = column.Worksheet
  member __.width with get() = column.Width
  member __.first_cell with get() = XlCell(column.FirstCell())
  member __.first_cell_used with get() = XlCell(column.FirstCellUsed())
  member __.last_cell with get() = XlCell(column.LastCell())
  member __.last_cell_used with get() = XlCell(column.LastCellUsed())
  // TODO
  member __.style with get() = column.Style

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
  
  member __.adjust() = XlColumn(column.AdjustToContents())
  member __.adjust(start_row: int) = XlColumn(column.AdjustToContents(start_row))
  member __.adjust(start_row: int, end_row: int) = XlColumn(column.AdjustToContents(start_row, end_row))
  member __.adjust(min_width: float, max_width: float) = XlColumn(column.AdjustToContents(min_width, max_width))
  member __.adjust(start_row: int, min_width: float, max_width: float) = XlColumn(column.AdjustToContents(start_row, min_width, max_width))
  member __.adjust(start_row: int, end_row: int, min_width: float, max_width: float) = XlColumn(column.AdjustToContents(start_row, end_row, min_width, max_width))
  
  member __.clear(?options: ClearOption) = match options with Some opt -> column.Clear(opt) | None -> column.Clear()
  member __.hide() = column.Hide()
  member __.unhide() = column.Unhide()

  member __.group() = XlColumn(column.Group())
  member __.group(outline_level: int) = XlColumn(column.Group(outline_level))
  member __.group(collapse: bool) = XlColumn(column.Group(collapse))
  member __.group(outline_level: int, collapse: bool) = XlColumn(column.Group(outline_level, collapse))
  member __.ungroup() = XlColumn(column.Ungroup())
  member __.ungroup(from_all: bool) = XlColumn(column.Ungroup(from_all))
  member __.expand() = XlColumn(column.Expand())
  member __.collapse() = XlColumn(column.Collapse())

  member __.add_vertical_pagebreak() = XlColumn(column.AddVerticalPageBreak())
  member __.insert_after(number_of_columns: int) = XlColumns(column.InsertColumnsAfter(number_of_columns))
  member __.insert_before(number_of_columns: int) = XlColumns(column.InsertColumnsBefore(number_of_columns))

  interface IEnumerable<XlCell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = column.Cells() |> Seq.map(fun cell -> XlCell(cell))
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlCell> = 
      let cells = column.Cells() |> Seq.map(fun cell -> XlCell(cell))
      cells.GetEnumerator()
      

      
// TODO
and XlColumns internal (columns: IXLColumns) =
  member internal __.raw with get() = columns
  // TODO
  member __.style with get() = columns.Style
  


// TODO
and XlRangeColumns internal (range: IXLRangeRows) =
  member internal __.raw with get() = range
  // TODO
  member __.style with get() = range.Style
    

    
// TODO
and XlRange internal (range: IXLRange) =
  member internal __.raw with get() = range
  member __.value with set(value) = range.Value <- value
  member __.first_cell with get() = XlCell(range.FirstCell())
  member __.first_cell_used with get() = XlCell(range.FirstCellUsed())
  member __.last_cell with get() = XlCell(range.LastCell())
  member __.last_cell_used with get() = XlCell(range.LastCellUsed())
  // TODO
  member __.style with get() = range.Style

  member __.fx(value: obj) = range.FormulaA1 <- value.ToString()
  member __.set(value) = __.value <- box value
  member __.set_formula(value: string) = range.FormulaA1 <- value
  member __.set_formula_r1c1(value: string) = range.FormulaR1C1 <- value
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
  
  member __.clear(?options: ClearOption) = match options with Some opt -> range.Clear(opt) | None -> range.Clear()
  

  
// TODO
and XlTable internal (table: IXLTable) =
  member internal __.raw with get() = table
  // TODO
  member __.style with get() = table.Style
