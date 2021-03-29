namespace Owl.Spreadsheet

open System
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
  member __.copy_to(target: IXLCell) = XlCell(cell.CopyFrom(target))
  member __.copy_to(target: XlCell) = XlCell(cell.CopyFrom(target.raw))
  member __.copy_to(target: string) = XlCell(cell.CopyFrom(target))
