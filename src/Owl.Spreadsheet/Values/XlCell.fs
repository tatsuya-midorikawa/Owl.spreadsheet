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
       | t when t = typeof<int16> -> to_int16(cell.Value) |> unbox<'T>
       | t when t = typeof<int> -> to_int(cell.Value) |> unbox<'T>
       | t when t = typeof<int64> -> to_int64(cell.Value) |> unbox<'T>
       | t when t = typeof<float32> -> to_single(cell.Value) |> unbox<'T>
       | t when t = typeof<float> -> to_double(cell.Value) |> unbox<'T>
       | t when t = typeof<decimal> -> to_decimal(cell.Value) |> unbox<'T>
       | t when t = typeof<string> -> to_string(cell.Value) |> unbox<'T>
       | t when t = typeof<DateTime> -> to_datetime(cell.Value) |> unbox<'T>
       | t when t = typeof<obj> -> __.value |> unbox<'T>
       | _ -> raise(exn "")

  member __.set<'T>(value: 'T) = __.value <- box value
  member __.fx(value: obj) = cell.FormulaA1 <- value.ToString()
  member __.get_formula() = cell.FormulaA1 
  member __.set_formula(value: string) = cell.FormulaA1  <- value
  member __.get_formula_r1c1() = cell.FormulaR1C1
  member __.set_formula_r1c1(value: string) = cell.FormulaR1C1  <- value
