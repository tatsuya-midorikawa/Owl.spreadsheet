namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

type Cell(cell: IXLCell) =
  member __.as_int with get() =
    match cell.Value with
    | null -> 0
    | :? double as d -> Convert.ToInt32 d
    | :? string as s -> match Int32.TryParse s with (true, value) -> value | _ -> 0
    | _ -> match Int32.TryParse $"{cell.Value}" with (true, value) -> value | _ -> 0
    
  member __.as_long with get() =
    match cell.Value with
    | null -> 0L
    | :? double as d -> Convert.ToInt64 d
    | :? string as s -> match Int64.TryParse s with (true, value) -> value | _ -> 0L
    | _ -> match Int64.TryParse $"{cell.Value}" with (true, value) -> value | _ -> 0L

  member __.as_number with get() =
    match cell.Value with
    | null -> 0.
    | :? double as d -> d
    | :? string as s -> match Double.TryParse s with (true, value) -> value | _ -> 0.
    | _ -> match Double.TryParse $"{cell.Value}" with (true, value) -> value | _ -> 0.
    
  member __.as_money with get() =
    match cell.Value with
    | null -> 0.m
    | :? double as d -> Convert.ToDecimal d
    | :? string as s -> match Decimal.TryParse s with (true, value) -> value | _ -> 0.m
    | _ -> match Decimal.TryParse $"{cell.Value}" with (true, value) -> value | _ -> 0.m

  member __.as_string with get() = 
    match cell.Value with null -> "" | :? string as s -> s | value -> value.ToString()
    
  member __.as_datetime with get() =
    match cell.Value with
    | null -> DateTime.MinValue
    | :? string as s -> match DateTime.TryParse s with (true, date) -> date | _ -> DateTime.MinValue
    | _ -> match DateTime.TryParse $"{cell.Value}" with (true, date) -> date | _ -> DateTime.MinValue


  member __.value with get() = cell.Value and set (value) = cell.Value <- value
