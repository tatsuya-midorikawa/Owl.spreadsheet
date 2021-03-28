namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

type XlCell internal (cell: IXLCell) =
  member private __.as_short 
    with get() =
      match cell.Value with
      | null -> 0s
      | :? double as d -> Convert.ToInt16 d
      | :? string as s -> match Int16.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a short.")
      | _ -> match Int16.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a short.")
    and set(value: int16) = cell.Value <- value

  member private __.as_int 
    with get() =
      match cell.Value with
      | null -> 0
      | :? double as d -> Convert.ToInt32 d
      | :? string as s -> match Int32.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a integer.")
      | _ -> match Int32.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a integer.")
    and set(value: int) = cell.Value <- value

  member private __.as_long
    with get() =
      match cell.Value with
      | null -> 0L
      | :? double as d -> Convert.ToInt64 d
      | :? string as s -> match Int64.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a long.")
      | _ -> match Int64.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a long.")
    and set(value: int64) = cell.Value <- value

  member private __.as_single
    with get() =
      match cell.Value with
      | null -> 0.f
      | :? double as d -> Convert.ToSingle d
      | :? string as s -> match Single.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a single.")
      | _ -> match Single.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a single.")
    and set(value: float32) = cell.Value <- value

  member private __.as_number
    with get() =
      match cell.Value with
      | null -> 0.
      | :? double as d -> d
      | :? string as s -> match Double.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a number.")
      | _ -> match Double.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a number.")
    and set(value: float) = cell.Value <- value

  member private __.as_money
    with get() =
      match cell.Value with
      | null -> 0.m
      | :? double as d -> Convert.ToDecimal d
      | :? string as s -> match Decimal.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a money.")
      | _ -> match Decimal.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a money.")
    and set(value: decimal) = cell.Value <- value

  member private __.as_string
    with get() = match cell.Value with null -> "" | :? string as s -> s | value -> value.ToString()
    and set(value: string) = cell.Value <- value
    
  member private __.as_datetime
    with get() =
      match cell.Value with
      | null -> DateTime.MinValue
      | :? string as s -> match DateTime.TryParse s with (true, date) -> date | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a datetime.")
      | _ -> match DateTime.TryParse $"{cell.Value}" with (true, date) -> date | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a datetime.")
    and set(value: DateTime) = cell.Value <- value
    
  member internal __.raw with get() = cell
  member __.value with get() = cell.Value and set(value) = cell.Value <- value
  member __.worksheet with get() = cell.Worksheet
  member __.get() = __.value
  member __.get<'T>() = 
    match typeof<'T> with
       | t when t = typeof<int16> -> __.as_short |> unbox<'T>
       | t when t = typeof<int> -> __.as_int |> unbox<'T>
       | t when t = typeof<int64> -> __.as_long |> unbox<'T>
       | t when t = typeof<float32> -> __.as_single |> unbox<'T>
       | t when t = typeof<float> -> __.as_number |> unbox<'T>
       | t when t = typeof<decimal> -> __.as_money |> unbox<'T>
       | t when t = typeof<string> -> __.as_string |> unbox<'T>
       | t when t = typeof<DateTime> -> __.as_datetime |> unbox<'T>
       | t when t = typeof<obj> -> __.value |> unbox<'T>
       | _ -> raise(exn "")

  member __.set<'T>(value: 'T) = __.value <- box value
  member __.fx(value: obj) = cell.FormulaA1 <- value.ToString()
  member __.get_formula() = cell.FormulaA1 
  member __.set_formula(value: string) = cell.FormulaA1  <- value
  member __.get_formula_r1c1() = cell.FormulaR1C1
  member __.set_formula_r1c1(value: string) = cell.FormulaR1C1  <- value
