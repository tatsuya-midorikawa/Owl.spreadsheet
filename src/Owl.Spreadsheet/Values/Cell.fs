namespace Owl.Spreadsheet

open System
open ClosedXML.Excel

type Cell(cell: IXLCell) =
  member __.as_int 
    with get() =
      match cell.Value with
      | null -> 0
      | :? double as d -> Convert.ToInt32 d
      | :? string as s -> match Int32.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a integer.")
      | _ -> match Int32.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a integer.")
    and set(value: int) = cell.Value <- value

  member __.as_long
    with get() =
      match cell.Value with
      | null -> 0L
      | :? double as d -> Convert.ToInt64 d
      | :? string as s -> match Int64.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a long.")
      | _ -> match Int64.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a long.")
    and set(value: int64) = cell.Value <- value

  member __.as_number
    with get() =
      match cell.Value with
      | null -> 0.
      | :? double as d -> d
      | :? string as s -> match Double.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a number.")
      | _ -> match Double.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a number.")
    and set(value: float) = cell.Value <- value

  member __.as_money
    with get() =
      match cell.Value with
      | null -> 0.m
      | :? double as d -> Convert.ToDecimal d
      | :? string as s -> match Decimal.TryParse s with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a money.")
      | _ -> match Decimal.TryParse $"{cell.Value}" with (true, value) -> value | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a money.")
    and set(value: decimal) = cell.Value <- value

  member __.as_string
    with get() = match cell.Value with null -> "" | :? string as s -> s | value -> value.ToString()
    and set(value: string) = cell.Value <- value

  member __.as_datetime
    with get() =
      match cell.Value with
      | null -> DateTime.MinValue
      | :? string as s -> match DateTime.TryParse s with (true, date) -> date | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a datetime.")
      | _ -> match DateTime.TryParse $"{cell.Value}" with (true, date) -> date | _ -> raise(InvalidCastException $"%A{cell.Value} can't be cast as a datetime.")
    and set(value: DateTime) = cell.Value <- value

  member __.value with get() = cell.Value and set (value) = cell.Value <- value
