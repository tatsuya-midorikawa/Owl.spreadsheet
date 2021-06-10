namespace Owl.Spreadsheet

open System.Linq
open Owl.Spreadsheet.Convert

module internal Xlfunc =
  let is_number value =
    match value.GetType() with
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> true
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> true
    | _ -> false

  let is_number_or_not_empty value =
    match value.GetType() with
    | t when t = typeof<bool> -> true
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> true
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> true
    | t when t = typeof<string> -> not(System.String.IsNullOrEmpty(to_string value))
    | t when t = typeof<datetime> -> not(System.String.IsNullOrEmpty(to_string value))
    | _ -> false

  let is_number_or_boolean value =
    match value.GetType() with
    | t when t = typeof<bool> -> to_bool(value)
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> true
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> true
    | _ -> false
    
  let to_number value =
    match value.GetType() with
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> to_double(value)
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> to_double(value)
    | _ -> raise(exn "Not supported type")

  let to_force_number value =
    match value.GetType() with
    | t when t = typeof<bool> -> if (value |> unbox<bool>) then 1.0 else 0.0
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> to_double(value)
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> to_double(value)
    | _ -> 0.0

  let value (cell: XlCell) = cell.value 

[<AbstractClass;Sealed;>]
type Xlfunc private() =
  static member IF(exp:bool, when_true:unit -> obj, when_false:unit -> obj) = if exp then when_true() else when_false()
  static member IF(exp:bool, when_true:obj, when_false:obj) = if exp then when_true else when_false

  static member AND(args: bool seq) = args.All(fun arg -> arg)
  static member AND(cells: XlCell seq) = cells |> Seq.map (Xlfunc.value >> to_bool) |> Xlfunc.AND
  static member AND(range: XlRange) = range.cells() |> Xlfunc.AND

  static member OR(args: bool seq) = args.Any(fun arg -> arg)
  static member OR(cells: XlCell seq) = cells |> Seq.map (Xlfunc.value >> to_bool) |> Xlfunc.OR
  static member OR(range: XlRange) = range.cells() |> Xlfunc.OR
  
  static member TODAY() = datetime.Today
  static member NOW() = datetime.Now

  static member MAX(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.max
  static member MAX(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MAX
  static member MAX(range: XlRange) = range.cells() |> Xlfunc.MAX

  static member MAXA(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.max
  static member MAXA(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MAXA
  static member MAXA(range: XlRange) = range.cells() |> Xlfunc.MAXA  

  static member MIN(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.min
  static member MIN(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MIN
  static member MIN(range: XlRange) = range.cells() |> Xlfunc.MIN

  static member MINA(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.min
  static member MINA(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MINA
  static member MINA(range: XlRange) = range.cells() |> Xlfunc.MINA

  static member SMALL(args: seq<#obj>, rank: int) =
    let xs = args |> Seq.filter Xlfunc.is_number
    let index = rank - 1
    if Seq.length xs < index then "#N/A" :> obj
    else xs |> Seq.map Xlfunc.to_number |> Seq.sort |> Seq.item index :> obj
  static member SMALL(cells: XlCell seq, rank: int) = (cells |> Seq.map Xlfunc.value, rank) |> Xlfunc.SMALL
  static member SMALL(range: XlRange, rank: int) = (range.cells(), rank) |> Xlfunc.SMALL
  
  static member public SUM(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.sum
  static member public SUM(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.SUM
  static member public SUM(range: XlRange) = range.cells() |> Xlfunc.SUM
  
  static member public SUMIF(condition_range: seq<#obj>, condition: obj -> bool, targets: seq<#obj>) = 
    Seq.zip condition_range targets 
    |> Seq.filter(fun (c, _) -> condition(c))
    |> Seq.map(fun (_, v) -> Xlfunc.to_number v)
    |> Seq.sum
  static member public SUMIF(condition_range: XlCell seq, condition: obj -> bool, targets: XlCell seq) =
    let cs = condition_range |> Seq.map Xlfunc.value
    let ts = targets |> Seq.map Xlfunc.value
    Xlfunc.SUMIF(cs, condition, ts)
  static member public SUMIF(condition_range: XlRange, condition: obj -> bool, targets: XlRange) =
    Xlfunc.SUMIF(condition_range.cells(), condition, targets.cells())

  static member public AVERAGE(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.average
  static member public AVERAGE(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.AVERAGE
  static member public AVERAGE(range: XlRange) = range.cells() |> Xlfunc.AVERAGE

  static member public AVERAGEA(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.average
  static member public AVERAGEA(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.AVERAGEA
  static member public AVERAGEA(range: XlRange) = range.cells() |> Xlfunc.AVERAGEA

  static member public COUNT(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.length
  static member public COUNT(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.COUNT
  static member public COUNT(range: XlRange) = range.cells() |> Xlfunc.COUNT

  static member public COUNTA(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.length
  static member public COUNTA(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.COUNTA
  static member public COUNTA(range: XlRange) = range.cells() |> Xlfunc.COUNTA

  static member public FLOOR(number: obj) = number |> (to_double >> floor)
  static member public FLOOR(number: XlCell) = Xlfunc.FLOOR number.value
  
  static member public ABS(number: obj) = number |> (to_double >> abs)
  static member public ABS(number: XlCell) = Xlfunc.ABS number.value

  static member public SIN(number: obj) = number |> (to_double >> sin)
  static member public SIN(number: XlCell) = Xlfunc.SIN number.value

  static member public COS(number: obj) = number |> (to_double >> cos)
  static member public COS(number: XlCell) = Xlfunc.COS number.value
  
  static member public TAN(number: obj) = number |> (to_double >> tan)
  static member public TAN(number: XlCell) = Xlfunc.TAN number.value

  static member public SINH(number: obj) = number |> (to_double >> sinh)
  static member public SINH(number: XlCell) = Xlfunc.SINH number.value

  static member public COSH(number: obj) = number |> (to_double >> cosh)
  static member public COSH(number: XlCell) = Xlfunc.COSH number.value

  static member public TANH(number: obj) = number |> (to_double >> tanh)
  static member public TANH(number: XlCell) = Xlfunc.TANH number.value
  
  static member public MOD(number: obj, divisor: obj) = to_double(number) % to_double(divisor)
  static member public MOD(number: XlCell, divisor: XlCell) = Xlfunc.MOD(number.value, divisor.value)
  
  static member public POWER(number1: obj, number2: obj) = to_double(number1) ** to_double(number2)
  static member public POWER(cell1: XlCell, cell2: XlCell) = Xlfunc.POWER(cell1.value, cell2.value)
  
  static member public PRODUCT(number1: obj, number2: obj) = to_double(number1) * to_double(number2)
  static member public PRODUCT(cell1: XlCell, cell2: XlCell) = Xlfunc.PRODUCT(cell1.value, cell2.value)

  // TODO
  static member public LOOKUP() = raise(exn "")

  static member public VLOOKUP(target:obj, range:XlRange, column:int) =
    let found = range.raw.Column(column).Cells().FirstOrDefault(fun c -> target = c.Value)
    if found = null then "#N/A" |> box else found.Value
  // TODO
  static member public VLOOKUP(target:obj, columns:XlColumns, row:int) =
    raise(exn "")
  static member public VLOOKUP(target:XlCell, range:XlRange, column:int) = Xlfunc.VLOOKUP(target.value, range, column)
  static member public VLOOKUP(target:XlCell, columns:XlColumns, column:int) = Xlfunc.VLOOKUP(target.value, columns, column)

  static member public HLOOKUP(target:obj, range:XlRange, row:int) =
    let found = range.raw.Row(row).Cells().FirstOrDefault(fun c -> target = c.Value)
    if found = null then "#N/A" |> box else found.Value
  // TODO
  static member public HLOOKUP(target:obj, rows:XlRows, row:int) =
    raise(exn "")
  static member public HLOOKUP(target:XlCell, range:XlRange, row:int) = Xlfunc.HLOOKUP(target.value, range, row)
  static member public HLOOKUP(target:XlCell, rows:XlRows, row:int) = Xlfunc.HLOOKUP(target.value, rows, row)
  
  // TODO
  static member public XLOOKUP() = raise(exn "")

  static member public RAND() = System.Random().NextDouble() |> box
  static member public RANDBETWEEN(min, max) = System.Random().Next(min, max)
