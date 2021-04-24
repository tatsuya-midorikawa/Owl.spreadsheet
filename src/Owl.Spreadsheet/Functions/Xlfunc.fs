namespace Owl.Spreadsheet

open System.Linq
open Owl.Spreadsheet.Convert

module internal Xlfunc =
  /// <summary></summary>
  let is_number value =
    match value.GetType() with
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> true
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> true
    | _ -> false

  /// <summary></summary>
  let is_number_or_not_empty value =
    match value.GetType() with
    | t when t = typeof<bool> -> true
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> true
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> true
    | t when t = typeof<string> -> not(System.String.IsNullOrEmpty(to_string value))
    | t when t = typeof<datetime> -> not(System.String.IsNullOrEmpty(to_string value))
    | _ -> false

  /// <summary></summary>
  let is_number_or_boolean value =
    match value.GetType() with
    | t when t = typeof<bool> -> to_bool(value)
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> true
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> true
    | _ -> false
    
  /// <summary></summary>
  let to_number value =
    match value.GetType() with
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> to_double(value)
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> to_double(value)
    | _ -> raise(exn "Not supported type")

  /// <summary></summary>
  let to_force_number value =
    match value.GetType() with
    | t when t = typeof<bool> -> if (value |> unbox<bool>) then 1.0 else 0.0
    | t when t = typeof<int16> || t = typeof<int> || t = typeof<int64> -> to_double(value)
    | t when t = typeof<float32> || t = typeof<float> || t = typeof<decimal> -> to_double(value)
    | _ -> 0.0

  /// <summary></summary>
  let value (cell: XlCell) = cell.value 

[<AbstractClass;Sealed;>]
type Xlfunc private() =
  static member IF(expression:bool, when_true:unit -> obj, when_false:unit -> obj) =
    if expression then when_true() else when_false()
  static member IF(expression:bool, when_true:obj, when_false:obj) =
    if expression then when_true else when_false

  static member AND(args: bool seq) = args.All(fun arg -> arg)
  static member OR(args: bool seq) = args.Any(fun arg -> arg)
  static member TODAY() = datetime.Today
  static member NOW() = datetime.Now

  static member MAX(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.max
  static member MAX(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MAX
  static member MAX(cells: XlCells) = cells :> XlCell seq |> Xlfunc.MAX
  static member MAX(range: XlRange) = range.cells() |> Xlfunc.MAX
  static member MAX(row: XlRow) = row :> XlCell seq |> Xlfunc.MAX
  static member MAX(column: XlColumn) = column :> XlCell seq |> Xlfunc.MAX

  static member MAXA(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.max
  static member MAXA(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MAXA
  static member MAXA(cells: XlCells) = cells  :> XlCell seq |> Xlfunc.MAXA
  static member MAXA(range: XlRange) = range.cells() |> Xlfunc.MAXA  
  static member MAXA(row: XlRow) = row :> XlCell seq |> Xlfunc.MAXA
  static member MAXA(column: XlColumn) = column :> XlCell seq |> Xlfunc.MAXA

  static member MIN(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.min
  static member MIN(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MIN
  static member MIN(cells: XlCells) = cells :> XlCell seq |> Xlfunc.MIN
  static member MIN(range: XlRange) = range.cells() |> Xlfunc.MIN
  static member MIN(row: XlRow) = row :> XlCell seq |> Xlfunc.MIN
  static member MIN(column: XlColumn) = column :> XlCell seq |> Xlfunc.MIN

  static member MINA(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.min
  static member MINA(cells: XlCell seq) = cells |> Seq.map Xlfunc.value |> Xlfunc.MINA
  static member MINA(cells: XlCells) = cells :> XlCell seq |> Xlfunc.MINA
  static member MINA(range: XlRange) = range.cells() |> Xlfunc.MINA
  static member MINA(row: XlRow) = row :> XlCell seq |> Xlfunc.MINA
  static member MINA(column: XlColumn) = column :> XlCell seq |> Xlfunc.MINA

  static member SMALL(args: seq<#obj>, rank: int) =
    let xs = args |> Seq.filter Xlfunc.is_number
    let index = rank - 1
    if Seq.length xs < index then "#N/A" :> obj
    else xs |> Seq.map Xlfunc.to_number |> Seq.sort |> Seq.item index :> obj
  static member SMALL(cells: XlCell seq, rank: int) = (cells |> Seq.map Xlfunc.value, rank) |> Xlfunc.SMALL
  static member SMALL(cells: XlCells, rank: int) = (cells :> XlCell seq, rank) |> Xlfunc.SMALL
  static member SMALL(range: XlRange, rank: int) = (range.cells(), rank) |> Xlfunc.SMALL
  static member SMALL(row: XlRow, rank: int) = (row :> XlCell seq, rank) |> Xlfunc.SMALL
  static member SMALL(column: XlColumn, rank: int) = (column :> XlCell seq, rank) |> Xlfunc.SMALL
  
  static member public SUM(args: seq<#obj>) = args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.sum 
  static member public SUM(args: XlCell seq) = args |> Seq.map Xlfunc.value |> Xlfunc.SUM

  /// <summary></summary>
  static member public AVERAGE(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.average

  /// <summary></summary>
  static member public AVERAGE(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.AVERAGE

  /// <summary></summary>
  static member public AVERAGEA(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.average

  /// <summary></summary>
  static member public AVERAGEA(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.AVERAGEA

  /// <summary></summary>
  static member public COUNT(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number |> Seq.length

  /// <summary></summary>
  static member public COUNT(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.COUNT

  /// <summary></summary>
  static member public COUNTA(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.length

  /// <summary></summary>
  static member public COUNTA(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.COUNTA

  /// <summary></summary>
  static member public FLOOR(number: obj) = number |> (to_double >> floor)
  /// <summary></summary>
  static member public FLOOR(number: XlCell) = Xlfunc.FLOOR number.value
  
  /// <summary></summary>
  static member public ABS(number: obj) = number |> (to_double >> abs)
  /// <summary></summary>
  static member public ABS(number: XlCell) = Xlfunc.ABS number.value

  /// <summary></summary>
  static member public SIN(number: obj) = number |> (to_double >> sin)
  /// <summary></summary>
  static member public SIN(number: XlCell) = Xlfunc.SIN number.value

  /// <summary></summary>
  static member public COS(number: obj) = number |> (to_double >> cos)
  /// <summary></summary>
  static member public COS(number: XlCell) = Xlfunc.COS number.value
  
  /// <summary></summary>
  static member public TAN(number: obj) = number |> (to_double >> tan)
  /// <summary></summary>
  static member public TAN(number: XlCell) = Xlfunc.TAN number.value

  /// <summary></summary>
  static member public SINH(number: obj) = number |> (to_double >> sinh)
  /// <summary></summary>
  static member public SINH(number: XlCell) = Xlfunc.SINH number.value

  /// <summary></summary>
  static member public COSH(number: obj) = number |> (to_double >> cosh)
  /// <summary></summary>
  static member public COSH(number: XlCell) = Xlfunc.COSH number.value

  /// <summary></summary>
  static member public TANH(number: obj) = number |> (to_double >> tanh)
  /// <summary></summary>
  static member public TANH(number: XlCell) = Xlfunc.TANH number.value
  
  /// <summary></summary>
  static member public MOD(number: obj, divisor: obj) = to_double(number) % to_double(divisor)
  /// <summary></summary>
  static member public MOD(number: XlCell, divisor: XlCell) = Xlfunc.MOD(number.value, divisor.value)
  
  /// <summary></summary>
  static member public POWER(number1: obj, number2: obj) = to_double(number1) ** to_double(number2)
  /// <summary></summary>
  static member public POWER(cell1: XlCell, cell2: XlCell) = Xlfunc.POWER(cell1.value, cell2.value)
  
  /// <summary></summary>
  static member public PRODUCT(number1: obj, number2: obj) = to_double(number1) * to_double(number2)
  /// <summary></summary>
  static member public PRODUCT(cell1: XlCell, cell2: XlCell) = Xlfunc.PRODUCT(cell1.value, cell2.value)

  // TODO
  /// <summary></summary>
  static member public LOOKUP() = raise(exn "")

  /// <summary></summary>
  static member public VLOOKUP(target:obj, range:XlRange, column:int) =
    let found = range.raw.Column(column).Cells().FirstOrDefault(fun cell -> target = cell.Value)
    if found = null then "#N/A" |> box else found.Value

  /// <summary></summary>
  static member public VLOOKUP(target:XlCell, range:XlRange, column:int) = 
    Xlfunc.VLOOKUP(target.value, range, column)

  // TODO
  /// <summary></summary>
  static member public HLOOKUP() = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public XLOOKUP() = raise(exn "")

  /// <summary></summary>
  static member public RAND() =
    let rand = System.Random()
    rand.NextDouble() |> box
    
  /// <summary></summary>
  static member public RANDBETWEEN(min, max) = 
    let rand = System.Random()
    rand.Next(min, max)
