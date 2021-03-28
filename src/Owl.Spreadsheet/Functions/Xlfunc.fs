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

[<AbstractClass;Sealed;>]
type Xlfunc private() =
  /// <summary></summary>
  static member public IF(expression:bool, when_true:unit -> obj, when_false:unit -> obj) =
    if expression then when_true() else when_false()

  /// <summary></summary>
  static member public IF(expression:bool, when_true:obj, when_false:obj) =
    if expression then when_true else when_false

  /// <summary></summary>
  static member public AND(args: bool seq) = args.All(fun arg -> arg)
  /// <summary></summary>
  static member  public OR(args: bool seq) = args.Any(fun arg -> arg)
  /// <summary></summary>
  static member public TODAY() = datetime.Today
  /// <summary></summary>
  static member  public NOW() = datetime.Now

  /// <summary></summary>
  static member public MAX(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.max

    
  /// <summary></summary>
  static member public MAX(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.MAX

  /// <summary></summary>
  static member public MAXA(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.max

  /// <summary></summary>
  static member public MAXA(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.MAXA
    
  /// <summary></summary>
  static member public MIN(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.min
    
  /// <summary></summary>
  static member public MIN(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.MIN

  /// <summary></summary>
  static member public MINA(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number_or_not_empty |> Seq.map Xlfunc.to_force_number |> Seq.min

  /// <summary></summary>
  static member public MINA(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.MINA

  /// <summary></summary>
  static member public SMALL(args: seq<#obj>, rank: int) =
    let xs = args |> Seq.filter Xlfunc.is_number
    let index = rank - 1
    if Seq.length xs < index then "#N/A" :> obj
    else xs |> Seq.map Xlfunc.to_number |> Seq.sort |> Seq.item index :> obj

  /// <summary></summary>
  static member public SUM(args: seq<#obj>) =
    args |> Seq.filter Xlfunc.is_number |> Seq.map Xlfunc.to_number |> Seq.sum 

  /// <summary></summary>
  static member public SUM(args: XlCell seq) =
    args |> Seq.map (fun cell -> cell.value) |> Xlfunc.SUM

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
