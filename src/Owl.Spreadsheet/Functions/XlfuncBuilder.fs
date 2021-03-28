namespace Owl.Spreadsheet

open System.Linq

[<AbstractClass;Sealed;>]
type XlfuncBuilder private() =
  // TODO
  /// <summary></summary>
  static member public IF(expression, when_true, when_false) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public AND(args: bool seq) =  raise(exn "")
  // TODO
  /// <summary></summary>
  static member  public OR(args: bool seq) =  raise(exn "")
  /// <summary></summary>
  static member public TODAY() = "TODAY()"
  /// <summary></summary>
  static member  public NOW() = "NOW()"

  // TODO
  /// <summary></summary>
  static member public MAX(args: ^T seq) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public MIN(args: ^T seq) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public SMALL(args: ^T seq, rank: int) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public SUM(args: ^T seq) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public AVERAGE(args: ^T seq) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public COUNT() = raise(exn "")
  
  // TODO
  /// <summary></summary>
  static member public FLOOR(number: float) = raise(exn "")
  // TODO  
  /// <summary></summary>
  static member public ABS(number: float) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public SIN(number: float) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public COS(number: float) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public TAN(number: float) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public SINH(number: float) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public COSH(number: float) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public TANH(number: float) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public MOD(number: float32, divisor: float32) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public POWER(number: float32, index: float32) = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public PRODUCT(number1: number, number2: number) = raise(exn "")

  // TODO
  /// <summary></summary>
  static member public LOOKUP() = raise(exn "")

  /// <summary></summary>
  static member public VLOOKUP(target:obj, range:XlRange, column:int, mode:bool) =
    $"VLOOKUP({target},{range.raw.RangeAddress},{column},{mode.ToString().ToUpper()})"
  /// <summary></summary>
  static member public VLOOKUP(target:XlCell, range:XlRange, column:int, mode:bool) =
    $"VLOOKUP({target.raw.Address},{range.raw.RangeAddress},{column},{mode.ToString().ToUpper()})"

  // TODO
  /// <summary></summary>
  static member public HLOOKUP() = raise(exn "")
  // TODO
  /// <summary></summary>
  static member public XLOOKUP() = raise(exn "")
  /// <summary></summary>
  static member public RAND() = raise(exn "")
  /// <summary></summary>
  static member public RANDBETWEEN(min, max) = raise(exn "")
