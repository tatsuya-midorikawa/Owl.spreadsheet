namespace Owl.Spreadsheet

open System.Linq

module ExcelFunction =
  let inline IF(expression, when_true, when_false) =
    if expression then when_true else when_false
  let inline AND(args: bool seq) = args.All(fun arg -> arg)
  let inline OR(args: bool seq) = args.Any(fun arg -> arg)
  let inline TODAY() = datetime.Today
  let inline NOW() = datetime.Now
  let inline MAX(args: ^T seq) = Seq.max args
  let inline MIN(args: ^T seq) = Seq.min args
  let inline SMALL(args: ^T seq, rank: int) = Seq.sort args |> Seq.item rank
  let inline SUM(args: ^T seq) = Seq.sum args
  let inline AVERAGE(args: ^T seq) = Seq.average args
  // TODO
  let inline COUNT() = raise(exn "")
  let inline FLOOR(number) = floor number
  let inline ABS(number) = abs number
  let inline SIN(number) = sin number
  let inline COS(number) = cos number
  let inline TAN(number) = tan number
  let inline SINH(number) = sinh number
  let inline COSH(number) = cosh number
  let inline TANH(number) = tanh number
  let inline MOD(number: ^T, divisor: ^T) = number % divisor
  let inline POWER(number: ^T, divisor: ^T) = number ** divisor
  let inline PRODUCT(number1: ^T, number2: ^T) = number1 * number2
  // TODO
  let inline LOOKUP() = raise(exn "")
  // TODO
  let inline VLOOKUP() = raise(exn "")
  // TODO
  let inline HLOOKUP() = raise(exn "")
  // TODO
  let inline XLOOKUP() = raise(exn "")
