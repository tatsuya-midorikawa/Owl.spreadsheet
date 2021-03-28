namespace Owl.Spreadsheet

open System
open System.Linq

module ExcelFunction =

  let inline private compare(lhs:obj, rhs:obj) =
    match rhs.GetType() with
    | rt when rt = typeof<int16> ->
      let (l, r) = (Convert.ToInt16 lhs, Convert.ToInt16 rhs)
      if l = r then 0 else if l < r then -1 else 1
    | rt when rt = typeof<int> ->
      let (l, r) = (Convert.ToInt32 lhs, Convert.ToInt32 rhs)
      if l = r then 0 else if l < r then -1 else 1
    | rt when rt = typeof<int64> ->
      let (l, r) = (Convert.ToInt64 lhs, Convert.ToInt64 rhs)
      if l = r then 0 else if l < r then -1 else 1
    | rt when rt = typeof<single> ->
      let (l, r) = (Convert.ToSingle lhs, Convert.ToSingle rhs)
      if l = r then 0 else if l < r then -1 else 1
    | rt when rt = typeof<double> ->
      let (l, r) = (Convert.ToDouble lhs, Convert.ToDouble rhs)
      if l = r then 0 else if l < r then -1 else 1
    | rt when rt = typeof<money> ->
      let (l, r) = (Convert.ToDecimal lhs, Convert.ToDecimal rhs)
      if l = r then 0 else if l < r then -1 else 1
    | rt when rt = typeof<datetime> ->
      let (l, r) = ((Convert.ToDateTime lhs).Ticks, (Convert.ToDateTime rhs).Ticks)
      if l = r then 0 else if l < r then -1 else 1
    | _ ->
      let (l, r) = (Convert.ToString lhs, Convert.ToString rhs)
      l.CompareTo(r)

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
  let inline POWER(number: ^T, index: ^T) = number ** index
  let inline PRODUCT(number1: ^T, number2: ^T) = number1 * number2
  // TODO
  let inline LOOKUP() = raise(exn "")
  // TODO
  let inline VLOOKUP(target:obj, range:XlRange, column:int, mode:bool) = 
    if mode then
      let mutable cache = System.Collections.Generic.List<ClosedXML.Excel.IXLCell>()
      for cell in range.raw.Column(column).Cells() do
        if compare(cell, target) <= 0 then cache.Add(cell)  
      if cache.Any() then
        cache.Sort()
        let last = cache.LastOrDefault()
        if last = null then obj() else last.Value
      else
        obj()
    else 
      let found = range.raw.Column(column).Cells().FirstOrDefault(fun cell -> obj.Equals(target, cell.Value))
      if found = null then obj() else found.Value
  // TODO
  let inline HLOOKUP() = raise(exn "")
  // TODO
  let inline XLOOKUP() = raise(exn "")
