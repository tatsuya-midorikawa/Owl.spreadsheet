namespace Owl.Spreadsheet

open System.Collections.Generic
open System.Collections.Concurrent
open System.Text
open System.Runtime.CompilerServices

module internal SpreadsheetHelper =
  let column_name_storage = ConcurrentDictionary<int, string>()
  [<Literal>]
  let column_name_table = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  let table_size = column_name_table.Length

[<Extension>]
type SpreadsheetHelper =
  [<Extension>]
  static member private as_string(stack: Stack<char>) =
    let acc = StringBuilder()
    stack |> Seq.iter (fun c -> acc.Append(c) |> ignore)
    acc.ToString()
    
  [<Extension>]
  static member private to_column_name_rec'(number: int, accumulator: Stack<char>) =
    if 0 <= number then 
      accumulator.Push(SpreadsheetHelper.column_name_table.[number % SpreadsheetHelper.table_size])
      if SpreadsheetHelper.table_size <= number then
        SpreadsheetHelper.to_column_name_rec'(number / SpreadsheetHelper.table_size - 1, accumulator )
    else
      ()
      
  [<Extension>]
  static member private to_column_name_rec(number: int) =
    let accumulator = Stack<char>()
    number.to_column_name_rec'(accumulator)
    accumulator

  /// <summary>
  /// 数値をSpreadsheetの列名に変換する
  /// </summary>
  [<Extension>]
  static member public to_column_name(number: int) =
    let n = number - 1
    match SpreadsheetHelper.column_name_storage.TryGetValue n with
    | (true, value) -> value
    | (false, _) -> 
      let column = n.to_column_name_rec().as_string()
      SpreadsheetHelper.column_name_storage.GetOrAdd(n, column)
      
  /// <summary>
  /// 行列番号をSpreadsheetのアドレスに変換する
  /// </summary>
  [<Extension>]
  static member public to_address((row, column): int * int) =
    $"%s{column.to_column_name()}%d{row}"
