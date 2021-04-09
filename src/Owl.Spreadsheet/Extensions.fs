namespace Owl.Spreadsheet

open System
open System.Collections.Generic
open System.Collections.Concurrent
open System.Linq
open System.Text
open System.Runtime.CompilerServices

module internal SpreadsheetHelper =
  let column_name_storage = ConcurrentDictionary<int, string>()
  [<Literal>]
  let column_name_table = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  let table_size = column_name_table.Length

  [<Literal>]
  let alphabets = "ZABCDEFGHIJKLMNOPQRSTUVWXYzabcdefghijklmnopqrstuvwxy"
  [<Literal>]
  let upper_alphabets = "ZABCDEFGHIJKLMNOPQRSTUVWXY"
  [<Literal>]
  let lower_alphabets = "zabcdefghijklmnopqrstuvwxy"

  let is_alphabet (c: char) =
    Seq.exists (fun alphabet -> alphabet = c) alphabets
  let is_upper_alphabet (c: char) =
    Seq.exists (fun alphabet -> alphabet = c) upper_alphabets
  let is_lower_alphabet (c: char) =
    Seq.exists (fun alphabet -> alphabet = c) lower_alphabets

  let are_consists_only_of_alphabets (target: string) =
    target.All(fun c -> is_alphabet c)

  let rec column_number_rec (name: ReadOnlySpan<char>) (acc: int) =
    let digit = name.Length - 1
    if 0 <= digit then
      let n = (int name.[0]) - 64
      let sum = acc + n * (26.0 ** float digit |> int)
      column_number_rec (name.Slice(1)) sum
    else
      acc

  let column_number (name: string) =
    if name |> String.IsNullOrEmpty || not (are_consists_only_of_alphabets name) then
      raise (exn $"invalid column name: {name}")
    else
      column_number_rec (name.ToUpper().AsSpan()) 0

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

  /// <summary>
  /// Spreadsheetの列名を数値に変換する
  /// </summary>
  static member public to_column_number(name: string) =
    SpreadsheetHelper.column_number name
