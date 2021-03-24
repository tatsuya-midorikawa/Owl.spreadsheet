namespace Owl.Spreadsheet

[<Struct>]
type Address = { row: int; column: int }
with member __.to_tuple() = (__.row, __.column)
