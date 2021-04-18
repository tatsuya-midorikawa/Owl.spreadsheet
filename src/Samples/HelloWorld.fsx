#r "nuget: Owl.Spreadsheet"
open Owl.Spreadsheet.Tools

let file_name = "sample.xlsx"
let workbook = create file_name
let sheet = fst workbook

workbook.at(1).cell(1,1).set("")
workbook |> at 1 |> cell (1,1) |> set ""
