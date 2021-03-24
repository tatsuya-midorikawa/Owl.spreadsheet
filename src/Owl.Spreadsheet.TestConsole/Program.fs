open Owl.Spreadsheet.Spreadsheet
open Owl.Spreadsheet.Worksheet

//let workbook = new_workbook_with "./sample.xlsx"
//let worksheet = workbook |> get_sheet_at 1
//let cell = worksheet |> get_cell

//for i in 1..9 do
//  for j in 1..9 do
//    cell(i, j).Value <- i * j

//workbook |> save |> close

let workbook = open_workbook "./sample.xlsx"
let cell = workbook |> get_sheet_at 1 |> get_cell
for i in 1..9 do
  for j in 1..9 do
    let value = cell(i, j).value
    cell(i, j).value <- value

workbook |> save_and_close
