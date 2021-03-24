open Owl.Spreadsheet.Spreadsheet
open Owl.Spreadsheet.Worksheet

let workbook = "./sample.xlsx" |> new_workbook_with
let worksheet = workbook |> get_sheet_at 1
let cell = worksheet |> get_cell

for i in 1..9 do
  for j in 1..9 do
    cell(i, j).Value <- i * j

workbook |> save |> close
