open Owl.Spreadsheet
open Owl.Spreadsheet.Spreadsheet
open Owl.Spreadsheet.Worksheet

// ========================
// Sample 1.

//let workbook = new_workbook_with "./sample.xlsx"
//let worksheet = workbook |> get_sheet_at 1
//let cell = worksheet |> get_cell

//for i in 1..9 do
//  for j in 1..9 do
//    cell(i, j).value <- i * j

//workbook |> save |> close



// ========================
// Sample 2.

//let workbook = open_workbook "./sample.xlsx"
//let cell = workbook |> get_sheet_at 1 |> get_cell
//for i in 1..9 do
//  for j in 1..9 do
//    let value = cell(i, j).as_int
//    cell(i, j).as_int <- value

//workbook |> save_and_close



// ========================
// Sample 3.

let workbook = new_workbook_with "./sample.xlsx"
let worksheet = workbook |> get_sheet_at 1
let cell = worksheet |> get_cell

for i in 1..9 do
  for j in 1..9 do
    cell(i, j).set<string>($"{i}x{j} = {i*j}")

for i in 1..9 do
  for j in 11..20 do
    cell(i, j).set<int>(i * j)

let offset = 11
for i in 1..9 do
  for j in 11..20 do
    let value = cell(i, j).get<number>()
    cell(i+offset, j).set(value/2.)

workbook |> save |> close
