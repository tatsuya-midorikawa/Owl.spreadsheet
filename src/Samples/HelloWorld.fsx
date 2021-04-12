#r "nuget: Owl.Spreadsheet"
open Owl.Spreadsheet.Spreadsheet
open Owl.Spreadsheet.Worksheet

let file_path = "./output/HelloWorld/sample.xlsx"

// Create a new .xlsx file
let workbook = new_workbook_with file_path
// Get the first worksheet
let worksheet = workbook |> get_sheet_at 1
// Get accessor to cell
let cell = worksheet |> get_cell

// Outputs multiplication table
for i in 1..9 do
 for j in 1..9 do
   cell(i, j).value <- i * j

// Save and close the .xlsx file
workbook |> save |> close



// Open an existing .xlsx file
let workbook' = open_workbook file_path
// And get accessor to cell
let cell' = workbook' |> get_sheet_at 1 |> get_cell

// Double the value in cell
for i in 1..9 do
  for j in 1..9 do
    let value = cell'(i, j).as_int
    cell'(i, j).value <- 2 * value

// Save and close it
workbook' |> save_and_close
