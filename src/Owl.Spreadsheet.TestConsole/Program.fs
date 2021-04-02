open Owl.Spreadsheet
open Owl.Spreadsheet.Spreadsheet
open Owl.Spreadsheet.XlWorksheet
open type Owl.Spreadsheet.Xlfunc
//open Owl.Spreadsheet.Convert
//open type Owl.Spreadsheet.XlfuncBuilder

// ========================
// Sample 1.

//let workbook = new_workbook_with "./sample.xlsx"
//let worksheet = workbook |> get_sheet_at 1
//let cell = worksheet |> get_cell

//for i in 1..9 do
//  for j in 1..9 do
//    //cell(i, j).set(i * j)
//    cell(i, j).fx(PRODUCT(i, j))

//workbook |> save |> close



// ========================
// Sample 2.

//let workbook = open_workbook "./sample.xlsx"
//let cell = workbook |> get_sheet_at 1 |> get_cell
//for i in 1..9 do
//  for j in 1..9 do
//    let value = cell(i, j).get<int>()
//    cell(i, j).set(value)

//workbook |> save_and_close



// ========================
// Sample 3.

//let workbook = new_workbook_with "./sample.xlsx"
//let worksheet = workbook |> get_sheet_at 1
//let cell = worksheet |> get_cell

//for i in 1..9 do
//  for j in 1..9 do
//    cell(i, j).set<string>($"{i}x{j} = {i*j}")

//for i in 1..9 do
//  for j in 11..20 do
//    cell(i, j).set<int>(i * j)

//let offset = 11
//for i in 1..9 do
//  for j in 11..20 do
//    let value = cell(i, j).get<number>()
//    cell(i+offset, j).set(value/2.)

//workbook |> save |> close



// ========================
// Sample 4.

//let workbook = new_workbook_from "./template.xlsx"
//let cell = workbook |> get_sheet_at 1 |> get_cell

//// Set the number of lines in the header line as the offset value
//let offset = 1
//let today = datetime.Today
//for i in 1..10 do
//  let row = i + offset
//  cell(row, 1).set<string>(identifier.NewGuid().ToString())
//  cell(row, 2).set<string>($"midoliy {i:D3}")
//  cell(row, 3).set<datetime>(datetime(year=1989+i, month=9, day=13+i))
//  cell(row, 4).set_formula($"=DATEDIF({(row, 3).to_address()},TODAY(),\"Y\")")
//  // cell(row, 4).set<int>((integer.Parse(datetime.Today.ToString("yyyyMMdd")) - integer.Parse(cell(row, 3).get<datetime>().ToString("yyyyMMdd"))) / 10000)

//workbook |> save_as "./output.xlsx" |> close



// ========================
// Sample 5.

let workbook = new_workbook_with "./sample.xlsx"
let worksheet = workbook |> get_sheet_at 1
let cell = worksheet |> get_cell
let range = worksheet |> get_range

for i in 1..10 do
  cell(i, 1).set(RANDBETWEEN(1, 10))
  cell(i, 2).set(RANDBETWEEN(1, 10))
  cell(i, 3).set(POWER(cell(i, 1), cell(i, 2)))

cell(1, 4).set(11)
cell(2, 4).set(VLOOKUP(cell(1, 4), range(1, 1)(10, 1), 1))

workbook |> save |> close



//// ========================
//// Sample 6.

//let workbook = new_workbook_with "./sample.xlsx"
//let worksheet = workbook |> get_sheet_at 1
//let cell = worksheet |> get_cell
//let range = worksheet |> get_range

//cell(1, 1).set(datetime.Now)
//let v = cell(1, 1).value
//printfn $"{System.String.IsNullOrEmpty(to_string(v))}"

//cell(1, 1).set(null)
//let v2 = cell(1, 1).value
//printfn $"{v2.GetType()}"
////cell(2, 3).fx(VLOOKUP(cell(1, 2).value, range(1, 1)(10, 1), 1, TRUE))

//workbook |> save |> close
