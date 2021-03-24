open Owl.Spreadsheet

"./sample.xlsx"
|> Spreadsheet.create_with_auto_save true
|> Spreadsheet.close

//"./sample.xlsx"
//|> Spreadsheet.create
//|> Spreadsheet.save
//|> Spreadsheet.close
