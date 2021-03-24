open Owl.Spreadsheet

//let workbook = Spreadsheet.new_workbook_with "./sample.xlsx"

//workbook.Cell("1").Value <- 500

//Spreadsheet.close workbook

seq{ 0..30 }
|> Seq.map (fun i -> i.to_column_name())
|> Seq.iter (printf "%s, ")
