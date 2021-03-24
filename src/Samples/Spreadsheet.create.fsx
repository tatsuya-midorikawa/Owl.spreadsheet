#r "nuget: Owl.Spreadsheet"
open Owl.Spreadsheet

let doc = Spreadsheet.create "./sample.xlsx"
doc.Close()
