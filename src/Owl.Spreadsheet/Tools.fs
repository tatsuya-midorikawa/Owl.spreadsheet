namespace Owl.Spreadsheet

open System.Linq
open ClosedXML.Excel

module Tools =
  [<Literal>]
  let TRUE = true
  [<Literal>]
  let FALSE = false

  let load filepath = XlWorkbook.open_workbook filepath
  let create name = XlWorkbook.new_workbook_with name
  let workbook (sheet: XlWorksheet) = sheet.raw.Workbook

  let fst (workbook: XLWorkbook) = workbook.Worksheet(1) |> XlWorksheet
  let snd (workbook: XLWorkbook) = workbook.Worksheet(2) |> XlWorksheet
  let thd (workbook: XLWorkbook) = workbook.Worksheet(3) |> XlWorksheet
  let last (workbook: XLWorkbook) = workbook.Worksheet(workbook.Worksheets.Count) |> XlWorksheet
  let at (n: int) (workbook: XLWorkbook) = workbook.Worksheet(n) |> XlWorksheet
  let by (name: string) (workbook: XLWorkbook) = workbook.Worksheet(name) |> XlWorksheet

  let add (name: string) (workbook: XLWorkbook) = workbook.Worksheets.Add(name) |> XlWorksheet
  let del (name: string) (workbook: XLWorkbook) = workbook.Worksheets.Delete(name)
  let del_at (n: int) (workbook: XLWorkbook) = workbook.Worksheets.Delete(n)
  let del_by (n: int) (worksheet: XlWorksheet) = worksheet.raw.Workbook.Worksheets.Delete(worksheet.raw.Name)

  let save (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.save |> ignore
  let save_as (name: string) (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.save_as name |> ignore
  let save_and_close (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.save_and_close |> ignore
  let close (sheet: XlWorksheet) = sheet |> workbook |> XlWorkbook.close |> ignore

  