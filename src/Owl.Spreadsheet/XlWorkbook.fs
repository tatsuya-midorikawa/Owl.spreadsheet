namespace Owl.Spreadsheet

open System.Collections
open System.Collections.Generic
open ClosedXML.Excel

type XlWorkbook (book: XLWorkbook) =
  member internal __.raw with get() = book
  member __.Item with get(nth: int) = book.Worksheet(nth) |> XlWorksheet
  member __.Item with get(name: string) = book.Worksheet(name) |> XlWorksheet
  member __.worksheets with get() = book.Worksheets |> XlWorksheets
  member __.worksheet(nth: int) = book.Worksheet(nth) |> XlWorksheet
  member __.worksheet(name: string) = book.Worksheet(name) |> XlWorksheet
  member __.add(sheet_name: string) = book.Worksheets.Add(sheet_name) |> ignore
  member __.save() = book.Save()
  member __.save_as(filepath: string) = book.SaveAs(filepath)
  member __.save_and_close() = __.save(); __.close()
  member __.close() = book.Dispose()
  member __.at(nth: int) = book.Worksheet(nth) |> XlWorksheet
  member __.by(name: string) = book.Worksheet(name) |> XlWorksheet
  static member create() =
    let workbook = new XLWorkbook()
    workbook.Worksheets.Add("Sheet1") |> ignore
    XlWorkbook workbook
  static member create(filepath: string) =
    let book = XlWorkbook.create()
    book.save_as filepath
    book
  static member create_from(template: string) = XLWorkbook.OpenFromTemplate(template) |> XlWorkbook
  static member load(filepath: string) = new XLWorkbook(filepath) |> XlWorkbook

and XlWorksheet (sheet: IXLWorksheet) =
  member internal __.raw with get() = sheet
  member __.workbook with get() = sheet.Workbook |> XlWorkbook
  member __.save() = sheet.Workbook.Save()
  member __.save_as(filepath: string) = sheet.Workbook.SaveAs(filepath)
  member __.close() = sheet.Workbook.Dispose()
  member __.save_and_close() = __.save(); __.close()
  member __.Item
    with get(row: int, column: string) = sheet.Cell(row, column) |> XlCell
    and set(row: int, column: string) (value: obj) = sheet.Cell(row, column).Value <- value
  member __.Item 
    with get(row: int, column: int) = sheet.Cell(row, column) |> XlCell
    and set(row: int, column: int) (value: obj) = sheet.Cell(row, column).Value <- value
  member __.Item 
    with get(start', end') = sheet.Range($"%s{start'}:%s{end'}") |> XlRange
    and set(start', end') (value: obj) = sheet.Range($"%s{start'}:%s{end'}").Value <- value
  member __.Item
    with get(address: string) = sheet.Cells(address) |> XlCells
    and set(address: string) (value: obj) = sheet.Cells(address).Value <- value
  member __.cell(address: string) = sheet.Cell address |> XlCell
  member __.cell(row:int, colmun:int) = sheet.Cell(row, colmun) |> XlCell
  member __.cells(range: string) = sheet.Cells range |> XlCells
  member __.cells(from': int * int, to': int * int) = __.cells $"%s{from'.to_address()}:%s{to'.to_address()}"
  member __.range(range: string) = sheet.Range range |> XlRange
  member __.range(from': int * int, to': int * int) = __.range $"%s{from'.to_address()}:%s{to'.to_address()}"
  member __.range(from', to') = __.range $"%s{from'}:%s{to'}"
  member __.column(column: int) = sheet.Column(column) |> XlColumn
  member __.column(column: string) = sheet.Column(column) |> XlColumn
  member __.columns() = sheet.Columns() |> XlColumns
  member __.columns(from': int, to': int) = sheet.Columns(from', to') |> XlColumns
  member __.columns(columns: string) = sheet.Columns(columns) |> XlColumns
  member __.columns(from': string, to': string) = sheet.Columns(from', to') |> XlColumns
  member __.row(row) = sheet.Row(row) |> XlRow
  member __.rows() = sheet.Rows() |> XlRows
  member __.rows(first': int, last': int) = sheet.Rows(first', last') |> XlRows
  member __.rows(rows: string) = sheet.Rows(rows) |> XlRows

and XlWorksheets (sheets: IXLWorksheets) =
  member internal __.raw with get() = sheets
  member __.count with get() = sheets.Count
  member __.add(nth: int) = sheets.Add(nth) |> XlWorksheet
  member __.add(sheet_name: string) = sheets.Add(sheet_name) |> XlWorksheet
  member __.delete(nth: int) = sheets.Delete(nth)
  member __.delete(sheet_name: string) = sheets.Delete(sheet_name)
  member __.contains(sheet_name: string) = sheets.Contains(sheet_name)
  member __.worksheet(nth: int) = sheets.Worksheet(nth) |> XlWorksheet
  member __.worksheet(sheet_name: string) = sheets.Worksheet(sheet_name) |> XlWorksheet

  interface IEnumerable<XlWorksheet> with
    member __.GetEnumerator(): IEnumerator = 
      let cs = sheets |> Seq.map XlWorksheet
      (cs :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<XlWorksheet> =
      let cs = sheets |> Seq.map XlWorksheet
      cs.GetEnumerator()
