namespace Owl.Spreadsheet

open System.Collections
open System.Collections.Generic
open ClosedXML.Excel

type XlWorkbook (book: XLWorkbook) =
  member internal __.raw with get() = book
  member __.worksheets with get() = book.Worksheets |> XlWorksheets
  member __.worksheet(nth: int) = book.Worksheet(nth) |> XlWorksheet
  member __.worksheet(name: string) = book.Worksheet(name) |> XlWorksheet
  member __.save() = book.Save()
  member __.save_as(filepath: string) = book.SaveAs(filepath)
  member __.close() = book.Dispose()
  

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
    with get(start': string, end': string) = sheet.Range($"{start'}:{end'}") |> XlRange
    and set(start': string, end': string) (value: obj) = sheet.Range($"{start'}:{end'}").Value <- value
  member __.Item
    with get(address: string) = sheet.Cells(address) |> XlCells
    and set(address: string) (value: obj) = sheet.Cells(address).Value <- value


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


module XlWorkbook =
  /// <summary>
  /// ワークブックを閉じる
  /// </summary>
  let public close (workbook: XlWorkbook) =
    workbook.raw.Dispose()

  /// <summary>
  /// ワークブックを名前をつけて保存する
  /// </summary>
  let public save_as (filepath: string) (workbook: XlWorkbook) =
    workbook.raw.SaveAs(filepath)
    workbook
  
  /// <summary>
  /// ワークブックを保存する
  /// </summary>
  let public save (workbook: XlWorkbook) =
    workbook.raw.Save()
    workbook
    
  /// <summary>
  /// ワークブックを保存したあとに閉じる
  /// </summary>
  let public save_and_close (workbook: XlWorkbook) =
    workbook |> (save >> close)

  /// <summary>
  /// ワークブックから指定の位置に存在するワークシートを取得する
  /// </summary>
  let public get_sheet_at (position: int) (workbook: XlWorkbook) =
    workbook.raw.Worksheet(position) |> XlWorksheet

  /// <summary>
  /// ワークブックから指定のシート名と一致するワークシートを取得する
  /// </summary>
  let public get_sheet_with (name: string) (workbook: XlWorkbook) =
    workbook.raw.Worksheet(name) |> XlWorksheet

  /// <summary>
  /// ワークブックに新規ワークシートを追加する
  /// </summary>
  let public add_sheet (sheet_name: string) (workbook: XlWorkbook) =
    workbook.raw.Worksheets.Add(sheet_name) |> ignore
    workbook

  /// <summary>
  /// 新規ワークブックを作成する
  /// </summary>
  let public new_workbook() =
    let workbook = new XLWorkbook()
    let worksheet = workbook.Worksheets.Add("Sheet1")
    workbook |> XlWorkbook
    
  /// <summary>
  /// 指定したファイルパスで新規ワークブックを作成する
  /// </summary>
  let public new_workbook_with (filepath: string) =
    new_workbook() |> save_as filepath
    
  /// <summary>
  /// テンプレートから新規ワークブックを作成する
  /// </summary>
  let public new_workbook_from (template: string) =
    XLWorkbook.OpenFromTemplate(template) |> XlWorkbook

  /// <summary>
  /// 既存のワークブックを開く
  /// </summary>
  let public open_workbook (filepath: string) =
    new XLWorkbook(filepath) |> XlWorkbook


module XlWorksheet =
  let public get_cell_at (sheet: XlWorksheet) (address: string) =
    sheet.raw.Cell address |> XlCell
  let public get_cell (sheet: XlWorksheet) (row:int, colmun:int) =
    sheet.raw.Cell(row, colmun) |> XlCell
    
  let public get_cells_at (sheet: XlWorksheet) (range: string)  =
    sheet.raw.Cells range |> XlCells
  let public get_cells (sheet: XlWorksheet) (from': int * int) (to': int * int) =
    get_cells_at sheet $"%s{from'.to_address()}:%s{to'.to_address()}"

  let public get_range_at (sheet: XlWorksheet) (range: string) =
    sheet.raw.Range range |> XlRange
  let public get_range (sheet: XlWorksheet) (from': int * int) (to': int * int) =
    get_range_at sheet $"%s{from'.to_address()}:%s{to'.to_address()}"
  let public get_range_by (sheet: XlWorksheet) (from': string) (to': string) =
    get_range_at sheet $"{from'}:{to'}"
    
  let public get_column (sheet: XlWorksheet) (column: int) =
    sheet.raw.Column(column) |> XlColumn
  let public get_column_at (sheet: XlWorksheet) (column: string) =
    sheet.raw.Column(column) |> XlColumn
    
  let public get_all_columns (sheet: XlWorksheet) =
    sheet.raw.Columns() |> XlColumns
  let public get_columns (sheet: XlWorksheet) (from': int, to': int) =
    sheet.raw.Columns(from', to') |> XlColumns
  let public get_columns_at (sheet: XlWorksheet) (columns: string) =
    sheet.raw.Columns(columns) |> XlColumns
  let public get_columns_by (sheet: XlWorksheet) (from': string, to': string) =
    sheet.raw.Columns(from', to') |> XlColumns
    
  let public get_row (sheet: XlWorksheet) (row: int) =
    sheet.raw.Row(row) |> XlRow
  
  let public get_all_rows (sheet: XlWorksheet) =
    sheet.raw.Rows() |> XlRows
  let public get_rows (sheet: XlWorksheet) (first': int, last': int) =
    sheet.raw.Rows(first', last') |> XlRows
  let public get_rows_at (sheet: XlWorksheet) (rows: string) =
    sheet.raw.Rows(rows) |> XlRows
