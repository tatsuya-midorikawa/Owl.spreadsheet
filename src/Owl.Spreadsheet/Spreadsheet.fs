namespace Owl.Spreadsheet

open ClosedXML.Excel

module Spreadsheet =
  [<Literal>]
  let TRUE = true
  [<Literal>]
  let FALSE = false

  /// <summary>
  /// ワークブックを閉じる
  /// </summary>
  let public close (workbook: XLWorkbook) =
    workbook.Dispose()

  /// <summary>
  /// ワークブックを名前をつけて保存する
  /// </summary>
  let public save_as (filepath: string) (workbook: XLWorkbook) =
    workbook.SaveAs(filepath)
    workbook
  
  /// <summary>
  /// ワークブックを保存する
  /// </summary>
  let public save (workbook: XLWorkbook) =
    workbook.Save()
    workbook
    
  /// <summary>
  /// ワークブックを保存したあとに閉じる
  /// </summary>
  let public save_and_close (workbook: XLWorkbook) =
    workbook |> (save >> close)

  /// <summary>
  /// ワークブックから指定の位置に存在するワークシートを取得する
  /// </summary>
  let public get_sheet_at (position: int) (workbook: XLWorkbook) =
    workbook.Worksheet(position) |> XlWorksheet

  /// <summary>
  /// ワークブックから指定のシート名と一致するワークシートを取得する
  /// </summary>
  let public get_sheet_with (name: string) (workbook: XLWorkbook) =
    workbook.Worksheet(name) |> XlWorksheet

  /// <summary>
  /// ワークブックに新規ワークシートを追加する
  /// </summary>
  let public add_sheet (sheet_name: string) (workbook: XLWorkbook) =
    workbook.Worksheets.Add(sheet_name) |> ignore
    workbook

  /// <summary>
  /// 新規ワークブックを作成する
  /// </summary>
  let public new_workbook() =
    let workbook = new XLWorkbook()
    let worksheet = workbook.Worksheets.Add("Sheet1")
    workbook
    
  /// <summary>
  /// 指定したファイルパスで新規ワークブックを作成する
  /// </summary>
  let public new_workbook_with (filepath: string) =
    new_workbook() |> save_as filepath
    
  /// <summary>
  /// テンプレートから新規ワークブックを作成する
  /// </summary>
  let public new_workbook_from (template: string) =
    XLWorkbook.OpenFromTemplate(template)

  /// <summary>
  /// 既存のワークブックを開く
  /// </summary>
  let public open_workbook (filepath: string) =
    new XLWorkbook(filepath)
   

