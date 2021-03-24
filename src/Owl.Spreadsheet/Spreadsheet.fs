namespace Owl.Spreadsheet

open System.IO
open System.Linq
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet

module Spreadsheet =
  let private (| Xlsx | _ |) (path: string) = 
    if Path.GetExtension(path) = ".xlsx" then Some(path) else None
  
  let private (| Xlsm | _ |) (path: string) = 
    if Path.GetExtension(path) = ".xlsm" then Some(path) else None

  let private get_workbook_part (document: SpreadsheetDocument) =
    if document.WorkbookPart = null then
      document.AddWorkbookPart(Workbook = Workbook())
    else
      document.WorkbookPart
      
  let private get_worksheets (document: SpreadsheetDocument) =
    if document.WorkbookPart.Workbook.Sheets = null then
      document.WorkbookPart.Workbook.AppendChild(Sheets())
    else
      document.WorkbookPart.Workbook.Sheets

  /// <summary>
  /// Spreadsheetドキュメントにワークシートを追加する
  /// </summary>
  let public add_sheet (name: string) (document: SpreadsheetDocument) =
    let workbook_part = get_workbook_part document
    let worksheet_part = workbook_part.AddNewPart<WorksheetPart>()
    worksheet_part.Worksheet <- Worksheet(SheetData() :> OpenXmlElement)
    let sheets = get_worksheets document
    let sheet =
      Sheet(
        Id = StringValue(document.WorkbookPart.GetIdOfPart(worksheet_part)),
        SheetId = UInt32Value(uint(sheets.Count() + 1)),
        Name = StringValue(name))
    sheets.Append([| sheet :> OpenXmlElement |])
    printfn $"sheets count= {sheets.Count()}"
    workbook_part.Workbook.Save()
    document
  
  /// <summary>
  /// Spreadsheetドキュメントを保存する
  /// </summary>
  let public save (document: SpreadsheetDocument) =
    document.Save()
    document
    
  /// <summary>
  /// Spreadsheetドキュメントを名前を付けて保存する
  /// </summary>
  let public save_as filepath (document: SpreadsheetDocument) =
    document.SaveAs(filepath) :?> SpreadsheetDocument

  /// <summary>
  /// Spreadsheetドキュメントを閉じる
  /// </summary>
  let public close (document: SpreadsheetDocument) =
    document.Close()
    
  /// <summary>
  /// 空のSpreadsheetドキュメントを新規作成する
  /// </summary>
  let public create_with_auto_save autosave filepath =
    let kind = 
      match filepath with
        | Xlsx _ -> SpreadsheetDocumentType.Workbook
        | Xlsm _ -> SpreadsheetDocumentType.MacroEnabledWorkbook
        | _ -> raise (exn "invalid file path")
    SpreadsheetDocument.Create(filepath, kind, autosave)
    |> add_sheet "Sheet1"

  /// <summary>
  /// 空のSpreadsheetドキュメントを新規作成する
  /// </summary>
  let public create filepath =
    create_with_auto_save false filepath
