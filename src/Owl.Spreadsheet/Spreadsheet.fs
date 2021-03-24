namespace Owl.Spreadsheet

open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet

module Spreadsheet =
  let private (| Xlsx | _ |) (path: string) = 
    if Path.GetExtension(path) = ".xlsx" then Some(path) else None
  let private (| Xlsm | _ |) (path: string) = 
    if Path.GetExtension(path) = ".xlsm" then Some(path) else None

  /// <summary>
  /// Spreadsheetドキュメントを新規作成する
  /// </summary>
  let public create filepath =
    let kind = 
      match filepath with
        | Xlsx _ -> SpreadsheetDocumentType.Workbook
        | Xlsm _ -> SpreadsheetDocumentType.MacroEnabledWorkbook
        | _ -> raise (exn "invalid file path")
    SpreadsheetDocument.Create(filepath, kind)
