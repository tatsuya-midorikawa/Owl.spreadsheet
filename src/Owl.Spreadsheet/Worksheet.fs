namespace Owl.Spreadsheet

open ClosedXML.Excel

module Worksheet =

  let public used_all_cells (sheet: IXLWorksheet) = sheet.Cells()
  let public cells (row, column) (sheet: IXLWorksheet) = ()
