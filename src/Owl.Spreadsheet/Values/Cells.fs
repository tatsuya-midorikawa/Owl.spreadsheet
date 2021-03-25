namespace Owl.Spreadsheet

open System
open ClosedXML.Excel
open System.Collections
open System.Collections.Generic

type Cells internal (cells: IXLCells) =
  member __.value with set(value) = cells.Value <- value
  member __.get() = cells |> Seq.map(fun cell -> Cell(cell))
  member __.set(value) = __.value <- box value
  member __.set_formula(value: string) = cells.FormulaA1  <- value
  member __.set_formula_r1c1(value: string) = cells.FormulaR1C1  <- value

  interface IEnumerable<Cell> with
    member __.GetEnumerator(): IEnumerator = 
      let cells = cells |> Seq.map(fun cell -> Cell(cell))
      (cells :> IEnumerable).GetEnumerator()
    member __.GetEnumerator(): IEnumerator<Cell> = 
      let cells = cells |> Seq.map(fun cell -> Cell(cell))
      cells.GetEnumerator()
    
