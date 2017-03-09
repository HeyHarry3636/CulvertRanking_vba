Attribute VB_Name = "bDeleteBlankRows"
Sub DeleteBlankRows()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim rng As Range
    Set rng = Worksheets("AUtrue").Range("A:A").SpecialCells(xlCellTypeBlanks)

    rng.EntireRow.Delete
  
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
