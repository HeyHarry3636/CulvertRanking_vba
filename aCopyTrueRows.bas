Attribute VB_Name = "aCopyTrueRows"
Sub CopyTrueRows()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each Cell In Sheets("data").Range("AU:AU")
        If Cell.Value = True Then
            matchRow = Cell.Row
            rows(matchRow & ":" & matchRow).Select
            Selection.Copy

            Sheets("AUtrue").Select
            ActiveSheet.rows(matchRow).Select
            ActiveSheet.Paste
            Sheets("data").Select
        End If
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
