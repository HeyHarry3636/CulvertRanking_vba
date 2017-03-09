Attribute VB_Name = "dCalcScoringTotals"
Sub CalcScoringTotals()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim strFormulas(1 To 2) As Variant
    Dim ws As Worksheet
    Set ws = Worksheets("Ranked")
    
    Dim rngSht As Worksheet
    Set rngSht = Worksheets("Score Matrix")
    
    Dim lastRow As Long
    lastRow = (rngSht.Range("A13").CurrentRegion.rows.Count)
    
    'With ThisWorkbook.Sheets("Score Matrix")
    With ws
        strFormulas(1) = "='Score Matrix'!A13"
        strFormulas(2) = "='Score Matrix'!J13"
        
        .Range("A2:B2").Formula = strFormulas
        .Range("A2:B" & lastRow).FillDown
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

