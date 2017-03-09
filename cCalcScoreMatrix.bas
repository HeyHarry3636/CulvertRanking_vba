Attribute VB_Name = "cCalcScoreMatrix"
Sub CalcScoreMatrix()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim strFormulas(1 To 10) As Variant
    Dim ws As Worksheet
    Set ws = Worksheets("Score Matrix")
    
    Dim rngSht As Worksheet
    Set rngSht = Worksheets("AUtrue")
    
    Dim lastRow As Long
    lastRow = (rngSht.Range("A1").CurrentRegion.rows.Count) + 12
    'Added +12 to lastRow because cells starts at A13 on 'Score Matrix' sheet
    
    'With ThisWorkbook.Sheets("Score Matrix")
    With ws
        strFormulas(1) = "=AUtrue!C1"
        strFormulas(2) = "=(IF(AUtrue!K1>$H$2, $H$1, IF(AUtrue!K1=$I$2, $I$1, IF(AUtrue!K1=$J$2, $J$1, IF(AUtrue!K1=$K$2, $K$1, 3)))))*$L$2"
        strFormulas(3) = "=IF( OR( AUtrue!AK1<$H$4, AUtrue!AK1>$H$3 ), $H$1, IF( OR( AUtrue!AK1<$I$4, AUtrue!AK1>$I$3 ), $I$1, IF( OR( AUtrue!AK1<$J$4, AUtrue!AK1>$J$3 ), $J$1, IF( OR( AUtrue!AK1<$K$3, AUtrue!AK1>$K$4 ), $K$1)))) * $L$3"
        strFormulas(4) = "=(IF(AUtrue!AB1=""Y"",$I$1, IF(AUtrue!AB1=""N"",$K$1)))*$L$5"
        strFormulas(5) = "=(IF(AUtrue!AF1=""Y"", $H$1, IF(AUtrue!AF1=""N"", $K$1)))*$L$6"
        strFormulas(6) = "=(IF(AUtrue!AC1=""Y"", $I$1, IF(AUtrue!AC1=""N"", $K$1)))*$L$7"
        strFormulas(7) = "=(IF(AUtrue!AJ1<$K$8, $K$1, IF(AUtrue!AJ1<$J$8, $J$1, IF(AUtrue!AJ1<$I$8, $I$1, IF(AUtrue!AJ1>=$H$8, $H$1)))))*$L$8"
        strFormulas(8) = "0"
        strFormulas(9) = "=(IF(AUtrue!AS1=""Poor"",$K$1,IF(AUtrue!AS1=""Fair"",$J$1,IF(AUtrue!AS1=""Good"",$I$1,IF(AUtrue!AS1=""Excellent"",$H$1)))))*$L$10"
        strFormulas(10) = "=SUM(B13:I13)"
        
        .Range("A13:J13").Formula = strFormulas
        .Range("A13:J" & lastRow).FillDown
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
