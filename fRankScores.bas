Attribute VB_Name = "fRankScores"
Sub RankScores()
Attribute RankScores.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' RankScores Macro
'

'
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("Ranked").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Ranked").Sort.SortFields.Add Key:=Range("B2:B1000") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Ranked").Sort
        .SetRange Range("A1:B1000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    
    Dim ws As Worksheet
    Set ws = Worksheets("Ranked")
    
    Dim rngSht As Worksheet
    Set rngSht = Worksheets("Score Matrix")
    
    Dim lastRow As Long
    lastRow = (rngSht.Range("A13").CurrentRegion.rows.Count)
    
    ws.Range("A1:B" & lastRow).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    
    
    
    
End Sub
