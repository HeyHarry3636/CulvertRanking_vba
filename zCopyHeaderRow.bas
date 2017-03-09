Attribute VB_Name = "zCopyHeaderRow"
Sub CopyHeaderRow()

    Dim ws As Worksheet
    Set ws = Worksheets("AUtrue")
    Dim HeaderRow As Range
    Set HeaderRow = Worksheets("data").Range("1:1")
    HeaderRow.Copy
    
    ws.Range(HeaderRow.Address).PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    ws.Paste
    
End Sub
