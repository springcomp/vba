Attribute VB_Name = "modSample"
Option Explicit

Public Sub Sample()

    Dim nCount As Long
    Dim nIndex As Integer
    
    Let nCount = ThisWorkbook.Worksheets.Count
    Let nIndex = 1
    For nIndex = 1 To nCount
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Item(nIndex)
        MsgBox ws.Name
    Next
    
    Set ws = Nothing

End Sub
