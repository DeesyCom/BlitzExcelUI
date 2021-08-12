Attribute VB_Name = "mdUI"
'@Folder("Main")
Option Private Module
Option Explicit


'------------------------------------------------
Sub clearTableContent(pWsName As String, TableName As String)
    Dim aWs As Worksheet
    Dim aListObj As ListObject
    
    Set aWs = ThisWorkbook.Worksheets(pWsName)
    Set aListObj = GetTableObject(aWs, TableName)
    If Not aListObj Is Nothing And Not aListObj.DataBodyRange Is Nothing Then
        aListObj.DataBodyRange.ClearContents
        aListObj.DataBodyRange.Delete
    End If
End Sub


Sub ClearLogTable()
    Dim aListObj As ListObject
    
    clearTableContent wsLogName, tblLogName

    Set aListObj = GetTableObject(ThisWorkbook.Worksheets(wsLogName), tblLogName)
    If Not aListObj Is Nothing And Not aListObj.DataBodyRange Is Nothing Then
    End If
End Sub


Sub CleanWorkbook()
    ClearLogTable
    Dim aWs As Worksheet
    Application.DisplayAlerts = False
    For Each aWs In ThisWorkbook.Worksheets
        If LCase(Left(aWs.CodeName, 2)) <> "ws" And aWs.Visible = xlSheetVisible Then
            aWs.Delete
        End If
    Next aWs
    Application.DisplayAlerts = True
End Sub


Function addRow_Log(aResult, aFunction, aMessage)
    On Error GoTo ErrorHandler
    Dim lo As ListObject
    Dim Ws As Worksheet
    Dim rng As Range

    Set Ws = ThisWorkbook.Sheets(tblLogName)
    Set lo = GetTableObject(Ws, tblLogName)

    If lo.InsertRowRange Is Nothing Then
        Set rng = lo.ListRows.Add.Range
    Else
        Set rng = lo.InsertRowRange
    End If
    rng.Cells(1, 1).Value = getIdNew()
    rng.Cells(1, 2).Value = aResult
    rng.Cells(1, 3).Value = Now()
    rng.Cells(1, 4).Value = aFunction
    rng.Cells(1, 5).Value = aMessage

    addRow_Log = rng.Cells(1, 1)
    Exit Function
ErrorHandler:
    addRow_Log = 0
End Function



Function getCellValue(pWsName As String, pCellName As String) As Variant
    getCellValue = ThisWorkbook.Worksheets(pWsName).Range(pCellName).Value
End Function


Sub setCellValue(pWsName As String, pCellName As String, pValue As Variant)
    ThisWorkbook.Worksheets(pWsName).Range(pCellName).Value = pValue
End Sub

