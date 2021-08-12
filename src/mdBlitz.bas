Attribute VB_Name = "mdBlitz"
'@Folder("Custom")
Option Private Module
Option Explicit


Public Const gBlitzFolder As String = "sp_blitz"
Public Const gBlitzScriptFull As String = "Install-All-Scripts.sql"
Public Const gBlitzScriptCoreNoStore As String = "Install-Core-Blitz-No-Query-Store.sql"
Public Const gBlitzScriptCoreStore As String = "Install-Core-Blitz-With-Query-Store.sql"


Function getBlitzScriptPath()
    getBlitzScriptPath = getAbsolutePath(getCurrentFolder()) & "\" & gBlitzFolder & "\" & gBlitzScriptFull
End Function


Sub Install_sp_Blitz()
    Dim aFso As FileSystemObject
    Dim aFileRead As TextStream
    Dim aSQL As String
    
    Set aFso = New Scripting.FileSystemObject
    
    Set aFileRead = aFso.OpenTextFile(getAbsolutePath(gBlitzFolder & "\" & gBlitzScriptFull), ForReading, False)
    aSQL = aFileRead.ReadAll()

    Set aFileRead = Nothing
    Set aFso = Nothing
End Sub

Function getWorkbookToExtract() As Workbook
    Dim rgOutput As String
    Dim aWb As Workbook
    
    rgOutput = GetRadioShapeValue(wsControlCentreName, "rsOutput")
    Set aWb = Nothing
    If rgOutput = cOutputThisFile Then
        Set aWb = ThisWorkbook
    ElseIf rgOutput = cOutputNewFile Then
        Set aWb = createWorkbook(WorkbookTitle:=wbBlitzName)
    End If
    Set getWorkbookToExtract = aWb
End Function

Sub BlitzAllChecked()
    On Error GoTo ErrorHandler
    Dim aWb As Workbook
    Dim rgOutput As String
    
    rgOutput = GetRadioShapeValue(wsControlCentreName, "rsOutput")
    Set aWb = Nothing
    
    If rgOutput = cOutputThisFile Then
        Set aWb = ThisWorkbook
    ElseIf rgOutput = cOutputNewFile Then
        Set aWb = createWorkbook(WorkbookTitle:=wbBlitzName)
    End If
    
    If CheckShape_IsChecked(wsControlCentreName, cInclude_sp_Blitz) Then
        If rgOutput = cOutputIndividualFile Then
            Set aWb = createWorkbook(WorkbookTitle:=wbBlitzName)
        End If
        
        DatasetByNameToWorksheet Wb:=aWb, DataSetName:=dsBlitzName
    End If
    
    If CheckShape_IsChecked(wsControlCentreName, cInclude_sp_BlitzFirst) Then
        If rgOutput = cOutputIndividualFile Then
            Set aWb = createWorkbook(WorkbookTitle:=wbBlitzFirstName)
        End If
        
        DatasetByNameToWorksheet Wb:=aWb, DataSetName:=dsBlitzFirstName
    End If
    
    If CheckShape_IsChecked(wsControlCentreName, cInclude_sp_BlitzIndex) Then
        If rgOutput = cOutputIndividualFile Then
            Set aWb = createWorkbook(WorkbookTitle:=wbBlitzIndexName)
        End If
        
        DatasetByNameToWorksheet Wb:=aWb, DataSetName:=dsBlitzIndexName
    End If
    
    If CheckShape_IsChecked(wsControlCentreName, cInclude_sp_BlitzCache) Then
        If rgOutput = cOutputIndividualFile Then
            Set aWb = createWorkbook(WorkbookTitle:=wbBlitzCacheName)
        End If
        
        DatasetByNameToWorksheet Wb:=aWb, DataSetName:=dsBlitzCacheName
    End If
    
    If CheckShape_IsChecked(wsControlCentreName, cInclude_sp_BlitzWho) Then
        If rgOutput = cOutputIndividualFile Then
            Set aWb = createWorkbook(WorkbookTitle:=wbBlitzWhoName)
        End If
        
        DatasetByNameToWorksheet Wb:=aWb, DataSetName:=dsBlitzWhoName
    End If
    
Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "BlitzAllChecked", Err.Description
    GoTo Finally
End Sub


Sub sp_BlitzAny(DataSetName As String, WorkbookTitle As String, Optional ReportWorkbook As Workbook = Nothing)
    On Error GoTo ErrorHandler
    Dim rgOutput As String
    
    rgOutput = GetRadioShapeValue(wsControlCentreName, "rsOutput")
    If ReportWorkbook Is Nothing Then
        If rgOutput = cOutputThisFile Then
            Set ReportWorkbook = ThisWorkbook
        ElseIf rgOutput = cOutputIndividualFile Then
            Set ReportWorkbook = createWorkbook(WorkbookTitle:=WorkbookTitle)
        End If
    End If
    DatasetByNameToWorksheet Wb:=ReportWorkbook, DataSetName:=DataSetName
Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "sp_BlitzAny " & DataSetName, "Dataset: " & DataSetName & Err.Description
    GoTo Finally
End Sub

Sub sp_Blitz(Optional pReportWorkbook As Workbook = Nothing)
    On Error GoTo ErrorHandler
    
    If pReportWorkbook Is Nothing Then
        If GetRadioShapeValue(wsControlCentreName, "rsOutput") = cOutputThisFile Then
            Set pReportWorkbook = ThisWorkbook
        Else
            Set pReportWorkbook = createWorkbook(WorkbookTitle:=wbBlitzName)
        End If
    End If
    DatasetByNameToWorksheet Wb:=pReportWorkbook, DataSetName:=dsBlitzName
Finally:
    Exit Sub

ErrorHandler:
    'addRow_Log logError, "GetTextFromFile", Err.Description
    GoTo Finally
End Sub




Sub TickAllBlitz()
    On Error GoTo ErrorHandler
    Dim aResult As Boolean
    aResult = CheckBox_IsChecked(wsControlCentreName, "cbTurnAllOnOff")
    'Application.ScreenUpdating = False
    SetCheckBoxValue wsControlCentreName, "cbIncludeBlitz", aResult
    SetCheckBoxValue wsControlCentreName, "cbIncludeBlitzFirst", aResult
    SetCheckBoxValue wsControlCentreName, "cbIncludeBlitzIndex", aResult
    SetCheckBoxValue wsControlCentreName, "cbIncludeBlitzCache", aResult
    SetCheckBoxValue wsControlCentreName, "cbIncludeBlitzWho", aResult
    'Application.ScreenUpdating = True
Finally:
    Exit Sub
    
ErrorHandler:
    addRow_Log logError, "TickAllBlitz", Err.Description
    GoTo Finally

End Sub

Sub PerformanceTest()
    On Error GoTo ErrorHandler
    Dim aWb As Workbook
    Dim rgOutput As String
    
    rgOutput = GetRadioShapeValue(wsControlCentreName, "rsOutput")
    Set aWb = Nothing
    
    If rgOutput = cOutputThisFile Then
        Set aWb = ThisWorkbook
    ElseIf rgOutput = cOutputNewFile Then
        Set aWb = createWorkbook(WorkbookTitle:=wbBlitzName)
    End If
    
    If CheckShape_IsChecked(wsControlCentreName, cInclude_sp_Blitz) Then
        If rgOutput = cOutputIndividualFile Then
            Set aWb = createWorkbook(WorkbookTitle:=wbBlitzName)
        End If
        
        DatasetByNameToWorksheet Wb:=aWb, DataSetName:="Performance_Check"
    End If
    
    
Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "PerformanceTest", Err.Description
    GoTo Finally
End Sub
