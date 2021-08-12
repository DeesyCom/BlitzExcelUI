Attribute VB_Name = "mdGeneric"
'@Folder("Main")
Option Private Module
Option Explicit

'------------------------------------------------
Sub ResetAll()
    setValidation False
    closeConnection
    setCellValue wsControlCentreName, cCellServerName, getXmlAttributeFromNode(getXmlConfigNode(), cXmlAttributeNameServer)
    wsControlCentre.Activate
    wsControlCentre.Range(getXmlAttributeFromNode(getXmlConfigNode(), cXmlAttributeNameActiveCell)).Activate
    closeXML
    setStatusBarMessage vbNullString
End Sub

Sub Init()
    ResetAll
End Sub

Sub Finalize()
    ResetAll
    setValidation True
    checkLog
End Sub

Sub setValidation(aValue As Boolean)
    Application.ScreenUpdating = aValue
    'Application.DisplayStatusBar = aValue
    'Application.DisplayStatusBar = True
    Application.EnableEvents = aValue
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = aValue
    
    If aValue Then
        'Application.Calculation = xlCalculationAutomatic
    Else
        
        'Application.Calculation = xlCalculationManual
    End If
    Application.Calculation = xlCalculationAutomatic
End Sub


Sub setStatusBarMessage(Message As String)
    Application.StatusBar = Message
    DoEvents
End Sub


Sub checkLog()
    If Not GetTableObject(wsLog, tblLogName) Is Nothing Then
        If GetTableObject(wsLog, tblLogName).ListRows.Count > 0 Then
            wsLog.Activate
        End If
    End If
End Sub

Function createWorkbook(WorkbookTitle As String, Optional WorkbookFilename As String = vbNullString) As Workbook
    Set createWorkbook = Workbooks.Add(xlWBATWorksheet)
    With createWorkbook
        .Title = WorkbookTitle
        .Subject = vbNullString
        If WorkbookFilename <> vbNullString Then
            .SaveAs Filename:=WorkbookFilename
        End If
    End With
End Function


Function getIdNew() As Long
    getIdNew = 0
    Dim id As Long
    id = getCellValue(wsConfigName, cellLatestId) + 1
    Call setCellValue(wsConfigName, cellLatestId, id)
    getIdNew = id
End Function


Function getAbsolutePath(pPath As String) As String
    Dim aPath As String
    
    aPath = pPath
    
    If Mid(aPath, 2, 2) = ":\" Then
        'Acceptable as S:\Bla bla bla\
    ElseIf Mid(aPath, 1, 2) = "\\" Then
        'Acceptable  as \\myserver\local
    ElseIf Mid(aPath, 1, 1) = "\" Then
        'Acceptable  as \local, need to add current drive!
        aPath = Mid(ThisWorkbook.Path, 1, 2) & aPath
    Else
        aPath = ThisWorkbook.Path & "\" & aPath
    End If
    
    getAbsolutePath = aPath
End Function


Function getCurrentFolder() As String
    getCurrentFolder = ThisWorkbook.Path
End Function

Function GetTextFromFile() As String
    On Error GoTo ErrorHandler
    Dim aFso As FileSystemObject
    Dim aFileRead As TextStream
    Dim aFileText As String

    Set aFso = New Scripting.FileSystemObject
    
    Set aFileRead = aFso.OpenTextFile(getAbsolutePath(gBlitzFolder & "\" & gBlitzScriptFull), ForReading, False)
    aFileText = aFileRead.ReadAll()
    
    GetTextFromFile = aFileText
    
Finally:
    If Not aFileRead Is Nothing Then
        aFileRead.Close
        Set aFileRead = Nothing
    End If
    
    If Not aFso Is Nothing Then
        Set aFso = Nothing
    End If
    Exit Function

ErrorHandler:
    addRow_Log logError, "GetTextFromFile", Err.Description
    GoTo Finally

End Function


Function GetCheckBox(pWorkSheetName As String, pControlName As String) As CheckBox
    On Error GoTo ErrorHandler
    
    Dim aWs As Worksheet
    Dim aCheckBox
    Set aWs = ThisWorkbook.Worksheets(pWorkSheetName)
    If Not aWs Is Nothing Then
        For Each aCheckBox In aWs.CheckBoxes
            If aCheckBox.Name = pControlName Then
                Set GetCheckBox = aCheckBox
                Exit Function
            End If
        Next
    End If

Finally:
    Exit Function

ErrorHandler:
    addRow_Log logError, "GetCheckBox", Err.Description
    GoTo Finally
End Function


Sub SetCheckBoxValue(pWorkSheetName As String, pControlName As String, pCheck As Boolean)
    On Error GoTo ErrorHandler
    
    Dim aWs As Worksheet
    Dim aCheckBox
    Set aWs = ThisWorkbook.Worksheets(pWorkSheetName)
    If Not aWs Is Nothing Then
        For Each aCheckBox In aWs.CheckBoxes
            If aCheckBox.Name = pControlName Then
                aCheckBox.Value = pCheck
                Exit Sub
            End If
        Next
    End If

Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "SetCheckBoxValue", Err.Description
    GoTo Finally
End Sub

Function GetCheckBoxValueAsString(aValue As Boolean)
    If aValue Then
        GetCheckBoxValueAsString = "1"
    Else
        GetCheckBoxValueAsString = "0"
    End If
End Function


Function GetReportWorkbook(pWb As Workbook) As Workbook
    On Error GoTo ErrorHandler
    
    Dim aCb As CheckBox
    
    If pWb Is Nothing Then
        Set aCb = GetCheckBox(wsControlCentreName, cbDedicatedFileName)
        If Not aCb Is Nothing Then
            If aCb.Value = 1 Then
                Set GetReportWorkbook = createWorkbook(wbNewBookTitlePrefix & Format(Date, "yyyymmdd"))
            Else
                Set GetReportWorkbook = ThisWorkbook
            End If
        End If
    End If
    
Finally:
    Exit Function

ErrorHandler:
    'addRow_Log logError, "GetTextFromFile", Err.Description
    GoTo Finally
End Function



Function prepareWorksheet(pWb As Workbook, pWsName As String, Optional pClean As Boolean = True) As Worksheet
    On Error Resume Next
    Set prepareWorksheet = Nothing
    If Not pWb Is Nothing Then
        Dim aWs As Worksheet
        Set aWs = pWb.Sheets(pWsName)
        If aWs Is Nothing Then
            On Error GoTo 0
            Set aWs = pWb.Worksheets.Add(, pWb.Worksheets(pWb.Worksheets.Count))
            aWs.Name = pWsName
        ElseIf pClean Then
            aWs.Cells.Delete Shift:=xlUp
            'aWs.Cells.ClearContents
            'aWs.Cells.ClearFormats
            'aWs.Cells.Clear
            'aWs.CodeName = pWsName
        Else
            Application.DisplayAlerts = False
            aWs.Delete
            Application.DisplayAlerts = True
            Set aWs = pWb.Worksheets.Add(, pWb.Worksheets(pWb.Worksheets.Count))
            aWs.Name = pWsName
        End If
        Set prepareWorksheet = aWs
        removeDefaultWorksheets pWb
    End If
Finally:
    Exit Function
ErrorHandler:
    Set prepareWorksheet = aWs
    addRow_Log logError, "prepareWorksheet", "Failed: " & Err.Description
End Function


Sub removeDefaultWorksheets(pWb As Workbook)
    If Not pWb Is Nothing Then
        Dim Ws As Worksheet
        Dim currentState As Boolean
        currentState = Application.DisplayAlerts
        Application.DisplayAlerts = False

        For Each Ws In pWb.Worksheets
            If Left(Ws.Name, 5) = "Sheet" Then
                Ws.Delete
            End If
        Next Ws
        Application.DisplayAlerts = currentState
    End If
End Sub


Function IsRangePartOfTable(pRange As Range) As Boolean
    On Error Resume Next
    IsRangePartOfTable = (pRange.ListObject.Name <> vbNullString)
    On Error GoTo 0
End Function

Function TableExtract(TargetRange As Range, ByRef pRs As ADODB.RecordSet, DataSetName As String, Optional aShift = xlDown, Optional formatAsTable As Boolean = True) As Range
    On Error GoTo ErrorHandler
    Dim aColCount As Long, aRowCount As Long
    Dim aInsertInto As Range, aFullRange As Range
    Dim aTargetAddress
    
    If Not TargetRange Is Nothing And Not pRs Is Nothing Then
        If pRs.Fields.Count > 0 Then
            aColCount = pRs.Fields.Count
            aRowCount = pRs.RecordCount
            If aRowCount < 0 Then
                aRowCount = 0
            End If
            aTargetAddress = TargetRange.Address
            
            If IsRangePartOfTable(TargetRange) Then
                Exit Function
            End If
            
            On Error GoTo ErrorHandler
            Set aInsertInto = TargetRange.Worksheet.Range(TargetRange.Worksheet.Range(aTargetAddress), TargetRange.Worksheet.Range(aTargetAddress).Offset(aRowCount, aColCount - 1))
            aInsertInto.Insert Shift:=aShift, CopyOrigin:=xlFormatFromLeftOrAbove
    
            Set TargetRange = TargetRange.Worksheet.Range(aTargetAddress)
            formatPrintTableHeader TargetRange, pRs
            formatPrintTableData TargetRange.Offset(1), pRs
            
            Set aFullRange = TargetRange.Worksheet.Range(TargetRange.Worksheet.Range(aTargetAddress), TargetRange.Worksheet.Range(aTargetAddress).Offset(aRowCount, aColCount - 1))
            If formatAsTable Then
                TargetRange.Worksheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=aFullRange, XlListObjectHasHeaders:=xlYes).Name = DataSetName
            End If
            Set TableExtract = aFullRange
        End If
    End If
    
Finally:
    Exit Function
ErrorHandler:
    addRow_Log logError, "TableExtract", Err.Description
    Resume Next
    GoTo Finally
End Function


Function formatPrintTableHeader(TargetRange As Range, ByRef pRs As ADODB.RecordSet) As Range
    On Error GoTo ErrorHandler
    Dim aColCount As Long, iCounter As Integer
    
    aColCount = pRs.Fields.Count
    If Not pRs Is Nothing And Not TargetRange Is Nothing Then
        For iCounter = 0 To pRs.Fields.Count - 1
            TargetRange.Offset(ColumnOffset:=iCounter).Value = pRs.Fields(iCounter).Name
        Next iCounter
    End If
    
Finally:
    Exit Function
ErrorHandler:
    addRow_Log logError, "formatPrintTableHeader", Err.Description
    GoTo Finally
End Function


Function formatPrintTableData(TargetRange As Range, ByRef pRs As ADODB.RecordSet) As Range
    On Error GoTo ErrorHandler
    
    
    If Not TargetRange Is Nothing Then
        If Not pRs.RecordCount = 0 Then
            pRs.MoveFirst
        End If
        TargetRange.CopyFromRecordset pRs
    End If
    
Finally:
    Exit Function
ErrorHandler:
    addRow_Log logError, "formatPrintTableData", Err.Description
    GoTo Finally
End Function


Function PrintTableData(ByRef pWs As Worksheet, ByRef pRs As ADODB.RecordSet, ByRef pRow As Long) As Long
    On Error GoTo ErrorHandler
    Dim iCounter As Long, aColumnCount As Long
    aColumnCount = pRs.Fields.Count
    Dim aRange As Range

    While Not pRs.EOF
        For iCounter = 0 To aColumnCount - 1
            Set aRange = pWs.Cells(pRow, iCounter + 1)
            If (pRs.Fields(iCounter).Type = adBigInt) Or (pRs.Fields(iCounter).Type = adInteger) Then
                aRange.Value = vbNullString & pRs(iCounter).Value
            Else
                aRange.Value = pRs(iCounter).Value
            End If
        Next iCounter
        pRs.MoveNext
        pRow = pRow + 1
    Wend
    PrintTableData = pRow
Finally:
    Exit Function
ErrorHandler:
    addRow_Log logError, "PrintTableData", Err.Description
    GoTo Finally
End Function

Function CopyFromRecordset(ByRef pRs As ADODB.RecordSet, TargetRange As Range) As Long
    On Error GoTo ErrorHandler
    Dim iCounter As Long
    
    If Not pRs Is Nothing And Not TargetRange Is Nothing Then
        For iCounter = 0 To pRs.Fields.Count - 1
            TargetRange.Offset(ColumnOffset:=iCounter).Value = pRs.Fields(iCounter).Name
        Next iCounter
        
        TargetRange.Offset(1).CopyFromRecordset pRs
    End If
    
Finally:
    Exit Function
ErrorHandler:
    addRow_Log logError, "CopyFromRecordset", Err.Description
    GoTo Finally
End Function


Sub setShapeShadow(ActionShape As Shape)
    On Error GoTo ErrorHandler
    If Not ActionShape Is Nothing Then
        With ActionShape.Shadow
            '.Type = msoShadow25
            .Visible = msoTrue
            '.Style = msoShadowStyleOuterShadow
            .Blur = 15
            .OffsetX = 2.8284271247
            .OffsetY = 2.8284271247
            .RotateWithShape = msoFalse
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0.5
            .Size = 100
        End With
    End If
Finally:
    Exit Sub
ErrorHandler:
    addRow_Log logError, "setShapeShadow", Err.Description
    GoTo Finally
End Sub


Sub ExecuteByShapeName(ShapeName As String)
    Dim aWb As Workbook
    Dim rgOutput As String

    Init
    rgOutput = GetRadioShapeValue(wsControlCentreName, "rsOutput")
    Set aWb = Nothing
    
    If rgOutput = cOutputThisFile Then
        Set aWb = ThisWorkbook
    Else
        Set aWb = createWorkbook(WorkbookTitle:=wbBlitzName)
    End If
    
    If ShapeName = Empty Then
        Exit Sub
    ElseIf ShapeName = "csBlitz" Then
        sp_BlitzAny DataSetName:=wbBlitzName, WorkbookTitle:=wbBlitzName, ReportWorkbook:=aWb 'sp_Blitz
    ElseIf ShapeName = "csBlitzFirst" Then
        sp_BlitzAny DataSetName:=wbBlitzFirstName, WorkbookTitle:=wbBlitzFirstName, ReportWorkbook:=aWb 'sp_BlitzFirst
    ElseIf ShapeName = "csBlitzIndex" Then
        sp_BlitzAny DataSetName:=wbBlitzIndexName, WorkbookTitle:=wbBlitzIndexName, ReportWorkbook:=aWb 'sp_BlitzIndex
    ElseIf ShapeName = "csBlitzCache" Then
        sp_BlitzAny DataSetName:=wbBlitzCacheName, WorkbookTitle:=wbBlitzCacheName, ReportWorkbook:=aWb 'sp_BlitzCache
    ElseIf ShapeName = "csBlitzWho" Then
        sp_BlitzAny DataSetName:=wbBlitzWhoName, WorkbookTitle:=wbBlitzWhoName, ReportWorkbook:=aWb 'sp_BlitzWho
    End If
    Finalize
End Sub


Sub RadioCheckShapeGroupOff(WorkSheetName As String, GroupPrefix As String, Optional Active As String = vbNullString)
    On Error GoTo ErrorHandler
    Dim aWorkSheet As Worksheet, aShape As Shape
    
    If WorkSheetName = Empty Or GroupPrefix = Empty Then
        Exit Sub
    End If
    
    Set aWorkSheet = ActiveWorkbook.Sheets(WorkSheetName)
    If Not aWorkSheet Is Nothing Then
        For Each aShape In aWorkSheet.Shapes
            If aShape.Name = Active Then
                aShape.ShapeStyle = cCheckShapeStyleActive
            ElseIf aShape.Name Like GroupPrefix & "*" Then
                aShape.ShapeStyle = cCheckShapeStyleRadioNonActive
            End If
        Next
    End If
Finally:
    Exit Sub
ErrorHandler:
    addRow_Log logError, "RadioCheckShapeGroupOff", Err.Description
    GoTo Finally
End Sub

Function GetRadioShapeValue(WorkSheetName As String, RadioGroupName As String) As String
    On Error GoTo ErrorHandler
    Dim aWs As Worksheet, aShape As Shape
    
    GetRadioShapeValue = vbNullString
    If WorkSheetName = vbNullString Or RadioGroupName = vbNullString Then
        Exit Function
    End If
    
    Set aWs = ActiveWorkbook.Worksheets(WorkSheetName)
    If Not aWs Is Nothing Then
        For Each aShape In aWs.Shapes
            If aShape.Name Like RadioGroupName & "*" Then
                If aShape.ShapeStyle = cCheckShapeStyleActive Then
                    GetRadioShapeValue = Mid(aShape.Name, InStr(1, aShape.Name, "_") + 1, 255)
                End If
            End If
        Next
    End If

Finally:
    Exit Function
ErrorHandler:
    addRow_Log logError, "GetRadioShapeValue", Err.Description
    GoTo Finally
End Function

Sub RadioCheckShape()
    On Error GoTo ErrorHandler
    Dim aShapeName As String, aShape As Shape, aGroupPrefix As String
    Dim aShiftPressed As Integer, aCtrlPressed As Integer, aAltPressed As Integer
    
    ResetAll
    aShapeName = Application.Caller
    If aShapeName = Empty Or Not aShapeName Like cRadioShapeLikePrefix Then
        Exit Sub
    End If
    
    aShiftPressed = GetKeyState(SHIFT_KEY)
    aCtrlPressed = GetKeyState(CTRL_KEY)
    aAltPressed = GetKeyState(ALT_KEY)
    
    Set aShape = ActiveSheet.Shapes(Application.Caller)
    aGroupPrefix = vbNullString
    If InStr(1, aShapeName, "_") > 0 Then
        aGroupPrefix = Left(aShapeName, InStr(1, aShapeName, "_"))
    End If
    
    setStatusBarMessage wsControlCentre.Shapes(Application.Caller).ShapeStyle

    If aShiftPressed < 0 Then
        'ExecuteByShapeName aShapeName
        Exit Sub
    End If
    
    With aShape
        If .ShapeStyle = cCheckShapeStyleActive Then
            '.ShapeStyle = cCheckShapeStyleNonActive
        Else
            RadioCheckShapeGroupOff ActiveSheet.Name, aGroupPrefix, aShapeName
            .ShapeStyle = cCheckShapeStyleActive
        End If
    End With
Finally:
    Exit Sub
ErrorHandler:
    addRow_Log logError, "TickCheckShape", Err.Description
    GoTo Finally
End Sub


Sub TickCheckShape()
    On Error GoTo ErrorHandler
    Dim aShapeName As String, aShape As Shape
    Dim aShiftPressed As Integer, aCtrlPressed As Integer, aAltPressed As Integer
      
    ResetAll
    aShapeName = Application.Caller
    If aShapeName = Empty Or Not aShapeName Like cCheckShapeLikePrefix Then
        Exit Sub
    End If
    
    aShiftPressed = GetKeyState(SHIFT_KEY)
    aCtrlPressed = GetKeyState(CTRL_KEY)
    aAltPressed = GetKeyState(ALT_KEY)
    
    Set aShape = ActiveSheet.Shapes(Application.Caller)
    'setStatusBarMessage wsControlCentre.Shapes(Application.Caller).ShapeStyle

    If aShiftPressed < 0 Then
        ExecuteByShapeName aShapeName
        Exit Sub
    End If
    
    With aShape
        If .ShapeStyle = cCheckShapeStyleActive Then
            .ShapeStyle = cCheckShapeStyleNonActive
        Else
            .ShapeStyle = cCheckShapeStyleActive
        End If
        'setShapeShadow aShape
    End With
Finally:
    Exit Sub
ErrorHandler:
    addRow_Log logError, "TickCheckShape", Err.Description
    GoTo Finally
End Sub


Function WorksheetToWindow(pWs As Worksheet) As WorksheetView
    Dim aWsV As WorksheetView
    Set WorksheetToWindow = Nothing
    If Not pWs Is Nothing Then
        For Each aWsV In pWs.Parent.Windows(1).SheetViews
            If aWsV.Sheet.Name = pWs.Name Then
                Set WorksheetToWindow = aWsV
                Exit Function
            End If
        Next
    End If
End Function

Sub WorksheetHideGrid(pWs As Worksheet, Optional pHide As Boolean = False)
    With WorksheetToWindow(pWs)
        .DisplayGridlines = Not pHide
    End With
End Sub

Sub WorksheetHideHeading(pWs As Worksheet, Optional pHide As Boolean = False)
    With WorksheetToWindow(pWs)
        .DisplayHeadings = Not pHide
    End With
End Sub

Function GetTableObject(pWs As Worksheet, TableName As String) As ListObject
    On Error GoTo ErrorHandler
    If pWs Is Nothing Or TableName = "" Then
        Set GetTableObject = Nothing
        Exit Function
    End If
    Set GetTableObject = pWs.ListObjects(TableName)
Finally:
    Exit Function

ErrorHandler:
    addRow_Log logError, "GetTableObject", Err.Description
    Resume Next
    GoTo Finally
End Function

Sub TableSortDataByColumn(TargetRange As Range, TableName As String, SortColumnName As String)
    On Error GoTo ErrorHandler
    
    If TargetRange Is Nothing Or TableName = "" Or SortColumnName = "" Then
        Exit Sub
    End If
    Dim aWs As Worksheet
    Dim aListObject As ListObject
    Dim aSortColumnName As String
    Dim aSortOrder As Variant
    
    Set aWs = TargetRange.Worksheet
    If Not aWs Is Nothing Then
        Set aListObject = GetTableObject(aWs, TableName)
        If Left(SortColumnName, 1) = "-" Then
            aSortOrder = xlDescending
            aSortColumnName = Mid(SortColumnName, 2, Len(SortColumnName))
        Else
            aSortOrder = xlAscending
            aSortColumnName = SortColumnName
        End If
        If Not aListObject Is Nothing Then
            With aListObject.Sort.SortFields
                .Clear
                .Add Key:=Range(TableName & "[" & getEscapedFieldName(aSortColumnName) & "]"), _
                    Order:=aSortOrder, SortOn:=xlSortOnValues, DataOption:=xlSortNormal
            End With
            With aListObject.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    End If
Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "TableSortDataByColumn", Err.Description
    Resume Next
    GoTo Finally
End Sub

Sub TableDeleteColumn()

End Sub


Function getEscapedFieldName(FieldName As String) As String
    getEscapedFieldName = Replace(Replace(Replace(FieldName, "#", "'#"), "[", "'["), "]", "']")
End Function

Function getColumnExists(TableName As String, ColumnName As String) As Boolean
    Dim aTable As ListObject, aColumn As ListColumn, aRange As Range
    On Error GoTo ErrorHandler
    getColumnExists = False
    
    Set aRange = Range(TableName & "[" & getEscapedFieldName(ColumnName) & "]")
    
    'Set aTable = GetTableObject(TableName)
    'Set aColumn = aTable.ListColumns(ColumnName)
    getColumnExists = True
    Exit Function
    
ErrorHandler:
    
End Function

Sub TableFormatColumn(ColumnFormatNode As IXMLDOMNode, TableName As String)
    On Error GoTo ErrorHandler
    Dim aColumnName As String
    Dim aFormatRange As Range, aRange As Range
    Dim aHRefAddress As String, aValue As String
    Dim aColumnWidth As String

    
    aColumnName = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameName)
    On Error GoTo NoFieldFound
    Set aFormatRange = Range(TableName & "[" & getEscapedFieldName(aColumnName) & "]")
    On Error GoTo ErrorHandler

    If Not aFormatRange Is Nothing Then
        If getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnDelete) = cXmlAttributeValueYes Then
            aFormatRange.Delete
            Exit Sub
        End If
        
        If getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnForceToNumber) = cXmlAttributeValueYes Then
            aFormatRange.TextToColumns Destination:=aFormatRange, TextQualifier:=xlDoubleQuote
        End If
        
        aValue = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnTextWrap)
        If aValue = cXmlAttributeValueYes Then
            aFormatRange.WrapText = True
        ElseIf aValue = cXmlAttributeValueNo Then
            aFormatRange.WrapText = False
        End If
        
        aValue = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnVerticalAlignment)
        If aValue = cXmlAttributeNameColumnVerticalAlignmentTop Then
            aFormatRange.VerticalAlignment = xlTop
        ElseIf aValue = cXmlAttributeNameColumnVerticalAlignmentCenter Then
            aFormatRange.VerticalAlignment = xlCenter
        ElseIf aValue = cXmlAttributeNameColumnVerticalAlignmentBottom Then
            aFormatRange.VerticalAlignment = xlBottom
        End If
        
        aValue = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnHorizontalAlignment)
        If aValue = cXmlAttributeNameColumnHorizontalAlignmentLeft Then
            aFormatRange.HorizontalAlignment = xlLeft
        ElseIf aValue = cXmlAttributeNameColumnHorizontalAlignmentCenter Then
            aFormatRange.HorizontalAlignment = xlCenter
        ElseIf aValue = cXmlAttributeNameColumnHorizontalAlignmentRight Then
            aFormatRange.HorizontalAlignment = xlRight
        ElseIf aValue = cXmlAttributeNameColumnHorizontalAlignmentGeneral Then
            aFormatRange.HorizontalAlignment = xlGeneral
        End If
        
        aColumnWidth = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnWidth)
        If aColumnWidth <> "" Then
            If aColumnWidth = cXmlAttributeNameColumnWidthAuto Then
                aFormatRange.EntireColumn.AutoFit
            ElseIf IsNumeric(aColumnWidth) Then
                aFormatRange.ColumnWidth = CDec(aColumnWidth)
            End If
        End If
        
        If getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnHRefFrom) <> "" Then
            aHRefAddress = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnHRefFrom)
            If getColumnExists(TableName, aHRefAddress) Then
                For Each aRange In aFormatRange
                    aRange.FormulaR1C1 = "=HYPERLINK([@" & aHRefAddress & "],""" & aRange.Value & """)"
                Next
            End If
        End If
        
        aFormatRange.NumberFormat = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnFormat)
        
        aValue = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnColorScale)
        If aValue = cXmlAttributeNameColumnColorScaleGYR Or aValue = cXmlAttributeNameColumnColorScaleRYG Then
            aFormatRange.FormatConditions.AddColorScale ColorScaleType:=3
            If aValue = cXmlAttributeNameColumnColorScaleGYR Then
                aFormatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 7039480
                aFormatRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 8109667
            ElseIf aValue = cXmlAttributeNameColumnColorScaleRYG Then
                aFormatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 8109667
                aFormatRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 7039480
            End If
        End If
        
        If getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnDatabar) = cXmlAttributeValueYes Then
            aFormatRange.FormatConditions.AddDatabar
            With aFormatRange.FormatConditions(1)
                .ShowValue = True
                .SetFirstPriority
                .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
                .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
                .BarColor.Color = getXmlAttributeFromNodeAsDouble(ColumnFormatNode, cXmlAttributeNameColumnBarColor, cXmlAttributeNameColumnBarColorDefault)
                .BarColor.TintAndShade = 0
                .BarBorder.Color.TintAndShade = 0
                .BarBorder.Type = xlDataBarBorderSolid
                .BarFillType = xlDataBarFillGradient
                .Direction = xlContext
                
                .NegativeBarFormat.ColorType = xlDataBarColor
                .NegativeBarFormat.BorderColorType = xlDataBarColor
                .NegativeBarFormat.Color.Color = 255
                .NegativeBarFormat.Color.TintAndShade = 0
                .NegativeBarFormat.BorderColor.Color = 255
                .NegativeBarFormat.BorderColor.TintAndShade = 0
                
                .AxisPosition = xlDataBarAxisAutomatic
                .AxisColor.Color = 0
                .AxisColor.TintAndShade = 0
            End With
        
        End If
        
        If getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnHide) = cXmlAttributeValueYes Then
            aFormatRange.EntireColumn.Hidden = True
        ElseIf getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnHide) = cXmlAttributeValueNo Then
            aFormatRange.EntireColumn.Hidden = False
        End If
        
    End If

Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "TableFormatColumn", Err.Description & aColumnName
    Resume Next
    GoTo Finally

NoFieldFound:
    'addRow_Log logError, "TableFormatColumn", "Field: " & aColumnName & " was not found for dataset " & TableName & Err.Description
    'Resume Next
    GoTo Finally
End Sub

Sub TableFormatColumnPost(ColumnFormatNode As IXMLDOMNode, TableName As String)
    On Error GoTo ErrorHandler
    Dim aColumnName As String
    Dim aFormatRange As Range
    Dim aColumnWidth As String
    
    If getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnDelete) = cXmlAttributeValueYes Then
        Exit Sub
    End If
    
    aColumnName = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameName)
    On Error GoTo NoFieldFound
    Set aFormatRange = Range(TableName & "[" & getEscapedFieldName(aColumnName) & "]")
    On Error GoTo ErrorHandler

    If Not aFormatRange Is Nothing Then
        If getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnHide) = cXmlAttributeValueYes Then
            aFormatRange.EntireColumn.Hidden = True
            Exit Sub
        End If
        
        aColumnWidth = getXmlAttributeFromNode(ColumnFormatNode, cXmlAttributeNameColumnWidth)
        If aColumnWidth <> "" Then
            If aColumnWidth = cXmlAttributeNameColumnWidthAuto Then
                aFormatRange.EntireColumn.AutoFit
            ElseIf IsNumeric(aColumnWidth) Then
                aFormatRange.ColumnWidth = CDec(aColumnWidth)
            End If
        End If
    End If

Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "TableFormatColumnPost", Err.Description & aColumnName
    Resume Next
    GoTo Finally

NoFieldFound:
    'addRow_Log logError, "TableFormatColumnPost", "Field " & aColumnName & " was not found for dataset " & TableName & Err.Description
    GoTo Finally
End Sub

Sub TableFormatData(pFormatNode As IXMLDOMNode, TargetRange As Range, TableName As String)
    On Error GoTo ErrorHandler
    Dim aWs As Worksheet
    Dim aColumnFormat As IXMLDOMNode
    Dim aCellAddress As String
    Dim aTable As ListObject
    Dim aColumn As ListColumn
    'Dim aGlobalColumnConfig As IXMLDOMNode
    
    If pFormatNode Is Nothing Or TargetRange Is Nothing Then
        Exit Sub
    End If
    
    If getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractAllColumnsAutofit) = cXmlAttributeValueYes Then
        TargetRange.EntireColumn.AutoFit
    End If
    If getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractRowHeight) <> "" Then
        TargetRange.RowHeight = getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractRowHeight)
    ElseIf getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractAllRowsAutofit) = cXmlAttributeValueYes Then
        TargetRange.EntireRow.AutoFit
    End If
    
        
    TableSortDataByColumn TargetRange, TableName, getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractOrderBy)
    
    For Each aColumnFormat In pFormatNode.ChildNodes
        If aColumnFormat.nodeName = cXmlColumnNodeName Then
            TableFormatColumn aColumnFormat, TableName
        End If
    Next
    
    Set aTable = GetTableObject(TargetRange.Worksheet, TableName)
    If Not aTable Is Nothing Then
        For Each aColumn In aTable.ListColumns
            Set aColumnFormat = getXmlGlobalColumnConfigByPath(getGlobalConfigColumnByName, aColumn.Name)
            If Not aColumnFormat Is Nothing Then
                TableFormatColumn aColumnFormat, TableName
            End If
        Next
    End If
    
    Set aWs = TargetRange.Worksheet
    WorksheetHideGrid aWs, getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractHideGrid) = cXmlAttributeValueYes
    WorksheetHideHeading aWs, getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractHideHeading) = cXmlAttributeValueYes
    
    aCellAddress = getXmlAttributeFromNode(pFormatNode, cXmlAttributeNameExtractFreeze)
    If aCellAddress <> "" Then
        FreezePane aWs, aCellAddress
    End If
    
    For Each aColumnFormat In pFormatNode.ChildNodes
        If aColumnFormat.nodeName = cXmlColumnNodeName Then
            TableFormatColumnPost aColumnFormat, TableName
        End If
    Next
    
    
Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "TableFormatData", Err.Description
    'Resume Next
    GoTo Finally

End Sub

Sub SetActiveWorksheet(ActivateWorksheet As Worksheet)
    If Not ActivateWorksheet Is Nothing And Application.ActiveSheet.Name <> ActivateWorksheet.Name Then
        ActivateWorksheet.Activate
    End If
End Sub

Sub FreezePane(Ws As Worksheet, FreezeAddress As String)
    Dim aActiveCell As Range
    If Not Ws Is Nothing And FreezeAddress <> "" Then
        Set aActiveCell = Application.ActiveCell
        SetActiveWorksheet Ws
        ActiveWindow.FreezePanes = False
        On Error GoTo ErrorWrongAddress
        Ws.Range(FreezeAddress).Select
        On Error GoTo ErrorHandler
        ActiveWindow.FreezePanes = True
    End If
Finally:
    Exit Sub

ErrorHandler:
    addRow_Log logError, "FreezePane", Err.Description
    GoTo Finally

ErrorWrongAddress:
    addRow_Log logError, "FreezePane", "Wrong address to Freeze - " & FreezeAddress & " on " & Ws.Name & Err.Description
    GoTo Finally
End Sub


Function GetCorrectFormatNode(OriginalNode As IXMLDOMNode, RecordSet As ADODB.RecordSet, Optional FormatNode As IXMLDOMNode) As IXMLDOMNode
    On Error GoTo ErrorHandler
    Dim aResult As IXMLDOMNode, aChild As IXMLDOMNode, aFieldName As String
    Dim aCounter As Integer
    
    If OriginalNode Is Nothing Or RecordSet Is Nothing Then
        Exit Function
    End If
    
    If Not FormatNode Is Nothing Then
        aCounter = 0
        Set aResult = FormatNode
        For Each aChild In FormatNode.ChildNodes
            If aChild.nodeName = cXmlColumnNodeName Then
                aFieldName = getXmlAttributeFromNode(aChild, cXmlAttributeNameName)
                On Error GoTo ErrorFieldNotFound
                If RecordSet.Fields(aFieldName).Name = aFieldName Then
                    aCounter = aCounter + 1
                End If
                On Error GoTo ErrorHandler

                If aCounter > 10 Then
                    Set GetCorrectFormatNode = FormatNode
                    Exit Function
                End If
            End If
        Next
    End If
    
Finally:
    Set GetCorrectFormatNode = aResult
    Exit Function

ErrorFieldNotFound:
    addRow_Log logInfo, "XML Field " & aFieldName & " was not found in " & getXmlAttributeFromNode(FormatNode, cXmlAttributeNameName), Err.Description
    Resume Next
    
ErrorHandler:
    addRow_Log logError, "GetCorrectFormatNode " & aFieldName, Err.Description
    GoTo Finally

ErrorWrongAddress:
    addRow_Log logError, "FreezeGetCorrectFormatNodePane " & aFieldName, Err.Description
    GoTo Finally

End Function


Function DatasetByNameToWorksheet(Wb As Workbook, DataSetName As String)
    On Error GoTo ErrorHandler
    Dim aStatusBar As String, aSQL As String
    Dim aExecStart As Date, aExecCompleted As Date, aExtractCompleted As Date
    Dim aRsCount As Integer
    Dim aWs As Worksheet
    Dim aRs As ADODB.RecordSet
    Dim aNode As IXMLDOMNode, aFormatNode As IXMLDOMNode
    Dim aActiveWorksheetName As String, aWsName As String, aActiveWs As Worksheet
    Dim aTableRange As Range
    Dim aShowAllDataset As String, aShowEverything As String
    
    aRsCount = 1
    Set aNode = getXmlQueryNode(DataSetName)
    aShowAllDataset = getXmlAttributeFromNode(aNode, cXmlAttributeQueryShowAllDataSet)
    aShowEverything = getXmlAttributeFromNode(aNode, cXmlAttributeQueryShowEverything)
    aActiveWorksheetName = getXmlAttributeFromNode(aNode, cXmlAttributeNameExtractSetActive)
    aSQL = getXmlQueryBody(DataSetName)
    aExecStart = Now()
    Set aRs = execAsRecordSet(aSQL)
    aExecCompleted = Now()
    While Not aRs Is Nothing
        If aRs.Fields.Count > 0 Then
            Set aFormatNode = getXmlQueryFormatNode(DataSetName, DataSetName & "_" & Format(aRsCount))
            Set aFormatNode = GetCorrectFormatNode(aNode, aRs, aFormatNode)
            If aShowAllDataset = cXmlAttributeValueYes _
            Or aShowEverything = cXmlAttributeValueYes _
            Or Not getXmlAttributeFromNode(aFormatNode, cXmlAttributeNameExtractSkip) = cXmlAttributeValueYes Then
                If aRs.RecordCount > 0 _
                Or (aRs.RecordCount = 0 And getXmlAttributeFromNode(aFormatNode, cXmlAttributeNameExtractHideIfEmpty) = cXmlAttributeValueYes) Then
                    aWsName = getXmlAttributeFromNode(aFormatNode, cXmlAttributeNameExtractWorksheetName, DataSetName & "_" & Format(aRsCount))
                    Set aWs = prepareWorksheet(Wb, aWsName)
                    If Not aWs Is Nothing And aActiveWorksheetName = aWs.Name Then
                        Set aActiveWs = aWs
                    End If
                    Set aTableRange = TableExtract(aWs.Range(getXmlAttributeFromNode(aFormatNode, cXmlAttributeNameExtractAddress, cXmlAttributeNameExtractAddressDefault)), aRs, aWsName)
                    TableFormatData aFormatNode, aTableRange, aWsName
                End If
            End If
        End If
        Set aRs = aRs.NextRecordset
        If Not aRs Is Nothing Then
            aRsCount = aRsCount + 1
        End If
    Wend
    aExtractCompleted = Now()
    If Not aActiveWs Is Nothing Then
        aActiveWs.Activate
    End If
    If Not aRs Is Nothing Then
        aStatusBar = GetQueryStats(aExecStart, aExecCompleted, aExecCompleted, aExtractCompleted, aRs.RecordCount)
    End If
    setStatusBarMessage aStatusBar

Finally:
    Exit Function

ErrorHandler:
    addRow_Log logError, "DatasetByNameToWorksheet", Err.Description
    GoTo Finally
End Function


Function ExtractDatasetToWorksheetByXMLNode(pWs As Worksheet, pDatasetNode As IXMLDOMNode)
    On Error GoTo ErrorHandler
    If Not pWs Is Nothing And Not pDatasetNode Is Nothing Then
        
    End If
Finally:
    Exit Function

ErrorHandler:
    addRow_Log logError, "ExtractDatasetToWorksheet", Err.Description
    GoTo Finally

End Function


