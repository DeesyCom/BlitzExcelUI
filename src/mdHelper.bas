Attribute VB_Name = "mdHelper"
'@Folder("Main")
Option Private Module
Option Explicit


Function CheckShape_IsChecked(WorkSheetName As String, ShapeName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim aWs As Worksheet, result As Boolean
    result = False
    
    If WorkSheetName = Empty Or ShapeName = Empty Then
        CheckShape_IsChecked = False
        Exit Function
    End If
    Set aWs = ThisWorkbook.Worksheets(WorkSheetName)
    If Not aWs Is Nothing Then
        result = ThisWorkbook.Worksheets(WorkSheetName).Shapes(ShapeName).ShapeStyle = cCheckShapeStyleActive
    End If
    
Finally:
    CheckShape_IsChecked = result
    Exit Function

ErrorHandler:
    addRow_Log logError, "CheckShape_IsChecked", "WorkSheetName: " & WorkSheetName & "; ShapeName: " & ShapeName & Err.Description
    GoTo Finally
End Function


Function CheckBox_IsChecked(WorkSheetName As String, ControlName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim aWs As Worksheet
    Dim aCheckBox As CheckBox
    Set aWs = ThisWorkbook.Worksheets(WorkSheetName)
    If Not aWs Is Nothing Then
        For Each aCheckBox In aWs.CheckBoxes
            If aCheckBox.Name = ControlName Then
                CheckBox_IsChecked = aCheckBox.Value = 1
                Exit Function
            End If
        Next
    End If

Finally:
    Exit Function

ErrorHandler:
    addRow_Log logError, "CheckBox_IsChecked", "Worksheet: " & WorkSheetName & "; ControlName: " & ControlName & " " & Err.Description
    GoTo Finally
End Function


