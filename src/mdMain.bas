Attribute VB_Name = "mdMain"
'@Folder("Main")
Option Private Module
Option Explicit


Sub InstallBlitz()
    Init
    If execDMOQuery(getBlitzScriptPath()) Then
        MsgBox "All good"
    End If
    
    
    'aSqlCode = GetTextFromFile(getBlitzScriptPath())
    'execDMOQuery (getBlitzScriptPath())
    'MsgBox Len(aSqlCode)
    'If aSqlCode <> "" Then
    '    SQLExecuteCreate aSqlCode
    'End If
    
    Finalize
End Sub
