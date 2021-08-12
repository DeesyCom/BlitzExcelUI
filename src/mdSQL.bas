Attribute VB_Name = "mdSQL"
'@Folder("Database")
Option Private Module
Option Explicit

Private gConn As ADODB.connection
Private cmdADOCreate As ADODB.Command
Private cmdADORead As ADODB.Command
Private cmdADOUpdate As ADODB.Command
Private cmdADODelete As ADODB.Command
Private cmdADOLongRead As ADODB.Command

Public Const gParamServer = "{Server}"
Public Const gParamDatabase = "{Database}"
Public Const gParamApplication = "{Application}"
Public Const cCmdTimeout = 60 * 60 ' 1 hour

Public Const gSQLServerConnectionLine = "Provider=sqloledb;Data Source=" & gParamServer & ";Initial Catalog=" & gParamDatabase & ";Integrated Security=SSPI;Application Name=" & gParamApplication & ";"
Public Const gSQLServer = "DESKTOP-A7IIVA2"

Private gServerName As String
Private gDatabaseName As String
Private gSchemaName As String


'------------------------------------------------
Sub closeAndNullCmd(ByRef aCmd As ADODB.Command)
    If Not aCmd Is Nothing Then
        Set aCmd = Nothing
    End If
End Sub


Sub closeConnection()
    closeAndNullCmd cmdADOCreate
    closeAndNullCmd cmdADORead
    closeAndNullCmd cmdADOUpdate
    closeAndNullCmd cmdADODelete
    closeAndNullCmd cmdADOLongRead

    If Not gConn Is Nothing Then
        gConn.Close
        Set gConn = Nothing
    End If
End Sub


Function getSQLServerName()
    If gServerName = "" Then
        gServerName = getXmlAttributeFromNode(getXmlConfigNode(), cXmlAttributeNameServer)
    End If
    getSQLServerName = gServerName
End Function


Function getSQLDatabaseName()
    If gDatabaseName = "" Then
        gDatabaseName = getXmlAttributeFromNode(getXmlConfigNode(), cXmlAttributeNameDatabase)
    End If
    getSQLDatabaseName = gDatabaseName
End Function


Function getSQLSchemaName()
    If gSchemaName = "" Then
        gSchemaName = getXmlAttributeFromNode(getXmlConfigNode(), cXmlAttributeNameSchema)
    End If
    getSQLSchemaName = gSchemaName
End Function


Function getConnectionString() As String
    Dim aConfigNode As IXMLDOMNode
    Dim aConnString As String
    Dim aDatabase As String
    Dim aSchema As String
    
    Set aConfigNode = getXmlConfigNode()
    aDatabase = getSQLDatabaseName()
    aSchema = getSQLSchemaName()
    
    aConnString = gSQLServerConnectionLine
    aConnString = Replace(Replace(aConnString, gParamServer, getSQLServerName()), gParamDatabase, aDatabase)
    
    getConnectionString = aConnString
End Function


Function getConnection() As ADODB.connection
    On Error GoTo ErrorHandler
    
    Dim aConn As ADODB.connection
    Set aConn = New ADODB.connection
    aConn.ConnectionString = getConnectionString
    aConn.ConnectionTimeout = cCmdTimeout
    aConn.CursorLocation = adUseClient
    aConn.Open
    Set getConnection = aConn
Finally:
    Exit Function
ErrorHandler:
    Dim adoErr As ADODB.Error
    Dim aIsAdoError As Boolean
    aIsAdoError = False
    If Not aConn Is Nothing Then
        For Each adoErr In aConn.Errors
            aIsAdoError = True
            addRow_Log logError, "getConnection", adoErr & vbCr & Err.Description
        Next adoErr
    End If
    
    If Not aIsAdoError Then
        addRow_Log logError, "getConnection", Err.Description
    End If
    GoTo Finally
End Function


Sub releaseADOCommand(ByRef pCmd As ADODB.Command)
    If Not pCmd Is Nothing Then
        Set pCmd = Nothing
    End If
End Sub


Function getADOCmd(pType As String) As ADODB.Command
    On Error GoTo ErrorHandler
    Dim aConn As ADODB.connection
    Dim aADOCmd As ADODB.Command
    
    Set aConn = getConnection()
    
    If pType = sqlTypeCreate Then       '--- Create
        If cmdADOCreate Is Nothing Then
            Set cmdADOCreate = New ADODB.Command
            cmdADOCreate.CommandTimeout = cCmdTimeout
            Set cmdADOCreate.ActiveConnection = aConn
        End If
        Set aADOCmd = cmdADOCreate
    ElseIf pType = sqlTypeRead Then       '--- Read
        If cmdADORead Is Nothing Then
            Set cmdADORead = New ADODB.Command
            cmdADORead.CommandTimeout = cCmdTimeout
            Set cmdADORead.ActiveConnection = aConn
        End If
        Set aADOCmd = cmdADORead
    ElseIf pType = sqlTypeUpdate Then       '--- Update
        If cmdADOUpdate Is Nothing Then
            Set cmdADOUpdate = New ADODB.Command
            cmdADOUpdate.CommandTimeout = cCmdTimeout
            Set cmdADOUpdate.ActiveConnection = aConn
        End If
        Set aADOCmd = cmdADOUpdate
    ElseIf pType = sqlTypeDelete Then       '--- Delete
        If cmdADODelete Is Nothing Then
            Set cmdADODelete = New ADODB.Command
            cmdADODelete.CommandTimeout = cCmdTimeout
            Set cmdADODelete.ActiveConnection = aConn
        End If
        Set aADOCmd = cmdADODelete
    Else                                    '--- Long Read
        If cmdADOCreate Is Nothing Then
            Set cmdADOCreate = New ADODB.Command
            cmdADOCreate.CommandTimeout = cCmdTimeout
            Set cmdADOCreate.ActiveConnection = aConn
        End If
        Set aADOCmd = cmdADOCreate
    End If
    

    Set getADOCmd = aADOCmd
Finally:
    Exit Function
ErrorHandler:
    Dim aAdoErr As ADODB.Error
    For Each aAdoErr In gConn.Errors
        addRow_Log logError, "getADOCmd", "ADO Error: " + aAdoErr + vbCr
    Next aAdoErr
    addRow_Log logError, "getADOCmd", "Failed with error#: " & CStr(Err.Number) & " and Description: " & Err.Description
    MsgBox "Connection to Database failed"
    GoTo Finally
End Function


Function ExecSqlByType(pSQL As String, pType As String) As Boolean
    On Error GoTo ErrorHandler
    Dim aCmd As ADODB.Command
    Set aCmd = getADOCmd(pType)
    ExecSqlByType = False
    aCmd.CommandTimeout = cCmdTimeout
    aCmd.CommandText = pSQL

    ExecSqlByType = execQuery(aCmd, "ExecSql")
Finally:
    releaseADOCommand aCmd
    Exit Function
RecordNotFound:
ErrorHandler:
    addRow_Log logError, "ExecSql", Err.Description & " ErrNumber: " & CStr(Err.Number)
    GoTo Finally
End Function



'Function ExecSql(pSql As String) As Boolean
'    On Error GoTo ErrorHandler
'
'    Dim aCmd As ADODB.Command
'    Set aCmd = getADOCmd()
'    ExecSql = False
'    aCmd.CommandTimeout = cCmdTimeout
'    aCmd.CommandText = pSql
'
'    ExecSql = execQuery(aCmd, "ExecSql")
'Finally:
'    releaseADOCommand aCmd
'    Exit Function
'RecordNotFound:
'ErrorHandler:
'    addRow_Log logError, "ExecSql", Err.Description & " ErrNumber: "
'    GoTo Finally
'End Function

Sub SQLExecuteCreate(pSQL As String)
    On Error GoTo ErrorHandler
    
    If ExecSqlByType(pSQL, sqlTypeCreate) Then
        MsgBox ("All done!")
    End If
    
Finally:
    Exit Sub
RecordNotFound:
ErrorHandler:
    addRow_Log logError, "SQLExecuteCreate", Err.Description & " ErrNumber: " & CStr(Err.Number) & pSQL
    GoTo Finally
End Sub


Function execQuery(ByRef pCmd As ADODB.Command, Optional pFunctionName As String = "", Optional pFunctionalPartName As String = "", Optional pSuppressErrors As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    execQuery = False
    If Not pCmd Is Nothing Then
        'logText pCmd.CommandText, "Function: " & pFunctionName & ", Function part: " & pFunctionalPartName
        pCmd.Execute
        execQuery = True
    End If
Finally:
    Exit Function
RecordNotFound:
ErrorHandler:
    If pSuppressErrors Then
        GoTo Finally
    End If
    Dim aIsAdoError As Boolean
    Dim adoErr As ADODB.Error
    Dim aSQL As String
    aIsAdoError = False
    If Not pCmd Is Nothing Then
        aSQL = pCmd.CommandText
        If Not pCmd.ActiveConnection Is Nothing Then
            If pCmd.ActiveConnection.State = 1 Then
                For Each adoErr In pCmd.ActiveConnection.Errors
                    aIsAdoError = True
                    addRow_Log logError, pFunctionName, adoErr & vbCr & Err.Description & vbCrLf & aSQL & vbCrLf & pFunctionalPartName
                Next adoErr
            End If
        End If
    End If
    
    If Not aIsAdoError Then
        addRow_Log logError, pFunctionName, pFunctionalPartName & vbCr & aSQL & Err.Description & " ErrNumber: " & CStr(Err.Number)
    End If
    GoTo Finally
End Function


Function execDMOQuery(pQueryPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim aSqlServer, aDatabase As Object
    
    Set aSqlServer = CreateObject("SQLDMO.SQLServer")
    Set aDatabase = CreateObject("SQLDMO.Database")
    
    aSqlServer.LoginSecure = True
    aSqlServer.Connect getSQLServerName()
    
    Set aDatabase = aSqlServer.Databases(getSQLDatabaseName(), getSQLSchemaName())
    aDatabase.ExecuteImmediate (pQueryPath)
    
Finally:
    
    Exit Function
RecordNotFound:
ErrorHandler:
    addRow_Log logError, "execDMOQuery", Err.Description & " ErrNumber: " & CStr(Err.Number) & vbCr & pQueryPath

    GoTo Finally
    
End Function

Function execAsRecordSet(ByRef pSQL As String) As ADODB.RecordSet
    On Error GoTo ErrorHandler
    
    Dim aRs As ADODB.RecordSet
    Dim aCmd As ADODB.Command
    
    Set aCmd = New ADODB.Command
    aCmd.CommandTimeout = cCmdTimeout
    aCmd.CommandText = pSQL
    Set aCmd.ActiveConnection = getConnection()
    
    Set aRs = aCmd.Execute
    
    Set execAsRecordSet = aRs
Finally:
    
    Exit Function
ErrorHandler:
    addRow_Log logError, "execAsRecordSet", Err.Description & " ErrNumber: " & CStr(Err.Number) & vbCr & pSQL

    GoTo Finally
    
End Function

Function execAsRecordSet2(ByRef pSQL As String) As ADODB.RecordSet
    On Error GoTo ErrorHandler
    
    Dim aRs As ADODB.RecordSet
    
    Set aRs = New ADODB.RecordSet
    With aRs
        .CursorLocation = adUseServer
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .ActiveConnection = getConnection()
        .Open pSQL
    End With
Finally:
    
    Exit Function
ErrorHandler:
    addRow_Log logError, "execAsRecordSet", Err.Description & " ErrNumber: " & CStr(Err.Number) & vbCr & pSQL

    GoTo Finally
    
End Function

