Attribute VB_Name = "mdXML"
'@Folder("Main")
Option Private Module
Option Explicit

Private gXmlConfigDoc As MSXML2.DOMDocument60
Private gXmlConfigNode As IXMLDOMNode
Private gXmlGlobalConfigNode As IXMLDOMNode

Private gSchema As String
Private gSchema_Template As String

Private lConfigFileName As String
'--------------------------------------------
Function getXmlConfDoc() As DOMDocument60
    lConfigFileName = ThisWorkbook.Path + "\" + gXMLConfigFileName '<-- replace with argument from command line
    If gXmlConfigDoc Is Nothing Then
        Set gXmlConfigDoc = New DOMDocument60
        If Not gXmlConfigDoc.Load(lConfigFileName) Then
            MsgBox Replace("Huston, we have a problem with {filename} config", "{filename}", lConfigFileName)
            End
        Else
        End If
    End If
    Set getXmlConfDoc = gXmlConfigDoc
End Function


Function getXmlNode(pXmlPath As String) As IXMLDOMNode
    Dim aConfDoc As DOMDocument60
    Set aConfDoc = getXmlConfDoc()
    If Not aConfDoc Is Nothing Then
        Set getXmlNode = aConfDoc.SelectSingleNode(pXmlPath)
    End If
End Function


Function getXmlConfigNode() As IXMLDOMNode
    If gXmlConfigNode Is Nothing Then
        Set gXmlConfigNode = getXmlNode(gXMLGetConfigNode)
        Set gXmlGlobalConfigNode = getXmlNode(getGlobalConfigByName)
    End If
    Set getXmlConfigNode = gXmlConfigNode
End Function


Sub closeXML()
    If Not gXmlConfigNode Is Nothing Then
        Set gXmlConfigNode = Nothing
    End If
    
    If Not gXmlConfigDoc Is Nothing Then
        Set gXmlConfigDoc = Nothing
    End If
End Sub


Function getXmlQuery(pXPath As String) As String
    getXmlQuery = ""
    Dim aConfDoc As DOMDocument60
    Dim aNode As IXMLDOMNode
    Dim aAttrs As IXMLDOMNamedNodeMap
    Set aConfDoc = getXmlConfDoc()
    Set aNode = aConfDoc.SelectSingleNode(pXPath)
    If aNode Is Nothing Then
        addRow_Log logError, "getXmlQuery", "XPAth " + pXPath + " was not found " & Err.Description & " ErrNumber: " & CStr(Err.Number)
        'MsgBox ("XPath " + pXPath + " was not found")
    Else
        getXmlQuery = Replace(Trim(aNode.Text), getXMLSQLSchema_Template(), getXMLSQLSchema())
        Set aAttrs = aNode.Attributes
        gNoLog = False
        If Not aAttrs.getNamedItem("nolog") Is Nothing Then
            gNoLog = aAttrs.getNamedItem("nolog").Text = "true"
        End If
    End If
    Set aNode = Nothing
    Set aAttrs = Nothing
    Set aConfDoc = Nothing
End Function


Function getXmlQueryByPath(pXPath As String, pQueryName As String) As String
    getXmlQueryByPath = getXmlQuery(Replace(pXPath, "{QueryName}", pQueryName))
End Function


Function getXmlGlobalColumnConfigByPath(XPath As String, ColumnName As String) As IXMLDOMNode
    Set getXmlGlobalColumnConfigByPath = getXmlNode(Replace(XPath, "{ColumnName}", ColumnName))
End Function


Function getXmlQueryBody(pQueryName As String) As String
    getXmlQueryBody = getXmlQuery(Replace(getQueryBodyByName, "{QueryName}", pQueryName))
End Function


Function getXmlQueryFormatNode(pQueryName As String, DataSetName As String) As IXMLDOMNode
    Set getXmlQueryFormatNode = getXmlNode(Replace(Replace(getQueryFormatNodeByName, "{QueryName}", pQueryName), "{DatasetName}", DataSetName))
End Function


Function getXmlQueryNode(pQueryName As String) As IXMLDOMNode
    Set getXmlQueryNode = getXmlNode(Replace(getQueryByName, "{QueryName}", pQueryName))
End Function


Function getXMLSQLSchema()
    If gSchema = "" Then
        gSchema = getXmlAttributeFromNode(getXmlConfigNode(), cXmlAttributeNameSchema)
    End If
    getXMLSQLSchema = gSchema
End Function


Function getXMLSQLSchema_Template()
    If gSchema_Template = "" Then
        gSchema_Template = getXmlAttributeFromNode(getXmlConfigNode(), cXmlAttributeNameSchemaTemplate)
    End If
    getXMLSQLSchema_Template = gSchema_Template
End Function


Function getXmlAttributeFromNode(ByRef Node As IXMLDOMNode, AttributeName As String, Optional DefaultValue As String = "") As String
    Dim aAttrs As IXMLDOMNamedNodeMap
    On Error GoTo ErrorHandler
    
    If Not Node Is Nothing Then
        Set aAttrs = Node.Attributes
        If Not aAttrs Is Nothing Then
            If Not aAttrs.getNamedItem(AttributeName) Is Nothing Then
                getXmlAttributeFromNode = Trim(aAttrs.getNamedItem(AttributeName).Text)
            Else
                getXmlAttributeFromNode = getXmlAttributeFromNode(Node.ParentNode, AttributeName, DefaultValue)
            End If
        Else
            getXmlAttributeFromNode = getXmlAttributeFromNode(Node.ParentNode, AttributeName, DefaultValue)
        End If
    Else
        getXmlAttributeFromNode = DefaultValue
    End If
Finally:
    Exit Function

ErrorHandler:
    addRow_Log logError, "getXmlAttributeFromNode: " & AttributeName, Err.Description
    GoTo Finally
End Function

Function getXmlAttributeFromNodeAsDouble(ByRef Node As IXMLDOMNode, AttributeName As String, Optional DefaultValue As Double = 0#) As Double
    Dim aValue As String
    aValue = getXmlAttributeFromNode(Node, AttributeName, CStr(DefaultValue))
    If IsNumeric(aValue) Then
        getXmlAttributeFromNodeAsDouble = CDec(aValue)
    Else
        getXmlAttributeFromNodeAsDouble = DefaultValue
    End If
End Function

