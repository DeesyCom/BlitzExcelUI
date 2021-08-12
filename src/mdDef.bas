Attribute VB_Name = "mdDef"
'@Folder("Define")
Option Private Module
Option Explicit


Global gNoLog As Boolean

Global Const SHIFT_KEY = 16
Global Const CTRL_KEY = 17
Global Const ALT_KEY = 18


Global Const gXMLConfigFileName As String = "config.xml"
Global Const gXMLGetConfigNode As String = "//configuration"
Global Const getGlobalConfigByName As String = gXMLGetConfigNode & "/queries/global"
Global Const getGlobalConfigColumnByName As String = gXMLGetConfigNode & "/queries/global/column[@name='{ColumnName}']"
Global Const getQueryByName As String = gXMLGetConfigNode & "/queries/query[@name='{QueryName}']"
Global Const getQueryBodyByName As String = gXMLGetConfigNode & "/queries/query[@name='{QueryName}']/body"
Global Const getQueryFormatNodeByName As String = gXMLGetConfigNode & "/queries/query[@name='{QueryName}']/format[@name='{DatasetName}']"
Global Const getQueryFormatColumnNodeByName As String = gXMLGetConfigNode & "/queries/query[@name='{QueryName}']/format[@name='{DatasetName}']/column[@name='{ColumnName}']"
Global Const getQueryFormatNodeByName2 As String = gXMLGetConfigNode & "/queries/query[@name='{QueryName}']/format"

Public Const tblLogName As String = "Log"
Public Const wsLogName As String = "Log"
Public Const wsConfigName As String = "Config"
Public Const cCellServerName = "vServer"

Public Const cellLatestId = "cellLatestId"

Public Const logError = "Error"
Public Const logInfo = "Info"


Public Const cXmlColumnNodeName = "column"

Public Const cXmlAttributeNameServer = "server"
Public Const cXmlAttributeNameDatabase = "database"
Public Const cXmlAttributeNameSchema = "schema"
Public Const cXmlAttributeNameSchemaTemplate = "schema_template"
Public Const cXmlAttribute_CellShift = "cell_shift"
Public Const cXmlAttribute_HeaderOrientation = "header_orientation"
Public Const cXmlAttributeNameActiveCell = "active_cell"

Public Const cXmlAttributeNameName = "name"
Public Const cXmlAttributeValueYes = "yes"
Public Const cXmlAttributeValueNo = "no"
Public Const cXmlAttributeQueryShowAllDataSet = "show_all_dataset"
Public Const cXmlAttributeQueryShowEverything = "show_everything"
Public Const cXmlAttributeNameExtractSetActive = "set_active"
Public Const cXmlAttributeNameExtractStyle = "style"
Public Const cXmlAttributeNameExtractAddress = "extract_address"
Public Const cXmlAttributeNameExtractAddressDefault = "A1"
Public Const cXmlAttributeNameExtractWorksheetName = "worksheet_name"
Public Const cXmlAttributeNameExtractSkip = "skip"
Public Const cXmlAttributeNameExtractFreeze = "freeze"
Public Const cXmlAttributeNameExtractHideGrid = "hide_grid"
Public Const cXmlAttributeNameExtractHideHeading = "hide_heading"
Public Const cXmlAttributeNameExtractAllColumnsAutofit = "all_columns_autofit"
Public Const cXmlAttributeNameExtractAllRowsAutofit = "all_rows_autofit"
Public Const cXmlAttributeNameExtractOrderBy = "order_by"
Public Const cXmlAttributeNameExtractRowHeight = "all_rows_height"
Public Const cXmlAttributeNameExtractHideIfEmpty = "hide_if_empty"
Public Const cXmlAttributeNameColumnFormat = "format"
Public Const cXmlAttributeNameColumnDatabar = "databar"
Public Const cXmlAttributeNameColumnBarColor = "bar_color"
Public Const cXmlAttributeNameColumnDelete = "delete"
Public Const cXmlAttributeNameColumnHide = "hide"
Public Const cXmlAttributeNameColumnWidth = "width"
Public Const cXmlAttributeNameColumnWidthAuto = "auto"
Public Const cXmlAttributeNameColumnHRefFrom = "href_from"
Public Const cXmlAttributeNameColumnBarColorDefault = 5920255
Public Const cXmlAttributeNameColumnHorizontalAlignment = "halign"
Public Const cXmlAttributeNameColumnHorizontalAlignmentLeft = "left"
Public Const cXmlAttributeNameColumnHorizontalAlignmentCenter = "center"
Public Const cXmlAttributeNameColumnHorizontalAlignmentRight = "right"
Public Const cXmlAttributeNameColumnHorizontalAlignmentGeneral = "general"
Public Const cXmlAttributeNameColumnVerticalAlignment = "valign"
Public Const cXmlAttributeNameColumnVerticalAlignmentTop = "top"
Public Const cXmlAttributeNameColumnVerticalAlignmentCenter = "center"
Public Const cXmlAttributeNameColumnVerticalAlignmentBottom = "bottom"
Public Const cXmlAttributeNameColumnTextWrap = "text_wrap"
Public Const cXmlAttributeNameColumnColorScale = "colorscale"
Public Const cXmlAttributeNameColumnColorScaleGYR = "gyr"
Public Const cXmlAttributeNameColumnColorScaleRYG = "ryg"
Public Const cXmlAttributeNameColumnForceToNumber = "force_to_number"



Public Const sqlTypeCreate As String = "Create"
Public Const sqlTypeRead As String = "Read"
Public Const sqlTypeUpdate As String = "Update"
Public Const sqlTypeDelete As String = "Delete"
Public Const sqlTypeLongRead As String = "LongRead"

Public Const cbDedicatedFileName As String = "cbDedicatedFile"
Public Const wsControlCentreName As String = "ControlCentre"

Public Const wbNewBookTitlePrefix As String = "Blitz"
Public Const wbBlitzName As String = "sp_Blitz"
Public Const dsBlitzName As String = "sp_Blitz"
Public Const wbBlitzFirstName As String = "sp_BlitzFirst"
Public Const dsBlitzFirstName As String = "sp_BlitzFirst"
Public Const wbBlitzIndexName As String = "sp_BlitzIndex"
Public Const dsBlitzIndexName As String = "sp_BlitzIndex"
Public Const wbBlitzCacheName As String = "sp_BlitzCache"
Public Const dsBlitzCacheName As String = "sp_BlitzCache"
Public Const wbBlitzWhoName As String = "sp_BlitzWho"
Public Const dsBlitzWhoName As String = "sp_BlitzWho"

Public Const cQueryStatTemplateExecTime = "{exec_time}"
Public Const cQueryStatTemplateExtractTime = "{extract_time}"
Public Const cQueryStatTemplateRecordCount = "{record_count}"
Public Const cQueryStatTemplate = "Exec time: " & cQueryStatTemplateExecTime & " ; Extract time: " & cQueryStatTemplateExtractTime & "; Record count: " & cQueryStatTemplateRecordCount

Public Const cCheckShapeLikePrefix = "cs*"
Public Const cRadioShapeLikePrefix = "rs*"
Public Const cCheckShapeStyleActive = 41
Public Const cCheckShapeStyleNonActive = 39
Public Const cCheckShapeStyleRadioNonActive = 4

Public Const cOutputThisFile = "ThisFile"
Public Const cOutputNewFile = "NewFile"
Public Const cOutputIndividualFile = "Individual"

Public Const cInclude_sp_Blitz = "csBlitz"
Public Const cInclude_sp_BlitzFirst = "csBlitzFirst"
Public Const cInclude_sp_BlitzIndex = "csBlitzIndex"
Public Const cInclude_sp_BlitzCache = "csBlitzCache"
Public Const cInclude_sp_BlitzWho = "csBlitzWho"

