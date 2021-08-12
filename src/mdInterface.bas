Attribute VB_Name = "mdInterface"
'@Folder("Main")
Option Explicit


Sub intf_ClearLog()
    Init
    ClearLogTable
    Finalize
End Sub


Sub intf_CleanWorkbook()
    Init
    CleanWorkbook
    Finalize
End Sub


Sub iTickCheckShape()
    TickCheckShape
End Sub


Sub iRadioCheckShape()
    RadioCheckShape
End Sub


Sub intf_InstallBlitz()
    InstallBlitz
End Sub


Sub intf_BlitzAllChecked()
    Init
    BlitzAllChecked
    Finalize
End Sub


Sub intf_PerformanceTest()
    Init
    PerformanceTest
    Finalize
End Sub


Sub intf_Blitz()
    Init
    sp_Blitz
    Finalize
End Sub


Sub intf_TickAllBlitz()
    Init
    TickAllBlitz
    Finalize
End Sub




