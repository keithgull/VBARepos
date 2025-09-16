VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoggerSheetOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: SheetOutput
Option Explicit
Implements iLoggerOutput

Private targetSheet As Worksheet
Private currentRow As Long
Private currentCol As Long

Public Sub Init(ByVal ws As Worksheet, ByVal logStartRow As Long, ByVal logStartCol As Long)
    Set targetSheet = ws
    currentRow = logStartRow
    currentCol = logStartCol
End Sub

Private Sub ILoggerOutput_OutputLog(ByVal message As String)
    targetSheet.Cells(currentRow, currentCol).Value = message
    currentRow = currentRow + 1
End Sub

