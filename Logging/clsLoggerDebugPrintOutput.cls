VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoggerDebugPrintOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: DebugOutput
Option Explicit
Implements iLoggerOutput

Private Sub ILoggerOutput_OutputLog(ByVal message As String)
    Debug.Print message
End Sub

