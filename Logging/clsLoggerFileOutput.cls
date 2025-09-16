VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoggerFileOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: FileOutput
Option Explicit
Implements iLoggerOutput

Private logFilePath As String

Public Sub Init(ByVal path As String, ByVal filename As String)
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    logFilePath = path & filename
End Sub

Private Sub ILoggerOutput_OutputLog(ByVal message As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open logFilePath For Append As #fnum
    Print #fnum, message
    Close #fnum
End Sub

