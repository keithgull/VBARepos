VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====== Class Module: Logger ==================
' VBA 汎用Logger
'  使い方
'    Dim log As New clsLogger
'
'    ' Enumを使って初期化
'    log.InitializeLogger LOGGER_TYPE_DEBUGPRINT, "", ""                        ' Debug.printを使ったログ
'    ' log.InitializeLogger LOGGER_TYPE_LOGFILE, "C:\Work", "vbalogout.log"     ' ログファイルへ出力するログ
'    ' log.InitializeLogger LOGGER_TYPE_LOGSHEET, "Sheet1", "C10"               ' シートに出力するログ
'
'    log.SetLogLevel LOGLEVEL_DEBUG, True                                       ' ログレベルの個別設定
'    log.SetLogLevelAll true, true, true, true, true                            ' ログレベルの一括設定（INFO、ERROR、WARN、DEBUG、TRACEの順で出力・非出力を設定）
'
'    log.InfoLog "開始"
'    log.DebugLog "デバッグ情報"
'

#If VBA7 Then
    Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private logInstance As clsLogger
Private outputHandler As iLoggerOutput
Private logLevels As Scripting.Dictionary
Private formatTemplate As String
Private useModuleName As Boolean
Private moduleNameLen As Integer

Private Sub Class_Initialize()
    'Set logInstance = GetLogger()
    Set logLevels = New Scripting.Dictionary
    logLevels.Add LOGLEVEL_INFO, True
    logLevels.Add LOGLEVEL_DEBUG, False
    logLevels.Add LOGLEVEL_WARN, True
    logLevels.Add LOGLEVEL_ERROR, True
    logLevels.Add LOGLEVEL_TRACE, False
    formatTemplate = "{time} {level} {module} {message}"
    useModuleName = False
End Sub

' シングルトン取得
Public Function GetLogger() As clsLogger
    If logInstance Is Nothing Then
        Set logInstance = New clsLogger
    End If
    Set GetLogger = logInstance
End Function

' 出力先の初期化
Public Sub InitializeLogger(ByVal loggerType As LOGGER_TYPE, ByVal param1 As String, ByVal param2 As String)
    Select Case loggerType
        Case LOGGER_TYPE_DEBUGPRINT
            Dim dlog As New clsLoggerDebugPrintOutput
            Set outputHandler = dlog

        Case LOGGER_TYPE_LOGFILE
            Dim flog As New clsLoggerFileOutput
            flog.Init param1, param2
            Set outputHandler = flog

        Case LOGGER_TYPE_LOGSHEET
            Dim slog As New clsLoggerSheetOutput
            Dim ws As Worksheet
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(param1)
            On Error GoTo 0
            If ws Is Nothing Then Err.Raise vbObjectError + 1, , "指定されたシートが存在しません: " & param1

            Dim col As Long, row As Long
            col = Range(param2).Column
            row = Range(param2).row
            slog.Init ws, col, row
            Set outputHandler = slog

        Case Else
            Err.Raise vbObjectError + 2, , "不明なログ種別: " & loggerType
    End Select
End Sub

Public Sub SetOutput(ByVal handler As iLoggerOutput)
    Set outputHandler = handler
End Sub

Public Sub SetLogLevel(ByVal level As String, ByVal enabled As Boolean)
    If logLevels.Exists(level) Then
        logLevels(level) = enabled
    End If
End Sub

Public Sub SetLogLevelAll(ByVal infoLvl As Boolean, ByVal errorLvl As Boolean, warnLvl As Boolean, debugLvl As Boolean, traceLvl As Boolean)
    SetLogLevel LOGLEVEL_INFO, infoLvl
    SetLogLevel LOGLEVEL_ERROR, errorLvl
    SetLogLevel LOGLEVEL_WARN, warnLvl
    SetLogLevel LOGLEVEL_DEBUG, debugLvl
    SetLogLevel LOGLEVEL_TRACE, traceLvl
End Sub

Public Sub SetUseModuleName(ByVal useFlag As Boolean, Optional nameLen As Integer = 0)
    useModuleName = useFlag
    moduleNameLen = nameLen
End Sub

Public Sub SetLogFormat(ByVal template As String)
    formatTemplate = template
End Sub

Public Sub log(ByVal level As String, ByVal moduleName As String, ByVal message As String)
    If Not logLevels.Exists(level) Or Not logLevels(level) Then
        Exit Sub
    End If

    Dim formatted As String
    formatted = Replace(formatTemplate, "{time}", Format(Now, "yyyy-mm-dd hh:nn:ss"))
    formatted = Replace(formatted, "{level}", level)
    formatted = IIf(useModuleName, Replace(formatted, "{module}", moduleName), "")
    formatted = Replace(formatted, "{message}", message)

    If Not outputHandler Is Nothing Then
        outputHandler.OutputLog formatted
    End If
End Sub

' 各ログレベル用メソッド
Public Sub InfoLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_INFO, moduleName, message
End Sub

Public Sub DebugLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_DEBUG, moduleName, message
End Sub

Public Sub WarnLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_WARN, moduleName, message
End Sub

Public Sub ErrorLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_ERROR, moduleName, message
End Sub

Public Sub TraceLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_TRACE, moduleName, message
End Sub

Public Sub InfoLog(ByVal message As String)
    OutputLog LOGLEVEL_INFO, "", message
End Sub

Public Sub DebugLog(ByVal message As String)
    OutputLog LOGLEVEL_DEBUG, "", message
End Sub

Public Sub WarnLog(ByVal message As String)
    OutputLog LOGLEVEL_WARN, "", message
End Sub

Public Sub ErrorLog(ByVal message As String)
    OutputLog LOGLEVEL_ERROR, "", message
End Sub

Public Sub TraceLog(ByVal message As String)
    OutputLog LOGLEVEL_TRACE, "", message
End Sub


' 共通ログ出力処理
Private Sub OutputLog(ByVal level As String, ByVal moduleName As String, ByVal message As String)
    If Not logLevels.Exists(level) Or Not logLevels(level) Then Exit Sub

    Dim paddedLevel As String
    paddedLevel = "[" & Format(level, "@@@@@") & "]" ' 5文字で桁揃え

    Dim formatted As String
    formatted = Replace(formatTemplate, "{time}", GetPreciseTime("-"))
    formatted = Replace(formatted, "{level}", paddedLevel)
    formatted = Replace(formatted, "{module}", moduleName)
    formatted = Replace(formatted, "{message}", message)

    If Not outputHandler Is Nothing Then
        outputHandler.OutputLog formatted
    End If
End Sub

Function GetPreciseTime(datedelim As String) As String
    Dim st As SYSTEMTIME
    GetLocalTime st
    GetPreciseTime = Format(st.wYear, "0000") & datedelim & _
                     Format(st.wMonth, "00") & datedelim & _
                     Format(st.wDay, "00") & " " & _
                     Format(st.wHour, "00") & ":" & _
                     Format(st.wMinute, "00") & ":" & _
                     Format(st.wSecond, "00") & "." & _
                     Format(st.wMilliseconds, "000")
End Function

