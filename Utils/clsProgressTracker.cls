VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgressTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================
' clsProgressTracker - 進捗管理クラス
' 処理の進捗表示とログ管理を行う汎用クラス
'=========================================
Option Explicit

Private m_StartTime As Date
Private m_TaskName As String
Private m_TotalItems As Long
Private m_ProcessedItems As Long
Private m_EnableDebugLog As Boolean
Private m_EnableStatusBar As Boolean

' ========================================
' プロパティ
' ========================================

Public Property Let taskName(ByVal Value As String)
    m_TaskName = Value
End Property

Public Property Get taskName() As String
    taskName = m_TaskName
End Property

Public Property Let totalItems(ByVal Value As Long)
    m_TotalItems = Value
End Property

Public Property Get totalItems() As Long
    totalItems = m_TotalItems
End Property

Public Property Get ProcessedItems() As Long
    ProcessedItems = m_ProcessedItems
End Property

Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = (Now() - m_StartTime) * 86400
End Property

Public Property Let EnableDebugLog(ByVal Value As Boolean)
    m_EnableDebugLog = Value
End Property

Public Property Let EnableStatusBar(ByVal Value As Boolean)
    m_EnableStatusBar = Value
End Property

' ========================================
' 初期化・終了処理
' ========================================

Private Sub Class_Initialize()
    m_StartTime = Now()
    m_TaskName = "処理"
    m_TotalItems = 0
    m_ProcessedItems = 0
    m_EnableDebugLog = True
    m_EnableStatusBar = True
End Sub

Private Sub Class_Terminate()
    If m_EnableStatusBar Then
        Application.StatusBar = False
    End If
End Sub

' ========================================
' パブリックメソッド
' ========================================

' 処理開始
Public Sub StartTask(ByVal taskName As String, Optional ByVal totalItems As Long = 0)
    m_StartTime = Now()
    m_TaskName = taskName
    m_TotalItems = totalItems
    m_ProcessedItems = 0
    
    LogMessage "=========================================="
    LogMessage m_TaskName & " 開始: " & Format(m_StartTime, "yyyy/mm/dd (aaa) hh:nn:ss")
    LogMessage "=========================================="
    
    UpdateStatus m_TaskName & "を開始しています..."
End Sub

' 進捗更新
Public Sub UpdateProgress(Optional ByVal currentItem As Long = -1, _
                          Optional ByVal message As String = "")
    If currentItem >= 0 Then
        m_ProcessedItems = currentItem
    Else
        m_ProcessedItems = m_ProcessedItems + 1
    End If
    
    Dim statusMsg As String
    If m_TotalItems > 0 Then
        Dim percent As Double
        percent = (m_ProcessedItems / m_TotalItems) * 100
        statusMsg = m_TaskName & ": " & m_ProcessedItems & " / " & m_TotalItems & _
                    " (" & Format(percent, "0.0") & "%)"
    Else
        statusMsg = m_TaskName & ": " & m_ProcessedItems & " 件処理済"
    End If
    
    If Len(message) > 0 Then
        statusMsg = statusMsg & " - " & message
    End If
    
    UpdateStatus statusMsg
    
    ' 定期的にDoEventsを実行
    If m_ProcessedItems Mod 50 = 0 Then
        DoEvents
    End If
End Sub

' ログメッセージ出力
Public Sub LogMessage(ByVal message As String)
    If m_EnableDebugLog Then
        Debug.Print message
    End If
End Sub

' 処理完了
Public Sub CompleteTask(Optional ByVal resultMessage As String = "")
    Dim endTime As Date
    Dim elapsed As Double
    
    endTime = Now()
    elapsed = (endTime - m_StartTime) * 86400 ' 秒に変換
    
    LogMessage "----------------------------------------"
    If Len(resultMessage) > 0 Then
        LogMessage resultMessage
    End If
    LogMessage m_TaskName & " 完了: " & Format(endTime, "yyyy/mm/dd (aaa) hh:nn:ss")
    LogMessage "処理時間: " & FormatElapsedTime(elapsed)
    LogMessage "=========================================="
    
    UpdateStatus ""
End Sub

' ========================================
' プライベートメソッド
' ========================================

' ステータスバー更新
Private Sub UpdateStatus(ByVal message As String)
    If m_EnableStatusBar Then
        If Len(message) > 0 Then
            Application.StatusBar = message
        Else
            Application.StatusBar = False
        End If
    End If
End Sub

' 経過時間をフォーマット
Private Function FormatElapsedTime(ByVal seconds As Double) As String
    If seconds < 60 Then
        FormatElapsedTime = Format(seconds, "0.0") & " 秒"
    ElseIf seconds < 3600 Then
        FormatElapsedTime = Format(Int(seconds / 60), "0") & " 分 " & _
                            Format(Int(seconds Mod 60), "0") & " 秒"
    Else
        FormatElapsedTime = Format(Int(seconds / 3600), "0") & " 時間 " & _
                            Format(Int((seconds Mod 3600) / 60), "0") & " 分"
    End If
End Function

' クラスの使用例
Sub SampleUsage()
    ' HTTP通信クラス
    Dim httpClient As New clsHttpClient
    Dim html As String
    
    httpClient.Timeout = 60000 ' 60秒
    html = httpClient.FetchHTML("https://example.com")
    
    If httpClient.LastStatus = 200 Then
        ' HTML解析クラス
        Dim parser As New clsHtmlParser
        parser.HtmlText = html
        
        ' テーブルを抽出
        Dim tables As Collection
        Set tables = parser.ExtractTags("table")
        
        ' 進捗管理クラス
        Dim progress As New clsProgressTracker
        progress.StartTask "テーブル処理", tables.Count
        
        Dim i As Long
        For i = 1 To tables.Count
            ' 処理...
            progress.UpdateProgress
        Next i
        
        progress.CompleteTask "処理完了: " & tables.Count & " テーブル"
    Else
        Debug.Print "エラー: " & httpClient.LastError
    End If
End Sub

