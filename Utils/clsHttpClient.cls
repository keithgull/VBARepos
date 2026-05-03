VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================
' clsHttpClient - HTTP通信用クラス
' Webサイトからデータを取得する汎用クラス
'=========================================
Option Explicit

Private m_Http As Object
Private m_LastStatus As Long
Private m_LastError As String
Private m_Timeout As Long

' ========================================
' プロパティ
' ========================================

Public Property Get LastStatus() As Long
    LastStatus = m_LastStatus
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Property Let Timeout(ByVal Value As Long)
    m_Timeout = Value
End Property

Public Property Get Timeout() As Long
    Timeout = m_Timeout
End Property

' ========================================
' 初期化・終了処理
' ========================================

Private Sub Class_Initialize()
    m_Timeout = 30000 ' デフォルト30秒
    m_LastStatus = 0
    m_LastError = ""
End Sub

Private Sub Class_Terminate()
    Set m_Http = Nothing
End Sub

' ========================================
' パブリックメソッド
' ========================================

' URLからHTMLをUTF-8で取得する
Public Function FetchHTML(ByVal url As String) As String
    Dim HtmlText As String
    
    m_LastError = ""
    m_LastStatus = 0
    
    ' HTTPオブジェクトを作成
    If Not CreateHttpObject() Then
        FetchHTML = ""
        Exit Function
    End If
    
    ' リクエスト送信
    If Not SendRequest(url) Then
        FetchHTML = ""
        Exit Function
    End If
    
    ' レスポンスをUTF-8でデコード
    HtmlText = DecodeResponseAsUTF8()
    
    Set m_Http = Nothing
    FetchHTML = HtmlText
End Function

' URLからバイナリデータを取得する
Public Function FetchBinary(ByVal url As String) As Byte()
    Dim result() As Byte
    
    m_LastError = ""
    m_LastStatus = 0
    
    If Not CreateHttpObject() Then
        FetchBinary = result
        Exit Function
    End If
    
    If Not SendRequest(url) Then
        FetchBinary = result
        Exit Function
    End If
    
    result = m_Http.responseBody
    Set m_Http = Nothing
    FetchBinary = result
End Function

' ========================================
' プライベートメソッド
' ========================================

' HTTPオブジェクトを作成する
Private Function CreateHttpObject() As Boolean
    On Error Resume Next
    
    ' 複数のオブジェクトタイプを試行
    Set m_Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set m_Http = CreateObject("MSXML2.XMLHTTP.6.0")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Set m_Http = CreateObject("MSXML2.XMLHTTP")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Set m_Http = CreateObject("Microsoft.XMLHTTP")
    End If
    
    On Error GoTo 0
    
    If m_Http Is Nothing Then
        m_LastError = "HTTPオブジェクトを作成できませんでした"
        CreateHttpObject = False
    Else
        CreateHttpObject = True
    End If
End Function

' HTTPリクエストを送信する
Private Function SendRequest(ByVal url As String) As Boolean
    On Error Resume Next
    
    m_Http.Open "GET", url, False
    If Err.Number <> 0 Then
        m_LastError = "HTTP Open エラー: " & Err.Description
        SendRequest = False
        Exit Function
    End If
    
    ' タイムアウト設定（ServerXMLHTTPの場合のみ）
    If TypeName(m_Http) Like "*ServerXMLHTTP*" Then
        m_Http.setTimeouts m_Timeout, m_Timeout, m_Timeout, m_Timeout
    End If
    
    ' リクエストヘッダー設定
    m_Http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    m_Http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    m_Http.setRequestHeader "Accept-Language", "ja,en-US;q=0.7,en;q=0.3"
    m_Http.setRequestHeader "Accept-Charset", "UTF-8"
    
    m_Http.send
    
    If Err.Number <> 0 Then
        m_LastError = "HTTP Send エラー: " & Err.Description
        SendRequest = False
        Exit Function
    End If
    
    On Error GoTo 0
    
    m_LastStatus = m_Http.Status
    
    If m_Http.Status <> 200 Then
        m_LastError = "HTTPステータスエラー: " & m_Http.Status
        SendRequest = False
        Exit Function
    End If
    
    SendRequest = True
End Function

' レスポンスをUTF-8でデコードする
Private Function DecodeResponseAsUTF8() As String
    Dim stream As Object
    Dim HtmlText As String
    
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    
    If Err.Number <> 0 Then
        ' フォールバック
        DecodeResponseAsUTF8 = m_Http.responseText
        Exit Function
    End If
    
    stream.Type = 1 ' バイナリ
    stream.Open
    stream.Write m_Http.responseBody
    stream.Position = 0
    stream.Type = 2 ' テキスト
    stream.Charset = "UTF-8"
    HtmlText = stream.ReadText
    stream.Close
    
    On Error GoTo 0
    
    Set stream = Nothing
    DecodeResponseAsUTF8 = HtmlText
End Function
