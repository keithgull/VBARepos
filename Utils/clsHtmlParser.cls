VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHtmlParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================
' clsHtmlParser - HTML解析用クラス
' HTMLテキストから要素を抽出する汎用クラス
'=========================================
Option Explicit

Private m_Regex As Object
Private m_HtmlText As String

' ========================================
' プロパティ
' ========================================

Public Property Let HtmlText(ByVal Value As String)
    m_HtmlText = Value
End Property

Public Property Get HtmlText() As String
    HtmlText = m_HtmlText
End Property

' ========================================
' 初期化・終了処理
' ========================================

Private Sub Class_Initialize()
    Set m_Regex = CreateObject("VBScript.RegExp")
    m_Regex.Global = True
    m_Regex.IgnoreCase = True
    m_Regex.Multiline = True
End Sub

Private Sub Class_Terminate()
    Set m_Regex = Nothing
End Sub

' ========================================
' パブリックメソッド
' ========================================

' 正規表現パターンに一致する全てのマッチを取得
Public Function FindAllMatches(ByVal pattern As String, _
                                Optional ByVal targetText As String = "") As Collection
    Dim matches As Object
    Dim match As Object
    Dim results As New Collection
    Dim searchText As String
    
    searchText = IIf(Len(targetText) > 0, targetText, m_HtmlText)
    
    m_Regex.pattern = pattern
    Set matches = m_Regex.Execute(searchText)
    
    For Each match In matches
        results.Add match
    Next match
    
    Set FindAllMatches = results
End Function

' 指定タグの内容を全て抽出
Public Function ExtractTagContents(ByVal tagName As String, _
                                    Optional ByVal targetText As String = "") As Collection
    Dim pattern As String
    Dim matches As Object
    Dim match As Object
    Dim results As New Collection
    Dim searchText As String
    
    searchText = IIf(Len(targetText) > 0, targetText, m_HtmlText)
    
    pattern = "<" & tagName & "[^>]*>([\s\S]*?)</" & tagName & ">"
    m_Regex.pattern = pattern
    Set matches = m_Regex.Execute(searchText)
    
    For Each match In matches
        If match.SubMatches.Count > 0 Then
            results.Add match.SubMatches(0)
        End If
    Next match
    
    Set ExtractTagContents = results
End Function

' 指定タグ全体を全て抽出（属性付き）
Public Function ExtractTags(ByVal tagName As String, _
                            Optional ByVal targetText As String = "") As Collection
    Dim pattern As String
    Dim matches As Object
    Dim match As Object
    Dim results As New Collection
    Dim searchText As String
    
    searchText = IIf(Len(targetText) > 0, targetText, m_HtmlText)
    
    pattern = "<" & tagName & "[^>]*>([\s\S]*?)</" & tagName & ">"
    m_Regex.pattern = pattern
    Set matches = m_Regex.Execute(searchText)
    
    For Each match In matches
        results.Add match.Value
    Next match
    
    Set ExtractTags = results
End Function

' 属性値を抽出
Public Function ExtractAttribute(ByVal tagHtml As String, _
                                  ByVal attributeName As String) As String
    Dim pattern As String
    Dim matches As Object
    
    pattern = attributeName & "=[""']?([^""'>\s]+)[""']?"
    m_Regex.pattern = pattern
    Set matches = m_Regex.Execute(tagHtml)
    
    If matches.Count > 0 Then
        ExtractAttribute = matches(0).SubMatches(0)
    Else
        ExtractAttribute = ""
    End If
End Function

' rowspan属性を抽出
Public Function ExtractRowspan(ByVal cellHtml As String) As Long
    Dim result As String
    result = ExtractAttribute(cellHtml, "rowspan")
    
    If Len(result) > 0 And IsNumeric(result) Then
        ExtractRowspan = CLng(result)
    Else
        ExtractRowspan = 1
    End If
End Function

' colspan属性を抽出
Public Function ExtractColspan(ByVal cellHtml As String) As Long
    Dim result As String
    result = ExtractAttribute(cellHtml, "colspan")
    
    If Len(result) > 0 And IsNumeric(result) Then
        ExtractColspan = CLng(result)
    Else
        ExtractColspan = 1
    End If
End Function

' テーブル行からセルを抽出
Public Function ExtractTableCells(ByVal rowHtml As String) As Collection
    Dim pattern As String
    Dim matches As Object
    Dim match As Object
    Dim results As New Collection
    Dim cellInfo As Variant
    
    pattern = "<t[dh][^>]*>([\s\S]*?)</t[dh]>"
    m_Regex.pattern = pattern
    Set matches = m_Regex.Execute(rowHtml)
    
    For Each match In matches
        ' 配列で格納: (0)=全体HTML, (1)=内容, (2)=rowspan, (3)=colspan
        ReDim cellInfo(0 To 3)
        cellInfo(0) = match.Value
        cellInfo(1) = StripHTML(match.SubMatches(0))
        cellInfo(2) = ExtractRowspan(match.Value)
        cellInfo(3) = ExtractColspan(match.Value)
        results.Add cellInfo
    Next match
    
    Set ExtractTableCells = results
End Function

' HTMLタグを除去
Public Function StripHTML(ByVal htmlString As String) As String
    m_Regex.pattern = "<[^>]+>"
    StripHTML = m_Regex.Replace(htmlString, "")
    
    ' HTMLエンティティをデコード
    StripHTML = DecodeHTMLEntities(StripHTML)
    
    ' 空白を正規化
    StripHTML = NormalizeWhitespace(StripHTML)
End Function

' HTMLエンティティをデコード
Public Function DecodeHTMLEntities(ByVal text As String) As String
    text = Replace(text, "&nbsp;", " ")
    text = Replace(text, "&amp;", "&")
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&apos;", "'")
    text = Replace(text, "&#39;", "'")
    text = Replace(text, "&#9670;", "")
    text = Replace(text, "&#x25C6;", "")
    DecodeHTMLEntities = text
End Function

' 空白を正規化
Public Function NormalizeWhitespace(ByVal text As String) As String
    text = Replace(text, vbCrLf, " ")
    text = Replace(text, vbLf, " ")
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbTab, " ")
    
    ' 連続する空白を1つに
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    
    NormalizeWhitespace = Trim(text)
End Function

' 特定の文字列の出現回数をカウント
Public Function CountOccurrences(ByVal searchString As String, _
                                  Optional ByVal targetText As String = "") As Long
    Dim searchText As String
    searchText = IIf(Len(targetText) > 0, targetText, m_HtmlText)
    
    CountOccurrences = (Len(searchText) - Len(Replace(searchText, searchString, ""))) / Len(searchString)
End Function

