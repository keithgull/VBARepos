VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================
' clsExcelHelper - Excel操作ヘルパクラス
' セル操作、書式設定、数式入力を簡略化
'=========================================
Option Explicit

Private m_Sheet As Worksheet
Private m_LastError As String

' ========================================
' プロパティ
' ========================================

Public Property Set Sheet(ByVal ws As Worksheet)
    Set m_Sheet = ws
End Property

Public Property Get Sheet() As Worksheet
    Set Sheet = m_Sheet
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

' ========================================
' 初期化
' ========================================

Private Sub Class_Initialize()
    m_LastError = ""
End Sub

' アクティブシートで初期化
Public Sub InitWithActiveSheet()
    Set m_Sheet = ActiveSheet
End Sub

' シート名で初期化
Public Function InitWithSheetName(ByVal sheetName As String) As Boolean
    On Error Resume Next
    Set m_Sheet = ThisWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then
        m_LastError = "シート '" & sheetName & "' が見つかりません"
        InitWithSheetName = False
        Exit Function
    End If
    On Error GoTo 0
    InitWithSheetName = True
End Function

' ========================================
' セル値の設定・取得
' ========================================

' セルに値を設定（行列番号指定）
Public Sub SetValue(ByVal Row As Long, ByVal Col As Long, ByVal Value As Variant)
    m_Sheet.cells(Row, Col).Value = Value
End Sub

' セルに値を設定（アドレス指定）
Public Sub SetValueAt(ByVal address As String, ByVal Value As Variant)
    m_Sheet.Range(address).Value = Value
End Sub

' セル値を取得（行列番号指定）
Public Function GetValue(ByVal Row As Long, ByVal Col As Long) As Variant
    GetValue = m_Sheet.cells(Row, Col).Value
End Function

' セル値を取得（アドレス指定）
Public Function GetValueAt(ByVal address As String) As Variant
    GetValueAt = m_Sheet.Range(address).Value
End Function

' 範囲に配列で値を一括設定（高速）
Public Sub SetValues(ByVal startRow As Long, ByVal startCol As Long, ByRef values As Variant)
    Dim rows As Long, cols As Long
    rows = UBound(values, 1) - LBound(values, 1) + 1
    cols = UBound(values, 2) - LBound(values, 2) + 1
    m_Sheet.cells(startRow, startCol).Resize(rows, cols).Value = values
End Sub

' 範囲の値を配列で一括取得（高速）
Public Function GetValues(ByVal address As String) As Variant
    GetValues = m_Sheet.Range(address).Value
End Function

' ========================================
' 数式の設定
' ========================================

' セルに数式を設定
Public Sub SetFormula(ByVal Row As Long, ByVal Col As Long, ByVal formula As String)
    m_Sheet.cells(Row, Col).formula = formula
End Sub

' セルに数式を設定（アドレス指定）
Public Sub SetFormulaAt(ByVal address As String, ByVal formula As String)
    m_Sheet.Range(address).formula = formula
End Sub

' R1C1形式の数式を設定
Public Sub SetFormulaR1C1(ByVal Row As Long, ByVal Col As Long, ByVal formula As String)
    m_Sheet.cells(Row, Col).FormulaR1C1 = formula
End Sub

' 配列数式を設定
Public Sub SetArrayFormula(ByVal address As String, ByVal formula As String)
    m_Sheet.Range(address).FormulaArray = formula
End Sub

' ========================================
' 書式設定 - 基本
' ========================================

' フォント設定（一括）
Public Sub SetFont(ByVal address As String, _
                   Optional ByVal fontName As String = "", _
                   Optional ByVal fontSize As Single = 0, _
                   Optional ByVal bold As Variant = Empty, _
                   Optional ByVal italic As Variant = Empty, _
                   Optional ByVal fontColor As Variant = Empty)
    With m_Sheet.Range(address).Font
        If Len(fontName) > 0 Then .Name = fontName
        If fontSize > 0 Then .Size = fontSize
        If Not IsEmpty(bold) Then .bold = bold
        If Not IsEmpty(italic) Then .italic = italic
        If Not IsEmpty(fontColor) Then .Color = fontColor
    End With
End Sub

' 背景色設定
Public Sub SetBackColor(ByVal address As String, ByVal Color As Long)
    m_Sheet.Range(address).Interior.Color = Color
End Sub

' 背景色設定（RGB指定）
Public Sub SetBackColorRGB(ByVal address As String, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)
    m_Sheet.Range(address).Interior.Color = RGB(r, g, b)
End Sub

' フォント色設定
Public Sub SetFontColor(ByVal address As String, ByVal Color As Long)
    m_Sheet.Range(address).Font.Color = Color
End Sub

' フォント色設定（RGB指定）
Public Sub SetFontColorRGB(ByVal address As String, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)
    m_Sheet.Range(address).Font.Color = RGB(r, g, b)
End Sub

' ========================================
' 書式設定 - 配置
' ========================================

' 水平配置
Public Sub SetHAlign(ByVal address As String, ByVal alignment As String)
    Dim ha As XlHAlign
    Select Case LCase(alignment)
        Case "left", "l", "左": ha = xlHAlignLeft
        Case "center", "c", "中央": ha = xlHAlignCenter
        Case "right", "r", "右": ha = xlHAlignRight
        Case "justify", "j", "均等": ha = xlHAlignJustify
        Case Else: ha = xlHAlignGeneral
    End Select
    m_Sheet.Range(address).HorizontalAlignment = ha
End Sub

' 垂直配置
Public Sub SetVAlign(ByVal address As String, ByVal alignment As String)
    Dim va As XlVAlign
    Select Case LCase(alignment)
        Case "top", "t", "上": va = xlVAlignTop
        Case "center", "c", "中央": va = xlVAlignCenter
        Case "bottom", "b", "下": va = xlVAlignBottom
        Case Else: va = xlVAlignCenter
    End Select
    m_Sheet.Range(address).VerticalAlignment = va
End Sub

' 折り返し設定
Public Sub SetWrapText(ByVal address As String, ByVal wrap As Boolean)
    m_Sheet.Range(address).WrapText = wrap
End Sub

' インデント設定
Public Sub SetIndent(ByVal address As String, ByVal Level As Integer)
    m_Sheet.Range(address).IndentLevel = Level
End Sub

' ========================================
' 書式設定 - 罫線
' ========================================

' 罫線スタイル定数
Public Enum BorderStyle
    bsNone = xlLineStyleNone
    bsThin = xlContinuous         ' 細線
    bsMedium = xlMedium           ' 中線
    bsThick = xlThick             ' 太線
    bsDashed = xlDash             ' 破線
    bsDotted = xlDot              ' 点線
    bsDouble = xlDouble           ' 二重線
End Enum

' 全周罫線
Public Sub SetBorderAll(ByVal address As String, _
                        Optional ByVal style As BorderStyle = bsThin, _
                        Optional ByVal Color As Long = 0)
    With m_Sheet.Range(address)
        .Borders.LineStyle = style
        If Color <> 0 Then .Borders.Color = Color
    End With
End Sub

' 外枠罫線のみ
Public Sub SetBorderOutline(ByVal address As String, _
                            Optional ByVal style As BorderStyle = bsThin, _
                            Optional ByVal Color As Long = 0)
    With m_Sheet.Range(address)
        .BorderAround LineStyle:=style, Color:=IIf(Color = 0, vbBlack, Color)
    End With
End Sub

' 個別罫線設定
Public Sub SetBorderEdge(ByVal address As String, _
                         ByVal edge As String, _
                         Optional ByVal style As BorderStyle = bsThin, _
                         Optional ByVal Color As Long = 0)
    Dim borderIndex As XlBordersIndex
    Select Case LCase(edge)
        Case "top", "上": borderIndex = xlEdgeTop
        Case "bottom", "下": borderIndex = xlEdgeBottom
        Case "left", "左": borderIndex = xlEdgeLeft
        Case "right", "右": borderIndex = xlEdgeRight
        Case "insideh", "内横": borderIndex = xlInsideHorizontal
        Case "insidev", "内縦": borderIndex = xlInsideVertical
    End Select
    
    With m_Sheet.Range(address).Borders(borderIndex)
        .LineStyle = style
        If Color <> 0 Then .Color = Color
    End With
End Sub

' ========================================
' 書式設定 - 表示形式
' ========================================

' 数値書式設定
Public Sub SetNumberFormat(ByVal address As String, ByVal formatCode As String)
    m_Sheet.Range(address).NumberFormat = formatCode
End Sub

' よく使う書式のショートカット
Public Sub SetFormatNumber(ByVal address As String, Optional ByVal decimals As Integer = 0)
    m_Sheet.Range(address).NumberFormat = "#,##0" & IIf(decimals > 0, "." & String(decimals, "0"), "")
End Sub

Public Sub SetFormatPercent(ByVal address As String, Optional ByVal decimals As Integer = 1)
    m_Sheet.Range(address).NumberFormat = "0" & IIf(decimals > 0, "." & String(decimals, "0"), "") & "%"
End Sub

Public Sub SetFormatCurrency(ByVal address As String, Optional ByVal symbol As String = "\")
    m_Sheet.Range(address).NumberFormat = symbol & "#,##0"
End Sub

Public Sub SetFormatDate(ByVal address As String, Optional ByVal pattern As String = "yyyy/mm/dd")
    m_Sheet.Range(address).NumberFormat = pattern
End Sub

Public Sub SetFormatText(ByVal address As String)
    m_Sheet.Range(address).NumberFormat = "@"
End Sub

' ========================================
' セル操作
' ========================================

' セル結合
Public Sub MergeCells(ByVal address As String)
    m_Sheet.Range(address).Merge
End Sub

' セル結合解除
Public Sub UnmergeCells(ByVal address As String)
    m_Sheet.Range(address).UnMerge
End Sub

' 行高さ設定
Public Sub SetRowHeight(ByVal Row As Long, ByVal height As Single)
    m_Sheet.rows(Row).RowHeight = height
End Sub

' 列幅設定
Public Sub SetColumnWidth(ByVal Col As Variant, ByVal Width As Single)
    m_Sheet.Columns(Col).ColumnWidth = Width
End Sub

' 列幅自動調整
Public Sub AutoFitColumn(ByVal Col As Variant)
    m_Sheet.Columns(Col).AutoFit
End Sub

' 範囲の列幅自動調整
Public Sub AutoFitColumns(ByVal address As String)
    m_Sheet.Range(address).Columns.AutoFit
End Sub

' 行の表示/非表示
Public Sub SetRowVisible(ByVal Row As Long, ByVal Visible As Boolean)
    m_Sheet.rows(Row).Hidden = Not Visible
End Sub

' 列の表示/非表示
Public Sub SetColumnVisible(ByVal Col As Variant, ByVal Visible As Boolean)
    m_Sheet.Columns(Col).Hidden = Not Visible
End Sub

' ========================================
' 範囲操作
' ========================================

' 範囲クリア（値のみ）
Public Sub ClearValues(ByVal address As String)
    m_Sheet.Range(address).ClearContents
End Sub

' 範囲クリア（書式含む）
Public Sub ClearAll(ByVal address As String)
    m_Sheet.Range(address).Clear
End Sub

' 範囲コピー
Public Sub CopyRange(ByVal sourceAddress As String, ByVal destAddress As String)
    m_Sheet.Range(sourceAddress).Copy Destination:=m_Sheet.Range(destAddress)
End Sub

' データ最終行を取得
Public Function GetLastRow(Optional ByVal Col As Variant = 1) As Long
    GetLastRow = m_Sheet.cells(m_Sheet.rows.Count, Col).End(xlUp).Row
End Function

' データ最終列を取得
Public Function GetLastColumn(Optional ByVal Row As Long = 1) As Long
    GetLastColumn = m_Sheet.cells(Row, m_Sheet.Columns.Count).End(xlToLeft).Column
End Function

' ========================================
' 名前の定義
' ========================================

' 名前の定義を追加
Public Sub AddNamedRange(ByVal Name As String, ByVal address As String, Optional ByVal isLocal As Boolean = True)
    If isLocal Then
        m_Sheet.Names.Add Name:=Name, RefersTo:=m_Sheet.Range(address)
    Else
        ThisWorkbook.Names.Add Name:=Name, RefersTo:=m_Sheet.Range(address)
    End If
End Sub

' 名前の定義を削除
Public Sub DeleteNamedRange(ByVal Name As String)
    On Error Resume Next
    ThisWorkbook.Names(Name).Delete
    On Error GoTo 0
End Sub

' ========================================
' 条件付き書式
' ========================================

' 条件付き書式を追加（数式ベース）
Public Sub AddConditionalFormat(ByVal address As String, _
                                 ByVal formula As String, _
                                 Optional ByVal backColor As Long = -1, _
                                 Optional ByVal fontColor As Long = -1, _
                                 Optional ByVal bold As Variant = Empty)
    Dim cf As FormatCondition
    Set cf = m_Sheet.Range(address).FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    
    If backColor >= 0 Then cf.Interior.Color = backColor
    If fontColor >= 0 Then cf.Font.Color = fontColor
    If Not IsEmpty(bold) Then cf.Font.bold = bold
End Sub

' 条件付き書式をクリア
Public Sub ClearConditionalFormats(ByVal address As String)
    m_Sheet.Range(address).FormatConditions.Delete
End Sub

' ========================================
' データ入力規則（ドロップダウン）
' ========================================

' ドロップダウンリストを設定（配列から）
Public Sub SetDropdownList(ByVal address As String, ByVal items As Variant)
    Dim listString As String
    If IsArray(items) Then
        listString = Join(items, ",")
    Else
        listString = CStr(items)
    End If
    
    With m_Sheet.Range(address).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=listString
        .ShowError = True
    End With
End Sub

' ドロップダウンリストを設定（範囲参照）
Public Sub SetDropdownFromRange(ByVal address As String, ByVal sourceRange As String)
    With m_Sheet.Range(address).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=" & sourceRange
        .ShowError = True
    End With
End Sub

' ========================================
' ユーティリティ
' ========================================

' 列番号を列文字に変換
Public Function ColNumToLetter(ByVal colNum As Long) As String
    ColNumToLetter = Split(m_Sheet.cells(1, colNum).address, "$")(1)
End Function

' 列文字を列番号に変換
Public Function ColLetterToNum(ByVal colLetter As String) As Long
    ColLetterToNum = m_Sheet.Range(colLetter & "1").Column
End Function

' セルアドレス生成
Public Function MakeAddress(ByVal Row As Long, ByVal Col As Long, _
                            Optional ByVal absRow As Boolean = False, _
                            Optional ByVal absCol As Boolean = False) As String
    MakeAddress = m_Sheet.cells(Row, Col).address(absRow, absCol)
End Function

' 範囲アドレス生成
Public Function MakeRangeAddress(ByVal startRow As Long, ByVal startCol As Long, _
                                  ByVal endRow As Long, ByVal endCol As Long) As String
    MakeRangeAddress = m_Sheet.Range(m_Sheet.cells(startRow, startCol), _
                                      m_Sheet.cells(endRow, endCol)).address
End Function

