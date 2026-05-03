VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================
' clsTableBuilder - テーブル作成ヘルパクラス
' 表形式のデータを簡単に作成・書式設定
'=========================================
Option Explicit

Private m_Helper As clsExcelHelper
Private m_StartRow As Long
Private m_StartCol As Long
Private m_CurrentRow As Long
Private m_Headers() As String
Private m_ColCount As Long

' ========================================
' 初期化
' ========================================

Private Sub Class_Initialize()
    Set m_Helper = New clsExcelHelper
    m_StartRow = 1
    m_StartCol = 1
    m_CurrentRow = 1
    m_ColCount = 0
End Sub

' シートとテーブル開始位置を設定
Public Sub Init(ByVal ws As Worksheet, Optional ByVal startRow As Long = 1, Optional ByVal startCol As Long = 1)
    Set m_Helper.Sheet = ws
    m_StartRow = startRow
    m_StartCol = startCol
    m_CurrentRow = startRow
End Sub

' ========================================
' ヘッダー設定
' ========================================

' ヘッダーを設定して書き込み
Public Sub SetHeaders(ParamArray headers() As Variant)
    Dim i As Long
    m_ColCount = UBound(headers) - LBound(headers) + 1
    ReDim m_Headers(0 To m_ColCount - 1)
    
    For i = 0 To m_ColCount - 1
        m_Headers(i) = CStr(headers(i))
        m_Helper.SetValue m_StartRow, m_StartCol + i, m_Headers(i)
    Next i
    
    ' ヘッダーのデフォルト書式
    Dim headerRange As String
    headerRange = m_Helper.MakeRangeAddress(m_StartRow, m_StartCol, m_StartRow, m_StartCol + m_ColCount - 1)
    m_Helper.SetFont headerRange, bold:=True
    m_Helper.SetBackColorRGB headerRange, 68, 114, 196  ' 青系
    m_Helper.SetFontColorRGB headerRange, 255, 255, 255 ' 白
    m_Helper.SetHAlign headerRange, "center"
    
    m_CurrentRow = m_StartRow + 1
End Sub

' ========================================
' データ行追加
' ========================================

' 1行追加
Public Sub AddRow(ParamArray values() As Variant)
    Dim i As Long
    For i = 0 To UBound(values)
        If i < m_ColCount Then
            m_Helper.SetValue m_CurrentRow, m_StartCol + i, values(i)
        End If
    Next i
    m_CurrentRow = m_CurrentRow + 1
End Sub

' 配列から複数行追加
Public Sub AddRows(ByRef data As Variant)
    Dim i As Long, j As Long
    For i = LBound(data, 1) To UBound(data, 1)
        For j = LBound(data, 2) To UBound(data, 2)
            m_Helper.SetValue m_CurrentRow, m_StartCol + j - LBound(data, 2), data(i, j)
        Next j
        m_CurrentRow = m_CurrentRow + 1
    Next i
End Sub

' ========================================
' テーブル書式設定
' ========================================

' テーブル全体に罫線を設定
Public Sub ApplyBorders(Optional ByVal style As BorderStyle = bsThin)
    Dim tableRange As String
    tableRange = GetTableRange()
    m_Helper.SetBorderAll tableRange, style
End Sub

' 縞模様（ゼブラストライプ）を適用
Public Sub ApplyZebraStripes(Optional ByVal evenColor As Long = 15790320) ' 薄いグレー
    Dim i As Long
    Dim rowRange As String
    
    For i = m_StartRow + 1 To m_CurrentRow - 1 Step 2
        rowRange = m_Helper.MakeRangeAddress(i, m_StartCol, i, m_StartCol + m_ColCount - 1)
        m_Helper.SetBackColor rowRange, evenColor
    Next i
End Sub

' 列幅を自動調整
Public Sub AutoFitColumns()
    m_Helper.AutoFitColumns GetTableRange()
End Sub

' 特定列の書式を設定
Public Sub SetColumnFormat(ByVal colIndex As Long, ByVal formatCode As String)
    Dim colRange As String
    colRange = m_Helper.MakeRangeAddress(m_StartRow + 1, m_StartCol + colIndex - 1, _
                                          m_CurrentRow - 1, m_StartCol + colIndex - 1)
    m_Helper.SetNumberFormat colRange, formatCode
End Sub

' 特定列を数値書式に
Public Sub SetColumnAsNumber(ByVal colIndex As Long, Optional ByVal decimals As Integer = 0)
    Dim colRange As String
    colRange = m_Helper.MakeRangeAddress(m_StartRow + 1, m_StartCol + colIndex - 1, _
                                          m_CurrentRow - 1, m_StartCol + colIndex - 1)
    m_Helper.SetFormatNumber colRange, decimals
End Sub

' 特定列をパーセント書式に
Public Sub SetColumnAsPercent(ByVal colIndex As Long, Optional ByVal decimals As Integer = 1)
    Dim colRange As String
    colRange = m_Helper.MakeRangeAddress(m_StartRow + 1, m_StartCol + colIndex - 1, _
                                          m_CurrentRow - 1, m_StartCol + colIndex - 1)
    m_Helper.SetFormatPercent colRange, decimals
End Sub

' ========================================
' ユーティリティ
' ========================================

' テーブル範囲を取得
Public Function GetTableRange() As String
    GetTableRange = m_Helper.MakeRangeAddress(m_StartRow, m_StartCol, _
                                               m_CurrentRow - 1, m_StartCol + m_ColCount - 1)
End Function

' データ範囲を取得（ヘッダー除く）
Public Function GetDataRange() As String
    GetDataRange = m_Helper.MakeRangeAddress(m_StartRow + 1, m_StartCol, _
                                              m_CurrentRow - 1, m_StartCol + m_ColCount - 1)
End Function

' 現在の行番号を取得
Public Property Get CurrentRow() As Long
    CurrentRow = m_CurrentRow
End Property

' 行数を取得（ヘッダー含む）
Public Property Get RowCount() As Long
    RowCount = m_CurrentRow - m_StartRow
End Property

' ヘルパーオブジェクトを取得（詳細操作用）
Public Property Get Helper() As clsExcelHelper
    Set Helper = m_Helper
End Property

