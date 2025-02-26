Attribute VB_Name = "modSheetUtils"
Option Explicit

Function GetCellAddress(rowNum As Long, colNum As Integer) As String
    GetCellAddress = Cells(rowNum, colNum).Address
End Function

Public Sub ApplyPriceFormat(ws As Worksheet, target As Range, formatStr As String)
    target.NumberFormat = formatStr
End Sub

Function AddSheet(templateSheetName As String, newSheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim template As Worksheet
    
    ' テンプレートシートの取得
    Set template = GetWorksheetByName(templateSheetName)
    
    If template Is Nothing Then
        MsgBox "テンプレートシートが見つかりません。", vbExclamation
        Set AddSheet = Nothing
        Exit Function
    End If

    ' 新しいシート名が既に存在するか確認
    If WorksheetExists(newSheetName) Then
        MsgBox "シート名が既に存在します。", vbExclamation
        Set AddSheet = Nothing
        Exit Function
    End If
    
    ' テンプレートシートをコピーして新しいシートを追加
    template.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ws.Name = newSheetName
    
    ' 追加されたシートを戻り値として返す
    Set AddSheet = ws
End Function

Sub DeleteSheet(targetSheetName As String)
    Dim ws As Worksheet
    
    ' 対象シートの取得
    Set ws = GetWorksheetByName(targetSheetName)
    
    If ws Is Nothing Then
        MsgBox "対象シートが見つかりません。", vbExclamation
        Exit Sub
    End If
End Sub

Function GetWorksheetByName(sheetName As String, silentmode As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        if silentmode = false then
        	MsgBox "対象シートが見つかりません。", vbExclamation
        end if
        Exit Function
    End If
    Set GetWorksheetByName = ws
End Function

Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    WorksheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function
