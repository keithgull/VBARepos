Attribute VB_Name = "modStringUtils"
Option Explicit

Public Function LPad(basestr As String, padStr As String, totalLen As Integer) As String
    Dim strlen As Integer
    strlen = Len(basestr)
    If totalLen > strlen Then
        LPad = WorksheetFunction.Rept(padStr, totalLen - strlen) & basestr
    ElseIf totalLen < strlen Then
        LPad = Right(basestr, totalLen)
    End If
End Function

Public Function RPad(basestr As String, padStr As String, totalLen As Integer) As String
    Dim strlen As Integer
    strlen = Len(basestr)
    If totalLen > strlen Then
        RPad = basestr & WorksheetFunction.Rept(padStr, totalLen - strlen)
    ElseIf totalLen < strlen Then
        RPad = Left(basestr, totalLen)
    End If
End Function


Public Function LPadSpc(basestr As String, totalLen As Integer) As String
    LPadSpc = LPad(basestr, " ", totalLen)
End Function
