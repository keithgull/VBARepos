Attribute VB_Name = "Module2"
Option Explicit

Public Const PAGE_EXECUTE_READWRITE = &H40
Public Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
Public Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
Public Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
Public Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Public Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer

Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim projectFunction As Long
Dim Flag As Boolean

Public Function GetPtr(ByVal value As LongPtr) As LongPtr
    GetPtr = value
End Function

Public Sub RecoverBytes()
    If Flag Then MoveMemory ByVal projectFunction, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function MyDialogBoxParamater(ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
  If pTemplateName = 4070 Then
    MyDialogBoxParamater = 1
  Else
    RecoverBytes
    MyDialogBoxParamater = MyDialogBoxParamater(hInstance, pTemplateName, hWndParent, lpDialogFunc, dwInitParam)
    HookFlag
  End If
End Function

Public Function HookFlag() As Boolean
  Dim TmpBytes(0 To 5) As Byte
  Dim p As LongPtr
  Dim OriginProtect As LongPtr
 
  HookFlag = False
  projectFunction = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")

  If VirtualProtect(ByVal projectFunction, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
    MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal projectFunction, 6
    If TmpBytes(0) <> &H68 Then
      MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal projectFunction, 6
      p = GetPtr(AddressOf MyDialogBoxParamater)
      HookBytes(0) = &H68
      MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
      HookBytes(5) = &HC3
      MoveMemory ByVal projectFunction, ByVal VarPtr(HookBytes(0)), 6
      Flag = True
      HookFlag = True
    End If
  End If
End Function

'Sub unlockVBAProject()
'  If HookFlag Then
'    MsgBox "VBA Project ���������܂����B", vbInformation, "�������܂����B"
'  End If
'End Sub

Sub unlockVBAProject(password As String)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim vbProjName As String

    Set vbProj = ThisWorkbook.VBProject
    On Error Resume Next
    vbProj.vbComponents("ThisWorkbook").Activate
    Application.VBE.CommandBars("Menu Bar").Controls("�c�[��").Controls("VBAProject �v���p�e�B").Execute
    SendKeys password & "~", True
    On Error GoTo 0
End Sub

Sub lockVBAProject(password As String)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim vbProjName As String
    
    Set vbProj = ThisWorkbook.VBProject
    On Error Resume Next
    vbProj.vbComponents("ThisWorkbook").Activate
    Application.VBE.CommandBars("Menu Bar").Controls("�c�[��").Controls("VBAProject �v���p�e�B").Execute
    On Error GoTo 0

    SendKeys password & "~", True
    SendKeys password & "~", True

End Sub
