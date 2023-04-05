Attribute VB_Name = "CommonModule"
Option Explicit

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const CP_ACP = 0 ' default to ANSI code page
Public Const CP_UTF8 = 65001 ' default to UTF-8 code page

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal HWND As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Long) As Long
Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal length As Long)
Public Declare Function MessageBox _
               Lib "user32" _
               Alias "MessageBoxA" (ByVal HWND As Long, _
                                    ByVal lpText As String, _
                                    ByVal lpCaption As String, _
                                    ByVal wType As Long) As Long

Public Declare Function MultiByteToWideChar _
               Lib "kernel32 " (ByVal CodePage As Long, _
                                ByVal dwFlags As Long, _
                                ByVal lpMultiByteStr As Long, _
                                ByVal cchMultiByte As Long, _
                                ByVal lpWideCharStr As Long, _
                                ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte _
               Lib "kernel32 " (ByVal CodePage As Long, _
                                ByVal dwFlags As Long, _
                                ByVal lpWideCharStr As Long, _
                                ByVal cchWideChar As Long, _
                                ByVal lpMultiByteStr As Long, _
                                ByVal cchMultiByte As Long, _
                                ByVal lpDefaultChar As Long, _
                                ByVal lpUsedDefaultChar As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long


'×Ö·û×ª UTF8
Public Function StringToUtf8(ByVal sData As String) As Byte()  ' Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize   As Long
    nSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), -1, 0, 0, 0, 0) - 1

    If nSize = 0 Or nSize = -1 Then Exit Function
    ReDim aRetn(0 To nSize - 1) As Byte
    WideCharToMultiByte CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize, 0, 0
    StringToUtf8 = aRetn
    Erase aRetn
End Function

' UTF8 ×ª×Ö·û
Public Function Utf8ToString(ByVal sData As Long) As Byte() ' Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize   As Long
    nSize = MultiByteToWideChar(CP_UTF8, 0, sData, -1, 0, 0) - 1

    If nSize = 0 Or nSize = -1 Then Exit Function
    ReDim aRetn(0 To 2 * nSize - 1) As Byte
    MultiByteToWideChar CP_UTF8, 0, sData, -1, VarPtr(aRetn(0)), nSize
    Utf8ToString = aRetn
    Erase aRetn
End Function

Function FuncAddressOf(addr As Long) As Long
    FuncAddressOf = addr
End Function

Function AppendLog(strContent As String, Optional strFileName As String = "Log.txt")
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open App.Path + "/" + strFileName For Append As #f
    Print #f, strContent
    Close #f
End Function


'Dim lsMb As Dictionary
