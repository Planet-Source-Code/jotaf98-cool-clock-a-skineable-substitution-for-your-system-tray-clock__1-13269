Attribute VB_Name = "mdlIniReadAndWrite"

'This module has 2 functions to read and write to INI files.

'APIs:
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpINIPath As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpINIPath As String) As Long

'Vars:
Dim INIfPath As String

Public Function ReadINI(ByVal Section As String, ByVal Key As String, ByVal INIPath As String) As String
    Dim RetStr As String
    
    If Dir(INIPath) = "" Then Exit Function
    
    RetStr = String(255, Chr(0))
    ReadINI = Left(RetStr, GetPrivateProfileString(Section, ByVal Key, "", RetStr, Len(RetStr), INIPath))
End Function

Public Function WriteINI(ByVal Section As String, ByVal Key As String, ByVal KeyValue As String, ByVal INIPath As String) As Integer
    'Function returns 1 if successful and 0 if unsuccessful
    
    If Dir(INIPath) = "" Then Exit Function
    WritePrivateProfileString Section, Key, KeyValue, INIPath
    WriteINI = 1
End Function
