Attribute VB_Name = "mdlMain"

'APIs:

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


'Constants
Enum eSkinType
    stText = False
    stImage = True
End Enum


'Variables:
Global INIFilePath As String

Global SkinType As eSkinType
Global SkinName As String
Global SeparatorChar As String
Global TextColor As Long
Global BackgroundColor As Long
Global TextFontName As String
Global TextSize As Integer

Global ctlControl As Control
