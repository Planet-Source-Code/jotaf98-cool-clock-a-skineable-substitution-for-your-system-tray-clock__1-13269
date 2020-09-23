VERSION 5.00
Begin VB.Form frmClock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cool Clock"
   ClientHeight    =   735
   ClientLeft      =   -90
   ClientTop       =   -660
   ClientWidth     =   1845
   ControlBox      =   0   'False
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   2850
      Left            =   960
      ScaleHeight     =   2850
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picTime 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      ScaleHeight     =   285
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrMainLoop 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   240
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    Dim hWnd As Long, RcTemp As RECT
    
    'Find the system clock's handle
    hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    hWnd = FindWindowEx(hWnd, 0, "TrayNotifyWnd", vbNullString)
    hWnd = FindWindowEx(hWnd, 0, "TrayClockWClass", vbNullString)
    
    'Get its rect
    GetWindowRect hWnd, RcTemp
    
    'Set the new coordinates
    With Me
        .Top = 0
        .Left = 0
        .Height = Me.Height * (RcTemp.Bottom - RcTemp.Top) / Me.ScaleHeight
        .Width = Me.Width * (RcTemp.Right - RcTemp.Left) / Me.ScaleWidth
    End With
    
    'Set this form as a system clock's child...
    SetParent Me.hWnd, hWnd
    
    'Position the display in the middle
    picTime.Left = Me.ScaleWidth \ 2 - picTime.Width \ 2 - 1
    picTime.Top = Me.ScaleHeight \ 2 - picTime.Height \ 2 - 1
    
    'Get the settings
    GetSettings
    
    'Apply them
    ApplySettings
    
    'Start the main loop!
    tmrMainLoop.Enabled = True
End Sub

Public Sub GetSettings()
    Dim TempStr As String
    Dim RGBArray() As String
    
    'Set the INI file path
    INIFilePath = App.Path
    If Right(INIFilePath, 1) <> "\" Then INIFilePath = INIFilePath & "\"
    INIFilePath = INIFilePath & "Options.ini"
    
    'Determine the skin type
    TempStr = ReadINI("General", "SkinType", INIFilePath)
    If TempStr = "Image" Then
        SkinType = stImage
    Else
        SkinType = stText
    End If
    
    'Get the skin image file name
    SkinName = ReadINI("ImageSkin", "SkinName", INIFilePath)
    
    'Get the separator character
    SeparatorChar = ReadINI("TextSkin", "SeparatorChar", INIFilePath)
    
    'Get the text color
    TempStr = ReadINI("TextSkin", "TextColor", INIFilePath)
    RGBArray = Split(TempStr, " ")
    TextColor = RGB(RGBArray(0), RGBArray(1), RGBArray(2))
    
    'Get the background color
    TempStr = ReadINI("TextSkin", "BackgroundColor", INIFilePath)
    RGBArray = Split(TempStr, " ")
    BackgroundColor = RGB(RGBArray(0), RGBArray(1), RGBArray(2))
    
    'Get the font name
    TextFontName = ReadINI("TextSkin", "TextFontName", INIFilePath)
    
    'Get the font size
    TextSize = ReadINI("TextSkin", "TextSize", INIFilePath)
End Sub

Public Sub ApplySettings()
    Dim SkinPath As String
    
    On Error Resume Next
    
    'Load the skin...
    SkinPath = App.Path
    If Right(SkinPath, 1) <> "\" Then SkinPath = SkinPath & "\"
    picSkin.Picture = LoadPicture(SkinPath & "Skins\" & SkinName)
    
    'Set the font and colors
    frmClock.BackColor = BackgroundColor
    picBuffer.ForeColor = TextColor
    picBuffer.BackColor = BackgroundColor
    picBuffer.FontName = TextFontName
    picBuffer.FontSize = TextSize
    
    'This simple loop will set the colors/font of all the
    'controls in the menu
    Load frmMenu
    frmMenu.BackColor = BackgroundColor
    For Each ctlControl In frmMenu.Controls
        If TypeOf ctlControl Is Label Then
            ctlControl.ForeColor = TextColor
            ctlControl.BackColor = BackgroundColor
        ElseIf TypeOf ctlControl Is Line Then
            ctlControl.BorderColor = TextColor
        End If
    Next ctlControl
End Sub

Private Sub tmrMainLoop_Timer()
    Dim TimeStr As String
    Dim i As Integer
    Dim TempCurX As Long, TempCurY As Long
    
    'This is the time as it will be displayed
    TimeStr = Format(Hour(Time), "00") & SeparatorChar & Format(Minute(Time), "00")
    
    If SkinType = stText Then 'Simple text
        'Set the printing coordinates
        TempCurX = Me.ScaleWidth \ 2 - picBuffer.TextWidth(TimeStr) \ 2 - 1
        TempCurY = Me.ScaleHeight \ 2 - picBuffer.TextHeight(TimeStr) \ 2 - 1
        
        'Clear the screen (note it resets CurrentX/Y too so I
        'had to keep them in different variables first)
        picBuffer.Cls
        
        'Print to these coordinates in the buffer
        picBuffer.CurrentX = TempCurX
        picBuffer.CurrentY = TempCurY
        picBuffer.Print TimeStr
    ElseIf SkinType = stImage Then 'Image skin
        'Get the individual numbers from the time string and
        'print them to the buffer
        For i = 1 To 5
            DrawToBuffer Mid(TimeStr, i, 1), i - 1
        Next i
    End If
    
    'Now, copy the ready image from the buffer to the front
    'display
    BitBlt frmClock.hDC, 0, 0, 41, 19, picBuffer.hDC, 0, 0, vbSrcCopy
End Sub

Private Sub DrawToBuffer(Character As String, Position As Long)
    'This function draws a digit/separator to the buffer.
    'The positions are hard-coded because their widths are
    'not always the same.
    Select Case Position
        Case 0
            BitBlt picBuffer.hDC, 0, 0, 10, 19, picSkin.hDC, 0, Character * 19, vbSrcCopy
        Case 1
            BitBlt picBuffer.hDC, 10, 0, 10, 19, picSkin.hDC, 10, Character * 19, vbSrcCopy
        Case 2
            BitBlt picBuffer.hDC, 20, 0, 1, 19, picSkin.hDC, 20, 0, vbSrcCopy
        Case 3
            BitBlt picBuffer.hDC, 21, 0, 10, 19, picSkin.hDC, 21, Character * 19, vbSrcCopy
        Case 4
            BitBlt picBuffer.hDC, 31, 0, 10, 19, picSkin.hDC, 31, Character * 19, vbSrcCopy
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Exit
    End
End Sub

Private Sub Form_Click()
    'Show the menu
    frmMenu.Show
End Sub

Private Sub picTime_Click()
    'Show the menu
    frmMenu.Show
End Sub

