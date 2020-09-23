VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cool Clock Menu"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   157
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   1920
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Choose a new skin..."
      Filter          =   "Image Files (*.bmp;*.gif;*.jpg;*.jpeg)|*.bmp;*.gif;*.jpg;*.jpeg"
      FontName        =   "Times New Roman"
      FontSize        =   12
      Max             =   16
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H0000FF00&
      Index           =   3
      Visible         =   0   'False
      X1              =   0
      X2              =   157
      Y1              =   146
      Y2              =   146
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H0000FF00&
      Index           =   2
      Visible         =   0   'False
      X1              =   0
      X2              =   157
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H0000FF00&
      Index           =   1
      Visible         =   0   'False
      X1              =   156
      X2              =   156
      Y1              =   0
      Y2              =   146
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H0000FF00&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   146
   End
   Begin VB.Line lneSeparator 
      BorderColor     =   &H0000FF00&
      Index           =   1
      Visible         =   0   'False
      X1              =   0
      X2              =   157
      Y1              =   125
      Y2              =   125
   End
   Begin VB.Line lneSeparator 
      BorderColor     =   &H0000FF00&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   157
      Y1              =   99
      Y2              =   99
   End
   Begin VB.Label lblSkinType 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Tag             =   "» Skin Type:"
      Top             =   0
      Width           =   2355
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Tag             =   " » Text"
      Top             =   195
      Width           =   2355
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Tag             =   "» Exit"
      Top             =   1950
      Width           =   2355
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Tag             =   "» About..."
      Top             =   1560
      Width           =   2355
   End
   Begin VB.Label lblSkinImage 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Tag             =   "  » Set Skin Image..."
      Top             =   1170
      Width           =   2355
   End
   Begin VB.Label lblGraphic 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Tag             =   " » Graphic"
      Top             =   975
      Width           =   2355
   End
   Begin VB.Label lblFont 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Tag             =   "  » Set Font..."
      Top             =   780
      Width           =   2355
   End
   Begin VB.Label lblBackColor 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Tag             =   "  » Set Background Color..."
      Top             =   585
      Width           =   2355
   End
   Begin VB.Label lblTextColor 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Tag             =   "  » Set Text Color..."
      Top             =   390
      Width           =   2355
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_GotFocus()
    'Position the menu in the place where the user
    'clicked
    Dim CurPos As POINTAPI
    GetCursorPos CurPos
    Me.Left = CurPos.x * 15 - Me.Left
    Me.Top = CurPos.y * 15 - Me.Top
    
    'Either draw the skin's menu (SkinType=Image) or
    'set the background/text colors and use labels
    If SkinType = stImage Then
        'Set all of the labels' captions to nothing
        '(because the text is already in the skin image)
        'and hide the lines
        For Each ctlControl In frmMenu.Controls
            If TypeOf ctlControl Is Label Then
                ctlControl.Caption = ""
            ElseIf TypeOf ctlControl Is Line Then
                ctlControl.Visible = False
            End If
        Next ctlControl
        
        'I *WOULD* use BitBlt here, but for some odd
        'reason it doesn't work. So I'll just stick to
        'good old PaintPicture (and slow! But we don't
        'need speed in this case :p  )
        
        'BitBlt picBack.hDC, 0, 0, 157, 147, _
          frmClock.picSkin.hDC, 41, 0, vbSrcCopy
        
        On Error GoTo InvalidSkin
        
        frmMenu.PaintPicture frmClock.picSkin.Picture, _
          0, 0, 157, 147, 41, 0, 157, 147, vbSrcCopy
    ElseIf SkinType = stText Then
        'Show the labels' captions and the lines
        '(NOTE: I kept each of the labels' captions
        'in their Tag property, because if I make them
        'invisible, they would not receive Click events.)
        For Each ctlControl In frmMenu.Controls
            If TypeOf ctlControl Is Label Then
                ctlControl.Caption = ctlControl.Tag
            ElseIf TypeOf ctlControl Is Line Then
                ctlControl.Visible = True
            End If
        Next ctlControl
        
        'Also, clear the skin image
        frmMenu.Cls
    End If
    
    Exit Sub
    
InvalidSkin: 'There was an error with PaintPicture: this
             'is not a valid skin. Switch to Text Skin.
    MsgBox "This is not a valid skin: switching to Text Skin.", vbOKOnly, "Cool Clock - Error"
    lblText_Click
    Form_GotFocus
End Sub

Private Sub Form_LostFocus()
    'Hide, like in regular menus
    Me.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Hide, like in regular menus
    Me.Hide
End Sub

Private Sub lblText_Click()
    'Set the skin type to Text and redraw accordingly
    SkinType = stText
    Form_GotFocus
    frmClock.ApplySettings
    
    'Write to INI file and hide menu
    WriteINI "General", "SkinType", "Text", INIFilePath
    Me.Hide
End Sub

Private Sub lblTextColor_Click()
    'Show the Choose (Text) Color dialog
    cd.Flags = cdlCCFullOpen Or cdlCCRGBInit
    cd.Color = TextColor
    cd.ShowColor
    TextColor = cd.Color
    frmClock.ApplySettings
    
    'Write to INI file and hide menu
    WriteINI "TextSkin", "TextColor", LongToRGB(TextColor), INIFilePath
    Me.Hide
End Sub

Private Sub lblBackColor_Click()
    'Show the Choose (Back) Color dialog
    cd.Flags = cdlCCFullOpen Or cdlCCRGBInit
    cd.Color = BackgroundColor
    cd.ShowColor
    BackgroundColor = cd.Color
    frmClock.ApplySettings
    
    'Write to INI file and hide menu
    WriteINI "TextSkin", "BackgroundColor", LongToRGB(BackgroundColor), INIFilePath
    Me.Hide
End Sub

Private Sub lblFont_Click()
    'Show the Choose Font dialog
    cd.Flags = cdlCFBoth Or cdlCFLimitSize
    cd.FontName = TextFontName
    cd.FontSize = TextSize
    cd.ShowFont
    TextFontName = cd.FontName
    TextSize = cd.FontSize
    frmClock.ApplySettings
    
    'Write to INI file and hide menu
    WriteINI "TextSkin", "TextFontName", TextFontName, INIFilePath
    WriteINI "TextSkin", "TextSize", TextSize, INIFilePath
    Me.Hide
End Sub

Private Sub lblGraphic_Click()
    'Set the skin type to Image and redraw accordingly
    SkinType = stImage
    Form_GotFocus
    
    'Write to INI file and hide menu
    WriteINI "General", "SkinType", "Image", INIFilePath
    Me.Hide
End Sub

Private Sub lblSkinImage_Click()
    Dim TempStrArr() As String
    
    'Show the Choose (Skin) File dialog
    MsgBox "Remember: You can only choose files that are in the ""Cool Clock\Skins"" folder.", vbOKOnly, "Cool Clock"
    cd.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    cd.ShowOpen
    
    'Do nothing if there's no file selected
    If cd.FileName = "" Then GoTo HideMenu
    
    'Get only the last part of the file name (take out
    'the path)
    TempStrArr = Split(cd.FileName, "\")
    SkinName = TempStrArr(UBound(TempStrArr))
    
    'Apply the new settings
    frmClock.ApplySettings
    Form_GotFocus
    
    'Write to INI file and hide menu
    WriteINI "ImageSKin", "SkinName", SkinName, INIFilePath
HideMenu:
    Me.Hide
End Sub

Private Sub lblAbout_Click()
    'Show the About dialog box
    frmAbout.Show
    Me.Hide
End Sub

Private Sub lblExit_Click()
    'Exit program
    End
End Sub

Public Function LongToRGB(LongColor As Long) As String
    'This function extracts the RGB values of a color and
    'returns a string with their individual values separated
    'by spaces (ready to write to the INI file).
    
    Dim R As Byte, G As Byte, B As Byte
    
    R = LongColor And 255
    G = (LongColor And 65280) \ 256
    B = (LongColor And 16711680) \ 65535
    
    LongToRGB = R & " " & G & " " & B
End Function
