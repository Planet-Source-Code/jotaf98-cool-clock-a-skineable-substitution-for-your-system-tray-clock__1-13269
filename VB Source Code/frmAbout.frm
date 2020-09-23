VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Cool Clock"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblOKFront 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H80000016&
      Height          =   225
      Index           =   1
      Left            =   1350
      TabIndex        =   12
      Top             =   2745
      Width           =   2025
   End
   Begin VB.Label lblOKFront 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H80000010&
      Height          =   225
      Index           =   0
      Left            =   1335
      TabIndex        =   10
      Top             =   2730
      Width           =   2025
   End
   Begin VB.Label lblOKBack 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   1335
      TabIndex        =   11
      Top             =   2655
      Width           =   2025
   End
   Begin VB.Line lneOK 
      BorderColor     =   &H8000000F&
      Index           =   2
      X1              =   224
      X2              =   224
      Y1              =   176
      Y2              =   200
   End
   Begin VB.Line lneOK 
      BorderColor     =   &H8000000F&
      Index           =   0
      X1              =   88
      X2              =   88
      Y1              =   176
      Y2              =   200
   End
   Begin VB.Line lneOK 
      BorderColor     =   &H8000000F&
      Index           =   3
      X1              =   88
      X2              =   224
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line lneOK 
      BorderColor     =   &H8000000F&
      Index           =   1
      X1              =   88
      X2              =   224
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Page:"
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   615
      TabIndex        =   9
      Top             =   1935
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   615
      TabIndex        =   8
      Top             =   1455
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   8
      X2              =   305
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   39
      X2              =   273
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblWebSite 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.firehawk.cjb.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Web Page:"
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "jotaf98@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "E-mail:"
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   40
      X2              =   272
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   9
      X2              =   304
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "by Jotaf98 (João F. S. Henriques)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   375
      Left            =   735
      TabIndex        =   7
      Top             =   660
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cool Clock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   615
      Left            =   135
      TabIndex        =   6
      Top             =   15
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "by Jotaf98 (João F. S. Henriques)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   645
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Cool Clock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This is for e-mailing/executing URLs
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub lblEmail_Click()
    'Send me an e-mail :)
    ShellExecute 0&, vbNullString, "mailto:jotaf98@hotmail.com", vbNullString, "C:\", 1
End Sub

Private Sub lblWebSite_Click()
    'Visit my web page
    ShellExecute 0&, vbNullString, "http://www.firehawk.cjb.net", vbNullString, "C:\", 1
End Sub


' -- NOTE --
'
'All of the following code is for the "OK" button to work
'like the ones in MS Internet Explorer 5 (flat buttons), so
'I didn't comment them much. You'll have to look at each
'object's name to understand it.

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then Exit Sub 'Clicking
    
    'Remove button border
    lneOK(0).BorderColor = vbButtonFace
    lneOK(1).BorderColor = vbButtonFace
    lneOK(2).BorderColor = vbButtonFace
    lneOK(3).BorderColor = vbButtonFace
End Sub

Private Sub lblOKBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then Exit Sub 'Clicking
    
    'Make the flat button border
    lneOK(0).BorderColor = Label7.ForeColor 'Button Light Shadow
    lneOK(1).BorderColor = Label7.ForeColor 'Button Light Shadow
    lneOK(2).BorderColor = Label1.ForeColor 'Button Shadow
    lneOK(3).BorderColor = Label1.ForeColor 'Button Shadow
End Sub

Private Sub lblOKFront_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then Exit Sub 'Clicking
    
    'Make the flat button border
    lneOK(0).BorderColor = Label7.ForeColor 'Button Light Shadow
    lneOK(1).BorderColor = Label7.ForeColor 'Button Light Shadow
    lneOK(2).BorderColor = Label1.ForeColor 'Button Shadow
    lneOK(3).BorderColor = Label1.ForeColor 'Button Shadow
End Sub

Private Sub lblOKFront_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Make the inverted flat button border (pressed)
    lneOK(0).BorderColor = Label1.ForeColor 'Button Shadow
    lneOK(1).BorderColor = Label1.ForeColor 'Button Shadow
    lneOK(2).BorderColor = Label7.ForeColor 'Button Light Shadow
    lneOK(3).BorderColor = Label7.ForeColor 'Button Light Shadow
    
    '"Press" the button's caption too
    lblOKFront(0).Left = lblOKFront(0).Left + 1
    lblOKFront(0).Top = lblOKFront(0).Top + 1
    lblOKFront(1).Left = lblOKFront(1).Left + 1
    lblOKFront(1).Top = lblOKFront(1).Top + 1
End Sub

Private Sub lblOKFront_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Make the inverted flat button border (pressed)
    lneOK(0).BorderColor = Label7.ForeColor 'Button Light Shadow
    lneOK(1).BorderColor = Label7.ForeColor 'Button Light Shadow
    lneOK(2).BorderColor = Label1.ForeColor 'Button Shadow
    lneOK(3).BorderColor = Label1.ForeColor 'Button Shadow
    
    '"Press" the button's caption too
    lblOKFront(0).Left = lblOKFront(0).Left - 1
    lblOKFront(0).Top = lblOKFront(0).Top - 1
    lblOKFront(1).Left = lblOKFront(1).Left - 1
    lblOKFront(1).Top = lblOKFront(1).Top - 1
    
    '-- This is where the real action takes place --
    
    'Close
    Unload Me
End Sub
