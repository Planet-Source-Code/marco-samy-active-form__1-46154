VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Our Main Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Shape           =   2  'Oval
      Top             =   120
      Width           =   7335
   End
   Begin VB.Image Image5 
      Height          =   1500
      Left            =   960
      Picture         =   "Form1.frx":0000
      Top             =   5040
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Image Image4 
      Height          =   8580
      Left            =   7440
      Picture         =   "Form1.frx":86B4
      Stretch         =   -1  'True
      Top             =   75
      Width           =   75
   End
   Begin VB.Image Image3 
      Height          =   8580
      Left            =   0
      Picture         =   "Form1.frx":8C47
      Stretch         =   -1  'True
      Top             =   75
      Width           =   75
   End
   Begin VB.Image Image2 
      Height          =   75
      Left            =   0
      Picture         =   "Form1.frx":91DA
      Top             =   8640
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Image Image1 
      Height          =   75
      Left            =   120
      Picture         =   "Form1.frx":9EC9
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Stopped As Boolean

Dim Ox, Oy

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'keep current Mouse points in memory
Ox = X: Oy = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'move the form when the moue key is down (1, 2 or 3)
If CBool(Button) Then Move Left + X - Ox, Top + Y - Oy
End Sub

Private Sub Form_GotFocus()
'when this form is  active

Stopped = False

'show bars
Image3.Visible = True: Image4.Visible = True

'activate the title bar
Label1.ForeColor = RGB(255, 255, 255) 'white

'start animation
Timing
End Sub

Function Timing()
Dim OldVal As Long
'loop body
Reloop:
Sleep (20)
If Stopped Then Exit Function

'how many mSecs to waint -------\/
If ((timeGetTime) - OldVal) >= 120 Then GoTo StartAnim: GoTo Reloop

StartAnim:

Static PicLeft As Long
If PicLeft >= ScaleWidth Then PicLeft = 100 Else PicLeft = PicLeft + 100
'draw the upper bar
PaintPicture Image1.Picture, PicLeft, 0
'draw the first part of the upper bar
PaintPicture Image1.Picture, 0, 0, PicLeft, Image1.Height, ScaleWidth - PicLeft
'draw the lower bar
PaintPicture Image2.Picture, PicLeft, Height - Image1.Height
'draw the first part of the lower bar
PaintPicture Image2.Picture, 0, Height - Image1.Height, PicLeft, Image1.Height, ScaleWidth - PicLeft
'draw the middle photo
PaintPicture Image5.Picture, PicLeft, (Height - Image5.Height) / 2
'draw the first part of the middle photo
PaintPicture Image5.Picture, 0, (Height - Image5.Height) / 2, PicLeft, Image5.Height, ScaleWidth - PicLeft

'check events
DoEvents
'return to loop area
GoTo Reloop
End Function

Private Sub Form_Load()
Form2.Show
End Sub

Private Sub Form_LostFocus()
'when the form is not active ( anohter form in this project is active

'stop animation
Stopped = True

'inactivate title bar
Label1.ForeColor = RGB(80, 80, 80) 'dark gray

'hide the vertical bars
Image3.Visible = False: Image4.Visible = False

'clear form
Cls
End Sub
