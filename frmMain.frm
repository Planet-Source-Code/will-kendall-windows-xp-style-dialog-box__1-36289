VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DFE9EF&
   BorderStyle     =   0  'None
   Caption         =   "0"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1560
      Left            =   120
      Picture         =   "frmMain.frx":058A
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":5EDE
      Height          =   855
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Download the XP-Style Button control for the whole window to look XP-style"
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Make sure if you are dimming any button, make Enabled=False"
      Height          =   735
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The code for the visual aspects of the form are in Form_Resize"
      Height          =   855
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   1
      Left            =   5160
      Picture         =   "frmMain.frx":5F6D
      Top             =   1560
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   0
      Left            =   4560
      Picture         =   "frmMain.frx":6671
      Top             =   1560
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   1
      Left            =   5160
      Picture         =   "frmMain.frx":6D75
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   0
      Left            =   4560
      Picture         =   "frmMain.frx":7479
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   2
      Left            =   5280
      Picture         =   "frmMain.frx":7B7D
      Top             =   840
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   1
      Left            =   4920
      Picture         =   "frmMain.frx":8101
      Top             =   840
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   0
      Left            =   4560
      Picture         =   "frmMain.frx":8685
      Top             =   840
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   140
      Picture         =   "frmMain.frx":8C09
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   2
      Left            =   5280
      Picture         =   "frmMain.frx":9193
      Top             =   480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   1
      Left            =   4920
      Picture         =   "frmMain.frx":9717
      Top             =   480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   0
      Left            =   4560
      Picture         =   "frmMain.frx":9C9B
      Top             =   480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   2
      Left            =   5280
      Picture         =   "frmMain.frx":A21F
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   1
      Left            =   4920
      Picture         =   "frmMain.frx":A7A3
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   0
      Left            =   4560
      Picture         =   "frmMain.frx":AD27
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaxButton 
      Height          =   315
      Left            =   2880
      Picture         =   "frmMain.frx":B2AB
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgMinButton 
      Height          =   315
      Left            =   3240
      Picture         =   "frmMain.frx":B82F
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgExit 
      Height          =   315
      Left            =   3600
      Picture         =   "frmMain.frx":BDB3
      Top             =   0
      Width           =   315
   End
   Begin VB.Label lblTest 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows XP-Style Dialog BOX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   135
      Width           =   2460
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Left            =   3240
      Picture         =   "frmMain.frx":C337
      Stretch         =   -1  'True
      Top             =   360
      Width           =   630
   End
   Begin VB.Image imgRight 
      Height          =   375
      Left            =   4320
      Picture         =   "frmMain.frx":C4FB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   45
   End
   Begin VB.Image imgTopRight 
      Height          =   435
      Left            =   3960
      Picture         =   "frmMain.frx":C797
      Top             =   0
      Width           =   195
   End
   Begin VB.Image imgTopLeft 
      Height          =   435
      Left            =   0
      Picture         =   "frmMain.frx":CC63
      Top             =   0
      Width           =   195
   End
   Begin VB.Image imgLeft 
      Height          =   495
      Left            =   4200
      Picture         =   "frmMain.frx":D12F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblTestShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Eraser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.Image imgTop 
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":D3CB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim mouseX, mouseY As Integer


Public Sub Form_Resize()
imgExit.Top = 70
imgBottom.Left = 0
imgLeft.Top = 0
imgRight.Top = 0
imgTop.Top = 0
imgLeft.Left = 0
imgTopRight.Left = Me.ScaleWidth - imgTopRight.Width
imgTopLeft.Left = 0
imgTop.Width = Me.ScaleWidth
imgLeft.Height = Me.ScaleHeight
imgRight.Height = Me.ScaleHeight
imgRight.Left = Me.ScaleWidth - imgRight.Width
imgBottom.Top = Me.ScaleHeight - imgBottom.Height
imgBottom.Width = Me.ScaleWidth
imgExit.Left = Me.ScaleWidth - imgExit.Width - imgLeft.Width - 40
imgMaxButton.Top = imgExit.Top
imgMinButton.Top = imgExit.Top
imgMaxButton.Left = imgExit.Left - imgMinButton.Width - 50
imgMinButton.Left = imgExit.Left - imgMinButton.Width - imgMaxButton.Width - 100


imgLeft.ZOrder
imgRight.ZOrder
imgBottom.ZOrder
imgTopRight.ZOrder
imgTopLeft.ZOrder

imgExit.ZOrder

lblTestShadow.Left = lblTest.Left + 10
lblTestShadow.Top = lblTest.Top + 10
lblTest_Change
lblTest.ZOrder
imgIcon.ZOrder
End Sub

Private Sub imgCorner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mouseX = X
mouseY = Y
End If
End Sub

Private Sub imgCorner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Me.Width + X - mouseX < lblTest.Left + lblTest.Width + 1500 Then Me.Width = lblTest.Left + lblTest.Width + 1501: GoTo nexter
Me.Width = Me.Width + X - mouseX
nexter:
If Me.Height + Y - mouseY < 3000 Then Me.Height = 3001: Exit Sub

Me.Height = Me.Height + Y - mouseY
End If
End Sub

Private Sub imgExit_Click()
End
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExit.Picture = imgDisabled(2).Picture
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExit.Picture = imgEnabled(2).Picture
End Sub

Private Sub imgIcon_DblClick()
If imgExit.Enabled = True Then
imgExit_Click
End If

End Sub

Private Sub imgMaxButton_Click()
Select Case Me.WindowState

    Case 0
    Me.WindowState = 2
    
    Case 2
    Me.WindowState = 0
    
End Select
End Sub

Private Sub imgMaxButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMaxButton.Picture = imgDisabled(0).Picture
End Sub

Private Sub imgMaxButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMaxButton.Picture = imgEnabled(0).Picture
End Sub

Private Sub imgMinButton_Click()
Me.WindowState = 1
End Sub

Private Sub imgMinButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMinButton.Picture = imgDisabled(1).Picture
End Sub

Private Sub imgMinButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMinButton.Picture = imgEnabled(1).Picture
End Sub

Private Sub imgTop_DblClick()
If imgMaxButton.Enabled = True Then
imgMaxButton_Click
End If

End Sub

Private Sub imgTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mouseX = X
mouseY = Y
End If
End Sub

Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Me.Left = Me.Left + X - mouseX
Me.Top = Me.Top + Y - mouseY
End If
End Sub

Private Sub imgTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTop_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgTopRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTop_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblTest_Change()
lblTestShadow.Caption = lblTest.Caption
SetWindowText Me.hwnd, lblTest.Caption
End Sub

Private Sub lblTest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTop_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblTest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTop_MouseMove Button, Shift, X, Y

End Sub
