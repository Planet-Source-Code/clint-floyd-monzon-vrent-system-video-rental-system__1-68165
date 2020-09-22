VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   240
      Picture         =   "frmMessage.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      Shape           =   2  'Oval
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5400
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   5400
      X2              =   5400
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Label message2 
      BackStyle       =   0  'Transparent
      Caption         =   "aasdfasdf"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label message1 
      BackStyle       =   0  'Transparent
      Caption         =   "asdfasdfasdfsadfasdfasdf"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label mainTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   0
      Picture         =   "frmMessage.frx":5348
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5640
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   0
      Picture         =   "frmMessage.frx":1267A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label2.BackColor = vbWhite
Label2.ForeColor = vbRed

If Button = 1 Then 'Ensure that the mouse left-button is clicked and hold
formdrag Me
End If

End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label2.BackColor = vbRed
Label2.ForeColor = vbWhite
End Sub

Private Sub Label3_Click()

End Sub

