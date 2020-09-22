VERSION 5.00
Begin VB.Form frmVoid 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmVoid.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Picture         =   "frmVoid.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Picture         =   "frmVoid.frx":46A0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter System Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   160
      Width           =   3720
   End
   Begin VB.Image Image20 
      Height          =   810
      Left            =   0
      Picture         =   "frmVoid.frx":8D34
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to Expand"
      Top             =   0
      Width           =   8280
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Picture         =   "frmVoid.frx":16068
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmVoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> systempass Then
    MsgBox "Password Mismatched!", vbInformation, "Error"
    Text1.SetFocus
    SendKeys "{HOME}+{END}"
    Exit Sub
Else
    MsgBox "Password Accepted! Please click OK to continue.", vbOKOnly, "Success"
    frmPOS.Picture5.Visible = True
    Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
