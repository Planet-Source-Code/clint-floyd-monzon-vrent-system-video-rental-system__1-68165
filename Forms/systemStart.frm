VERSION 5.00
Begin VB.Form systemStart 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Picture         =   "systemStart.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   6375
   End
End
Attribute VB_Name = "systemStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim a, b, c, d As String






Label1.Caption = App.Major & "." & App.Minor

On Error GoTo hell

Open App.Path & "\systemdata.xtm" For Input As #1
Input #1, a, b, c, d
Close #1

systemlogos = c
systemname = a
systempass = b
rateperday = d
Label2.Caption = a
On Error GoTo hell
Shell "regsvr32 /s " & App.Path & "\components\Flash9.ocx", vbHide

Exit Sub
hell:
    Unload Me
    frmSetup.Show 1


End Sub

Private Sub Timer1_Timer()
Unload Me
frmLogin.Show
End Sub
    

