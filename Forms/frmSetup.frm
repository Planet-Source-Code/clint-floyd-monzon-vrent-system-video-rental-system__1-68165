VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Setup"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Text            =   "10"
      Top             =   4440
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   360
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox systempassword2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "•"
      TabIndex        =   2
      Top             =   3240
      Width           =   5655
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
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox systempassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   2160
      Width           =   5655
   End
   Begin VB.TextBox shopname 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Rate: (PerDay)"
      Height          =   615
      Left            =   360
      TabIndex        =   14
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "System Logo:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-type System Password:"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "This password is the Global password for the system. You will need this when you are voiding a product."
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "System Password:"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2230
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Video Shop Name:"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3855
      Left            =   240
      Top             =   1440
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Thank you for using VRent system. Please fill up the necessary fields needed to run this software. All fields are required."
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   7815
   End
   Begin VB.Image Image4 
      Height          =   795
      Left            =   1920
      Picture         =   "frmSetup.frx":1601A
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   6360
   End
   Begin VB.Image Image3 
      Height          =   795
      Left            =   -5040
      Picture         =   "frmSetup.frx":306B6
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   15360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to VRent Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3885
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   -120
      Picture         =   "frmSetup.frx":702FA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8655
   End
   Begin VB.Image Image2 
      Height          =   4860
      Left            =   -240
      Picture         =   "frmSetup.frx":95B3E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   8940
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n, p, l, r As String

If systempassword.Text = "" And shopname.Text = "" And systempassword2.Text = "" And Text1.Text = "" Then
    MsgBox "All Fields are required!", vbInformation, "Error"
Else
If systempassword2.Text <> systempassword.Text Then
    MsgBox "Password not Matched!", vbCritical, "Error"
Else

    'Write the settings to a text file
    n = shopname.Text
    p = systempassword.Text
    l = Text1.Text
    r = Text2.Text
    Open App.Path & "\systemdata.xtm" For Output As #1
    Write #1, n, p, l, r
    Close #1
    
    MsgBox "Settings Saved.", vbInformation, "Message"
    Unload Me
    
    If mainstat <> 1 Then
        systemStart.Show
    Else
        Unload Me
    End If
    
End If
End If
End Sub

Private Sub Command2_Click()
With cdl
        
        .DialogTitle = "Open File"
        .Filter = "Bitmap Files|*.bmp|JPEG Files|*.jpg"
        .ShowOpen
        
End With
        Text1.Text = cdl.FileName
End Sub

Private Sub Form_Load()
Dim a, b, c As String
On Error Resume Next
Open App.Path & "\systemdata.xtm" For Input As #1
Input #1, a, b, c

shopname.Text = a
systempassword.Text = b
Text1.Text = c

Close #1
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub
