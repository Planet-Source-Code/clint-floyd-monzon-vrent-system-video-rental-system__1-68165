VERSION 5.00
Begin VB.Form frmSystempassword 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmSystempassword.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter System Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   5205
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   -360
      Picture         =   "frmSystempassword.frx":1601A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15465
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   -120
      Picture         =   "frmSystempassword.frx":3B85E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   7575
   End
End
Attribute VB_Name = "frmSystempassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b
If Text1.Text <> systempass Then
    MsgBox "Incorrect Password!", vbInformation, "Error"
Else
    a = mainuser
    b = Date & " -- " & Time
    MsgBox "For Security Purposes, this transaction will be log.", vbInformation, "Success"
    
    
        
        Open App.Path & "\data\VRSLOG.vrs" For Append As #1
        Print #1, a, b
        Close #1
    
    
    
        
    Select Case myforms
            Case "frmSettings"
                    frmSettings.Show 1
                    Unload Me
            Case "frmNewUser"
                    frmNewUser.Show 1
                    Unload Me
    End Select
    Unload Me
    
    'frmNewUser.Show 1
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
