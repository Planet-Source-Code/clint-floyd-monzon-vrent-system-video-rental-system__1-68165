VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   2970
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmAbout.frx":2BF64
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v."
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
      Height          =   285
      Left            =   2130
      TabIndex        =   1
      Top             =   1320
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Created: February 2007"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'*V-Rent System; a Video rental Solution                  *
'*________________________________________________________*
'*Created by:                                             *
'*          clint monzon                                  *
'*          charlie peñafiel                              *
'*          raiza costelo                                 *
'*          cherry lyn molina                             *
'*Date Started/Edited:                                    *
'*          February 14, 2007                             *
'**********************************************************
'                A B O U T   F O R M                      *
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Private Sub Form_Load()


Label2.Caption = "v. " & App.Major & "." & App.Minor & " for Demonstration"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then

    formdrag Me

End If

Label3.BackColor = vbWhite
Label3.ForeColor = vbRed


End Sub

Private Sub Timer1_Timer()
Label4.Caption = Label4.Caption + 1
Text1.SelStart = Val(Label4.Caption)

If Label4.Caption >= Label3.Caption Then
   Label4.Caption = "1"
End If

End Sub

Private Sub Label4_Click()
End Sub

Private Sub Label3_Click()
Unload Me 'Unloads this form
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.BackColor = vbRed
Label3.ForeColor = vbWhite
End Sub
