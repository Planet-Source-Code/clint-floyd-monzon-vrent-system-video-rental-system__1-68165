VERSION 5.00
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select type"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Close"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Member List"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrowers List"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Movie Master List"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   -360
      Picture         =   "frmReports.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PopupMenu frmImages.mnumovielist, , Command1.Left, Command2.Top + Command2.Height

End Sub

Private Sub Command2_Click()
PopupMenu frmImages.mnuBorrowed, , Command2.Left, Command2.Top + Command2.Height
End Sub

Private Sub Command3_Click()
PopupMenu frmImages.mnumem, , Command3.Left, Command3.Top + Command3.Height
End Sub

Private Sub Command4_Click()
Dim sqlstr As String

Set movRS = New ADODB.Recordset
sqlstr = "SELECT * FROM movies where movieStat = 'OUT'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly
rptMov.Caption = "Movies"
Set rptMov.DataSource = movRS
rptMov.Show vbModal

End Sub

Private Sub Command5_Click()
Dim sqlstr As String

Set movRS = New ADODB.Recordset
sqlstr = "SELECT * FROM movies where movieStat = 'IN'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly
rptMov.Caption = "Movies"
Set rptMov.DataSource = movRS
rptMov.Show vbModal

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call dbConnect
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmImages
Unload rptMov
Unload rptborrower
Unload rptMembers
Unload Me
End Sub
