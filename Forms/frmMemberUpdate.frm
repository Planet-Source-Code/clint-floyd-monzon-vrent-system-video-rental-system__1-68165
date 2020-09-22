VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMemberUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " MEMBER UPDATE"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmMemberUpdate.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      DataField       =   "memberOthers"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2160
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      DataField       =   "memberContactNumber"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      DataField       =   "memberAddress"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   1080
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label4 
      Caption         =   "Others:"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Contact #:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "You can only change the fields below."
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmMemberUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
With Adodc1.Recordset

    .Fields("memberContactNumber") = Text2.Text
    .Fields("memberAddress") = Text1.Text
    .Fields("memberOthers") = Text3.Text
    .Update
    MsgBox "Successfully Updated the information of " & .Fields(1)
    Unload Me
End With
End Sub

Private Sub Form_Load()
Dim t As String

t = InputBox("Please Enter Membership ID", "VRENT")

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM members WHERE memberID = '" & t & "'"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True Then
   MsgBox "No User Found!"
   Unload Me
Else
Me.Caption = "Edit " & UCase(Adodc1.Recordset.Fields(1)) & "'s information."
End If
End Sub

Private Sub Text1_Change()
Text1.Text = UCase(Text1.Text)
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub Text3_Change()
Text3.Text = UCase(Text3.Text)
Text3.SelStart = Len(Text3.Text)
End Sub
