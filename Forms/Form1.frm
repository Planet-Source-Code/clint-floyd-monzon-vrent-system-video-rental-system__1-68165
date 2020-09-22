VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Member ID"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3960
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.CommandButton Command1 
      Caption         =   "&ACCEPT"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Picture         =   "Form1.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Member &Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2280
      Picture         =   "Form1.frx":46A0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Member &ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   480
      Picture         =   "Form1.frx":8D34
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   45
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaxLength       =   5
      TabIndex        =   0
      Top             =   960
      Width           =   6375
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1560
      Top             =   1320
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc2"
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   6375
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   6840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   6375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Members Area"
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
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1770
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Member by:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Image frmMember 
      Height          =   810
      Left            =   0
      Picture         =   "Form1.frx":D3C8
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to Expand"
      Top             =   0
      Width           =   8280
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   0
      Picture         =   "Form1.frx":1A6FC
      Stretch         =   -1  'True
      Top             =   480
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mm As String

Private Sub Command1_Click()
Dim bcount As Integer

bcount = Adodc2.Recordset.RecordCount
borrowerID = Adodc1.Recordset.Fields(0)
borrowerName = Adodc1.Recordset.Fields(1)




If bcount <= 0 Then
   frmPOS.Text1.Locked = False
            
Else
   If bcount >= 5 Then
         
        frmPOS.cart.TextRTF = ""
        frmPOS.lblTotal.Caption = "0"
        frmPOS.lblTotalShadow.Caption = "0"

        frmPOS.lblItemCount.Caption = "0"

        frmPOS.cart.SelStart = Len(frmPOS.cart.Text)
        frmPOS.lblMemberName.Caption = ""
        frmPOS.Text1.Text = "Press Space..."
        frmPOS.Text1.Locked = True
        MsgBox "You've Reached the maximum limit of 5.", vbInformation + vbOKOnly, "VRENT"
        Unload Me
   Else
        frmPOS.lblRentCount.Caption = Adodc2.Recordset.RecordCount
        frmPOS.lblMemberName.Caption = Adodc2.Recordset.Fields(1) & ", you have  " & Adodc2.Recordset.RecordCount & " unreturned movie(s)."
   End If
   
End If
        

frmPOS.Text1.Locked = False

Unload Me

End Sub

Private Sub Form_Load()
mm = "memberID"
End Sub

Private Sub Label5_Click()
Unload Me

frmPOS.Text1.Locked = True
frmPOS.Text1 = "Press Space..."
frmPOS.lblMemberName.Caption = ""
End Sub

Private Sub Option1_Click()
mm = "memberID"
Text1.Text = ""
Text1.MaxLength = 5
Text1.SetFocus


End Sub

Private Sub Option2_Click()
mm = "memberFullName"
Text1.Text = ""
Text1.MaxLength = 255
Text1.SetFocus
End Sub

Private Sub Text1_Change()
Dim search As String

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
search = Text1.Text 'sets the search value
Adodc1.CommandType = adCmdText 'set the Command Type as text
Adodc2.CommandType = adCmdText 'set the Command Type as text
Adodc1.RecordSource = "Select * From members Where " & mm & " LIKE '" & search & "'" 'SQL Query
Adodc2.RecordSource = "Select * From borrowed Where " & mm & " LIKE '" & search & "'" 'SQL Query
Adodc1.Refresh 'Refresh the database
Adodc2.Refresh 'Refresh the database

If Text1.Text = "" Then

Label3.Caption = ""
Label4.Caption = ""
Command1.Enabled = False
Exit Sub
End If

    If Adodc1.Recordset.EOF Then
        'If no User Found on System
        Label3.Caption = "No Such USER Matched."
        Label4.Caption = ""
        Command1.Enabled = False
        Exit Sub
    Else
    Adodc1.Recordset.MoveFirst
    Label3.Caption = "Member ID     : " & Adodc1.Recordset.Fields(0)
    Label4.Caption = "Member Name: " & Adodc1.Recordset.Fields(1)
    'On Error Resume Next
    frmPOS.lblMemberName.Caption = "Welcome " & Adodc1.Recordset.Fields(1)
    Command1.Enabled = True
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Option1.Value = True Then
checknum KeyAscii
Else
End If
End Sub
