VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "User Login"
   ClientHeight    =   4470
   ClientLeft      =   3885
   ClientTop       =   0
   ClientWidth     =   7005
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox mname 
      DataField       =   "myname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   13560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox mpassword 
      DataField       =   "mypassword"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   13560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox muname 
      DataField       =   "myuname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   13560
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   13560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   120
      Picture         =   "frmLogin.frx":1601A
      Stretch         =   -1  'True
      ToolTipText     =   "User Menu"
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label myVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v 0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3600
      TabIndex        =   10
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   5520
      Picture         =   "frmLogin.frx":1B24E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to VRent System! Please fill out the field below to start."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   6495
   End
   Begin VB.Image Image4 
      Height          =   1275
      Left            =   -120
      Picture         =   "frmLogin.frx":20D40
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   7080
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   0
      Picture         =   "frmLogin.frx":3B3DC
      Top             =   0
      Width           =   7500
   End
   Begin VB.Image Image7 
      Height          =   3420
      Left            =   0
      Picture         =   "frmLogin.frx":5A62C
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   6945
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmLogin"
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
'                L O G I N    F O R M                     *
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Private Sub Form_Load()
myVersion.Caption = "v. " & App.Major & "." & App.Minor & " for Demonstration"
'Timer1.Enabled = True
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
formdrag Me  'to drag the form even without toolbox
End Sub


Private Sub cmdCancel_Click()
If mainstat <> 1 Then 'to ensure that it is a log-off or new login
      End
Else

   Unload Me          'Unload or Exit this Window.
   frmMain.Show       'Show the Main Form
End If
End Sub

Private Sub cmdOK_Click()

Dim clientname As String

'Connect to the database
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"

'Returns if the main form is visible or not
If mainstat <> 1 Then

    search = Text1.Text 'sets the search value
    
    Adodc1.CommandType = adCmdText 'set the Command Type as text
    Adodc1.RecordSource = "Select * From systemuser Where myuname LIKE '" & search & "%'" 'SQL Query
    Adodc1.Refresh 'Refresh the database
    
    'Determine if a record is found on the Database
    If Adodc1.Recordset.EOF Then
        mBox "Error", "User Not Found!", "Please Try Again.", "error"
        'MsgBox "No such user found!"        'if not found on the database
        Text1.SetFocus
        SendKeys h
        Exit Sub
    Else
   Adodc1.Recordset.MoveFirst           'Move the database pointer to the first
   If Text2.Text = mpassword.Text Then  'Determine if password is match to the db
     
    mainuser = mname.Text                  'set the Currently logon on the system
    Text1.Visible = False
    Text2.Visible = False
    Label1(0).Visible = False
    Label1(1).Visible = False
    Label1(2).Visible = False
    
    Label2.Visible = True
    Label2.Caption = Adodc1.Recordset.Fields(3)   'Show the user level of the user
    mBox "Welcome", "Welcome " & Adodc1.Recordset.Fields(0), "", "info"
    
    'MsgBox "Welcome " & mname.Text & "!"   'Welcomes the new logged in user
    frmMain.Show                           'Show the Main Form
    Unload Me                              'Unload this Window.
    
   Else
       mBox "Error", "Password Not Match!", "Please Try Again.", "error"
      'MsgBox "Wrong Password"              'Prompt the user for the wrong password
      Text2.SetFocus
      SendKeys h
   End If
End If

Else
   Unload frmMain                          'Unload the Main form for the new user
   search = Text1.Text                     'sets the search value


Adodc1.CommandType = adCmdText             'Set the Command Type as Text
Adodc1.RecordSource = "Select * From systemuser Where myuname LIKE '" & search & "%'"  'SQL Query string
Adodc1.Refresh                             'Refreshes the Database

If Adodc1.Recordset.EOF Then
    mBox "Error", "User Not Found!", "Please Try Again.", "error"
    'MsgBox "No Record Found!"              'If no record Found
    Text1.SetFocus
    SendKeys h
    Exit Sub

Else
   Adodc1.Recordset.MoveFirst
   If Text2.Text = mpassword.Text Then     'Matching the Passwords
     
    mainuser = mname.Text                  'Set the logged in user Globally
    Text1.Visible = False
    Text2.Visible = False
    Label1(0).Visible = False
    Label1(1).Visible = False
    Label1(2).Visible = False
    
    Label2.Visible = True
    Label2.Caption = Adodc1.Recordset.Fields(3)   'Show the user level of the user
    mBox "Welcome", "Welcome Administrator!", "Log-In Date: " & Date, "error"
    'MsgBox "Welcome " & mname.Text & "!"   'Welcomes the New User
    frmMain.Show                           'Shows the Main Form
    Unload Me                              'Exits this Window
    
   Else
        mBox "Error", "Password Not Match!", "Please Try Again.", "error"
      'MsgBox "Wrong Password"              'Prompt the user for Wrong Password
      Text2.SetFocus
      SendKeys h
   End If
End If
End If

End Sub

Private Sub Image5_Click()
PopupMenu frmImages.addnewuser, , Image5.Left + 200, Image5.Top + Image5.Height
End Sub

Private Sub Text2_Click()
SendKeys h

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()
Dim x As Integer

    
    
If mainstat <> 1 Then 'to ensure that it is a log-off or new login
   For x = 3645 To 0 Step -10
Me.Top = x
    If x <= 0 Then
    Timer2.Enabled = False
    End If
Next x
    Timer2.Enabled = False
    End
Else
For x = 3645 To 0 Step -10
Me.Top = x
    If x <= 0 Then
    Timer2.Enabled = False
    End If
Next x
    Timer2.Enabled = False
   Unload Me          'Unload or Exit this Window.
   frmMain.Show       'Show the Main Form
End If

End Sub
