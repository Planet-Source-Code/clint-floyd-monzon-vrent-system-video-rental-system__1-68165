VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNewUser 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   4350
   ClientTop       =   2565
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   9810
      TabIndex        =   6
      Top             =   0
      Width           =   9870
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "frmNewUser.frx":0000
         Top             =   4080
         Width           =   5535
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   360
         ItemData        =   "frmNewUser.frx":0184
         Left            =   1320
         List            =   "frmNewUser.frx":018E
         TabIndex        =   15
         Text            =   "Employee"
         Top             =   3600
         Width           =   5535
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   50
         PasswordChar    =   "•"
         TabIndex        =   3
         Top             =   3120
         Width           =   8175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   50
         PasswordChar    =   "•"
         TabIndex        =   2
         Top             =   2520
         Width           =   8175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   2040
         Width           =   8175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   0
         Top             =   1560
         Width           =   8175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Level:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype  Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill-up the field. All fields are REQUIRED."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   4515
      End
      Begin VB.Image Image3 
         Height          =   2640
         Left            =   6960
         Picture         =   "frmNewUser.frx":01AB
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New User Sign-Up"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New User Sign-Up"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   1
         Left            =   290
         TabIndex        =   8
         Top             =   160
         Width           =   3315
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   0
         Picture         =   "frmNewUser.frx":467B
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   10200
      End
      Begin VB.Image Image20 
         Height          =   795
         Left            =   0
         Picture         =   "frmNewUser.frx":119AF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10200
      End
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Combo1.Text = "Employee"
End Sub

Private Sub Command1_Click()
'On Error GoTo er1

If Text1.Text <> "" Or Text2.Text <> "" Or Text3.Text <> "" Or Text4.Text <> "" Or Combo1.Text <> "" Then


Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc1.CommandType = adCmdTable
Adodc1.RecordSource = "systemuser"



search = Text2.Text 'sets the search value

Adodc1.CommandType = adCmdText 'set the Command Type as text
Adodc1.RecordSource = "Select * From systemuser Where myuname LIKE '" & search & "%'" 'SQL Query
Adodc1.Refresh 'Refresh the database

'Determine if a record is found on the Database
If Adodc1.Recordset.EOF Then
   'If Not Found then Proceed
    
    With Adodc1.Recordset
    .AddNew
    .Fields(0) = Text1.Text
    .Fields(1) = Text2.Text
    .Fields(2) = Text3.Text
    .Fields(3) = Combo1.Text
    .Update
    MsgBox "New User added successfully to database."
    Unload Me
    End With





Else
   'If Found then Go Back to change username.
   MsgBox "The username you've typed is already existing." & vbCr & "Please try another username."
   Text2.SetFocus
   SendKeys h 'highlight the username box.
End If


Else



    mBox "Error", "Some fields are empty.", "All fields are required!", "error"




End If

Exit Sub
er1:
   MsgBox Err.Description

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text4_Change()
If Text4.Text = Text3.Text Then
   Text4.BackColor = &HC0FFC0
   Text4.ForeColor = vbBlack
Else
   'Text4.BackColor = &H80FF&
   'Text4.ForeColor = vbWhite
End If
End Sub
