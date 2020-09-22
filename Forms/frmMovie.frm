VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMovie 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   4350
   ClientTop       =   2565
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   9720
      TabIndex        =   13
      Top             =   0
      Width           =   9780
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   480
         Top             =   6600
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
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   9135
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   2295
            Left            =   6960
            TabIndex        =   29
            Top             =   1320
            Width           =   1815
            Begin VB.CommandButton Command2 
               Caption         =   "&Cancel"
               Height          =   495
               Left            =   240
               TabIndex        =   12
               Top             =   1320
               Width           =   1335
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&Add"
               Default         =   -1  'True
               Height          =   495
               Left            =   240
               TabIndex        =   11
               Top             =   600
               Width           =   1335
            End
            Begin VB.Shape Shape1 
               BackStyle       =   1  'Opaque
               FillColor       =   &H00E0E0E0&
               FillStyle       =   0  'Solid
               Height          =   2175
               Left            =   0
               Top             =   120
               Width           =   1815
            End
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   10
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   9
            Top             =   3240
            Width           =   2535
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            MaxLength       =   100
            TabIndex        =   8
            Top             =   2400
            Width           =   2535
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   7
            Text            =   "IN"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.ComboBox Combo3 
            Height          =   390
            ItemData        =   "frmMovie.frx":0000
            Left            =   3840
            List            =   "frmMovie.frx":0013
            TabIndex        =   6
            Text            =   "1"
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MaxLength       =   100
            TabIndex        =   5
            Top             =   3120
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            Height          =   390
            Left            =   240
            TabIndex        =   4
            Top             =   2160
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMovie.frx":0026
            Left            =   240
            List            =   "frmMovie.frx":0033
            TabIndex        =   3
            Text            =   "VCD"
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MaxLength       =   100
            TabIndex        =   2
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Price:"
            Height          =   270
            Index           =   8
            Left            =   6960
            TabIndex        =   28
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Price:"
            Height          =   270
            Index           =   7
            Left            =   3840
            TabIndex        =   27
            Top             =   2880
            Width           =   1410
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Stocks:"
            Height          =   270
            Index           =   6
            Left            =   3840
            TabIndex        =   26
            Top             =   2040
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   270
            Index           =   5
            Left            =   3840
            TabIndex        =   25
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Days:"
            Height          =   270
            Index           =   4
            Left            =   3840
            TabIndex        =   24
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Date:"
            Height          =   270
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Top             =   2760
            Width           =   1380
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year:"
            Height          =   270
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   1800
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Format:"
            Height          =   270
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category:"
            Height          =   270
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         MaxLength       =   255
         TabIndex        =   1
         Top             =   1800
         Width           =   6255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   15
         TabIndex        =   0
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movie Title"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   3240
         TabIndex        =   18
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movie ID (Barcode Number)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   2520
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill-up the following fields below."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movie Management"
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
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   120
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movie Management"
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
         Index           =   0
         Left            =   390
         TabIndex        =   14
         Top             =   160
         Width           =   3705
      End
      Begin VB.Image Image20 
         Height          =   795
         Left            =   0
         Picture         =   "frmMovie.frx":0046
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10200
      End
   End
End
Attribute VB_Name = "frmMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
SendKeys h
End Sub

Private Sub Combo2_Click()
SendKeys h
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub Combo3_Click()
SendKeys h
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub Command1_Click()
On Error GoTo er1

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc1.CommandType = adCmdTable
Adodc1.RecordSource = "movies"




search = Text1.Text 'sets the search value

Adodc1.CommandType = adCmdText 'set the Command Type as text
Adodc1.RecordSource = "Select * From movies Where movieCodeNumber LIKE '" & search & "%'" 'SQL Query
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
    .Fields(4) = Combo2.Text
    .Fields(5) = Text4.Text
    .Fields(6) = Combo3.Text
    .Fields(7) = Text5.Text
    .Fields(8) = Text6.Text
    .Fields(9) = Text6.Text
    .Fields(10) = Text7.Text
    .Fields(11) = Text8.Text
    .Update
    mBox "Success", "Movie Successfully Added to Database", "", "info"
    'MsgBox "Movie Entry successfully added to database."
    Unload Me
    End With





Else
   'If Found then Go Back to change username.
   mBox "Error", "The ID you've typed is already existing.", "", "error"
   'MsgBox "The ID you've typed is already existing." & vbCr & "Please try another username."
   Text1.SetFocus
   SendKeys h 'highlight the username box.
End If


Exit Sub
er1:
MsgBox Err.Description

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer

Combo2.Text = Year(Date)

For x = Year(Date) To 1980 Step -1
Combo2.AddItem x
Next x

SendKeys "{HOME}+{END}"
End Sub

Private Sub Text1_Click()
SendKeys h
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub Text2_Change()
Text2.Text = UCase(Text2.Text)
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text2_Click()
SendKeys h
End Sub

Private Sub Text3_Change()
Text3.Text = UCase(Text3.Text)
Text3.SelStart = Len(Text3.Text)
End Sub

Private Sub Text3_Click()
SendKeys h
End Sub

Private Sub Text4_Click()
SendKeys h
End Sub

Private Sub Text6_Click()
SendKeys h
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
checknum KeyAscii 'call the checknum function located at the Global_module
End Sub

Private Sub Text7_Click()
SendKeys h
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub Text8_Click()
SendKeys h
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub
