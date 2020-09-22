VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VRENT: MEMBERS"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewMember.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl 
      Left            =   8280
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   4320
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
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&CANCEL"
      Height          =   495
      Left            =   7200
      TabIndex        =   14
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&BROWSE"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   975
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   5655
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   9
      Top             =   3000
      Width           =   5655
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1200
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2520
      Width           =   5655
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   5655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   6480
      TabIndex        =   15
      Top             =   280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   7080
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OTHERS:"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE:"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT #"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS:"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FULL NAME:"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MEMBER ID:"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   -120
      Picture         =   "frmNewMember.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmNewMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ppath As String

Private Sub Command1_Click()
With cdl

    .DialogTitle = "Open Picture File"
    .CancelError = False
    .Filter = "JPEG FILE|*.jpg |BMP FILE|*.bmp|ALL PICTURE FILES|*.bmp, *.jpg"
    .ShowOpen

End With
    ppath = cdl.FileName
    Image2.Picture = LoadPicture(cdl.FileName)
End Sub

Private Sub Command2_Click()
On Error GoTo hell
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then

    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * FROM members"
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount <= 0 Then
    Else
    Adodc1.Recordset.MoveLast
    End If
    
    
            
            With Adodc1.Recordset
                    
                    .AddNew
                    .Fields(0) = Text1.Text
                    .Fields(1) = Text2.Text
                    .Fields(2) = Text3.Text
                    .Fields(3) = Text4.Text
                    .Fields(4) = Text5.Text
                    .Fields(5) = Text6.Text
                    .Fields(6) = ppath
                    .Update
                    
                    MsgBox "New Member Successfully Added.", vbInformation, "OK"
                    Unload Me
                    
            End With


Else
MsgBox "All fields are required! Please supply information in the first 4 text boxes.", vbInformation, "ERROR"
End If

Exit Sub
hell:
    MsgBox Err.Description

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text5.Text = Date
ppath = ""
End Sub

Private Sub Text1_Change()
Label2.Caption = 5 - Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub Text2_Change()
Text2.Text = UCase(Text2.Text)
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text3_Change()
Text3.Text = UCase(Text3.Text)
Text3.SelStart = Len(Text3.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub Text6_Change()
Text6.Text = UCase(Text6.Text)
Text6.SelStart = Len(Text6.Text)
End Sub
