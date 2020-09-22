VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmData 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   3795
   ClientTop       =   2265
   ClientWidth     =   11130
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
   ScaleHeight     =   7290
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7275
      ScaleWidth      =   11070
      TabIndex        =   1
      Top             =   0
      Width           =   11130
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Format"
         Height          =   375
         Left            =   7560
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Movie Title"
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID Code"
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   1200
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtsearch 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         MaxLength       =   15
         TabIndex        =   0
         Top             =   1200
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmData.frx":0000
         Height          =   5055
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Movie Database"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   6840
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By:"
         Height          =   270
         Left            =   3600
         TabIndex        =   6
         Top             =   840
         Width           =   945
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         Height          =   615
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   7095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         Height          =   615
         Index           =   0
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
         Height          =   270
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   675
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         Top             =   840
         Width           =   10575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   10815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Viewer"
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
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   10680
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   -120
         Picture         =   "frmData.frx":0015
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11400
      End
      Begin VB.Image Image1 
         Height          =   18000
         Left            =   0
         Picture         =   "frmData.frx":D349
         Top             =   -3480
         Width           =   24000
      End
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tbl As String

Private Sub Command1_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
search = txtsearch.Text 'sets the search value
    
    Adodc1.CommandType = adCmdText 'set the Command Type as text
    Adodc1.RecordSource = "Select * From movies Where " & tbl & " LIKE '" & search & "%'" 'SQL Query
    Adodc1.Refresh 'Refresh the database
    
    'Determine if a record is found on the Database
    If Adodc1.Recordset.EOF Then
        'MsgBox "Record Not Found!"        'if not found on the database
        txtsearch.SetFocus
        SendKeys h
        
        Exit Sub
    Else
    Adodc1.Recordset.MoveFirst
    
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc1.CommandType = adCmdText 'set the Command Type as text
Adodc1.RecordSource = "Select * From movies" 'SQL Query
Adodc1.Refresh 'Refresh the database
    
tbl = "movieCodeNumber"
    
    
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.BackColor = vbWhite
Label1.ForeColor = vbRed
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.BackColor = vbRed
Label1.ForeColor = vbWhite
End Sub

Private Sub Option1_Click()
tbl = "movieCodeNumber"
txtsearch.MaxLength = 15
End Sub

Private Sub Option2_Click()

tbl = "movieTitle"
txtsearch.MaxLength = 255
End Sub

Private Sub Option3_Click()
tbl = "movieStat"
txtsearch.MaxLength = 3
End Sub

Private Sub Option4_Click()
tbl = "movieFormat"
txtsearch.MaxLength = 5
End Sub

Private Sub txtsearch_Change()
Command1_Click
End Sub
