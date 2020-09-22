VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReturn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VRENT"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmReturn.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1485
      Visible         =   0   'False
      Width           =   8895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmReturn.frx":1601A
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "movieCodeNumber"
         Caption         =   "CODE NUMBER"
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
         DataField       =   "movieTitle"
         Caption         =   "TITLE"
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
      BeginProperty Column02 
         DataField       =   "dateBorrowed"
         Caption         =   "BORROW DATE"
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
      BeginProperty Column03 
         DataField       =   "dateReturned"
         Caption         =   "DUE DATE"
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
      BeginProperty Column04 
         DataField       =   "memberID"
         Caption         =   "memberID"
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
      BeginProperty Column05 
         DataField       =   "memberName"
         Caption         =   "Name"
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
         MarqueeStyle    =   5
         ScrollBars      =   2
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4500.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   7080
      Top             =   6960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   9240
      TabIndex        =   24
      Top             =   6960
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   2400
      TabIndex        =   22
      Top             =   7800
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60555265
      CurrentDate     =   39152
   End
   Begin VB.CommandButton mlast 
      Caption         =   "››"
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
      Height          =   375
      Left            =   9840
      TabIndex        =   19
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton mnext 
      Caption         =   "›"
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
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton mprev 
      Caption         =   "‹"
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
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton mfirst 
      Caption         =   "‹‹"
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
      Height          =   375
      Left            =   8040
      TabIndex        =   20
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   0
      Top             =   960
      Width           =   8895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   5160
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60555265
      CurrentDate     =   39152
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   8280
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1680
      Top             =   8280
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2880
      Top             =   8280
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   4080
      Top             =   8280
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
      Caption         =   "Adodc4"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DUE DATE:"
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   16
      Top             =   5595
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Borrowed:"
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   14
      Top             =   5595
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TITLE:"
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   12
      Top             =   5115
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CODE #:"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   6075
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MEM. NAME:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5595
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MEMBER ID:"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   5115
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   120
      Top             =   840
      Width           =   10455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Movie"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2085
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   0
      Picture         =   "frmReturn.frx":1602F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10665
   End
   Begin VB.Image Image4 
      Height          =   1155
      Left            =   -120
      Picture         =   "frmReturn.frx":3B873
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   10800
   End
   Begin VB.Image Image1 
      Height          =   7095
      Left            =   -240
      Picture         =   "frmReturn.frx":55F0F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mm As String


Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Command1_Click()
Dim k, j

If Adodc1.Recordset.RecordCount <= 0 Then
   MsgBox "No Record Found!"
   Exit Sub
End If

On Error GoTo er1

MonthView1.Value = Text5.Text   'set the value of the date
MonthView2.Value = Text6.Text


k = MonthView1.Value - MonthView2.Value ' subtract the value from each calendar

'if k returns negative then ok, else then elapsed date will occur

If MonthView1.Value < Date Then

    MonthView1.Value = Text6.Text
    MonthView2.Value = Date
    k = MonthView2.Value - MonthView1.Value
End If


If k <= 0 Then
        'put the data in returned database
        Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
        Adodc3.CommandType = adCmdText 'set the Commasssssnd Type as text
        Adodc3.RecordSource = "Select * From returned"
        Adodc3.Refresh
        
        With Adodc3.Recordset
        
                .AddNew
                .Fields(0) = Text1.Text
                .Fields(1) = Text2.Text
                .Fields(2) = Text3.Text
                .Fields(3) = Text4.Text
                .Fields(4) = Text5.Text
                .Fields(5) = Text6.Text
                .Fields(6) = "0"
                .Update
            
        End With
                
        Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
        Adodc4.CommandType = adCmdText 'set the Command Type as text
        Adodc4.RecordSource = "Select * From movies WHERE movieCodeNumber LIKE '" & Text3.Text & "%'" 'SQL Query
        Adodc4.Refresh
        
        With Adodc4.Recordset
            .Fields("movieStock") = .Fields("movieStock") + 1
            .Update
            
            If .Fields("movieStock") >= 1 Then
               .Fields("movieStat") = "IN"
               .Update
            End If
            
        End With
        
        'delete data in borrowed database
        Adodc2.Recordset.Delete adAffectCurrent
        

   MsgBox "Thank You for Returning on Time!", vbOKOnly, "VRENT"
   Exit Sub
Else

frmDueDate.Text1.Text = Text6.Text
frmDueDate.Text3.Text = k
frmDueDate.Text4.Text = rateperday * k
frmDueDate.Show 1

End If

er1:

If Err.Description = "" Then
   MsgBox "Thank You! Come Again!", vbOKOnly, "VRENT"
   Exit Sub
Else
   MsgBox Err.Description, vbOKOnly, "VRENT"
End If
Exit Sub


End Sub

Private Sub DataGrid1_Click()
List1.Visible = False

End Sub

Private Sub Form_Load()
mm = "memberID"
End Sub

Private Sub List1_Click()
txtCustomer.Text = List1.Text
List1.Visible = False
End Sub

Private Sub Option1_Click()
mm = "memberID"
End Sub

Private Sub Option2_Click()
mm = "memberFullName"
End Sub

Private Sub mfirst_Click()
On Error GoTo hell
Adodc2.Recordset.MoveFirst
MonthView1.Value = Text6.Text
Exit Sub
hell:

    'mprev.Enabled = False
    mfirst.Enabled = False
   ' mnext.Enabled = False
    'mlast.Enabled = False
    
End Sub

Private Sub mlast_Click()
On Error GoTo hell
Adodc2.Recordset.MoveLast
MonthView1.Value = Text6.Text
Exit Sub
hell:

   ' mprev.Enabled = False
   ' mfirst.Enabled = False
   ' mnext.Enabled = False
    mlast.Enabled = False
End Sub

Private Sub mnext_Click()
On Error GoTo hell
Adodc2.Recordset.MoveNext
MonthView1.Value = Text6.Text
Exit Sub
hell:
If Adodc1.Recordset.EOF Then
    mprev_Click
    
End If
    'mprev.Enabled = False
    'mfirst.Enabled = False
    ' mnext.Enabled = False
    'mlast.Enabled = False
End Sub

Private Sub mprev_Click()
On Error GoTo hell

Adodc2.Recordset.MovePrevious
MonthView1.Value = Text6.Text
Exit Sub
hell:
    If Adodc1.Recordset.BOF Then
     mnext_Click
    End If
    'mprev.Enabled = False
    'mfirst.Enabled = False
    'mnext.Enabled = False
    'mlast.Enabled = False
End Sub

Private Sub Text4_Change()
On Error Resume Next
MonthView1.Value = Text6.Text
End Sub

Private Sub Timer1_Timer()
On Error GoTo er
If Adodc1.Recordset.RecordCount <= 0 Then
   Command1.Enabled = False
Else
   Command1.Enabled = True
End If

Exit Sub
er:
    Command1.Enabled = False
End Sub

Private Sub txtCustomer_Change()
Dim search As String

List1.Clear
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
'search = txtCustomer.Text 'sets the search value
    
Adodc2.CommandType = adCmdText 'set the Command Type as text
Adodc1.RecordSource = "Select * From members Where " & mm & " LIKE '" & txtCustomer.Text & "%'" 'SQL Query
Adodc2.RecordSource = "Select * From borrowed Where  memberID  LIKE '" & txtCustomer.Text & "%'" 'SQL Query
Adodc1.Refresh 'Refresh the database
Adodc2.Refresh


Text1.DataField = "memberID"
Text2.DataField = "memberName"
Text3.DataField = "movieCodeNumber"
Text4.DataField = "movieTitle"
Text5.DataField = "dateBorrowed"
Text6.DataField = "dateReturned"

Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Label3(0).Visible = False
Label3(1).Visible = False
Label3(2).Visible = False
Label3(3).Visible = False
Label3(4).Visible = False
Label4.Visible = False


If txtCustomer.Text = "" Then
    DataGrid1.Visible = False
    List1.Visible = False
    List1.Clear
    mprev.Enabled = False
    mfirst.Enabled = False
    mnext.Enabled = False
    mlast.Enabled = False
    
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Label3(0).Visible = False
Label3(1).Visible = False
Label3(2).Visible = False
Label3(3).Visible = False
Label3(4).Visible = False
Label4.Visible = False
    
    Exit Sub
End If

If Adodc1.Recordset.EOF Then

   

        'If no User Found on System
        
    mprev.Enabled = False
    mfirst.Enabled = False
    mnext.Enabled = False
    mlast.Enabled = False
        DataGrid1.Visible = False
        List1.Clear
        List1.Visible = False
        List1.Clear
        Exit Sub
    Else
        DataGrid1.Visible = True
        List1.Clear
        List1.Visible = True
        
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Label3(0).Visible = True
Label3(1).Visible = True
Label3(2).Visible = True
Label3(3).Visible = True
Label3(4).Visible = True
Label4.Visible = True
        
            mprev.Enabled = True
            mfirst.Enabled = True
            mnext.Enabled = True
            mlast.Enabled = True
            On Error Resume Next
            MonthView1.Value = Text6.Text
        Adodc1.Recordset.MoveFirst
        List1.Clear
        Do Until Adodc1.Recordset.EOF
        List1.AddItem Adodc1.Recordset.Fields("memberID") & "  --  " & Adodc1.Recordset.Fields("memberFullName")
        Adodc1.Recordset.MoveNext
        Loop
    
        
        If Len(txtCustomer.Text) >= 5 Then
        List1.Visible = False
        End If
        
End If



End Sub
