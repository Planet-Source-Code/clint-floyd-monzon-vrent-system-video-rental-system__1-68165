VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPenaltyChecker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penalty Checker"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmPenaltyChecker.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   120
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPenaltyChecker.frx":1601A
      Height          =   3135
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60358657
      CurrentDate     =   39182
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60358657
      CurrentDate     =   39149
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while the system computes the penalties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmPenaltyChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x
x = MonthView2.Value - MonthView1.Value
MsgBox x
End Sub

Private Sub Timer1_Timer()
Dim x, tmpX, z, a As Integer
Dim tmpMonth1, tmpMonth2, tmpMonth3
Dim tot
Dim totality

On Error GoTo ex1

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc1.CommandType = adCmdText 'set the Command Type as text
Adodc1.RecordSource = "Select * From borrowed"
Adodc1.Refresh 'Refresh the database
a = 0
x = Adodc1.Recordset.RecordCount
pb.Max = x
Adodc1.Recordset.MoveFirst
Timer1.Enabled = False
For z = 1 To x Step 1
    pb.Value = z
    
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
    Adodc1.CommandType = adCmdText 'set the Command Type as text
    Adodc1.RecordSource = "Select * From borrowed"
    Adodc1.Refresh 'Refresh the database
    Adodc1.Recordset.MoveFirst
    For a = z To a + 1 Step 1
    
            Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF Then
                Unload Me
                Exit Sub
                
            End If
    Next a
        
    MonthView1.Value = Adodc1.Recordset.Fields("dateBorrowed")
    MonthView2.Value = Adodc1.Recordset.Fields("dateReturned")
    tot = MonthView2.Value - MonthView1.Value
            
                If tot <= 0 Then
                    
                    MonthView1.Value = Adodc1.Recordset.Fields("dateReturned")
                    MonthView2.Value = Date
                    tot = MonthView2.Value - MonthView1.Value
                    totality = tot * rateperday
                End If
            
            'MsgBox totality
            
            With Adodc1.Recordset
                
                .Fields("Penalty") = totality
                .Update
            End With
    
    Adodc1.Recordset.Close
Next z
    
    Unload Me
    
Exit Sub
ex1:
    Unload frmPenaltyChecker
    Exit Sub
End Sub
