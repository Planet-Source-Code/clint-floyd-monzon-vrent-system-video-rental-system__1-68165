VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "V-Rent System"
   ClientHeight    =   11460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11460
   ScaleWidth      =   15360
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   13920
      Top             =   8760
   End
   Begin VB.PictureBox navipane 
      BorderStyle     =   0  'None
      Height          =   9020
      Left            =   0
      ScaleHeight     =   9015
      ScaleWidth      =   3135
      TabIndex        =   23
      Top             =   1280
      Width           =   3135
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "<<<"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Hide Navigational Pane"
         Top             =   0
         Width           =   3135
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   3120
         X2              =   3120
         Y1              =   0
         Y2              =   48740
      End
      Begin VB.Image Image14 
         Height          =   1095
         Left            =   120
         Picture         =   "MainForm.frx":1601A
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rent-a-CD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1200
         TabIndex        =   31
         Top             =   750
         Width           =   1305
      End
      Begin VB.Image Image15 
         Height          =   1095
         Left            =   120
         Picture         =   "MainForm.frx":1C45F
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return CD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1200
         TabIndex        =   30
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Image Image16 
         Height          =   1305
         Left            =   120
         Picture         =   "MainForm.frx":228A4
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manage"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1320
         TabIndex        =   29
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1560
         TabIndex        =   28
         Top             =   3960
         Width           =   1245
      End
      Begin VB.Image Image17 
         Height          =   1200
         Left            =   120
         Picture         =   "MainForm.frx":271D2
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manage"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1320
         TabIndex        =   27
         Top             =   4920
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movies"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1800
         TabIndex        =   26
         Top             =   5280
         Width           =   945
      End
      Begin VB.Image Image18 
         Height          =   1185
         Left            =   240
         Picture         =   "MainForm.frx":2E390
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1440
         TabIndex        =   25
         Top             =   6600
         Width           =   1005
      End
      Begin VB.Image Image19 
         Height          =   1200
         Left            =   120
         Picture         =   "MainForm.frx":31D1D
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   390
         Left            =   1560
         TabIndex        =   24
         Top             =   8040
         Width           =   1035
      End
      Begin VB.Image Image8 
         Height          =   1410
         Left            =   240
         Picture         =   "MainForm.frx":379D4
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2760
      End
      Begin VB.Image Image9 
         Height          =   1410
         Left            =   240
         Picture         =   "MainForm.frx":44D08
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   2760
      End
      Begin VB.Image Image10 
         Height          =   1410
         Left            =   240
         Picture         =   "MainForm.frx":5203C
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   2760
      End
      Begin VB.Image Image11 
         Height          =   1410
         Left            =   240
         Picture         =   "MainForm.frx":5F370
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   2760
      End
      Begin VB.Image Image12 
         Height          =   1410
         Left            =   240
         Picture         =   "MainForm.frx":6C6A4
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   2760
      End
      Begin VB.Image Image13 
         Height          =   1410
         Left            =   240
         Picture         =   "MainForm.frx":799D8
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   2760
      End
      Begin VB.Image Image6 
         Height          =   9180
         Left            =   0
         Picture         =   "MainForm.frx":86D0C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3180
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   9240
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2000
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   9840
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7080
      TabIndex        =   10
      Text            =   "Search..."
      Top             =   1440
      Width           =   2655
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3150
      ScaleHeight     =   735
      ScaleWidth      =   15480
      TabIndex        =   35
      Top             =   1280
      Width           =   15480
      Begin VB.Label myTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rent-a-CD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   240
         TabIndex        =   36
         Top             =   0
         Width           =   2025
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   0
         Picture         =   "MainForm.frx":8E280
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15465
      End
   End
   Begin VB.Timer picture2Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6480
      Top             =   480
   End
   Begin VB.Timer lblTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   480
   End
   Begin VB.Timer naviTime 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show   Nav i ga t  i  ona l      Pane"
      Height          =   9015
      Left            =   0
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Show Navigational Pane"
      Top             =   1270
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3840
      Top             =   2760
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   15360
      TabIndex        =   0
      Top             =   10245
      Width           =   15360
      Begin VB.Label welcome 
         Appearance      =   0  'Flat
         BackColor       =   &H00693616&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   5895
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   1680
         X2              =   1680
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Label cmdHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00693616&
         Caption         =   " Help "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   960
         X2              =   960
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Label cmdAbout 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00693616&
         Caption         =   "about"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   300
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label thedate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created: February 13, 2007"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   8355
         TabIndex        =   5
         Top             =   240
         Width           =   2220
      End
      Begin VB.Image Image4 
         Height          =   1275
         Left            =   5520
         Picture         =   "MainForm.frx":B3AC4
         Top             =   0
         Width           =   6360
      End
      Begin VB.Image image5 
         Height          =   1275
         Left            =   0
         Picture         =   "MainForm.frx":CE160
         Top             =   0
         Width           =   15360
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3120
      Top             =   10440
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      Height          =   3135
      Left            =   3120
      ScaleHeight     =   3075
      ScaleWidth      =   12195
      TabIndex        =   14
      Top             =   9645
      Width           =   12250
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   11400
         TabIndex        =   37
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1920
         Top             =   120
      End
      Begin VB.Label lblPrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6000
         TabIndex        =   34
         Top             =   960
         Width           =   150
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   4920
         TabIndex        =   22
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lblCount3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2280
         TabIndex        =   21
         Top             =   2160
         Width           =   150
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrow Count:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label lblCount2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2280
         TabIndex        =   19
         Top             =   1560
         Width           =   150
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "   User Counts:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movie Counts:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2280
         TabIndex        =   16
         Top             =   960
         Width           =   150
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Summary"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image Image20 
         Height          =   570
         Left            =   -120
         Picture         =   "MainForm.frx":10DDA4
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to Expand"
         Top             =   0
         Width           =   16920
      End
   End
   Begin VB.Label myVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v. 1.0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
   Begin VB.Label myVersionShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v. 1.0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   3620
      TabIndex        =   2
      Top             =   390
      Width           =   480
   End
   Begin VB.Label cmdOff 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00693616&
      Caption         =   " Log-Off "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   795
   End
   Begin VB.Label cmdMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00693616&
      Caption         =   " Minimize "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   6435
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label cmdClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00693616&
      Caption         =   " Close "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   7515
      MouseIcon       =   "MainForm.frx":11B0D8
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   1275
      Left            =   7440
      Picture         =   "MainForm.frx":11B3E2
      Top             =   0
      Width           =   15360
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   0
      Picture         =   "MainForm.frx":15B026
      Top             =   0
      Width           =   7500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   10680
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Image Image7 
      Height          =   8460
      Left            =   0
      MouseIcon       =   "MainForm.frx":17A276
      MousePointer    =   99  'Custom
      Picture         =   "MainForm.frx":17A3C8
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   10980
   End
End
Attribute VB_Name = "frmMain"
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
'                M A I N    F O R M                       *
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Dim lf As Integer

Private Sub cmdAbout_Click()
frmAbout.Show 1   'Shows the About Window with Modal, meaning the other form is disable.
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdAbout.BackColor = frmImages.color2.BackColor 'Change the Back Color of the button to light blue
cmdAbout.ForeColor = vbBlack 'change the font color to black
End Sub

Private Sub cmdClose_Click()
End     'Terminate the Program
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdClose.BackColor = vbRed 'Change the Close button's back color to red
cmdClose.ForeColor = vbWhite 'Changes the font color to white
End Sub



Private Sub cmdHelp_Click()
frmHelp.Show
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdHelp.BackColor = frmImages.color2.BackColor 'Changes the bg color to light blue
cmdHelp.ForeColor = frmImages.color1.BackColor 'changes the font color to dark violet

End Sub

Private Sub cmdMin_Click()
Me.WindowState = 1 'To minimize this window, set the Window State to 1 or Minimize
End Sub

Private Sub cmdMin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdMin.BackColor = frmImages.color2.BackColor 'Change bg color to light blue
cmdMin.ForeColor = frmImages.color1.BackColor 'changes the font color to violet

End Sub

Private Sub cmdOff_Click()
'To log off
mainstat = 1 'set 1 if the main window is open
Me.Hide      'hides this window
frmLogin.Show 'Show the Log-in form.
End Sub

Private Sub cmdOff_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdOff.BackColor = frmImages.color2.BackColor 'changes the bg color to light blue
cmdOff.ForeColor = vbBlack 'changes the font color to black
End Sub

Private Sub cmdSearch_Click()
Dim x, y As Integer
Dim tmpsearch, sp As String
Dim search As String


If txtSearch.Text = "Search..." Then
    
txtSearch.Text = ""

End If


cmdSearch.Enabled = False



Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
search = txtSearch.Text 'sets the search value
    
    Adodc1.CommandType = adCmdText 'set the Command Type as text
    Adodc1.RecordSource = "Select * From movies Where movieTitle LIKE '" & search & "%'" 'SQL Query
    Adodc1.Refresh 'Refresh the database
    
    'Determine if a record is found on the Database
    If Adodc1.Recordset.EOF Then
        MsgBox "Record Not Found!"        'if not found on the database
        txtSearch.SetFocus
        SendKeys h
        cmdSearch.Enabled = True
        Exit Sub
    Else
    Adodc1.Recordset.MoveFirst           'Move the database pointer to the first
    
    List1.Clear
    
    Do Until Adodc1.Recordset.EOF
    y = 0
    sp = ""
    tmpsearch = ""
    tmpsearch = Adodc1.Recordset.Fields(1)
    
    If Len(tmpsearch) < 26 Then
       For y = Len(tmpsearch) To 25 Step 1
            sp = sp & "."
       Next y
           tmpsearch = tmpsearch & sp
    ElseIf Len(tmpsearch) > 25 Then
       
       tmpsearch = Mid(tmpsearch, 1, 22) & "...."
    Else
       tmpsearch = tmpsearch
    
    End If
    
    Dim k As String  ' for the new adodc recordset value
    
    If Len(Adodc1.Recordset.Fields(7)) < 3 Then
    k = Adodc1.Recordset.Fields(7) & " "
    Else
    k = Adodc1.Recordset.Fields(7)
    End If
    
    
    List1.AddItem (tmpsearch & "   " & k & "  " & Adodc1.Recordset.Fields(8) & " of " & Adodc1.Recordset.Fields(9))
    Adodc1.Recordset.MoveNext
    
    Loop
    
    Dim totList As Integer
    
    totList = List1.ListCount
    totList = 315 * totList

    List1.Visible = True
    For x = 315 To totList Step 10
    List1.Height = x
    Next x
    Command1.Top = List1.Top + List1.Height + 100
    Command1.Left = List1.Left + List1.Width - Command1.Width
    Command1.Visible = True
    Command1.Default = True
End If
End Sub

Private Sub Command1_Click()
Dim x, totList As Integer

cmdSearch.Enabled = True

totList = List1.ListCount
totList = 315 * totList

Command1.Visible = False
Command1.Default = False
cmdSearch.Default = True

For x = totList To 315 Step -10
List1.Height = x
Next x
List1.Visible = False
'Command1.Value = False


End Sub

Private Sub Command2_Click()
naviTime.Enabled = True
End Sub

Private Sub Command3_Click()
naviTime.Enabled = True
End Sub

Private Sub Form_Activate()
Dim mb As VbMsgBoxResult
    Adodc2.RecordSource = "Select * From movies" 'SQL Query
    Adodc2.Refresh 'Refresh the database
    
    lblCount.Caption = Adodc2.Recordset.RecordCount
If lblCount.Caption <= 0 Then
    mb = MsgBox("It seems that your shop doesn't have a movie saved on database." & vbCr & "Do you want to import list of movies?", vbQuestion + vbYesNo, "VRENT")
    
    If mb = vbYes Then
       frmMovieSelection.Show 1
       Exit Sub
    Else
       Exit Sub
    End If
    
End If
End Sub

Private Sub Form_Load()
Dim mo, dy 'declare the variable for Month and Day
welcome.Caption = "Welcome " & mainuser 'Welcomes the User
    Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
    Adodc2.CommandType = adCmdText 'set the Command Type as text
    Adodc2.RecordSource = "Select * From movies" 'SQL Query
    Adodc2.Refresh 'Refresh the database
    
    lblCount.Caption = Adodc2.Recordset.RecordCount
    
    'Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
    'Adodc2.CommandType = adCmdText 'set the Command Type as text
    Adodc2.RecordSource = "Select * From systemuser" 'SQL Query
    Adodc2.Refresh 'Refresh the database
    
    lblCount2.Caption = Adodc2.Recordset.RecordCount
    
    Adodc2.RecordSource = "Select * From borrowed"
    Adodc2.Refresh
    
    lblCount3.Caption = Adodc2.Recordset.RecordCount
    
Select Case Weekday(Date, vbSunday)     'select the weekday or the day today eg. Monday, Tuesday, etc.
       Case "1"              '1 for Sunday, 2 for Monday, 3 for Tuesday, etc.
       dy = "Sunday"
       Case "2"
       dy = "Monday"
       Case "3"
       dy = "Tuesday"
       Case "4"
       dy = "Wednesday"
       Case "5"
       dy = "Thursday"
       Case "6"
       dy = "Friday"
       Case "7"
       dy = "Sunday"
End Select                'End the Selection

Select Case Month(Date)   'Select for the Month for Today (eg. January, February, etc)
       Case "1"           '1 for January, 2 for February, 3 for March, etc
       mo = "January"
       Case "2"
       mo = "February"
       Case "3"
       mo = "March"
       Case "4"
       mo = "April"
       Case "5"
       mo = "May"
       Case "6"
       mo = "June"
       Case "7"
       mo = "July"
       Case "8"
       mo = "August"
       Case "9"
       mo = "September"
       Case "10"
       mo = "October"
       Case "11"
       mo = "November"
       Case "12"
       mo = "December"
End Select              'End Selection
thedate.Caption = "Today is " & dy & ", " & mo & " " & Day(Date) & ", " & Year(Date) 'Print out the Gathered data regarding the date.

Me.Move 0, 0, Screen.Width, Screen.Height 'moves the form to the 0,0 axis of the screen
Line1.X2 = Screen.Width                   'Formats the line to be as the width of the monitor screen
Image4.Move Me.Width - Image4.Width       'Moves the Position of each objects
Image3.Width = Me.Width

'Count Printers Installed on the Computer
If VB.Printers.Count <= 0 Then
lblPrint.Caption = "No Printer Installed"   'Show this Message if No Printer Found
Else
lblPrint.Caption = VB.Printers.Count & " Printer(s) Available" ' Show this if one or more printer is FOund
End If

'Align the Objects
'Image3.Move 0, 1275, Me.Width
lf = myTitle.Left
Image6.Height = Me.Height
Image7.Height = Me.Height
Image7.Width = Me.Width
Image7.Left = 0
cmdClose.Left = Me.Width - cmdClose.Width - 100
cmdMin.Left = cmdClose.Left - cmdMin.Width - 200
cmdOff.Left = cmdMin.Left - cmdOff.Width - 200
myVersion.Caption = "v. " & App.Major & "." & App.Minor
myVersionShadow.Caption = "v. " & App.Major & "." & App.Minor
thedate.Move Me.Width - thedate.Width - 100



cmdSearch.Left = Me.Width - cmdSearch.Width - 100
txtSearch.Left = cmdSearch.Left - txtSearch.Width - 100
List1.Left = txtSearch.Left - List1.Width + txtSearch.Width

'set the status of the system
'1 = loaded
'2 = hidden/logoff
mainstat = 1

frmPenaltyChecker.Show 1

End Sub

Private Sub Image10_Click()
frmMemberSelect.Show 1
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageDown 0, 0, 1, 0, 0, 0   'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Manage Customers"
ImageHover 0, 0, 1, 0, 0, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image11_Click()
frmMovieSelection.Show 1
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageDown 0, 0, 0, 1, 0, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageHover 0, 0, 0, 1, 0, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image12_Click()
frmReports.Show 1
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageDown 0, 0, 0, 0, 1, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Reports"
ImageHover 0, 0, 0, 0, 1, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image13_Click()
myforms = "frmSettings"
frmSystempassword.Show 1
End Sub

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageDown 0, 0, 0, 0, 0, 1 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Settings"
ImageHover 0, 0, 0, 0, 0, 1 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image14_Click()
frmPOS.Show 1
End Sub

Private Sub Image14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Rent-a-CD"
End Sub

Private Sub Image15_Click()
frmReturn.Show 1
End Sub

Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Return CD"
End Sub

Private Sub Image16_Click()
frmMemberSelect.Show 1
End Sub

Private Sub Image16_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Manage Customers"
End Sub

Private Sub Image17_Click()
frmMovieSelection.Show 1
End Sub

Private Sub Image17_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Manage Movies"
End Sub

Private Sub Image18_Click()
frmReports.Show 1
End Sub

Private Sub Image18_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Reports"
End Sub

Private Sub Image19_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Settings"
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'returns the default colors of all the command button on top

cmdClose.BackColor = frmImages.color1.BackColor '
cmdClose.ForeColor = vbWhite
cmdMin.BackColor = frmImages.color1.BackColor
cmdMin.ForeColor = vbWhite
cmdOff.BackColor = frmImages.color1.BackColor
cmdOff.ForeColor = vbWhite
End Sub

Private Sub Image20_DblClick()
Timer1.Enabled = True
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'return the default colors of all the command button on top
cmdAbout.BackColor = frmImages.color1.BackColor
cmdAbout.ForeColor = vbWhite
cmdHelp.BackColor = frmImages.color1.BackColor
cmdHelp.ForeColor = vbWhite
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageDown 0, 0, 0, 0, 0, 0
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then ' 1 - Left Click, 2 - Right Click

PopupMenu frmImages.mnuMain, 0, x, y + 1200

End If
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageDown 1, 0, 0, 0, 0, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
frmPOS.Show 1

End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Rent-a-CD"
ImageHover 1, 0, 0, 0, 0, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image9_Click()
frmReturn.Show 1
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ImageDown 0, 1, 0, 0, 0, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Return CD"
ImageHover 0, 1, 0, 0, 0, 0 'set 1 if it is active, 0 if inactive. This is for the mouse over and mouse remove function for button
End Sub


Private Sub Label1_Click()
frmPOS.Show 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Rent-a-CD"
End Sub

Private Sub Label2_Click()
frmReturn.Show 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Return CD"
End Sub

Private Sub Label3_Click()
frmMemberSelect.Show 1
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Manage Customers"
End Sub

Private Sub Label4_Click()
frmMemberSelect.Show 1
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Manage Customers"
End Sub

Private Sub Label5_Click()
frmMovieSelection.Show 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Manage Movies"
End Sub

Private Sub Label6_Click()
frmMovieSelection.Show 1
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Manage Movies"
End Sub

Private Sub Label7_Click()
frmReports.Show 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Reports"
End Sub

Private Sub Label8_Click()
myforms = "frmSettings"
frmSystempassword.Show 1
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
myTitle.Caption = "Settings"
End Sub

Private Sub lblTimer_Timer()
Dim x As Integer






If Picture3.Left >= 3100 Then

    picture2Timer.Enabled = True

   For x = 3150 To 260 Step -1
   
       Picture3.Left = x
       
        
   Next x
   
   
   
   lblTimer.Enabled = False
   
Else

    picture2Timer.Enabled = True

   For x = 260 To 3150 Step 1
   
        Picture3.Left = x
       
        
   Next x
     
       lblTimer.Enabled = False
   
End If


End Sub

Private Sub naviTime_Timer()
Dim x As Integer


If navipane.Left >= 0 Then
   
       lblTimer.Enabled = True
       
   For x = 0 To -3135 Step -1
   
       navipane.Left = x
       
          
   Next x
   
   
   

   
   
   naviTime.Enabled = False

Else

    lblTimer.Enabled = True

   For x = -3135 To 0 Step 1
   
        navipane.Left = x
        
   Next x
   
    naviTime.Enabled = False
   
End If
   


End Sub

Private Sub picture2Timer_Timer()

Dim x As Integer

If Picture2.Left = 3120 Then

    Picture2.Width = Me.Width - 255
    
    For x = 3120 To 255 Step -1
    
        Picture2.Left = x
    
    Next x
    
        picture2Timer.Enabled = False
        
Else

    For x = 255 To 3120 Step 1
    
        Picture2.Left = x
    
    Next x


    Picture2.Width = 12250

    picture2Timer.Enabled = False
    
    
End If



End Sub

Private Sub Text1_Change()
If Text1.Text >= 240000 Then
   frmPenaltyChecker.Show 1
   Text1.Text = 0
End If
End Sub

Private Sub Timer1_Timer()
Dim x As Integer


If Picture2.Top <> 7200 Then

        For x = 9650 To 7200 Step -1
        Picture2.Top = x
        Next x

        Timer1.Enabled = False
        Image20.ToolTipText = "Double-Click to Collapse"
Else
    
        For x = 7200 To 9650 Step 1
        Picture2.Top = x
        Next x

        Timer1.Enabled = False
        Image20.ToolTipText = "Double-Click to Expand"
        
End If

End Sub

Private Sub Timer2_Timer()
Text1.Text = Text1.Text + 1

End Sub

Private Sub txtSearch_Click()
If txtSearch.Text = "Search..." Then txtSearch.Text = ""
SendKeys h
End Sub




