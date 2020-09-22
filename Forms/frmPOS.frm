VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOS 
   BackColor       =   &H0086591E&
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   195
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   Picture         =   "frmPOS.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   3720
      ScaleHeight     =   5295
      ScaleWidth      =   5775
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Default         =   -1  'True
         Height          =   615
         Left            =   2880
         TabIndex        =   42
         Top             =   4440
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear All"
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   4440
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   375
         Left            =   5160
         TabIndex        =   37
         Top             =   0
         Width           =   495
      End
      Begin VB.ListBox l2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Double-Click item you want to remove."
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Void"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   105
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   5655
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H8000000A&
         FillStyle       =   0  'Solid
         Height          =   5415
         Left            =   -120
         Top             =   -240
         Width           =   5775
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   5415
         Left            =   0
         Top             =   120
         Width           =   5775
      End
   End
   Begin VB.ListBox l6 
      Height          =   1815
      Left            =   6240
      TabIndex        =   38
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   12960
      ScaleHeight     =   195
      ScaleWidth      =   1275
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   11520
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1455
      Begin VB.PictureBox Picture2 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   3840
         ScaleHeight     =   195
         ScaleWidth      =   1035
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Press Space..."
      Top             =   8040
      Width           =   5895
   End
   Begin VB.PictureBox picStat 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   9960
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9375
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   10800
      Width           =   9375
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - EXIT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8130
         TabIndex        =   11
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Multiply  Items"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6240
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - RETURNS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4965
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Void/Delete All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F2 - Total"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2025
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F1 - HELP"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
      Begin VB.Image Image5 
         Height          =   690
         Left            =   7800
         Picture         =   "frmPOS.frx":240044
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1560
      End
      Begin VB.Image Image4 
         Height          =   690
         Left            =   6240
         Picture         =   "frmPOS.frx":24D378
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1560
      End
      Begin VB.Image Image3 
         Height          =   690
         Left            =   4680
         Picture         =   "frmPOS.frx":25A6AC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1560
      End
      Begin VB.Image Image2 
         Height          =   690
         Left            =   3120
         Picture         =   "frmPOS.frx":2679E0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   690
         Left            =   1560
         Picture         =   "frmPOS.frx":274D14
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1560
      End
      Begin VB.Image btn1 
         Height          =   690
         Left            =   0
         Picture         =   "frmPOS.frx":282048
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1560
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   13200
      Top             =   3480
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
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComCtl2.MonthView mv 
      Height          =   2370
      Left            =   7560
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   8804638
      Appearance      =   1
      StartOfWeek     =   60489729
      CurrentDate     =   39148
   End
   Begin VB.ListBox l5 
      Height          =   1815
      Left            =   4800
      TabIndex        =   33
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox l4 
      Height          =   1815
      Left            =   3600
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox l3 
      Height          =   1815
      Left            =   2400
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox l1 
      Height          =   1815
      Left            =   960
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox cart 
      Height          =   5055
      Left            =   1200
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2880
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPOS.frx":28F37C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   8520
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F9 - Movie Lookup"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9375
      TabIndex        =   39
      Top             =   11040
      Width           =   1515
   End
   Begin VB.Image Image7 
      Height          =   690
      Left            =   9360
      Picture         =   "frmPOS.frx":28F3FE
      Stretch         =   -1  'True
      Top             =   10800
      Width           =   1560
   End
   Begin VB.Label lblRentCount 
      Caption         =   "0"
      Height          =   375
      Left            =   8160
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblMemberName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   1320
      TabIndex        =   31
      Top             =   9240
      Width           =   150
   End
   Begin VB.Label lblStockCount 
      Caption         =   "0"
      Height          =   495
      Left            =   1440
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Movie Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1440
      TabIndex        =   29
      Top             =   8100
      Width           =   2220
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   14280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00BB7D2B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5775
      Left            =   1320
      Top             =   3000
      Width           =   8415
   End
   Begin VB.Label lblprinter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Printer Found on System!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   390
      Left            =   11280
      TabIndex        =   12
      Top             =   9960
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image Image6 
      Height          =   1215
      Left            =   10680
      Picture         =   "frmPOS.frx":29C732
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Item(s):"
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
      Left            =   9960
      TabIndex        =   19
      Top             =   7800
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cashier Name:"
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
      Left            =   9960
      TabIndex        =   18
      Top             =   5880
      Width           =   4335
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "unknown"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   17
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Number:"
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
      Left            =   9960
      TabIndex        =   16
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   15
      Top             =   7080
      Width           =   4335
   End
   Begin VB.Label lblItemCount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   14
      Top             =   8040
      Width           =   4335
   End
   Begin VB.Image shopLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   9960
      Picture         =   "frmPOS.frx":29E7FA
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1800.95"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C9862C&
      Height          =   1665
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Php"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblTotalShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1800.95"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   1125
      TabIndex        =   3
      Top             =   405
      Width           =   5175
   End
   Begin VB.Image btnImageUp 
      Height          =   735
      Left            =   0
      Picture         =   "frmPOS.frx":2F667E
      Stretch         =   -1  'True
      Top             =   10800
      Width           =   15465
   End
End
Attribute VB_Name = "frmPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
picStat.BackColor = vbRed
End Sub

Private Sub btn1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btn1.Picture = frmImages.btnImageHover.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub cart_Click()
Text1.SetFocus
End Sub

Private Sub Command1_Click()
Picture5.Visible = False
Text1.SetFocus
End Sub

Private Sub Command2_Click()
l1.Clear
l2.Clear
l3.Clear
l4.Clear
l5.Clear
l6.Clear
lblTotal.Caption = "0"
lblTotalShadow.Caption = "0"
lblItemCount.Caption = "0"
cart.Text = ""

End Sub

Private Sub Command3_Click()
Picture5.Visible = False
End Sub

Private Sub Form_Load()
Dim ln, ss As String
Dim fs As New FileSystemObject
'On Error Resume Next

On Error Resume Next
frmPOS.shopLogo.Picture = LoadPicture(systemlogos)
frmPOS.lblTotal.Caption = "0"
frmPOS.lblTotalShadow.Caption = "0"

frmPOS.lblItemCount.Caption = "0"
Dim pth As String
If fs.FolderExists(App.Path & "\transactions\Rent Date -" & Date$) = True Then
        
        pth = App.Path & "\transactions\Rent Date -" & Date$ & "\"     'set the saving path
            
        Else
            fs.CreateFolder App.Path & "\transactions\Rent Date -" & Date$
            
End If

File1.Path = App.Path & "\transactions\Rent Date -" & Date$ 'Set the Path of the file list box for future counting.

Me.Move 0, 0, Screen.Width, Screen.Height  'Resize the form as large as the monitor


lblName.Caption = mainuser          'Sets the Name of the Current user


'COUNT THE PRINTERS INSTALLED ON THE COMPUTER
'TAKE NOTE THAT ALL PRINTERS ARE COUNTED AS WELL AS THE OTHER DRIVER THAT ACT AS PRINTER

If VB.Printers.Count <= 0 Then
lblprinter.Visible = True
Else
lblprinter.Visible = False
End If

'SET THE VISIBILITY OF THE INFORMATION WHEN THERE IS NO PRINTER INSTALLED.

If lblprinter.Visible <> False Then
   Image6.Visible = True
Else
    Image6.Visible = False
End If

Label4.Caption = File1.ListCount

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'TO HIGHLIGHT THE BUTTON WHEN IT RECEIVES THE MOUSE MOVE SIGNAL

btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture


End Sub

Private Sub Image1_Click()

'CHECKS IF THE TOTAL IS NOT EQUAL TO ZERO.
'IF ZERO, IT CANNOT BE TOTALED.

    If lblTotal.Caption = "0" Then
       MsgBox "Nothing to be total. The Cart is EMPTY!", vbCritical, "Error"
    Else

    Label4.Caption = Int(Label4.Caption + 1)
    
    cart.Text = cart.Text & vbNewLine & vbNewLine
           

    cart.Text = cart.Text & "Total        :             P" & lblTotal.Caption
    
    cart.SelStart = Len(cart.Text)
    
    frmCash.Show 1
    
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Picture = frmImages.btnImageUp.Picture
btn1.Picture = frmImages.btnImageUp.Picture
'Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Image2_Click()
frmVoid.Show 1
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
'Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
'Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Image4_Click()
'IN ABLE TO MULTIPLY THE PRODUCT, YOU MUST FIRST SUPPLY THE PRODUCT CODE.

    If Text1.Text = "" Then
        MsgBox "Enter first the Movie Code!", vbCritical, "Error"
    Else
        frmMultiply.Show 1
    End If
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image4.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
'Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Image5_Click()
Unload Me
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image5.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
'Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub l2_Click()
l1.ListIndex = l2.ListIndex
l3.ListIndex = l2.ListIndex
l4.ListIndex = l2.ListIndex
l5.ListIndex = l2.ListIndex
l6.ListIndex = l2.ListIndex







End Sub

Private Sub l2_DblClick()
Dim xx As Integer
Dim kk, ss As String
Dim y As Integer
lblTotal.Caption = Val(lblTotal.Caption) - l6.Text
lblTotalShadow.Caption = lblTotal.Caption
l2.RemoveItem l2.ListIndex
l1.RemoveItem l1.ListIndex
l3.RemoveItem l3.ListIndex
l4.RemoveItem l4.ListIndex
l5.RemoveItem l5.ListIndex
l6.RemoveItem l6.ListIndex
lblItemCount.Caption = Val(lblItemCount.Caption) - 1

cart.Text = ""

For xx = 0 To l2.ListCount Step 1

       l2.ListIndex = xx - 1
       
       kk = l2.Text
       
       If Len(kk) >= 15 Then
          kk = Mid(kk, 1, 15)
       ElseIf Len(kk) < 15 Then
                
                For y = Len(kk) To 15 Step 1
                    ss = ss & " "
                Next y
                    kk = kk '& ss
       Else
       
           kk = kk
       End If
       
       
       cart.Text = cart.Text & kk & "     " & l5.Text & "    " & l6.Text & vbNewLine
       
        
Next xx







End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
'Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Label11_Click()
    If Text1.Text = "" Then
        MsgBox "Enter first the Movie Code!", vbCritical, "Error"
    Else
        frmMultiply.Show 1
    End If
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image4.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
'Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Label12_Click()
Unload Me
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image5.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
'Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'btn1.Picture = frmImages.btnImageHover.Picture
Image1.Picture = frmImages.btnImageUp.Picture
btn1.Picture = frmImages.btnImageHover.Picture
'Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Label8_Click()
'***********************************************************************
'*   CHECK IF THE CART TOTAL IS NOT EQUAL TO ZERO
'***********************************************************************

    If lblTotal.Caption = "0" Then
       MsgBox "Nothing to be total. The Cart is EMPTY!", vbCritical, "Error"
    Else
'***********************************************************************
'*  IF CART IS NOT EQUAL TO ZERO THEN PROCEED!
'***********************************************************************
    Label4.Caption = Int(Label4.Caption + 1)
    cart.Text = cart.Text & vbNewLine & vbNewLine
    cart.Text = cart.Text & "Total        :             P" & lblTotal.Caption
    cart.SelStart = Len(cart.Text)
    frmCash.Show 1
    
    End If
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Picture = frmImages.btnImageHover.Picture
'Image1.Picture = frmImages.btnImageUp.Picture
btn1.Picture = frmImages.btnImageUp.Picture
'Image1.Picture = frmImages.btnImageUp.Picture
Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Label9_Click()
frmVoid.Show 1
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Picture = frmImages.btnImageHover.Picture
btn1.Picture = frmImages.btnImageUp.Picture
Image1.Picture = frmImages.btnImageUp.Picture
'Image2.Picture = frmImages.btnImageUp.Picture
Image3.Picture = frmImages.btnImageUp.Picture
Image4.Picture = frmImages.btnImageUp.Picture
Image5.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub Text1_Change()
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"

    'SET THE COMMAND TYPE TO TEXT
    Adodc1.CommandType = adCmdText 'set the Command Type as text
    Adodc1.RecordSource = "Select * From movies Where movieCodeNumber LIKE '" & Text1.Text & "%'" 'SQL Query
    Adodc1.Refresh 'Refresh the database
    On Error GoTo ex1
    lblStockCount.Caption = Adodc1.Recordset.Fields(8)
Exit Sub
ex1:
    Exit Sub
    
End Sub

Private Sub Text1_Click()
' IF TEXTBOX 1 IS CLICKED THEN HIGHLIGHT THE TEXT ON IT!
SendKeys h
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim x, y As Integer
Dim d, search As String
Dim tot As Double
Dim t, s As Integer
Dim MovieTitle, sp As String
Dim memID As Integer
Dim bcount As Integer



'POSITION THE TEXT BY ALIGNING IT WITH SPACES.
d = "            "
'SET THE TEXT TO SEARCH
search = Text1.Text

If KeyCode = vbKeySpace Then
    Text1.Text = ""
    If Val(lblItemCount.Caption) <= 0 Then
        Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
        Adodc1.CommandType = adCmdText 'set the Command Type as text
    
    
        
    If lblItemCount.Caption <= 0 Then
e:
        Form1.Show 1

       
                
    End If
    If Val(lblItemCount.Caption) >= 5 Then
       MsgBox "Cannot go beyond 5 items.", vbInformation + vbOKOnly, "VRENT"
       Text1.Text = ""
       Text1.SetFocus
       Exit Sub
    End If
    
    'Text1.Locked = False
    
    Else
    Exit Sub
    End If

End If



If KeyCode = 13 Then
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
    Dim kkk As Integer
    
    If Text1.Text = "Press Space..." Then
        Exit Sub
    End If
    
    kkk = Val(frmPOS.lblItemCount.Caption) + Val(frmPOS.lblRentCount.Caption)

    If kkk >= 5 Then
            MsgBox "Cannot rent more than 5 items.", vbInformation + vbOKOnly
            Text1.Text = ""
            Text1.SetFocus
            Exit Sub
    End If
    
    
    'CONNECT TO THE DATABASE WITH THIS CONNECTION STRING, THIS WILL _
    ENSURE THAT THE PROGRAM IS ERROR FREE ON CONNECTING THE DATABASE
    
    

    'SET THE COMMAND TYPE TO TEXT
    Adodc1.CommandType = adCmdText 'set the Command Type as text
    Adodc1.RecordSource = "Select * From movies Where movieCodeNumber LIKE '" & search & "'" 'SQL Query
    Adodc1.Refresh 'Refresh the database
    
    'Determine if a record is found on the Database
    If Adodc1.Recordset.EOF Then
        mBox "Error", "Movie Not Found!", "Please Try Again.", "error"
        'MsgBox "Record Not Found!"        'if not found on the database
        Text1.SetFocus
        SendKeys h
        Exit Sub
    Else
        y = 0
        
        If Adodc1.Recordset.Fields("movieStock") <= 0 Then
            MsgBox "Movie Out of Stock!", vbInformation, "Error"
            Exit Sub
        End If
        'SET THE MOVIE TITLE TO THIS VARIABLE
        MovieTitle = Adodc1.Recordset.Fields("movieTitle")
        l1.AddItem Adodc1.Recordset.Fields("movieCodeNumber")
        l2.AddItem Adodc1.Recordset.Fields("movieTitle")
        l3.AddItem Date
        mv.Value = Date + Adodc1.Recordset.Fields("movieNoDays") - 1
         l4.AddItem mv.Value
        l5.AddItem Adodc1.Recordset.Fields("movieFormat")
        l6.AddItem Adodc1.Recordset.Fields("movieRentPrice")
        
        t = Len(MovieTitle)         'CHECKS THE LENGTH OF THE STRING - MovieTitle
        
        If t < 15 Then              'IF THE LENGTH IS LESS THAN 15 THEN
            'ADD SPACES TO MAKE THE MOVIE LENGTH UPTO 15 CHARACTERS.
            'ONLY 15 CHARACTERS IS ALLOWED
            For s = t To 14 Step 1
            
            sp = sp & " "
            
            Next s
            
            MovieTitle = MovieTitle & sp
        
        End If
    
        If t > 15 Then     'IF MOVIE TITLE IS GREATER THAT 15 THEN TRIM IT TO BECOME 15 CHARACTERS LONG
                    'MID(STRING,START,LENGTH OF STRING)
        MovieTitle = Mid(MovieTitle, 1, 15)
        
        End If
        
        
        cart.Text = cart.Text & vbNewLine & MovieTitle & "     " & Adodc1.Recordset.Fields("movieFormat") & "    " & Adodc1.Recordset.Fields("movieRentPrice")
        cart.SelStart = Len(cart.Text)
        Text2.Text = Val(Adodc1.Recordset.Fields("movieRentPrice"))
        tot = Val(tot + Text2.Text)
        lblTotal.Caption = Val(lblTotal.Caption) + Text2.Text
        lblTotalShadow.Caption = Val(lblTotalShadow.Caption) + Text2.Text
        lblItemCount.Caption = Int(lblItemCount.Caption + 1)
        Text1.Text = ""
        Text1.SetFocus
        
    End If
End If



If KeyCode = vbKeyF12 Then
    Unload Me
End If

If KeyCode = vbKeyF2 Then

    If lblTotal.Caption = "0" Then
       MsgBox "Nothing to be total. The Cart is EMPTY!", vbCritical, "Error"
    Else

    Label4.Caption = Int(Label4.Caption + 1)
    
    cart.Text = cart.Text & vbNewLine & vbNewLine
           
                                          '             a
    cart.Text = cart.Text & "Total        :             P" & lblTotal.Caption
    
    cart.SelStart = Len(cart.Text)
    
    
    'Form1.Show 1
    
    frmCash.Show 1
    
    End If

End If

If KeyCode = vbKeyF5 Then
    If Text1.Text = "" Then
        MsgBox "Enter first the Movie Code!", vbCritical, "Error"
    Else
        frmMultiply.Show 1
    End If
End If

If KeyCode = vbKeyF3 Then
    'Picture5.Visible = True
    frmVoid.Show 1
End If

If KeyCode = vbKeyF9 Then
    frmSelection.Show 1
End If

Exit Sub
Exit Sub
End Sub





Private Sub Text1_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub
