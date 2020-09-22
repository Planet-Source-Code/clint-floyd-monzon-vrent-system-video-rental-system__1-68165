VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   6000
   ClientLeft      =   4590
   ClientTop       =   2595
   ClientWidth     =   9315
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   4815
      Left            =   240
      ScaleHeight     =   4755
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
         Height          =   4095
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   6615
         _cx             =   11668
         _cy             =   7223
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Transparent"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   0   'False
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
         AllowNetworking =   "all"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   2280
         X2              =   8400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the type you want to edit."
         Height          =   735
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   6255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
      Begin VB.Image Image2 
         Height          =   1905
         Left            =   120
         Picture         =   "frmSettings.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1920
      End
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "frmSettings.frx":30044
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   9495
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub flash_FSCommand(ByVal command As String, ByVal args As String)
Select Case command
        Case "sys"
        
        frmSetup.Show 1
        Case "users"
        frmNewUser.Show 1

End Select

End Sub

Private Sub Form_Load()
flash.LoadMovie 0, App.Path & "\data\systemsettings.vrs"
End Sub

