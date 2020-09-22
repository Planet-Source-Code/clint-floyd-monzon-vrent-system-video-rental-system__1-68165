VERSION 5.00
Begin VB.Form frmMovieSelection 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   4470
   ClientTop       =   3840
   ClientWidth     =   8655
   Icon            =   "frmMovieSelection.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View All Data"
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
         Height          =   480
         Left            =   4680
         TabIndex        =   6
         Top             =   1560
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   8640
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Image imagebutton4 
         Height          =   1410
         Left            =   4200
         Picture         =   "frmMovieSelection.frx":000C
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label lblcancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
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
         Height          =   240
         Left            =   6225
         TabIndex        =   1
         Top             =   1680
         Width           =   1485
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblonebyone 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "One-by-One Registration"
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
         Height          =   480
         Left            =   585
         TabIndex        =   3
         Top             =   1560
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblreg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulk Registration"
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
         Height          =   480
         Left            =   2505
         TabIndex        =   2
         Top             =   1560
         Width           =   1545
         WordWrap        =   -1  'True
      End
      Begin VB.Label l2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Management Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label l1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Management Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2190
         TabIndex        =   5
         Top             =   150
         Width           =   3615
      End
      Begin VB.Image imagebutton1 
         Height          =   1410
         Left            =   600
         Picture         =   "frmMovieSelection.frx":D33E
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Image imagebutton2 
         Height          =   1410
         Left            =   2400
         Picture         =   "frmMovieSelection.frx":1A670
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Image imagebutton3 
         Height          =   1410
         Left            =   6000
         Picture         =   "frmMovieSelection.frx":279A2
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Image i3 
         Height          =   690
         Left            =   0
         Picture         =   "frmMovieSelection.frx":34CD4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8880
      End
      Begin VB.Image Image1 
         Height          =   4935
         Left            =   0
         Picture         =   "frmMovieSelection.frx":42008
         Stretch         =   -1  'True
         Top             =   600
         Width           =   8655
      End
   End
End
Attribute VB_Name = "frmMovieSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton1.Picture = frmImages.btnImageUp.Picture
imagebutton2.Picture = frmImages.btnImageUp.Picture
imagebutton3.Picture = frmImages.btnImageUp.Picture
imagebutton4.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub imagebutton1_Click()

Unload Me
frmMovie.Show 1
End Sub

Private Sub imagebutton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton1.Picture = frmImages.btnImageDown
End Sub

Private Sub imagebutton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton1.Picture = frmImages.btnImageHover.Picture
imagebutton2.Picture = frmImages.btnImageUp.Picture
imagebutton3.Picture = frmImages.btnImageUp.Picture
imagebutton4.Picture = frmImages.btnImageUp.Picture

End Sub

Private Sub imagebutton2_Click()

Unload Me
frmbulk.Show 1
End Sub

Private Sub imagebutton2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton2.Picture = frmImages.btnImageDown.Picture
End Sub

Private Sub imagebutton2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton2.Picture = frmImages.btnImageHover.Picture
imagebutton1.Picture = frmImages.btnImageUp.Picture
imagebutton3.Picture = frmImages.btnImageUp.Picture
imagebutton4.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub imagebutton3_Click()

Unload Me
End Sub

Private Sub imagebutton3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton3.Picture = frmImages.btnImageDown.Picture
End Sub

Private Sub imagebutton3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton3.Picture = frmImages.btnImageHover.Picture
imagebutton2.Picture = frmImages.btnImageUp.Picture
imagebutton1.Picture = frmImages.btnImageUp.Picture
imagebutton4.Picture = frmImages.btnImageUp.Picture
End Sub

Private Sub imagebutton4_Click()
Unload Me
frmData.Show 1
End Sub

Private Sub imagebutton4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton4.Picture = frmImages.btnImageDown.Picture
End Sub

Private Sub imagebutton4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imagebutton4.Picture = frmImages.btnImageHover.Picture
imagebutton2.Picture = frmImages.btnImageUp.Picture
imagebutton1.Picture = frmImages.btnImageUp.Picture
imagebutton3.Picture = frmImages.btnImageUp.Picture
End Sub
