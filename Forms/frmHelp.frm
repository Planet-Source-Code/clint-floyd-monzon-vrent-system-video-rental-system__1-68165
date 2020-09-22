VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHelp 
   Caption         =   "VRENT ::: Web Help"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5940
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   0
      Width           =   8115
      Begin MSComctlLib.ProgressBar p 
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton Command2 
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
   End
   Begin SHDocVwCtl.WebBrowser w 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2175
      ExtentX         =   3836
      ExtentY         =   2778
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
w.GoBack
End Sub

Private Sub Command2_Click()
On Error Resume Next
w.GoForward
End Sub

Private Sub Form_Load()
w.Navigate "http://bsit.us.to/vrent"
End Sub

Private Sub Form_Resize()
w.Move 0, 0 + Picture1.Height, Me.Width - 100, Me.Height - 1000
End Sub

Private Sub w_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
p.Max = ProgressMax
p.Value = Progress
End Sub

Private Sub w_TitleChange(ByVal Text As String)
Me.Caption = Text

If Mid(Text, 1, 4) = "HTTP" Then
   w.Navigate App.Path & "\error\index.htm"
End If
End Sub
