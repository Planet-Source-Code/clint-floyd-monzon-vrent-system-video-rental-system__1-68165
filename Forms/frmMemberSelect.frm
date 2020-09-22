VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9.ocx"
Begin VB.Form frmMemberSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VRENT"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmMemberSelect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _cx             =   13150
      _cy             =   5741
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
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Picture         =   "frmMemberSelect.frx":1601A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmMemberSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub flash_FSCommand(ByVal command As String, ByVal args As String)
On Error Resume Next
Select Case command

        Case "newuser"
            frmNewMember.Show 1
        
        Case "edituser"
            frmMemberUpdate.Show 1
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
flash.LoadMovie 0, App.Path & "\data\memberselect.vrs"
End Sub
