VERSION 5.00
Begin VB.Form frmImages 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox color1 
      BackColor       =   &H00693616&
      Height          =   495
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00EEC577&
      Height          =   495
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Picture         =   "frmImages.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image imgInfo 
      Height          =   1095
      Left            =   2040
      Picture         =   "frmImages.frx":D332
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image imgError 
      Height          =   1095
      Left            =   1080
      Picture         =   "frmImages.frx":1352D
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image imgQuestion 
      Height          =   1095
      Left            =   0
      Picture         =   "frmImages.frx":16EAA
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image btnImageDown 
      Height          =   810
      Left            =   0
      Picture         =   "frmImages.frx":1C1F2
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image btnImageUp 
      Height          =   810
      Left            =   960
      Picture         =   "frmImages.frx":29526
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image btnImageHover 
      Height          =   810
      Left            =   1920
      Picture         =   "frmImages.frx":3685A
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu addnewuser 
      Caption         =   "Add New User"
      Begin VB.Menu addUser 
         Caption         =   "Add New User"
      End
      Begin VB.Menu editUser 
         Caption         =   "Edit User"
      End
      Begin VB.Menu separtt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main Menu"
      Begin VB.Menu rentacd 
         Caption         =   "&Rent-a-CD"
      End
      Begin VB.Menu returnCD 
         Caption         =   "R&eturn CD"
      End
      Begin VB.Menu manageCustomer 
         Caption         =   "&Manage Customer"
      End
      Begin VB.Menu manageMovies 
         Caption         =   "M&anage Movies"
      End
      Begin VB.Menu reports 
         Caption         =   "Re&ports"
      End
      Begin VB.Menu mnuSlash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnumem 
      Caption         =   "frmMem"
      Begin VB.Menu mnuViewAllMembers 
         Caption         =   "View All Members"
      End
      Begin VB.Menu mnuVei 
         Caption         =   "View Member By"
         Begin VB.Menu menumemdate 
            Caption         =   "Membership Date"
         End
         Begin VB.Menu mnumemloc 
            Caption         =   "Location"
         End
      End
   End
   Begin VB.Menu mnumovielist 
      Caption         =   "movielist"
      Begin VB.Menu viewalldata 
         Caption         =   "View All Data"
      End
      Begin VB.Menu mnuviewby 
         Caption         =   "View by"
         Begin VB.Menu viewbystatus 
            Caption         =   "Status"
         End
         Begin VB.Menu viewbycategory 
            Caption         =   "Category"
         End
         Begin VB.Menu viewbystocks 
            Caption         =   "Stocks"
         End
      End
   End
   Begin VB.Menu mnuBorrowed 
      Caption         =   "Borrowed"
      Begin VB.Menu viewalldataborrowed 
         Caption         =   "View All Data"
      End
      Begin VB.Menu viewalllduedates 
         Caption         =   "View All Due Dates"
      End
   End
End
Attribute VB_Name = "frmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addUser_Click()
Dim i As String

i = InputBox("Please Enter System Password", "VRENT-ADMIN")

If i <> systempass Then
    MsgBox "Invalid Password!", vbInformation, "VRENT"
Else
frmNewUser.Show 1
End If
End Sub

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
Private Sub btnImageUp_Click()

End Sub

Private Sub manageCustomer_Click()
frmMemberUpdate.Show 1
End Sub

Private Sub manageMovies_Click()
Unload Me
frmMovieSelection.Show 1
End Sub

Private Sub menumemdate_Click()
Dim sqlstr As String
Dim dt As String
Set movRS = New ADODB.Recordset
dt = InputBox("Enter Date ex. 05/14/2006", "VRENT")
sqlstr = "SELECT * FROM members where memberSince = '" & dt & "'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly

Set rptMembers.DataSource = movRS

rptMembers.Show vbModal
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnumemloc_Click()
Dim dt As String
Set movRS = New ADODB.Recordset
dt = InputBox("Enter Location/Address", "VRENT")
sqlstr = "SELECT * FROM members where memberAddress LIKE '%" & dt & "%'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly

Set rptMembers.DataSource = movRS

rptMembers.Show vbModal
End Sub

Private Sub mnuSettings_Click()
frmSystempassword.Show 1
myforms = "frmSettings"
End Sub

Private Sub mnuViewAllMembers_Click()
Dim sqlstr As String
Set movRS = New ADODB.Recordset
sqlstr = "SELECT * FROM members"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly

Set rptMembers.DataSource = movRS

rptMembers.Show vbModal
End Sub

Private Sub rentacd_Click()
Unload Me
frmPOS.Show 1
End Sub

Private Sub reports_Click()
frmReports.Show 1
End Sub

Private Sub returnCD_Click()
frmReturn.Show 1
End Sub

Private Sub viewalldata_Click()
Dim sqlstr As String

Set movRS = New ADODB.Recordset
sqlstr = "SELECT * FROM movies"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly
rptMov.Caption = "Movies"
Set rptMov.DataSource = movRS
rptMov.Show vbModal
End Sub

Private Sub viewalldataborrowed_Click()
Dim sqlstr As String
Set movRS = New ADODB.Recordset
sqlstr = "SELECT * FROM borrowed"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly

Set rptborrower.DataSource = movRS

rptborrower.Show vbModal
End Sub

Private Sub viewalllduedates_Click()
Dim sqlstr As String
Dim t
Set movRS = New ADODB.Recordset
t = InputBox("Enter Date ex. 05/14/2006", "VRENT", Date)
sqlstr = "SELECT * FROM borrowed Where dateReturned < '" & t & "'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly

Set rptborrower.DataSource = movRS

rptborrower.Show vbModal
End Sub

Private Sub viewbycategory_Click()
Dim sqlstr As String
Dim t As String
Set movRS = New ADODB.Recordset

t = InputBox("Enter Category ex. Family, Thriller", "VRENT")

sqlstr = "SELECT * FROM movies where movieCategory LIKE '%" & t & "%'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly
rptMov.Caption = "Movies"
Set rptMov.DataSource = movRS
rptMov.Show vbModal
End Sub

Private Sub viewbystatus_Click()
Dim sqlstr As String
Dim t As String
Set movRS = New ADODB.Recordset

t = InputBox("ENTER STATUS EITHER IN or OUT", "VRENT")

sqlstr = "SELECT * FROM movies where moviestat ='" & t & "'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly
rptMov.Caption = "Movies"
Set rptMov.DataSource = movRS
rptMov.Show vbModal
End Sub

Private Sub viewbystocks_Click()
Dim sqlstr As String
Dim t As String
Set movRS = New ADODB.Recordset

t = InputBox("Enter Stock ex. 10,20", "VRENT")

sqlstr = "SELECT * FROM movies where movieStock LIKE '%" & t & "%'"

movRS.Open sqlstr, vidCon, adOpenKeyset, adLockReadOnly
rptMov.Caption = "Movies"
Set rptMov.DataSource = movRS
rptMov.Show vbModal
End Sub
