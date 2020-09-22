VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VRent:POS"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   ControlBox      =   0   'False
   Icon            =   "frmCash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1440
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Default         =   -1  'True
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   2040
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
      Height          =   375
      Left            =   1680
      Top             =   2040
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
   Begin RichTextLib.RichTextBox t2 
      Height          =   1095
      Left            =   2160
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmCash.frx":000C
   End
   Begin RichTextLib.RichTextBox t1 
      Height          =   1095
      Left            =   600
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmCash.frx":0097
   End
   Begin RichTextLib.RichTextBox tr 
      Height          =   735
      Left            =   600
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmCash.frx":0122
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblChange 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6120
      Picture         =   "frmCash.frx":01AD
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Cash Tendered"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image20 
      Height          =   810
      Left            =   0
      Picture         =   "frmCash.frx":D4DF
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to Expand"
      Top             =   0
      Width           =   8280
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Picture         =   "frmCash.frx":1A813
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim f As New FileSystemObject

Dim a, b, c As Double
Dim sv As String




If Val(Text1.Text) < Val(frmPOS.lblTotal.Caption) Then
        MsgBox "The value must be greater than or equal to the TOTAL Value."
        Text1.Text = ""
        Text1.SetFocus
        Exit Sub
End If
        a = Val(frmPOS.lblTotal.Caption)
        b = Val(Text1.Text)
        c = b - a

        c = Math.Round(c, 2)








'Edit the Receipt Interface
shopname = systemname              'sets the Shopname

shopname = shopname
ln = "--------------------------------------"
ss = " Item#                     Price      "


tr.TextRTF = shopname & vbNewLine & vbNewLine & ln & vbNewLine & ss & vbNewLine & ln






frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & "Cash Tendered:             P" & Text1.Text
frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & "Change       :             P" & c
frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & "______________________________________" & vbNewLine
frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & "Customer #: " & frmPOS.Label4.Caption
frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & "Cashier: " & Mid(frmPOS.lblName.Caption, 1, 20)
frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & "Time/Date: " & Time & " - " & Date

frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & vbNewLine & " This serve as your official receipt.  "
frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & "        Thank You! Come Again!"

tr.TextRTF = tr.Text & vbNewLine & frmPOS.cart.Text
frmPOS.cart.SelStart = Len(frmPOS.cart.Text)

t1.Text = tr.Text
t2.Text = EncryptString(t1.Text, "VRENTSYSTEM")

'save transactions
repeate_action:

If f.FolderExists(App.Path & "\transactions") = True Then
'If folder exist then
    Dim fs As New FileSystemObject
    
        'create a folder according to the date
repsave:
Dim pth As String
        If fs.FolderExists(App.Path & "\transactions\Rent Date -" & Date$) = True Then
        
        pth = App.Path & "\transactions\Rent Date -" & Date$ & "\"     'set the saving path
        
        
        
        sv = Mid(Time$, 1, 2) & "-" & Mid(Time$, 4, 2) & "-" & Mid(Time$, 7, 2) & ".vrs"
             
             
             
             t2.SaveFile pth & sv, 1
            'frmPOS.cart.SaveFile "C:\a.txt", 1
            
            GoTo cont
            
        Else
            fs.CreateFolder App.Path & "\transactions\Rent Date -" & Date$
            GoTo repsave
        End If
        
    
Exit Sub
Else
'if the folder does not exist then create that folder!!!
'after creating the folder, go to previous action.

f.CreateFolder App.Path & "\transactions"
GoTo repeate_action
End If

cont:

'Print Receipt
On Error Resume Next
If frmPOS.lblprinter.Visible = True Then
Else
On Error Resume Next
frmPOS.PrintForm
End If



lblChange.Caption = c     'c is the variable that holds the value of change

Label2.Caption = "Close"

Text1.Visible = False


'Dim ln, ss As String

'shopname = shopname
'ln = "--------------------------------------"
'ss = " Item#                     Price      "


frmPOS.cart.TextRTF = ""
frmPOS.lblTotal.Caption = "0"
frmPOS.lblTotalShadow.Caption = "0"

frmPOS.lblItemCount.Caption = "0"

frmPOS.cart.SelStart = Len(frmPOS.cart.Text)

Command1.Enabled = False
Command2.Enabled = True
Command2.Default = True
tr.Text = ""
Label1.Caption = "Change"

'_____________________________________________________________________________
'*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*'
'*
'*              PUT ALL DATA ON THE BORROWERS DATABASE!!
'*
'*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*8*'
'_____________________________________________________________________________

Dim ilist, il As Integer
Dim cdn As String
ilist = frmPOS.l1.ListCount    'Count the total items on the cart



For il = 1 To ilist Step 1

'On Error GoTo er1
'Open a Connection
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"

Adodc1.CommandType = adCmdText 'set the Command Type as text
Adodc2.CommandType = adCmdText 'set the Command Type as text

Adodc1.RecordSource = "Select * From borrowed"


Adodc1.Refresh 'Refresh the database


'highlight the listbox by selecting their list index - 1; list index start from 0
frmPOS.l1.ListIndex = il - 1
frmPOS.l2.ListIndex = il - 1
frmPOS.l3.ListIndex = il - 1
frmPOS.l4.ListIndex = il - 1
frmPOS.l5.ListIndex = il - 1


Dim tmpstock As String
Dim k
Adodc2.CommandType = adCmdText 'set the Command Type as text
Adodc2.RecordSource = "SELECT * FROM movies WHERE movieCodeNumber LIKE '" & frmPOS.l1.Text & "%'" 'SQL Query
Adodc2.Refresh
tmpstock = Adodc2.Recordset.Fields("movieStock")

With Adodc2.Recordset
    
    .Fields("movieStock") = tmpstock - 1
    .Update
End With

If Adodc2.Recordset.Fields("movieStock") <= 0 Then
    
    With Adodc2.Recordset
    
    .Fields("movieStat") = "OUT"
    .Update
End With

'
    
End If

    With Adodc1.Recordset
    .AddNew
    .Fields("memberID") = borrowerID
    .Fields("memberName") = borrowerName
    .Fields("movieCodeNumber") = frmPOS.l1.Text
    .Fields("movieTitle") = frmPOS.l2.Text
    .Fields("dateBorrowed") = frmPOS.l3.Text
     
    .Fields("dateReturned") = frmPOS.l4.Text
    .Fields("penalty") = ""
    .Update
   End With
Adodc1.Recordset.Close 'Close the Connection
Adodc2.Recordset.Close
Next il

frmPOS.Text1.Locked = True
frmPOS.Text1.Text = "Press Space..."
frmPOS.lblMemberName.Caption = ""



Exit Sub
er1:
MsgBox Err.Description & Err.Number
Exit Sub
End Sub

Private Sub Command2_Click()
lblChange.Caption = ""
frmPOS.l1.Clear
frmPOS.l2.Clear
frmPOS.l4.Clear
frmPOS.l3.Clear
frmPOS.l5.Clear
frmPOS.l6.Clear
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "Enter Cash Tendered:"
Label2.Caption = "Accept"
Text1.Visible = True



End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Picture = frmImages.btnImageUp
End Sub

Private Sub Image2_Click()
If Label2.Caption <> "Accept" Then
    Command2_Click
Else
    Command1_Click
End If

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Picture = frmImages.btnImageDown
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Picture = frmImages.btnImageHover
End Sub

Private Sub Label2_Click()
If Label2.Caption <> "Accept" Then
    Command2_Click
Else
    Command1_Click
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Picture = frmImages.btnImageHover
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

