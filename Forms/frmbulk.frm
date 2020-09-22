VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmbulk 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   4275
   ClientTop       =   2310
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7275
      ScaleWidth      =   10290
      TabIndex        =   0
      Top             =   0
      Width           =   10350
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   9120
         TabIndex        =   16
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Apply"
         Height          =   375
         Left            =   8160
         TabIndex        =   17
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   7200
         TabIndex        =   18
         Top             =   6840
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Movie Properties"
         Height          =   3735
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   9855
         Begin VB.TextBox txtRentPrice 
            Height          =   375
            Left            =   4800
            TabIndex        =   41
            Top             =   2760
            Width           =   2415
         End
         Begin VB.TextBox txtPurchasePrice 
            Height          =   375
            Left            =   4800
            TabIndex        =   39
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox txtStock 
            Height          =   375
            Left            =   4800
            TabIndex        =   38
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox txtStatus 
            Height          =   375
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "IN"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.ComboBox txtNoDays 
            Height          =   315
            ItemData        =   "frmbulk.frx":0000
            Left            =   4800
            List            =   "frmbulk.frx":0016
            TabIndex        =   33
            Text            =   "1"
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txtPurchaseDate 
            Height          =   375
            Left            =   1080
            TabIndex        =   29
            Top             =   2640
            Width           =   2415
         End
         Begin VB.ComboBox txtYear 
            Height          =   315
            Left            =   1080
            TabIndex        =   28
            Top             =   1920
            Width           =   2415
         End
         Begin VB.ComboBox txtFormat 
            Height          =   315
            ItemData        =   "frmbulk.frx":002C
            Left            =   1080
            List            =   "frmbulk.frx":0039
            TabIndex        =   26
            Text            =   "VCD"
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox txtCategory 
            Height          =   375
            Left            =   1080
            TabIndex        =   24
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtTitle 
            Height          =   405
            Left            =   4800
            TabIndex        =   22
            Top             =   360
            Width           =   4935
         End
         Begin VB.TextBox txtcode 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   20
            Top             =   360
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            Height          =   2295
            Left            =   7320
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total List Count:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   390
            Index           =   2
            Left            =   7500
            TabIndex        =   44
            Top             =   1680
            Width           =   2085
         End
         Begin VB.Label lblCount 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   540
            Left            =   8520
            TabIndex        =   42
            Top             =   2160
            Width           =   225
         End
         Begin VB.Label lblCountShadow 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   540
            Left            =   8550
            TabIndex        =   43
            Top             =   2190
            Width           =   225
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Price:"
            Height          =   195
            Index           =   8
            Left            =   3840
            TabIndex        =   40
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Price:"
            Height          =   195
            Index           =   7
            Left            =   3600
            TabIndex        =   37
            Top             =   2400
            Width           =   1125
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Stock"
            Height          =   195
            Index           =   6
            Left            =   3810
            TabIndex        =   36
            Top             =   1920
            Width           =   825
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Index           =   5
            Left            =   4140
            TabIndex        =   34
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Days:"
            Height          =   195
            Index           =   4
            Left            =   3750
            TabIndex        =   32
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            Height          =   195
            Index           =   3
            Left            =   525
            TabIndex        =   31
            Top             =   2760
            Width           =   390
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   30
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year:"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   27
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Format:"
            Height          =   195
            Left            =   360
            TabIndex        =   25
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movie Title:"
            Height          =   195
            Left            =   3840
            TabIndex        =   21
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movie Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Movie Que"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   9855
         Begin VB.ListBox listrentprice 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   9120
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.ListBox listpprice 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   8520
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
         Begin VB.ListBox liststock 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   8040
            TabIndex        =   12
            Top             =   360
            Width           =   495
         End
         Begin VB.ListBox liststat 
            Appearance      =   0  'Flat
            Height          =   1590
            ItemData        =   "frmbulk.frx":004C
            Left            =   7560
            List            =   "frmbulk.frx":004E
            TabIndex        =   11
            Top             =   360
            Width           =   495
         End
         Begin VB.ListBox listnday 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   7080
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
         Begin VB.ListBox listpdate 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   6000
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.ListBox listyear 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   5280
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
         Begin VB.ListBox listformat 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   4560
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.ListBox listcategory 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   3480
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.ListBox listtitle 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   2175
         End
         Begin VB.ListBox listcode 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   480
         Top             =   6240
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulk Movie Registration"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulk Movie Registration"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   1
         Left            =   270
         TabIndex        =   1
         Top             =   165
         Width           =   4365
      End
      Begin VB.Image Image20 
         Height          =   795
         Left            =   0
         Picture         =   "frmbulk.frx":0050
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10440
      End
   End
End
Attribute VB_Name = "frmbulk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim xx As Integer

'On Error GoTo er1

If lblCount.Caption = 0 Then
    MsgBox "No Movie on Waiting List", vbInformation, "Error"
    Exit Sub
End If

For xx = 1 To lblCount.Caption Step 1
'set the index of each listbox to see the text inside it
listcode.ListIndex = xx - 1
listtitle.ListIndex = xx - 1
listcategory.ListIndex = xx - 1
listformat.ListIndex = xx - 1
listyear.ListIndex = xx - 1
listpdate.ListIndex = xx - 1
listnday.ListIndex = xx - 1
liststat.ListIndex = xx - 1
liststock.ListIndex = xx - 1
listpprice.ListIndex = xx - 1
listrentprice.ListIndex = xx - 1

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
Adodc1.CommandType = adCmdTable
Adodc1.RecordSource = "movies"

Adodc1.CommandType = adCmdText 'set the Command Type as text
Adodc1.RecordSource = "Select * From movies"
Adodc1.Refresh 'Refresh the database

'search = listcode.Text 'sets the search value


'Determine if a record is found on the Database
'If Adodc1.Recordset.EOF Then
   'If Not Found then Proceed
    
    With Adodc1.Recordset
    .AddNew
    .Fields(0) = listcode.Text
    .Fields(1) = listtitle.Text
    .Fields(2) = listcategory.Text
    .Fields(3) = listformat.Text
    .Fields(4) = listyear.Text
    .Fields(5) = listpdate.Text
    .Fields(6) = listnday.Text
    .Fields(7) = liststat.Text
    .Fields(8) = liststock.Text
    .Fields(9) = liststock.Text
    .Fields(10) = listpprice.Text
    .Fields(11) = listrentprice.Text
    .Update
    End With


Adodc1.Recordset.Close





Next xx

    MsgBox xx - 1 & " Movie Entry successfully added to database."
    Unload Me


Exit Sub
er1:
MsgBox Err.Description

End Sub

Private Sub Command3_Click()
Dim x As VbMsgBoxResult

x = MsgBox("Is the Information you've enter is correct?", vbYesNo, "VRent:Bulk")
If x = vbNo Then
    Exit Sub
Else
   GoTo co
End If

co:
If Len(txtcode.Text) > 0 And Len(txtTitle.Text) > 0 And _
    Len(txtCategory.Text) > 0 And Len(txtFormat.Text) > 0 _
    And Len(txtYear.Text) > 0 And Len(txtPurchaseDate.Text) > 0 _
    And Len(txtNoDays.Text) > 0 And Len(txtStatus.Text) > 0 _
    And Len(txtStock.Text) > 0 And Len(txtPurchasePrice.Text) > 0 Then
    
    

listcode.AddItem txtcode.Text
listtitle.AddItem txtTitle.Text
listcategory.AddItem txtCategory.Text
listformat.AddItem txtFormat.Text
listyear.AddItem txtYear.Text
listpdate.AddItem txtPurchaseDate.Text
listnday.AddItem txtNoDays.Text
liststat.AddItem txtStatus.Text
liststock.AddItem txtStock.Text
listpprice.AddItem txtPurchasePrice.Text
listrentprice.AddItem txtRentPrice.Text


listcode.ListIndex = listcode.ListCount - 1
listtitle.ListIndex = listtitle.ListCount - 1
listcategory.ListIndex = listcategory.ListCount - 1
listformat.ListIndex = listformat.ListCount - 1
listyear.ListIndex = listyear.ListCount - 1
listpdate.ListIndex = listpdate.ListCount - 1
listnday.ListIndex = listnday.ListCount - 1
liststat.ListIndex = liststat.ListCount - 1
liststock.ListIndex = liststock.ListCount - 1
listpprice.ListIndex = listpprice.ListCount - 1
listrentprice.ListIndex = listrentprice.ListCount - 1

lblCount.Caption = listcode.ListCount
lblCountShadow.Caption = listcode.ListCount


txtcode.Text = ""
txtTitle.Text = ""
txtCategory.Text = ""
txtFormat.Text = "VCD"
txtYear.Text = ""
txtPurchaseDate.Text = ""
txtNoDays.Text = "1"
txtStock.Text = ""
txtPurchasePrice.Text = ""
txtRentPrice.Text = ""
txtcode.SetFocus

Else

MsgBox "Cannot contain an EMPTY value." & vbCr & "Please fill up the fields accordingly.", vbInformation, "Error"
    Exit Sub

End If

End Sub

Private Sub Form_Load()
Dim x As Integer
txtYear.Text = Year(Date)

For x = Year(Date) To 1960 Step -1
txtYear.AddItem x
Next x

End Sub

Private Sub txtCategory_Change()
txtCategory.Text = UCase(txtCategory.Text)
txtCategory.SelStart = Len(txtCategory.Text)
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub txtFormat_Change()
txtFormat.Text = "VCD"
End Sub

Private Sub txtNoDays_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub txtPurchasePrice_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub txtRentPrice_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub

Private Sub txtTitle_Change()
txtTitle.Text = UCase(txtTitle.Text)
txtTitle.SelStart = Len(txtTitle.Text)
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub
