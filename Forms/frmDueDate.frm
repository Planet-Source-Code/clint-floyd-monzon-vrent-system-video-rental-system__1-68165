VERSION 5.00
Begin VB.Form frmDueDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "You've Elapsed the Due date"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   ControlBox      =   0   'False
   Icon            =   "frmDueDate.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Please Enter Amount Received:"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Penalty Charge:"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Elapsed Day:"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "And Today  is:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "You're Due Date is:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmDueDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As Integer
'On Error Resume Next
If Val(Text5.Text) >= Val(Text4.Text) Then
'put the data in database
        frmReturn.Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
        frmReturn.Adodc3.CommandType = adCmdText 'set the Command Type as text
        frmReturn.Adodc3.RecordSource = "Select * From returned"
        frmReturn.Adodc3.Refresh
        
        With frmReturn.Adodc3.Recordset
        
                .AddNew
                .Fields(0) = frmReturn.Text1.Text
                .Fields(1) = frmReturn.Text2.Text
                .Fields(2) = frmReturn.Text3.Text
                .Fields(3) = frmReturn.Text4.Text
                .Fields(4) = frmReturn.Text5.Text
                .Fields(5) = frmReturn.Text6.Text
                .Fields(6) = Text4.Text
                .Update
            
        End With
        frmReturn.Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
        frmReturn.Adodc4.CommandType = adCmdText 'set the Command Type as text
        frmReturn.Adodc4.RecordSource = "Select * From movies WHERE movieCodeNumber LIKE '" & frmReturn.Text3.Text & "%'" 'SQL Query
        frmReturn.Adodc4.Refresh
        
        With frmReturn.Adodc4.Recordset
            .Fields("movieStock") = .Fields("movieStock") + 1
            .Update
               
               If .Fields("movieStock") >= 1 Then
                    .Fields("movieStat") = "IN"
                    .Update
               End If
        End With
                
     'delete data in borrowed database
     frmReturn.Adodc2.Recordset.Delete
     
     If frmReturn.Adodc2.Recordset.EOF = True Then
             frmReturn.Adodc2.Recordset.MovePrevious
     Else
             frmReturn.Adodc2.Recordset.MoveNext
     End If
     
     
     x = Text5.Text - Text4.Text
         
     MsgBox "Change: " & x, vbInformation + vbOKOnly, "VRENT"
     
     
    'frmReturn.Adodc2.Recordset.Delete adAffectCurrent
   'frmReturn.Adodc2.Recordset.Fields(0).Delete
   'frmReturn.Adodc2.Recordset.Fields(1).Delete
   'frmReturn.Adodc2.Recordset.Fields(2).Delete
   'frmReturn.Adodc2.Recordset.Fields(3).Delete
   'frmReturn.Adodc2.Recordset.Fields(4).Delete
   'frmReturn.Adodc2.Recordset.Fields(5).Delete
   'frmReturn.Adodc2.Recordset.Fields(6).Delete
   'frmReturn.Adodc2.Recordset.Update
   Unload Me
Else
   MsgBox "The Value must greater than or equal to the Amount.", vbInformation, "ERROR"
   Text5.Text = ""
   Text5.SetFocus
End If
End Sub

Private Sub Form_Load()
Text2.Text = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub
