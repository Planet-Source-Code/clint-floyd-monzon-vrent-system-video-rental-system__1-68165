VERSION 5.00
Begin VB.Form frmMultiply 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmMultiply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
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
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiply Item By:"
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
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   0
      Picture         =   "frmMultiply.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   4335
      Left            =   -1320
      Picture         =   "frmMultiply.frx":D33E
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   8055
   End
End
Attribute VB_Name = "frmMultiply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim search, d As String
Dim t, s, x, y As Integer
Dim MovieTitle, sp As String
Dim k, kkk As Integer

If frmPOS.Adodc1.Recordset.RecordCount <= 0 Then
    MsgBox "No Movie Found on Database!"
    Exit Sub
End If

kkk = Val(frmPOS.lblRentCount) + Val(frmPOS.lblItemCount)

If kkk >= 5 Then
    MsgBox "Cannnot rent more than 5 items."
    Unload Me
    Exit Sub
End If


x = 1

d = "            "

k = Int(Text1.Text) + frmPOS.lblItemCount.Caption

If Val(frmPOS.lblStockCount.Caption) <= 0 Then
   MsgBox "Movie Out of Stock.", vbInformation + vbOKOnly, "VRENT"
   Unload Me
End If

If Text1.Text > 5 Then
   MsgBox "Cannot Rent more than 5.", vbInformation + vbOKOnly, "VRENT"
   Exit Sub
End If

If k > 5 Then
    
    MsgBox "Cannot go beyond 5 items.", vbInformation + vbOKOnly, "VRENT"
    Exit Sub
End If
   
If Text1.Text > Val(frmPOS.lblStockCount.Caption) Then
   MsgBox "There's only " & frmPOS.lblStockCount.Caption & " stock available."
   Exit Sub
End If

If Text1.Text = "" Then

MsgBox "Please Enter a Number.", vbInformation, "Error"

Else

    If Val(Text1.Text) = 0 Then
        
        MsgBox "Invalid Data Input!", vbInformation, "Error"
    Else

search = frmPOS.Text1.Text


    frmPOS.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"


    frmPOS.Adodc1.CommandType = adCmdText 'set the Command Type as text
    frmPOS.Adodc1.RecordSource = "Select * From movies Where movieCodeNumber LIKE '" & search & "'" 'SQL Query
    frmPOS.Adodc1.Refresh 'Refresh the database
    
    'Determine if a record is found on the Database
    If frmPOS.Adodc1.Recordset.EOF Then
        mBox "Error", "Movie Not Found!", "Please Try Again.", "error"
        'MsgBox "Record Not Found!"        'if not found on the database
        'frmPOS.Text1.SetFocus
        SendKeys h
        Unload Me
        Exit Sub
    Else
        
        
        MovieTitle = frmPOS.Adodc1.Recordset.Fields("movieTitle")
        frmPOS.lblItemCount.Caption = Int(frmPOS.lblItemCount.Caption) + Int(Text1.Text)
        'Loop until it reach the maximum entry
        For x = 1 To Int(Text1.Text) Step 1
                frmPOS.l1.AddItem frmPOS.Adodc1.Recordset.Fields("movieCodeNumber")
                frmPOS.l2.AddItem frmPOS.Adodc1.Recordset.Fields("movieTitle")
                frmPOS.l3.AddItem Date
                frmPOS.mv.Value = Date + frmPOS.Adodc1.Recordset.Fields("movieNoDays") - 1
                frmPOS.l4.AddItem frmPOS.mv.Value
                frmPOS.l5.AddItem frmPOS.Adodc1.Recordset.Fields("movieFormat")
                frmPOS.l6.AddItem frmPOS.Adodc1.Recordset.Fields("movieRentPrice")
                
        Next
        
        
        
        t = Len(MovieTitle)
        
        If t < 15 Then
        
            For s = t To 14 Step 1
            
            sp = sp & " "
            
            Next s
            
            MovieTitle = MovieTitle & sp
        
        End If
    
        If t > 15 Then
        
        MovieTitle = Mid(MovieTitle, 1, 15)
        
        End If
        
        Dim p As Double
        
        p = frmPOS.Adodc1.Recordset.Fields("movieRentPrice") * Text1.Text
        
        frmPOS.cart.Text = frmPOS.cart.Text & vbNewLine & MovieTitle & "     " & frmPOS.Adodc1.Recordset.Fields("movieFormat") & "    " & frmPOS.Adodc1.Recordset.Fields("movieRentPrice") & vbNewLine & "       @ " & Text1.Text & "                 " & p
        frmPOS.cart.SelStart = Len(frmPOS.cart.Text)
        frmPOS.Text2.Text = Val(frmPOS.Adodc1.Recordset.Fields("movieRentPrice"))
        tot = Val(tot + (frmPOS.Text2.Text * Text1.Text))
        frmPOS.lblTotal.Caption = Val(frmPOS.lblTotal.Caption) + (frmPOS.Text2.Text * Text1.Text)
        frmPOS.lblTotalShadow.Caption = Val(frmPOS.lblTotalShadow.Caption) + (frmPOS.Text2.Text * Text1.Text)
        
        frmPOS.Text1.Text = ""
        'frmPOS.Text1.SetFocus
        
        Unload Me
        
    End If
    End If
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
checknum KeyAscii
End Sub
