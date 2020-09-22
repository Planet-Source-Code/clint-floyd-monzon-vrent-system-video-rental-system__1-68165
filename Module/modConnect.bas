Attribute VB_Name = "modConnect"
Public vidCon As ADODB.Connection
Public movRS As ADODB.Recordset
Public membership As ADODB.Connection
Public memUpdate As ADODB.Recordset

Sub dbConnect()
    Set vidCon = New ADODB.Connection
    vidCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
End Sub



Sub cnnMember()
    Set membership = New ADODB.Connection
    vidCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\systemdata.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=password"
End Sub

