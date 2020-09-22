Attribute VB_Name = "Module_MessageBox"


Public Function mBox(Main_Title As String, message1 As String, message2 As String, mtype As String)
Attribute mBox.VB_Description = "Replaces original Message box with the brand new one"


Select Case mtype
        Case "error"
             frmMessage.Image3.Picture = frmImages.imgError.Picture
        Case "question"
                frmMessage.Image3.Picture = frmImages.imgQuestion.Picture
        Case "info"
                frmMessage.Image3.Picture = frmImages.imgInfo.Picture
End Select



frmMessage.mainTitle.Caption = Main_Title
frmMessage.message1.Caption = message1
frmMessage.message2.Caption = message2
 
frmMessage.Show 1

End Function
