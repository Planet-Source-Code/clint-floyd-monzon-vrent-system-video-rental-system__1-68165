Attribute VB_Name = "Global_Module"
'                GLOBAL MODULE                            *
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Const h = "{HOME}+{END}"

Global mainstat As Integer '0 if not loaded, 1 if loaded
Global mainuser
Global shopname
Global borrowerID As String
Global borrowerName As String
Global systempass, systemname, systemlogos As String
Global rateperday As Double


Declare Sub ReleaseCapture Lib "user32" ()

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Public Sub formdrag(theform As Form)

    ReleaseCapture
    Call SendMessage(theform.hwnd, &HA1, 2, 0&)
End Sub



Public Function ImageHover(valImage1, valImage2, valImage3, valImage4, valImage5, valImage6 As Integer)
If valImage1 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageHover.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage2 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageHover.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
  Exit Function
End If
If valImage3 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageHover.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage4 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageHover.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage5 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageHover.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage6 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageHover.Picture
   Exit Function
End If

End Function

Public Function ImageDown(valImage1, valImage2, valImage3, valImage4, valImage5, valImage6 As Integer)
If valImage1 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageDown.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage2 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageDown.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
  Exit Function
End If
If valImage3 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageDown.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage4 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageDown.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage5 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageDown.Picture
   frmMain.Image13.Picture = frmImages.btnImageUp.Picture
   Exit Function
End If
If valImage6 = 1 Then
   frmMain.Image8.Picture = frmImages.btnImageUp.Picture
   frmMain.Image9.Picture = frmImages.btnImageUp.Picture
   frmMain.Image10.Picture = frmImages.btnImageUp.Picture
   frmMain.Image11.Picture = frmImages.btnImageUp.Picture
   frmMain.Image12.Picture = frmImages.btnImageUp.Picture
   frmMain.Image13.Picture = frmImages.btnImageDown.Picture
   Exit Function
End If
End Function


Public Property Let hl(ByVal vNewValue As Variant)

End Property

Public Sub checknum(KeyAscii As Integer)
'Accepts only numeric input
Select Case KeyAscii
  Case vbKey0 To vbKey9
  Case vbKeyBack, vbKeyClear, vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
End Select
End Sub
