Attribute VB_Name = "moduleCipher"
Public Function EncryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i As Integer
 
 ' Initilize i and make sure the EncryptKey is long enough
 i = 0
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 ' Loop through the string to encrypt each character.
 Do
  i = i + 1
  OldChar = Asc(Mid(InString, i, 1))
  CryptChar = Asc(Mid(TempKey, i, 1))
  ' If it's an even character, add the ASCII value of the
  ' appropriate character in the Key, otherwise, subract it.
  ' Also, make sure the value is between 0 and 127.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
   Case Else   'Odd Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
  End Select
  ' If the value is less than 35, add 40 to it (to make sure
  ' it's in the display range) and put it in an escape
  ' sequence (using ! [ASCII Value 33] as the escape char)
  If NewChar < 35 Then
   OutString = OutString + "!" + Chr(NewChar + 40)
  Else
   OutString = OutString + Chr(NewChar)
  End If
 Loop Until i = Len(InString)
 
 EncryptString = OutString

End Function

Public Function DecryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i, c As Integer
 
 ' Initialize c and i (loop variables)
 c = 0       ' c is used for InString
 i = 0       ' i is used for EncryptKey
 ' Make sure the EncryptKey is long enough
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 Do
  ' In the decrypt function, two integers are need keeping
  ' track of location (becuase the escape sequence it two
  ' chars long, but only has one placeholder in the key)
  i = i + 1
  c = c + 1
  OldChar = Asc(Mid(InString, c, 1))
  ' If this is an escape sequence, get the next character and
  ' subtract 40 from it's value.
  If OldChar = 33 Then
   c = c + 1
   OldChar = Asc(Mid(InString, c, 1))
   OldChar = OldChar - 40
  End If
  CryptChar = Asc(Mid(TempKey, i, 1))
  ' If it's an even character, subract the appropriate key
  ' value... also, if it's out of range, bring it back in.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
   Case Else   'Odd Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
  End Select
  OutString = OutString + Chr(NewChar)
 Loop Until c = Len(InString)
 
 DecryptString = OutString

End Function


