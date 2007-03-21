'Copyright Pacific Northwest National Laboratory / Battelle
'File:  PwdDecode.vb
'File Contents:  Class for simple decoding of obfuscated passwords using "Pwd.exe" (written by
'                Dave Clark
'Author(s):  Nathan Trimble, Dave Clark

Public Class PwdDecoder

    'Class object not createable or static
    Private Sub New()
    End Sub

    Public Shared Function DecodePassword(ByVal EnPwd As String) As String
        'Decrypts password received from ini file
        ' Password was created by alternately subtracting or adding 1 to the ASCII value of each character

        Dim CharCode As Byte
        Dim TempStr As String
        Dim Indx As Integer

        TempStr = ""

        Indx = 1
        Do While Indx <= Len(EnPwd)
            CharCode = CByte(Asc(Mid(EnPwd, Indx, 1)))
            If Indx Mod 2 = 0 Then
                CharCode = CharCode - CByte(1)
            Else
                CharCode = CharCode + CByte(1)
            End If
            TempStr = TempStr & Chr(CharCode)
            Indx = Indx + 1
        Loop

        Return TempStr
    End Function
End Class
