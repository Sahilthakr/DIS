Attribute VB_Name = "PinInput"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function SetTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)

Private Declare Function KillTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Const EM_SETPASSWORDCHAR = &HCC
Public Const NV_INPUTBOX As Long = &H5000&
Public WindowTitle As String
Public Strkey As String

Private pbBase64Byt(0 To 63)    As Byte             ' base 64 encoder byte array
   
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="


Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim EditHwnd As Long
    EditHwnd = FindWindowEx(FindWindow("#32770", WindowTitle), 0, "Edit", "")
    
    Call SendMessage(EditHwnd, EM_SETPASSWORDCHAR, Asc("*"), 0)
    KillTimer hwnd, idEvent
End Sub

Public Function PasswordInputbox(Prompt As String, Optional Title As String, Optional Default As String, Optional XPos As Long, Optional YPos As Long, Optional HelpFile As Long, Optional Context As Long) As String
        Dim ret As String
    If Title = "" Then
        WindowTitle = App.Title
    Else
        WindowTitle = Title
    End If

    SetTimer 0, 0, 1, AddressOf TimerProc
    PasswordInputbox = InputBox(Prompt, WindowTitle, Default, XPos, YPos, "", Context)
End Function

Public Function Decrypt(ByVal str2Decode As String) As String

Dim lPtr            As Long

    Dim iValue          As Integer

    Dim iLen            As Integer

    Dim iCtr            As Integer

    Dim Bits(1 To 4)    As Byte

    Dim strDecode       As String


    ' for each 4 character group....

    For lPtr = 1 To Len(str2Decode) Step 4

        iLen = 4

        For iCtr = 0 To 3

            ' retrive the base 64 value, 4 at a time

            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)

            Select Case iValue

                ' A~Za~z0~9+/

                Case 1 To 64: Bits(iCtr + 1) = iValue - 1

                ' =

                Case 65

                    iLen = iCtr

                    Exit For

                ' not found

                Case 0: Exit Function

            End Select

        Next


        ' convert the 4, 6 bit values into 3, 8 bit values

        Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10

        Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4

        Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)


        ' add the three new characters to the output string

        For iCtr = 1 To iLen - 1

            strDecode = strDecode & Chr$(Bits(iCtr))

        Next


    Next


    Decrypt = strDecode


End Function

Function Encrypt(InputStr As String) As String

    Dim State As Integer

    Dim i As Integer

    Dim Base64

    Dim current As Integer

    Dim old As Integer

    Dim first_new As String

    Dim second_new As String

    Dim Output As String

   

    Base64 = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
                   "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                   "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", _
                   "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", _
                   "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")

   

    For i = 1 To Len(InputStr)

         State = State + 1

         Select Case State

         Case 1

             current = Asc(Mid(InputStr, i, 1))

             first_new = Base64((Int(current / 4) And &H3F))

             Output = Output & first_new

         Case 2

             current = Asc(Mid(InputStr, i, 1))

             first_new = Base64((((old * 16) And &H30) Or ((Int(current / 16) And &HF))))

             Output = Output & first_new

         Case 3:

             current = Asc(Mid(InputStr, i, 1))

             first_new = Base64((((old * 4) And &H3C) Or ((Int(current / 64) And &H3))))

             second_new = Base64((current And &H3F))

             Output = Output & first_new & second_new

             State = 0

         End Select

         old = current

     Next

     Select Case State

     Case 1

         first_new = Base64(((old * 16) And &H30))

         Output = Output & first_new & "=="

     Case 2

         first_new = Base64(((old * 4) And &H3C))

         Output = Output & first_new & "="

     End Select

     Encrypt = Output

End Function


