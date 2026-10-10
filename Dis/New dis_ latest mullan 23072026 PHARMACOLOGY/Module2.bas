Attribute VB_Name = "Module2"

Option Explicit
Public mainuser1 As String
Public mainpass1 As String
Public RdEnv As rdoEnvironment
Public RdCn As rdoConnection
Public varipAddr As String

Public Sub Main1()
mainuser = "TMH"
mainpass = "TMH"
Set RdEnv = rdoEnvironments(0)
RdEnv.CursorDriver = rdUseOdbc
Set RdCn = RdEnv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.1.4.2;UID=" & mainuser & ";PWD=" & mainpass & ";database=DISLIB")
'Set con = rnv.OpenConnection(DSNAME:="stores", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.1.4.2;UID=" & mainuser & ";PWD=" & mainpass & ";database=TOOLSADMN")End Sub
End Sub


Public Function ValidateDate(ByVal InDt As String) As Boolean
    DoEvents
    ValidateDate = False
    If Not IsDate(InDt) Then
        MsgBox "Not a valid Date", vbInformation
        Exit Function
    End If
    Dim Dt, mn, Yr As String
    Dt = Left(InDt, 2)
    mn = Mid(InDt, 4, 2)
    Yr = Right(InDt, 4)
    If Val(Dt) <= 0 Then
        MsgBox "Invalid Date - Day cannot be zero", vbInformation
        ValidateDate = False
        Exit Function
    Else
        Select Case Val(mn)
            Case 1, 3, 5, 7, 8, 10, 12:
                If Val(Dt) > 31 Then
                    MsgBox "Invalid Date - Days exceeding the actual number of days in the month", vbInformation
                    ValidateDate = False
                    Exit Function
                End If
            Case 4, 6, 9, 11:
                If Val(Dt) > 30 Then
                    MsgBox "Invalid Date - Days exceeding the actual number of days in the month", vbInformation
                    ValidateDate = False
                    Exit Function
                End If
            Case 2:
                If Val(Yr) Mod 4 = 0 Then
                    If Val(Yr) Mod 100 = 0 And Val(Yr) Mod 400 <> 0 Then
                        If Val(Dt) > 28 Then
                            MsgBox "Invalid Date - Days exceeding the actual number of days in the month", vbInformation
                            ValidateDate = False
                            Exit Function
                        End If
                    Else
                        If Val(Dt) > 29 Then
                            MsgBox "Invalid Date - Days exceeding the actual number of days in the month", vbInformation
                            ValidateDate = False
                            Exit Function
                        End If
                    End If
                Else
                    If Val(Dt) > 28 Then
                        MsgBox "Invalid Date - Days exceeding the actual number of days in the month", vbInformation
                        ValidateDate = False
                        Exit Function
                    End If
                End If
        End Select
    End If
    If Len(Trim(mn)) = 0 Then
        MsgBox "Invalid Date - Month cannot be spaces", vbInformation
        ValidateDate = False
        Exit Function
    ElseIf Val(mn) <= 0 Then
        MsgBox "Invalid Date - Month cannot be zero", vbInformation
        ValidateDate = False
        Exit Function
    ElseIf Val(mn) > 12 Then
        MsgBox "Invalid Date - Month cannot be > 12", vbInformation
        ValidateDate = False
        Exit Function
    End If
    If Val(Yr) <= 1800 Then
        MsgBox "Invalid Date - Year is less than 1800", vbInformation
        ValidateDate = False
        Exit Function
    End If
    ValidateDate = True
End Function




Public Function ValidateTIME(ByVal InTIME As String) As Boolean
    DoEvents
    ValidateTIME = False
    
    Dim HH, MM, SS As String
    HH = Left(InTIME, 2)
    MM = Mid(InTIME, 4, 2)
    SS = Right(InTIME, 2)
    If HH > 24 Or MM >= 60 Or SS >= 60 Then
       ValidateTIME = False
       MsgBox "INVALID TIME ENTERED"
        Exit Function
    End If
    
    
    ValidateTIME = True
End Function




