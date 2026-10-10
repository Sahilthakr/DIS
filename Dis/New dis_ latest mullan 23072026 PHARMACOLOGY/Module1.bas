Attribute VB_Name = "Module1"
Public rnv As rdoEnvironment
Public con As rdoConnection
Public dept As rdoQuery
Public code As rdoQuery

Public Dt As rdoQuery
Public dt1 As rdoQuery

Public unit As rdoQuery
Public item As rdoQuery
Public itemchk As rdoQuery
Public expiry As rdoQuery
Public area As rdoQuery
Public areaname1 As rdoQuery
Public RS As rdoResultset
Dim str As String
Dim month As Integer
Public ip As String
Dim dd As Long
Public strlogin As String

Public mainuser As String
Public mainpass As String

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'UPDATE FOR 2D,ENDOSCOPY DONE BY SHWETA 02-04-2013
Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEC_QR1 As rdoQuery
'''''''

Private Sub Main()
mainuser1 = "TMH"
mainpass1 = "TMH"
Set RdEnv = rdoEnvironments(0)
RdEnv.CursorDriver = rdUseOdbc
Set RdCn = RdEnv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.1.4.2;UID=" & mainuser & ";PWD=" & mainpass & ";database=DISLIB")
'Set con = rnv.OpenConnection(DSNAME:="stores", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.1.4.2;UID=" & mainuser & ";PWD=" & mainpass & ";database=TOOLSADMN")End Sub
End Sub

Public Function UPDRTIMELOG(ByVal REQNO As String, ByVal MOD1 As String) As Boolean
If MOD1 = "EXAM" Then
  str = "UPDATE RISLIB.RTIMELOG SET EXAMDATE=CURRENT DATE,EXAMTIME=CURRENT TIME ,SAMRECDATE=CURRENT DATE, SAMRECTIME=CURRENT TIME WHERE REQNNO=?"
  ElseIf MOD1 = "SAVE" Then
  str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=CURRENT DATE,SAVREPTIME=CURRENT TIME  WHERE REQNNO=?"
  
ElseIf MOD1 = "COMMIT" Then
  str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=CURRENT DATE,COMREPTIME=CURRENT TIME  WHERE REQNNO=?"
Else

End If
On Error GoTo MSG1

Set UPDRTIMEC_QR = rdb.CreateQuery("", str)
UPDRTIMEC_QR(0) = REQNO
UPDRTIMEC_QR.Execute
UPDRTIMELOG = True
Exit Function
MSG1:
MsgBox Err.Description
UPDRTIMELOG = False
End Function
