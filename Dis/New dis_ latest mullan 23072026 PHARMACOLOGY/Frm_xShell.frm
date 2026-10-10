VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Frm_xShell 
   Caption         =   "Shell_Program"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6795
   Icon            =   "Frm_xShell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Winsock11 
      Height          =   480
      Left            =   840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   1200
      Width           =   1200
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2640
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_xShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EMP_CCNo
Dim strEXE_UPDATE_DT
Dim RemoteSourceFile
Dim sProcess
Dim newCCNo_EXE_UPDATE
Dim runExe
Dim adminPwd
Dim newEXE_UPDATE_DT
Dim strGetServTm
Dim IPADDR
Dim newIPADDR_EXEUPDATE
Dim strGetServDT
Public RE As rdoEnvironment
Public RCon As rdoConnection
Public RCon1 As rdoConnection
Public RConAct As rdoConnection
Dim EXETST As Boolean
Dim sLiberationKey As String
'Dim G_ReleaseDt As String
Private Sub Form_Load()
    Dim filesys
    EXETST = False
    Call Main
    sProcess = 0
    ''''''''''''''''''''''''''''''''''''''ADDED BY VARINDER
    Dim EIPADD, EADD, EFOLD
    IPQRY = "SELECT * FROM FMSLIB.EXE_IP"
    Set ERQ = RCon.CreateQuery("", IPQRY)
        Set IP_EXE = ERQ.OpenResultset(1, 1)
        If IP_EXE.RowCount > 0 Then
        EIPADD = Trim(IP_EXE("IPADD"))
        EFOLD = Trim(IP_EXE("FOLDER"))
        EADD = EIPADD & EFOLD
        End If
    '''''''''''''''''''''''''''''''''''''''''''''
    strGetServTm = Format(Now, "HH:MM:SS")    'G_GetServTm
    strGetServDT = Format(Now, "YYYY-mm-dd")
    G_IpAddr = Winsock1.LocalIP
    sLiberationKey = Trim$(GetSetting("DIS", "Lock", "LiberationKey", ""))
    Set filesys = CreateObject("Scripting.FileSystemObject")
    '***--- Commented and added by Bhushan on 20/05/2022 to fix the exe location on Mullanpur server
    '    RemoteSourceFile = "\\10.80.1.62\MULLANPUR_EXES\DIS\DIS.EXE"
    RemoteSourceFile = EADD & "DIS\DIS.EXE"
    '***--- Till here commented and added by Bhushan on 20/05/2022 to fix the exe location on Mullanpur server
    newIPADDR_EXEUPDATE = Format("2015-10-31", "YYYY-mm-dd")
    If Trim(Format(strGetServTm, "HH:MM:SS")) > Trim(Format("9:00:00", "HH:MM:SS")) And Trim(Format(strGetServTm, "HH:MM:SS")) < Trim(Format("23:00:00", "HH:MM:SS")) Then
        If filesys.FileExists(RemoteSourceFile) Then
            Call CheckExeUpdates(G_IpAddr)
        End If
    End If
    cmdarg = COMMAND$
    'cmdarg = "210056:IT/00001:B2    "
    If Trim(cmdarg) <> "" Then
                USERID_PARM = Trim(Mid(cmdarg, 1, 6))
                CASENO_PARM = Trim(Mid(cmdarg, 8, 8))
                TOKEN_NO_PARM = Trim(Mid(cmdarg, 17, 8))
    End If
    If Trim(USERID_PARM) <> "" And Trim(CASENO_PARM) <> "" And Trim(TOKEN_NO_PARM) <> "" Then
          logonfrm.show
           'Call Form4.Form_Load 'Commented By Bhushan on 21/05/2021 to call the login page
            
           Unload Me
           Exit Sub
    Else
         If EXETST = True Then
              Call openExe
              End
         Else
              'Form4.Show 'Commented By Bhushan on 21/05/2021 to call the login page
              logonfrm.show
         End If
         Unload Me
    End If
    
End Sub
Sub Main()

On Error GoTo errmsg
    
    Set RE = rdoEnvironments(0)
    RE.CursorDriver = rdUseOdbc
      '***--- Below DSNAME changed By Bhushan on 19/05/2022 to provide orgnization based DSN name, Here we check the exe update at Mullanpur server only
      'Set RCon = RE.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=ECIL;PWD=ECIL;database=DISLIB")
      Set RCon = RE.OpenConnection(DSNAME:="MULLANPUR", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=ECIL;PWD=ECIL;database=DISLIB")
      '***--- Till here DSNAME changed By Bhushan on 19/05/2022 to provide orgnization based DSN name, Here we check the exe update at Mullanpur server only
  '  Set RCon1 = RE.OpenConnection(dsname:="DISDSN", prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=DISuser;PWD=DISuser;database=DISLIB")
    RCon.QueryTimeout = 1500
    
   
    
    
    
    Exit Sub

errmsg:

    MsgBox "There is a Communication Link Failure", vbCritical, G_Title
    End
    
End Sub

Private Sub CheckExeUpdates(IPADDR)
    Dim RQ As rdoQuery, RSV As rdoResultset, Qry As String
    Dim RQExe As rdoQuery, RSVExe As rdoResultset, QryExe As String
    Dim RQ_InsCCNo As rdoQuery, RSV_InsCCNo As rdoResultset, Qry_InsCCNo As String
    Dim RQUsrSts As rdoQuery, StrUsrSts As String
   
    Dim RSV_Update As rdoResultset, Qry_Update As String
    Dim RQ_Update As rdoQuery
    Dim RQUpdateStr As rdoQuery
    Dim QryUpdateStr As String
    
    
CASE1:  Qry = "Select * from FMSLIB.EXE_ACCESS where IPADDR='" & IPADDR & "' AND EXENAME='DIS'"
        Set RQ = RCon.CreateQuery("", Qry)
        Set RSV = RQ.OpenResultset(1, 1)
        If RSV.RowCount > 0 Then
            While Not RSV.EOF
                strEXE_UPDATE_DT = IIf(IsNull(Trim(RSV("EXEUPDATE"))), "", Trim(RSV("EXEUPDATE")))
                sProcess = 1
                RSV.MoveNext
            Wend
        Else
            Qry_InsCCNo = "Insert Into FMSLIB.EXE_ACCESS(IPADDR,EXEUPDATE,EXENAME) Values('" & IPADDR & "','" & newIPADDR_EXEUPDATE & "','DIS')"
            Set RQ_InsCCNo = RCon.CreateQuery("", Qry_InsCCNo)
            RQ_InsCCNo.Execute
            
            GoTo CASE1
        End If
    
        If strEXE_UPDATE_DT <> "" And sProcess = 1 Then
            QryExe = "Select * from FMSLIB.FMS_EXEDT where EXE_DATE >'" & Format(strEXE_UPDATE_DT, "YYYY-MM-DD") & "' AND IsActive='1' AND FLD1='DIS' ORDER BY EXE_DATE DESC FETCH FIRST 1 ROWS ONLY "
            Set RQExe = RCon.CreateQuery("", QryExe)
            Set RSVExe = RQExe.OpenResultset(1, 1)
            If RSVExe.RowCount > 0 Then
                If MsgBox("Some changes done in DIS." & vbCrLf & "New Exe is available please update it ", vbYesNo + vbQuestion, G_Title) = vbYes Then
                    While Not RSVExe.EOF
                        newEXE_UPDATE_DT = Format(IIf(IsNull(Trim(RSVExe("EXE_DATE"))), "", Trim(RSVExe("EXE_DATE"))), "YYYY-MM-DD")
                        Call RenameFile("DIS.exe", RSVExe("EXE_DATE"))
                        Call CopyFile
                        sProcess = 1
                        RSVExe.MoveNext
                      '  G_ReleaseDt = RSVExe("EXE_DATE")
                    Wend
                     EXETST = True
                Else
                    sProcess = 0
                End If
            Else
                If (sLiberationKey = "") Or (G_IpAddr <> sLiberationKey) Then
                    'Call RenameFile("DIS.exe", strGetServDT)
                    'Call CopyFile
                End If
                
                sProcess = 0
            End If
        End If
    
        RSVExe.Close
        RQExe.Close
        G_IpAddr = Winsock1.LocalIP
    
        If sProcess = 1 Then
             StrUsrSts = "Update FMSLIB.EXE_ACCESS Set EXEUPDATE='" & Format(newEXE_UPDATE_DT, "YYYY-MM-DD") & "' Where IPADDR='" & Trim(G_IpAddr) & "' AND EXENAME='DIS'"
            Set RQUsrSts = RCon.CreateQuery("", StrUsrSts)
            
            RQUsrSts.Execute
        End If
        
        
    
    sProcess = 0
    

End Sub


Private Sub RenameFile(sFile, sdate)
Dim strSource As String
Dim strDest As String
Dim strReName As String
Dim filesys
Dim strDestination As String
Dim filesys1
Dim rnd As Integer
    
Set filesys1 = CreateObject("Scripting.FileSystemObject")
Set filesys = CreateObject("Scripting.FileSystemObject")

strSource = App.PATH & "\"
strDestination = App.PATH & "\DIS.exe"

rnd = Random(1, 1000)

If filesys.FileExists(RemoteSourceFile) Then
    strDest = "DIS.exe"
    If filesys.FileExists(App.PATH & "\DIS.exe") Then
        strReName = "DIS" & Replace(sdate, "/", "_") & rnd & ".exe"

        strDest = Replace(strDest, "DIS.exe", strReName)
        strDest = strSource & strDest

        
        Name strDestination As strDest
     End If
Else
    MsgBox "Exe not available in Source"
    End

End If

End Sub


Private Sub CopyFile()
Dim FilePath
Dim sourceFile
FilePath = App.PATH & "\"


Dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject")

sourceFile = RemoteSourceFile

If filesys.FileExists(sourceFile) Then
    filesys.CopyFile sourceFile, FilePath
End If



End Sub


Private Sub openExe()
EXETST = False
  runExe = App.PATH & "\DIS.exe"
Shell runExe, 1
End Sub



Function Random(Lowerbound As Long, Upperbound As Long)
Randomize
Random = Int(rnd * Upperbound) + Lowerbound
End Function


