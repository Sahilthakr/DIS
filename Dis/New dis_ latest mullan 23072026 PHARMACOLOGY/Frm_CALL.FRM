VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form logonfrm 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "LOGIN"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_CALL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   1725
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H0080FF80&
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   0
      Top             =   2160
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      FillColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   8835
      TabIndex        =   5
      Top             =   120
      Width           =   8895
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   4440
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.PictureBox Picture2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   600
         Picture         =   "Frm_CALL.frx":0442
         ScaleHeight     =   1875
         ScaleWidth      =   1755
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HBCHRC, Mullanpur"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "Save & Commit Tiem & pharmicology test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "   HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE, MULLANPUR"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8295
      End
      Begin VB.Image Image1 
         Height          =   1965
         Left            =   6120
         Picture         =   "Frm_CALL.frx":3380
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UserName :-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "DIS"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   50
         Left            =   6480
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "DIAGNOSTIC INFORMATION SYSTEM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Maintained by I.T. Department"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   4800
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   3255
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   7935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RELEASE DATE:23/07/2026"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.Label LBL_MSG 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
End
Attribute VB_Name = "logonfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PWDFORM As String
Public usrid As String
Dim STRSQL As String
Dim SELDOCDTL_RS As rdoResultset
Public MYRS As rdoResultset
Dim Cqy As New rdoQuery

'Public UID As String
Dim pwschk As Boolean
Dim flags As Integer
Public FLG_PWD As Integer
Dim SCREEN1 As String
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdlogin_Click()
On Error GoTo MSG:
UID = "A" + txtusername.Text
pwschk = True
PASSWD = txtpassword.Text
usr_name = txtusername
'LBL_MSG.Caption = "WORKING..." 'Shraddha 17/07/09


'************* Shraddha Pandit 17/07/09 *************
ProgressBar1.Visible = True 'Shraddha Pandit 17/07/09
ProgressBar1.Min = 0
ProgressBar1.Max = 100
ProgressBar1.Value = 0

While ProgressBar1.Value < 30
    MYDELAY (1)
    lbl_msg.Caption = "Password Checking..."
    ProgressBar1.Value = ProgressBar1.Value + 10
Wend
'****************************************************


DoEvents
Main
pwdchk
 If pwschk = True Then
 Else
     ProgressBar1.Value = 0
    lbl_msg.Caption = ""
    txtusername.SelStart = 0
    txtusername.SetFocus
    flags = 0
    FLG_PWD = 0
    txtusername.SetFocus
 Exit Sub
 End If
 
   '********* Shraddha Pandit 17/07/09 **************
While ProgressBar1.Value < 80
    MYDELAY (1)
    lbl_msg.Caption = "Authorising ..."
    ProgressBar1.Value = ProgressBar1.Value + 10
Wend
'*************************************************
 Dim SELDOCDTL_RS As rdoResultset
Set SELDOCDTL_RS = rdb.OpenResultset("SELECT EMPCODE FROM EMPLIB.EMPOADINFO WHERE EMPCODE='" & txtusername.Text & "' AND CESSYN='4'   ", 1, 1)
    If Trim(txtusername) <> "" Then
        empcode = False
        'SELDOCDTL_RS.MoveFirst
        While Not SELDOCDTL_RS.EOF
            If Trim(txtusername.Text) = Trim(SELDOCDTL_RS(0)) Then
                empcode = True
                txtpassword.SetFocus
                userid = txtusername
                'Exit Sub
            End If
            SELDOCDTL_RS.MoveNext
        Wend
        If empcode = False Then
            MsgBox "User Id is not available", vbInformation
            txtusername.SetFocus
            txtusername = ""
            txtpassword = ""
            Exit Sub
        End If
        Else
        Exit Sub
    End If
    
    SELDOCDTL_RS.Close
    Set SELDOCDTL_RS = Nothing
 
 '**************************************
  
'''
Dim RS_dislnk As rdoResultset
'Set RS_dislnk = rdb.OpenResultset("SELECT * from  dislib.dislnkusr WHERE USR_ID='" & UID & "'", 1, 1)
Set RS_dislnk = rdb.OpenResultset("SELECT * from  dislib.dislnkusr WHERE USR_ID ='" & UID & "' ", 1, 1)
If Not RS_dislnk.EOF Then
 USRGRP = "SYSADM" 'MYRS.rdoColumns(0).Value
        'USRGRP = ""
        USRDEPT = Trim(RS_dislnk("DEPT_CODE")) 'USERNAME
        'usr_name = txtUserName.Text '** comment sangrur by rahul dated on 02.04.2020
        usr_name = Trim(RS_dislnk("USERNAME"))
        
        PSWD = txtpassword.Text
        flags = 1
        FLG_PWD = 1
        usrid = "A" + txtusername.Text
        PASSWD = txtpassword.Text
Else
MsgBox ("You Are Not authorised to Use ")
Exit Sub
End If

'************* Shraddha Pandit 17/07/09 *************
   While ProgressBar1.Value < 99
    MYDELAY (1)
    lbl_msg.Caption = "Successful Login..."
    ProgressBar1.Value = ProgressBar1.Value + 10
Wend
   ProgressBar1.Value = 100 'Shraddha Pandit 17/07/09

'************* Shraddha Pandit 17/07/09 *************



'If Trim(txtUserName) = "DISCLK" Then
'    txtUserName.Text = ""
'    txtPassword.Text = ""
'    Exit Sub
'End If

'If Trim(txtUserName) = "DISCLK" Or Trim(txtUserName) = "DISDOC" Or Trim(txtUserName) = "DISADM" Then
'
'Else
'    rdb.QueryTimeout = 3600
    ''''' Commented at Sangrur
                    '    STRSQL = "CALL dislib.SECPROC('" & UID & "' , '" & PASSWD & "')"
                    '    Set MYRS = rdb.OpenResultset(STRSQL, rdOpenStatic, rdConcurValues)
                    '    If MYRS.RowCount = 0 Then
                    '        MsgBox "Invalid Login/Password Combination"
                    '        '********** Shraddha Pandit 17/07/09 **********
                    '        ProgressBar1.Value = 0
                    '        lbl_msg.Caption = ""
                    '        '**********************************************
                    '        txtusername.SelStart = 0
                    '        txtusername.SelLength = Len(Trim(txtusername))
                    '        txtusername.SetFocus
                    '        flags = 0
                    '        FLG_PWD = 0
                    '        txtusername.SetFocus
                    '        Exit Sub
                    '    Else
                    '         USRGRP = MYRS.rdoColumns(0).Value
                    '        'USRGRP = ""
                    '         USRDEPT = MYRS.rdoColumns(1).Value
                    '
                    '        usr_name = txtusername.Text '**17/7/99
                    '        PSWD = txtpassword.Text
                    '        flags = 1
                    '        FLG_PWD = 1
                    '        usrid = txtusername.Text
                    '        PASSWD = txtpassword.Text
                    '    End If
                    
                    
                    
                    
                    
                    
  ''' COMMENT FOR SANGRUR BY DATED 02.04.2020    FOR LOGIN
                    
''        USRGRP = "SYSADM" 'MYRS.rdoColumns(0).Value
''        'USRGRP = ""
''        USRDEPT = "FEC" 'MYRS.rdoColumns(1).Value
''
''        usr_name = txtUserName.Text '**17/7/99
''        PSWD = txtPassword.Text
''        flags = 1
''        FLG_PWD = 1
''        usrid = txtUserName.Text
''        PASSWD = txtPassword.Text

 ''' END  COMMENT FOR SANGRUR BY DATED 02.04.2020    FOR LOGIN
'End If


'If txtUserName = "DISCLK" Then
' If txtPassword <> "DISCLK" Then
'    MsgBox " Incorrect password"
'    txtPassword.SetFocus
'    Exit Sub
' End If
'End If
'
'If txtUserName = "DISDOC" Then
' If txtPassword <> "BILL" Then
'    MsgBox " In correct password"
'    txtPassword.SetFocus
'    Exit Sub
' End If
'End If
'
'If txtUserName = "DISADM" Then
' If txtPassword <> "MARS" Then
'    MsgBox " In correct password"
'    txtPassword.SetFocus
'    Exit Sub
' End If
'End If
 
 
'ProgressBar1.Value = 100 'Shraddha Pandit 17/07/09
 
  ''''added by charulata
Dim PCARD As New PISWriteCard.ClsWriteCard
If txtusername = "999999" Then
    MsgBox ("Place the Master card on the reader.")
    Strkey = PCARD.ReadMasterKey("admin1", "tmhit1")
Else
    Strkey = "1A4F101A4F10"
End If
Set PCARD = Nothing
'''upto here

 DIS_MAIN.show
 
 Unload Me
 Exit Sub
MSG:
 MsgBox Err.Description
 'Resume Next
End Sub

Private Sub Command1_Click()
MICROMANUALREC.show
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
flags = 0
lbl_msg.Caption = ""
ProgressBar1.Visible = False 'Shraddha Pandit 17/07/09



'******* Added by Shraddha Pandit on 14/09/09 to check regional settings *******
    
If Not IsDDMMYYYY Then End

'*******************************************************************************

End Sub
'******* Added by Shraddha Pandit on 14/09/09 *******
Public Function IsDDMMYYYY() As Boolean
'Checking Date Format in Regional Settings(DD/MM/YYYY)
On Error Resume Next
Dim X As Long
Dim Dt As String
    IsDDMMYYYY = False
    
    Dt = CStr(Date)
    If Len(Dt) = 10 And InStr(3, Dt, "/") = 3 And InStr(6, Dt, "/") = 6 And Day("10/11/2002") = 10 Then
        IsDDMMYYYY = True
    Else
        MsgBox "Please set the Date in Regional Settings as " & vbCrLf & vbTab & "'DD/MM/YYYY'" & vbCrLf & vbCrLf & "Run the application again.", vbCritical ', MsgTitle
       ' X = Shell("Rundll32.exe shell32.dll,Control_RunDLL intl.cpl ,4")
    End If
End Function
'******************************************************

Private Sub lblLabels_Click(Index As Integer)

End Sub






Private Sub txtPassword_KeyPress(KeyAscii As Integer)
'If KeyAscii >= 97 And KeyAscii <= 122 Then
'    KeyAscii = KeyAscii - 32
'End If
End Sub

Private Sub txtPassword_LostFocus()
UID = txtusername.Text
PASSWD = txtpassword.Text
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then
KeyAscii = KeyAscii - 32
End If
End Sub

Public Sub Main()
Dim er As rdoError
On Error GoTo CnEh
Set renv = rdoEnvironments(0)
renv.CursorDriver = rdUseOdbc
       
        If Mid(Trim(UID), 2) <> "DISCLK" Then
'            Set rdb = renv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=" & Trim(UID) & ";PWD=" & Trim(UID) & ";Database = EC_TMH")      'THIS CODE IS USED TO CONTECT TO WITH SECURITY.
            Set rdb = renv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=PATOOLS;PWD=PATOOLS;Database = 10.80.1.2")
           ' Set rdb = renv.OpenConnection(DSNAME:="ACTDISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH")
            LOGIN = Trim(UID)
        Else
            Set rdb = renv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=PATOOLS;PWD=PATOOLS;Database = 10.80.1.2")
          '  Set rdb = renv.OpenConnection(DSNAME:="ACTDISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH")
            LOGIN = "PATOOLS"
        End If
Exit Sub

CnEh:
        Debug.Print Err, ERROR
        For Each er In rdoErrors
            Debug.Print er.Description, er.Number
        Next er
        MsgBox "Either this user dosen't exist on Server or " & Chr(13) & " there is a communication link failure"
    MsgBox ERROR(Err)
    End
   
   
   
   
End Sub
'********* Shraddha Pandit 17/07/09 *************
Public Sub MYDELAY(inttime As String)
Dim sttime
Dim endtime1 As Long
Dim TESTTIME As Long
'sttime = Timer

endtime1 = Timer + inttime
TESTTIME = Timer

Do While Timer < endtime1
 '   Text1 = endtime1
 '   Text2 = TESTTIME
    DoEvents
  '  TESTTIME = Timer
Loop

End Sub
'*************************************************
Sub pwdchk()
Dim UserData As String
Dim USRARR(10)  As String
Dim PASSARR(10) As String
Dim LNUSR As Integer
Dim LNPAS As Integer
Dim STRUSR As String
Dim STRPASS As String
Dim NUMUSR As Integer
Dim NUMPASS As Integer


Dim RS_LOG As rdoResultset
Set RS_LOG = rdb.OpenResultset("SELECT TMHPWD from  PATOOLSLIB.LOGPATIENT WHERE USERID='" & txtusername.Text & "'", 1, 1)
If Not RS_LOG.EOF Then
    If Trim(RS_LOG(0)) <> Trim(txtpassword.Text) Then
        MsgBox ("WRONG PASSWORD.")
        txtpassword.Text = ""
        txtusername.SelStart = 0
        txtusername.SelLength = Len(Trim(txtusername))
        txtusername.SetFocus
        flags = 0
        FLG_PWD = 0
        pwschk = False
        Exit Sub
        End If
       
Else
    MsgBox ("INVALID USER/PASSWORD COMBINATION")
     txtusername.SelStart = 0
    txtusername.SelLength = Len(Trim(txtusername))
    txtusername.SetFocus
    flags = 0
    FLG_PWD = 0
     pwschk = False
    Exit Sub
End If


End Sub

