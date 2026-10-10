VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form tmh_ris_apptqry 
   Caption         =   "APPOINTMENTS QUERY"
   ClientHeight    =   9465
   ClientLeft      =   165
   ClientTop       =   -1290
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crpt_applst 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmd_print 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmd_show 
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox cbo_modality 
      Height          =   315
      ItemData        =   "tmh_ris_apptqry.frx":0000
      Left            =   2400
      List            =   "tmh_ris_apptqry.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ComboBox cbo_deptname 
      Height          =   315
      ItemData        =   "tmh_ris_apptqry.frx":0004
      Left            =   2400
      List            =   "tmh_ris_apptqry.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9300
      TabIndex        =   6
      Top             =   7800
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   1755
      Left            =   7200
      TabIndex        =   8
      Top             =   1320
      Width           =   4335
      Begin MSFlexGridLib.MSFlexGrid msf_timecnt 
         Height          =   1335
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         FormatString    =   "^              TIME SLOT     |^      NO. OF APPOINTMENTS"
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4365
      Left            =   120
      TabIndex        =   0
      Top             =   3060
      Width           =   11460
      Begin MSFlexGridLib.MSFlexGrid msf_timeslot 
         Height          =   4035
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   7117
         _Version        =   393216
         Rows            =   400
         Cols            =   10
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"tmh_ris_apptqry.frx":0008
      End
   End
   Begin MSMask.MaskEdBox msk_aptdate 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1500
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_sysdate 
      Height          =   315
      Left            =   10080
      TabIndex        =   16
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_systime 
      Height          =   315
      Left            =   10200
      TabIndex        =   17
      Top             =   960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label Label5 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   19
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   18
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label22 
      Caption         =   "Message Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   15
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label lbl_msg 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   7800
      Width           =   5595
   End
   Begin VB.Label Label1 
      Caption         =   "Appointment  Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1500
      Width           =   1935
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Modality"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   12
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "APPOINTMENTS IN A DEPARTMENT FOR A MODALITY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1320
      TabIndex        =   10
      Top             =   480
      Width           =   7725
   End
End
Attribute VB_Name = "tmh_ris_apptqry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rq_appslot As rdoQuery
Dim rs_appslot As rdoResultset
Dim flg_appslot As Boolean
Dim RS As rdoResultset
Dim rq_apptmcnt As rdoQuery
Dim rs_apptmcnt As rdoResultset
Dim flg_apptmcnt As Boolean

Dim rq_modality As rdoQuery
Dim rs_modality As rdoResultset
Dim modalityflag As Boolean

Dim rq_dept As rdoQuery
Dim rs_dept As rdoResultset
Dim DEPTFLAG As Boolean


Dim var_deptcode As String
Dim var_modcode As String


Dim str As String
Dim risau2_qr As rdoQuery
Dim risau2_rs As rdoResultset
Sub modalities_in_dept()
    var_deptcode = ""
    msf_timeslot.Rows = 1
    msf_timecnt.Rows = 1
   If Len(Trim(cbo_deptname.Text)) <> 0 Then
        Dim K As Integer
        K = InStr(1, cbo_deptname.Text, "$")
        var_deptcode = Trim(Mid(cbo_deptname, K + 1))
    End If
 '**********SELECTING ONLY MODALITIES BELONGING TO THAT DEPARTMENT*******
    rq_modality(0) = Trim(var_deptcode)
    If modalityflag = False Then
        Set rs_modality = rq_modality.OpenResultset(1, 1)
        modalityflag = True
    Else
        rs_modality.MoreResults
        rs_modality.Requery
    End If
    
    If rs_modality.RowCount > 0 Then
        cbo_modality.clear
        cbo_modality.AddItem "ALL"
        While Not rs_modality.EOF
            cbo_modality.AddItem Trim(rs_modality("MODLITYDES")) & "$" & rs_modality("MODALITYCD")
        rs_modality.MoveNext
        Wend
        lbl_msg.Caption = "SELECT MODALITY"
    Else
        cbo_modality.clear
        lbl_msg.Caption = "NO MODALITIES IN THE SELECTED DEPARTMENT"
        msf_timeslot.Rows = 1
        msf_timecnt.Rows = 1
    End If


End Sub

Private Sub cbo_deptname_Click()
lbl_msg.Caption = ""
Call modalities_in_dept

End Sub



Private Sub cbo_deptname_GotFocus()
If QRYFLAG = True Then
cbo_modality.SetFocus
End If
End Sub

Private Sub cbo_modality_Click()
lbl_msg.Caption = ""
msf_timeslot.Rows = 1
msf_timecnt.Rows = 1

lbl_msg.Caption = "CLICK 'SHOW' COMMAND BUTTON TO SEE THE RESULTS IN THE GRID"
End Sub

Private Sub cbo_modality_GotFocus()
If cbo_modality.ListCount > 0 Then
cbo_modality = cbo_modality.List(0)
End If
lbl_msg.Caption = ""
msf_timeslot.Rows = 1
msf_timecnt.Rows = 1

lbl_msg.Caption = "CLICK 'SHOW' COMMAND BUTTON TO SEE THE RESULTS IN THE GRID"
End Sub

Private Sub cmd_back_Click()


'tmh_ris_appointment.ActiveControl
If QRYFLAG = True Then
tmh_ris_appointment.show
Unload Me
Else
DIS_MAIN.show
Unload Me
End If
'If QRYFLAG = True Then
'Unload Me
'Else
'Unload Me
'End If
QRYFLAG = False
End Sub

Private Sub cmd_print_Click()
Dim ydt As Integer, mdt As Integer, ddt As Integer, pos1 As Integer
Dim dcd1 As String, mcd1 As String
Dim str1 As String
Dim str2 As String
Dim STR3 As String

ydt = Right(msk_aptdate.Text, 4)
mdt = Mid(msk_aptdate.Text, 4, 2)
ddt = Mid(msk_aptdate.Text, 1, 2)
pos1 = InStr(cbo_deptname.Text, "$")
dcd1 = Trim(Mid(cbo_deptname.Text, pos1 + 1))
pos1 = InStr(cbo_modality.Text, "$")
mcd1 = Trim(Mid(cbo_modality.Text, pos1 + 1))


lbl_msg.Caption = "PRINTING..."
DoEvents
     str1 = ip
     str2 = Format(DATE11, "dd/mm/yyyy")
     STR3 = time1
    str1 = str1 & Space(5) & str2 & Space(5) & STR3
    str1 = Space(3) & str1


'{APT_PATDET.APNTDT} = Date (1998,10 ,15 ) and  {APT_PATDET.REQNNO} [1 to 3] = 'RAD' and {APT_PATDET.REQNNO} [4 to 5] = 'XS'
crpt_applst.SelectionFormula = "{APT_PATDET.APNTDT} = Date (" & ydt & "," & mdt & "," & ddt & ") and  {APT_PATDET.REQNNO} [1 to 3] = '" & dcd1 & "' and {APT_PATDET.REQNNO} [4 to 5] = '" & mcd1 & "'"
crpt_applst.Connect = "UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH"

crpt_applst.ReportFileName = "C:\RISREPORTS\cr_applst.rpt"

   crpt_applst.Formulas(0) = "apptdate = '" & msk_aptdate.Text & "'"
   Dim i As Integer
   Dim S As String
   i = InStr(1, cbo_deptname, "$")
   
   S = Left(cbo_deptname, i - 1)
   crpt_applst.Formulas(1) = "dept = '" & S & "'"
    i = InStr(1, cbo_modality, "$")
   
   S = Left(cbo_modality, i - 1)
   crpt_applst.Formulas(2) = "modality = '" & S & "'"
   crpt_applst.Formulas(3) = "curtime = '" & msk_systime & "'"
     crpt_applst.Formulas(4) = "us_id = '" & str1 & "'"
    crpt_applst.WindowState = crptMaximized
    'CRPT_APPLST.Destination =crptToPrinter
    crpt_applst.Action = 1
lbl_msg.Caption = ""

Exit Sub
prtmsg: MsgBox Error(Err)
lbl_msg.Caption = ""

End Sub

Private Sub cmd_show_Click()
Dim j
Dim slno As Integer

lbl_msg.Caption = ""
var_modcode = ""
If Len(Trim(cbo_deptname.Text)) = 0 Then
    j = MsgBox("Please select the DEPARTMENT.", vbCritical)
    Exit Sub
End If
If Len(Trim(cbo_modality.Text)) = 0 Then
    j = MsgBox("Please select the MODALITY.", vbCritical)
    Exit Sub
End If

Screen.MousePointer = 11
lbl_msg.Caption = "Retrieving... Please wait..."
DoEvents
If Len(Trim(cbo_modality.Text)) <> 0 Then
        Dim K As Integer
        K = InStr(1, cbo_modality.Text, "$")
        var_modcode = Trim(Mid(cbo_modality.Text, K + 1))
End If



'********5/7/2001
'rq_apptmcnt(0) = GET_REVDATE(msk_aptdate.Text)
'rq_apptmcnt(1) = var_modcode
'rq_apptmcnt(2) = var_deptcode
'If flg_apptmcnt = False Then
'    Set rs_apptmcnt = rq_apptmcnt.OpenResultset("rq_apptmcnt_stmt", str)
'    flg_apptmcnt = True
'Else
'    rs_apptmcnt.MoreResults
'    rs_apptmcnt.Requery
'End If

'**************5/7/2001 added till here
If cbo_modality.Text = "ALL" Then

str = "select  APNTTM, count(APNTTM)  from rislib.RREQUISITN where APNTDT = '" & GET_REVDATE(msk_aptdate.Text) & "'  and dept_code = '" & var_deptcode & "' group by APNTTM ORDER BY APNTTM"  '**29052001 added
Set rs_apptmcnt = rdb.OpenResultset(str, 1, 3)
Else
str = "select  APNTTM, count(APNTTM)  from rislib.RREQUISITN where APNTDT = '" & GET_REVDATE(msk_aptdate.Text) & "' and MODALITYCD='" & var_modcode & "' and dept_code = '" & var_deptcode & "' group by APNTTM ORDER BY APNTTM"  '**29052001 added
Set rs_apptmcnt = rdb.OpenResultset(str, 1, 3)

End If

 Dim i As Integer
    i = 0
  
  '*******COUNTING NO. OF APPTS IN THE SAME TIME SLOT  ********
  '******tms_arr CONTAINS TIME SLOT VS NO. OF APPTS FOR THAT SLOT*****
  
 ' rs_apptmcnt.MoveFirst '*****SINCE ALREADY MOVED TO LAST ROW
  If rs_apptmcnt.RowCount > 0 Then
  msf_timecnt.Rows = 1
  ReDim tms_arr(rs_apptmcnt.RowCount, 2) As String
    While Not rs_apptmcnt.EOF
        If IsNull(rs_apptmcnt(0)) = False Then
        tms_arr(i, 0) = rs_apptmcnt(0)
        tms_arr(i, 1) = rs_apptmcnt(1)
        msf_timecnt.AddItem rs_apptmcnt("APNTTM") & vbTab & rs_apptmcnt(1) '& vbTab & var_patname & vbTab & var_patage & vbTab & rs_appslot("sex") & vbTab & Trim(rs_appslot("ser_det"))
        Else
        tms_arr(i, 0) = ""
        tms_arr(i, 1) = rs_apptmcnt(1)
        msf_timecnt.AddItem "-" & vbTab & rs_apptmcnt(1) '& vbTab & var_patname & vbTab & var_patage & vbTab & rs_appslot("sex") & vbTab & Trim(rs_appslot("ser_det"))
        End If
        
        If ((i + 1) Mod 2) = 0 Then
            msf_timecnt.row = i + 1
            msf_timecnt.CellBackColor = &H8000000F '&HC000&
        End If
        
        rs_apptmcnt.MoveNext
        i = i + 1
    Wend
    Else
    msf_timecnt.Rows = 1
    
    End If
   If cbo_modality.Text = "ALL" Then
      
   str = "SELECT APNTDT,APNTTM,REQNNO,caseno,tag,first_name,mid_name,last_name,birth_date,sex,ser_code,ser_det,EXAMINATN,PAT_CATG FROM rislib.APT_PATDET WHERE APNTDT = '" & GET_REVDATE(msk_aptdate.Text) & "'  and dept_code = '" & var_deptcode & "' order by APNTTM"  '" & GET_REVDATE(lbl_aptdate.Caption), 1, 1)'**add 29052001
   Set rs_appslot = rdb.OpenResultset(str, 1, 3)
   Else
   str = "SELECT APNTDT,APNTTM,REQNNO,caseno,tag,first_name,mid_name,last_name,birth_date,sex,ser_code,ser_det,EXAMINATN,pat_catg   FROM rislib.APT_PATDET WHERE APNTDT = '" & GET_REVDATE(msk_aptdate.Text) & "' and MODALITYCD='" & var_modcode & "' and dept_code = '" & var_deptcode & "' order by APNTTM"  '" & GET_REVDATE(lbl_aptdate.Caption), 1, 1)'**add 29052001
   Set rs_appslot = rdb.OpenResultset(str, 1, 3)
   End If
    msf_timeslot.Rows = 1
    Dim var_apttm
    Dim var_msfrowindex As Integer
    Dim var_patname As String
    Dim var_patage As Integer
    
    If rs_appslot.RowCount > 0 Then
    var_msfrowindex = 0
    slno = 0
    While Not rs_appslot.EOF
        var_apttm = rs_appslot("APNTTM")
        var_patname = rs_appslot("tag") & " " & Trim(rs_appslot("first_name")) & " " & Trim(rs_appslot("mid_name")) & " " & Trim(rs_appslot("last_name"))
        var_patage = DateDiff("yyyy", rs_appslot("birth_date"), tmh_ris_appointment.msk_sysdate.Text)
        slno = slno + 1
        'msf_timeslot.AddItem rs_appslot("APNTTM") & vbTab & rs_appslot("CASENO") & vbTab & var_patname & vbTab & var_patage & vbTab & rs_appslot("sex") & vbTab & Trim(rs_appslot("ser_det")) & vbTab & req_chglen(Trim(rs_appslot("REQNNO"))) & vbTab & Trim(rs_appslot("EXAMINATN"))
         msf_timeslot.AddItem CStr(slno) & vbTab & rs_appslot("APNTTM") & vbTab & rs_appslot("CASENO") & vbTab & var_patname & vbTab & var_patage & vbTab & rs_appslot("sex") & vbTab & Trim(rs_appslot("pat_catg")) & vbTab & Trim(rs_appslot("ser_det")) & vbTab & req_chglen(Trim(rs_appslot("REQNNO"))) & vbTab & Trim(rs_appslot("EXAMINATN"))
        rs_appslot.MoveNext
    Wend
    Else
        Dim p
        lbl_msg.Caption = ""
        Screen.MousePointer = 1
        p = MsgBox("Appointments not yet issued on this day" & Chr(13) & "for this modality in this department", vbInformation)
    End If

'End If

'******marking colors for different times on grids


 If rs_apptmcnt.RowCount > 0 Then
'    For j = 0 To rs_apptmcnt.RowCount Step 2
    For j = 1 To rs_apptmcnt.RowCount Step 2

    msf_timeslot.row = 1
    For K = 0 To msf_timeslot.Rows
    If msf_timeslot.Text = tms_arr(j, 0) Then
         msf_timeslot.CellBackColor = &H8000000F '&HC000&
    End If
        If msf_timeslot.row <> msf_timeslot.Rows - 1 Then
        msf_timeslot.row = msf_timeslot.row + 1
        End If
        Next K
    Next j

End If

lbl_msg.Caption = ""
Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
Dim rowno As Integer
Dim timeslot As Integer

Dim tms_arr() As String

Dim var_deptcode As String '**29052001

'Main
Dim str As String
appslt_viewed = 1
Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = RS("dt1") '**add 24/11/2k
time1 = Format(RS("TM1"), "hh:mm:ss") '***2/2/99
msk_systime.Text = time1
msk_sysdate.Text = DATE11




flg_appslot = False
flg_apptmcnt = False

 '***5/7/2001********add
 'str = "select  APNTTM, count(APNTTM)  from RREQUISITN where APNTDT = '" & GET_REVDATE(msk_aptdate.Text) & "' and MODALITYCD='" & var_modcode & "' and dept_code = '" & var_deptcode & "' group by APNTTM"  '**29052001 added
'str = "select  APNTTM, count(APNTTM)  from RREQUISITN where APNTDT = ? and MODALITYCD = ? and dept_code = ? group by APNTTM"   '**29052001 added
'Set rq_apptmcnt = rdb.CreateQuery("apptmcnt_stmt", str)
'**********
 
'str = "select dept_code,dept_name from PATOOLSLIB.ma106 where dept_code in(select distinct dept_code from RMODALITY)"'****department for which modalities are present
str = "select dept_code,dept_name from PATOOLSLIB.ma106 where (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' OR DEPT_CODE='IZZ' )" '*********all departments.
Set rq_dept = rdb.CreateQuery("rq_dept_reqstmt1", str)

Set rs_dept = rq_dept.OpenResultset(1, 1)
While rs_dept.EOF = False
cbo_deptname.AddItem Trim(rs_dept("dept_name")) & "$" & rs_dept("dept_code")
rs_dept.MoveNext
Wend

DEPTFLAG = False
modalityflag = False
str = "select distinct MODALITYCD,MODLITYDES from rislib.RMODALITY where DEPT_CODE = ?" '**ADD 4/7/2001
Set rq_modality = rdb.CreateQuery("rq_modality_reqstmt", str)

'STR = "select * from rislib.risauth where usrgroup=? and scrn_desc=?"
'Set risau2_qr = rdb.CreateQuery("risau2_qr1", STR)
'risau2_qr(0) = Trim(USRGRP)
'risau2_qr(1) = "APPOINTMENT QUERY"
'Set risau2_rs = risau2_qr.OpenResultset(1, 1)
'If risau2_rs.RowCount = 0 Then
'    MsgBox "No Button Authorisation for this User Group & Screen"
'     End
'End If
cmd_show.Enabled = True
'CMDPRINT
cmd_back.Enabled = True


End Sub

Private Sub Form_Unload(cancel As Integer)
QRYFLAG = False
Dim RS As rdoResultset
Dim RQ As rdoQuery
For Each RS In rdb.rdoResultsets
    RS.Close
Next

For Each RQ In rdb.rdoQueries
    RQ.Close
Next

End Sub


Private Sub msk_aptdate_GotFocus()
lbl_msg.Caption = "ENTER DATE ON WHICH APPOINTMENTS ARE TO BE VIEWED"
End Sub

Private Sub msk_aptdate_LostFocus()
   Dim ind_cbodept As Integer

lbl_msg.Caption = ""
'*****5/7/2001***ADD
If msk_aptdate.FormattedText <> "  /  /    " Then
If Len(Trim(cbo_deptname.Text)) = 0 Then

'MsgBox Trim(msk_aptdate.FormattedText)
If Trim(cbo_deptname) = "" Then
    For ind_cbodept = 0 To (cbo_deptname.ListCount - 1)
        If Trim(cbo_deptname.List(ind_cbodept)) = "GEN. MEDICINE SER.$CAZ" Then
            cbo_deptname.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
End If
'****5/7/2001 ADDED TILL HERE



If cbo_deptname.Enabled = True Then
cbo_deptname.SetFocus
End If

Else
'If cmd_close.Enabled = True Then
'cmd_close.SetFocus
'End If

End If
End Sub

Public Sub CMDPRINT()
If risau2_rs("PRINT") = 1 Then
cmd_print.Enabled = True
Else
cmd_print.Enabled = False
End If
End Sub
