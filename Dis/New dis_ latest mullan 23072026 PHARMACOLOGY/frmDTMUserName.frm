VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmDTMUserName 
   Caption         =   "frmDTMUserName"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame COMMAND 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   9
      Top             =   7680
      Width           =   9855
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7800
         MouseIcon       =   "frmDTMUserName.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6480
         MouseIcon       =   "frmDTMUserName.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3960
         MouseIcon       =   "frmDTMUserName.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "frmDTMUserName.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1440
         MouseIcon       =   "frmDTMUserName.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   240
         MouseIcon       =   "frmDTMUserName.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5160
         MouseIcon       =   "frmDTMUserName.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   9015
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2055
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton opt_patient 
            Caption         =   "Patient"
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.ComboBox cmb_rh 
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox cmb_abo 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   2400
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Msk_caseno 
         Height          =   375
         Left            =   3000
         TabIndex        =   0
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_dtmno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   1
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "DTM No."
         Height          =   255
         Left            =   6000
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Case No."
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lbl_ccno 
         Caption         =   "ABO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lbl_name 
         Caption         =   "Rh"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6360
         TabIndex        =   7
         Top             =   2400
         Width           =   615
      End
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   360
      Width           =   10215
   End
   Begin VB.Label lbl_cap 
      AutoSize        =   -1  'True
      Caption         =   "DTM NO"
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
      Left            =   4800
      TabIndex        =   15
      Top             =   1080
      Width           =   1185
   End
End
Attribute VB_Name = "frmDTMUserName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'for view mode
'Dim SELDTMUSER_QR As rdoQuery
'Dim SELDTMUSER_RS As rdoResultset
'Dim SELDTMUSER_FLAG As Boolean
''to insert data into dtmuser table
'Dim insdtmuser_qr As rdoQuery
'Dim insdtmuser_rs As rdoResultset
'Dim insdtmuser_flag As Boolean
''to delete record
'Dim deldtmuser_qr As rdoQuery
''to modify the record
'Dim upddtmuser_qr As rdoQuery
'Dim upddtmuser_flag As Boolean
Dim modsave_flag As Boolean
Dim newsave_flag As Boolean
Dim VIEW_FLAG As Boolean
Dim NEWFLAG As Boolean
Dim RS As rdoResultset
Dim QR As rdoQuery
Dim rs_flag As Boolean

Dim dtm_flag As Boolean
Dim INSDTMNO_QR As rdoQuery
Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset

'***** Shraddha 06/08/10 *****
Dim SEL_DTMTAB_QR As rdoQuery
Dim SEL_DTMTAB_RS As rdoResultset
Dim UPD_DTMNO_QR As rdoQuery
'*****************************

Dim SELDTMNO_FLAG As Boolean
Dim CASENO As String
Dim UPDDTMNO_QR As rdoQuery


Private Sub fieldclear()
txt_dtmno.Text = ""
cmb_abo.Text = ""
cmb_rh.Text = ""
If opt_patient.Value = True Then
msk_caseno.Mask = ""
msk_caseno.Text = ""
msk_caseno.Mask = ">??/#####"
End If
If opt_staff.Value = True Then
msk_caseno.Mask = ""
msk_caseno.Text = ""
msk_caseno.Mask = "#######"
End If
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    txt_dtmno.Enabled = True
    cmb_abo.Enabled = True
    cmb_rh.Enabled = True
    msk_caseno.Enabled = True
    
End If
If str1 = "DIS" Then
    txt_dtmno.Enabled = False
    cmb_abo.Enabled = False
    cmb_rh.Enabled = False
    msk_caseno.Enabled = False
End If
If str1 = "VIEW" Then
    txt_dtmno.Enabled = False
    cmb_abo.Enabled = False
    cmb_rh.Enabled = False
    msk_caseno.Enabled = True
End If
If str1 = "MODIFY" Then
    txt_dtmno.Enabled = True '09/08/10 Shraddha, Uncommented again by Shraddha on 09/09/10
    cmb_abo.Enabled = True
    cmb_rh.Enabled = True
    msk_caseno.Enabled = False
End If
End Sub

Private Sub cmd_back_Click()
On Error GoTo MSG:
btnenabl ("INI")
FLDENABL ("DIS")
fieldclear
Me.Hide
DIS_MAIN.show
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Private Sub cmd_cancel_Click()
On Error GoTo MSG:
btnenabl ("INI")

FLDENABL "DIS"
fieldclear
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Private Sub cmd_delete_Click()

Dim ver As Byte 'to get the choice
On Error GoTo MSG:
btnenabl ("INI")
'If txt_empcode.Text = "" Then
'    MsgBox "Emp. Code cannot be blank"
'Else
'    ver = MsgBox("Do You Want to Delete The Record", vbYesNo)
'    If ver = vbYes Then
'        deldtmuser_qr(0) = Trim(txt_empcode.Text)
'        deldtmuser_qr.Execute
'        MsgBox "Record Deleted Successfully"
'        fieldclear
'    End If
' End If
Exit Sub
MSG:
MsgBox Err.Description, vbCritical, "DTM USER !!"

End Sub





Private Sub cmd_modify_Click()
On Error GoTo MSG:

btnenabl ("NEW")
modsave_flag = True
FLDENABL "MODIFY"
Exit Sub
MSG:
MsgBox Err.Description, vbCritical, "DTM USER !!"

End Sub

Private Sub cmd_new_Click()
On Error GoTo MSG:
btnenabl ("NEW")
newsave_flag = True
NEWFLAG = True
FLDENABL ("ENA")
fieldclear
'txt_empcode.SetFocus
'new_flag = True
If opt_patient.Value = True Then
    msk_caseno.Mask = "??/#####"
    msk_caseno.SetFocus
End If
If opt_staff.Value = True Then
    msk_caseno.Mask = "######"
    msk_caseno.SetFocus
End If
opt_patient.Value = True
Exit Sub
MSG:
MsgBox Err.Description, vbCritical, "DTM USER !!"
End Sub


Private Sub cmd_save_Click()
'On Error GoTo MSG:
If Trim(msk_caseno.Text) = "" Then
  MsgBox "Please Enter Case No.!!", vbInformation, "DTM"
  msk_caseno.SetFocus
  Exit Sub
End If
If Trim(txt_dtmno.Text) = "" Then
  MsgBox "Please Enter DTM No.!!", vbInformation, "DTM"
  txt_dtmno.SetFocus
  Exit Sub
End If
If Trim(cmb_abo.Text) = "" Then
  MsgBox "Please Enter ABO!!", vbInformation, "DTM"
  cmb_abo.SetFocus
  Exit Sub
End If
If Trim(cmb_rh.Text) = "" Then
  MsgBox "Please Enter Rh!!", vbInformation, "DTM"
  cmb_rh.SetFocus
  Exit Sub
End If

btnenabl ("SAVE")
'SHWETA 18-AUG-2010
'TO CHECK DUP DTMNO
 Dim QR_DUPDTMNO As rdoQuery
        Dim RS_DUPDTMNO As rdoResultset
        Dim DUPSTR As String
        DUPSTR = "SELECT * FROM DISLIB.DTMNO WHERE DTMNO = '" & Trim(txt_dtmno.Text) & "'"
        Set QR_DUPDTMNO = rdb.CreateQuery("DUP12", DUPSTR)
        Set RS_DUPDTMNO = QR_DUPDTMNO.OpenResultset(1, 1)
        If RS_DUPDTMNO.EOF Then
        Else
        MsgBox "DTMNO IS ALREADY EXIST"
        'STOP4DTMNO = True
        rdb.RollbackTrans
        Exit Sub
        End If
        QR_DUPDTMNO.Close
        '''''''''''''''''''''''''
'

''''''''''''''''''''''''

If NEWFLAG = True Then
    NEWFLAG = False
        INSDTMNO_QR(0) = msk_caseno.FormattedText
        INSDTMNO_QR(1) = Trim(txt_dtmno.Text)
        INSDTMNO_QR(2) = Trim(cmb_abo.Text) & " " & Trim(cmb_rh.Text)
        INSDTMNO_QR.Execute
        
        '***** Shraddha 06/08/10 *****
        datetime
        UPD_DTMNO_QR(0) = ipadd
        UPD_DTMNO_QR(1) = Format(DATE11, "YYYY-MM-DD")
        UPD_DTMNO_QR(2) = Format(time1, "HH:MM:SS") & UID
        UPD_DTMNO_QR(3) = msk_caseno.FormattedText
        UPD_DTMNO_QR.Execute
        '*****************************
        
        MsgBox "Record Saved Successfully"
    
End If

If modsave_flag Then
    UPDDTMNO_QR(0) = Trim(txt_dtmno.Text)
    UPDDTMNO_QR(1) = Trim(cmb_abo.Text) & " " & Trim(cmb_rh.Text)
    UPDDTMNO_QR(2) = msk_caseno.FormattedText
    UPDDTMNO_QR.Execute
    
    '***** Shraddha 06/08/10 *****
    'Insert into DTMNO FLD2,FLD3,FLD4
    'str = "UPDATE DISLIB.DTMNO SET FLD2=?, FLD3=?, FLD4=? WHERE CASENO=?"
 '   UPDRTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
'UPDRTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
    datetime
    UPD_DTMNO_QR(0) = ipadd
    UPD_DTMNO_QR(1) = Format(DATE11, "YYYY-MM-DD")
    UPD_DTMNO_QR(2) = Format(time1, "HH:MM:SS") & UID
    UPD_DTMNO_QR(3) = msk_caseno.FormattedText
    UPD_DTMNO_QR.Execute
    '*****************************
    
    MsgBox "Record Updated Successfully"
    modsave_flag = False
    'cmd_cancel_Click
End If
Exit Sub
MSG:
    MsgBox Err.Description, vbCritical, "DTM USER !!"
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
FLDENABL "VIEW"
VIEW_FLAG = True
'txt_empcode.Enabled = True
'txt_empcode.SetFocus
'opt_patient.Value = True
If opt_patient.Value = True Then
    msk_caseno.Mask = "??/#####"
    msk_caseno.SetFocus
End If
If opt_staff.Value = True Then
    msk_caseno.Mask = "######"
    msk_caseno.SetFocus
End If
msk_caseno.SetFocus
Exit Sub
MSG:
MsgBox Err.Description, vbCritical, "DTM USER !!"

End Sub

Private Sub Form_Load()
'Main
'fillvalue
btnenabl "INI"
FLDENABL "DIS"
'to view record for entered ccno
If dtm_flag = False Then

'STR = "select * from dislib.dtmuser where empcode= ? "
'Set SELDTMUSER_QR = rdb.CreateQuery("seldtmuser_str", STR)
''to insert a new record
'STR = "insert into dislib.dtmuser(empcode,username,password) values (?,?,?)"
'Set insdtmuser_qr = rdb.CreateQuery("insdtmuser_str", STR)
''to delete selected record
'STR = "delete from dislib.dtmuser where empcode=?"
'Set deldtmuser_qr = rdb.CreateQuery("deldtmuser_str", STR)
''to modify the record
'STR = "update dislib.dtmuser set username=?,password=? where empcode=?"
'Set upddtmuser_qr = rdb.CreateQuery("upddtmuser_str", STR)
'
'STR = "select count(empcode) empcode from dislib.dtmuser where empcode=?"
'Set QR = rdb.CreateQuery("str_dtm", STR)
  
str = "INSERT INTO DISLIB.DTMNO(CASENO,DTMNO,FLD1) VALUES(?,?,?)"
Set INSDTMNO_QR = rdb.CreateQuery("INSDTMNO_STR55", str)

str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR44", str)

'***** Shraddha 06/08/10 *****
str = "SELECT * FROM DISLIB.DTMTAB WHERE CASENO=?"
Set SEL_DTMTAB_QR = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.DTMNO SET FLD2=?, FLD3=?, FLD4=? WHERE CASENO=?"
Set UPD_DTMNO_QR = rdb.CreateQuery("", str)
'*****************************

str = "UPDATE DISLIB.DTMNO SET DTMNO=?,FLD1=? WHERE CASENO=?"
Set UPDDTMNO_QR = rdb.CreateQuery("UPDDTMNO_STRFF", str)


  cmb_abo.AddItem "A"
  cmb_abo.AddItem "B"
  cmb_abo.AddItem "AB"
  cmb_abo.AddItem "O"
  
  cmb_rh.AddItem "Negative"
  cmb_rh.AddItem "Positive"
  cmb_rh.AddItem "Weak D"
dtm_flag = True
End If
End Sub

Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_view.Enabled = False
    cmd_delete.Enabled = False
ElseIf cap = "VIEW" Then
    NEWFLAG = False
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_view.Enabled = False
    cmd_delete.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_view.Enabled = True
    cmd_delete.Enabled = False
    'NEWFLAG = True
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_view.Enabled = True
    cmd_delete.Enabled = False
End If
End Sub

Private Sub datanav()
End Sub




Private Sub msk_caseno_LostFocus()
'On Error GoTo Errhandler
If opt_patient.Value = False And opt_staff.Value = False Then
    Exit Sub
End If
If NEWFLAG = True Then
'  NEWFLAG = False
  If opt_patient.Value = True Then
    CASENO = padzero(5, "C", UCase(msk_caseno.Text), 3)
    msk_caseno.Text = UCase(CASENO)
  Else
    CASENO = msk_caseno.FormattedText
  End If
  If SELDTMNO_FLAG = False Then
    SELDTMNO_QR(0) = CASENO
    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    SELDTMNO_FLAG = True
  Else
    SELDTMNO_QR(0) = CASENO
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
  End If
  If Not SELDTMNO_RS.EOF Then
    If Len(Trim(SELDTMNO_RS(0))) > 0 Then
        MsgBox "This Case No. is Already Present in the DataBase", vbInformation, "DTM"
        'Msk_caseno.SetFocus
        '***** Shraddha 09/08/10 *****
        cmd_modify.Enabled = False
        cmd_save.Enabled = False
        txt_dtmno.Enabled = False
        cmb_abo.Enabled = False
        cmb_rh.Enabled = False

        '*****************************
    Else
        MsgBox "Blood Group not done"
    End If
  End If
End If
If VIEW_FLAG = True Then
  VIEW_FLAG = False
  
  
  
  If opt_patient.Value = True Then
    CASENO = padzero(5, "C", UCase(msk_caseno.Text), 3)
    msk_caseno.Text = UCase(CASENO)
  Else
    CASENO = msk_caseno.FormattedText
  End If
  
    '***** Shraddha 06/08/10 *****
    SEL_DTMTAB_QR(0) = CASENO
    Set SEL_DTMTAB_RS = SEL_DTMTAB_QR.OpenResultset(1, 3)
    
    'cmd_modify
    If Not SEL_DTMTAB_RS.EOF Then
        cmd_modify.Enabled = True
    Else
        MsgBox "Blood Group not done"
        'cmd_modify.Enabled = False 'Commented on 09/08/10
    End If
    '*****************************
  
  If SELDTMNO_FLAG = False Then
    SELDTMNO_QR(0) = CASENO
    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    SELDTMNO_FLAG = True
  Else
    SELDTMNO_QR(0) = CASENO
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
  End If
  If Not SELDTMNO_RS.EOF Then
    Dim i
    txt_dtmno.Text = Trim(SELDTMNO_RS("DTMNO"))
    
    i = InStr(1, Trim(SELDTMNO_RS("FLD1")), " ")
    cmb_abo.Text = Mid(Trim(SELDTMNO_RS("FLD1")), 1, i)
    
    'If condition added by Shraddha on 06/08/10 *****
    If Trim(SELDTMNO_RS("FLD1")) <> "" Then
        cmb_rh.Text = Mid(Trim(SELDTMNO_RS("FLD1")), i)
    End If
  End If
End If



Exit Sub
Errhandler:
MsgBox Err.Description, vbCritical, "DTMNO"
Exit Sub
End Sub
