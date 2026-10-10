VERSION 5.00
Begin VB.Form Dis_security 
   Caption         =   "DIS Security"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12645
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
   ScaleHeight     =   7620
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   11775
      Begin VB.ComboBox cmb_addendum 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3120
         Width           =   2535
      End
      Begin VB.ComboBox cmb_final 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   2535
      End
      Begin VB.ComboBox cmb_prov 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox cmb_gross 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ComboBox cmb_enter 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txt_modality 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6480
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txt_ccno 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl_addendum 
         Caption         =   "Addendum By"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lbl_final 
         Caption         =   "Final By"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lbl_prov 
         Caption         =   "Provisional By"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lbl_gross 
         Caption         =   "Gross By"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lbl_enter 
         Caption         =   "Enter By"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lbl_modality 
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbl_name 
         Caption         =   "CC Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_ccno 
         Caption         =   "CC No."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
   End
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
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   11775
      Begin VB.CommandButton cmd_delete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6960
         MouseIcon       =   "Dis_security.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Dis_security.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1920
         MouseIcon       =   "Dis_security.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3600
         MouseIcon       =   "Dis_security.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   5280
         MouseIcon       =   "Dis_security.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   8640
         MouseIcon       =   "Dis_security.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10320
         MouseIcon       =   "Dis_security.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   1215
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label lbl_cap 
      AutoSize        =   -1  'True
      Caption         =   "DIS - SECURITY"
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
      Left            =   4920
      TabIndex        =   16
      Top             =   840
      Width           =   2190
   End
End
Attribute VB_Name = "Dis_security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'for view mode
Dim sel_dissecview_qr As rdoQuery
Dim sel_dissecview_rs As rdoResultset
Dim sel_dissecview_flag As Boolean
'to insert data in table
Dim ins_dissecnew_qr As rdoQuery

Dim ins_dissecnew_rs As rdoResultset
Dim ins_dissecnew_flag As Boolean
'to delete record
Dim del_dissecdel_qr As rdoQuery
'to modify the record
Dim upd_dissecmod_qr As rdoQuery
Dim upd_dissecmod_flag As Boolean
Dim modsave_flag As Boolean
Dim newsave_flag As Boolean
Dim new_flag As Boolean


Private Sub fieldclear()
cmb_addendum.ListIndex = -1
cmb_enter.ListIndex = -1
cmb_final.ListIndex = -1
cmb_gross.ListIndex = -1
cmb_prov.ListIndex = -1
txt_ccno.Text = ""
txt_name.Text = ""

End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    txt_ccno.Enabled = True
    txt_name.Enabled = True
    txt_modality.Enabled = True
    cmb_enter.Enabled = True
    cmb_final.Enabled = True
    cmb_prov.Enabled = True
    cmb_addendum.Enabled = True
    cmb_gross.Enabled = True
End If
If str1 = "DIS" Then
    txt_ccno.Enabled = False
    txt_name.Enabled = False
    txt_modality.Enabled = False
    cmb_enter.Enabled = False
    cmb_final.Enabled = False
    cmb_prov.Enabled = False
    cmb_addendum.Enabled = False
    cmb_gross.Enabled = False
End If
If str1 = "VIEW" Then
    txt_ccno.Enabled = True
    txt_name.Enabled = False
    txt_modality.Enabled = False
    cmb_enter.Enabled = False
    cmb_final.Enabled = False
    cmb_prov.Enabled = False
    cmb_addendum.Enabled = False
    cmb_gross.Enabled = False
End If
If str1 = "MODIFY" Then
    txt_ccno.Enabled = False
    txt_name.Enabled = False
    txt_modality.Enabled = True
    cmb_enter.Enabled = True
    cmb_final.Enabled = True
    cmb_prov.Enabled = True
    cmb_addendum.Enabled = True
    cmb_gross.Enabled = True
End If
End Sub

Private Sub cmd_back_Click()
On Error GoTo MSG:
btnenabl ("INI")
FLDENABL ("DIS")
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
If txt_ccno.Text = "" Then
    MsgBox "CC No. cannot be blank"
Else
    ver = MsgBox("Do You Want to Delete The Record", vbYesNo)
    If ver = vbYes Then
        del_dissecdel_qr(0) = txt_ccno.Text
        del_dissecdel_qr.Execute
        MsgBox "Record Deleted Successfully"
    End If
 End If
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Private Sub cmd_modify_Click()
On Error GoTo MSG:
btnenabl ("NEW")
modsave_flag = True
FLDENABL "MODIFY"
fillvalue
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Private Sub cmd_new_Click()
On Error GoTo MSG:
btnenabl ("NEW")
newsave_flag = True
new_flag = False
FLDENABL ("ENA")

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_save_Click()
On Error GoTo MSG:
btnenabl ("SAVE")
If newsave_flag Then
    If txt_ccno = "" Then
        MsgBox "CCNo. cannot be blank"
    Else
        ins_dissecnew_qr(0) = txt_ccno.Text
        ins_dissecnew_qr(1) = txt_name.Text
        ins_dissecnew_qr(2) = cmb_enter.Text
        ins_dissecnew_qr(3) = cmb_prov.Text
        ins_dissecnew_qr(4) = cmb_gross.Text
        ins_dissecnew_qr(5) = cmb_final.Text
        ins_dissecnew_qr(6) = cmb_addendum.Text
        ins_dissecnew_qr(7) = txt_modality.Text
        ins_dissecnew_qr.Execute
        newsave_flag = False
    End If
End If

If modsave_flag Then
    upd_dissecmod_qr(0) = txt_name.Text
    upd_dissecmod_qr(1) = cmb_enter.Text
    upd_dissecmod_qr(2) = cmb_prov.Text
    upd_dissecmod_qr(3) = cmb_gross.Text
    upd_dissecmod_qr(4) = cmb_final.Text
    upd_dissecmod_qr(5) = cmb_addendum.Text
    upd_dissecmod_qr(6) = txt_ccno.Text
    upd_dissecmod_qr(7) = txt_modality.Text
    upd_dissecmod_qr.Execute
    MsgBox "Record Updated Successfully"
    
    modsave_flag = False
    cmd_cancel_Click
End If
Exit Sub
MSG:
If Err.Number = 40002 Then
    MsgBox "CC No. Already Exists"
    txt_ccno.Text = ""
Else
    MsgBox Err.Description
End If

End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
new_flag = True
FLDENABL "VIEW"
Exit Sub
MSG:
MsgBox Err.Description

End Sub
Private Sub fillvalue()
'enter values to all combo
cmb_enter.clear
cmb_enter.AddItem "YES"
cmb_enter.AddItem "NO"

cmb_prov.clear
cmb_prov.AddItem "YES"
cmb_prov.AddItem "NO"

cmb_final.clear
cmb_final.AddItem "YES"
cmb_final.AddItem "NO"

cmb_gross.clear
cmb_gross.AddItem "YES"
cmb_gross.AddItem "NO"

cmb_addendum.clear
cmb_addendum.AddItem "YES"
cmb_addendum.AddItem "NO"

End Sub
Private Sub Form_Load()
'Main
fillvalue
btnenabl "INI"
FLDENABL "DIS"
'to view record for entered ccno
str = "select ccname,enterby,provby,grossby,finalby,addendumby,modality from dislib.dissec where ccno= ? "
Set sel_dissecview_qr = rdb.CreateQuery("sel1_deptqrcg", str)
'to insert a new record
str = "insert into dislib.dissec(ccno,ccname,enterby,provby,grossby,finalby,addendumby,modality) values (?,?,?,?,?,?,?,?)"
Set ins_dissecnew_qr = rdb.CreateQuery("ins_dissecnew", str)
'to delete selected record
str = "delete from dislib.dissec where ccno= ? "
Set del_dissecdel_qr = rdb.CreateQuery("del_dissecdel", str)
'to modify the record
str = "update dislib.dissec set ccname=?,enterby=?,provby=?,grossby=?,finalby=?,addendumby=?,modality=? where ccno=?"
Set upd_dissecmod_qr = rdb.CreateQuery("upd_dissecmod", str)

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

Private Sub txt_ccno_LostFocus()

If new_flag Then
    If Not sel_dissecview_flag Then
        sel_dissecview_qr(0) = txt_ccno.Text
        Set sel_dissecview_rs = sel_dissecview_qr.OpenResultset(1, 3)
        sel_dissecview_flag = True
    Else
        sel_dissecview_qr(0) = txt_ccno.Text
        'Set sel_dissecview_rs = sel_dissecview_qr.OpenResultset(1, 3)
        sel_dissecview_rs.MoreResults
        sel_dissecview_rs.Requery
    End If
    If sel_dissecview_rs.RowCount > 0 Then
        txt_name.Text = sel_dissecview_rs("ccname")
        'cmb_enter.clear
        cmb_enter.AddItem sel_dissecview_rs("enterby")
        cmb_enter.Text = sel_dissecview_rs("enterby")
        'cmb_final.clear
        cmb_final.AddItem sel_dissecview_rs("finalby")
        cmb_final.Text = sel_dissecview_rs("finalby")
        'cmb_prov.clear
        cmb_prov.AddItem sel_dissecview_rs("provby")
        cmb_prov.Text = sel_dissecview_rs("provby")
        'cmb_addendum.clear
        cmb_addendum.AddItem sel_dissecview_rs("addendumby")
        cmb_addendum.Text = sel_dissecview_rs("addendumby")
        'cmb_gross.clear
        cmb_gross.AddItem sel_dissecview_rs("grossby")
        cmb_gross.Text = sel_dissecview_rs("grossby")
        txt_modality.Text = sel_dissecview_rs("modality")
    Else
        MsgBox "No Record"
        txt_ccno.Text = ""
    End If
 End If
End Sub


