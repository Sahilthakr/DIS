VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_microbatch 
   Caption         =   "Batch report for Microbiology"
   ClientHeight    =   3345
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form2"
   ScaleHeight     =   3345
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\disreportsec\microbiologyindex.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   6975
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
         Left            =   5160
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
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
         Left            =   720
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmb_ref 
         Caption         =   "&Refresh"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Msk_todt 
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Msk_fromdt 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LBL_TODT 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   705
      End
      Begin VB.Label LBL_FROMDT 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   885
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Micro Biology Batch Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frm_microbatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selma105qr As rdoQuery
Dim selma105rs As rdoResultset
Dim ma105flag As Boolean

' select user for enterby from dislnkusr
Dim sel_dislnkusr_qr As rdoQuery
Dim sel_dislnkusr_rs As rdoResultset
Dim sel_dislnkusr_flg As Boolean

Private Sub cmb_patcatg_GotFocus()
'cmb_patcatg.Text = "ALL"
End Sub

Private Sub cmb_patstat_LostFocus()
'If Trim(cmb_patstat.Text) = "IP" Then
'           If selma105flag = False Then
'                Set selma105rs = selma105qr.OpenResultset(1, 3)
'                selma105flag = True
'            Else
'                selma105rs.MoreResults
'                selma105rs.Requery
'            End If
'            If selma105rs.RowCount > 0 Then
'                cmb_ward.Enabled = True
'                cmb_ward.clear
'                cmb_ward.AddItem "ALL"
'                While Not selma105rs.EOF
'
'                    cmb_ward.AddItem Trim(selma105rs("ward_name"))
'                    selma105rs.MoveNext
'                Wend
'             End If
'
'Else
'   cmb_ward.Enabled = False
'  cmb_ward.clear
'End If
End Sub

Private Sub cmb_pattype_Change()
'If Trim(cmb_pattype.Text) = "G" Then
'    cmb_patcatg.clear
'    cmb_patcatg.AddItem "ALL"
'    cmb_patcatg.AddItem "C"
'    cmb_patcatg.AddItem "NC"
'End If
'If Trim(cmb_pattype.Text) = "P" Then
'    cmb_patcatg.clear
'    cmb_patcatg.AddItem "ALL"
'    cmb_patcatg.AddItem "A"
'    cmb_patcatg.AddItem "B"
'    cmb_patcatg.AddItem "D"
'    cmb_patcatg.AddItem "F"
'End If
'If Trim(cmb_pattype.Text) = "R" Then
'    cmb_patcatg.clear
'    cmb_patcatg.AddItem "ALL"
'    cmb_patcatg.AddItem "C"
'    cmb_patcatg.AddItem "F"
'End If
End Sub

Private Sub cmb_ref_Click()
msk_fromdt.Text = ""
msk_todt.Text = ""
'cmb_patcatg.Text = ""
'cmb_patstat.Text = ""
'cmb_pattype.Text = ""
'cmb_ward.Text = ""
End Sub

Private Sub cmd_back_Click()
msk_fromdt.Text = ""
msk_todt.Text = ""
'cmb_patcatg.Text = ""
'cmb_patstat.Text = ""
'cmb_pattype.Text = ""
'cmb_ward.Text = ""
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_print_Click()

''Dim var_date_f As String, var_date_t As String
''Dim StrType As String, StrCat As String, StrStat As String
''Dim StrWard As String, StrEnterby As String
''If Trim(cmb_patcatg.Text) = "" Or Trim(cmb_patcatg.Text) = "ALL" Then
''    StrCat = ""
'''    MsgBox "Patient Category is Empty. Select the Category", vbInformation, "Patient Cateogy"
'''    cmb_patcatg.SetFocus
'''    Exit Sub
''Else
''  StrCat = "PAT_CATG='" & UCase(Trim(cmb_patcatg.Text)) & "'  AND "
''End If
''
''If Trim(cmb_patstat.Text) = "" Or Trim(cmb_patstat.Text) = "ALL" Then
''    StrStat = ""
'''    MsgBox "Patient Status is Empty. Select the Patient Status", vbInformation, "Patient Status"
'''    cmb_patstat.SetFocus
'''    Exit Sub
''Else
''    StrStat = "PATNSTAT='" & UCase(Trim(cmb_patstat.Text)) & "'  AND "
''End If
''
''If Trim(cmb_enterby.Text) <> "" Or Trim(cmb_enterby.Text) <> "ALL" Then
''      StrEnterby = " ENTERBY LIKE '" & (cmb_enterby.Text) & "%'  AND  "
''End If
''
''If Trim(cmb_patstat.Text) = "IP" Then
''    If Trim(cmb_ward.Text) = "" Or Trim(cmb_ward.Text) = "ALL" Then
'''        MsgBox "Ward is Empty. Select the Ward", vbInformation, "Ward Name"
'''        cmb_ward.SetFocus
'''        Exit Sub
''    StrWard = ""
''    Else
''    StrWard = "REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN WHERE WARD_NAME='" & UCase(Trim(cmb_ward.Text)) & "' )  AND "
''    End If
''End If

var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

'justin
'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F') AND PREPORTDT BETWEEN ? AND ?  "
''STR = "SELECT * FROM RISLIB.RREQUISITN WHERE " & StrWard & StrType & StrStat & StrCat & StrEnterby & _
''"MODALITYCD='BM' AND (REQNSTAT='T' OR REQNSTAT='F') AND  PREPORTDT BETWEEN ? AND ?"


CrystalReport1.ReportFileName = "C:\disreportsec\microbiologyindexOTHERS.rpt"
CrystalReport1.Formulas(0) = "FROMDT='" & Trim(msk_fromdt.FormattedText) & "'"
CrystalReport1.Formulas(2) = "TODT='" & Trim(msk_todt.FormattedText) & "'"
CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date_f & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date_t & ") AND ({RREQUISITN.PATNSTAT}='IP' and {RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='GZZ'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Destination = crptToWindow
CrystalReport1.Action = 1

CrystalReport1.ReportFileName = "C:\disreportsec\microbiologyindexPVTOPD.rpt"
CrystalReport1.Formulas(0) = "FROMDT='" & Trim(msk_fromdt.FormattedText) & "'"
CrystalReport1.Formulas(2) = "TODT='" & Trim(msk_todt.FormattedText) & "'"
CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date_f & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date_t & ") AND {RREQUISITN.PATNSTAT}='OP' and ({RREQUISITN.PAT_CATG}<>'C' or {RREQUISITN.PAT_CATG}<>'NC') and ({RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='GZZ'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Destination = crptToWindow
CrystalReport1.Action = 1

CrystalReport1.ReportFileName = "C:\disreportsec\microbiologyindexGENOPD.rpt"
CrystalReport1.Formulas(0) = "FROMDT='" & Trim(msk_fromdt.FormattedText) & "'"
CrystalReport1.Formulas(2) = "TODT='" & Trim(msk_todt.FormattedText) & "'"
'CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date_f & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date_t & ") AND (RREQUISITN.PATNSTAT}='OP' and ({RREQUISITN.PAT_CATG}='C' or {RREQUISITN.PAT_CATG}='NC') and ({RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='GZZ'"
CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date_f & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date_t & ") AND {RREQUISITN.PATNSTAT}='OP' and ({RREQUISITN.PAT_CATG}='C' or {RREQUISITN.PAT_CATG}='NC') and ({RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='GZZ'"

CrystalReport1.WindowState = crptMaximized
CrystalReport1.Destination = crptToWindow
CrystalReport1.Action = 1

CrystalReport1.ReportFileName = "C:\disreportsec\microbiologyindexST.rpt"
CrystalReport1.Formulas(0) = "FROMDT='" & Trim(msk_fromdt.FormattedText) & "'"
CrystalReport1.Formulas(2) = "TODT='" & Trim(msk_todt.FormattedText) & "'"
CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date_f & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date_t & ") AND ({RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='GZZ'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Destination = crptToWindow
CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()

DATE11 = Date
''STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=8 AND DEPT_CODE='FEC' ORDER BY USERNAME"
''Set sel_dislnkusr_qr = rdb.CreateQuery("sel_dislnkusr_qr", STR)
''
''If sel_dislnkusr_flg = False Then
''    Set sel_dislnkusr_rs = sel_dislnkusr_qr.OpenResultset(1, 3)
''    sel_dislnkusr_flg = True
''Else
''    sel_dislnkusr_rs.MoreResults
''    sel_dislnkusr_rs.Requery
''End If
''
''  If sel_dislnkusr_rs.RowCount > 0 Then
''    Cmb_Enterby.clear
''    sel_dislnkusr_rs.MoveFirst
''     While Not sel_dislnkusr_rs.EOF
''        Cmb_Enterby.AddItem Trim(sel_dislnkusr_rs("USERNAME"))
''        sel_dislnkusr_rs.MoveNext
''     Wend
''  End If

'cmb_patcatg.clear
'cmb_patstat.clear
'
'    cmb_patstat.AddItem "ALL"
'    cmb_patstat.AddItem "IP"
'    cmb_patstat.AddItem "OP"

str = " select ward_name from patoolslib.ma105"
Set selma105qr = rdb.CreateQuery("selma105qrbatretre", str)
End Sub

Private Sub Msk_fromdt_GotFocus()
msk_fromdt.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub Msk_todt_GotFocus()
msk_todt.Text = Format(DATE11, "dd/mm/yyyy")
End Sub
