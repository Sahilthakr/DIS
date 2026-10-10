VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form tmbatch 
   Caption         =   "Tumour Marker Batch Printing"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
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
   ScaleHeight     =   9180
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   1080
      TabIndex        =   38
      Top             =   5280
      Width           =   9135
      Begin VB.TextBox txt_tolbno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4680
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txt_frmlabno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Msk_fromdt 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
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
      Begin MSMask.MaskEdBox Msk_todt 
         Height          =   375
         Left            =   4680
         TabIndex        =   1
         Top             =   120
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
      Begin VB.Label LBL_FROMDT 
         AutoSize        =   -1  'True
         Caption         =   "To Lab No."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LBL_FROMDT 
         AutoSize        =   -1  'True
         Caption         =   "From Lab No."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   40
         Top             =   240
         Width           =   705
      End
      Begin VB.Label LBL_FROMDT 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   1080
      TabIndex        =   22
      Top             =   1560
      Width           =   9135
      Begin VB.ComboBox Cmb_Enterby 
         Height          =   315
         ItemData        =   "tmbatch.frx":0000
         Left            =   1920
         List            =   "tmbatch.frx":0002
         TabIndex        =   36
         Top             =   2160
         Width           =   3975
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "Back"
         Height          =   375
         Left            =   5520
         TabIndex        =   35
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox cmb_ward 
         Height          =   315
         Left            =   6000
         TabIndex        =   28
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox cmb_patstat 
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cmb_patcatg 
         Height          =   315
         Left            =   6000
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cmb_pattype 
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton cmb_ref 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Entered by"
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Ward"
         Height          =   255
         Left            =   5400
         TabIndex        =   34
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Patient Status"
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Patient Category"
         Height          =   255
         Left            =   4440
         TabIndex        =   32
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Patient Type"
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LBL_TODT 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   30
         Top             =   480
         Width           =   705
      End
      Begin VB.Label LBL_FROMDT 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   29
         Top             =   480
         Width           =   885
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   2895
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   9135
      Begin VB.CommandButton cmd_ref 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton CMD_PRINT1 
         Caption         =   "&Print"
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txt_tolabno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txt_fromlabno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox MSK_TODT1 
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MSK_FROMDT1 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.Label LBL_TODT1 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4800
         TabIndex        =   17
         Top             =   360
         Width           =   705
      End
      Begin VB.Label LBL_FROMDT1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lbl_labno2 
         AutoSize        =   -1  'True
         Caption         =   "To Lab No."
         Height          =   195
         Left            =   4560
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_labno1 
         AutoSize        =   -1  'True
         Caption         =   "From Lab No."
         Height          =   195
         Left            =   720
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITALAND RESEARCH CENTRE"
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
      Left            =   480
      TabIndex        =   21
      Top             =   240
      Width           =   9975
   End
   Begin VB.Label LBL_CAP 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TUMOUR MARKER  REPORTS"
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
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label LBL_PRINTED 
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   8280
      TabIndex        =   19
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   1320
      TabIndex        =   18
      Top             =   7320
      Width           =   6735
   End
End
Attribute VB_Name = "tmbatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selma105qr As rdoQuery
Dim selma105rs As rdoResultset
Dim ma105flag As Boolean
Dim str As String
Dim selrrequisitnqr As rdoQuery
Dim selrrequisitnrs As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim SELRREQNOQR As rdoQuery
Dim SELRREQNORS As rdoResultset
Dim SELRREQNOFALG As Boolean
Dim SELRREQNOQRALL As rdoQuery
Dim SELRREQNORSALL As rdoResultset
Dim SELRREQNOFLAGALL As Boolean

Dim selrreqnoipallqr As rdoQuery
Dim selrreqnoipallrs As rdoResultset
Dim selrreqnoipallflag As Boolean

Dim SELTMBTCH_QR As rdoQuery
Dim SELTMBTCH_RS As rdoResultset
Dim SELTMBTCHFLAG As Boolean
Dim UPDRTIMEPB_QR As rdoQuery

Dim seldmprocedurqr As rdoQuery
Dim seldmprocedurrs As rdoResultset
Dim seldmprocedurflag As Boolean

'************************
Dim selpreqnoqrEnt As rdoQuery
Dim selpreqnorsEnt As rdoResultset
Dim selpreqnoflagEnt As Boolean
Dim enter_by1 As String
'************************

Dim selpreqnoqr As rdoQuery
Dim selpreqnors As rdoResultset
Dim selpreqnoflag As Boolean

Dim seltmpreqnoqr As rdoQuery
Dim seltmpreqnors As rdoResultset
Dim seltmpreqnoflag As Boolean

Dim SELTMREQNOQR As rdoQuery
Dim SELTMREQNORS As rdoResultset
Dim SELTMREQNOFLAG As Boolean

Dim SELREQNOQR As rdoQuery
Dim SELREQNORS As rdoResultset
Dim SELREQNOFLAG As Boolean

Dim SELMULPROQR As rdoQuery
Dim SELMULPRORS As rdoResultset
Dim SELMULPROFLAG As Boolean
Dim TMFYESFALG As Boolean
Dim TMFFFALG As Boolean

Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset

' select user for enterby from dislnkusr
Dim sel_dislnkusr_qr As rdoQuery
Dim sel_dislnkusr_rs As rdoResultset
Dim sel_dislnkusr_flg As Boolean
' select user for enterby from dislnkusr




Private Sub cmb_enterby_Click()
    If Trim(cmb_enterby.Text) <> "" Then
        enter_by1 = (cmb_enterby.Text) & "%"
    End If
End Sub

Private Sub cmb_patcatg_GotFocus()
cmb_patcatg.Text = "ALL"
End Sub

Private Sub cmb_patstat_GotFocus()
cmb_patstat.Text = "ALL"
End Sub

Private Sub cmb_patstat_LostFocus()
If Trim(cmb_patstat.Text) = "IP" Then
           If selma105flag = False Then
                Set selma105rs = selma105qr.OpenResultset(1, 3)
                selma105flag = True
            Else
                selma105rs.MoreResults
                selma105rs.Requery
            End If
            If selma105rs.RowCount > 0 Then
                cmb_ward.Enabled = True
                cmb_ward.clear
                cmb_ward.AddItem "ALL"
                While Not selma105rs.EOF
                    
                    cmb_ward.AddItem Trim(selma105rs("ward_name"))
                    selma105rs.MoveNext
                Wend
             End If
                
Else
   cmb_ward.Enabled = False
 cmb_ward.clear
End If
End Sub

Private Sub cmb_pattype_Click()
If Trim(cmb_pattype.Text) = "G" Then
    cmb_patcatg.clear
    cmb_patcatg.AddItem "ALL"
    cmb_patcatg.AddItem "C"
    cmb_patcatg.AddItem "NC"
End If
If Trim(cmb_pattype.Text) = "P" Then
    cmb_patcatg.clear
    cmb_patcatg.AddItem "ALL"
    cmb_patcatg.AddItem "A"
    cmb_patcatg.AddItem "B"
    cmb_patcatg.AddItem "D"
    cmb_patcatg.AddItem "F"
End If
If Trim(cmb_pattype.Text) = "R" Then
    cmb_patcatg.clear
    cmb_patcatg.AddItem "ALL"
    cmb_patcatg.AddItem "C"
    cmb_patcatg.AddItem "F"
End If
End Sub

Private Sub cmb_ref_Click()
msk_fromdt.Text = ""
msk_todt.Text = ""
cmb_patcatg.Text = ""
cmb_patstat.Text = ""
cmb_pattype.Text = ""
cmb_ward.Text = ""
Label1.Caption = ""
End Sub

Private Sub cmd_back_Click()
msk_fromdt.Text = ""
msk_todt.Text = ""
cmb_patcatg.Text = ""
cmb_patstat.Text = ""
cmb_pattype.Text = ""
cmb_ward.Text = ""
Label1.Caption = ""
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_print_Click()
'On Error GoTo prtmsg:
Dim pcount As Integer
Dim str As String
Dim str1 As String
Dim VAR_REQNNO As String
Dim var_yr_f As Integer, var_yr_t As Integer
Dim var_mon_f As Integer, var_mon_t As Integer
Dim var_day_f As Integer, var_day_t As Integer
Dim var_date_f As String, var_date_t As String
Dim StrType As String, StrCat As String, StrStat As String
Dim StrWard As String, StrEnterby As String

If Trim(cmb_pattype.Text) = "" Or Trim(cmb_pattype.Text) = "ALL" Then
    StrType = ""
'    MsgBox "Patient Type is Empty. Select the type 'G' or 'P'", vbInformation, "Patient Type"
'    cmb_pattype.SetFocus
'    Exit Sub
Else

  StrType = "CASENO IN(SELECT CASE_NO FROM PATOOLSLIB.TA101 WHERE PAT_REG_TY ='" & Trim(cmb_pattype.Text) & "' )  AND  "
End If

If Trim(cmb_patcatg.Text) = "" Or Trim(cmb_patcatg.Text) = "ALL" Then
    StrCat = ""
'    MsgBox "Patient Category is Empty. Select the Category", vbInformation, "Patient Cateogy"
'    cmb_patcatg.SetFocus
'    Exit Sub
Else
  StrCat = "PAT_CATG='" & UCase(Trim(cmb_patcatg.Text)) & "'  AND "
End If

If Trim(cmb_patstat.Text) = "" Or Trim(cmb_patstat.Text) = "ALL" Then
    StrStat = ""
'    MsgBox "Patient Status is Empty. Select the Patient Status", vbInformation, "Patient Status"
'    cmb_patstat.SetFocus
'    Exit Sub
Else
    StrStat = "PATNSTAT='" & UCase(Trim(cmb_patstat.Text)) & "'  AND "
End If

If Trim(cmb_enterby.Text) <> "" Or Trim(cmb_enterby.Text) <> "ALL" Then
      StrEnterby = " ENTERBY LIKE '" & (cmb_enterby.Text) & "%'  AND  "
End If

If Trim(cmb_patstat.Text) = "IP" Then
    If Trim(cmb_ward.Text) = "" Or Trim(cmb_ward.Text) = "ALL" Then
'        MsgBox "Ward is Empty. Select the Ward", vbInformation, "Ward Name"
'        cmb_ward.SetFocus
'        Exit Sub
    StrWard = ""
    Else
    StrWard = "REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN WHERE WARD_NAME='" & UCase(Trim(cmb_ward.Text)) & "' )  AND "
    End If
End If

var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

'justin
'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F') AND PREPORTDT BETWEEN ? AND ?  "
str = "SELECT * FROM RISLIB.RREQUISITN WHERE " & StrWard & StrType & StrStat & StrCat & StrEnterby & _
"MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F') AND  PREPORTDT BETWEEN ? AND ?"


Set selrrequisitnqr = rdb.CreateQuery("selpreqnoqrbrTM1", str)

    selrrequisitnqr(0) = msk_fromdt.FormattedText
    selrrequisitnqr(1) = msk_todt.FormattedText
    Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)

If selrrequisitnrs.RowCount > 0 Then
  
     
  While Not selrrequisitnrs.EOF
                
                If SELMULPROFLAG = False Then
                    SELMULPROQR(0) = Trim(selrrequisitnrs("REQNNO"))
                    Set SELMULPRORS = SELMULPROQR.OpenResultset(1, 3)
                    SELMULPROFLAG = True
                Else
                    SELMULPROQR(0) = Trim(selrrequisitnrs("REQNNO"))
                    SELMULPRORS.MoreResults
                    SELMULPRORS.Requery
                End If
                
                If SELMULPRORS.RowCount > 0 Then         '3
                    While Not SELMULPRORS.EOF
                        If SELMULPRORS("PRO_CODE") = "F816" Or SELMULPRORS("PRO_CODE") = "F828" Then
                            TMFYESFALG = True
                        Else
                            TMFFFALG = True
                        End If
                        SELMULPRORS.MoveNext
                    Wend
                End If
                
                
                
                
                
                If TMFFFALG = True Then                 '4
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTMFORBATCH.rpt"
                    'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM.rpt"
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selrrequisitnrs("REQNNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selrrequisitnrs("REQNNO")) & "'"
                    
                    
                    
                    If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                        str1 = ""
                        str1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    End If
                    
                    'CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'                    CrystalReport1.WindowState = crptMaximized
                    'CrystalReport1.Destination = crptToWindow

                    CrystalReport1.Destination = crptToPrinter
                    CrystalReport1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFFFALG = False
                End If                                        '4
                
                If TMFYESFALG = True Then                       '5
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1FORBATCH.RPT"
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selrrequisitnrs("REQNNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selrrequisitnrs("REQNNO")) & "'"
                    
                    
                    
                    
                    
'                    If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                        CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
'                        STR1 = ""
'                        STR1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
'                        CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'                    Else
'                        CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
'                        CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'                    End If
                    
                    'CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
                    'CrystalReport1.WindowState = crptMaximized
                    'CrystalReport1.Destination = crptToWindow 'SHWETA 31 JAN 2012
                    CrystalReport1.Destination = crptToPrinter
                    'CRYSTALREPORT1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFYESFALG = False
               End If                                         '5
            
            
        CC = CC + 1
        LBL_PRINTED.Caption = CC
        

        selrrequisitnrs.MoveNext
    Wend
        selrrequisitnqr.Close
        MsgBox "Printing is completed"
        lbl_cap.Caption = ""
    Else
       selrrequisitnqr.Close
        MsgBox "No Records are there for Printing"
        lbl_cap.Caption = ""
    End If                                                      '2
   
 
                
                
                
                
'justin


'''If Trim(cmb_patstat.Text) = "OP" Then
'''
'''     If SELRREQUISITNFLAG = False Then
'''        selrrequisitnqr(0) = Trim(cmb_patstat.Text)
'''        selrrequisitnqr(1) = Trim(cmb_patcatg.Text)
'''        selrrequisitnqr(2) = msk_fromdt.FormattedText
'''        selrrequisitnqr(3) = msk_todt.FormattedText
'''        Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
'''        SELRREQUISITNFLAG = True
'''     Else
'''        selrrequisitnqr(0) = Trim(cmb_patstat.Text)
'''        selrrequisitnqr(1) = Trim(cmb_patcatg.Text)
'''        selrrequisitnqr(2) = msk_fromdt.FormattedText
'''        selrrequisitnqr(3) = msk_todt.FormattedText
'''        selrrequisitnrs.MoreResults
'''        selrrequisitnrs.Requery
'''     End If
'''     Label1.Caption = "No. of Reports in this Combination is " & selrrequisitnrs.RowCount
'''     While Not selrrequisitnrs.EOF
'''        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
'''        STR = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
'''
'''      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
'''       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(selrrequisitnrs("REQNNO")) & "' "
'''
'''         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selrrequisitnrs("REQNNO")) & "'"
'''
'''        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            STR1 = ""
'''            STR1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''         Else
'''
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''        End If
'''
'''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'''
'''        If seldmprocedurflag = False Then
'''            seldmprocedurqr(0) = selrrequisitnrs("REQNNO")
'''            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
'''            seldmprocedurflag = True
'''        Else
'''             seldmprocedurqr(0) = selrrequisitnrs("REQNNO")
'''             seldmprocedurrs.MoreResults
'''             seldmprocedurrs.Requery
'''        End If
'''        If seldmprocedurrs.RowCount > 0 Then
'''            pcount = 0
'''            While Not seldmprocedurrs.EOF
'''                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
'''                    pcount = pcount + 1
'''                End If
'''                seldmprocedurrs.MoveNext
'''            Wend
'''        End If
'''
'''        If pcount = 0 Then
'''            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
'''        Else
'''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'''        End If
'''
'''        CrystalReport1.WindowState = crptMaximized
'''        CrystalReport1.Destination = crptToPrinter
'''        'CRYSTALREPORT1.WindowState = crptMaximized
'''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'''        CrystalReport1.Action = 1
'''        selrrequisitnrs.MoveNext
'''        Wend
'''
'''End If
'''If Trim(cmb_patstat.Text) = "IP" And Trim(cmb_ward.Text) <> "ALL" Then
'''
'''    If SELRREQNOFALG = False Then
'''        SELRREQNOQR(0) = Trim(cmb_ward.Text)
'''        SELRREQNOQR(1) = Trim(cmb_patstat.Text)
'''        SELRREQNOQR(2) = Trim(cmb_patcatg.Text)
'''        SELRREQNOQR(3) = msk_fromdt.FormattedText
'''        SELRREQNOQR(4) = msk_todt.FormattedText
'''        Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
'''        SELRREQNOFALG = True
'''     Else
'''        SELRREQNOQR(0) = Trim(cmb_ward.Text)
'''        SELRREQNOQR(1) = Trim(cmb_patstat.Text)
'''        SELRREQNOQR(2) = Trim(cmb_patcatg.Text)
'''        SELRREQNOQR(3) = msk_fromdt.FormattedText
'''        SELRREQNOQR(4) = msk_todt.FormattedText
'''        SELRREQNORS.MoreResults
'''        SELRREQNORS.Requery
'''     End If
'''     Label1.Caption = "No. of Reports in this Combination is " & SELRREQNORS.RowCount
'''     While Not SELRREQNORS.EOF
'''        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
'''        STR = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
'''
'''
'''      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
'''       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORS("REQNNO")) & "' "
'''
'''         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(SELRREQNORS("REQNNO")) & "'"
'''
'''         If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            STR1 = ""
'''            STR1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''         Else
'''
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''        End If
'''
'''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'''        '******
'''        If seldmprocedurflag = False Then
'''            seldmprocedurqr(0) = SELRREQNORS("REQNNO")
'''            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
'''            seldmprocedurflag = True
'''        Else
'''             seldmprocedurqr(0) = SELRREQNORS("REQNNO")
'''             seldmprocedurrs.MoreResults
'''             seldmprocedurrs.Requery
'''        End If
'''        If seldmprocedurrs.RowCount > 0 Then
'''            pcount = 0
'''            While Not seldmprocedurrs.EOF
'''                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
'''                    pcount = pcount + 1
'''                End If
'''                seldmprocedurrs.MoveNext
'''            Wend
'''        End If
'''
'''        If pcount = 0 Then
'''            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
'''        Else
'''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'''        End If
'''
'''
'''        '******
'''        CrystalReport1.WindowState = crptMaximized
'''        CrystalReport1.Destination = crptToPrinter
'''        'CRYSTALREPORT1.WindowState = crptMaximized
'''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'''        CrystalReport1.Action = 1
'''
'''
'''        SELRREQNORS.MoveNext
'''        Wend
'''End If
'''If cmb_patstat = "IP" And cmb_ward.Text = "ALL" Then
'''    If selrreqnoipallflag = False Then
'''
'''        selrreqnoipallqr(0) = Trim(cmb_patstat.Text)
'''        selrreqnoipallqr(1) = Trim(cmb_patcatg.Text)
'''        selrreqnoipallqr(2) = msk_fromdt.FormattedText
'''        selrreqnoipallqr(3) = msk_todt.FormattedText
'''        Set selrreqnoipallrs = selrreqnoipallqr.OpenResultset(1, 3)
'''        selrreqnoipallflag = True
'''     Else
'''
'''        selrreqnoipallqr(0) = Trim(cmb_patstat.Text)
'''        selrreqnoipallqr(1) = Trim(cmb_patcatg.Text)
'''        selrreqnoipallqr(2) = msk_fromdt.FormattedText
'''        selrreqnoipallqr(3) = msk_todt.FormattedText
'''        selrreqnoipallrs.MoreResults
'''        selrreqnoipallrs.Requery
'''     End If
'''     Label1.Caption = "No. of Reports in this Combination is " & selrreqnoipallrs.RowCount
'''     While Not selrreqnoipallrs.EOF
'''        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
'''
'''        STR = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
'''
'''
'''      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
'''       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORS("REQNNO")) & "' "
'''
'''         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selrreqnoipallrs("REQNNO")) & "'"
'''
'''         If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            STR1 = ""
'''            STR1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''         Else
'''
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''        End If
'''
'''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'''         '********
'''           If seldmprocedurflag = False Then
'''            seldmprocedurqr(0) = selrreqnoipallrs("REQNNO")
'''            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
'''            seldmprocedurflag = True
'''        Else
'''             seldmprocedurqr(0) = selrreqnoipallrs("REQNNO")
'''             seldmprocedurrs.MoreResults
'''             seldmprocedurrs.Requery
'''        End If
'''        If seldmprocedurrs.RowCount > 0 Then
'''            pcount = 0
'''            While Not seldmprocedurrs.EOF
'''                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
'''                    pcount = pcount + 1
'''                End If
'''                seldmprocedurrs.MoveNext
'''            Wend
'''        End If
'''
'''        If pcount = 0 Then
'''            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
'''        Else
'''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'''        End If
'''
'''
'''
'''         '********
'''
'''        CrystalReport1.WindowState = crptMaximized
'''        CrystalReport1.Destination = crptToPrinter
'''        'CRYSTALREPORT1.WindowState = crptMaximized
'''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'''        CrystalReport1.Action = 1
'''
'''
'''        SELRREQNORS.MoveNext
'''        Wend
'''End If
'''
'''If Trim(cmb_patcatg.Text) = "ALL" And Trim(cmb_patstat.Text) = "ALL" Then
'''
'''    If SELRREQNOFLAGALL = False Then
'''
'''        SELRREQNOQRALL(0) = msk_fromdt.FormattedText
'''        SELRREQNOQRALL(1) = msk_todt.FormattedText
'''        Set SELRREQNORSALL = SELRREQNOQRALL.OpenResultset(1, 3)
'''        SELRREQNOFLAGALL = True
'''     Else
'''
'''        SELRREQNOQRALL(0) = msk_fromdt.FormattedText
'''        SELRREQNOQRALL(1) = msk_todt.FormattedText
'''        SELRREQNORSALL.MoreResults
'''        SELRREQNORSALL.Requery
'''     End If
'''     Label1.Caption = "No. of Reports in this Combination is " & SELRREQNORSALL.RowCount
'''     While Not SELRREQNORSALL.EOF
'''        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
'''
'''        STR = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
'''
'''
'''      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
'''       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORSALL("REQNNO")) & "' "
'''
'''         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(SELRREQNORSALL("REQNNO")) & "'"
'''
'''         If Trim(SELRREQNORSALL("REQNSTAT")) = "T" Then
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            STR1 = ""
'''            STR1 = "Entered By " & Trim(SELRREQNORSALL("ENTERBY"))
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''         Else
'''
'''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''        End If
'''
'''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'''        '******
'''        If seldmprocedurflag = False Then
'''            seldmprocedurqr(0) = SELRREQNORSALL("REQNNO")
'''            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
'''            seldmprocedurflag = True
'''        Else
'''             seldmprocedurqr(0) = SELRREQNORSALL("REQNNO")
'''             seldmprocedurrs.MoreResults
'''             seldmprocedurrs.Requery
'''        End If
'''        If seldmprocedurrs.RowCount > 0 Then
'''            pcount = 0
'''            While Not seldmprocedurrs.EOF
'''                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
'''                    pcount = pcount + 1
'''                End If
'''                seldmprocedurrs.MoveNext
'''            Wend
'''        End If
'''
'''        If pcount = 0 Then
'''            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
'''        Else
'''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'''        End If
'''
'''
'''
'''        '******
'''        CrystalReport1.WindowState = crptMaximized
'''        CrystalReport1.Destination = crptToPrinter
'''        'CRYSTALREPORT1.WindowState = crptMaximized
'''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'''        CrystalReport1.Action = 1
'''
'''
'''        SELRREQNORSALL.MoveNext
'''        Wend
'''
'''
'''End If
'''MsgBox "Printing is Completed"
'''lbl_cap.Caption = ""
'''Exit Sub
'''
'prtmsg: MsgBox Error(Err)
'''
'''
End Sub

Private Sub CMD_PRINT1_Click()
Dim CC As Integer
CC = 0
On Error GoTo MSG:
TMFYESFALG = False
TMFFFALG = False
If Trim(MSK_FROMDT1.Text) = "" Then
    MsgBox "Enter the From Labno."
    Exit Sub
End If
If Trim(MSK_TODT1.Text) = "" Then
    MsgBox "Enter the to Lab No."
    Exit Sub
End If
    selpreqnoqr.QueryTimeout = 300
    selpreqnoqrEnt.QueryTimeout = 300
If Trim(cmb_enterby.Text) <> "" Then                   '1
        
            selpreqnoqrEnt(0) = Format(MSK_FROMDT1.FormattedText, "YYYY-MM-DD")
            selpreqnoqrEnt(1) = Format(MSK_TODT1.FormattedText, "YYYY-MM-DD")
            selpreqnoqrEnt(2) = Trim(enter_by1)
        If selpreqnoflagEnt = False Then
            Set selpreqnorsEnt = selpreqnoqrEnt.OpenResultset(1, 3)
            selpreqnoflagEnt = True
        Else
            selpreqnorsEnt.MoreResults
            selpreqnorsEnt.Requery
        End If
        
   '*******************************************************
        If selpreqnorsEnt.RowCount > 0 Then               '2
            Label1.Caption = "No. of Reports in this Combination is " & selpreqnorsEnt.RowCount
            LBL_PRINTED.Caption = CC
            str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
            While Not selpreqnorsEnt.EOF
                If SELMULPROFLAG = False Then
                    SELMULPROQR(0) = Trim(selpreqnorsEnt("REQNNO"))
                    Set SELMULPRORS = SELMULPROQR.OpenResultset(1, 3)
                    SELMULPROFLAG = True
                Else
                    SELMULPROQR(0) = Trim(selpreqnorsEnt("REQNNO"))
                    SELMULPRORS.MoreResults
                    SELMULPRORS.Requery
                End If
                If SELMULPRORS.RowCount > 0 Then         '3
                    While Not SELMULPRORS.EOF
                        If SELMULPRORS("PRO_CODE") = "F816" Or SELMULPRORS("PRO_CODE") = "F828" Then
                            TMFYESFALG = True
                        Else
                            TMFFFALG = True
                        End If
                    SELMULPRORS.MoveNext
                    Wend
                
                If TMFFFALG = True Then                 '4
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM.rpt"
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selpreqnorsEnt("REQNNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqnorsEnt("REQNNO")) & "'"
                    
                    If Trim(selpreqnorsEnt("REQNSTAT")) = "T" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                        str1 = ""
                        str1 = "Entered By " & Trim(selpreqnorsEnt("ENTERBY"))
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    End If
                    
                    CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
                    CrystalReport1.WindowState = crptMaximized
                    'CrystalReport1.Destination = crptToWindow
                    CrystalReport1.Destination = crptToPrinter
                    'CRYSTALREPORT1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFFFALG = False
                End If                                        '4
                
                If TMFYESFALG = True Then                       '5
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1.RPT"
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selpreqnorsEnt("REQNNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqnorsEnt("REQNNO")) & "'"
                    If Trim(selpreqnorsEnt("REQNSTAT")) = "T" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
                        str1 = ""
                        str1 = "Entered By " & Trim(selpreqnorsEnt("ENTERBY"))
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    End If
                    
                    CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
                    CrystalReport1.WindowState = crptMaximized
                    'CrystalReport1.Destination = crptToPrinter
                    CrystalReport1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFYESFALG = False
                 End If                                         '5
            End If                                              '3
        CC = CC + 1
        LBL_PRINTED.Caption = CC
        

        selpreqnorsEnt.MoveNext
    Wend
        MsgBox "Printing is completed"
        lbl_cap.Caption = ""
    Else
        MsgBox "No Records are there for Printing"
        lbl_cap.Caption = ""
    End If                                                      '2
    
   '*******************************************************
Else                                                            '1
    
    If selpreqnoflag = False Then
        selpreqnoqr(0) = Trim(MSK_FROMDT1.FormattedText)
        selpreqnoqr(1) = Trim(MSK_TODT1.FormattedText)
        Set selpreqnors = selpreqnoqr.OpenResultset(1, 3)
        selpreqnoflag = True
    Else
        selpreqnoqr(0) = Trim(MSK_FROMDT1.FormattedText)
        selpreqnoqr(1) = Trim(MSK_TODT1.FormattedText)
        selpreqnors.MoreResults
        selpreqnors.Requery
    End If
'----------------------------------------------------------
If selpreqnors.RowCount > 0 Then                                '6
        
     Label1.Caption = "No. of Reports in this Combination is " & selpreqnors.RowCount
     LBL_PRINTED.Caption = CC
     str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
     While Not selpreqnors.EOF
        
            If SELMULPROFLAG = False Then
                SELMULPROQR(0) = Trim(selpreqnors("REQNNO"))
                Set SELMULPRORS = SELMULPROQR.OpenResultset(1, 3)
                SELMULPROFLAG = True
            Else
               SELMULPROQR(0) = Trim(selpreqnors("REQNNO"))
               SELMULPRORS.MoreResults
               SELMULPRORS.Requery
            End If
            If SELMULPRORS.RowCount > 0 Then                        '7
                While Not SELMULPRORS.EOF
                    If SELMULPRORS("PRO_CODE") = "F816" Or SELMULPRORS("PRO_CODE") = "F828" Then
                        TMFYESFALG = True
                    Else
                        TMFFFALG = True
                    End If
                    SELMULPRORS.MoveNext
                Wend
                
                If TMFFFALG = True Then                             '8
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM.rpt"
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selpreqnors("REQNNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqnors("REQNNO")) & "'"
                    If Trim(selpreqnors("REQNSTAT")) = "T" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                        str1 = ""
                        str1 = "Entered By " & Trim(selpreqnors("ENTERBY"))
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                     Else
            
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    End If
                    CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
                    CrystalReport1.WindowState = crptMaximized
                    'CrystalReport1.Destination = crptToWindow
                    CrystalReport1.Destination = crptToPrinter
                    'CRYSTALREPORT1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFFFALG = False
                End If                                              '8
                If TMFYESFALG = True Then                           '9
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1.RPT"
                    
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selpreqnors("REQNNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqnors("REQNNO")) & "'"
                    If Trim(selpreqnors("REQNSTAT")) = "T" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
                        str1 = ""
                        str1 = "Entered By " & Trim(selpreqnors("ENTERBY"))
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                     Else
            
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    End If
                    CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
                    CrystalReport1.WindowState = crptMaximized
                    CrystalReport1.Destination = crptToPrinter
                    'CRYSTALREPORT1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFYESFALG = False
                    
                End If                                      '9
            End If                                          '7
        CC = CC + 1
        LBL_PRINTED.Caption = CC
        

        selpreqnors.MoveNext
    Wend
        MsgBox "Printing is completed"
        lbl_cap.Caption = ""
    Else
        MsgBox "No Records are there for Printing"
        lbl_cap.Caption = ""
    End If                                                  '6
End If                                                      '1

    Exit Sub
MSG:
    MsgBox Err.Description
    
End Sub

Private Sub cmd_ref_Click()
txt_frmLabNo.Text = ""
txt_tolbno.Text = ""
MSK_FROMDT1.Text = ""
MSK_TODT1.Text = ""
cmb_enterby.Text = ""
End Sub

Private Sub Command1_Click()
Dim CC As Integer
CC = 0
'On Error GoTo MSG:
If Trim(txt_frmLabNo.Text) = "" Then
    MsgBox "Enter the Lab no. From"
    Exit Sub
End If

If Trim(txt_tolbno.Text) = "" Then
    MsgBox "Enter the Lab No. upto"
    Exit Sub
End If

'seltmpreqnoqr.QueryTimeout = 300
'***** Commented and added by Shraddha on 01/02/10 *****
'seltmpreqnoqr.QueryTimeout = 1000
'Commented and added by Shraddha on 25/02/10
seltmpreqnoqr.QueryTimeout = 5000
If seltmpreqnoflag = False Then
    seltmpreqnoqr(0) = Trim(txt_frmLabNo.Text)
    seltmpreqnoqr(1) = Trim(txt_tolbno.Text)
    Set seltmpreqnors = seltmpreqnoqr.OpenResultset(1, 3)
    seltmpreqnoflag = True
Else
    seltmpreqnoqr(0) = Trim(txt_frmLabNo.Text)
    seltmpreqnoqr(1) = Trim(txt_tolbno.Text)
    seltmpreqnors.MoreResults
    seltmpreqnors.Requery
End If
'RTIMELOG UPDATION 14-11-2006

'If SELTMBTCHFLAG = False Then
'    SELTMBTCH_QR(0) = Trim(txt_frmlabno.Text)
'    SELTMBTCH_QR(1) = Trim(txt_tolabno.Text)
'    Set SELTMBTCH_RS = SELTMBTCH_QR.OpenResultset(1, 3)
'    SELTMBTCHFLAG = True
'Else
'    SELTMBTCH_QR(0) = Trim(txt_frmlabno.Text)
'    SELTMBTCH_QR(1) = Trim(txt_tolabno.Text)
'    SELTMBTCH_RS.MoreResults
'    SELTMBTCH_RS.Requery
'End If

While seltmpreqnors.EOF = False
    UPDRTIMEPB_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    UPDRTIMEPB_QR(1) = Time 'txt_time.Text
    UPDRTIMEPB_QR(2) = seltmpreqnors("REQNO")
    UPDRTIMEPB_QR.Execute
    seltmpreqnors.MoveNext
Wend

For i = 0 To 13
    CrystalReport1.Formulas(i) = ""
Next
If seltmpreqnors.RowCount > 0 Then
        
    Label1.Caption = "No. of Reports in this Combination is " & seltmpreqnors.RowCount
    LBL_PRINTED.Caption = CC
    seltmpreqnors.MoveFirst
    While Not seltmpreqnors.EOF
        CC = CC + 1
        LBL_PRINTED.Caption = CC

        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")


      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
      'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
      'CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(seltmpreqnors("REQNO")) & "'"
            
            
                If SELMULPROFLAG = False Then
                    SELMULPROQR(0) = Trim(seltmpreqnors("REQNO"))
                    Set SELMULPRORS = SELMULPROQR.OpenResultset(1, 3)
                    SELMULPROFLAG = True
                Else
                    SELMULPROQR(0) = Trim(seltmpreqnors("REQNO"))
                    SELMULPRORS.MoreResults
                    SELMULPRORS.Requery
                End If
                
                If SELMULPRORS.RowCount > 0 Then         '3
                    While Not SELMULPRORS.EOF
                        If SELMULPRORS("PRO_CODE") = "F816" Or SELMULPRORS("PRO_CODE") = "F828" Then
                            TMFYESFALG = True
                        Else
                            TMFFFALG = True
                        End If
                    SELMULPRORS.MoveNext
                    Wend
                End If
                
                
                If SELTMREQNOFLAG = False Then
                    SELTMREQNOQR(0) = seltmpreqnors("REQNO")
                    Set SELTMREQNORS = SELTMREQNOQR.OpenResultset(1, 3)
                    SELTMREQNOFLAG = True
                Else
                    SELTMREQNOQR(0) = seltmpreqnors("REQNO")
                    SELTMREQNORS.MoreResults
                    SELTMREQNORS.Requery
                End If
                
                If TMFFFALG = True Then
                '4
                '***** Shraddha 02/07/10 for staff report in batch printing *****
                If Right(txt_frmLabNo.Text, 3) >= "5101" And Right(txt_frmLabNo, 3) <= "5200" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTMS.rpt"
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
                    
                    
                    CrystalReport1.Formulas(0) = "REQNO = '" & Trim(seltmpreqnors("REQNO")) & "'"
                    'FOLLOWING CONDITION ADDED BY SHWETA 31-01-2012 AS TOLD BY KALPANA MAM
                    'REASON DATA IS NOT GETTING PRINTED
                ElseIf (Right(txt_frmLabNo.Text, 4) >= "1901" And Right(txt_frmLabNo, 4) <= "2000") Or (Right(txt_frmLabNo.Text, 4) >= "201" And Right(txt_frmLabNo, 4) <= "299") Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1FORBATCH.rpt"
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
                       
                      '''''TILL HERE SHWETA 31-01-2012
                CrystalReport1.Formulas(0) = "REQNO = '" & Trim(seltmpreqnors("REQNO")) & "'"
                Else
                '****************************************************************
                    'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTMFORBATCH.rpt"
                    'Commented and added by Shraddha on 17/09/10
                    CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISTMFORBATCH.rpt"
                    CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(seltmpreqnors("REQNO")) & "'"
                End If
                    'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM.rpt"
                    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
                    'CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
                    
                   ' CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
                   ' CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(seltmpreqnors("REQNO")) & "'"
                    
                    
                    If SELTMREQNORS.RowCount > 0 Then
                    
                        If Trim(SELTMREQNORS("REQNSTAT")) = "T" Then
                            CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                            str1 = ""
                            str1 = "Entered By " & Trim(SELTMREQNORS("ENTERBY"))
                            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        Else
                            CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        End If
                        
                    End If
                    
                    
'                    CrystalReport1.WindowState = crptMaximized
                    'CrystalReport1.Destination = crptToWindow
                    CrystalReport1.Destination = crptToPrinter
                    CrystalReport1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFFFALG = False
                End If                                        '4
                
                If TMFYESFALG = True Then                       '5
                    
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1FORBATCH.RPT"
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(seltmpreqnors("REQNO")) & "' "
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(seltmpreqnors("REQNO")) & "'"
                    
                    If SELTMREQNORS.RowCount > 0 Then
                    
                        If Trim(SELTMREQNORS("REQNSTAT")) = "T" Then
                            CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
                            str1 = ""
                            str1 = "Entered By " & Trim(SELTMREQNORS("ENTERBY"))
                            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        Else
                            CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
                            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        End If
                        
                    End If
                           
                    
                    'CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
                    CrystalReport1.WindowState = crptMaximized
                    CrystalReport1.Destination = crptToWindow
                    'CrystalReport1.Destination = crptToPrinter
                    'CRYSTALREPORT1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                    TMFYESFALG = False
               End If                                         '5
    
     seltmpreqnors.MoveNext
Wend
End If

    MsgBox "Printing is Completed"
    lbl_cap.Caption = ""
Exit Sub
'MSG:
'    MsgBox Err.Description

End Sub

Private Sub Command2_Click()
msk_fromdt.Text = ""
msk_todt.Text = ""
Label1.Caption = ""
Me.Hide
DIS_MAIN.show
End Sub

Private Sub Command3_Click()
msk_fromdt.Text = ""
msk_todt.Text = ""
txt_frmLabNo.Text = ""
txt_tolbno.Text = ""
Label1.Caption = ""
End Sub

Private Sub Form_Load()
Dim str As String
' select user for enterby from dislnkusr
str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=8 AND DEPT_CODE='FEC' ORDER BY USERNAME"
Set sel_dislnkusr_qr = rdb.CreateQuery("sel_dislnkusr_qr", str)
' select user for enterby from dislnkusr

'STR = "SELECT * FROM PATOOLSLIB/TA101 WHERE CASE_NO=?"
'Set SELTA101_QR = rdb.CreateQuery("selta101_qr", STR)
'SELTA101_QR(0) = ""
'Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 1)

'********************** Shraddha Pandit 18/03/09 ********************************
'STR = "SELECT REQNO,PATHNO FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO"
'Set seltmpreqnoqr = rdb.CreateQuery("seltmpreqnoqrbr", STR)

'str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY REFDOC,PATHNO"

str = "SELECT REQNO, PATHNO, REFDOC FROM DISLIB.BCTMLF WHERE MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F') AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO"

Set seltmpreqnoqr = rdb.CreateQuery("seltmpreqnoqrbr", str)

'********************************************************************************

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELTMREQNOQR = rdb.CreateQuery("SELTMREQNOQRCC", str)

str = "SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F') AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO"
Set SELTMBTCH_QR = rdb.CreateQuery("SELTMBTCH_QR", str)

str = "UPDATE RISLIB.RTIMELOG SET PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
Set UPDRTIMEPB_QR = rdb.CreateQuery("UPDRTIMEPB_QR", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set selrrequisitnqr = rdb.CreateQuery("SELRREQUISITNQRTMBCR", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN WHERE WARD_NAME=?) AND  MODALITYCD='TM' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQRBCTM1 ", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND (REQNSTAT= 'T' OR REQNSTAT='F')AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQRALL = rdb.CreateQuery("SELRREQNOQRALLBCTM", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN)  AND  MODALITYCD='TM' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?    order by dislib.drequisitn.pathno "
Set selrreqnoipallqr = rdb.CreateQuery("SELRREQNOQRALLipTM", str)

str = "SELECT pvalue FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set seldmprocedurqr = rdb.CreateQuery("seldmprocedurqrBPTM", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE (MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F') AND (PREPORTDT BETWEEN ? AND ?) AND (ENTERBY like ?))"
Set selpreqnoqrEnt = rdb.CreateQuery("selpreqnoqrEnt", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='TM' AND (REQNSTAT='T' OR REQNSTAT='F') AND PREPORTDT BETWEEN ? AND ?  "
Set selpreqnoqr = rdb.CreateQuery("selpreqnoqrbrTM", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNOQR = rdb.CreateQuery("SELREQNOQRCTMC", str)

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SELMULPROQR = rdb.CreateQuery("SELMULPROQRTM", str)

msk_fromdt.Text = ""
msk_todt.Text = ""
cmb_patcatg.clear
cmb_patstat.clear
cmb_pattype.clear
cmb_ward.clear
datetime

    cmb_pattype.AddItem "G"
    cmb_pattype.AddItem "P"
    cmb_pattype.AddItem "R"
    cmb_patstat.AddItem "ALL"
    cmb_patstat.AddItem "IP"
    cmb_patstat.AddItem "OP"
str = " select ward_name from patoolslib.ma105"
Set selma105qr = rdb.CreateQuery("selma105qrbatrep6", str)

' ************************
If sel_dislnkusr_flg = False Then
    Set sel_dislnkusr_rs = sel_dislnkusr_qr.OpenResultset(1, 3)
    sel_dislnkusr_flg = True
Else
    sel_dislnkusr_rs.MoreResults
    sel_dislnkusr_rs.Requery
End If
        
  If sel_dislnkusr_rs.RowCount > 0 Then
    cmb_enterby.clear
    sel_dislnkusr_rs.MoveFirst
     While Not sel_dislnkusr_rs.EOF
        cmb_enterby.AddItem Trim(sel_dislnkusr_rs("USERNAME"))
        sel_dislnkusr_rs.MoveNext
     Wend
  End If
' ************************
End Sub



Private Sub Msk_fromdt_GotFocus()
msk_fromdt.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub MSK_FROMDT1_GotFocus()
MSK_FROMDT1.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub Msk_todt_GotFocus()
msk_todt.Text = Format(DATE11, "dd/mm/yyyy")
End Sub



Private Sub MSK_TODT1_GotFocus()
MSK_TODT1.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub txt_frmlabno_LostFocus()
Dim NN As String
If Trim(txt_frmLabNo.Text) = "" Then
    Exit Sub
End If
NN = Right(msk_fromdt.FormattedText, 2) & Mid(msk_fromdt.FormattedText, 4, 2) & Left(msk_fromdt.FormattedText, 2)
txt_frmLabNo.Text = NN & txt_frmLabNo.Text

End Sub

Private Sub txt_tolbno_LostFocus()
Dim NN As String
If Trim(txt_tolbno.Text) = "" Then
    Exit Sub
End If
NN = Right(msk_todt.FormattedText, 2) & Mid(msk_todt.FormattedText, 4, 2) & Left(msk_todt.FormattedText, 2)
txt_tolbno.Text = NN & txt_tolbno.Text

End Sub
