VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form BATCH_REPORT 
   Caption         =   "Batch Reports"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2760
      TabIndex        =   18
      Top             =   4320
      Width           =   9135
      Begin VB.CommandButton cmd_ref 
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
         Left            =   3840
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton CMD_PRINT1 
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
         Left            =   2280
         TabIndex        =   29
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt_tolabno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   28
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txt_fromlabno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Top             =   720
         Width           =   1335
      End
      Begin MSMask.MaskEdBox MSK_TODT1 
         Height          =   375
         Left            =   5640
         TabIndex        =   25
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox MSK_FROMDT1 
         Height          =   375
         Left            =   1920
         TabIndex        =   23
         Top             =   240
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
      Begin VB.Label LBL_TODT1 
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
         Left            =   4800
         TabIndex        =   26
         Top             =   360
         Width           =   705
      End
      Begin VB.Label LBL_FROMDT1 
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
         Index           =   1
         Left            =   960
         TabIndex        =   24
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lbl_labno2 
         AutoSize        =   -1  'True
         Caption         =   "To Lab No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl_labno1 
         AutoSize        =   -1  'True
         Caption         =   "From Lab No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   720
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Width           =   9135
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
         Left            =   3840
         TabIndex        =   17
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "Back"
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
         Left            =   5760
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cmb_pattype 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmb_patcatg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmb_patstat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cmb_ward 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   5
         Top             =   1680
         Width           =   2775
      End
      Begin MSMask.MaskEdBox Msk_todt 
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Left            =   1920
         TabIndex        =   0
         Top             =   240
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
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   885
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
         Left            =   5160
         TabIndex        =   14
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Patient Type"
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
         Left            =   720
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Patient Category"
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
         Left            =   4440
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Patient Status"
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
         Left            =   600
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Ward"
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
         Left            =   5400
         TabIndex        =   10
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE,MULLANPUR"
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
      Height          =   495
      Left            =   240
      TabIndex        =   30
      Top             =   120
      Width           =   13095
   End
   Begin VB.Label LBL_PRINTED 
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   7560
      TabIndex        =   22
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   7920
      Width           =   5055
   End
   Begin VB.Label LBL_CAP 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "BIOCHEMISTRY REPORTS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "BATCH_REPORT"
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

Dim seldmprocedurqr As rdoQuery
Dim seldmprocedurrs As rdoResultset
Dim seldmprocedurflag As Boolean
Dim selpreqnoqr As rdoQuery
Dim selpreqnors As rdoResultset
Dim selpreqnoflag As Boolean

Dim SELREQNOQR As rdoQuery
Dim SELREQNORS As rdoResultset
Dim SELREQNOFLAG As Boolean

'****** Added by Shraddha Pandit on 06/05/2009 ****
Dim sel_patstat_qr As rdoQuery
Dim sel_patstat_rs As rdoResultset

Dim sel_patstatOPIP_qr As rdoQuery
Dim sel_patstatOPIP_RS As rdoResultset
'***************************************************


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
On Error GoTo prtmsg:
Dim pcount As Integer
Dim str As String
Dim str1 As String
Dim VAR_REQNNO As String
Dim var_yr_f As Integer, var_yr_t As Integer
Dim var_mon_f As Integer, var_mon_t As Integer
Dim var_day_f As Integer, var_day_t As Integer
Dim var_date_f As String, var_date_t As String
If Trim(cmb_pattype.Text) = "" Then
    MsgBox "Patient Type is Empty. Select the type 'G' or 'P'", vbInformation, "Patient Type"
    cmb_pattype.SetFocus
    Exit Sub
End If
If Trim(cmb_patcatg.Text) = "" Then
    MsgBox "Patient Category is Empty. Select the Category", vbInformation, "Patient Cateogy"
    cmb_patcatg.SetFocus
    Exit Sub
End If
If Trim(cmb_patstat.Text) = "" Then
    MsgBox "Patient Status is Empty. Select the Patient Status", vbInformation, "Patient Status"
    cmb_patstat.SetFocus
    Exit Sub
End If
If Trim(cmb_patstat.Text) = "IP" Then
    If Trim(cmb_ward.Text) = "" Then
        MsgBox "Ward is Empty. Select the Ward", vbInformation, "Ward Name"
        cmb_ward.SetFocus
        Exit Sub
    End If
End If

var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

If Trim(cmb_patstat.Text) = "OP" Then

     If SELRREQUISITNFLAG = False Then
        selrrequisitnqr(0) = Trim(cmb_patstat.Text)
        selrrequisitnqr(1) = Trim(cmb_patcatg.Text)
        selrrequisitnqr(2) = msk_fromdt.FormattedText
        selrrequisitnqr(3) = msk_todt.FormattedText
        Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
     Else
        selrrequisitnqr(0) = Trim(cmb_patstat.Text)
        selrrequisitnqr(1) = Trim(cmb_patcatg.Text)
        selrrequisitnqr(2) = msk_fromdt.FormattedText
        selrrequisitnqr(3) = msk_todt.FormattedText
        selrrequisitnrs.MoreResults
        selrrequisitnrs.Requery
     End If
     Label1.Caption = "No. of Reports in this Combination is " & selrrequisitnrs.RowCount
     While Not selrrequisitnrs.EOF
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
      'T F condition added on 03/09/10
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(selrrequisitnrs("REQNNO")) & "' AND ({RREQUISITN.REQNSTAT}= 'T' OR {RREQUISITN.REQNSTAT}= 'F')"
        
         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selrrequisitnrs("REQNNO")) & "'"
        
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
  
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        If seldmprocedurflag = False Then
            seldmprocedurqr(0) = selrrequisitnrs("REQNNO")
            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
            seldmprocedurflag = True
        Else
             seldmprocedurqr(0) = selrrequisitnrs("REQNNO")
             seldmprocedurrs.MoreResults
             seldmprocedurrs.Requery
        End If
        If seldmprocedurrs.RowCount > 0 Then
            pcount = 0
            While Not seldmprocedurrs.EOF
                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                    pcount = pcount + 1
                End If
                seldmprocedurrs.MoveNext
            Wend
        End If
        If pcount = 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        End If
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
        'CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        selrrequisitnrs.MoveNext
        Wend
End If
If Trim(cmb_patstat.Text) = "IP" And Trim(cmb_ward.Text) <> "ALL" Then
    
    If SELRREQNOFALG = False Then
        SELRREQNOQR(0) = Trim(cmb_ward.Text)
        SELRREQNOQR(1) = Trim(cmb_patstat.Text)
        SELRREQNOQR(2) = Trim(cmb_patcatg.Text)
        SELRREQNOQR(3) = msk_fromdt.FormattedText
        SELRREQNOQR(4) = msk_todt.FormattedText
        Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
        SELRREQNOFALG = True
     Else
        SELRREQNOQR(0) = Trim(cmb_ward.Text)
        SELRREQNOQR(1) = Trim(cmb_patstat.Text)
        SELRREQNOQR(2) = Trim(cmb_patcatg.Text)
        SELRREQNOQR(3) = msk_fromdt.FormattedText
        SELRREQNOQR(4) = msk_todt.FormattedText
        SELRREQNORS.MoreResults
        SELRREQNORS.Requery
     End If
     Label1.Caption = "No. of Reports in this Combination is " & SELRREQNORS.RowCount
     While Not SELRREQNORS.EOF
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
      'T F condition added on 03/09/10
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORS("REQNNO")) & "' AND ({RREQUISITN.REQNSTAT}= 'T' OR {RREQUISITN.REQNSTAT}= 'F')"
        
         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(SELRREQNORS("REQNNO")) & "'"
        
         If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
  
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        '******
        If seldmprocedurflag = False Then
            seldmprocedurqr(0) = SELRREQNORS("REQNNO")
            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
            seldmprocedurflag = True
        Else
             seldmprocedurqr(0) = SELRREQNORS("REQNNO")
             seldmprocedurrs.MoreResults
             seldmprocedurrs.Requery
        End If
        If seldmprocedurrs.RowCount > 0 Then
            pcount = 0
            While Not seldmprocedurrs.EOF
                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                    pcount = pcount + 1
                End If
                seldmprocedurrs.MoveNext
            Wend
        End If
        If pcount = 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        End If
        
        
        '******
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        

        SELRREQNORS.MoveNext
        Wend
End If

If cmb_patstat = "IP" And cmb_ward.Text = "ALL" Then
    If selrreqnoipallflag = False Then
        
        selrreqnoipallqr(0) = Trim(cmb_patstat.Text)
        selrreqnoipallqr(1) = Trim(cmb_patcatg.Text)
        selrreqnoipallqr(2) = msk_fromdt.FormattedText
        selrreqnoipallqr(3) = msk_todt.FormattedText
        Set selrreqnoipallrs = selrreqnoipallqr.OpenResultset(1, 3)
        selrreqnoipallflag = True
     Else
        
        selrreqnoipallqr(0) = Trim(cmb_patstat.Text)
        selrreqnoipallqr(1) = Trim(cmb_patcatg.Text)
        selrreqnoipallqr(2) = msk_fromdt.FormattedText
        selrreqnoipallqr(3) = msk_todt.FormattedText
        selrreqnoipallrs.MoreResults
        selrreqnoipallrs.Requery
     End If
     Label1.Caption = "No. of Reports in this Combination is " & selrreqnoipallrs.RowCount
     While Not selrreqnoipallrs.EOF
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
      'T F condition added on 03/09/10
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORS("REQNNO")) & "' AND ({RREQUISITN.REQNSTAT}= 'T' OR {RREQUISITN.REQNSTAT}= 'F')"
         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selrreqnoipallrs("REQNNO")) & "'"
         If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
         '********
        If seldmprocedurflag = False Then
            seldmprocedurqr(0) = selrreqnoipallrs("REQNNO")
            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
            seldmprocedurflag = True
        Else
             seldmprocedurqr(0) = selrreqnoipallrs("REQNNO")
             seldmprocedurrs.MoreResults
             seldmprocedurrs.Requery
        End If
        If seldmprocedurrs.RowCount > 0 Then
            pcount = 0
            While Not seldmprocedurrs.EOF
                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                    pcount = pcount + 1
                End If
                seldmprocedurrs.MoveNext
            Wend
        End If
        If pcount = 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        End If
         
         
         
         '********
        
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        SELRREQNORS.MoveNext
        Wend
End If
If Trim(cmb_patcatg.Text) = "ALL" And Trim(cmb_patstat.Text) = "ALL" Then
    If SELRREQNOFLAGALL = False Then
        SELRREQNOQRALL(0) = msk_fromdt.FormattedText
        SELRREQNOQRALL(1) = msk_todt.FormattedText
        Set SELRREQNORSALL = SELRREQNOQRALL.OpenResultset(1, 3)
        SELRREQNOFLAGALL = True
     Else
        SELRREQNOQRALL(0) = msk_fromdt.FormattedText
        SELRREQNOQRALL(1) = msk_todt.FormattedText
        SELRREQNORSALL.MoreResults
        SELRREQNORSALL.Requery
     End If
     Label1.Caption = "No. of Reports in this Combination is " & SELRREQNORSALL.RowCount
     While Not SELRREQNORSALL.EOF
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
      'T F condition added on 03/09/10
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORSALL("REQNNO")) & "' AND ({RREQUISITN.REQNSTAT}= 'T' OR {RREQUISITN.REQNSTAT}= 'F')"
        
         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(SELRREQNORSALL("REQNNO")) & "'"
        
         If Trim(SELRREQNORSALL("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORSALL("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
  
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        '******
        If seldmprocedurflag = False Then
            seldmprocedurqr(0) = SELRREQNORSALL("REQNNO")
            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
            seldmprocedurflag = True
        Else
             seldmprocedurqr(0) = SELRREQNORSALL("REQNNO")
             seldmprocedurrs.MoreResults
             seldmprocedurrs.Requery
        End If
        If seldmprocedurrs.RowCount > 0 Then
            pcount = 0
            While Not seldmprocedurrs.EOF
                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                    pcount = pcount + 1
                End If
                seldmprocedurrs.MoveNext
            Wend
        End If
        If pcount = 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        End If
        '******
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        SELRREQNORSALL.MoveNext
        Wend
End If
MsgBox "Printing is Completed"
lbl_cap.Caption = ""
Exit Sub

prtmsg: MsgBox Error(Err)
End Sub

Private Sub CMD_PRINT1_Click()
Dim CC As Integer
CC = 0
'On Error GoTo MSG:
If Trim(txt_fromlabno.Text) = "" Then
    MsgBox "Enter the Labno. From"
    Exit Sub
End If
If Trim(txt_tolabno.Text) = "" Then
    MsgBox "Enter the Lab No. upto"
    Exit Sub
End If
selpreqnoqr.QueryTimeout = 1500
If selpreqnoflag = False Then
    selpreqnoqr(0) = Trim(txt_fromlabno.Text)
    selpreqnoqr(1) = Trim(txt_tolabno.Text)
    Set selpreqnors = selpreqnoqr.OpenResultset(1, 3)
    selpreqnoflag = True
Else
    selpreqnoqr(0) = Trim(txt_fromlabno.Text)
    selpreqnoqr(1) = Trim(txt_tolabno.Text)
    selpreqnors.MoreResults
    selpreqnors.Requery
End If
For i = 0 To 13
    CrystalReport1.Formulas(i) = ""
Next

If selpreqnors.RowCount > 0 Then
    Label1.Caption = "No. of Reports in this Combination is " & selpreqnors.RowCount
    LBL_PRINTED.Caption = CC
    While Not selpreqnors.EOF
        CC = CC + 1
        LBL_PRINTED.Caption = CC
        Dim AA
        '******** Commented & added by Shraddha on 06/05/2009 ******
        var_date_f = Year(MSK_FROMDT1.FormattedText) & "," & month(MSK_FROMDT1.FormattedText) & "," & Day(MSK_FROMDT1.FormattedText)
        var_date_t = Year(MSK_TODT1.FormattedText) & "," & month(MSK_TODT1.FormattedText) & "," & Day(MSK_TODT1.FormattedText)
        '***** STAFF *****
        sel_patstat_qr(0) = selpreqnors("REQNO")
        sel_patstat_qr(1) = var_date_f
        sel_patstat_qr(2) = var_date_t
        Set sel_patstat_rs = sel_patstat_qr.OpenResultset(1, 3)
        '*****************
        
        sel_patstatOPIP_qr(0) = selpreqnors("REQNO")
        sel_patstatOPIP_qr(1) = var_date_f
        sel_patstatOPIP_qr(2) = var_date_t

        Set sel_patstatOPIP_RS = sel_patstatOPIP_qr.OpenResultset(1, 3)

'        AA = Right(txt_fromlabno.Text, 3)
'        If AA >= 300 And AA <= 350 Then
'
        If Not sel_patstat_rs.EOF Then 'ST
        '********************** Till here ***************************
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCBATCHUNI.rpt"
            'Commented and added by Shraddha on 26/08/10
            CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISBCUNI.rpt"
            'ADDED BY SHWETA 17-AUG-2010
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCUNI.rpt"
        End If
        
        If Not sel_patstatOPIP_RS.EOF Then 'IP OP
        'Else
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCFORBATCHUNI.rpt"
            'Commented and added by Shraddha on 26/08/10
            CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISBCUNI.rpt"
            'ADDED BY SHWETA 17-AUG-2010
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCUNI.rpt"
        End If
        
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")


        'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selpreqnors("REQNO")) & "' "
        'Commented and added by Shraddha on 26/05/10
        'T F condition added by Shraddha on 03/09/10
        'Order by Condition added by Shraddha on 04/10/10
        'CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(selpreqnors("REQNO")) & "' AND ({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F') order by pathno"
        CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(selpreqnors("REQNO")) & "' AND ({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F')"

        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqnors("REQNO")) & "'"

        If SELREQNOFLAG = False Then
            SELREQNOQR(0) = selpreqnors("REQNO")
            Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
            SELREQNOFLAG = True6
        Else
            SELREQNOQR(0) = selpreqnors("REQNO")
            SELREQNORS.MoreResults
            SELREQNORS.Requery
        End If
        
        If SELREQNORS.RowCount > 0 Then
            If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(SELREQNORS("ENTERBY"))
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             End If
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        '******
        If seldmprocedurflag = False Then
            seldmprocedurqr(0) = selpreqnors("REQNO")
            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
            seldmprocedurflag = True
        Else
             seldmprocedurqr(0) = selpreqnors("REQNO")
             seldmprocedurrs.MoreResults
             seldmprocedurrs.Requery
        End If
        If seldmprocedurrs.RowCount > 0 Then
            pcount = 0
            While Not seldmprocedurrs.EOF
                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                    pcount = pcount + 1
                End If
                seldmprocedurrs.MoveNext
            Wend
        End If
        If pcount = 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        End If

        '******
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
        'CrystalReport1.Destination = crptToWindow
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
     selpreqnors.MoveNext
Wend
End If


If selpreqnors.RowCount = 0 Then
    MsgBox "No records found"
Else
    MsgBox "Printing is completed"
End If
    lbl_cap.Caption = ""
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_ref_Click()
txt_fromlabno.Text = ""
txt_tolabno.Text = ""
MSK_FROMDT1.Text = ""
MSK_TODT1.Text = ""
End Sub

Private Sub Form_Load()
Dim str As String

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set selrrequisitnqr = rdb.CreateQuery("SELRREQUISITNQRBCR", str)

'********** Added by Shraddha Pandit on 06/05/2009 ********
str = "SELECT * FROM RISLIB.RREQUISITN WHERE  reqnno=? and PATNSTAT='ST' AND FREPORTDT>=? AND FREPORTDT<=?"
Set sel_patstat_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE  reqnno=? AND (PATNSTAT='IP' OR PATNSTAT='OP') AND FREPORTDT>=? AND FREPORTDT<=?"
Set sel_patstatOPIP_qr = rdb.CreateQuery("", str)

'str = "SELECT * FROM RISLIB.RREQUISITN WHERE  reqnno=? AND (PATNSTAT='IP' OR PATNSTAT='OP') AND FREPORTDT>=? AND FREPORTDT<=?"
'Set sel_patstatOPIP_qr = rdb.CreateQuery("", str)
'**********************************************************

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN WHERE WARD_NAME=?) AND  MODALITYCD='BC' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQRBC1 ", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND (REQNSTAT= 'T' OR REQNSTAT='F')AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQRALL = rdb.CreateQuery("SELRREQNOQRALLBC", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN)  AND  MODALITYCD='BC' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ? "  ' order by DISLIB.drequisitn.ward_name "
Set selrreqnoipallqr = rdb.CreateQuery("SELRREQNOQRALLip", str)

str = "SELECT pvalue FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set seldmprocedurqr = rdb.CreateQuery("seldmprocedurqrBP", str)

'************* Commented & added by Shraddha Pandit on 18/03/09 *************
'STR = "SELECT REQNO,PATHNO FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO"
'Set selpreqnoqr = rdb.CreateQuery("selpreqnoqrbr", STR)

str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY REFDOC,PATHNO"
'Commented and added by Shraddha on 04/02/10
'Commented and added by Shraddha on 08/09/10
'str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO, REFDOC"
'str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.DISBCUNILF WHERE REQNO LIKE 'FECBC%' AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO, REFDOC"
'str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.dREQUISITN WHERE REQNO LIKE 'FECBC%' AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO, REFDOC"
Set selpreqnoqr = rdb.CreateQuery("selpreqnoqrbr", str)

'*****************************************************************************
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNOQR = rdb.CreateQuery("SELREQNOQRCC", str)


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
Set selma105qr = rdb.CreateQuery("selma105qrbatrep", str)
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

Private Sub txt_fromlabno_LostFocus()
Dim NN As String
If Trim(txt_fromlabno.Text) = "" Then
    Exit Sub
End If
NN = Right(MSK_FROMDT1.FormattedText, 2) & Mid(MSK_FROMDT1.FormattedText, 4, 2) & Left(MSK_FROMDT1.FormattedText, 2)
txt_fromlabno.Text = NN & txt_fromlabno.Text
End Sub

Private Sub txt_tolabno_LostFocus()
Dim NN As String
If Trim(txt_tolabno.Text) = "" Then
    Exit Sub
End If
NN = Right(MSK_TODT1.FormattedText, 2) & Mid(MSK_TODT1.FormattedText, 4, 2) & Left(MSK_TODT1.FormattedText, 2)
txt_tolabno.Text = NN & txt_tolabno.Text
End Sub


