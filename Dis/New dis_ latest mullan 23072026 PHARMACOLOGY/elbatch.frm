VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form elbatch 
   Caption         =   "Emergency Lab Batch Mode Printing"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
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
   ScaleHeight     =   7650
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   1080
      TabIndex        =   11
      Top             =   1320
      Width           =   9135
      Begin VB.ComboBox cmb_ward 
         Height          =   315
         Left            =   6000
         TabIndex        =   18
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox cmb_patstat 
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cmb_patcatg 
         Height          =   315
         Left            =   6000
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmb_pattype 
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "Back"
         Height          =   375
         Left            =   5760
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmb_ref 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Msk_todt 
         Height          =   375
         Left            =   6000
         TabIndex        =   19
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
      Begin MSMask.MaskEdBox Msk_fromdt 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
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
      Begin VB.Label Label7 
         Caption         =   "Ward"
         Height          =   255
         Left            =   5400
         TabIndex        =   26
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Patient Status"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Patient Category"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Patient Type"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LBL_TODT 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   22
         Top             =   360
         Width           =   705
      End
      Begin VB.Label LBL_FROMDT 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   3360
      TabIndex        =   0
      Top             =   4320
      Width           =   9135
      Begin VB.TextBox txt_fromlabno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txt_tolabno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton CMD_PRINT1 
         Caption         =   "&Print"
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmd_ref 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
      Begin MSMask.MaskEdBox MSK_TODT1 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
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
         TabIndex        =   2
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
      Begin VB.Label lbl_labno1 
         AutoSize        =   -1  'True
         Caption         =   "From Lab No."
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lbl_labno2 
         AutoSize        =   -1  'True
         Caption         =   "To Lab No."
         Height          =   195
         Left            =   4560
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LBL_FROMDT1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   885
      End
      Begin VB.Label LBL_TODT1 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   30
      Top             =   0
      Width           =   11655
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   1200
      TabIndex        =   29
      Top             =   6720
      Width           =   6735
   End
   Begin VB.Label LBL_PRINTED 
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   8400
      TabIndex        =   28
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label LBL_CAP 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "EMERGENCY LAB  REPORTS"
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
      Left            =   3720
      TabIndex        =   27
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "elbatch"
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

Dim SELBG_QR As rdoQuery
Dim SELBG_RS As rdoResultset

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
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(selrrequisitnrs("REQNNO")) & "' "
        
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
        'CRYSTALREPORT1.WindowState = crptMaximized
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
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORS("REQNNO")) & "' "
        
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
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORS("REQNNO")) & "' "
        
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
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
       CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.REQNNO}= '" & Trim(SELRREQNORSALL("REQNNO")) & "' "
        
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
On Error GoTo MSG:
If Trim(txt_fromlabno.Text) = "" Then
    MsgBox "Enter the From Labno."
    Exit Sub
End If
If Trim(txt_tolabno.Text) = "" Then
    MsgBox "Enter the to Lab No."
    Exit Sub
End If
    selpreqnoqr.QueryTimeout = 1300
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
If selpreqnors.RowCount > 0 Then
        
     Label1.Caption = "No. of Reports in this Combination is " & selpreqnors.RowCount
     LBL_PRINTED.Caption = CC
     While Not selpreqnors.EOF
        CC = CC + 1
        LBL_PRINTED.Caption = CC
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISELFORBATCH.rpt"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
                
      'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & "BC" & " ' AND {RREQUISITN.REQNSTAT} =  '" & "T" & "' OR {RREQUISITN.REQNSTAT}='" & "F" & "' AND {RREQUISITN.PATNSTAT}='" & "OP" & "'AND {RREQUISITN.PAT_CATG}='" & Trim(SELRREQUISITNRS("PAT_CATG")) & "' "
       CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}= '" & Trim(selpreqnors("REQNO")) & "' "
        
         CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqnors("REQNO")) & "'"
        
         If SELREQNOFLAG = False Then
            SELREQNOQR(0) = selpreqnors("REQNO")
            Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
            SELREQNOFLAG = True
         Else
            SELREQNOQR(0) = selpreqnors("REQNO")
            SELREQNORS.MoreResults
            SELREQNORS.Requery
         End If
         If SELREQNORS.RowCount > 0 Then
             If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(SELREQNORS("ENTERBY"))
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
    
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
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
        
        'ADDED BY MILIND ON 29-03-2007 FOR DISPALYING ABG OR VBG CAPTION IN REPORT
        SELBG_QR(0) = selpreqnors("REQNO") 'VAR_REQNNO
        Set SELBG_RS = SELBG_QR.OpenResultset(1, 1)
        
        If SELBG_RS.RowCount > 0 Then
            
                If Trim(SELBG_RS("PRO_CODE")) = "F901" And SELBG_RS.RowCount = 2 Then
                    SELBG_RS.MoveNext
                    If Trim(SELBG_RS("PRO_CODE")) = "F009" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "ARTERIAL BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                        
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F995" And SELBG_RS.RowCount = 2 Then
                    SELBG_RS.MoveNext
                    If Trim(SELBG_RS("PRO_CODE")) = "F009" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "VENOUS BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                        
                End If
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
        End If
        'UPTO HERE
        
        '******
        CrystalReport1.Formulas(18) = "ipaddress = '" & ipadd & "'"
        
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
        'CrystalReport1.Destination = crptToWindow
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        

        selpreqnors.MoveNext
        Wend



End If
    MsgBox "Printing is completed"
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

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND DEPT_CODE='FED' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set selrrequisitnqr = rdb.CreateQuery("SELRREQUISITNQRBCREL", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN WHERE WARD_NAME=?) AND  MODALITYCD='BC' AND DEPT_CODE='FED' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQRBCEL ", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND DEPT_CODE='FED' AND (REQNSTAT= 'T' OR REQNSTAT='F')AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQRALL = rdb.CreateQuery("SELRREQNOQRALLEL1", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN)  AND  MODALITYCD='BC' AND DEPT_CODE='FED' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ? "  ' order by DISLIB.drequisitn.ward_name "
Set selrreqnoipallqr = rdb.CreateQuery("SELRREQNOQRALLEL11", str)

str = "SELECT pvalue FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set seldmprocedurqr = rdb.CreateQuery("seldmprocedurqrEL1", str)

str = "SELECT REQNO,PATHNO FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND DEPT_CODE='FED' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO"

Set selpreqnoqr = rdb.CreateQuery("selpreqnoqrEL1", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNOQR = rdb.CreateQuery("SELREQNOQREL1", str)

str = "SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SELBG_QR = rdb.CreateQuery("", str)


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
Set selma105qr = rdb.CreateQuery("selma105qrbatrep3", str)
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

