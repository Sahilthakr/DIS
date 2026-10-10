VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_mis_report 
   Caption         =   "MIS REPORT"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   13245
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   1560
      TabIndex        =   15
      Top             =   1080
      Width           =   8055
      Begin VB.ComboBox cmb_modality 
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
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmb_dept 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Modality"
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
         Left            =   4440
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Department "
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
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2760
      TabIndex        =   14
      Top             =   7440
      Width           =   5775
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
         Left            =   3960
         TabIndex        =   9
         Top             =   240
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
         Left            =   480
         TabIndex        =   8
         Top             =   240
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
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Period "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   11
      Top             =   2880
      Width           =   8055
      Begin MSMask.MaskEdBox Msk_todt 
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label2 
         Caption         =   "To"
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
         Left            =   4080
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "From"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame FRAME_EL 
      Height          =   2415
      Left            =   1560
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ComboBox cmb_blood 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox cmb_heamato 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cmb_biochem 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Blood Gas"
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
         Left            =   840
         TabIndex        =   28
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Heamatological"
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
         Left            =   840
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Biochemical"
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
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame FRAME_MICRO 
      Height          =   2415
      Left            =   1560
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ComboBox cmb_specimen 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1800
         Width           =   3615
      End
      Begin VB.ComboBox cmb_ward 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox cmb_unit 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   3615
      End
      Begin VB.ComboBox cmb_diag 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label8 
         Caption         =   "Specimen"
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
         Left            =   960
         TabIndex        =   22
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
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Unit"
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
         Left            =   960
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Diagnosis"
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
         Left            =   960
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2400
      TabIndex        =   30
      Top             =   360
      Width           =   7095
   End
End
Attribute VB_Name = "frm_mis_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SELDTMLF_QR As rdoQuery
Dim SELDTMLF_RS As rdoResultset
Dim SELDTMLF_FLAG As Boolean
Dim dept As String
Dim MODALITY As String
Dim MISFLAG As Boolean

Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset
Dim SELRMODALITYFLAG As Boolean

Dim SELRECFORBC_QR As rdoQuery
Dim SELRECFORBC_RS As rdoResultset
Dim SELRECFORBC_FLAG As Boolean

Dim SELMA10121_QR As rdoQuery
Dim SELMA10121_RS As rdoResultset
Dim SELMA10121FLAG As Boolean

Dim SELMA1031_QR As rdoQuery
Dim SELMA1031_RS As rdoResultset
Dim SELMA1031FLAG As Boolean

Dim SELMA106_QR1 As rdoQuery
Dim SELMA106_RS1 As rdoResultset
Dim SELMA106_FLAG1 As Boolean

Dim SELSPEC_QR As rdoQuery
Dim SELSPEC_RS As rdoResultset
Dim SELSPEC_FLAG As Boolean

Dim SELSPEREP_QR As rdoQuery
Dim SELSPEREP_RS As rdoResultset
Dim SELSPEREP_FLAG As Boolean

Dim SELTOTSAMREP_QR As rdoQuery
Dim SELTOTSAMREP_RS As rdoResultset
Dim SELTOTSAMREP_FLAG As Boolean

Dim SELDIASAMREP_QR As rdoQuery
Dim SELDIASAMREP_RS As rdoResultset
Dim SELDIASAMREP_FLAG As Boolean

Dim SELUINSAMREP_QR As rdoQuery
Dim SELUINSAMREP_RS As rdoResultset
Dim SELUINSAMREP_FLAG As Boolean

Dim SELWARSAMREP_QR As rdoQuery
Dim SELWARSAMREP_RS As rdoResultset
Dim SELWARSAMREP_FLAG As Boolean

Dim SELELLABLF_QR As rdoQuery
Dim SELELLABLF_RS As rdoResultset
Dim SELELLABLF_FLAG As Boolean

Dim SELTM_QR As rdoQuery
Dim SELTM_RS As rdoResultset
Dim SELTM_FLAG As Boolean




Private Sub cmb_dept_Click()
Dim i
i = InStr(cmb_dept.Text, "$")
dept = Mid(cmb_dept.Text, i + 1, 4)
If Trim(dept) = "FEC" Then
    Label12.Caption = " BIOCHEMISTRY MIS REPORT"
End If
If Trim(dept) = "GZZ" Then
    Label12.Caption = "MICROBIOLOGY MIS REPORT"
End If
If Trim(dept) = "NZZ" Then
    Label12.Caption = "TRANSFUSSION MEDICINE MIS REPORT"
End If


If SELRMODALITYFLAG = False Then
    SELRMODALITY_QR(0) = Trim(dept)
    Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(1, 3)
    SELRMODALITYFLAG = True
 Else
    SELRMODALITY_QR(0) = Trim(dept)
    SELRMODALITY_RS.MoreResults
    SELRMODALITY_RS.Requery
End If
If SELRMODALITY_RS.RowCount > 0 Then
    cmb_modality.clear
    SELRMODALITY_RS.MoveFirst
    While Not SELRMODALITY_RS.EOF
        cmb_modality.AddItem Trim(SELRMODALITY_RS("MODLITYDES")) & "$" & Trim(SELRMODALITY_RS("MODALITYCD"))
        If Trim(dept) = "CBZ" Then
            cmb_modality.Text = Trim(SELRMODALITY_RS("MODLITYDES")) & "$" & Trim(SELRMODALITY_RS("MODALITYCD"))
        End If
        SELRMODALITY_RS.MoveNext
    Wend
End If

End Sub


Private Sub cmb_diag_GotFocus()
If SELMA10121FLAG = False Then
        Set SELMA10121_RS = SELMA10121_QR.OpenResultset(1, 3)
        SELMA10121FLAG = True
    Else
         Set SELMA10121_RS = SELMA10121_QR.OpenResultset(1, 3)
        SELMA10121_RS.MoreResults
        SELMA10121_RS.Requery
    End If

    If SELMA10121_RS.RowCount > 0 Then
        cmb_diag.clear
        While Not SELMA10121_RS.EOF
            cmb_diag.AddItem Trim(SELMA10121_RS("PRVDIAGDES"))
            SELMA10121_RS.MoveNext
        Wend
    End If
End Sub

Private Sub cmb_modality_LostFocus()
Dim i
Dim a
a = InStr(1, Trim(cmb_modality.Text), "$")
MODALITY = Mid(Trim(cmb_modality.Text), a + 1)
If Trim(dept) = "GZZ" And (Trim(MODALITY) = "MB" Or Trim(MODALITY) = "BA") Then
    FRAME_MICRO.Visible = True
End If
If Trim(dept) = "FEC" And Trim(MODALITY) = "EL" Then
    FRAME_EL.Visible = True
End If
End Sub

Private Sub cmb_ref_Click()
msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"
FRAME_MICRO.Visible = False
FRAME_EL.Visible = False
cmb_biochem.ListIndex = -1
cmb_blood.ListIndex = -1
cmb_heamato.ListIndex = -1
cmb_diag.ListIndex = -1
cmb_specimen.ListIndex = -1
cmb_unit.ListIndex = -1
cmb_ward.ListIndex = -1
End Sub

Private Sub cmb_specimen_GotFocus()
If SELSPEC_FLAG = False Then
    Set SELSPEC_RS = SELSPEC_QR.OpenResultset(1, 3)
    SELSPEC_FLAG = True
Else
    SELSPEC_RS.MoreResults
    SELSPEC_RS.Requery
End If
If Not SELSPEC_RS.EOF Then
    cmb_specimen.clear
    Do While Not SELSPEC_RS.EOF
        cmb_specimen.AddItem Trim(SELSPEC_RS("SPECIMEN"))
        SELSPEC_RS.MoveNext
    Loop
End If
End Sub

Private Sub cmb_unit_GotFocus()
If SELMA1031FLAG = False Then
            Set SELMA1031_RS = SELMA1031_QR.OpenResultset(1, 1)
            SELMA1031FLAG = True
        Else
            Set SELMA1031_RS = SELMA1031_QR.OpenResultset(1, 1)
            SELMA1031_RS.MoreResults
            SELMA1031_RS.Requery
        End If
        If SELMA1031_RS.RowCount > 0 Then
            cmb_unit.clear
            While Not SELMA1031_RS.EOF
                cmb_unit.AddItem Trim(SELMA1031_RS("UD_NAME"))
                SELMA1031_RS.MoveNext
            Wend
        End If
    ' **** REFERRING DOCTOR AND REFERRED TO DOCTOR FILLING
End Sub

Private Sub cmb_ward_GotFocus()
If SELMA106_FLAG1 = False Then
    Set SELMA106_RS1 = SELMA106_QR1.OpenResultset(1, 3)
    SELMA106_FLAG1 = True
Else
    SELMA106_RS1.MoreResults
    SELMA106_RS1.Requery
End If
If Not SELMA106_RS1.EOF Then
    cmb_ward.clear
    Do While Not SELMA106_RS1.EOF
        cmb_ward.AddItem SELMA106_RS1("DEPT_NAME")
        SELMA106_RS1.MoveNext
    Loop
End If
End Sub

Private Sub cmd_back_Click()
msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"
FRAME_MICRO.Visible = False
FRAME_EL.Visible = False
cmb_diag.ListIndex = -1
cmb_specimen.ListIndex = -1
cmb_unit.ListIndex = -1
cmb_ward.ListIndex = -1
cmb_biochem.ListIndex = -1
cmb_blood.ListIndex = -1
cmb_heamato.ListIndex = -1
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_print_Click()
If Trim(msk_fromdt.Text) <> "" And Trim(msk_todt.Text) <> "" Then
    For i = 0 To 28
        CrystalReport1.Formulas(i) = ""
    Next i
    Dim var_yr_f As Integer, var_yr_t As Integer
    Dim var_mon_f As Integer, var_mon_t As Integer
    Dim var_day_f As Integer, var_day_t As Integer
    Dim var_date_f As String, var_date_t As String
    
    var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
    If Trim(dept) = "NZZ" And Trim(MODALITY) = "SS" Then
        CrystalReport1.SelectionFormula = "{DTMLF.RECPTDT} >= Date(" & var_date_f & ") AND {DTMLF.RECPTDT} <= Date(" & var_date_t & ") AND {DTMLF.MODALITYCD}='SS' AND ({DTMLF.REQNSTAT}='E' OR {DTMLF.REQNSTAT}='T' OR {DTMLF.REQNSTAT}='F')"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DTMMISTEST.RPT"
    End If
    If Trim(dept) = "NZZ" And Trim(MODALITY) = "BP" Then
        CrystalReport1.SelectionFormula = "{DCOMRP.ON} >= Date(" & var_date_f & ") AND {DCOMRP.ON} <= Date(" & var_date_t & ") "
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISBP.RPT"
    End If
    If Trim(dept) = "NZZ" And Trim(MODALITY) = "HT" Then
        CrystalReport1.SelectionFormula = "{DTMLF.RECPTDT} >= Date(" & var_date_f & ") AND {DTMLF.RECPTDT} <= Date(" & var_date_t & ") AND {DTMLF.MODALITYCD}='HT' AND ({DTMLF.REQNSTAT}='E' OR {DTMLF.REQNSTAT}='T' OR {DTMLF.REQNSTAT}='F')"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DTMMISTEST.RPT"
    End If
    If Trim(dept) = "FEC" And Trim(MODALITY) = "BC" Then
        Dim UREA As Double
        Dim TRUESUGAR As Double
        Dim PPSUGAR As Double
        Dim RANDOMSUGAR As Double
        Dim FURINESUGAR As Double
        Dim PPURINESUGAR As Double
        Dim LFT As Double
        Dim RFT As Double
        Dim CALC As Double
        Dim PHOS As Double
        Dim LP As Double
        Dim VMA As Double
        Dim HIAA As Double
        Dim TOTAL As Double
        
        UREA = 0
        TRUESUGAR = 0
        PPSUGAR = 0
        RANDOMSUGAR = 0
        FURINESUGAR = 0
        PPURINESUGAR = 0
        LFT = 0
        RFT = 0
        CALC = 0
        PHOS = 0
        LP = 0
        VMA = 0
        HIAA = 0
        
        If SELRECFORBC_FLAG = False Then
            SELRECFORBC_QR(0) = GET_REVDATE(Format(msk_fromdt, "DD/MM/YYYY"))
            SELRECFORBC_QR(1) = GET_REVDATE(Format(msk_todt, "DD/MM/YYYY"))
            Set SELRECFORBC_RS = SELRECFORBC_QR.OpenResultset(1, 3)
            SELRECFORBC_FLAG = True
        Else
            SELRECFORBC_QR(0) = msk_fromdt.FormattedText
            SELRECFORBC_QR(1) = msk_todt.FormattedText
            SELRECFORBC_RS.MoreResults
            SELRECFORBC_RS.Requery
        End If
        If Not SELRECFORBC_RS.EOF Then
            Do While Not SELRECFORBC_RS.EOF
                If SELRECFORBC_RS("PRO_CODE") = "F802" Then
                    TRUESUGAR = TRUESUGAR + 1
                    LFT = LFT + 1
                    RFT = RFT + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F855" Then
                    UREA = UREA + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F842" Then
                    TRUESUGAR = TRUESUGAR + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F843" Then
                    PPSUGAR = PPSUGAR + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F841" Then
                    RANDOMSUGAR = RANDOMSUGAR + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F846" Then
                    FURINESUGAR = FURINESUGAR + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F847" Then
                    PPURINESUGAR = PPURINESUGAR + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F865" Then
                    LFT = LFT + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F854" Then
                    RFT = RFT + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F888" Then
                    CALC = CALC + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F890" Then
                    PHOS = PHOS + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F849" Then
                    LP = LP + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F897" Then
                    VMA = VMA + 1
                End If
                If SELRECFORBC_RS("PRO_CODE") = "F898" Then
                    HIAA = HIAA + 1
                End If
                SELRECFORBC_RS.MoveNext
            Loop
        End If
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}='BC' AND ({RREQUISITN.REQNSTAT}='T' OR {RREQUISITN.REQNSTAT}='F')"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISBC.RPT"
        
        If Val(UREA) > 0 Then
            CrystalReport1.Formulas(2) = "UREA='" & UREA & "'"
            CrystalReport1.Formulas(3) = "TXTUREA='" & "UREA" & "'"
        Else
            CrystalReport1.Formulas(2) = "UREA='" & "" & "'"
            CrystalReport1.Formulas(3) = "TXTUREA='" & "" & "'"
        End If
        If Val(TRUESUGAR) > 0 Then
            CrystalReport1.Formulas(4) = "TRUESUGAR='" & TRUESUGAR & "'"
            CrystalReport1.Formulas(5) = "TXTTRUESUGAR='" & "True Sugar" & "'"
        Else
            CrystalReport1.Formulas(4) = "TRUESUGAR='" & "" & "'"
            CrystalReport1.Formulas(5) = "TXTTRUESUGAR='" & "" & "'"
        End If
        If Val(PPSUGAR) > 0 Then
            CrystalReport1.Formulas(6) = "PPSUGAR='" & PPSUGAR & "'"
            CrystalReport1.Formulas(7) = "TXTPPSUGAR='" & "P P Sugar" & "'"
        Else
            CrystalReport1.Formulas(6) = "PPSUGAR='" & "" & "'"
            CrystalReport1.Formulas(7) = "TXTPPSUGAR='" & "" & "'"
        End If
        If Val(RANDOMSUGAR) > 0 Then
            CrystalReport1.Formulas(8) = "RANDOMSUGAR='" & RANDOMSUGAR & "'"
            CrystalReport1.Formulas(9) = "TXTRANDOMSUGAR='" & "Random Sugar" & "'"
        Else
            CrystalReport1.Formulas(8) = "RANDOMSUGAR='" & "" & "'"
            CrystalReport1.Formulas(9) = "TXTRANDOMSUGAR='" & "" & "'"
        End If
        If Val(FURINESUGAR) > 0 Then
            CrystalReport1.Formulas(10) = "FURINESUGAR='" & FURINESUGAR & "'"
            CrystalReport1.Formulas(11) = "TXTFURINESUGAR='" & "F Urine Sugar" & "'"
        Else
            CrystalReport1.Formulas(10) = "FURINESUGAR='" & "" & "'"
            CrystalReport1.Formulas(11) = "TXTFURINESUGAR='" & "" & "'"
        End If
        If Val(PPURINESUGAR) > 0 Then
            CrystalReport1.Formulas(12) = "PPURINESUGAR='" & PPURINESUGAR & "'"
            CrystalReport1.Formulas(13) = "TXTPPURINESUGAR='" & "P P Urine Sugar" & "'"
        Else
            CrystalReport1.Formulas(12) = "PPURINESUGAR='" & "" & "'"
            CrystalReport1.Formulas(13) = "TXTPPURINESUGAR='" & "" & "'"
        End If
        If Val(LFT) > 0 Then
            CrystalReport1.Formulas(14) = "LFT='" & LFT & "'"
            CrystalReport1.Formulas(15) = "TXTLFT='" & "Liver Function Tests" & "'"
        Else
            CrystalReport1.Formulas(14) = "LFT='" & "" & "'"
            CrystalReport1.Formulas(15) = "LFT='" & "" & "'"
        End If
        If Val(RFT) > 0 Then
            CrystalReport1.Formulas(16) = "RFT='" & RFT & "'"
            CrystalReport1.Formulas(17) = "TXTRFT='" & "Renal Function Tests" & "'"
        Else
            CrystalReport1.Formulas(16) = "RFT='" & "" & "'"
            CrystalReport1.Formulas(17) = "RFT='" & "" & "'"
        End If
        If Val(CALC) > 0 Then
            CrystalReport1.Formulas(18) = "CALC='" & CALC & "'"
            CrystalReport1.Formulas(19) = "TXTCALC='" & "Calcium" & "'"
        Else
            CrystalReport1.Formulas(18) = "CALC='" & "" & "'"
            CrystalReport1.Formulas(19) = "TXTCALC='" & "" & "'"
        End If
        If Val(PHOS) > 0 Then
            CrystalReport1.Formulas(20) = "PHOS='" & PHOS & "'"
            CrystalReport1.Formulas(21) = "TXTPHOS='" & "Phosphorus" & "'"
        Else
            CrystalReport1.Formulas(20) = "PHOS='" & "" & "'"
            CrystalReport1.Formulas(21) = "TXTPHOS='" & "" & "'"
        End If
        If Val(LP) > 0 Then
            CrystalReport1.Formulas(22) = "LP='" & LP & "'"
            CrystalReport1.Formulas(23) = "TXTLP='" & "Lipid Profile" & "'"
        Else
            CrystalReport1.Formulas(22) = "LP='" & "" & "'"
            CrystalReport1.Formulas(23) = "TXTLP='" & "" & "'"
        End If
        If Val(VMA) > 0 Then
            CrystalReport1.Formulas(24) = "VMA='" & VMA & "'"
            CrystalReport1.Formulas(25) = "TXTVMA='" & "Urinary VMA" & "'"
        Else
            CrystalReport1.Formulas(24) = "VMA='" & "" & "'"
            CrystalReport1.Formulas(25) = "TXTVMA='" & "" & "'"
        End If
        If Val(HIAA) > 0 Then
            CrystalReport1.Formulas(26) = "HIAA='" & HIAA & "'"
            CrystalReport1.Formulas(27) = "TXTHIAA='" & "Phosphorus" & "'"
        Else
            CrystalReport1.Formulas(26) = "HIAA='" & "" & "'"
            CrystalReport1.Formulas(27) = "TXTHIAA='" & "" & "'"
        End If
        TOTAL = UREA + TRUESUGAR + PPSUGAR + RANDOMSUGAR + FURINESUGAR + PPURINESUGAR + LFT + RFT + CALC + PHOS + LP + VMA + HIAA
        CrystalReport1.Formulas(28) = "TOTAL='" & TOTAL & "'"
    End If
    If Trim(dept) = "GZZ" And (Trim(MODALITY) = "BA" Or Trim(MODALITY) = "MB") Then
        Dim TOTSPECDATE As String
        Dim NOSAMPLE As String
        If Trim(cmb_diag.Text) = "" And Trim(cmb_unit.Text) = "" And Trim(cmb_ward.Text) = "" And Trim(cmb_specimen.Text) = "" Then
'            If SELTOTSAMREP_FLAG = False Then
'                SELTOTSAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELTOTSAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELTOTSAMREP_QR(2) = Trim(MODALITY)
'                Set SELTOTSAMREP_RS = SELTOTSAMREP_QR.OpenResultset(1, 3)
'                SELTOTSAMREP_FLAG = True
'            Else
'                SELTOTSAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELTOTSAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELTOTSAMREP_QR(2) = Trim(MODALITY)
'                SELTOTSAMREP_RS.MoreResults
'                SELTOTSAMREP_RS.Requery
'            End If
'            If Not SELTOTSAMREP_RS.EOF Then
'                Do While Not SELTOTSAMREP_RS.EOF
'                    NOSAMPLE = NOSAMPLE + 1
'                    SELTOTSAMREP_RS.MoveNext
'                Loop
'            End If
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISMBDT.RPT"
            CrystalReport1.SelectionFormula = "{MICOBIOLF.REQNDT} >= Date(" & var_date_f & ") AND {MICOBIOLF.REQNDT} <= Date(" & var_date_t & ") AND {MICOBIOLF.MODALITYCD}='" & Trim(MODALITY) & "' AND {MICOBIOLF.DEPT_CODE}='" & Trim(dept) & "' AND NOT ({MICOBIOLF.SPECIMEN} LIKE '.*')"
                'CRYSTALREPORT1.SelectionFormula = "{MICOBIOLF.SPECIMEN}=" & NOSAMPLE
                'CRYSTALREPORT1.Formulas(2) = "NOSAMPLE='" & TOTSPECDATE & "'"
        End If
        If Trim(cmb_diag.Text) <> "" And Trim(cmb_unit.Text) = "" And Trim(cmb_ward.Text) = "" And Trim(cmb_specimen.Text) = "" Then
'            If SELDIASAMREP_FLAG = False Then
'                SELDIASAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELDIASAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELDIASAMREP_QR(2) = Trim(cmb_diag.Text)
'                SELDIASAMREP_QR(3) = Trim(MODALITY)
'                Set SELDIASAMREP_RS = SELDIASAMREP_QR.OpenResultset(1, 3)
'                SELDIASAMREP_FLAG = True
'            Else
'                SELDIASAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELDIASAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELDIASAMREP_QR(2) = Trim(cmb_diag.Text)
'                SELDIASAMREP_QR(3) = Trim(MODALITY)
'                SELDIASAMREP_RS.MoreResults
'                SELDIASAMREP_RS.Requery
'            End If
            'If Not SELDIASAMREP_RS.EOF Then
                'TOTSPECDATE = SELDIASAMREP_RS("TOTALDIA")
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISMBDT.RPT"
            CrystalReport1.SelectionFormula = "{MICOBIOLF.REQNDT} >= Date(" & var_date_f & ") AND {MICOBIOLF.REQNDT} <= Date(" & var_date_t & ") AND {MICOBIOLF.MODALITYCD}='" & Trim(MODALITY) & "' AND {MICOBIOLF.DEPT_CODE}='" & Trim(dept) & "' AND {MICOBIOLF.PRVDIAGDES}='" & Trim(cmb_diag.Text) & "' AND NOT ({MICOBIOLF.SPECIMEN} LIKE '.*')"
                'CRYSTALREPORT1.Formulas(2) = "NOSAMPLE='" & TOTSPECDATE & "'"
            'End If
        End If
        If Trim(cmb_diag.Text) = "" And Trim(cmb_unit.Text) <> "" And Trim(cmb_ward.Text) = "" And Trim(cmb_specimen.Text) = "" Then
'            If SELUINSAMREP_FLAG = False Then
'                SELUINSAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELUINSAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELUINSAMREP_QR(2) = Trim(cmb_unit.Text)
'                SELUINSAMREP_QR(3) = Trim(MODALITY)
'                Set SELUINSAMREP_RS = SELUINSAMREP_QR.OpenResultset(1, 3)
'                SELUINSAMREP_FLAG = True
'            Else
'                SELUINSAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELUINSAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELUINSAMREP_QR(2) = Trim(cmb_unit.Text)
'                SELUINSAMREP_QR(3) = Trim(MODALITY)
'                SELUINSAMREP_RS.MoreResults
'                SELUINSAMREP_RS.Requery
'            End If
'            If Not SELUINSAMREP_RS.EOF Then
'                TOTSPECDATE = SELUINSAMREP_RS("TOTALUIN")
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISMBDT.RPT"
            CrystalReport1.SelectionFormula = "{MICOBIOLF.REQNDT} >= Date(" & var_date_f & ") AND {MICOBIOLF.REQNDT} <= Date(" & var_date_t & ") AND {MICOBIOLF.MODALITYCD}='" & Trim(MODALITY) & "' AND {MICOBIOLF.DEPT_CODE}='" & Trim(dept) & "' AND {MICOBIOLF.REFBYDOC}='" & Trim(cmb_unit.Text) & "' AND NOT ({MICOBIOLF.SPECIMEN} LIKE '.*')"
                'CRYSTALREPORT1.Formulas(2) = "NOSAMPLE='" & TOTSPECDATE & "'"
        End If
        If Trim(cmb_diag.Text) = "" And Trim(cmb_unit.Text) = "" And Trim(cmb_ward.Text) <> "" And Trim(cmb_specimen.Text) = "" Then
'            If SELWARSAMREP_FLAG = False Then
'                SELWARSAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELWARSAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELWARSAMREP_QR(2) = Trim(cmb_ward.Text)
'                SELWARSAMREP_QR(3) = Trim(MODALITY)
'                Set SELWARSAMREP_RS = SELWARSAMREP_QR.OpenResultset(1, 3)
'                SELWARSAMREP_FLAG = True
'            Else
'                SELWARSAMREP_QR(0) = GET_REVDATE(Format(Msk_fromdt, "DD/MM/YYYY"))
'                SELWARSAMREP_QR(1) = GET_REVDATE(Format(Msk_todt, "DD/MM/YYYY"))
'                SELWARSAMREP_QR(2) = Trim(cmb_ward.Text)
'                SELWARSAMREP_QR(3) = Trim(MODALITY)
'                SELWARSAMREP_RS.MoreResults
'                SELWARSAMREP_RS.Requery
'            End If
'            If Not SELWARSAMREP_RS.EOF Then
'                TOTSPECDATE = SELWARSAMREP_RS("TOTALWAR")
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISMBDT.RPT"
            CrystalReport1.SelectionFormula = "{MICOBIOLF.REQNDT} >= Date(" & var_date_f & ") AND {MICOBIOLF.REQNDT} <= Date(" & var_date_t & ") AND {MICOBIOLF.MODALITYCD}='" & Trim(MODALITY) & "' AND {MICOBIOLF.DEPT_CODE}='" & Trim(dept) & "' AND {MICOBIOLF.WARD_NAME}='" & Trim(cmb_ward.Text) & "' AND NOT ({MICOBIOLF.SPECIMEN} LIKE '.%')" ' AND {MICOBIOLF.SPECIMEN} NOT LIKE '*%'" 'Len{MICOBIOLF.SPECIMEN}>1" 'AND
                'CRYSTALREPORT1.Formulas(2) = "NOSAMPLE='" & TOTSPECDATE & "'"
            'End If
        End If
        If Trim(MODALITY) = "BA" Then
            CrystalReport1.Formulas(3) = "MOD='" & "BACTERIOLOGY" & "'"
        ElseIf Trim(MODALITY) = "MB" Then
            CrystalReport1.Formulas(3) = "MOD='" & "MICOBACTERIOLOGY" & "'"
        End If
    End If
    If Trim(dept) = "FEC" And Trim(MODALITY) = "EL" Then
        If Trim(cmb_biochem.Text) = "" And Trim(cmb_blood.Text) = "" And Trim(cmb_heamato.Text) = "" Then
            MsgBox "Please choose any option from the Combo Box !!!", vbInformation, "DIS"
            cmb_biochem.SetFocus
            Exit Sub
        End If
        If cmb_biochem.Text = "PAYING" Or cmb_blood.Text = "PAYING" Or cmb_heamato.Text = "PAYING" Then
            If SELELLABLF_FLAG = False Then
                SELELLABLF_QR(0) = GET_REVDATE(Format(msk_fromdt, "DD/MM/YYYY"))
                SELELLABLF_QR(1) = GET_REVDATE(Format(msk_todt, "DD/MM/YYYY"))
                SELELLABLF_QR(2) = "A"
                SELELLABLF_QR(3) = "B"
                SELELLABLF_QR(4) = "C"
                SELELLABLF_QR(5) = "D"
                SELELLABLF_QR(6) = "F"
                Set SELELLABLF_RS = SELELLABLF_QR.OpenResultset(1, 3)
                SELELLABLF_FLAG = True
            Else
                SELELLABLF_QR(0) = GET_REVDATE(Format(msk_fromdt, "DD/MM/YYYY"))
                SELELLABLF_QR(1) = GET_REVDATE(Format(msk_todt, "DD/MM/YYYY"))
                SELELLABLF_QR(2) = "A"
                SELELLABLF_QR(3) = "B"
                SELELLABLF_QR(4) = "C"
                SELELLABLF_QR(5) = "D"
                SELELLABLF_QR(6) = "F"
                SELELLABLF_RS.MoreResults
                SELELLABLF_RS.Requery
            End If
        End If
        If cmb_biochem.Text = "NON-PAYING" Or cmb_blood.Text = "NON-PAYING" Or cmb_heamato.Text = "NON-PAYING" Then
            If SELELLABLF_FLAG = False Then
                SELELLABLF_QR(0) = GET_REVDATE(Format(msk_fromdt, "DD/MM/YYYY"))
                SELELLABLF_QR(1) = GET_REVDATE(Format(msk_todt, "DD/MM/YYYY"))
                SELELLABLF_QR(2) = "NC"
                SELELLABLF_QR(3) = "NC"
                SELELLABLF_QR(4) = "NC"
                SELELLABLF_QR(5) = "NC"
                SELELLABLF_QR(6) = "SF"
                Set SELELLABLF_RS = SELELLABLF_QR.OpenResultset(1, 3)
                SELELLABLF_FLAG = True
            Else
                SELELLABLF_QR(0) = GET_REVDATE(Format(msk_fromdt, "DD/MM/YYYY"))
                SELELLABLF_QR(1) = GET_REVDATE(Format(msk_todt, "DD/MM/YYYY"))
                SELELLABLF_QR(2) = "NC"
                SELELLABLF_QR(3) = "NC"
                SELELLABLF_QR(4) = "NC"
                SELELLABLF_QR(5) = "NC"
                SELELLABLF_QR(6) = "SF"
                SELELLABLF_RS.MoreResults
                SELELLABLF_RS.Requery
            End If
        End If
        If Not SELELLABLF_RS.EOF Then
            If Trim(cmb_biochem.Text) <> "" And Trim(cmb_blood.Text) = "" And Trim(cmb_heamato.Text) = "" Then
                Dim FS As Double
                Dim PPS As Double
                Dim RS As Double
                Dim GC As Double
                Dim Na As Double
                Dim Ka As Double
                Dim Cl As Double
                Dim u As Double
                Dim CR As Double
                Dim SGOT As Double
                Dim LDH As Double
                Dim CPK As Double
                Dim CPKMB As Double
                Dim Ca As Double
                Dim CaI As Double
                Dim Mg As Double
                Dim AM As Double
                Dim LAC As Double
                FS = 0
                PPS = 0
                RS = 0
                GC = 0
                Na = 0
                Ka = 0
                Cl = 0
                u = 0
                CR = 0
                SGOT = 0
                LDH = 0
                CPK = 0
                CPKMB = 0
                Ca = 0
                CaI = 0
                Mg = 0
                AM = 0
                LAC = 0
                If Not SELELLABLF_RS.EOF Then
                    Do While Not SELELLABLF_RS.EOF
                        If SELELLABLF_RS("PRO_CODE") = "F860" Then
                            Na = Na + 1
                            Ka = Ka + 1
                            Cl = Cl + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F875" Then
                            SGOT = SGOT + 1
                            LDH = LDH + 1
                            CPK = CPK + 1
                            CPKMB = CPKMB + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F842" Then
                            FS = FS + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F843" Then
                            PPS = PPS + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F841" Then
                            RS = RS + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F848" Then
                            GC = GC + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F861" Then
                            Na = Na + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F862" Then
                            Ka = Ka + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F863" Then
                            Cl = Cl + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F855" Then
                            u = u + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F857" Then
                            CR = CR + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F877" Then
                            SGOT = SGOT + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F914" Then
                            LDH = LDH + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F878" Then
                            CPK = CPK + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F879" Then
                            CPKMB = CPKMB + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F888" Then
                            Ca = Ca + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F889" Then
                            CaI = CaI + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F891" Then
                            Mg = Mg + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F881" Then
                            AM = AM + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F892" Then
                            LAC = LAC + 1
                        End If
                        SELELLABLF_RS.MoveNext
                    Loop
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISEL_BIO.RPT"
                    If Val(FS) > 0 Then
                        CrystalReport1.Formulas(2) = "FS='" & Trim(FS) & "'"
                    Else
                        CrystalReport1.Formulas(2) = "FS='" & "" & "'"
                    End If
                    If Val(PPS) > 0 Then
                        CrystalReport1.Formulas(3) = "PPS='" & Trim(PPS) & "'"
                    Else
                        CrystalReport1.Formulas(3) = "PPS='" & "" & "'"
                    End If
                    If Val(RS) > 0 Then
                        CrystalReport1.Formulas(4) = "RS='" & Trim(RS) & "'"
                    Else
                        CrystalReport1.Formulas(4) = "RS='" & "" & "'"
                    End If
                    If Val(GC) > 0 Then
                        CrystalReport1.Formulas(5) = "GC='" & Trim(GC) & "'"
                    Else
                        CrystalReport1.Formulas(5) = "GC='" & "" & "'"
                    End If
                    If Val(Na) > 0 Then
                        CrystalReport1.Formulas(6) = "NA='" & Trim(Na) & "'"
                    Else
                        CrystalReport1.Formulas(6) = "NA='" & "" & "'"
                    End If
                    If Val(Ka) > 0 Then
                        CrystalReport1.Formulas(7) = "KA='" & Trim(Ka) & "'"
                    Else
                        CrystalReport1.Formulas(7) = "KA='" & "" & "'"
                    End If
                    If Val(Cl) > 0 Then
                        CrystalReport1.Formulas(8) = "CL='" & Trim(Cl) & "'"
                    Else
                        CrystalReport1.Formulas(8) = "CL='" & "" & "'"
                    End If
                    If Val(u) > 0 Then
                        CrystalReport1.Formulas(9) = "U='" & Trim(u) & "'"
                    Else
                        CrystalReport1.Formulas(9) = "U='" & "" & "'"
                    End If
                    If Val(CR) > 0 Then
                        CrystalReport1.Formulas(10) = "CR='" & Trim(CR) & "'"
                    Else
                        CrystalReport1.Formulas(10) = "CR='" & "" & "'"
                    End If
                    If Val(SGOT) > 0 Then
                        CrystalReport1.Formulas(11) = "SGOT='" & Trim(SGOT) & "'"
                    Else
                        CrystalReport1.Formulas(11) = "SGOT='" & "" & "'"
                    End If
                    If Val(LDH) > 0 Then
                        CrystalReport1.Formulas(12) = "LDH='" & Trim(LDH) & "'"
                    Else
                        CrystalReport1.Formulas(12) = "LDH='" & "" & "'"
                    End If
                    If Val(CPK) > 0 Then
                        CrystalReport1.Formulas(13) = "CPK='" & Trim(CPK) & "'"
                    Else
                        CrystalReport1.Formulas(13) = "CPK='" & "" & "'"
                    End If
                    If Val(CPKMB) > 0 Then
                        CrystalReport1.Formulas(14) = "CPKMB='" & Trim(CPKMB) & "'"
                    Else
                        CrystalReport1.Formulas(14) = "CPKMB='" & "" & "'"
                    End If
                    If Val(Ca) > 0 Then
                        CrystalReport1.Formulas(15) = "CA='" & Trim(Ca) & "'"
                    Else
                        CrystalReport1.Formulas(15) = "CA='" & "" & "'"
                    End If
                    If Val(CaI) > 0 Then
                        CrystalReport1.Formulas(16) = "CAI='" & Trim(CaI) & "'"
                    Else
                        CrystalReport1.Formulas(16) = "CAI='" & "" & "'"
                    End If
                    If Val(Mg) > 0 Then
                        CrystalReport1.Formulas(17) = "MG='" & Trim(Mg) & "'"
                    Else
                        CrystalReport1.Formulas(17) = "MG='" & "" & "'"
                    End If
                    If Val(AM) > 0 Then
                        CrystalReport1.Formulas(18) = "AM='" & Trim(AM) & "'"
                    Else
                        CrystalReport1.Formulas(18) = "AM='" & "" & "'"
                    End If
                    If Val(LAC) > 0 Then
                        CrystalReport1.Formulas(19) = "LAC='" & Trim(LAC) & "'"
                    Else
                        CrystalReport1.Formulas(19) = "LAC='" & "" & "'"
                    End If
                    TOTAL = FS + PPS + RS + GC + Na + Ka + Cl + u + CR + SGOT + LDH + CPK + CPKMB + Ca + CaI + Mg + AM + LAC
                    CrystalReport1.Formulas(20) = "TOTAL='" & TOTAL & "'"
                End If
            End If
            If Trim(cmb_biochem.Text) = "" And Trim(cmb_blood.Text) <> "" And Trim(cmb_heamato.Text) = "" Then
                Dim AR As Double
                Dim cap As Double
                AR = 0
                cap = 0
                If Not SELELLABLF_RS.EOF Then
                    Do While Not SELELLABLF_RS.EOF
                        If SELELLABLF_RS("PRO_CODE") = "F901" Then
                            AR = AR + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F902" Then
                            cap = cap + 1
                        End If
                        SELELLABLF_RS.MoveNext
                    Loop
                End If
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISEL_BLOOD.RPT"
                If Val(AR) > 0 Then
                    CrystalReport1.Formulas(2) = "AR='" & Trim(AR) & "'"
                Else
                    CrystalReport1.Formulas(2) = "AR='" & "" & "'"
                End If
                If Val(cap) > 0 Then
                    CrystalReport1.Formulas(3) = "CAP='" & Trim(cap) & "'"
                Else
                    CrystalReport1.Formulas(3) = "CAP='" & "" & "'"
                End If
                TOTAL = AR + cap
                CrystalReport1.Formulas(4) = "TOTAL='" & TOTAL & "'"
            End If
            If Trim(cmb_biochem.Text) = "" And Trim(cmb_blood.Text) = "" And Trim(cmb_heamato.Text) <> "" Then
                Dim HB As Double
                Dim PC As Double
                Dim WBC As Double
                Dim BT As Double
                Dim CT As Double
                Dim PT As Double
                Dim TT As Double
                Dim PTTK As Double
                Dim PCV As Double
                HB = 0
                PC = 0
                WBC = 0
                BT = 0
                CT = 0
                PT = 0
                TT = 0
                PTTK = 0
                PCV = 0
                If Not SELELLABLF_RS.EOF Then
                    Do While Not SELELLABLF_RS.EOF
                        If SELELLABLF_RS("PRO_CODE") = "F903" Then
                            HB = HB + 1
                            PC = PC + 1
                            WBC = WBC + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F908" Then
                            BT = BT + 1
                            CT = CT + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F904" Then
                            HB = HB + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F906" Then
                            PC = PC + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F905" Then
                            WBC = WBC + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F909" Then
                            BT = BT + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F910" Then
                            CT = CT + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F911" Then
                            PT = PT + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F913" Then
                            TT = TT + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F912" Then
                            PTTK = PTTK + 1
                        End If
                        If SELELLABLF_RS("PRO_CODE") = "F907" Then
                            PCV = PCV + 1
                        End If
                        SELELLABLF_RS.MoveNext
                    Loop
                End If
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISEL_HEAMATO.RPT"
                If Val(HB) > 0 Then
                    CrystalReport1.Formulas(2) = "HB='" & Trim(HB) & "'"
                Else
                    CrystalReport1.Formulas(2) = "HB='" & "" & "'"
                End If
                If Val(PC) > 0 Then
                    CrystalReport1.Formulas(3) = "PC='" & Trim(PC) & "'"
                Else
                    CrystalReport1.Formulas(3) = "PC='" & "" & "'"
                End If
                If Val(WBC) > 0 Then
                    CrystalReport1.Formulas(4) = "WBC='" & Trim(WBC) & "'"
                Else
                    CrystalReport1.Formulas(4) = "WBC='" & "" & "'"
                End If
                If Val(BT) > 0 Then
                    CrystalReport1.Formulas(5) = "BT='" & Trim(BT) & "'"
                Else
                    CrystalReport1.Formulas(5) = "BT='" & "" & "'"
                End If
                If Val(CT) > 0 Then
                    CrystalReport1.Formulas(6) = "CT='" & Trim(CT) & "'"
                Else
                    CrystalReport1.Formulas(6) = "CT='" & "" & "'"
                End If
                If Val(PT) > 0 Then
                    CrystalReport1.Formulas(7) = "PT='" & Trim(PT) & "'"
                Else
                    CrystalReport1.Formulas(7) = "PT='" & "" & "'"
                End If
                If Val(TT) > 0 Then
                    CrystalReport1.Formulas(8) = "TT='" & Trim(TT) & "'"
                Else
                    CrystalReport1.Formulas(8) = "TT='" & "" & "'"
                End If
                If Val(PTTK) > 0 Then
                    CrystalReport1.Formulas(9) = "PTTK='" & Trim(PTTK) & "'"
                Else
                    CrystalReport1.Formulas(9) = "PTTK='" & "" & "'"
                End If
                If Val(PCV) > 0 Then
                    CrystalReport1.Formulas(10) = "PCV='" & Trim(PCV) & "'"
                Else
                    CrystalReport1.Formulas(10) = "PCV='" & "" & "'"
                End If
                TOTAL = HB + PC + WBC + BT + CT + PT + TT + PTTK + PCV
                CrystalReport1.Formulas(11) = "TOTAL='" & TOTAL & "'"
            End If
        End If
    End If
    If Trim(dept) = "FEC" And Trim(MODALITY) = "TM" Then
        Dim CEA As Double
        Dim AFP As Double
        Dim BHCG As Double
        Dim CA125 As Double
        Dim CA199 As Double
        Dim PLAP As Double
        Dim B2MICRO As Double
        Dim PSA As Double
        Dim CA153 As Double
        Dim FERRITIN As Double
        Dim CYCLO As Double
        Dim VANCO As Double
        Dim SEP As Double
        Dim BJP As Double
        CEA = 0
        AFP = 0
        BHCG = 0
        CA125 = 0
        CA199 = 0
        PLAP = 0
        B2MICRO = 0
        PSA = 0
        CA153 = 0
        FERRITIN = 0
        CYCLO = 0
        VANCO = 0
        SEP = 0
        BJP = 0
        If SELTM_FLAG = False Then
            SELTM_QR(0) = GET_REVDATE(Format(msk_fromdt, "DD/MM/YYYY"))
            SELTM_QR(1) = GET_REVDATE(Format(msk_todt, "DD/MM/YYYY"))
            Set SELTM_RS = SELTM_QR.OpenResultset(1, 3)
            SELTM_FLAG = True
        Else
            SELTM_QR(0) = GET_REVDATE(Format(msk_fromdt, "DD/MM/YYYY"))
            SELTM_QR(1) = GET_REVDATE(Format(msk_todt, "DD/MM/YYYY"))
            SELTM_RS.MoreResults
            SELTM_RS.Requery
        End If
        If Not SELTM_RS.EOF Then
            Do While Not SELTM_RS.EOF
                If SELTM_RS("PRO_CODE") = "F818" Then
                    CEA = CEA + 1
                End If
                If SELTM_RS("PRO_CODE") = "F817" Then
                    AFP = AFP + 1
                End If
                If SELTM_RS("PRO_CODE") = "F819" Then
                    BHCG = BHCG + 1
                End If
                If SELTM_RS("PRO_CODE") = "F823" Then
                    CA125 = CA125 + 1
                End If
                If SELTM_RS("PRO_CODE") = "F824" Then
                    CA199 = CA199 + 1
                End If
                If SELTM_RS("PRO_CODE") = "F826" Then
                PLAP = PLAP + 1
                End If
                If SELTM_RS("PRO_CODE") = "F821" Then
                    B2MICRO = B2MICRO + 1
                End If
                If SELTM_RS("PRO_CODE") = "F820" Then
                    PSA = PSA + 1
                End If
                If SELTM_RS("PRO_CODE") = "F822" Then
                    CA153 = CA153 + 1
                End If
                If SELTM_RS("PRO_CODE") = "F833" Then
                    CYCLO = CYCLO + 1
                End If
                If SELTM_RS("PRO_CODE") = "F835" Then
                    VANCO = VANCO + 1
                End If
                If SELTM_RS("PRO_CODE") = "F816" Then
                    SEP = SEP + 1
                End If
                If SELTM_RS("PRO_CODE") = "F948" Then
                    BJP = BJP + 1
                End If
                SELTM_RS.MoveNext
            Loop
        End If
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MISTM.RPT"
        If Val(CEA) > 0 Then
            CrystalReport1.Formulas(2) = "CEA='" & Trim(CEA) & "'"
        Else
            CrystalReport1.Formulas(2) = "CEA='" & "" & "'"
        End If
        If Val(AFP) > 0 Then
            CrystalReport1.Formulas(3) = "AFP='" & Trim(AFP) & "'"
        Else
            CrystalReport1.Formulas(3) = "AFP='" & "" & "'"
        End If
        If Val(BHCG) > 0 Then
            CrystalReport1.Formulas(4) = "BHCG='" & Trim(BHCG) & "'"
        Else
            CrystalReport1.Formulas(4) = "BHCG='" & "" & "'"
        End If
        If Val(CA125) > 0 Then
            CrystalReport1.Formulas(5) = "CA125='" & Trim(CA125) & "'"
        Else
            CrystalReport1.Formulas(5) = "CA125='" & "" & "'"
        End If
        If Val(CA199) > 0 Then
            CrystalReport1.Formulas(6) = "CA199='" & Trim(CA199) & "'"
        Else
            CrystalReport1.Formulas(6) = "CA199='" & "" & "'"
        End If
        If Val(PLAP) > 0 Then
            CrystalReport1.Formulas(7) = "PLAP='" & Trim(PLAP) & "'"
        Else
            CrystalReport1.Formulas(7) = "PLAP='" & "" & "'"
        End If
        If Val(B2MICRO) > 0 Then
            CrystalReport1.Formulas(8) = "B2MICRO='" & Trim(B2MICRO) & "'"
        Else
            CrystalReport1.Formulas(8) = "B2MICRO='" & "" & "'"
        End If
        If Val(PSA) > 0 Then
            CrystalReport1.Formulas(9) = "PSA='" & Trim(PSA) & "'"
        Else
            CrystalReport1.Formulas(9) = "PSA='" & "" & "'"
        End If
        If Val(CA153) > 0 Then
            CrystalReport1.Formulas(10) = "CA153='" & Trim(CA153) & "'"
        Else
            CrystalReport1.Formulas(10) = "CA153='" & "" & "'"
        End If
        If Val(CYCLO) > 0 Then
            CrystalReport1.Formulas(11) = "CYCLO='" & Trim(CYCLO) & "'"
        Else
            CrystalReport1.Formulas(11) = "CYCLO='" & "" & "'"
        End If
        If Val(VANCO) > 0 Then
            CrystalReport1.Formulas(12) = "VANCO='" & Trim(VANCO) & "'"
        Else
            CrystalReport1.Formulas(12) = "VANCO='" & "" & "'"
        End If
        If Val(SEP) > 0 Then
            CrystalReport1.Formulas(13) = "SEP='" & Trim(SEP) & "'"
        Else
            CrystalReport1.Formulas(13) = "SEP='" & "" & "'"
        End If
        If Val(BJP) > 0 Then
            CrystalReport1.Formulas(14) = "BJP='" & Trim(BJP) & "'"
        Else
            CrystalReport1.Formulas(14) = "BJP='" & "" & "'"
        End If
        TOTAL = CEA + AFP + BHCG + CA125 + CA199 + PLAP + B2MICRO + PSA + CA153 + CYCLO + VANCO + SEP + BJP
        CrystalReport1.Formulas(12) = "TOTAL='" & TOTAL & "'"
    End If
    CrystalReport1.Formulas(0) = "FROMDT='" & Trim(msk_fromdt.Text) & "'"
    CrystalReport1.Formulas(1) = "TODT='" & Trim(msk_todt.Text) & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
End If

End Sub

Private Sub Form_Load()
If MISFLAG = False Then
    MISFLAG = True
    str = "select * from rislib.rmodality where dept_code=?"
    Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr19091", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' OR DEPT_CODE='IZZ' ) "
    Set SELMA106_QR = rdb.CreateQuery("SELMA106901_QR1", str)
    
    str = "SELECT DISLIB.DMULPROCE.*,RISLIB.RREQUISITN.* FROM DISLIB.DMULPROCE,RISLIB.RREQUISITN WHERE REQNO=REQNNO AND REQNSTAT IN('T','F') AND MODALITYCD='BC' AND RISLIB.RREQUISITN.REQNDT BETWEEN ? AND ?"
    Set SELRECFORBC_QR = rdb.CreateQuery("SELRECFORBC_STR", str)
    
    str = "SELECT * FROM DISLIB.ELLABLF WHERE MODALITYCD='EL' AND REQNDT BETWEEN ? AND ? AND (PAT_CATG=? OR PAT_CATG=? OR PAT_CATG=? OR PAT_CATG=? OR PAT_CATG=?)"
    Set SELELLABLF_QR = rdb.CreateQuery("SELRECFORBC_STR109", str)
    
    
    str = "SELECT * FROM DISLIB.ELLABLF WHERE MODALITYCD='TM' AND REQNDT BETWEEN ? AND ?"
    Set SELTM_QR = rdb.CreateQuery("SELRECFORBC_STR109009", str)
    
    str = "SELECT * FROM DISLIB.DTMLF WHERE MODALITYCD IN('SS','BP','HT') AND REQNSTAT IN('E','T','F') "
    Set SELDTMLF_QR = rdb.CreateQuery("SELDTMLF_STR", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA10121 ORDER BY PRVDIAGDES"
    Set SELMA10121_QR = rdb.CreateQuery("SELMA101211_QR11009", str)
    
    str = "SELECT DISTINCT UD_NAME FROM PATOOLSLIB.MA1031 ORDER BY UD_NAME"
    Set SELMA1031_QR = rdb.CreateQuery("SELMA1031_QR1009", str)
   
    str = "SELECT DEPT_NAME FROM PATOOLSLIB.MA106 ORDER BY DEPT_NAME"
    Set SELMA106_QR1 = rdb.CreateQuery("SELMA106_STRXXX", str)
    
    str = "SELECT DISTINCT SPECIMEN FROM DISLIB.DMSPONANTI ORDER BY SPECIMEN"
    Set SELSPEC_QR = rdb.CreateQuery("SELSPEC_STR11", str)

    str = "SELECT SPECIMEN FROM DISLIB.MICOBIOLF WHERE DEPT_CODE='GZZ' AND REQNDT BETWEEN ? AND ? AND MODALITYCD=? AND LENGTH(SPECIMEN)>1 AND SPECIMEN NOT LIKE '.%' AND SPECIMEN NOT LIKE '*%'"
    Set SELTOTSAMREP_QR = rdb.CreateQuery("SELTOTSAMREP_STR", str)

    str = "SELECT COUNT(SPECIMEN)TOTALDIA FROM DISLIB.MICOBIOLF WHERE DEPT_CODE='GZZ' AND REQNDT BETWEEN ? AND ? AND PRVDIAGDES=? AND MODALITYCD=?"
    Set SELDIASAMREP_QR = rdb.CreateQuery("SELDIATOTSAMREP_STR", str)
        
    str = "SELECT COUNT(SPECIMEN)TOTALUIN FROM DISLIB.MICOBIOLF WHERE DEPT_CODE='GZZ' AND REQNDT BETWEEN ? AND ? AND REFBYDOC=? AND MODALITYCD=?"
    Set SELUINSAMREP_QR = rdb.CreateQuery("SELUINTOTSAMREP_STR", str)
    
    str = "SELECT COUNT(SPECIMEN)TOTALWAR FROM DISLIB.MICOBIOLF WHERE DEPT_CODE='GZZ' AND REQNDT BETWEEN ? AND ? AND WARD_NAME=? AND MODALITYCD=? AND LENGTH(SPECIMEN)>1 AND SPECIMEN NOT LIKE '.%' AND SPECIMEN NOT LIKE '*%'"
    Set SELWARSAMREP_QR = rdb.CreateQuery("SELWARTOTSAMREP_STR", str)

End If
' **** DEPARTMENT NAME COMBO BOX FILLING
If SELMA106FLAG = False Then
    SELMA106
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106FLAG = True
Else
    SELMA106
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106_RS.MoreResults
    SELMA106_RS.Requery
End If
If SELMA106_RS.RowCount > 0 Then
    cmb_dept.clear
    While Not SELMA106_RS.EOF
    cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & SELMA106_RS("DEPT_CODE")
    SELMA106_RS.MoveNext
    Wend
End If
' **** DEPARTMENT NAME COMBO BOX FILLING
cmb_biochem.AddItem "PAYING"
cmb_biochem.AddItem "NON-PAYING"

cmb_blood.AddItem "PAYING"
cmb_blood.AddItem "NON-PAYING"

cmb_heamato.AddItem "PAYING"
cmb_heamato.AddItem "NON-PAYING"

msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"

datetime
End Sub

Private Sub Msk_fromdt_GotFocus()
msk_fromdt.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub Msk_todt_GotFocus()
msk_todt.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub CLEAR_FIELD()
msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"
FRAME_MICRO.Visible = False
FRAME_EL.Visible = False
cmb_diag.ListIndex = -1
cmb_unit.ListIndex = -1
cmb_ward.ListIndex = -1
End Sub
