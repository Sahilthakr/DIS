VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form HEMBMA 
   Caption         =   "Requisition for Bone Marrow Aspiration"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
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
   ScaleHeight     =   8700
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5400
      TabIndex        =   45
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   720
      TabIndex        =   39
      Top             =   5640
      Width           =   10695
      Begin VB.CheckBox chk_freezing 
         Caption         =   "Freezing"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frame_cyto 
      Caption         =   "Cyctochemistry"
      Height          =   1095
      Left            =   720
      TabIndex        =   30
      Top             =   6240
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CheckBox chk_mpo 
         Caption         =   "MPO"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_SBB 
         Caption         =   "SBB"
         Height          =   255
         Left            =   3120
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_PAS 
         Caption         =   "PAS"
         Height          =   255
         Left            =   7560
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_LAP 
         Caption         =   "LAP"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_NSE 
         Caption         =   "NSE"
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_TRAP 
         Caption         =   "TRAP"
         Height          =   255
         Left            =   7560
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "TRAP"
         Height          =   255
         Left            =   8520
         TabIndex        =   55
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "PAS"
         Height          =   255
         Left            =   8520
         TabIndex        =   54
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "NSE"
         Height          =   255
         Left            =   3960
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "SBB"
         Height          =   255
         Left            =   3960
         TabIndex        =   52
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "LAP"
         Height          =   255
         Left            =   960
         TabIndex        =   51
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "MPO"
         Height          =   255
         Left            =   960
         TabIndex        =   50
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<< &Back"
      Height          =   615
      Left            =   9120
      TabIndex        =   23
      Top             =   7680
      Width           =   1815
   End
   Begin VB.ComboBox cmb_bmaspno 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2760
      TabIndex        =   22
      Top             =   7440
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Studies Required"
      Height          =   1335
      Left            =   720
      TabIndex        =   2
      Top             =   4320
      Width           =   10695
      Begin VB.TextBox txtSurface 
         Height          =   285
         Left            =   4920
         TabIndex        =   60
         Top             =   960
         Width           =   5655
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Modul"
         Height          =   255
         Left            =   4920
         TabIndex        =   49
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox CHK_SMC 
         Caption         =   "Surface Marker - Complete"
         Height          =   195
         Left            =   4920
         TabIndex        =   43
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "HEMBMA.frx":0000
         Left            =   7440
         List            =   "HEMBMA.frx":000D
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chk_dnaindex 
         Caption         =   "DNA Index"
         Height          =   195
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chk_sm 
         Caption         =   "Surface Marker - Single"
         Height          =   195
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chk_cytochemistry 
         Caption         =   "Cytochemistry + Morphology"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   4335
      End
      Begin VB.CheckBox CHK_MORPHOLOGY 
         Caption         =   "Morphology"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   4800
         X2              =   4800
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Studies"
         Height          =   195
         Left            =   8400
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nature of Material "
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   3240
      Width           =   10695
      Begin VB.ComboBox CMB_NMATERIAL 
         Height          =   315
         ItemData        =   "HEMBMA.frx":0036
         Left            =   1800
         List            =   "HEMBMA.frx":004F
         TabIndex        =   41
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Out side BM Biopsy"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TXT_NOTHERS 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   25
         Top             =   360
         Width           =   5295
      End
      Begin VB.CheckBox CHK_OUTSIDEIMPRINT 
         Caption         =   "Outside Imprint"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chk_oaspimp 
         Caption         =   "Outside BM asp.+imp."
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chk_obas 
         Caption         =   "Outside BM asp.Slides"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox CHK_ASPIMP 
         Caption         =   "Aspirate+Imprint"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chk_oi 
         Caption         =   "Only Imprint"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chk_oa 
         Caption         =   "Only Aspirate"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Material"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Others"
         Height          =   195
         Left            =   4680
         TabIndex        =   26
         Top             =   480
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clinical Details"
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   10695
      Begin VB.CheckBox chk_hep 
         Caption         =   "Hepatomegaly"
         Height          =   255
         Left            =   6000
         TabIndex        =   58
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chk_spl 
         Caption         =   "Splenomegaly"
         Height          =   255
         Left            =   6000
         TabIndex        =   57
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chk_lym 
         Caption         =   "Lymphadenopathy"
         Height          =   255
         Left            =   6000
         TabIndex        =   56
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox CHK_OUTSIDESLIDES 
         Caption         =   "Outside Slides"
         Height          =   255
         Left            =   4200
         TabIndex        =   44
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TXT_OTHERS 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   7920
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox CHK_OTHERS 
         Caption         =   "Others"
         Height          =   195
         Left            =   7920
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chk_relapse 
         Caption         =   "Relapse"
         Height          =   195
         Left            =   2160
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chk_dm 
         Caption         =   "Day 10 Marrow"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chk_pb 
         Caption         =   "Post BMT"
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chk_ts 
         Caption         =   "Taken Steroids"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chk_pt 
         Caption         =   "Post Transfusion"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chk_Posttreatment 
         Caption         =   "Post Induction"
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chk_partiallytreated 
         Caption         =   "Partially Treated"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chk_nc 
         Caption         =   "New Case"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   2040
      TabIndex        =   46
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label10 
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
      Height          =   375
      Left            =   120
      TabIndex        =   59
      Top             =   120
      Width           =   11175
   End
   Begin VB.Label lbl_caseno 
      Caption         =   "Case No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   48
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   47
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lbl_cap 
      AutoSize        =   -1  'True
      Caption         =   "BONE MARROW ASPIRATION  - REQUISITION"
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
      Left            =   2760
      TabIndex        =   24
      Top             =   600
      Width           =   6405
   End
   Begin VB.Label lbl_pbmaspno 
      AutoSize        =   -1  'True
      Caption         =   "Previous Investigations"
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   7440
      Width           =   1995
   End
End
Attribute VB_Name = "HEMBMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''******REPORT********
Dim seldrequisitn_qr10 As rdoQuery
Dim seldrequisitn_rs10 As rdoResultset
Dim seldrequisitn_flag10 As Boolean

Dim seldrequisitn_qr11 As rdoQuery
Dim seldrequisitn_rs11 As rdoResultset
Dim seldrequisitn_flag11 As Boolean


Dim X
Dim DEPT1 As String
Dim MODALITY1 As String
Dim REQNO1 As String


Dim SELRREQNOQR As rdoQuery
Dim SELRREQNORS As rdoResultset
Dim SELRREQNOFLAG As Boolean

Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUM_FLAG As Boolean

Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset

Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset

Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset

Dim SELMA301QR As rdoQuery
Dim SELMA301RS As rdoResultset

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset

''HLREPORT
Dim SELDCBCREPORT_QR As rdoQuery
Dim SELDCBCREPORT_RS As rdoResultset
Dim SELDCBCREPORT_FLAG As Boolean
''HLREPORT
Dim SELDBMAREPORT_QR As rdoQuery
Dim SELDBMAREPORT_RS As rdoResultset
Dim SELDBMAREPORT_FLAG As Boolean
'''******REPORT********

Private Sub chk_cytochemistry_Click()
If chk_cytochemistry.Value = 1 Then
    'frame_cyto.Visible = True
    CHK_MORPHOLOGY.Value = 1
    
Else
    'frame_cyto.Visible = False
    CHK_MORPHOLOGY.Value = 0
End If
End Sub

Private Sub CHK_OTHERS_Click()
If CHK_OTHERS.Value = 1 Then
  txt_others.Visible = True
Else
  txt_others.Visible = False
End If
End Sub

Private Sub chk_sm_Click()
If chk_sm.Value = 1 Then
    CHK_SMC.Value = 0
    
End If
End Sub

Private Sub CHK_SMC_Click()
If CHK_SMC.Value = 1 Then
    chk_sm.Value = 0
End If
End Sub

Private Sub cmb_bmaspno_Click()
On Error GoTo Errhandler
Dim X
X = InStr(1, Trim(cmb_bmaspno.Text), "$")
REQNO1 = Mid(Trim(cmb_bmaspno.Text), 1, X - 1)
DEPT1 = Mid(Trim(REQNO1), 1, 3)
MODALITY1 = Mid(Trim(REQNO1), 5, 2)
If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
    SPREPORT
ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
    FAREPORT
ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
    CGCNREPORT
ElseIf DEPT1 = "UZZ" And MODALITY1 = "BM" Then
    BONEMARROW
ElseIf DEPT1 = "UZZ" And MODALITY1 = "BF" Then
    BFREPORT
ElseIf DEPT1 = "UZZ" And MODALITY1 = "HL" Then
    HLREPORT
End If
Exit Sub
Errhandler:
MsgBox Err.Description, vbCritical, "DIS"
End Sub

Private Sub cmd_back_Click()

pathology.CMB_TEXT.clear
ccount = 0
If BMFLAG = False Then
    meterial = cmb_nmaterial.Text & "+" & TXT_NOTHERS.Text
        ind1 = ""
        ind2 = ""
        ind3 = ""
        ind4 = ""
        ind5 = ""
        ind6 = ""
        ind7 = ""
        ind8 = ""
        ind9 = ""
        tp1 = ""
        tp2 = ""
        tp3 = ""
        
     '***** CLINICAL iNDICATION
     If chk_nc.Value = 1 Then
        ind1 = chk_nc.Caption
     Else
        ind1 = ""
     End If
     If chk_pt.Value = 1 Then
        ind2 = chk_pt.Caption
     Else
        ind2 = ""
     End If
     If chk_dm.Value = 1 Then
        ind3 = chk_dm.Caption
     Else
        ind3 = ""
     End If
     If chk_partiallytreated.Value = 1 Then
        ind4 = chk_partiallytreated.Caption
     Else
        ind4 = ""
     End If
      If chk_ts.Value = 1 Then
        ind5 = chk_ts.Caption
     Else
        ind5 = ""
     End If
     If chk_relapse.Value = 1 Then
        ind6 = chk_relapse.Caption
     Else
        ind6 = ""
     End If
     If chk_Posttreatment.Value = 1 Then
        ind7 = chk_Posttreatment.Caption
     Else
        ind7 = ""
     End If
     If chk_pb.Value = 1 Then
        ind8 = chk_pb.Caption
     Else
        ind8 = ""
     End If
     If CHK_OTHERS.Value = 1 Then
        ind9 = txt_others.Text
     Else
        ind9 = ""
     End If
     ''******BY MONO********
     If chk_lym.Value = 1 Then
        tp1 = chk_lym.Caption
     Else
        tp1 = ""
     End If
     If chk_spl.Value = 1 Then
        tp2 = chk_spl.Caption
     Else
        tp2 = ""
     End If
     If chk_hep.Value = 1 Then
        tp3 = chk_hep.Caption
     Else
        tp3 = ""
     End If
     ''******BY MONO********
     '***** CLINICAL INDICATION
    If CHK_MORPHOLOGY.Value = 0 And chk_cytochemistry.Value = 0 And chk_sm.Value = 0 And CHK_SMC.Value = 0 Then
        MsgBox ("Select (MOPHOLOGY/CYTOCHEMISTRY),SURFACE MARKER or ALL of them")
        Exit Sub
    End If
     
    If chk_cytochemistry.Value = 1 And CHK_MORPHOLOGY.Value = 1 Then
        If ccount > 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "U752"
        ccount = ccount + 1
    End If
    If CHK_MORPHOLOGY.Value = 1 And chk_cytochemistry.Value = 0 Then
        If ccount > 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "U751"
        ccount = ccount + 1
    End If
    
    If chk_sm.Value = 1 And CHK_SMC.Value = 0 Then
        If ccount > 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "U754"
        ccount = ccount + 1
    End If
    If CHK_SMC.Value = 1 And chk_sm.Value = 0 Then
        If ccount > 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "U753"
        ccount = ccount + 1
    End If
    
    If chk_sm.Value = 1 Or CHK_SMC.Value = 1 Then
        str_surface = Trim(txtSurface.Text)
    End If
    'If chk_dnaindex.Value = 1  Then
    '    If ccount > 11 Then
    '        MsgBox "You Can Select Max. of 11 Procedures only."
    '        Unload Me
    '        pathology.show
    '        Exit Sub
    '    End If
    '    pathology.CMB_TEXT.AddItem "F753"
    '    ccount = ccount + 1
    'End If
End If
BMFLAG = False
Unload Me
pathology.show
End Sub

Private Sub Form_Load()
'clear Public variables added by v.s.k. on 24-05-2012
str_surface = ""
'end
 
 ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ***********
If HEMBMAFLAG = False Then
    ''DCBCREPORT FOR HL
    str = "SELECT * FROM DISLIB.DCBCREPORT WHERE REQNO=?"
    Set SELDCBCREPORT_QR = rdb.CreateQuery("SELDCBCREPORTX_STRSDDED", str)
    
    ''DCBCREPORT FOR HL
    str = "SELECT * FROM DISLIB.DBMAREPORT WHERE REQNO=?"
    Set SELDBMAREPORT_QR = rdb.CreateQuery("SELDBMAREPORTX_STRCVGF", str)
    
    'STR = "SELECT * FROM RISLIB.RREQUISITN LEFT OUTER JOIN DISLIB.DREQUISITN ON RISLIB.RREQUISITN.REQNNO=DISLIB.DREQUISITN.REQNO WHERE RISLIB.RREQUISITN.CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')" 'ORDER BY RREQUISITN.REQNDT DESC"
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr10 = rdb.CreateQuery("seldrequisitn11X_1qr11XX0", str)
         
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr11 = rdb.CreateQuery("seldrequisitn11X_1qrXXREP", str)
     
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set SELRREQNOQR = rdb.CreateQuery("SELXRREQNOQR1SPREQXX", str)
     
    str = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUMX_QR11SPREQXX", str)
        
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1SPREQX_QR1XX", str)
    
    str = "SELECT REQNO,FSPREPORT,IHREPORT,IMPRESSION,GROSSDESC,GROSSEDBY,SPREPORT,REPORTIND,FNAC FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORTX_1QR1sprepXX", str)
    
    str = "SELECT * FROM PATOOLSLIB.TA1012 WHERE CASE_NO=?"
    Set selta1012_qr = rdb.CreateQuery("SELTA1012X_QR11kl", str)
    selta1012_qr(0) = "TT/00333"
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    
    str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
    Set SELMA301QR = rdb.CreateQuery("SELMA301QR1QSWX", str)
    
    SELMA301QR(0) = "F22"
    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
    
    ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ***********End Sub
    HEMBMAFLAG = True
End If

End Sub
Private Sub SPREPORT()
Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim i As Integer
On Error GoTo prtmsg:
DoEvents
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 150
If seldrequisitn_flag11 = False Then
    seldrequisitn_qr11(0) = msk_caseno.FormattedText
    Set seldrequisitn_rs11 = seldrequisitn_qr11.OpenResultset(1, 3)
    seldrequisitn_flag11 = True
Else
    seldrequisitn_qr11(0) = msk_caseno.FormattedText
    seldrequisitn_rs11.MoreResults
    seldrequisitn_rs11.Requery
End If
If seldrequisitn_rs11.RowCount > 0 Then
    prpathno = ""
    'CMB_PATHNO.clear
    While Not seldrequisitn_rs11.EOF
        prpathno = prpathno & " " & seldrequisitn_rs11("pathno")
        'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
        seldrequisitn_rs11.MoveNext
    Wend
    'txt_prepathreqno.Text = prpathno
End If
If SELDADDENDUM_FLAG = False Then
    SELDADDENDUM_QR(0) = VAR_REQNNO
    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
    SELDADDENDUM_FLAG = True
Else
    SELDADDENDUM_QR(0) = VAR_REQNNO
    SELDADDENDUM_RS.MoreResults
    SELDADDENDUM_RS.Requery
End If

For i = 0 To 25
    CRYSTALREPORT1.Formulas(i) = ""
Next i


str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"
If pathology.opt_patient.Value = True Then
        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
            Else
                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            End If
        Else
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CRYSTALREPORT1.Formulas(1) = "PREPATH = 'Prev. Path No." & Trim(prpathno) & "'"
        Else
             CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = msk_caseno.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        Dim z
        Dim Y
        Dim w
        Dim u
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            If Trim(selpathdrequisitn_rs("OPPROCEDUR")) <> "" Then
                z = InStr(1, Trim(selpathdrequisitn_rs("OPPROCEDUR")), "'")
                If z <> 0 Then
                    Y = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
                    u = Len(Y)
                    X = Left(Y, z - 1)
                    w = Right(Y, u - (z))
                    Y = X & w
                    CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Y & "'"
                Else
                    CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
                End If
            Else
                CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
            End If
        Else
            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
        End If
        
        CRYSTALREPORT1.Formulas(8) = "USRID = '" & str & "'"
        
        'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
        'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
        CRYSTALREPORT1.WindowState = crptMaximized
      '  CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.Action = 1
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
'        If SELEMPBASINFO_RS.RowCount > 0 Then
'            txt_sex.Text = SELEMPBASINFO_RS("GENDER")
'            txt_name.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
'            txt_age = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), CStr(DATE11))
'        End If

        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
            Else
                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
            End If
        Else
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(selpathdrequisitn_rs("prepathno")) & "'"
        Else
            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = MaskEdBox1.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301_QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
        Else
            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
         
        End If
        CRYSTALREPORT1.Formulas(8) = "USRID = '" & str & "'"
        CRYSTALREPORT1.Formulas(13) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CRYSTALREPORT1.Formulas(14) = "DOB = '" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT")) & "'"
        CRYSTALREPORT1.Formulas(15) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CRYSTALREPORT1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.Action = 1
    End If
'CrystalReport1.WindowState = crptMaximized
'CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox Error(Err)
'txt_msg.Text = ""
'btnenabl ("print")
Exit Sub
MSG:
MsgBox Error(Err)
End Sub
Private Sub FAREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 300
'If seldrequisitn_flag10 = False Then
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
'    seldrequisitn_flag10 = True
'Else
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    seldrequisitn_rs10.MoreResults
'    seldrequisitn_rs10.Requery
'End If
'If seldrequisitn_rs10.RowCount > 0 Then
'    prpathno = ""
'    CMB_PATHNO.clear
'    While Not seldrequisitn_rs10.EOF
'        prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
'        CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
'        seldrequisitn_rs10.MoveNext
'    Wend
'    txt_prepathreqno.Text = prpathno
'End If
'If SELDADDENDUM_FLAG = False Then
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'    SELDADDENDUM_FLAG = True
'Else
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    SELDADDENDUM_RS.MoreResults
'    SELDADDENDUM_RS.Requery
'End If

For i = 0 To 25
    CRYSTALREPORT1.Formulas(i) = ""
Next i

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"

If pathology.opt_patient.Value = True Then
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CRYSTALREPORT1.Formulas(2) = "ENT = '" & str1 & "'"
             Else
    
                CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CRYSTALREPORT1.Formulas(2) = "ENT = '" & str1 & "'"
            End If
            CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
        Else
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
        End If
        'CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.Action = 1
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CRYSTALREPORT1.Formulas(7) = "USRID = '" & str & "'"
        Else
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CRYSTALREPORT1.Formulas(7) = "USRID = '" & str & "'"
        End If
        CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
        CRYSTALREPORT1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.Action = 1
    End If
End Sub
Private Sub CGCNREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
SELPATHDREQUISITN_QR(0) = VAR_REQNNO
Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
'seldrequisitn_qr10.QueryTimeout = 300
'If seldrequisitn_flag10 = False Then
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
'    seldrequisitn_flag10 = True
'Else
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    seldrequisitn_rs10.MoreResults
'    seldrequisitn_rs10.Requery
'End If
'SELPATHDREQUISITN_QR(0) = VAR_REQNNO
'Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
'
'If seldrequisitn_rs10.RowCount > 0 Then
'    prpathno = ""
'    CMB_PATHNO.clear
'    While Not seldrequisitn_rs10.EOF
'        prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
'        CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
'        seldrequisitn_rs10.MoveNext
'    Wend
'    txt_prepathreqno.Text = prpathno
'End If
'If SELDADDENDUM_FLAG = False Then
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'    SELDADDENDUM_FLAG = True
'Else
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    SELDADDENDUM_RS.MoreResults
'    SELDADDENDUM_RS.Requery
'End If

'SELDADDENDUM_QR(0) = VAR_REQNNO
'SELDADDENDUM_RS.MoreResults
'SELDADDENDUM_RS.Requery

For i = 0 To 25
    CRYSTALREPORT1.Formulas(i) = ""
Next i

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"

If pathology.opt_patient.Value = True Then
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CRYSTALREPORT1.Formulas(7) = "USRID = '" & str & "'"
        CRYSTALREPORT1.Formulas(12) = "NOSLIDES=" & "" & selpathdrequisitn_rs("LABNO")
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) <> UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CRYSTALREPORT1.Formulas(11) = "P= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CRYSTALREPORT1.Formulas(12) = "F= '" & Trim(SELRREQNORS("FREPORTBY")) & "'"
            CRYSTALREPORT1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
            CRYSTALREPORT1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) = "" Then
            CRYSTALREPORT1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
            CRYSTALREPORT1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) = UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CRYSTALREPORT1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CRYSTALREPORT1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
    Else
        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CRYSTALREPORT1.Formulas(7) = "USRID = '" & str & "'"
        CRYSTALREPORT1.Formulas(12) = "NOSLIDES = " & "" & selpathdrequisitn_rs("LABNO")
    End If
Else
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CRYSTALREPORT1.Formulas(7) = "USRID = '" & str & "'"
    Else
        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CRYSTALREPORT1.Formulas(7) = "USRID = '" & str & "'"
    End If
    CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
    CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
    CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
End If
CRYSTALREPORT1.WindowState = crptMaximized
'CRYSTALREPORT1.WindowState = crptMaximized
CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CRYSTALREPORT1.Action = 1
End Sub

Private Sub BONEMARROW()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'''********BONE MARROW*************
For i = 0 To 25
    CRYSTALREPORT1.Formulas(i) = ""
Next i

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

   If SELDBMAREPORT_FLAG = False Then
        SELDBMAREPORT_QR(0) = VAR_REQNNO
        Set SELDBMAREPORT_RS = SELDBMAREPORT_QR.OpenResultset(1, 3)
        SELDBMAREPORT_FLAG = True
    Else
        SELDBMAREPORT_QR(0) = VAR_REQNNO
        SELDBMAREPORT_RS.MoreResults
        SELDBMAREPORT_RS.Requery
    End If
    If pathology.opt_patient.Value = True Then
        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBM2.RPT"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Not selta1012_rs.EOF Then
            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
        End If
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & SELRREQNORS("ENTERBY")
            If Len(Trim(SELRREQNORS("ENTERBY"))) > 0 Then
                CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CRYSTALREPORT1.Formulas(12) = "USRID = '" & str & "'"
        If IsNull(SELDBMAREPORT_RS("CD2")) And IsNull(SELDBMAREPORT_RS("CD3")) And IsNull(SELDBMAREPORT_RS("CD4")) And IsNull(SELDBMAREPORT_RS("CD5")) And IsNull(SELDBMAREPORT_RS("CD7")) And IsNull(SELDBMAREPORT_RS("CD8")) And IsNull(SELDBMAREPORT_RS("CYTOCD3")) Then
            CRYSTALREPORT1.Formulas(13) = "tmarkers = '" & "" & "'"
        Else
            CRYSTALREPORT1.Formulas(13) = "tmarkers = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("CD10")) And IsNull(SELDBMAREPORT_RS("CD19")) And IsNull(SELDBMAREPORT_RS("FMC7")) And IsNull(SELDBMAREPORT_RS("CD22")) And IsNull(SELDBMAREPORT_RS("CD23")) And IsNull(SELDBMAREPORT_RS("smig")) And IsNull(SELDBMAREPORT_RS("CD79B")) And IsNull(SELDBMAREPORT_RS("CD25")) And IsNull(SELDBMAREPORT_RS("CD11C")) And IsNull(SELDBMAREPORT_RS("cd103")) And IsNull(SELDBMAREPORT_RS("kappa")) And IsNull(SELDBMAREPORT_RS("lamda")) And IsNull(SELDBMAREPORT_RS("DBA44")) And IsNull(SELDBMAREPORT_RS("cytoEP22")) And IsNull(SELDBMAREPORT_RS("cytocd79a")) And IsNull(SELDBMAREPORT_RS("CYTOMU")) Then
            CRYSTALREPORT1.Formulas(14) = "bmarkers = '" & "" & "'"
        Else
            CRYSTALREPORT1.Formulas(14) = "bmarkers = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("CD34")) And IsNull(SELDBMAREPORT_RS("HLADR")) And IsNull(SELDBMAREPORT_RS("CD38")) And IsNull(SELDBMAREPORT_RS("TDT")) And IsNull(SELDBMAREPORT_RS("CD138")) And IsNull(SELDBMAREPORT_RS("ANTIMPO")) Then
            CRYSTALREPORT1.Formulas(15) = "others = '" & "" & "'"
        Else
                CRYSTALREPORT1.Formulas(15) = "others = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("cd13")) And IsNull(SELDBMAREPORT_RS("Cd14")) And IsNull(SELDBMAREPORT_RS("CD15")) And IsNull(SELDBMAREPORT_RS("CD33")) And IsNull(SELDBMAREPORT_RS("cd36")) And IsNull(SELDBMAREPORT_RS("cd4142")) And IsNull(SELDBMAREPORT_RS("Glycophoin")) And IsNull(SELDBMAREPORT_RS("CD68")) And IsNull(SELDBMAREPORT_RS("CD42")) And IsNull(SELDBMAREPORT_RS("CD117")) Then
           CRYSTALREPORT1.Formulas(16) = "mm = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(16) = "mm = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("IRONSTAIN")) And IsNull(SELDBMAREPORT_RS("GIDEROBLAS")) And IsNull(SELDBMAREPORT_RS("sudanblack")) And IsNull(SELDBMAREPORT_RS("TOLUIDINEB")) Then
           CRYSTALREPORT1.Formulas(17) = "cci = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(17) = "cci = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("myelopero")) And IsNull(SELDBMAREPORT_RS("nonspester")) And IsNull(SELDBMAREPORT_RS("trap")) Then
           CRYSTALREPORT1.Formulas(18) = "ccm = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(18) = "ccm = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("IMPRINT")) And IsNull(SELDBMAREPORT_RS("abcellblas")) Then
           CRYSTALREPORT1.Formulas(19) = "im = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(19) = "im = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("mcellulari")) And IsNull(SELDBMAREPORT_RS("morphology")) Then
           CRYSTALREPORT1.Formulas(20) = "ms = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(20) = "ms = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("ecellulari")) And IsNull(SELDBMAREPORT_RS("maturation")) Then
           CRYSTALREPORT1.Formulas(21) = "es = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(21) = "es = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("mature")) And IsNull(SELDBMAREPORT_RS("immature")) And IsNull(SELDBMAREPORT_RS("atypbinu")) And IsNull(SELDBMAREPORT_RS("plasmablas")) Then
           CRYSTALREPORT1.Formulas(22) = "pc = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(22) = "pc = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("MLYMPHOCYT")) And IsNull(SELDBMAREPORT_RS("ALYMPHOCYT")) And IsNull(SELDBMAREPORT_RS("Plasmacell")) Then
           CRYSTALREPORT1.Formulas(23) = "ls = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(23) = "ls = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("PROMYELOCY")) And IsNull(SELDBMAREPORT_RS("PROMYELOCY")) And IsNull(SELDBMAREPORT_RS("MMYELOCYTE")) And IsNull(SELDBMAREPORT_RS("polymorphs")) And IsNull(SELDBMAREPORT_RS("basophils")) And IsNull(SELDBMAREPORT_RS("eosinophil")) And IsNull(SELDBMAREPORT_RS("monocytes")) And IsNull(SELDBMAREPORT_RS("PROMONOCYT")) Then
           CRYSTALREPORT1.Formulas(24) = "mys = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(24) = "mys = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("ECELLULARI")) And IsNull(SELDBMAREPORT_RS("MERATIO")) Then
           CRYSTALREPORT1.Formulas(25) = "mor = '" & "" & "'"
        Else
           CRYSTALREPORT1.Formulas(25) = "mor = '" & "abc" & "'"
        End If
    End If
    CRYSTALREPORT1.WindowState = crptMaximized
    'CRYSTALREPORT1.WindowState = crptMaximized
    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CRYSTALREPORT1.Action = 1


End Sub
Private Sub BFREPORT()
For i = 0 To 25
    CRYSTALREPORT1.Formulas(i) = ""
Next i
selta1012_qr(0) = msk_caseno
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
If pathology.opt_patient.Value = True Then
             
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

    
    
     If Trim(SELRREQNORS("MODALITYCD")) = "BF" Then
         CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBF.RPT"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            If Len(Trim(SELRREQNORS("ENTERBY"))) > 0 Then
                CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else

            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CRYSTALREPORT1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If
       
End If
If pathology.opt_staff.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

 
     If Trim(SELRREQNORS("MODALITYCD")) = "BF" Then
         CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBFS.RPT"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & VAR_REQNNO & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CRYSTALREPORT1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If

End If
'***** STAFF
CRYSTALREPORT1.WindowState = crptMaximized
CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CRYSTALREPORT1.Action = 1
End Sub
Private Sub HLREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If

If SELDCBCREPORT_FLAG = False Then
    SELDCBCREPORT_QR(0) = VAR_REQNNO
    Set SELDCBCREPORT_RS = SELDCBCREPORT_QR.OpenResultset(1, 3)
    SELDCBCREPORT_FLAG = True
Else
    SELDCBCREPORT_QR(0) = VAR_REQNNO
    SELDCBCREPORT_RS.MoreResults
    SELDCBCREPORT_RS.Requery
End If
For i = 0 To 25
    CRYSTALREPORT1.Formulas(i) = ""
Next i
selta1012_qr(0) = SELRREQNORS("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

If pathology.opt_patient.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    'If Trim(txt_modcode.Text) = "HL" Then
    CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DIShl.RPT"
    CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'       If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'       Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'       End If
    CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
    If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
        CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        str1 = ""
        str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
        If Len(Trim(SELRREQNORS("enterby"))) > 0 Then
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
        End If
    Else
        CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
    End If
    CRYSTALREPORT1.Formulas(12) = "USRID = '" & str & "'"
             If Trim(SELDCBCREPORT_RS("lap")) = "" And Trim(SELDCBCREPORT_RS("MPO")) = "" And Trim(SELDCBCREPORT_RS("SBB")) = "" And Trim(SELDCBCREPORT_RS("PSA")) = "" And Trim(SELDCBCREPORT_RS("NSE")) = "" And Trim(SELDCBCREPORT_RS("TRAD")) = "" Then
                'If CYTOCHEMFLAG = True Then
                CRYSTALREPORT1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
                'End If
            End If
            If Trim(SELDCBCREPORT_RS("BT")) = "" And Trim(SELDCBCREPORT_RS("PT")) = "" And Trim(SELDCBCREPORT_RS("PLATELET1")) = "" And Trim(SELDCBCREPORT_RS("PTTK")) = "" And Trim(SELDCBCREPORT_RS("TT")) = "" And Trim(SELDCBCREPORT_RS("FIBRINOGEN")) = "" And Trim(SELDCBCREPORT_RS("FDPDDIMER")) = "" And Trim(SELDCBCREPORT_RS("OTHERS")) = "" Then
                CRYSTALREPORT1.Formulas(11) = "CP = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(11) = "CP = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("NORMBOBLAS")) = "" And Trim(SELDCBCREPORT_RS("OTHERCELLR")) = "" And Trim(SELDCBCREPORT_RS("rbcmorpho")) = "" And Trim(SELDCBCREPORT_RS("reticcount")) = "" And Trim(SELDCBCREPORT_RS("mparasite")) = "" And Trim(SELDCBCREPORT_RS("ESR")) = "" Then
                CRYSTALREPORT1.Formulas(15) = "RBC = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(15) = "RBC = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PMORPHOLOG")) = "" And Trim(SELDCBCREPORT_RS("FLD1")) = "" Then
                CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PROLYMPHOC")) = "" And Trim(SELDCBCREPORT_RS("Sezarycell")) = "" And Trim(SELDCBCREPORT_RS("hairycells")) = "" And Trim(SELDCBCREPORT_RS("othercells")) = "" Then
                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("blasts")) = "" And Trim(SELDCBCREPORT_RS("PROMYELOCY")) = "" And Trim(SELDCBCREPORT_RS("myelocytes")) = "" And Trim(SELDCBCREPORT_RS("metamyeloc")) = "" And Trim(SELDCBCREPORT_RS("plasmacell")) = "" And Trim(SELDCBCREPORT_RS("aTlymphEE")) = "" Then
                CRYSTALREPORT1.Formulas(18) = "blast = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(18) = "blast = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophi1")) = "" And Trim(SELDCBCREPORT_RS("lymphocyt1")) = "" And Trim(SELDCBCREPORT_RS("monocytE1")) = "" And Trim(SELDCBCREPORT_RS("eosinophi1")) = "" And Trim(SELDCBCREPORT_RS("basophill1")) = "" Then
                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HCT")) = "" And Trim(SELDCBCREPORT_RS("MCH")) = "" And Trim(SELDCBCREPORT_RS("MCV")) = "" And Trim(SELDCBCREPORT_RS("MCHC")) = "" And Trim(SELDCBCREPORT_RS("MPV")) = "" And Trim(SELDCBCREPORT_RS("RDW")) = "" And Trim(SELDCBCREPORT_RS("ABNORMAL")) = "" Then
                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophil")) = "" And Trim(SELDCBCREPORT_RS("lymphocyte")) = "" And Trim(SELDCBCREPORT_RS("monocytes")) = "" And Trim(SELDCBCREPORT_RS("eosinophil")) = "" And Trim(SELDCBCREPORT_RS("basophils")) = "" And Trim(SELDCBCREPORT_RS("immgranulo")) = "" Then
                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HB")) = "" And Trim(SELDCBCREPORT_RS("RBC")) = "" And Trim(SELDCBCREPORT_RS("TLC")) = "" And Trim(SELDCBCREPORT_RS("platelets")) = "" Then
                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "" & "'"
            Else
                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "123" & "'"
            End If
End If
If pathology.opt_staff.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    If Trim(txt_modcode.Text) = "EK" Then
        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISECS.RPT"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'           If Trim(prepath) = "" Then
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'           Else
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CRYSTALREPORT1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CRYSTALREPORT1.Formulas(12) = "USRID = '" & str & "'"
    End If
End If
'***** STAFF
CRYSTALREPORT1.WindowState = crptMaximized
CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CRYSTALREPORT1.Action = 1

End Sub

