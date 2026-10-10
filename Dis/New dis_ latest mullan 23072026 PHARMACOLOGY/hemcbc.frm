VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form hemcbc 
   Caption         =   "Haematology Requisition"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16275
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
   ScaleHeight     =   11040
   ScaleWidth      =   16275
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox TreeView1 
      Height          =   735
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   78
      Top             =   8640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox CHK_RETICCOUNT 
      Caption         =   "Reticulocyte Count"
      Height          =   255
      Left            =   11160
      TabIndex        =   74
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CheckBox CHK_ESR 
      Caption         =   "Erythrocyte Sadimentation Rate (ESR)"
      Height          =   255
      Left            =   11160
      TabIndex        =   69
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox txt_cyto_other 
         Height          =   285
         Left            =   1320
         TabIndex        =   67
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox chk_cytochemistry 
         Caption         =   "Cytochemistry on peripheral Blood"
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   3255
      End
      Begin VB.CheckBox CHK_TRAP1 
         Caption         =   "TRAP"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox CHK_NSE 
         Caption         =   "NSE"
         Height          =   195
         Left            =   4080
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_PAS 
         Caption         =   "PAS"
         Height          =   195
         Left            =   6600
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_SBB 
         Caption         =   "SBB"
         Height          =   195
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Other"
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   1560
         Width           =   735
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   240
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "TRAP"
         Height          =   255
         Left            =   6720
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "PAS"
         Height          =   255
         Left            =   7800
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "NSE"
         Height          =   255
         Left            =   5160
         TabIndex        =   50
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "SBB"
         Height          =   255
         Left            =   5160
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CheckBox CHK_MORPHOLOGYDLC1 
      Caption         =   "Haemogram,DLC,Morphology "
      Height          =   195
      Left            =   12855
      TabIndex        =   54
      Top             =   1275
      Visible         =   0   'False
      Width           =   2865
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   720
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
      Left            =   4920
      TabIndex        =   42
      Top             =   1080
      Width           =   5175
   End
   Begin VB.CheckBox CHK_OUTSIDESAMPLE1 
      Caption         =   "Outside sample"
      Height          =   255
      Left            =   12855
      TabIndex        =   40
      Top             =   1035
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox CHK_RETICCOUNT1 
      Caption         =   "Retic count"
      Height          =   255
      Left            =   12855
      TabIndex        =   39
      Top             =   555
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox CHK_MALARIALPARASITE1 
      Caption         =   "Malarial Parasite"
      Height          =   255
      Left            =   12855
      TabIndex        =   38
      Top             =   795
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CheckBox CHK_ESR1 
      Caption         =   "ESR"
      Height          =   255
      Left            =   12840
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox CMB_PREINVEST 
      BackColor       =   &H008080FF&
      Height          =   315
      Left            =   3360
      TabIndex        =   30
      Top             =   9360
      Width           =   4095
   End
   Begin VB.Frame Frame7 
      Height          =   1695
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   12975
      Begin VB.CheckBox chk_Heaparin 
         Caption         =   "On Heaparin"
         Height          =   255
         Left            =   6000
         TabIndex        =   73
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chk_Warffine 
         Caption         =   "On Warffine"
         Height          =   375
         Left            =   6000
         TabIndex        =   72
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox CHK_PLATELETS1 
         Caption         =   "Platelets"
         Height          =   2295
         Left            =   6840
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox CHK_CP 
         Caption         =   "Coagulation Profile"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   15
         TabIndex        =   29
         Top             =   15
         Width           =   1935
      End
      Begin VB.CheckBox CHK_PT 
         Caption         =   "Prothrombin Time(PT)/International Normalised Ratio (INR)"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   5655
      End
      Begin VB.CheckBox CHK_PTTK 
         Caption         =   "Activated Partial Thromboplastin Time (APTT)"
         Height          =   195
         Left            =   6000
         TabIndex        =   24
         Top             =   360
         Width           =   5535
      End
      Begin VB.CheckBox CHK_TT 
         Caption         =   "TT"
         Height          =   2115
         Left            =   5295
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox CHK_FDPDDIMER 
         Caption         =   " ( D- DIMER)"
         Height          =   195
         Left            =   2520
         TabIndex        =   22
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox CHK_FIBRINOGEN 
         Caption         =   "Fibrinogen"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox CHK_OTHERS 
         Caption         =   "History"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1095
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   1095
         Width           =   2415
      End
      Begin VB.CheckBox CHK_BT 
         Caption         =   "BT"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1815
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   10815
      Begin VB.CheckBox chk_CBCmplt 
         Caption         =   "Manual Platelet Count"
         Height          =   255
         Left            =   2760
         TabIndex        =   71
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chk_mdc 
         Caption         =   "Manual Differential Count"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chk_cbc 
         Caption         =   "Complete Blood Count(CBC)"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   0
         Width           =   2895
      End
      Begin VB.CheckBox CHK_MORPHOLOGYDLC 
         Caption         =   "Haemogram,DLC,Morphology "
         Height          =   195
         Left            =   8640
         TabIndex        =   65
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox CHK_OUTSIDESAMPLE 
         Caption         =   "Outside sample"
         Height          =   195
         Left            =   8640
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox CHK_MALARIALPARASITE 
         Caption         =   "Malarial Parasite"
         Height          =   255
         Left            =   8640
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox CHK_RETICCOUNTZ 
         Caption         =   "Reticulocyte Count"
         Height          =   255
         Left            =   8640
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox CHK_ESR2 
         Caption         =   "ESR"
         Height          =   255
         Left            =   8640
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_RBCINDICES 
         Caption         =   "RBC Indices"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1560
         TabIndex        =   55
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox CHK_HAEMOGRAM 
         Caption         =   "Peripheral Blood Smear (PBS)"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3120
         TabIndex        =   41
         Top             =   0
         Width           =   3015
      End
      Begin VB.Frame frame_SubPBS 
         Height          =   1455
         Left            =   3120
         TabIndex        =   35
         Top             =   120
         Width           =   5415
         Begin VB.CheckBox CHK_LAP 
            Caption         =   "LAP"
            Height          =   255
            Left            =   4440
            TabIndex        =   77
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chk_mpo 
            Caption         =   "MPO"
            Height          =   435
            Left            =   3480
            TabIndex        =   76
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txt_other 
            Height          =   285
            Left            =   720
            TabIndex        =   59
            Top             =   960
            Width           =   3255
         End
         Begin VB.CheckBox chk_MPLT 
            Caption         =   "MPLT"
            Height          =   255
            Left            =   2160
            TabIndex        =   58
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chk_mp 
            Caption         =   "Malarial parasites"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox CHK_RBC 
            Caption         =   "RBC Morphology"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox CHK_DIFFCOUNT 
            Caption         =   "MDC"
            Height          =   195
            Left            =   2160
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Cytochemistry"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3600
            TabIndex        =   75
            Top             =   120
            Width           =   1455
         End
         Begin VB.Shape Shape2 
            Height          =   615
            Left            =   3360
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lbl_Other 
            Caption         =   "Other"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.CheckBox CHK_HB 
         Caption         =   "Hb"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox CHK_TLC 
         Caption         =   "TLC"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox CHK_PLATELETS 
         Caption         =   "Platelets"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<< &Back"
      Height          =   495
      Left            =   9630
      TabIndex        =   10
      Top             =   8100
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   1560
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CheckBox Check12 
         Caption         =   "Peripheral Blood Smear / Slides for Surface markers"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Peripheral Blood Smear / Slides for morphology"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   600
         Width           =   4815
      End
      Begin VB.CheckBox Check24 
         Caption         =   "Outside Specimen"
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txt_pbos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txt_smpb 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Others"
         Height          =   195
         Left            =   5880
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl_smpb 
         AutoSize        =   -1  'True
         Caption         =   "Surface Markers on Peripheral Blood"
         Height          =   195
         Left            =   5040
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lbl_pbos 
         AutoSize        =   -1  'True
         Caption         =   "Peripheral Blood Outside Slides"
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   2685
      End
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   2160
      TabIndex        =   43
      Top             =   1080
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
   Begin VB.Frame Frame1 
      Caption         =   "Surface Marker"
      Height          =   1215
      Left            =   240
      TabIndex        =   26
      Top             =   7200
      Width           =   9135
      Begin VB.TextBox txtSurfaceMarker 
         Height          =   285
         Left            =   360
         MaxLength       =   150
         TabIndex        =   79
         Top             =   720
         Width           =   8415
      End
      Begin VB.CheckBox chk_sm 
         Caption         =   "Surface Marker - Single"
         Height          =   195
         Left            =   2505
         TabIndex        =   48
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox CHK_SMC 
         Caption         =   "Surface Marker - Complete"
         Height          =   195
         Left            =   5295
         TabIndex        =   47
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox CHK_SMPERIPHERALBLOOD 
         Caption         =   "Surface Markers"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin VB.Label Label9 
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
      TabIndex        =   53
      Top             =   0
      Width           =   11655
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
      Left            =   1200
      TabIndex        =   45
      Top             =   1080
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
      Left            =   4200
      TabIndex        =   44
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Previous Investigation"
      Height          =   195
      Left            =   1320
      TabIndex        =   31
      Top             =   9480
      Width           =   1905
   End
   Begin VB.Label lbl_cbc 
      AutoSize        =   -1  'True
      Caption         =   "HAEMATOLOGY - REQUISITION"
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
      Left            =   3960
      TabIndex        =   11
      Top             =   480
      Width           =   4485
   End
End
Attribute VB_Name = "hemcbc"
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

Dim chk As Boolean
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

Private Sub chk_cbc_Click()
    
If chk_cbc.Value = 1 Then
    frame_SubPBS.Visible = False
    CHK_HAEMOGRAM.Value = 0
    CHK_HAEMOGRAM.Visible = False
    CHK_ESR.Value = 0
    CHK_ESR.Visible = False
    CHK_RETICCOUNT.Value = 0
    CHK_RETICCOUNT.Visible = False
'   If CHK_HAEMOGRAM.Value = 1 Then
'      MsgBox " You can select either CBC or HEMOGRAM "
'      Exit Sub
'    End If
'    CHK_HB.Value = 1
'    CHK_TLC.Value = 1
'    CHK_PLATELETS.Value = 1
'    CHK_RBCINDICES.Value = 1
    
    
     Frame7.Enabled = False
    chk_pt.Enabled = False
    CHK_FIBRINOGEN.Enabled = False
    CHK_OTHERS.Enabled = False
    CHK_PTTK.Enabled = False
    CHK_TT.Enabled = False
    CHK_FDPDDIMER.Enabled = False
    CHK_PLATELETS1.Enabled = False
     CHK_CP.Enabled = False
     

    
     
     '***** Shraddha 27/09/2011 *****
'     CHK_HB.Enabled = False
'     CHK_TLC.Enabled = False
'     CHK_PLATELETS.Enabled = False
'     CHK_RBCINDICES.Enabled = False
     '*******************************
     
Else
'    CHK_HB.Value = 0
'    CHK_TLC.Value = 0
'    CHK_PLATELETS.Value = 0
'    CHK_RBCINDICES.Value = 0
'
    CHK_RETICCOUNT.Visible = True
    CHK_ESR.Visible = True
    frame_SubPBS.Visible = True
    CHK_HAEMOGRAM.Visible = True
     Frame7.Enabled = True
    chk_pt.Enabled = True
    CHK_FIBRINOGEN.Enabled = True
    CHK_OTHERS.Enabled = True
    CHK_PTTK.Enabled = True
    CHK_TT.Enabled = True
    CHK_FDPDDIMER.Enabled = True
    CHK_PLATELETS1.Enabled = True
     CHK_CP.Enabled = True
     
     '***** Shraddha 27/09/2011 *****
'     CHK_HB.Enabled = True
'     CHK_TLC.Enabled = True
'     CHK_PLATELETS.Enabled = True
'     CHK_RBCINDICES.Enabled = True
     '*******************************
End If

'    'CHK_RBCINDICES.Value = 1
'    'CHK_MORPHOLOGYDLC.Value = 1
'  CHK_HAEMOGRAM.Value = 0
'
'Else
'   If CHK_HAEMOGRAM.Value = 1 Then
'
'    CHK_HB.Value = 1
'    CHK_TLC.Value = 1
'    CHK_PLATELETS.Value = 1
'    CHK_RBCINDICES.Value = 1
'    CHK_MORPHOLOGYDLC.Value = 1
'    CHK_DIFFCOUNT.Value = 1
'   Else
'     CHK_HB.Value = 0
'    CHK_TLC.Value = 0
'    CHK_PLATELETS.Value = 0
'    CHK_RBCINDICES.Value = 0
'    CHK_MORPHOLOGYDLC.Value = 0
'    CHK_DIFFCOUNT.Value = 0
'    'CHK_RBCINDICES.Value = 0
'    'CHK_MORPHOLOGYDLC.Value = 1
' End If
'
'End If
    'CHK_RBCINDICES.Value = 0
  '  CHK_MORPHOLOGYDLC.Value = 0
    'CHK_DIFFCOUNT.Value = 0
    
   
If chk_cbc.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    'Clear all other checkboxes and show other frames
   '***** Shraddha 27/09/2011 *****
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
        End If
    Next
    '******************************
End If
End Sub

Private Sub chk_CBCmplt_Click()
If chk_CBCmplt.Value = 1 Then
        Frame7.Visible = False
        Frame2.Visible = False
        Frame1.Visible = False
    Else
        Frame7.Visible = True
        'Frame2.Visible = True
        Frame1.Visible = True
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Enabled = True
            End If
        Next
    End If
End Sub

Private Sub CHK_CP_Click()
Dim c As Object


If CHK_CP.Value = 1 Then
    chk_pt.Value = 1
    CHK_PTTK.Value = 1
    'CHK_BT.Value = 1
    'CHK_PLATELETS1.Value = 1
    'CHK_TT.Value = 1

For Each c In Controls
    If TypeOf c Is CheckBox Then
        c.Enabled = False
        Frame7.Enabled = True
        chk_pt.Enabled = True
        CHK_FIBRINOGEN.Enabled = True
        CHK_OTHERS.Enabled = True
        CHK_PTTK.Enabled = True
        CHK_TT.Enabled = True
        CHK_FDPDDIMER.Enabled = True
        CHK_PLATELETS1.Enabled = True
        CHK_CP.Enabled = True
    End If
    
Next

Else
    chk_pt.Value = 0
    CHK_PTTK.Value = 0
    
  If chk_cbc.Value = 1 Then
    
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
            
            CHK_HB.Enabled = False
            CHK_TLC.Enabled = False
            CHK_PLATELETS.Enabled = False
            CHK_RBCINDICES.Enabled = False
            CHK_DIFFCOUNT.Enabled = False
            
            Frame7.Enabled = False
            chk_pt.Enabled = False
            CHK_FIBRINOGEN.Enabled = False
            CHK_OTHERS.Enabled = False
            CHK_PTTK.Enabled = False
            CHK_TT.Enabled = False
            CHK_FDPDDIMER.Enabled = False
            CHK_PLATELETS1.Enabled = False
            CHK_CP.Enabled = False
        End If
    Next
    
  Else
    
        For Each c In Controls
        
            If TypeOf c Is CheckBox Then
                c.Enabled = True
            End If
        Next
        
  End If
    'CHK_BT.Value = 0
    'CHK_PLATELETS1.Value = 0
    'CHK_TT.Value = 0
End If
chk_Warffine.Enabled = True
chk_Heaparin.Enabled = True
If CHK_CP.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Clear all other checkboxes and show other frames
   '***** Shraddha 27/09/2011 *****
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
        End If
    Next
    '******************************
End If

End Sub

Private Sub chk_cytochemistry_Click()

If chk_cytochemistry.Value = 1 Then
EnableCloseButton Me.hwnd, False
    chk_mpo.Visible = True
    CHK_LAP.Visible = True
    Shape1.Visible = True
    'chk_mpo.Value = 1
'    CHK_SBB.Value = 1
'    CHK_PAS.Value = 1
    'CHK_LAP.Value = 1
'    CHK_NSE.Value = 1
'    CHK_TRAP.Value = 1
Frame7.Visible = False
Frame5.Visible = False
Frame1.Visible = False
Else
EnableCloseButton Me.hwnd, True
    chk_mpo.Value = 0
'    CHK_SBB.Value = 0
'    CHK_PAS.Value = 0
    CHK_LAP.Value = 0
'    CHK_NSE.Value = 0
'    CHK_TRAP.Value = 0
'    chk_mpo.Visible = False
'    CHK_LAP.Visible = False
'    Shape1.Visible = False
    
   'Clear all other checkboxes and show other frames
   '***** Shraddha 27/09/2011 *****
    Frame7.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
    '******************************
End If
End Sub

Private Sub CHK_DIFFCOUNT_Click()
'''If CHK_DIFFCOUNT.Value = 1 Then
''    CHK_HAEMOGRAM.Value = 1
''    chk_hb.Value = 1
''    CHK_TLC.Value = 1
''    CHK_PLATELETS.Value = 1
''    CHK_RBCINDICES.Value = 1
''Else
''    CHK_HAEMOGRAM.Value = 0
''    chk_hb.Value = 0
''    CHK_TLC.Value = 0
''    CHK_PLATELETS.Value = 0
''    CHK_RBCINDICES.Value = 0
''End If
If CHK_HAEMOGRAM.Value = 1 Then
If CHK_DIFFCOUNT.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    CHK_DIFFCOUNT.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_ESR_Click()
If CHK_ESR.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    Frame5.Visible = False
    CHK_RETICCOUNT.Visible = False
Else
    CHK_RETICCOUNT.Visible = True
    'Clear all other checkboxes and show other frames
   '***** Shraddha 27/09/2011 *****
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    Frame5.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
    '******************************
End If
End Sub

Private Sub CHK_FDPDDIMER_Click()
If CHK_FDPDDIMER.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If

If CHK_CP.Value = 1 Then
    CHK_FIBRINOGEN.Value = 1
Else
    If CHK_FDPDDIMER.Value = 1 Then
        CHK_FDPDDIMER.Value = 1
    Else
        CHK_FDPDDIMER.Value = 0
    End If
    CHK_FIBRINOGEN.Value = 0
End If

End Sub

Private Sub CHK_FIBRINOGEN_Click()

If CHK_FIBRINOGEN.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If

If CHK_CP.Value = 1 Then
    CHK_FDPDDIMER.Value = 1
Else
    If CHK_FIBRINOGEN.Value = 1 Then
        CHK_FIBRINOGEN.Value = 1
    Else
        CHK_FIBRINOGEN.Value = 0
    End If
    CHK_FDPDDIMER.Value = 0
End If


End Sub

Private Sub CHK_HAEMOGRAM_Click()

If CHK_HAEMOGRAM.Value = 1 Then
''    If chk_cbc.Value = 1 Then
''      MsgBox " You can select either CBC or HEMOGRAM "
''      Exit Sub
''    End If
'    chk_cbc.Value = 0
'
'    CHK_HB.Value = 1
'    CHK_TLC.Value = 1
'    CHK_PLATELETS.Value = 1
'    CHK_RBCINDICES.Value = 1
'   ' CHK_MORPHOLOGYDLC.Value = 1
'    CHK_DIFFCOUNT.Value = 1
'    'CHK_MALARIALPARASITE.Value = 1
'    CHK_RBC.Value = 1
    'Added by Shraddha on 26/09/2011
    EnableCloseButton Me.hwnd, False
     frame_SubPBS.Visible = True
     
     CHK_DIFFCOUNT.Enabled = True
     chk_mp.Enabled = True
     CHK_RBC.Enabled = True
     chk_MPLT.Enabled = True


Else
EnableCloseButton Me.hwnd, True
'    CHK_HB.Value = 0
'    CHK_TLC.Value = 0
'    CHK_PLATELETS.Value = 0
'    CHK_RBCINDICES.Value = 0
  '  CHK_MORPHOLOGYDLC.Value = 0
    CHK_DIFFCOUNT.Value = 0
    CHK_MALARIALPARASITE.Value = 0
    CHK_RBC.Value = 0
    chk_mp.Value = 0
End If

'    chk_cbc.Value = 0
'
'Else
'   If chk_cbc.Value = 1 Then
'     CHK_HB.Value = 1
'     CHK_TLC.Value = 1
'     CHK_PLATELETS.Value = 1
'   Else
'    CHK_HB.Value = 0
'    CHK_TLC.Value = 0
'    CHK_PLATELETS.Value = 0
'    CHK_RBCINDICES.Value = 0
'    CHK_MORPHOLOGYDLC.Value = 0
'    CHK_DIFFCOUNT.Value = 0
'   End If
'
'End If

If CHK_HAEMOGRAM.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
Else
    'Clear all other checkboxes and show other frames
   '***** Shraddha 27/09/2011 *****
    Frame7.Visible = True
    'Frame2.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
        End If
    Next
    '******************************
End If
End Sub

Private Sub chk_hb_Click()
If chk_cbc.Value = 1 Then
If CHK_HB.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    CHK_HB.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub chk_Heaparin_Click()
If CHK_CP.Value = 1 Then

    If chk_Heaparin.Value = 1 Then
        Frame2.Visible = False
        Frame5.Visible = False
        Frame1.Visible = False
    Else
        'Clear all other checkboxes and show other frames
       '***** Shraddha 27/09/2011 *****
        Frame7.Visible = True
        Frame5.Visible = True
        Frame1.Visible = True
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Enabled = True
            End If
        Next
        '******************************
    End If
Else
    chk_Heaparin.Value = 0
     'Clear all other checkboxes and show other frames
       '***** Shraddha 27/09/2011 *****
        Frame2.Visible = False
        Frame5.Visible = False
        Frame1.Visible = False
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Value = 0
                c.Enabled = True
            End If
        Next
        '******************************
    
End If
End Sub

Private Sub CHK_LAP_Click()
If CHK_LAP.Value = 1 Then
    CHK_DIFFCOUNT.Value = 1
End If

'**** below code on 12.04.2012 by v.s.k.
'If chk_cytochemistry.Value = 1 Then
'If CHK_LAP.Value = 1 Then
'    Frame7.Visible = False
'    Frame5.Visible = False
'    Frame1.Visible = False
'Else
'    'Clear all other checkboxes and show other frames
'   '***** Shraddha 27/09/2011 *****
'    Frame7.Visible = True
'    Frame5.Visible = True
'    Frame1.Visible = True
'    For Each c In Controls
'        If TypeOf c Is CheckBox Then
'            c.Enabled = True
'        End If
'    Next
'    '******************************
'End If
'Else
'    CHK_LAP.Value = 0
'    Frame7.Visible = False
'    Frame5.Visible = False
'    Frame1.Visible = False
'    For Each c In Controls
'        If TypeOf c Is CheckBox Then
'            c.Value = 0
'            c.Enabled = True
'        End If
'    Next
'End If
End Sub

Private Sub chk_mdc_Click()

    If chk_mdc.Value = 1 Then
        Frame7.Visible = False
        Frame2.Visible = False
        Frame1.Visible = False
    Else
        Frame7.Visible = True
        'Frame2.Visible = True
        Frame1.Visible = True
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Enabled = True
            End If
        Next
    End If

End Sub

Private Sub CHK_MORPHOLOGYDLC_Click()
'If CHK_MORPHOLOGYDLC.Value = 1 Then
'    CHK_HAEMOGRAM.Value = 1
'    CHK_HB.Value = 1
'    CHK_TLC.Value = 1
'    CHK_PLATELETS.Value = 1
'    CHK_DIFFCOUNT.Value = 1
'    CHK_RBCINDICES.Value = 1
'Else
'    CHK_HAEMOGRAM.Value = 0
'    CHK_HB.Value = 0
'    CHK_TLC.Value = 0
'    CHK_PLATELETS.Value = 0
'    CHK_DIFFCOUNT.Value = 0
'    CHK_RBCINDICES.Value = 0
'End If
End Sub


Private Sub chk_mp_Click()
If CHK_HAEMOGRAM.Value = 1 Then
If chk_mp.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    chk_mp.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub chk_MPLT_Click()
If CHK_HAEMOGRAM.Value = 1 Then
If chk_MPLT.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    chk_MPLT.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub chk_mpo_Click()
If chk_mpo.Value = 1 Then
    CHK_DIFFCOUNT.Value = 1
End If

'**** below code commented on 12-04-2012 by v.s.k.
'If chk_cytochemistry.Value = 1 Then
'
'    If chk_mpo.Value = 1 Then
'        Frame7.Visible = False
'        Frame5.Visible = False
'        Frame1.Visible = False
'    Else
'        'Clear all other checkboxes and show other frames
'       '***** Shraddha 27/09/2011 *****
'        Frame7.Visible = True
'        Frame5.Visible = True
'        Frame1.Visible = True
'        For Each c In Controls
'            If TypeOf c Is CheckBox Then
'                c.Enabled = True
'            End If
'        Next
'        '******************************
'    End If
'Else
'    chk_mpo.Value = 0
'     'Clear all other checkboxes and show other frames
'       '***** Shraddha 27/09/2011 *****
'        Frame7.Visible = False
'        Frame5.Visible = False
'        Frame1.Visible = False
'        For Each c In Controls
'            If TypeOf c Is CheckBox Then
'                c.Value = 0
'                c.Enabled = True
'            End If
'        Next
'        '******************************
'
'End If
End Sub

Private Sub CHK_OTHERS_Click()
If CHK_OTHERS.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
If CHK_OTHERS.Value = 1 Then
    'Text1.SetFocus
End If

End Sub

Private Sub CHK_PLATELETS_Click()
If chk_cbc.Value = 1 Then
If CHK_PLATELETS.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    CHK_PLATELETS.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_PLATELETS1_Click()
If CHK_PLATELETS1.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_PT_Click()
If chk_pt.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_PTTK_Click()
If CHK_PTTK.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_RBC_Click()
If CHK_HAEMOGRAM.Value = 1 Then
If CHK_RBC.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    CHK_RBC.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_RBCINDICES_Click()
If chk_cbc.Value = 1 Then
If CHK_RBCINDICES.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    CHK_RBCINDICES.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_RETICCOUNT_Click()
If CHK_RETICCOUNT.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    Frame5.Visible = False
    CHK_ESR.Visible = False
Else
    CHK_ESR.Visible = True
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    Frame5.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
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

Private Sub CHK_TLC_Click()
If chk_cbc.Value = 1 Then
If CHK_TLC.Value = 1 Then
    Frame7.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
Else
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
Else
    CHK_TLC.Value = 0
    Frame7.Visible = True
    'Frame2.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub CHK_TRAP1_Click()
If chk_cytochemistry.Value = 1 Then

    If CHK_TRAP1.Value = 1 Then
        Frame7.Visible = False
        Frame5.Visible = False
        Frame1.Visible = False
    Else
        'Clear all other checkboxes and show other frames
       '***** Shraddha 27/09/2011 *****
        Frame7.Visible = True
        Frame5.Visible = True
        Frame1.Visible = True
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Enabled = True
            End If
        Next
        '******************************
    End If
Else
    CHK_TRAP1.Value = 0
     'Clear all other checkboxes and show other frames
       '***** Shraddha 27/09/2011 *****
        Frame7.Visible = False
        Frame5.Visible = False
        Frame1.Visible = False
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Value = 0
                c.Enabled = True
            End If
        Next
        '******************************
    
End If
End Sub

Private Sub CHK_tt_Click()
If CHK_TT.Value = 1 Then
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = False
Else
    'Frame2.Visible = True
    Frame5.Visible = True
    Frame1.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
            c.Enabled = True
        End If
    Next
End If
End Sub

Private Sub chk_Warffine_Click()
If CHK_CP.Value = 1 Then

    If chk_Warffine.Value = 1 Then
        Frame2.Visible = False
        Frame5.Visible = False
        Frame1.Visible = False
    Else
        'Clear all other checkboxes and show other frames
       '***** Shraddha 27/09/2011 *****
        Frame7.Visible = True
        Frame5.Visible = True
        Frame1.Visible = True
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Enabled = True
            End If
        Next
        '******************************
    End If
Else
    chk_Warffine.Value = 0
     'Clear all other checkboxes and show other frames
       '***** Shraddha 27/09/2011 *****
        Frame2.Visible = False
        Frame5.Visible = False
        Frame1.Visible = False
        For Each c In Controls
            If TypeOf c Is CheckBox Then
                c.Value = 0
                c.Enabled = True
            End If
        Next
        '******************************
    
End If
End Sub



Private Sub CMB_PREINVEST_Click()
On Error GoTo Errhandler
Dim X
X = InStr(1, Trim(cmb_preinvest.Text), "$")
REQNO1 = Mid(Trim(cmb_preinvest.Text), 1, X - 1)
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
Exit Sub
End Sub

Private Sub cmd_back_Click()
Dim GRP1 As Boolean
Dim GRP2 As Boolean
'meterial = cmb_nmeterial.Text
pathology.CMB_TEXT.clear
ccount = 0
MPO = ""
SBB = ""
PAS = ""
LAP = ""
NSE = ""
TRAP = ""
cbc = " "
hem = " "
cg1 = " "
cyto = " "
smar = " "
TT = " "
oth = " "
othtext = " "

GRP1 = False
GRP2 = False

'***** Shraddha 1/12/2011 *****
chk = False
For Each c In Controls
If TypeOf c Is CheckBox Then
    If c.Value = 1 Then
        chk = True
    End If
End If
Next
If pathology.cmb_reqno.Visible = False Then
If chk = False Then
    MsgBox "Please select atleast 1 procedure"
    Exit Sub
End If
End If
'******************************

'Commented and added by Shraddha 06/05/2011
'If chk_mpo.Value = 1 Then
If chk_mpo.Value = 1 Or CHK_MALARIALPARASITE.Value = 1 Then
    MPO = "YES"
End If
If CHK_SBB.Value = 1 Then
    SSB = "YES"
End If
If CHK_PAS.Value = 1 Then
    PAS = "YES"
End If
If CHK_LAP.Value = 1 Then
    LAP = "YES"
End If
If CHK_NSE.Value = 1 Then
    NSE = "YES"
End If
If CHK_TRAP1.Value = 1 Then
    TRAP = "YES"
End If
If chk_cbc.Value = 1 Then
    cbc = "YES"
End If

If CHK_HAEMOGRAM.Value = 1 And CHK_MORPHOLOGYDLC.Value = 1 Then
       MsgBox "Both Haemogram,TLC,DLC,Paltelets and Haemogram ,TLC, DLC,Platelets, Morphology cannot be selected"
       Exit Sub
End If

'***** Shraddha 30/11/2011 *****
If CHK_CP.Value = 1 Then
    If Trim(Text1.Text) = "" Then
        MsgBox "Please enter History"
        Exit Sub
    End If
End If
'*******************************
       
If CHK_HAEMOGRAM.Value = 1 Then
 
 hem = "YES"
 
 
End If


If chk_cbc.Value = 1 Or CHK_HAEMOGRAM.Value = 1 Then
   cbc = "YES"
   GRP1 = True
End If

If chk_sm.Value = 1 Or CHK_SMC.Value = 1 Then
  smar = "YES"
  GRP2 = True
End If

If chk_cytochemistry.Value = 1 Then
  cyto = "YES"
  GRP2 = True
End If

If chk_pt.Value = 1 Or CHK_PTTK.Value = 1 Or CHK_FDPDDIMER.Value = 1 Or CHK_OTHERS.Value = 1 Or CHK_TT.Value = 1 Or CHK_FIBRINOGEN.Value = 1 Then
  cg1 = "YES"
  GRP1 = True
  
End If

If CHK_TT.Value = 1 Then
    TT = "YES"
End If

If CHK_OTHERS.Value = 1 Then
      oth = "YES"
      othtext = Trim(Text1.Text)
End If



'If GRP1 = True And GRP2 = True Then
'    MsgBox "You can Select either CBC/HEMOGRAM/COAGULATION  OR CYTOCHEMISTRY/SURFACE MARKER"
'    Exit Sub
'End If

    



'If HLFLAG = False Then
       ' ******** for cbc
       If chk_cbc.Value = 1 Then
          If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            'pathology.CMB_TEXT.AddItem "U721"
            'Commented and added by Shraddha on 26/09/2011
            pathology.CMB_TEXT.AddItem "U722"
            ccount = ccount + 1
        End If
        '**** Haemogram
        ''
                If CHK_HAEMOGRAM.Value = 1 And CHK_MORPHOLOGYDLC.Value = 0 Then
                   If ccount > 11 Then
                        MsgBox "You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                      'pathology.CMB_TEXT.AddItem "U722"
                      'Commented and added by Shraddha on 26/09/2011
                      pathology.CMB_TEXT.AddItem "U713"
                      ccount = ccount + 1
                End If
                 If CHK_HAEMOGRAM.Value = 0 And CHK_MORPHOLOGYDLC.Value = 1 Then
                   If ccount > 11 Then
                        MsgBox "You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                      pathology.CMB_TEXT.AddItem "U723"
                      ccount = ccount + 1
                End If
        '''
'        If CHK_MORPHOLOGYDLC.Value = 1 Then
'            If ccount > 11 Then
'                MsgBox "You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            pathology.CMB_TEXT.AddItem "U723"
'            ccount = ccount + 1
'        Else
'            If CHK_DIFFCOUNT.Value = 1 Then
'                If ccount > 11 Then
'                    MsgBox "You Can Select Max. of 11 Procedures only."
'                    Unload Me
'                    pathology.show
'                    Exit Sub
'                End If
'                pathology.CMB_TEXT.AddItem "U722"
'                ccount = ccount + 1
'            Else
'                If CHK_HAEMOGRAM.Value = 1 Then
'                    If ccount > 11 Then
'                        MsgBox "You Can Select Max. of 11 Procedures only."
'                        Unload Me
'                        pathology.show
'                        Exit Sub
'                    End If
'                    pathology.CMB_TEXT.AddItem "U721"
'                    ccount = ccount + 1
'                End If
'            End If
'        End If
        '**** Haemogram
        '**** ESR
        If CHK_ESR.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U706"
            ccount = ccount + 1
        End If
        '**** ESR
        
        '***** MDC *****
        If chk_mdc.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U702"
            ccount = ccount + 1
        End If
                
        '***** cbc mplt *****
        If chk_CBCmplt.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U703"
            ccount = ccount + 1
        End If
        '********************
        
        
        
        '**** MALARIAL PARASITE
        If CHK_MALARIALPARASITE.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U713"
            ccount = ccount + 1
        End If
        '**** MALARIAL PARASITE
        '**** RETIC COUNT
        If CHK_RETICCOUNT.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U724"
            ccount = ccount + 1
        End If
        'following condition commented by shweta 11-05-2012 as told by kalpana mam
        '**** RETIC COUNT
        '**** CYTOCHEMISTRY ON PERIPHERAL BLOOD
'        If chk_cytochemistry.Value = 1 Or chk_mpo.Value = 1 Or CHK_SBB.Value = 1 Or CHK_PAS.Value = 1 Or CHK_LAP.Value = 1 Or CHK_NSE.Value = 1 Or CHK_TRAP1.Value = 1 Then
'            If ccount > 11 Then
'                MsgBox "You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            'pathology.CMB_TEXT.AddItem "U752"
'            'Commented and added by Shraddha on 26/09/2011
'            pathology.CMB_TEXT.AddItem "U755"
'            ccount = ccount + 1
'        End If
        '**** CYTOCHEMISTRY ON PERIPHERAL BLOOD
        '**** SURFACE MARKER
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
        
        '**** SURFACE MARKER
        '**** COAGULATION PROFILE
        '**** COAGULATION PROFILE new
        If CHK_CP.Value = 1 And chk_pt.Value = 1 And CHK_PTTK.Value = 1 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 0 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U709"
            ccount = ccount + 1
        
        End If
        If CHK_CP.Value = 1 And chk_pt.Value = 1 And CHK_PTTK.Value = 1 And CHK_FIBRINOGEN.Value = 1 And CHK_FDPDDIMER.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U712"
            ccount = ccount + 1
        
        End If
        'If CHK_CP.Value = 1 And chk_pt.Value = 0 And CHK_PTTK.Value = 0 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 1 Then
        'Commented and added by Shraddha on 26/11/10
        If chk_pt.Value = 0 And CHK_PTTK.Value = 0 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U714"
            ccount = ccount + 1
        
        End If
        
        'If CHK_CP.Value = 1 And chk_pt.Value = 0 And CHK_PTTK.Value = 0 And CHK_FIBRINOGEN.Value = 1 And CHK_FDPDDIMER.Value = 0 Then
        'Commented and added by Shraddha on 26/11/10
        If chk_pt.Value = 0 And CHK_PTTK.Value = 0 And CHK_FIBRINOGEN.Value = 1 And CHK_FDPDDIMER.Value = 0 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U715"
            ccount = ccount + 1
        
        End If
        ' codes below are temporirily suspended
        
'        If CHK_CP.Value = 1 And chk_pt.Value = 1 And CHK_PTTK.Value = 1 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 0 Then
'            If ccount > 11 Then
'                MsgBox "You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            pathology.CMB_TEXT.AddItem "U709"
'            ccount = ccount + 1
        
'        End If
        
        If CHK_CP.Value = 1 And chk_pt.Value = 1 And CHK_PTTK.Value = 0 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 0 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U708"
            ccount = ccount + 1
        
        End If
        
         If CHK_CP.Value = 1 And chk_pt.Value = 0 And CHK_PTTK.Value = 1 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 0 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U710"
            ccount = ccount + 1
        
        End If
        ' FOR TT
        If CHK_CP.Value = 1 And CHK_TT.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U777" ' U777 CODE IS TO BE CHANGED WHEN ACTUAL CODE FOR TT IS DECIDED
            
            ccount = ccount + 1
        
        End If
        
'        If CHK_CP.Value = 0 And chk_pt.Value = 0 And CHK_PTTK.Value = 0 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 1 Then
'            If ccount > 11 Then
'                MsgBox "You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            pathology.CMB_TEXT.AddItem "U714"
'            ccount = ccount + 1
'
'        End If
        
        '**** COAGULATION PROFILE new
'        If CHK_CP.Value = 1 And CHK_FIBRINOGEN.Value = 1 And CHK_FDPDDIMER.Value = 1 Then
'            If ccount > 11 Then
'                MsgBox "You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            pathology.CMB_TEXT.AddItem "U712"
'            ccount = ccount + 1
'        End If
'
'        Else
'            If CHK_CP.Value = 1 And CHK_FIBRINOGEN.Value = 0 And CHK_FDPDDIMER.Value = 0 And chk_pt.Value = 1 And CHK_PTTK.Value = 1 Then
'                If ccount > 11 Then
'                    MsgBox "You Can Select Max. of 11 Procedures only."
'                    Unload Me
'                    pathology.show
'                    Exit Sub
'                End If
'                pathology.CMB_TEXT.AddItem "U711"
'                ccount = ccount + 1
'            Else
'                If chk_pt.Value = 1 And CHK_PTTK.Value = 1 Then
'                        If ccount > 11 Then
'                            MsgBox "You Can Select Max. of 11 Procedures only."
'                            Unload Me
'                            pathology.show
'                            Exit Sub
'                        End If
'                        pathology.CMB_TEXT.AddItem "U710"
'                        ccount = ccount + 1
'                Else
'                    If chk_pt.Value = 1 And chk_bt.Value = 0 Then
'                        If ccount > 11 Then
'                            MsgBox "You Can Select Max. of 11 Procedures only."
'                            Unload Me
'                            pathology.show
'                            Exit Sub
'                        End If
'                        pathology.CMB_TEXT.AddItem "U708"
'                        ccount = ccount + 1
'                    End If
'                    If chk_pt.Value = 0 And chk_bt.Value = 1 Then
'                        If ccount > 11 Then
'                            MsgBox "You Can Select Max. of 11 Procedures only."
'                            Unload Me
'                            pathology.show
'                            Exit Sub
'                        End If
'                        pathology.CMB_TEXT.AddItem "U707"
'                        ccount = ccount + 1
'                    End If
'                    If CHK_PTTK.Value = 1 Then
'                        If ccount > 11 Then
'                            MsgBox "You Can Select Max. of 11 Procedures only."
'                            Unload Me
'                            pathology.show
'                            Exit Sub
'                        End If
'                        pathology.CMB_TEXT.AddItem "U709"
'                        ccount = ccount + 1
'                    End If
'                    If CHK_FIBRINOGEN.Value = 1 Then
'                        If ccount > 11 Then
'                            MsgBox "You Can Select Max. of 11 Procedures only."
'                            Unload Me
'                            pathology.show
'                            Exit Sub
'                        End If
'                        pathology.CMB_TEXT.AddItem "U715"
'                        ccount = ccount + 1
'                    End If
'                    If CHK_FDPDDIMER.Value = 1 Then
'                        If ccount > 11 Then
'                            MsgBox "You Can Select Max. of 11 Procedures only."
'                            Unload Me
'                            pathology.show
'                            Exit Sub
'                        End If
'                        pathology.CMB_TEXT.AddItem "U714"
'                        ccount = ccount + 1
'                    End If
'                End If
'            End If
'        End If
'        '**** COAGULATION PROFILE
        
        
 
'End If

'***** Shraddha 26/09/2011 *****
str_pbs = ""
str_cytochem = ""
If CHK_DIFFCOUNT.Value = 1 Then
    str_pbs = str_pbs & "MDC, "
End If
If chk_mp.Value = 1 Then
    str_pbs = str_pbs & "MP, "
End If
If CHK_RBC.Value = 1 Then
    str_pbs = str_pbs & "RBC Morphology, "
End If
If chk_MPLT.Value = 1 Then
    str_pbs = str_pbs & "MPLT, "
End If

If Trim(txt_other.Text) <> "" Then
    str_pbs = str_pbs & Trim(txt_other.Text) & ", "
End If

If chk_mpo.Value = 1 Then
    str_pbs = str_pbs & "MPO, "
End If

If CHK_LAP.Value = 1 Then
    str_pbs = str_pbs & "LAP, "
End If


If Trim(txt_cyto_other.Text) <> "" Then
    str_cytochem = str_cytochem & ", " & Trim(txt_cyto_other.Text)
End If

If Trim(Text1.Text) <> "" Then
    str_cytochem = str_cytochem & ", " & Trim(Text1.Text)
End If

If CHK_TRAP1.Value = 1 Then
    str_cytochem = str_cytochem & "TRAP "
End If

If chk_Warffine.Value = 1 Then
    str_cytochem = str_cytochem & "On Warffine "
End If

If chk_Heaparin.Value = 1 Then
    str_cytochem = str_cytochem & "On Heaparin "
End If
'*******************************

'********** added by v.s.k. on 22-05-2012
str_surface = ""
If CHK_SMPERIPHERALBLOOD.Value = 1 Or chk_sm.Value = 1 Or CHK_SMC.Value = 1 Then
    str_surface = Trim(txtSurfaceMarker.Text)
End If
'****** end
If CHK_HAEMOGRAM.Value = 1 Then
    If chk_mp.Value = 0 And CHK_RBC.Value = 0 And CHK_DIFFCOUNT = 0 And chk_MPLT = 0 Then
        MsgBox "Please select any of the below test."
        Exit Sub
    End If
End If

'*** Added by V.S.K. on 9 Apr 2012
'    If chkCytoChem.Value = 1 Then
'        If optCyto(0).Value = False And optCyto(1).Value = False Then
'            MsgBox "Please selcet between MPO / LAP"
'            Exit Sub
'        End If
'    End If

'****end V.S.K.

'If chk_cytochemistry.Value = 1 Then
'    If chk_mpo.Value = 0 And CHK_LAP.Value = 0 And CHK_TRAP1.Value = 0 Then
'        MsgBox "Please selcet between MPO / LAP"
'        Exit Sub
'    End If
'End If

BCFLAG = False
Unload Me
pathology.show
'pathology.Biochemistry1.Visible = True

End Sub

Private Sub Form_Load()
 ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ***********
'clear Public variables added by v.s.k. on 24-05-2012
str_surface = ""
str_cytochem = ""
str_pbs = ""
'end
 EnableCloseButton Me.hwnd, False
If HEMCBCFLAG = False Then
    ''DCBCREPORT FOR HL
    str = "SELECT * FROM DISLIB.DCBCREPORT WHERE REQNO=?"
    Set SELDCBCREPORT_QR = rdb.CreateQuery("SELDCBCREPORT_STRSDDED", str)
    
    ''DCBCREPORT FOR HL
    str = "SELECT * FROM DISLIB.DBMAREPORT WHERE REQNO=?"
    Set SELDBMAREPORT_QR = rdb.CreateQuery("SELDBMAREPORT_STRCVGF", str)
    
    'STR = "SELECT * FROM RISLIB.RREQUISITN LEFT OUTER JOIN DISLIB.DREQUISITN ON RISLIB.RREQUISITN.REQNNO=DISLIB.DREQUISITN.REQNO WHERE RISLIB.RREQUISITN.CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')" 'ORDER BY RREQUISITN.REQNDT DESC"
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr10 = rdb.CreateQuery("seldrequisitn11_1qr11XX0", str)
         
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr11 = rdb.CreateQuery("seldrequisitn11_1qrXXREP", str)
     
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQR1SPREQXX", str)
     
    str = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11SPREQXX", str)
        
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1SPREQ_QR1XX", str)
    
    str = "SELECT REQNO,FSPREPORT,IHREPORT,IMPRESSION,GROSSDESC,GROSSEDBY,SPREPORT,REPORTIND,FNAC FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QR1sprepXX", str)
    
    str = "SELECT * FROM PATOOLSLIB.TA1012 WHERE CASE_NO=?"
    Set selta1012_qr = rdb.CreateQuery("SELTA1012_QR11kl", str)
    selta1012_qr(0) = "TT/00333"
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    
    str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
    Set SELMA301QR = rdb.CreateQuery("SELMA301QR1QSW", str)
    
    SELMA301QR(0) = "F22"
    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
    
    ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ***********End Sub
    HEMCBCFLAG = True
End If
Shape1.Visible = False
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
    CrystalReport1.Formulas(i) = ""
Next i


str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"
If pathology.Opt_patient.Value = True Then
        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = 'Prev. Path No." & Trim(prpathno) & "'"
        Else
             CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = msk_caseno.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        Dim z
        Dim Y
        Dim w
        Dim u
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
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
                    CrystalReport1.Formulas(4) = "OPPRO = '" & Y & "'"
                Else
                    CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
                End If
            Else
                CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
            End If
        Else
            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
        End If
        
        CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
        
        'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
        'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
        CrystalReport1.WindowState = crptMaximized
      '  CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
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
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(selpathdrequisitn_rs("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = MaskEdBox1.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301_QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
        Else
            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
         
        End If
        CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(13) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(14) = "DOB = '" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT")) & "'"
        CrystalReport1.Formulas(15) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
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
    CrystalReport1.Formulas(i) = ""
Next i

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"

If pathology.Opt_patient.Value = True Then
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(2) = "ENT = '" & str1 & "'"
             Else
    
                CrystalReport1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(2) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        End If
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        End If
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
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
    CrystalReport1.Formulas(i) = ""
Next i
SELPATHDREQUISITN_QR(0) = VAR_REQNNO
Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"

If pathology.Opt_patient.Value = True Then
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(4) = "NOSLIDES='" & Trim(selpathdrequisitn_rs("LABNO")) & "'"
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) <> UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "P= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(12) = "F= '" & Trim(SELRREQNORS("FREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
            CrystalReport1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) = "" Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) = UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(12) = "NOSLIDES = " & "" & selpathdrequisitn_rs("LABNO")
    End If
Else
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    End If
    CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
    CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
    CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
End If
CrystalReport1.WindowState = crptMaximized
'CRYSTALREPORT1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
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

    If SELDBMAREPORT_FLAG = False Then
        SELDBMAREPORT_QR(0) = VAR_REQNNO
        Set SELDBMAREPORT_RS = SELDBMAREPORT_QR.OpenResultset(1, 3)
        SELDBMAREPORT_FLAG = True
    Else
        SELDBMAREPORT_QR(0) = VAR_REQNNO
        SELDBMAREPORT_RS.MoreResults
        SELDBMAREPORT_RS.Requery
    End If
    If pathology.Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBM2.RPT"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Not selta1012_rs.EOF Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
        End If
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & SELRREQNORS("ENTERBY")
            If Len(Trim(SELRREQNORS("ENTERBY"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If IsNull(SELDBMAREPORT_RS("CD2")) And IsNull(SELDBMAREPORT_RS("CD3")) And IsNull(SELDBMAREPORT_RS("CD4")) And IsNull(SELDBMAREPORT_RS("CD5")) And IsNull(SELDBMAREPORT_RS("CD7")) And IsNull(SELDBMAREPORT_RS("CD8")) And IsNull(SELDBMAREPORT_RS("CYTOCD3")) Then
            CrystalReport1.Formulas(13) = "tmarkers = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "tmarkers = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("CD10")) And IsNull(SELDBMAREPORT_RS("CD19")) And IsNull(SELDBMAREPORT_RS("FMC7")) And IsNull(SELDBMAREPORT_RS("CD22")) And IsNull(SELDBMAREPORT_RS("CD23")) And IsNull(SELDBMAREPORT_RS("smig")) And IsNull(SELDBMAREPORT_RS("CD79B")) And IsNull(SELDBMAREPORT_RS("CD25")) And IsNull(SELDBMAREPORT_RS("CD11C")) And IsNull(SELDBMAREPORT_RS("cd103")) And IsNull(SELDBMAREPORT_RS("kappa")) And IsNull(SELDBMAREPORT_RS("lamda")) And IsNull(SELDBMAREPORT_RS("DBA44")) And IsNull(SELDBMAREPORT_RS("cytoEP22")) And IsNull(SELDBMAREPORT_RS("cytocd79a")) And IsNull(SELDBMAREPORT_RS("CYTOMU")) Then
            CrystalReport1.Formulas(14) = "bmarkers = '" & "" & "'"
        Else
            CrystalReport1.Formulas(14) = "bmarkers = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("CD34")) And IsNull(SELDBMAREPORT_RS("HLADR")) And IsNull(SELDBMAREPORT_RS("CD38")) And IsNull(SELDBMAREPORT_RS("TDT")) And IsNull(SELDBMAREPORT_RS("CD138")) And IsNull(SELDBMAREPORT_RS("ANTIMPO")) Then
            CrystalReport1.Formulas(15) = "others = '" & "" & "'"
        Else
                CrystalReport1.Formulas(15) = "others = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("cd13")) And IsNull(SELDBMAREPORT_RS("Cd14")) And IsNull(SELDBMAREPORT_RS("CD15")) And IsNull(SELDBMAREPORT_RS("CD33")) And IsNull(SELDBMAREPORT_RS("cd36")) And IsNull(SELDBMAREPORT_RS("cd4142")) And IsNull(SELDBMAREPORT_RS("Glycophoin")) And IsNull(SELDBMAREPORT_RS("CD68")) And IsNull(SELDBMAREPORT_RS("CD42")) And IsNull(SELDBMAREPORT_RS("CD117")) Then
           CrystalReport1.Formulas(16) = "mm = '" & "" & "'"
        Else
           CrystalReport1.Formulas(16) = "mm = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("IRONSTAIN")) And IsNull(SELDBMAREPORT_RS("GIDEROBLAS")) And IsNull(SELDBMAREPORT_RS("sudanblack")) And IsNull(SELDBMAREPORT_RS("TOLUIDINEB")) Then
           CrystalReport1.Formulas(17) = "cci = '" & "" & "'"
        Else
           CrystalReport1.Formulas(17) = "cci = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("myelopero")) And IsNull(SELDBMAREPORT_RS("nonspester")) And IsNull(SELDBMAREPORT_RS("trap")) Then
           CrystalReport1.Formulas(18) = "ccm = '" & "" & "'"
        Else
           CrystalReport1.Formulas(18) = "ccm = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("IMPRINT")) And IsNull(SELDBMAREPORT_RS("abcellblas")) Then
           CrystalReport1.Formulas(19) = "im = '" & "" & "'"
        Else
           CrystalReport1.Formulas(19) = "im = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("mcellulari")) And IsNull(SELDBMAREPORT_RS("morphology")) Then
           CrystalReport1.Formulas(20) = "ms = '" & "" & "'"
        Else
           CrystalReport1.Formulas(20) = "ms = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("ecellulari")) And IsNull(SELDBMAREPORT_RS("maturation")) Then
           CrystalReport1.Formulas(21) = "es = '" & "" & "'"
        Else
           CrystalReport1.Formulas(21) = "es = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("mature")) And IsNull(SELDBMAREPORT_RS("immature")) And IsNull(SELDBMAREPORT_RS("atypbinu")) And IsNull(SELDBMAREPORT_RS("plasmablas")) Then
           CrystalReport1.Formulas(22) = "pc = '" & "" & "'"
        Else
           CrystalReport1.Formulas(22) = "pc = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("MLYMPHOCYT")) And IsNull(SELDBMAREPORT_RS("ALYMPHOCYT")) And IsNull(SELDBMAREPORT_RS("Plasmacell")) Then
           CrystalReport1.Formulas(23) = "ls = '" & "" & "'"
        Else
           CrystalReport1.Formulas(23) = "ls = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("PROMYELOCY")) And IsNull(SELDBMAREPORT_RS("PROMYELOCY")) And IsNull(SELDBMAREPORT_RS("MMYELOCYTE")) And IsNull(SELDBMAREPORT_RS("polymorphs")) And IsNull(SELDBMAREPORT_RS("basophils")) And IsNull(SELDBMAREPORT_RS("eosinophil")) And IsNull(SELDBMAREPORT_RS("monocytes")) And IsNull(SELDBMAREPORT_RS("PROMONOCYT")) Then
           CrystalReport1.Formulas(24) = "mys = '" & "" & "'"
        Else
           CrystalReport1.Formulas(24) = "mys = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREPORT_RS("ECELLULARI")) And IsNull(SELDBMAREPORT_RS("MERATIO")) Then
           CrystalReport1.Formulas(25) = "mor = '" & "" & "'"
        Else
           CrystalReport1.Formulas(25) = "mor = '" & "abc" & "'"
        End If
    End If
    CrystalReport1.WindowState = crptMaximized
    'CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1

End Sub
Private Sub BFREPORT()
For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
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
If pathology.Opt_patient.Value = True Then
             
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

    
    
     If Trim(SELRREQNORS("MODALITYCD")) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            If Len(Trim(SELRREQNORS("ENTERBY"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If
       
End If
If pathology.opt_staff.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

 
     If Trim(SELRREQNORS("MODALITYCD")) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBFS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & VAR_REQNNO & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If

End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
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
    CrystalReport1.Formulas(i) = ""
Next i
selta1012_qr(0) = SELRREQNORS("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

If pathology.Opt_patient.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    'If Trim(txt_modcode.Text) = "HL" Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShl.RPT"
    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'       If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'       Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'       End If
    CrystalReport1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
    If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
        CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        str1 = ""
        str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
        If Len(Trim(SELRREQNORS("enterby"))) > 0 Then
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
        End If
    Else
        CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
    End If
    CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
             If Trim(SELDCBCREPORT_RS("lap")) = "" And Trim(SELDCBCREPORT_RS("MPO")) = "" And Trim(SELDCBCREPORT_RS("SBB")) = "" And Trim(SELDCBCREPORT_RS("PSA")) = "" And Trim(SELDCBCREPORT_RS("NSE")) = "" And Trim(SELDCBCREPORT_RS("TRAD")) = "" Then
                'If CYTOCHEMFLAG = True Then
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
                'End If
            End If
            If Trim(SELDCBCREPORT_RS("BT")) = "" And Trim(SELDCBCREPORT_RS("PT")) = "" And Trim(SELDCBCREPORT_RS("PLATELET1")) = "" And Trim(SELDCBCREPORT_RS("PTTK")) = "" And Trim(SELDCBCREPORT_RS("TT")) = "" And Trim(SELDCBCREPORT_RS("FIBRINOGEN")) = "" And Trim(SELDCBCREPORT_RS("FDPDDIMER")) = "" And Trim(SELDCBCREPORT_RS("OTHERS")) = "" Then
                CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
            Else
                CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("NORMBOBLAS")) = "" And Trim(SELDCBCREPORT_RS("OTHERCELLR")) = "" And Trim(SELDCBCREPORT_RS("rbcmorpho")) = "" And Trim(SELDCBCREPORT_RS("reticcount")) = "" And Trim(SELDCBCREPORT_RS("mparasite")) = "" And Trim(SELDCBCREPORT_RS("ESR")) = "" Then
                CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PMORPHOLOG")) = "" And Trim(SELDCBCREPORT_RS("FLD1")) = "" Then
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
            Else
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PROLYMPHOC")) = "" And Trim(SELDCBCREPORT_RS("Sezarycell")) = "" And Trim(SELDCBCREPORT_RS("hairycells")) = "" And Trim(SELDCBCREPORT_RS("othercells")) = "" Then
                CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
            Else
                CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("blasts")) = "" And Trim(SELDCBCREPORT_RS("PROMYELOCY")) = "" And Trim(SELDCBCREPORT_RS("myelocytes")) = "" And Trim(SELDCBCREPORT_RS("metamyeloc")) = "" And Trim(SELDCBCREPORT_RS("plasmacell")) = "" And Trim(SELDCBCREPORT_RS("aTlymphEE")) = "" Then
                CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
            Else
                CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophi1")) = "" And Trim(SELDCBCREPORT_RS("lymphocyt1")) = "" And Trim(SELDCBCREPORT_RS("monocytE1")) = "" And Trim(SELDCBCREPORT_RS("eosinophi1")) = "" And Trim(SELDCBCREPORT_RS("basophill1")) = "" Then
                CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HCT")) = "" And Trim(SELDCBCREPORT_RS("MCH")) = "" And Trim(SELDCBCREPORT_RS("MCV")) = "" And Trim(SELDCBCREPORT_RS("MCHC")) = "" And Trim(SELDCBCREPORT_RS("MPV")) = "" And Trim(SELDCBCREPORT_RS("RDW")) = "" And Trim(SELDCBCREPORT_RS("ABNORMAL")) = "" Then
                CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
            Else
                CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophil")) = "" And Trim(SELDCBCREPORT_RS("lymphocyte")) = "" And Trim(SELDCBCREPORT_RS("monocytes")) = "" And Trim(SELDCBCREPORT_RS("eosinophil")) = "" And Trim(SELDCBCREPORT_RS("basophils")) = "" And Trim(SELDCBCREPORT_RS("immgranulo")) = "" Then
                CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HB")) = "" And Trim(SELDCBCREPORT_RS("RBC")) = "" And Trim(SELDCBCREPORT_RS("TLC")) = "" And Trim(SELDCBCREPORT_RS("platelets")) = "" Then
                CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
            End If
End If
If pathology.opt_staff.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    If Trim(txt_modcode.Text) = "EK" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISECS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'           If Trim(prepath) = "" Then
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'           Else
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
    End If
End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1

End Sub

Private Sub Label10_Click()

End Sub

