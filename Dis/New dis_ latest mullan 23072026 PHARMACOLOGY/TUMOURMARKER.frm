VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form TUMOURMARKER 
   Caption         =   "Tumour Marker"
   ClientHeight    =   11055
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
   ScaleHeight     =   11055
   ScaleWidth      =   16275
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   4335
      Left            =   360
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame7 
         Caption         =   "Other Tests"
         Height          =   1695
         Left            =   8040
         TabIndex        =   65
         Top             =   2400
         Width           =   3375
         Begin VB.CheckBox chk_vitD 
            Caption         =   "VITAMIN D*"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chk_pth 
            Caption         =   "Parathormone (PTH) "
            Enabled         =   0   'False
            Height          =   435
            Left            =   1680
            TabIndex        =   67
            Top             =   1080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk_calcitonin 
            Caption         =   "Calcitonin"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1935
         Left            =   8040
         TabIndex        =   60
         Top             =   360
         Width           =   3375
         Begin VB.CheckBox chk_folate 
            Caption         =   "Folate"
            Height          =   195
            Left            =   240
            TabIndex        =   86
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chk_vitamin 
            Caption         =   "Vitamin B12"
            Height          =   195
            Left            =   1800
            TabIndex        =   85
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CheckBox chk_freeT4 
            Caption         =   "FREE T4*"
            Height          =   255
            Left            =   1800
            TabIndex        =   84
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chk_freeT3 
            Caption         =   "FREE T3*"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chk_TSH 
            Caption         =   "TSH"
            Height          =   195
            Left            =   2160
            TabIndex        =   64
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk_T4 
            Caption         =   "T4"
            Height          =   195
            Left            =   1200
            TabIndex        =   63
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk_T3 
            Caption         =   "T3"
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chk_tft 
            Caption         =   "Thyroid Function Tests"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "MARKERS"
         ForeColor       =   &H000000C0&
         Height          =   3615
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   7680
         Begin VB.TextBox txt_rem 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   58
            Top             =   3120
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.CheckBox serum_AFP 
            Caption         =   "Serum AFP"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   57
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CheckBox serum_PSA 
            Caption         =   "Serum PSA"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3600
            TabIndex        =   56
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chk_addmark 
            Caption         =   "Other Marker"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5880
            TabIndex        =   55
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox CHK_CYFRA 
            Caption         =   "CYFRA - 21"
            Enabled         =   0   'False
            Height          =   195
            Left            =   600
            TabIndex        =   54
            Top             =   1920
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox ChK_NSE 
            Caption         =   "NSE "
            Enabled         =   0   'False
            Height          =   195
            Left            =   3600
            TabIndex        =   53
            Top             =   1440
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chk_crp 
            Caption         =   "CRP"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5880
            TabIndex        =   52
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox serumCA19_9 
            Caption         =   "Serum CA 19-9"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3600
            TabIndex        =   51
            Top             =   2520
            Width           =   1695
         End
         Begin VB.CheckBox serumFerritin 
            Caption         =   "Serum Ferritin"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   5880
            TabIndex        =   50
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox serumCA_153 
            Caption         =   "Serum CA 15-3"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3600
            TabIndex        =   49
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox serumB2_Microglobin 
            Caption         =   "Serum B2 - Microglobulin"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   48
            Top             =   1380
            Width           =   2535
         End
         Begin VB.CheckBox serum_CA_125 
            Caption         =   "Serum CA 125"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   47
            Top             =   810
            Width           =   1575
         End
         Begin VB.CheckBox serumB_HCG 
            Caption         =   "Serum B-HCG"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   46
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CheckBox serum_CEA 
            Caption         =   "Serum CEA"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   45
            Top             =   330
            Width           =   1335
         End
         Begin VB.Label lbl_rem 
            AutoSize        =   -1  'True
            Caption         =   "Additional Markers"
            Height          =   195
            Left            =   600
            TabIndex        =   59
            Top             =   3120
            Visible         =   0   'False
            Width           =   1710
         End
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame 2"
      Height          =   2655
      Left            =   2760
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CheckBox chk_UFLC 
         Caption         =   "Urine Free Light chains (24 hours urine)"
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   960
         Width           =   3855
      End
      Begin VB.CheckBox chk_urlambda 
         Caption         =   "Urine Free Light Chain Lambda"
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox chk_urkappa 
         Caption         =   "Urine Free Light Chain Kappa"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chk_bjp 
         Caption         =   "BJP  (24 Hours  Urine)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame SPE 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3495
      Left            =   240
      TabIndex        =   26
      Top             =   1920
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CheckBox chk_comleteSPE 
         Caption         =   "Serum Protein Electrophoresis"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   40
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CheckBox chk_SPEIgLC 
         Caption         =   "Serum Protein Electrophoresis (SPE) + Immunoglobulin (Ig) +  Light Chains (LC)"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   480
         MaskColor       =   &H00000000&
         TabIndex        =   39
         Top             =   240
         Width           =   5655
      End
      Begin VB.CheckBox chk_IFE 
         Caption         =   "Immuno Fixation Electrophoresis (IFE)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         MaskColor       =   &H00000000&
         TabIndex        =   38
         Top             =   360
         Width           =   3615
      End
      Begin VB.Frame Frame4 
         Height          =   2415
         Left            =   480
         TabIndex        =   33
         Top             =   840
         Width           =   5655
         Begin VB.CheckBox chk_SPEIgLC1 
            Caption         =   "SPE + Ig + LC"
            Height          =   255
            Left            =   4080
            TabIndex        =   78
            Top             =   1560
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.CheckBox ImmunoelectrophoresisIEP 
            Caption         =   "Serum Protein Electrophoresis"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   240
            Width           =   2895
         End
         Begin VB.CheckBox chk_lambda 
            Caption         =   "Lambda Total"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox chk_kappa 
            Caption         =   "Kappa Total"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CheckBox chk_igs 
            Caption         =   "Serum Immunoglobulins (IgS)"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4080
            MaskColor       =   &H00000000&
            TabIndex        =   37
            Top             =   2040
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CheckBox chk_iga 
            Caption         =   "Immunoglobulin A (IgA)"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   36
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chk_igm 
            Caption         =   "Immunoglobulin M (IgM)"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   35
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chk_igg 
            Caption         =   "Immunoglobulin G (IgG)"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   34
            Top             =   1320
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   6360
         TabIndex        =   27
         Top             =   840
         Width           =   3735
         Begin VB.CheckBox chk_serligchains 
            Caption         =   "Serum Free Light Chains"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CheckBox chk_kappa1 
            Caption         =   "Kappa Total"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk_lambda1 
            Caption         =   "Lambda Total"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1440
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk_kappafree 
            Caption         =   "Free Kappa Light chain"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   2655
         End
         Begin VB.CheckBox chk_lambdafree 
            Caption         =   "Free Lambda Light chain "
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   720
            Width           =   2535
         End
      End
   End
   Begin VB.Frame Frame9 
      Height          =   855
      Left            =   480
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   10935
      Begin VB.OptionButton opt_3 
         Caption         =   "Estimation Of Drug Level"
         Height          =   255
         Left            =   8040
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton opt_2 
         Caption         =   "Serum Immuno Electrophoresis (SIEP) Profile "
         Height          =   375
         Left            =   4260
         TabIndex        =   24
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton opt_1 
         Caption         =   "Markers,Thyroid Function Test, Vitamin test"
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.TextBox txt_clininfo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   2040
      Width           =   7455
   End
   Begin VB.ComboBox CMB_NMETERIAL 
      Height          =   315
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<<&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   4
      Top             =   9600
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Msk_caseno 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7200
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estimation of Drug Level"
      ForeColor       =   &H000000C0&
      Height          =   2535
      Left            =   240
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox txt_hrs 
         Height          =   285
         Left            =   4440
         TabIndex        =   80
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox CMB_FR 
         Height          =   315
         ItemData        =   "TUMOURMARKER.frx":0000
         Left            =   6240
         List            =   "TUMOURMARKER.frx":000A
         TabIndex        =   72
         Top             =   1680
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox TXT_MM 
         Height          =   285
         Left            =   5640
         TabIndex        =   71
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXT_HH 
         Height          =   285
         Left            =   5040
         TabIndex        =   70
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmb_unit 
         Height          =   315
         ItemData        =   "TUMOURMARKER.frx":0016
         Left            =   10440
         List            =   "TUMOURMARKER.frx":0023
         TabIndex        =   42
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_amt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txt_Dosage 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   7815
      End
      Begin VB.CheckBox chk_metho 
         Caption         =   "Methotrexate"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6840
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox chk_vanc 
         Caption         =   "Vancomycin"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chk_amik 
         Caption         =   "Amikacin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chk_cyclos 
         Caption         =   "Cyclosporin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox msk_fromdt 
         Height          =   255
         Left            =   2880
         TabIndex        =   69
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
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
      Begin VB.Label Label6 
         Caption         =   "hours"
         Height          =   255
         Left            =   5400
         TabIndex        =   81
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "mg and after"
         Height          =   255
         Left            =   3240
         TabIndex        =   79
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "HH   :   MM"
         Height          =   255
         Left            =   5040
         TabIndex        =   73
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dosage of Drug Given at "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   68
         Top             =   1680
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Drug Dosage"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   1200
         Width           =   1125
      End
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   11535
   End
   Begin VB.Label LBL_CLINCAL 
      AutoSize        =   -1  'True
      Caption         =   "Clinical Information"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   1635
   End
   Begin VB.Label LBL_TM 
      AutoSize        =   -1  'True
      Caption         =   "Tumour Marker"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   2145
   End
   Begin VB.Label LBL_NMETERIAL 
      AutoSize        =   -1  'True
      Caption         =   "Nature of Material"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6480
      TabIndex        =   2
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label lbl_caseno 
      AutoSize        =   -1  'True
      Caption         =   "Case No."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1545
      TabIndex        =   0
      Top             =   1125
      Width           =   795
   End
End
Attribute VB_Name = "TUMOURMARKER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Object

Private Sub Check1_Click()

End Sub

Private Sub chk_addmark_Click()
If chk_addmark.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub chk_amik_Click()
If chk_amik.Value = 1 Then
    'MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
    'Shraddha 07/04/10
'    If txt_Dosage.Text = "" Then
'        MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
'        Exit Sub
'    End If
    
    chk_cyclos.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = False
        End If
    Next
    chk_amik.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True

    'txt_amt.SetFocus
Else
    chk_cyclos.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
        End If
    Next
    chk_amik.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True

    
End If
End Sub

Private Sub chk_bjp_Click()
If chk_bjp.Value = 1 Then
    ImmunoelectrophoresisIEP.Enabled = False
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = False
            txt_Dosage.Enabled = False
            chk_bjp.Enabled = True
        End If
    Next
    chk_UFLC.Enabled = True
    chk_urkappa.Enabled = True
    chk_urlambda.Enabled = True
            
Else
    ImmunoelectrophoresisIEP.Enabled = True
    For Each c In Controls
    If TypeOf c Is CheckBox Then
        c.Enabled = True
        txt_Dosage.Enabled = True
    End If
Next
End If


'If chk_bjp.Value = 1 Then ImmunoelectrophoresisIEP.Value = 0
'If chk_bjp.Value = 0 Then ImmunoelectrophoresisIEP.Value = 0
End Sub




Private Sub chk_calcitonin_Click()
If chk_calcitonin.Value = 1 Then
        Frame1.Visible = False
        Frame6.Visible = False
Else
        Frame1.Visible = True
        Frame6.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_comleteSPE_Click()
 Dim ctr As Object
    If chk_comleteSPE.Value = 1 Then
        ImmunoelectrophoresisIEP.Value = 1
        'chk_IFE.Value = 1 'Commented Shraddha 04/03/10
        'chk_igs.Value = 1 'Commented Shraddha 04/03/10
        chk_iga.Value = 1
        chk_igm.Value = 1
        chk_igg.Value = 1
        'chk_serligchains.Value = 1 'Commented Shraddha 04/03/10
        chk_kappa.Value = 1
        chk_lambda.Value = 1
        'chk_kappafree.Value = 1 'Commented Shraddha 04/03/10
        'chk_lambdafree.Value = 1 'Commented Shraddha 04/03/10
        
        For Each ctr In Controls
            If TypeOf ctr Is CheckBox Then
                ctr.Enabled = False
            End If
        Next
        '***** Added by Shraddha on 04/03/10 *****
        chk_IFE.Enabled = True
        chk_serligchains.Enabled = True
        chk_kappafree.Enabled = True
        chk_lambdafree.Enabled = True
        '*****************************************
        
        chk_comleteSPE.Enabled = True
     Else
        ImmunoelectrophoresisIEP.Value = 0
        'chk_IFE.Value = 0
        'chk_igs.Value = 0
        chk_iga.Value = 0
        chk_igm.Value = 0
        chk_igg.Value = 0
        'chk_serligchains.Value = 0
        chk_kappa.Value = 0
        chk_lambda.Value = 0
        'chk_kappafree.Value = 0
        'chk_lambdafree.Value = 0
        For Each ctr In Controls
            If TypeOf ctr Is CheckBox Then
                ctr.Enabled = True
            End If
        Next
        chk_comleteSPE.Enabled = True
     End If
End Sub

Private Sub chk_crp_Click()
If chk_crp.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub



Private Sub chk_cyclos_Click()
If chk_cyclos.Value = 1 Then
    'MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
    'Shraddha 07/04/10
'    If txt_Dosage.Text = "" Then
'        MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
'        Exit Sub
'    End If
    
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = False
        End If
    Next
    chk_cyclos.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True
    'txt_amt.SetFocus
Else

    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
        End If
    Next
    chk_cyclos.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True
    'txt_hrs.Enabled = True
    
End If

End Sub

Private Sub CHK_CYFRA_Click()
If CHK_CYFRA.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    txt_amt.Enabled = False
    cmb_unit.Enabled = False
    'txt_hrs.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True
    'txt_hrs.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub chk_folate_Click()
If chk_folate.Value = 1 Then
 '***** Shweta 26/04/2012 *****
'        Frame1.Visible = False
'        Frame7.Visible = False
        Frame6.Visible = True
Else
        Frame7.Visible = True
        Frame1.Visible = True
        Frame6.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_freeT3_Click()
If chk_freeT3.Value = 1 Then
 '***** Shweta 26/04/2012 *****
        'Frame1.Visible = False
        'Frame7.Visible = False
Else
        Frame1.Visible = True
        Frame7.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_freeT4_Click()
If chk_freeT4.Value = 1 Then
 '***** Shweta 26/04/2012 *****
'        Frame1.Visible = False
'        Frame7.Visible = False
Else
        Frame1.Visible = True
        Frame7.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_igs_Click()
Dim ctr As Object
    If chk_igs.Value = 1 Then
        chk_iga.Value = 1
        chk_igm.Value = 1
        chk_igg.Value = 1
        chk_iga.Enabled = False
        chk_igm.Enabled = False
        chk_igg.Enabled = False
        chk_igs.Enabled = True
        
     Else
        chk_iga.Value = 0
        chk_igm.Value = 0
        chk_igg.Value = 0
        chk_iga.Enabled = True
        chk_igm.Enabled = True
        chk_igg.Enabled = True
        chk_igs.Enabled = True
    End If
    
End Sub

Private Sub chk_metho_Click()
If chk_metho.Value = 1 Then
    'MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
    'Shraddha 07/04/10
'    If txt_Dosage.Text = "" Then
'        MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
'        Exit Sub
'    End If
    
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = False
        End If
    Next
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True
    'txt_amt.SetFocus
    
Else
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
        End If
    Next
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True
End If
End Sub

Private Sub ChK_NSE_Click()
If CHK_NSE.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    txt_amt.Enabled = False
    cmb_unit.Enabled = False
    'txt_hrs.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True
   ' txt_hrs.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub chk_pth_Click()
If chk_pth.Value = 1 Then
        Frame1.Visible = False
        Frame6.Visible = False
Else
        Frame1.Visible = True
        Frame6.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_serligchains_Click()
Dim ctr As Object
    If chk_serligchains.Value = 1 Then
        'chk_kappa.Value = 1 'Commented by Shraddha on 04/03/10
        'chk_lambda.Value = 1 'Commented by Shraddha on 04/03/10
        chk_kappafree.Value = 1
        chk_lambdafree.Value = 1
        'chk_kappa.Enabled = False 'Commented by Shraddha on 04/03/10
        'chk_lambda.Enabled = False 'Commented by Shraddha on 04/03/10
        chk_kappafree.Enabled = False
        chk_lambdafree.Enabled = False
         
        chk_serligchains.Enabled = True
     Else
        'chk_kappa.Value = 0 'Commented by Shraddha on 04/03/10
        'chk_lambda.Value = 0 'Commented by Shraddha on 04/03/10
        chk_kappafree.Value = 0
        chk_lambdafree.Value = 0
        'chk_kappa.Enabled = True 'Commented by Shraddha on 04/03/10
        'chk_lambda.Enabled = True 'Commented by Shraddha on 04/03/10
        chk_kappafree.Enabled = True
        chk_lambdafree.Enabled = True
        chk_serligchains.Enabled = True
    End If
End Sub

Private Sub chk_SPEIgLC_Click()
    If chk_SPEIgLC.Value = 1 Then
        ImmunoelectrophoresisIEP.Value = 1
        chk_iga.Value = 1
        chk_igm.Value = 1
        chk_igg.Value = 1
        chk_kappa.Value = 1
        chk_lambda.Value = 1
        
        For Each ctr In Controls
            If TypeOf ctr Is CheckBox Then
                ctr.Enabled = False
            End If
        Next
        
        chk_IFE.Enabled = True
        chk_serligchains.Enabled = True
        chk_kappafree.Enabled = True
        chk_lambdafree.Enabled = True
        chk_comleteSPE.Enabled = True
        chk_SPEIgLC.Enabled = True
     Else
        ImmunoelectrophoresisIEP.Value = 0
        chk_iga.Value = 0
        chk_igm.Value = 0
        chk_igg.Value = 0
        chk_kappa.Value = 0
        chk_lambda.Value = 0
        
        For Each ctr In Controls
            If TypeOf ctr Is CheckBox Then
                ctr.Enabled = True
            End If
        Next
        chk_comleteSPE.Enabled = True
        chk_SPEIgLC.Enabled = True
     End If
End Sub

Private Sub chk_T3_Click()
'***** Shraddha 26/09/2011 *****
If chk_T3.Value = 1 Then
 '***** Shweta 26/04/2012 *****
'        Frame1.Visible = False
'        Frame7.Visible = False
Else
        Frame1.Visible = True
        Frame7.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_T4_Click()
If chk_T4.Value = 1 Then
'        Frame1.Visible = False
'        Frame7.Visible = False
Else
        Frame1.Visible = True
        Frame7.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_tft_Click()
Dim ctr As Object

    If chk_tft.Value = 1 Then
        chk_T3.Value = 1
        chk_T4.Value = 1
        chk_TSH.Value = 1
        chk_T3.Enabled = False
        chk_T4.Enabled = False
        chk_TSH.Enabled = False
        chk_tft.Enabled = True
        '***** Shraddha 26/09/2011 *****
        '***** Shweta 26/04/2012 *****
        'Frame1.Visible = False
        'Frame7.Visible = False
        '*******************************
    Else
        chk_T3.Value = 0
        chk_T4.Value = 0
        chk_TSH.Value = 0
        chk_T3.Enabled = True
        chk_T4.Enabled = True
        chk_TSH.Enabled = True
        chk_tft.Enabled = True
        '***** Shraddha 26/09/2011 *****
        Frame1.Visible = True
        Frame7.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
        '*******************************
    End If
End Sub

Private Sub chk_TSH_Click()
If chk_TSH.Value = 1 Then
'        Frame1.Visible = False
'        Frame7.Visible = False
Else
        Frame1.Visible = True
        Frame7.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_UFLC_Click()
If chk_UFLC.Value = 1 Then
        
        chk_urkappa.Value = 1
        chk_urlambda.Value = 1
        
        chk_urkappa.Enabled = False
        chk_urlambda.Enabled = False
         
        chk_UFLC.Enabled = True
     Else
        
        chk_urkappa.Value = 0
        chk_urlambda.Value = 0
        
        chk_urkappa.Enabled = True
        chk_urlambda.Enabled = True
        
        chk_UFLC.Enabled = True
    End If
End Sub

Private Sub chk_vanc_Click()
If chk_vanc.Value = 1 Then
    'MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
    '***** Shraddha 07/04/10 *****
'    If txt_Dosage.Text = "" Then
'        MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
'        Exit Sub
'    End If
    '*****************************
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_metho.Enabled = False
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = False
        End If
    Next
    chk_vanc.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True
    'txt_amt.SetFocus
    
Else
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_metho.Enabled = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Enabled = True
        End If
    Next
    chk_vanc.Enabled = True
    txt_Dosage.Enabled = True
    txt_amt.Enabled = True
    cmb_unit.Enabled = True

End If
End Sub


Private Sub chk_vitamin_Click()
If chk_vitamin.Value = 1 Then
 '***** Shweta 26/04/2012 *****
      '  Frame1.Visible = False
        Frame6.Visible = True
        'Frame7.Visible = False
Else
        Frame7.Visible = True
        'Frame1.Visible = False
        Frame6.Visible = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub chk_vitD_Click()
If chk_vitD.Value = 1 Then
 '***** Shweta 26/04/2012 *****
        'Frame1.Enabled = False
        'Frame6.Enabled = False
         '***** Shweta 26/04/2012 *****
Else
        Frame1.Enabled = True
        Frame6.Visible = True
        Frame6.Enabled = True
        For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
End If
End Sub

Private Sub CMB_NMETERIAL_Click()
'MILIND DATE 09/01/06
txt_Dosage.Enabled = True
txt_rem.Enabled = True
serum_AFP.Enabled = True
serum_CA_125.Enabled = True
serum_CEA.Enabled = True
serum_PSA.Enabled = True
serumB_HCG.Enabled = True
serumB2_Microglobin.Enabled = True
serumCA_153.Enabled = True
serumCA19_9.Enabled = True
serumFerritin.Enabled = True
'CHK_CYFRA.Enabled = True 'Commented on 24/06/10 by Shraddha
chk_crp.Enabled = True
CHK_NSE.Enabled = True
chk_addmark.Enabled = True
ImmunoelectrophoresisIEP.Enabled = True
chk_bjp.Enabled = True
chk_cyclos.Enabled = True
chk_amik.Enabled = True
chk_vanc.Enabled = True
chk_metho.Enabled = True

    Frame10.Visible = False
    Frame10.Enabled = False
    Frame3.Visible = False
    Frame3.Enabled = False
'    Frame6.Visible = False
'    Frame6.Enabled = False
'    Frame7.Visible = False
'    Frame7.Enabled = False
    Frame2.Visible = False
    Frame2.Enabled = False

If Trim(CMB_NMETERIAL.Text) = "Blood" Then
    Frame9.Visible = True
    Frame9.Enabled = True
    OPT_1.Value = False
    OPT_2.Value = False
    OPT_3.Value = False
    Frame8.Visible = False
    Frame8.Enabled = False
ElseIf Trim(CMB_NMETERIAL.Text) = "Urine ( 24 Hours )" Then
    Frame9.Visible = False
    Frame9.Enabled = False
    Frame8.Visible = True
    Frame8.Enabled = True
    Frame8.Top = 3000
    Frame8.Left = 4500
    Frame8.Width = 4500
End If

Dim ctr As Object

For Each ctr In Controls
    If TypeOf ctr Is CheckBox Then
        ctr.Enabled = True
    End If
Next
End Sub

Private Sub cmb_nmeterial_GotFocus()
CMB_NMETERIAL.clear
CMB_NMETERIAL.AddItem "Blood"
CMB_NMETERIAL.AddItem "Urine ( 24 Hours )"
'cmb_nmeterial.AddItem "Pleural"
'cmb_nmeterial.AddItem "Ascitic"
'cmb_nmeterial.AddItem "Pancreatic"
'cmb_nmeterial.AddItem "Amniotic"
'cmb_nmeterial.AddItem "Cystic"
'cmb_nmeterial.AddItem "C.S.F"
CMB_NMETERIAL.Text = "Blood"
End Sub

Private Sub cmd_back_Click()
Dim ccount As Integer
pathology.CMB_TEXT.clear

Dim drug_dose As String
drug_dose = ""

If tmflag = False Then
'UnCommented on 28/06/10
If chk_cyclos.Value = 1 Or chk_amik.Value = 1 Or chk_vanc.Value = 1 Or chk_metho.Value = 1 Then

  'If Len(Trim(txt_amt.Text)) > 0 And Len(Trim(cmb_unit.Text)) > 0 And Len(Trim(msk_fromdt.FormattedText)) = 10 And Len(Trim(TXT_HH.Text)) > 0 And Len(Trim(TXT_MM.Text)) > 0 And Len(Trim(CMB_FR.Text)) > 0 Then
  'If Len(Trim(txt_amt.Text)) > 0 And Len(Trim(msk_fromdt.FormattedText)) = 10 And Len(Trim(TXT_HH.Text)) > 0 And Len(Trim(TXT_MM.Text)) > 0 And Len(Trim(CMB_FR.Text)) > 0 Then
  'Commented and added by Shraddha on 29/06/2011
  If Trim(txt_amt.Text) <> "" And Trim(txt_hrs.Text) <> "" Then
    'txt_Dosage.Text = Trim(txt_amt.Text) & " " & Trim(cmb_unit.Text)
    drug_dose = Trim(txt_amt.Text) & " XXX " & Trim(txt_hrs.Text)

   If Trim(CMB_FR.Text) = "PM" Then
        If Trim(TXT_HH.Text) <> "12" Then
             TXT_HH.Text = 12 + CInt(TXT_HH.Text)
        End If

   End If

   If Trim(CMB_FR.Text) = "AM" Then
        If Trim(TXT_HH.Text) = "12" And Trim(TXT_MM.Text) = "00" Then
             TXT_HH.Text = 12 + CInt(TXT_HH.Text)
        End If
        If Trim(TXT_HH.Text) = "12" And Trim(TXT_MM.Text) <> "00" Then
             TXT_HH.Text = "00"
        End If

   End If

   Dim Yr, mn, DY As String
   Yr = Right(msk_fromdt.FormattedText, 4)
   mn = Mid(msk_fromdt.FormattedText, 4, 2)
   DY = Left(msk_fromdt.FormattedText, 2)


'Commented on 30/06/2011
   'DATEREQ1 = Format(msk_fromdt.FormattedText, "YYYY-MM-DD")
   datereq = Yr & "/" & mn & "/" & DY
   
   
   'UPDRTIMEFORBC(0) = Format(msk_fromdt.FormattedText, "YYYY-MM-DD")
   timereq = Trim(TXT_HH.Text) & ":" & Trim(TXT_MM.Text) & ":" & "00"

   Else
        'MsgBox ("Please enter proper Drug Dosage details"), vbOKOnly
        'txt_amt.SetFocus
        Me.Hide
        pathology.show
        Exit Sub
   End If

End If
'UnCommenetd on 28/06/10
ccount = 0
'**** markers
If serum_CEA.Value = 1 Then
If ccount >= 11 Then
        MsgBox "You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
End If
pathology.CMB_TEXT.AddItem "F818"
ccount = ccount + 1
End If

If serum_AFP.Value = 1 Then
If ccount >= 11 Then
        MsgBox "You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F817"
ccount = ccount + 1
End If

If serumB_HCG.Value = 1 Then
If ccount >= 11 Then
        MsgBox "You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F819"
ccount = ccount + 1
End If

'If serumB_HCG.Value = 1 Then
'pathology.CMB_TEXT.AddItem "Serum B-HCG"
'End If

If serum_CA_125.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F823"
ccount = ccount + 1
End If

If serum_PSA.Value = 1 Then
If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F820"
ccount = ccount + 1
End If

If serumCA19_9.Value = 1 Then
If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F824"
ccount = ccount + 1
End If

If serumB2_Microglobin.Value = 1 Then
If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F821"
ccount = ccount + 1
End If

If serumCA_153.Value = 1 Then
If ccount >= 11 Then
        MsgBox "You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F822"
ccount = ccount + 1
End If

If serumFerritin.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F830"
ccount = ccount + 1
End If
'If Serum_PLAP.Value = 1 Then
'If ccount >= 11 Then
 '       MsgBox " You Can Select Max. of 11 Procedures only."
  '      Unload Me
   '     pathology.show
    '    Exit Sub
    'End If
'pathology.CMB_TEXT.AddItem "F826"
'ccount = ccount + 1
'End If
If chk_crp.Value = 1 Then
If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F829"
ccount = ccount + 1
End If
If chk_addmark.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F827"
ccount = ccount + 1
End If

If CHK_CYFRA.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F831"
ccount = ccount + 1
End If

If CHK_NSE.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F832"
ccount = ccount + 1
End If

If chk_tft.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
    pathology.CMB_TEXT.AddItem "F033"
    ccount = ccount + 1

Else

    If chk_T3.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F034"
        ccount = ccount + 1
    End If
    
    If chk_T4.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F035"
        ccount = ccount + 1
    End If
    
    If chk_TSH.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F036"
        ccount = ccount + 1
    End If
End If

If chk_folate.Value = 1 Then
            If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F037"
        ccount = ccount + 1
End If

If chk_calcitonin.Value = 1 Then
            If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F040"
        ccount = ccount + 1
End If

'***** Shraddha 16/08/2011 *****

If chk_vitD.Value = 1 Then
            If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F899"
        ccount = ccount + 1
End If

If chk_freeT3.Value = 1 Then
            If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F058"
        ccount = ccount + 1
End If

If chk_freeT4.Value = 1 Then
            If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F059"
        ccount = ccount + 1
End If
'*************************************************


If chk_vitamin.Value = 1 Then
            If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F038"
        ccount = ccount + 1
End If

If chk_pth.Value = 1 Then
            If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F039"
        ccount = ccount + 1
End If

'**** markers



'**** Immunoelectrophoresis
'***** Shraddha 17/03/10 *****
If chk_IFE.Value = 1 Then
        pathology.CMB_TEXT.AddItem "F052"
        ccount = ccount + 1
End If

If chk_kappa.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F050"
    ccount = ccount + 1
End If

If chk_lambda.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F051"
    ccount = ccount + 1
End If

If chk_serligchains.Value = 1 Then
        pathology.CMB_TEXT.AddItem "F049"
        ccount = ccount + 1
    Else
    
'        If chk_kappa.Value = 1 Then
'            pathology.CMB_TEXT.AddItem "F050"
'            ccount = ccount + 1
'        End If
'
'        If chk_lambda.Value = 1 Then
'            pathology.CMB_TEXT.AddItem "F051"
'            ccount = ccount + 1
'        End If
        
        If chk_kappafree.Value = 1 Then
            pathology.CMB_TEXT.AddItem "F041"
            ccount = ccount + 1
        End If
        
        If chk_lambdafree.Value = 1 Then
            pathology.CMB_TEXT.AddItem "F042"
            ccount = ccount + 1
        End If
    End If
'*****************************

If chk_comleteSPE.Value = 1 Then

    pathology.CMB_TEXT.AddItem "F043"
    ccount = ccount + 1


Else
    If ImmunoelectrophoresisIEP.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
    'pathology.CMB_TEXT.AddItem "F816"
    pathology.CMB_TEXT.AddItem "F044"
    ccount = ccount + 1
    End If
    
    
    
    If chk_igs.Value = 1 Then
        pathology.CMB_TEXT.AddItem "F045"
        ccount = ccount + 1
    Else
    
        If chk_iga.Value = 1 Then
            pathology.CMB_TEXT.AddItem "F046"
            ccount = ccount + 1
        End If
        
        If chk_igm.Value = 1 Then
            pathology.CMB_TEXT.AddItem "F047"
            ccount = ccount + 1
        End If
        
        If chk_igg.Value = 1 Then
            pathology.CMB_TEXT.AddItem "F048"
            ccount = ccount + 1
        End If
    End If
    
'    If chk_serligchains.Value = 1 Then
'        pathology.cmb_text.AddItem "F049"
'        ccount = ccount + 1
'    Else
'
'        If chk_kappa.Value = 1 Then
'            pathology.cmb_text.AddItem "F050"
'            ccount = ccount + 1
'        End If
'
'        If chk_lambda.Value = 1 Then
'            pathology.cmb_text.AddItem "F051"
'            ccount = ccount + 1
'        End If
'
'        If chk_kappafree.Value = 1 Then
'            pathology.cmb_text.AddItem "F041"
'            ccount = ccount + 1
'        End If
'
'        If chk_lambdafree.Value = 1 Then
'            pathology.cmb_text.AddItem "F042"
'            ccount = ccount + 1
'        End If
'    End If
End If

'**** Immunoelectrophoresis

'**** Urine 24 Hours

If chk_bjp.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F977"
ccount = ccount + 1
End If

If chk_urkappa.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F053"
    ccount = ccount + 1
End If


If chk_urlambda.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F054"
    ccount = ccount + 1

End If

'**** Urine 24 Hours

'**** Estimation of Druge Level
'Commented on 12/03/10
'UnCommented on 17/03/10
'If chk_cyclos.Value = 1 And txt_Dosage.Text = "" Then
'If txt_Dosage.Text = "" Then
'    MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
    'Exit Sub
'End If
'UnCommented on 17/03/10
'Commented on 12/03/10
If chk_cyclos.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F833"
ccount = ccount + 1
End If

If chk_amik.Value = 1 And txt_Dosage.Text = "" Then
    MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
    Exit Sub
End If

If chk_amik.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F834"
ccount = ccount + 1
End If

If chk_vanc.Value = 1 And txt_Dosage.Text = "" Then
    MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
    Exit Sub
End If


If chk_vanc.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F835"
ccount = ccount + 1
End If

'If chk_metho.Value = 1 And txt_Dosage.Text = "" Then
'    MsgBox "PLEASE ENTER THE DOSAGE OF DRUGS DETAILS ", vbInformation
'    Exit Sub
'End If


If chk_metho.Value = 1 Then
    If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F836"
ccount = ccount + 1
End If
End If

'************** Shraddha Pandit 09/03/09 *********************
If CMB_NMETERIAL.Text = "" Then
    MsgBox "Please Select Nature of material"
    Exit Sub
End If

If Frame9.Visible = True Then
    If OPT_1.Value = False And OPT_2.Value = False And OPT_3.Value = False Then
        MsgBox "Please Select an option"
        Exit Sub
    End If
End If

If Frame10.Visible Then
    If serum_CEA.Value = 0 And serum_PSA.Value = 0 And serumFerritin.Value = 0 And serum_CA_125.Value = 0 And serumCA_153.Value = 0 And chk_crp.Value = 0 And serumB2_Microglobin.Value = 0 And CHK_NSE.Value = 0 And chk_addmark.Value = 0 And CHK_CYFRA.Value = 0 And serumB_HCG.Value = 0 And serum_AFP.Value = 0 And serumCA19_9.Value = 0 And txt_rem.Text = "" And chk_T3.Value = 0 And chk_T4.Value = 0 And chk_TSH.Value = 0 And chk_folate.Value = 0 And chk_vitamin.Value = 0 And chk_calcitonin.Value = 0 And chk_pth.Value = 0 And chk_vitD.Value = 0 And chk_freeT3.Value = 0 And chk_freeT4.Value = 0 Then
        MsgBox "Please Select atleast one value"
        Exit Sub
    End If
End If

If Frame3.Visible = True Then
    If chk_iga.Value = 0 And chk_igm.Value = 0 And chk_igg.Value = 0 And chk_kappa.Value = 0 And chk_kappafree.Value = 0 And chk_lambda.Value = 0 And chk_lambdafree.Value = 0 And ImmunoelectrophoresisIEP.Value = 0 And chk_IFE.Value = 0 Then
        MsgBox "Please Select atleast one value"
        Exit Sub
    End If
End If

If Frame8.Visible = True Then
    If chk_bjp.Value = 0 And chk_urkappa.Value = 0 And chk_urlambda.Value = 0 Then
        MsgBox "Please Select atleast one value"
        Exit Sub
    End If
End If

If Frame2.Visible = True Then
    If txt_Dosage.Text = "" And chk_cyclos.Value = 0 And chk_amik.Value = 0 And chk_vanc.Value = 0 And chk_metho.Value = 0 And txt_amt.Text = "" Then
        MsgBox "Please Select atleast one value"
        Exit Sub
    End If
End If
'*************************************************************
meterial = CMB_NMETERIAL.Text
'DDG = txt_Dosage.Text
'commented and added by Shraddha on 30/06/2011
'DDG = drug_dose
'AM = txt_rem
Unload Me
pathology.show
End Sub


Public Sub CHKCLEAR()
Dim obj As Object
For Each obj In TUMOURMARKER.Controls
If TypeOf obj Is CheckBox Then
obj.Enabled = False
End If
Next
End Sub


Private Sub Form_Load()
'MILIND Date 09/01/06
txt_Dosage.Enabled = False
txt_rem.Enabled = False
serum_AFP.Enabled = False
serum_CA_125.Enabled = False
serum_CEA.Enabled = False
serum_PSA.Enabled = False
serumB_HCG.Enabled = False
serumB2_Microglobin.Enabled = False
serumCA_153.Enabled = False
serumCA19_9.Enabled = False
serumFerritin.Enabled = False
'CHK_CYFRA.Enabled = False Commented on 24/06/10 by Shraddha
chk_crp.Enabled = False
CHK_NSE.Enabled = False
chk_addmark.Enabled = False
ImmunoelectrophoresisIEP.Enabled = False
chk_bjp.Enabled = False
chk_cyclos.Enabled = False
chk_amik.Enabled = False
chk_vanc.Enabled = False
chk_metho.Enabled = False

Dim ctrl As Object

For Each ctrl In Controls
    If TypeOf ctrl Is Frame Then
        ctrl.Enabled = False
    End If
Next

End Sub

'''''Public Sub CHKLCK()
'''''Dim obj As Object
'''''For Each obj In TUMOURMARKER.Controls
'''''If TypeOf obj Is CheckBox Then
'''''obj.Locked = True
'''''End If
'''''Next
'''''End Sub

Private Sub ImmunoelectrophoresisIEP_Click()
'If ImmunoelectrophoresisIEP.Value = 1 Then
'    chk_bjp.Enabled = False
'    txt_Dosage.Enabled = False
'    For Each c In Controls
'        If TypeOf c Is CheckBox Then
'            c.Enabled = False
'            ImmunoelectrophoresisIEP.Enabled = True
'        End If
'    Next
'
'Else
'    chk_bjp.Enabled = True
'    txt_Dosage.Enabled = True
'    For Each c In Controls
'    If TypeOf c Is CheckBox Then
'        c.Enabled = True
'    End If
'Next
'End If
''If ImmunoelectrophoresisIEP.Value = 1 Then chk_bjp.Value = 0
''If ImmunoelectrophoresisIEP.Value = 0 Then chk_bjp.Value = 0
End Sub

Private Sub ImmunoelectrophoresisIEP1_Click()

End Sub

Private Sub opt_1_Click()
    Frame2.Visible = False
    Frame3.Visible = False
    Frame10.Visible = True
    Frame10.Enabled = True
    Frame1.Visible = True
    
    Frame1.Enabled = True
    Frame6.Enabled = True
    Frame7.Enabled = True
    Frame10.Top = 3500
    Frame10.Left = 500
'    Frame1.Width = 4500
    enableframes

End Sub

Private Sub opt_2_Click()
    Frame10.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame5.Enabled = True
    Frame3.Top = 3500
    Frame3.Left = 2000
'    Frame1.Width = 4500
    ImmunoelectrophoresisIEP.Enabled = True
    enableframes
    
End Sub

Private Sub opt_3_Click()
    Frame10.Visible = False
    Frame3.Visible = False
    Frame2.Visible = False
    Frame2.Enabled = True
    Frame2.Top = 3500
    Frame2.Left = 500
    enableframes
End Sub
Public Sub enableframes()

''''for markers
If serum_CEA.Value = 1 Or serum_CA_125.Value = 1 Or serumB2_Microglobin.Value = 1 Or CHK_CYFRA.Value = 1 Or serum_AFP.Value = 1 Or serum_PSA.Value = 1 Or serumCA_153.Value = 1 Or chk_crp.Value = 1 Or CHK_NSE.Value = 1 Or serumB_HCG.Value = 1 Or serumCA19_9.Value = 1 Or serumFerritin.Value = 1 Or chk_addmark.Value = 1 Then
    Frame2.Enabled = False
    Frame3.Enabled = False
    
End If
 
'''''for serum
If chk_comleteSPE.Value = 1 Or ImmunoelectrophoresisIEP.Value = 1 Or chk_igs.Value = 1 Or chk_iga.Value = 1 Or chk_igm.Value = 1 Or chk_igg.Value = 1 Or chk_IFE.Value = 1 Or chk_serligchains.Value = 1 Or chk_kappa.Value = 1 Or chk_lambda.Value = 1 Or chk_kappafree.Value = 1 Or chk_lambdafree.Value = 1 Then
    Frame10.Enabled = False
    Frame2.Enabled = False
    
End If

''''for drug
If chk_cyclos.Value = 1 Or chk_amik.Value = 1 Or chk_vanc.Value = 1 Or chk_metho.Value = 1 Then
    Frame10.Enabled = False
    Frame3.Enabled = False
End If

For Each c In Controls
    If TypeOf c Is CheckBox Then
        c.Enabled = True
    End If
Next
End Sub
Private Sub serum_AFP_Click()
If serum_AFP.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serum_CA_125_Click()
If serum_CA_125.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serum_CEA_Click()
If serum_CEA.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True

    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serum_PSA_Click()
If serum_PSA.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serumB_HCG_Click()
If serumB_HCG.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serumB2_Microglobin_Click()
If serumB2_Microglobin.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serumCA_153_Click()
If serumCA_153.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serumCA19_9_Click()
If serumCA19_9.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub

Private Sub serumFerritin_Click()
If serumFerritin.Value = 1 Then
    chk_bjp.Enabled = False
    ImmunoelectrophoresisIEP.Enabled = False
    chk_cyclos.Enabled = False
    chk_amik.Enabled = False
    chk_vanc.Enabled = False
    chk_metho.Enabled = False
    txt_Dosage.Enabled = False
    '***** Shraddha 26/09/2011 *****
     '***** Shweta 26/04/2012 *****
'    Frame6.Visible = False
'    Frame7.Visible = False
    '*******************************
Else
    chk_bjp.Enabled = True
    ImmunoelectrophoresisIEP.Enabled = True
    chk_cyclos.Enabled = True
    chk_amik.Enabled = True
    chk_vanc.Enabled = True
    chk_metho.Enabled = True
    txt_Dosage.Enabled = True
    '***** Shraddha 26/09/2011 *****
    Frame6.Visible = True
    Frame7.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
    '*******************************
End If
End Sub


Private Sub TXT_HH_LostFocus()
If Trim(TXT_HH.Text) > 12 Then
    MsgBox ("Please enter time in 12 hours format..."), vbOKOnly
    TXT_HH.Text = ""
    TXT_HH.SetFocus
End If

End Sub

Private Sub TXT_MM_Change()
If Trim(TXT_MM.Text) > 60 Or Trim(TXT_MM.Text) = 60 Then
    MsgBox ("Minutes can not exceeds 60..."), vbOKOnly
    TXT_MM.Text = ""
    TXT_MM.SetFocus
End If

End Sub
