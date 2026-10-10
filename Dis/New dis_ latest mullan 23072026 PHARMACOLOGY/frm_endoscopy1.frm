VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_endoscopy1 
   Caption         =   "Endoscopy Report"
   ClientHeight    =   8895
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14820
   ControlBox      =   0   'False
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
   ScaleHeight     =   8895
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10200
      TabIndex        =   75
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10200
      TabIndex        =   74
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H8000000D&
      Height          =   2775
      Left            =   0
      TabIndex        =   67
      Top             =   3960
      Width           =   11775
      Begin MSMask.MaskEdBox msk_procdate 
         Height          =   375
         Left            =   1800
         TabIndex        =   87
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmb_type 
         Height          =   315
         ItemData        =   "frm_endoscopy1.frx":0000
         Left            =   6720
         List            =   "frm_endoscopy1.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cmb_documentation 
         Height          =   315
         ItemData        =   "frm_endoscopy1.frx":0004
         Left            =   6720
         List            =   "frm_endoscopy1.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   22
         Text            =   "None"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txt_instrument 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   23
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txt_complication 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Text            =   "Nil"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.ComboBox cmb_procdone 
         Height          =   315
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txt_sampling 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton cmd_dis 
         Caption         =   "&DIS"
         Height          =   495
         Left            =   10320
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txt_report 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txt_procname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txt_impression 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txt_proccode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lbl_procdate 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Date"
         Height          =   195
         Left            =   330
         TabIndex        =   88
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label lbl_type 
         Caption         =   "Type"
         Height          =   195
         Left            =   6120
         TabIndex        =   86
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_complication 
         Caption         =   "Complication"
         Height          =   255
         Left            =   690
         TabIndex        =   83
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lbl_instrument 
         AutoSize        =   -1  'True
         Caption         =   "Instrument"
         Height          =   195
         Left            =   5640
         TabIndex        =   82
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl_procdone 
         AutoSize        =   -1  'True
         Caption         =   "Performed By"
         Height          =   195
         Left            =   5400
         TabIndex        =   81
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lbl_report 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1200
         TabIndex        =   79
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lbl_tissue 
         Caption         =   "Tissue Sampling"
         Height          =   255
         Left            =   330
         TabIndex        =   78
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lbl_procname 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Name"
         Height          =   195
         Left            =   360
         TabIndex        =   72
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lbl_impression 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5640
         TabIndex        =   71
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lbl_documentation 
         AutoSize        =   -1  'True
         Caption         =   "Documentation"
         Height          =   195
         Left            =   5280
         TabIndex        =   70
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label lbl_addendum 
         AutoSize        =   -1  'True
         Caption         =   "Addendum "
         Height          =   195
         Left            =   840
         TabIndex        =   69
         Top             =   2160
         Width           =   960
      End
      Begin VB.Label lbl_proccode 
         Caption         =   "Procedure Code"
         Height          =   255
         Left            =   5160
         TabIndex        =   68
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   63
      Top             =   6720
      Width           =   11775
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         Left            =   8880
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cmb_reportby 
         Height          =   315
         Left            =   4680
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cmb_enterby 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TXT_COMMENTS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   10215
      End
      Begin VB.Label Label8 
         Caption         =   "To be Finalized By"
         Height          =   255
         Left            =   7200
         TabIndex        =   73
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "Reported By"
         Height          =   195
         Left            =   3600
         TabIndex        =   66
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lbl_enter 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
         Height          =   195
         Left            =   240
         TabIndex        =   65
         Top             =   720
         Width           =   720
      End
      Begin VB.Label LBL_COMM 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   64
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame COMMAND 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   62
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7320
         MouseIcon       =   "frm_endoscopy1.frx":0008
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10320
         MouseIcon       =   "frm_endoscopy1.frx":0312
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8880
         MouseIcon       =   "frm_endoscopy1.frx":061C
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5880
         MouseIcon       =   "frm_endoscopy1.frx":0926
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "frm_endoscopy1.frx":0C30
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "frm_endoscopy1.frx":0F3A
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MouseIcon       =   "frm_endoscopy1.frx":1244
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         MouseIcon       =   "frm_endoscopy1.frx":154E
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame common 
      Height          =   2895
      Left            =   0
      TabIndex        =   30
      Top             =   1080
      Width           =   11775
      Begin VB.TextBox txt_lasttreat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8640
         TabIndex        =   84
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txt_provdiag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   42
         Top             =   720
         Width           =   1215
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_microno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   40
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   39
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   36
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   34
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   33
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   32
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   2160
         Width           =   3855
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
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
      Begin VB.Label lbl_lasttreat 
         AutoSize        =   -1  'True
         Caption         =   "Previous Treatment"
         Height          =   192
         Left            =   6960
         TabIndex        =   85
         Top             =   2160
         Width           =   1668
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   80
         Top             =   2520
         Width           =   1830
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   60
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   58
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   57
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Endoscopy No."
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   8880
         TabIndex        =   56
         Top             =   600
         Width           =   1284
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   55
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   54
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   53
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   52
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   51
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   49
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   48
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   47
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   46
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2160
         Width           =   2130
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
      Left            =   240
      TabIndex        =   89
      Top             =   0
      Width           =   9975
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9600
      TabIndex        =   77
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9600
      TabIndex        =   76
      Top             =   720
      Width           =   540
   End
   Begin VB.Label LBL_ENDOCAPTION 
      AutoSize        =   -1  'True
      Caption         =   "REPORT"
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
      TabIndex        =   61
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frm_endoscopy1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callstatus          As Boolean
Dim objWord             As Word.Application
Dim objDoc              As Word.Document
Dim i                   As Integer
Dim sel_deptqr          As rdo.rdoQuery
Dim sel_deptrs          As rdo.rdoResultset
Dim SEL_REPORT_QR       As rdo.rdoQuery
Dim SEL_REPORT_RS       As rdo.rdoResultset
Dim sel_moadlityqr      As rdo.rdoQuery
Dim sel_modalityrs      As rdo.rdoResultset
Dim sel_prdiag_qr       As rdo.rdoQuery 'to get provisional diagnosis
Dim sel_prdiag_rs       As rdo.rdoResultset
Dim upd_6_qr            As rdo.rdoQuery
Dim upd_1_qr            As rdo.rdoQuery 'used for updating addendum
Dim sel_addendum_qr     As rdo.rdoQuery 'used for addendum
Dim sel_addendum_rs     As rdo.rdoResultset
Dim sel_FINAL_qr        As rdo.rdoQuery ' to get status of commit flag
Dim sel_ENTER_qr        As rdo.rdoQuery ' to get ENTER BY USER LIST
Dim SEL_dreq_staff_qr   As rdo.rdoQuery ' to get staff detail
Dim sel_ENTER_rs        As rdo.rdoResultset ' to  get ENTER BY USER LIST
Dim sel_FINAL_rs        As rdo.rdoResultset ' to  get FINAL BY USER LIST
Dim SEL_dreq_staff_rs   As rdo.rdoResultset ' to get staff detail
Dim req_qr              As rdo.rdoQuery
Dim req_rs              As rdo.rdoResultset
Dim upd_7_qr            As rdo.rdoQuery     'TO UPDATE RECORD IN DEDREPORT TABLE
Dim ins_1_qr            As rdo.rdoQuery
Dim ins_2_qr            As rdo.rdoQuery    'to insert in addendum table
Dim INS_3_qr            As rdo.rdoQuery
Dim SELRREQUISITN_QR    As rdo.rdoQuery
Dim SELRREQUISITN_RS    As rdo.rdoResultset
Dim sel_micro_qr        As rdo.rdoQuery ' to  get ENDOSCOPY number
Dim sel_micro_rs        As rdo.rdoResultset ' to  get ENDOSCOPY number
Dim REFDR_QR            As rdo.rdoQuery
Dim REFDR_RS            As rdo.rdoResultset
Dim upd_11_qr           As rdo.rdoQuery
Dim upd_8_qr            As rdo.rdoQuery
Dim sel_forupd1_qr      As rdo.rdoQuery
Dim sel_forupd1_rs      As rdo.rdoResultset
Dim SEL_MODALITY_QR     As rdo.rdoQuery
Dim SEL_MODALITY_RS     As rdo.rdoResultset
Dim SEL_detail_qr       As rdo.rdoQuery
Dim SEL_detail_rs       As rdo.rdoResultset
Dim sel_bed_qr          As rdo.rdoQuery
Dim sel_bed_rs          As rdo.rdoResultset
Dim sel_proccode_qr     As rdo.rdoQuery
Dim sel_proccode_rs     As rdo.rdoResultset 'for selecting procedure code
Dim sel_prevdia_qr      As rdo.rdoQuery 'used for previous diagnosis
Dim sel_prevdia_rs      As rdo.rdoResultset
Dim mod_reqno           As String 'used to store reqno by calling req_chglen method
Dim new_flag            As Boolean
Dim mod_flag            As Boolean
Dim VIEW_FLAG           As Boolean
Dim no_req              As Boolean 'to check before printing report whether req is existing or not
Dim ADDENDUMFLAG        As Boolean 'USED FOR INSERTING ADDENDUM AFTER COMMIT
Dim S1                  As String
Dim S2                  As String
Dim label               As String

Dim SEL_MODALITYFLAG As Boolean
''******SECURITY***********

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String

Dim REPWTRNRS As rdoResultset
Dim REPWTRN_QR As rdoQuery

Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset

Dim seldedrep_qr As rdoQuery
Dim SELDEDREP_RS As rdoResultset

Private Sub cmb_dept_Click()
 Dim deptno As String
 Dim dept As Integer
Dim l As Integer
On Error GoTo MSG:
    If Trim(cmb_dept) = "" Then
        Exit Sub
    End If

    dept = InStr(cmb_dept.Text, "$")
    txt_deptcode.Text = Mid(cmb_dept.Text, dept + 1, 3)
   If SEL_MODALITYFLAG = False Then
        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
        Set sel_modalityrs = sel_moadlityqr.OpenResultset(2, 3)
        SEL_MODALITYFLAG = True
   Else
   
        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
        sel_modalityrs.MoreResults
        sel_modalityrs.Requery
    End If
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
    End If
Exit Sub
MSG:
MsgBox Error(Err)

End Sub

''******SECURITY***********
Private Sub cmb_modality_Click()
On Error GoTo MSG
    i = InStr(cmb_modality.Text, "$")
    txt_modcode.Text = Mid(cmb_modality.Text, i + 1, 2)
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_procdone_lostfocus()
    cmb_reportby.Text = cmb_procdone.Text
End Sub

Private Sub cmb_reqno_Click()
'On Error GoTo MSG
    clearlowerpanel
    If cmb_reqno.ListIndex > 0 And (cmb_reqno.Text) = "" Or Len(cmb_reqno.Text) < 1 Then
        MsgBox "Select Requisition First", vbOKOnly
        btnenabl "INI"
    Exit Sub
    End If
        mod_reqno = req_chglen(cmb_reqno) 'put change value to this var so that it used later on
        GETUSERLIST
        GETFINALLIST
        GETREPORTEDLIST
    If Opt_patient.Value Then
        If new_flag Then
            getupperpanel
            enable_field
        ElseIf VIEW_FLAG Then
            getupperpanel
            RETRIVE
            cmb_type.Text = "Diagnostic"
        End If
        no_req = True
    ElseIf opt_staff.Value Then 'FOR STAFF MEMBER
        'if staff call staff procedure
        getupperpanel
        STAFF
        no_req = True
    End If 'END OF TOP IF
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_cancel_Click()
On Error GoTo MSG
    btnenabl ("INI")
    clearupperpanel
    clearlowerpanel
    clearlastpanel
    disable_field
    CMD_NEW.SetFocus
    cmb_dept.Enabled = True
    cmb_modality.Enabled = True
    cmb_dept.Enabled = True
    cmb_reqno.Enabled = True
    MaskEdBox1.Enabled = True
    set_flag "Ni"
    SEL_MODALITYFLAG = False
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_commit_Click()
'On Error GoTo Errhandler
    If Len(Trim(cmb_finalby.Text)) = 0 Then
            MsgBox "The Finalized by field is EMPTY.", vbCritical
            'cmb_finalby.SetFocus
            Exit Sub
        End If
    Dim i As Integer
    i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
    If i = vbYes Then
    On Error GoTo MSG:
    
    ''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
            cmd_commit.SetFocus
            Exit Sub
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
            If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
                '  If loginusrid <> curusrdetrs("usr_id") Then
                MsgBox "Sorry! This Report is not to be Finalized by You"
                cmd_commit.SetFocus
                Exit Sub
            End If
        End If
        If CURUSRDET_RS("COMMIT") <> "YES" Then
            MsgBox "Sorry! You are not authorised to Commit the Report of this Modality"
            cmd_commit.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
    End If
''********SECURITY***************************

     rdb.BeginTrans
        upd_6_qr(0) = Left(Trim(cmb_finalby.Text), 30)
        upd_6_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        upd_6_qr(2) = "F"
        upd_6_qr(3) = "Y"
        upd_6_qr(4) = mod_reqno
        upd_6_qr.Execute
        
        
            'SHWETA UPDATE RTIME LOG 02-04-2013
      If UPDRTIMELOG(Trim(mod_reqno), "COMMIT") = True Then
      Else
      GoTo MSG
      End If
        'TILL HERE
        MsgBox "Record is Committed Successfully."
    End If
    rdb.CommitTrans
    btnenabl "INI"
    
    cmd_print.Enabled = True
    Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub cmd_modify_Click()
On Error GoTo MSG
    btnenabl ("NEW")
    new_flag = False
    VIEW_FLAG = False
    mod_flag = True
    If Len(Trim(MaskEdBox1.Text)) = 0 Then
        MsgBox "Case Number should not be empty!", vbCritical
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    If Len(Trim(cmb_reqno.Text)) = 0 Then
        MsgBox "Requisition Number should not be empty!", vbCritical
        If cmb_reqno.Enabled = True Then
            cmb_reqno.SetFocus
         '   SendKeys "^{f4}"
        End If
        Exit Sub
    End If
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = mod_reqno
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            If REFDR_RS.RowCount > 0 Then
                If REFDR_RS("commitflag") = "Y" Then
                  ADDENDUMFLAG = True
                End If
            End If
    If ADDENDUMFLAG Then
        'disable_field
        txt_addendum.Enabled = True
        txt_addendum.Visible = True
        lbl_addendum.Visible = True
        
    Else
        enable_field
    End If
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_new_Click()

Dim L1 As Integer
On Error GoTo MSG
    enable_field
    btnenabl ("NEW")
    new_flag = True
    VIEW_FLAG = False
    mod_flag = False
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
    Wend
    For i = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(i)) = "DIV OF DIG DIS & CLI NUTRITION$CBZ" Then
                cmb_dept.ListIndex = i
        End If
    Next i
     If Trim(cmb_dept) = "" Then
        Exit Sub
    End If
    cmb_enterby.Text = cmb_enterby.List(0)
    'cmb_enterby.Enabled = False
    For L1 = 0 To cmb_finalby.ListCount
        If Trim(cmb_enterby.List(0)) = cmb_finalby.List(L1) Then
            cmb_finalby.Text = cmb_finalby.List(L1)
        End If
    Next L1
    'cmb_finalby.Text = cmb_finalby.List(0)
    cmb_reportby.Text = cmb_reportby.List(0)
    cmb_procdone.Text = cmb_procdone.List(0)
        
    i = InStr(cmb_dept.Text, "$")
    txt_deptcode.Text = Mid(cmb_dept.Text, i + 1, 3)
    sel_moadlityqr(0) = Trim(txt_deptcode.Text)
    Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
        sel_modalityrs.Close
    Else
        sel_modalityrs.Close
    End If
    For i = 0 To (cmb_modality.ListCount - 1)
          If Trim(cmb_modality.List(i)) = "ENDOSCOPY$ED" Then
             cmb_modality.ListIndex = i
          End If
    Next i
sel_deptrs.Close
set_flag "NEW"
Opt_patient.Value = True
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
On Error GoTo MSG
    DoEvents
     If no_req Then 'check for entered selected req no.
        REFDR_QR(0) = mod_reqno
        REFDR_QR(1) = MaskEdBox1.FormattedText
        Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        If ADDENDUMFLAG Then
            'commented and added by shweta 23-sept-2010
            'crystalReport1.ReportFileName = "c:\DISREPORTSEC\Endoscopy_add_Report.rpt"
            CrystalReport1.ReportFileName = "c:\DISREPORTSECLF\Endoscopy_add_Report.rpt"
            ADDENDUMFLAG = False
        Else
            'commented and added by shweta 23-sept-2010
            'CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Endoscopy_Report.rpt"
            CrystalReport1.ReportFileName = "c:\DISREPORTSECLF\Endoscopy_Report.rpt"
        End If
        CrystalReport1.Formulas(1) = "proFIN = 'ENDOSCOPY REPORT'"
        CrystalReport1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enterby.Text) & "'"
        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        CrystalReport1.SelectionFormula = "{ENDOSCPYLF.REQNNO} = '" & mod_reqno & "'"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(7) = "pathnumber='" & txt_microno.Text & "'"
        If Not Trim(txt_ward.Text) = "" Then
             If InStr(1, txt_ward.Text, "WARD") = 0 Then
                CrystalReport1.Formulas(8) = "ward='" & txt_ward.Text & "'"
             Else
                 CrystalReport1.Formulas(8) = "ward='" & Trim(Mid(txt_ward.Text, 1, InStr(1, txt_ward.Text, "WARD") - 1)) & "'"
             End If
        Else
            CrystalReport1.Formulas(8) = "ward="
        End If
        CrystalReport1.Formulas(9) = "NAM='" & txt_name.Text & "'"
        CrystalReport1.Formulas(10) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
        If Opt_patient.Value Then
            CrystalReport1.Formulas(11) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
        Else
            CrystalReport1.Formulas(11) = "ST='Staff /" & txt_status.Text & "'"
        End If
        CrystalReport1.Formulas(12) = "label='" & label & "'"
        If Len(Trim(txt_report.Text)) > 2 Then
            CrystalReport1.Formulas(13) = "report='" & "Print" & "'"
        Else
            CrystalReport1.Formulas(13) = "report="
        End If
          
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
     Else
        MsgBox "Please select valid Requisition No.", vbOKOnly
     End If
Exit Sub
MSG:
    MsgBox Err.Description
DoEvents
End Sub

Private Sub cmd_save_Click()
On Error GoTo MSG
        If Len(Trim(cmb_enterby)) = 0 Then
            MsgBox "The Enter by field is EMPTY.", vbCritical
            cmb_enterby.SetFocus
        Exit Sub
        End If
'        If Len(Trim(cmb_finalby.Text)) = 0 Then
'            MsgBox "The Finalized by field is EMPTY.", vbCritical
'            cmb_finalby.SetFocus
'            Exit Sub
'        End If
        If Len(Trim(cmb_reportby.Text)) = 0 Then
            MsgBox "The Report by field is empty."
             cmb_reportby.SetFocus
             Exit Sub
        End If
If Not ADDENDUMFLAG Then
    If new_flag Then    'if new record then use insert query
        'write to save insert query
            btnenabl "SAVE"
            new_flag = False
            insert1
            disable_field
            'cmd_commit.Enabled = True
    ElseIf mod_flag Then 'else if modifying a record then use update query
     'write to update record update query
        mod_flag = False
        btnenabl "SAVE"
        
        seldedrep_qr(0) = mod_reqno
        Set SELDEDREP_RS = seldedrep_qr.OpenResultset(1, 3)
        
        If SELDEDREP_RS.RowCount > 0 Then
            MODIFY
        ElseIf SELDEDREP_RS.RowCount = 0 Then
            insert1
        End If
        disable_field
    End If
Else
    'add addendum
    ADDENDUM "INSERT"
    txt_addendum.Enabled = False
End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
new_flag = False
VIEW_FLAG = True
mod_flag = False
Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    'get the result
    While Not sel_deptrs.EOF
            cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
            sel_deptrs.MoveNext
    Wend
sel_deptrs.Close
        For i = 0 To (cmb_dept.ListCount - 1)
            If Trim(cmb_dept.List(i)) = "DIV OF DIG DIS & CLI NUTRITION$CBZ" Then
                cmb_dept.ListIndex = i
            End If
        Next i

    i = InStr(cmb_dept.Text, "$")
    txt_deptcode.Text = Mid(cmb_dept.Text, i + 1, 3)
    sel_moadlityqr(0) = Trim(txt_deptcode.Text)
    Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
        sel_modalityrs.Close
    Else
        sel_modalityrs.Close
    End If
    For i = 0 To (cmb_modality.ListCount - 1)
          If Trim(cmb_modality.List(i)) = "ENDOSCOPY$ED" Then
             cmb_modality.ListIndex = i
          End If
    Next i
set_flag "VIEW"
Opt_patient.Value = True
cmd_print.Enabled = True
cmb_dept.SetFocus
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_back_Click()
On Error GoTo MSG
    clearlastpanel
    clearlowerpanel
    clearupperpanel
    set_flag "All"
    btnenabl "INI"
    Me.Hide
    DIS_MAIN.show
    Exit Sub
MSG:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo MSG
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
msk_procdate.Text = Format(DATE11, "DD/MM/YYYY")
initialize_combo
set_flag "asl"
disable_field
'new_flag = False
'mod_flag = False
'view_flag = False
'query definition starts here
'*******************************************************************************************
'************       ALL SELECT QUERY            *****************************************
'*******************************************************************************************


'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1BIO", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = Trim(LOGINUSRNAME)
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Enabled = False
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNSTR", str)
If Trim(USRDEPT) = "ALL" Then
    REPWTRN_QR(0) = "CBZ"
Else
    REPWTRN_QR(0) = Trim(USRDEPT)
End If
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)

If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    'cmb_addendumby.clear
    cmb_reportby.clear
    cmb_finalby.clear
    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 1 Then
            cmb_reportby.AddItem Trim(REPWTRNRS("USERNAME"))
            
            cmb_procdone.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        If REPWTRNRS("USRTYPE") >= 4 Then
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        REPWTRNRS.MoveNext
    Wend
    'cmb_reportby.Text = cmb_reportby.List(0)
    'cmb_finalby.Text = cmb_finalby.List(0)
    'cmb_procdone.Text = cmb_procdone.List(0)
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BIO", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

'*********SECURITY**********************************


'for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("endo_select1", str)

'FOR MODALITY SELECTION
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("endo_SELECT2", str)
SEL_MODALITYFLAG = False

'to get entered by name
str = "select * from dislib.dissec where enterby='YES'AND MODALITY='ED'"
Set sel_ENTER_qr = rdb.CreateQuery("endo_select4", str)

'to get PROCEDURE by name
str = "select * from dislib.dissec where GROSSBY='YES'AND MODALITY='ED'"
Set SEL_REPORT_QR = rdb.CreateQuery("endo_select20", str)

'to get final by list name
str = "select * from dislib.dissec where FINALby='YES' AND MODALITY='ED'"
Set sel_FINAL_qr = rdb.CreateQuery("endo_select5", str)

'for staff query
str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
Set SEL_dreq_staff_qr = rdb.CreateQuery("endo_select6", str)

'select requisition no.
str = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? and DEPT_CODE=? AND REQNSTAT='E' ORDER BY REQNDT DESC"
Set req_qr = rdb.CreateQuery("endo_select7", str)

'for view mode
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND  DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
Set SELRREQUISITN_QR = rdb.CreateQuery("endo_SELECT8", str)

'***************
str = "SELECT * from  DISLIB.Drequisitn where reqno=?"
Set sel_micro_qr = rdb.CreateQuery("ENDo_SELECT11", str)

'TO GET MODALITY DESC.
str = "SELECT MODLITYDES,MODALITYCD FROM RISLIB.RMODALITY WHERE DEPT_CODE=? AND MODALITYCD=?"
Set SEL_MODALITY_QR = rdb.CreateQuery("ENDo_SELECT12", str)

'select other detail for patient
str = "select * from PATOOLSLIB.ta101 where case_no =?"
Set SEL_detail_qr = rdb.CreateQuery("ENDo_select13", str)

'for referred by dr. and commit flag
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
Set REFDR_QR = rdb.CreateQuery("endo_select14", str)

'for bed and ward retriving
str = "select * from PATOOLSLIB.ma104 where case_no =?"
Set sel_bed_qr = rdb.CreateQuery("endo_select15", str)

'FOR SELECTING WHICH PROCEDURE HE HAS SELECTED
str = "SELECT SER_CODE,SER_DET FROM DISLIB.MA101 WHERE SER_CODE IN(SELECT SER_CODE FROM RISLIB.RREQUISITN WHERE REQNNO= ?)"
Set sel_proccode_qr = rdb.CreateQuery("endo_select16", str)

'to get provisional diagnosis
str = "select PRVDIAGDES from patoolslib.ta1012 where case_no=?"
Set sel_prdiag_qr = rdb.CreateQuery("endo_select17", str)

'to get addendum for entered reqno
str = "select * from dislib.addendumed  where reqno=? "
Set sel_addendum_qr = rdb.CreateQuery("endo_select18", str)

'TO GET PREVIOUS DIAGNOSIS
str = "select prvdiagdes from patoolslib.ta1012 where case_no=?"
Set sel_prevdia_qr = rdb.CreateQuery("endo_select19", str)

'GET DATA TO MODIFY THE RECORD
str = "SELECT reqno,report,impression,docment,dttype,compcation,instrument,pperformed,remarks,fld1,TISSAMP from  DISLIB.Dedreport where reqno=?"
Set sel_forupd1_qr = rdb.CreateQuery("endo_SELECT10", str)


'*******************************************************************************************
'************       ALL UPDATE QUERY            *****************************************
'*******************************************************************************************

'update reqstat if he/she successfully entered data
str = "update rislib.rrequisitn set reqnstat='T' where caseno=? and reqnno=?"
Set upd_11_qr = rdb.CreateQuery("endo_update0", str)

'to update rislib for enterby and enter by
str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
Set upd_8_qr = rdb.CreateQuery("endo_update1", str)

'update commit flag
str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set upd_6_qr = rdb.CreateQuery("endo_update2", str)

'UPDATE RECORD
str = "UPDATE DISLIB.DEDREPORT SET REPORT=?,IMPRESSION=?,DOCMENT=?,DTTYPE=?,COMPcATION=?,INSTRUMENT=?,PPERFORMED=?,REMARKS=?,FLD1=?,TISSAMP=? WHERE REQNO=?"
Set upd_7_qr = rdb.CreateQuery("ENDO_UPDATE3", str)
'*******************************************************************************************
'************       ALL INSERT QUERY            *****************************************
'*******************************************************************************************
str = "INSERT INTO DISLIB.DEDREPORT(REQNO,REPORT,IMPRESSION,DOCMENT,DTTYPE,COMPCATION,INSTRUMENT,PPERFORMED,REMARKS,FLD1,TISSAMP)VALUES(?,?,?,?,?,?,?,?,?,?,?)"
Set ins_1_qr = rdb.CreateQuery("ENDO_INSERT1", str)

str = "INSERT INTO DISLIB. (REQNO,)VALUES(?,?)"
Set ins_2_qr = rdb.CreateQuery("ENDO_INSERT2", str)

str = "SELECT * FROM DISLIB.DEDREPORT WHERE REQNO=?"
Set seldedrep_qr = rdb.CreateQuery("DEDRPT", str)

'to insert new addendum
str = "INSERT INTO DISLIB.ADDENDUMED(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
Set INS_3_qr = rdb.CreateQuery("endo_insert3", str)

Exit Sub
MSG:
    MsgBox Err.Description
End Sub



Private Sub initialize_combo()
On Error GoTo MSG
cmb_documentation.AddItem "None"
cmb_documentation.AddItem "Picture"
cmb_documentation.AddItem "Mpeg"
cmb_documentation.ListIndex = -1
cmb_type.AddItem "Diagnostic"
cmb_type.AddItem "Therapeutic"
cmb_type.ListIndex = -1
For i = 0 To (cmb_documentation.ListCount - 1)
    If Trim(cmb_documentation.List(i)) = "None" Then
                cmb_documentation.ListIndex = i
        End If
Next i
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub clearupperpanel()
On Error GoTo MSG
    MaskEdBox1.Text = ""
    cmb_reqno.clear
    cmb_modality.clear
    txt_deptcode.Text = ""
    cmb_dept.clear
    txt_microno.Text = ""
    txt_name.Text = ""
    txt_age.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_status.Text = ""
    txt_ward.Text = ""
    txt_bedno.Text = ""
    txt_pridoc.Text = ""
    txt_modcode.Text = ""
    txt_refdoc.Text = ""
    txt_lasttreat.Text = ""
Exit Sub
MSG:
        MsgBox Err.Description
End Sub

Private Sub clearlowerpanel()
On Error GoTo MSG
    txt_procname.Text = ""
    txt_proccode.Text = ""
    txt_report.Text = ""
    txt_impression.Text = ""
    txt_addendum.Text = ""
    cmb_documentation.Text = ""
    cmb_procdone.Text = ""
    cmb_type.Text = ""
    txt_instrument.Text = ""
    txt_complication.Text = "Nil"
    'cmb_type.Text = ""
    txt_sampling.Text = ""
    txt_proccode.Text = ""
    msk_procdate.Mask = "##/##/####"
    msk_procdate.Text = Format(DATE11, "DD/MM/YYYY")
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub clearlastpanel()
On Error GoTo MSG
    txt_comments.Text = ""
    cmb_enterby.Text = ""
    cmb_reportby.Text = ""
    cmb_finalby.Text = ""
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("SAVE") = "YES" Then
            cmd_save.Enabled = True
        End If
    End If
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    'cmd_fullreport.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("MODIFY") = "YES" Then
            cmd_modify.Enabled = True
        Else
            cmd_modify.Enabled = False
        End If
    End If
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    'cmd_fullreport.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("COMMIT") = "YES" Then
            cmd_commit.Enabled = True
        Else
            cmd_commit.Enabled = False
        End If
    End If
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    'cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    'cmd_fullreport.Enabled = False
ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    cmd_modify.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    'cmd_fullreport.Enabled = False
ElseIf cap = "print" Then
     cmd_save.Enabled = False
     cmd_back.Enabled = True
     cmd_modify.Enabled = False
     cmd_print.Enabled = False
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False
End If

'If cap = "NEW" Then
'    cmd_save.Enabled = True
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'ElseIf cap = "VIEW" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'ElseIf cap = "SAVE" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = True
'    cmd_view.Enabled = True
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'ElseIf cap = "COMMIT" Then
'    cmd_save.Enabled = False
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'ElseIf cap = "print" Then
'    cmd_save.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_cancel.Enabled = True
'    CMD_NEW.Enabled = False
'ElseIf cap = "cancel" Then
'    cmd_save.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_cancel.Enabled = True
'    CMD_NEW.Enabled = False
'End If
End Sub

Private Sub set_flag(flag As String) 'used to flag status
If Trim(flag) = "NEW" Then
    new_flag = True
    VIEW_FLAG = False
    mod_flag = False
ElseIf Trim(flag) = "VIEW" Then
    new_flag = False
    VIEW_FLAG = True
    mod_flag = False
ElseIf Trim(flag) = "MOD" Then
    new_flag = False
    VIEW_FLAG = False
    mod_flag = True
Else
    new_flag = False
    VIEW_FLAG = False
    mod_flag = False
    ADDENDUMFLAG = False
End If
End Sub

Private Sub update_rislib()
On Error GoTo MSG
    upd_8_qr(0) = cmb_enterby.Text
    upd_8_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_8_qr(2) = cmb_reportby.Text
    upd_8_qr(3) = "T"
    upd_8_qr(4) = cmb_finalby.Text
    upd_8_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_8_qr(6) = Trim(mod_reqno)
    upd_8_qr.Execute
    MsgBox "Record Has Been Saved Successfully", vbOKOnly
    'SHWETA UPDATE RTIME LOG 02-04-2013
      If UPDRTIMELOG(Trim(mod_reqno), "SAVE") = True Then
  Else
  GoTo MSG
  End If
    ''TILL HERE
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub GETUSERLIST()
On Error GoTo MSG
'    Set sel_ENTER_rs = sel_ENTER_qr.OpenResultset(1, 3)
'    If sel_ENTER_rs.RowCount > 0 Then
'        cmb_enterby.clear
'            While Not sel_ENTER_rs.EOF
'                cmb_enterby.AddItem Trim(sel_ENTER_rs("CCNAME"))
'                sel_ENTER_rs.MoveNext
'            Wend
'        sel_ENTER_rs.Close
'    Else
'        sel_ENTER_rs.Close
'    End If
    'INSERT
'    For i = 0 To (cmb_enteryby.ListCount - 1)
'        If Trim(cmb_enteryby.List(i)) = "Mrs. Kalpana Varadkar" Then
'                cmb_enteryby.ListIndex = i
'        End If
'    Next i
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub GETFINALLIST()
On Error GoTo MSG
'    Set sel_FINAL_rs = sel_FINAL_qr.OpenResultset(1, 3)
'    If sel_FINAL_rs.RowCount > 0 Then
'        cmb_finalby.clear
'            While Not sel_FINAL_rs.EOF
'                cmb_finalby.AddItem Trim(sel_FINAL_rs("CCNAME"))
'                sel_FINAL_rs.MoveNext
'            Wend
'            sel_FINAL_rs.Close
'    Else
'        sel_FINAL_rs.Close
'    End If
'    For i = 0 To (cmb_final.ListCount - 1)
'        If Trim(cmb_final.List(i)) = "Dr.Rohini Kelkar" Then
'                cmb_final.ListIndex = i
'        End If
'    Next i
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub GETREPORTEDLIST()
On Error GoTo MSG
'    Set SEL_REPORT_RS = SEL_REPORT_QR.OpenResultset(1, 3)
'    If SEL_REPORT_RS.RowCount > 0 Then
'        cmb_reportby.clear
'        cmb_procdone.clear
'            While Not SEL_REPORT_RS.EOF
'                cmb_reportby.AddItem Trim(SEL_REPORT_RS("CCNAME"))
'                cmb_procdone.AddItem Trim(SEL_REPORT_RS("CCNAME"))
'                SEL_REPORT_RS.MoveNext
'            Wend
'            SEL_REPORT_RS.Close
'    Else
'        SEL_REPORT_RS.Close
'    End If
'    For i = 0 To (cmb_final.ListCount - 1)
'        If Trim(cmb_final.List(i)) = "Dr.Rohini Kelkar" Then
'                cmb_final.ListIndex = i
'        End If
'    Next i
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub MaskEdBox1_LostFocus()
On Error GoTo MSG:
If Trim(MaskEdBox1.Text) = "" Then
    btnenabl "INI"
    Exit Sub
End If
If Opt_patient.Value Then
'If Opt_patient.Value = True Then
   '' MaskEdBox1.Text = UCase(padzero(5, "C", MaskEdBox1.Text, 3)) ''commented by Monika on 05/12/2022
   '''' Commented PADZERO FUNCTION INSTEAD USED NEWNOPAD FUNCTION
'cno_pad = padzero(5, "C", msk_caseno.Text, 3)
    



    
  If IsNumeric(Left(Trim(MaskEdBox1.Text), 1)) Then
          Dim Ks As Integer
        Dim str1s As String
        Dim strs As String
        Dim cno_pad As String
            strs = Trim(MaskEdBox1.Text)
            Ks = Len(strs)
            str1s = Mid(strs, 5, Ks)
              Dim sel_CASE_rs  As rdo.rdoResultset
              Set sel_CASE_rs = rdb.OpenResultset(" SELECT * FROM PATOOLSLIB.NEWCASETBL WHERE CASE_CODE = '" & Left(UCase(MaskEdBox1.FormattedText), 4) & "' ", 1, 1)
              If Not sel_CASE_rs.EOF Then
                    Dim CSNO As String
                    CSNO = Trim(sel_CASE_rs(1))
                    CSNO = CSNO & str1s
                     cno_pad = NEWNOpad(CSNO)
              Else
               cno_pad = NEWNOpad(UCase(MaskEdBox1))
              End If
    
  Else
    cno_pad = NEWNOpad(UCase(MaskEdBox1))
  End If
     MaskEdBox1.Text = cno_pad
''''''''''''''''''''till here (01/12/2022)'''''''''
    If new_flag Then
            cmb_reqno.clear
            req_qr(0) = MaskEdBox1.FormattedText
            req_qr(1) = Trim(txt_modcode.Text)
            req_qr(2) = Trim(txt_deptcode.Text)
            Set req_rs = req_qr.OpenResultset(1, 3)
        If req_rs.RowCount > 0 Then
            no_req = True
            While Not req_rs.EOF
                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
                req_rs.MoveNext
            Wend
            req_rs.Close
        Else
            no_req = False
            req_rs.Close
            cmb_reqno.clear
            MsgBox "No Requisition for this Modality", vbOKOnly
            btnenabl "cancel"
            set_flag "asd"
            Exit Sub
            End If 'end of rowcount
        cmb_reqno.SetFocus
      '  SendKeys "^{f4}"
    End If  'end of new_flag
If VIEW_FLAG Then
            new_flag = False
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
            If SELRREQUISITN_RS.RowCount > 0 Then
                no_req = True
                cmb_reqno.clear
                While Not SELRREQUISITN_RS.EOF
                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                    SELRREQUISITN_RS.MoveNext
                Wend
                SELRREQUISITN_RS.Close
            Else
                no_req = False
                SELRREQUISITN_RS.Close
                MsgBox "No Requisition for this Modality", vbOKOnly
                btnenabl "cancel"
                set_flag "asd"
                Exit Sub
            End If
        End If
ElseIf opt_staff.Value = True Then
    If new_flag Then
        cmb_reqno.clear
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                req_qr(2) = Trim(txt_deptcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            If req_rs.RowCount > 0 Then
                no_req = True
                While Not req_rs.EOF
                    cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
                    req_rs.MoveNext
                Wend
            Else
                cmb_reqno.clear
                no_req = False
                MsgBox "No Requisition for this Modality", vbOKOnly
                btnenabl ("INI")
                Exit Sub
            End If
    End If
    If VIEW_FLAG Then
                cmb_reqno.clear
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
            'End If
            If SELRREQUISITN_RS.RowCount > 0 Then
                no_req = True
                While Not SELRREQUISITN_RS.EOF
                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                    SELRREQUISITN_RS.MoveNext
                Wend
                SELRREQUISITN_RS.Close
             Else
                no_req = False
                MsgBox " No Requisition for this Modality"
                SELRREQUISITN_RS.Close
                btnenabl ("INI")
                Exit Sub
             End If
        End If
End If
 Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub Opt_patient_Click()
    If Opt_patient.Value = True Then
       lbl_caseno.Caption = "Case No."
       ' MaskEdBox1.Mask = "??>/#####"
    End If
    
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
   'MaskEdBox1.Mask = "#######"         ' COMMENTED BY VARINDER
End If
End Sub

Private Sub txt_impression_gotfocus()
On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    'objDoc.CheckSpelling
    'objWord.Documents.add "C:\test.doc"
    objDoc.Content = txt_impression.Text
    If Len(txt_impression.Text) > 2 Then
        objDoc.Content = txt_impression.Text
    End If
    objWord.Visible = True
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub txt_impression_lostfocus()
On Error GoTo MSG
    txt_impression.Text = objDoc.Content
    objDoc.Content = ""
    ' Modified By Pankaj Baggal 03/09/2004
    'objDoc.SaveAs "c:\DISREPORTSEC\frozen.doc"
    objDoc.Close (wdDoNotSaveChanges)
    objWord.quit
    Set objWord = Nothing
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub

Private Sub txt_report_GotFocus()
On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    objDoc.Content = txt_report.Text
    objWord.Visible = True
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub

Private Sub txt_report_LostFocus()
On Error GoTo MSG
    txt_report.Text = objDoc.Content
    objDoc.Content = ""
    ' Modified By Pankaj Baggal 03/09/2004
    'objDoc.SaveAs "c:\DISREPORTSEC\frozen.doc"
    objDoc.Close (wdDoNotSaveChanges)
    objWord.quit
    Set objWord = Nothing
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub

'common procedure for new and view mode
Private Sub getupperpanel()
'On Error GoTo MSG
        'THIS TO GET DEPTCODE AND MODALITYCODE
        sel_micro_qr(0) = mod_reqno
        Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
        While Not sel_micro_rs.EOF
            txt_microno.Text = sel_micro_rs("PATHNO")
            'STILL TO WRITE FOR PROCEDURE HEADING
            'if  sel_micro_rs("METASTASIS")=true then
            If sel_micro_rs("METASTASIS") = Null Or Trim(sel_micro_rs("METASTASIS")) = "" Then
            label = ""
            Else
            label = sel_micro_rs("METASTASIS")
            End If
            LBL_ENDOCAPTION.Caption = Trim(label) & "  REPORT"
            'show this remarks but later on store in report table
            txt_comments.Text = sel_micro_rs("PREPATHNO")
            txt_lasttreat.Text = "" & sel_micro_rs("NMETERIALS")
            'txt_sampling.Text = sel_micro_rs("OUTSSPECIM")
            sel_micro_rs.MoveNext
        Wend
        sel_prevdia_qr(0) = MaskEdBox1.FormattedText
        Set sel_prevdia_rs = sel_prevdia_qr.OpenResultset(1, 3)
        If Not sel_prevdia_rs.EOF Then
            txt_provdiag.Text = "" & sel_prevdia_rs(0)
            sel_prevdia_rs.Close
        Else
            sel_prevdia_rs.Close
        End If
        S1 = Mid(cmb_reqno.Text, 1, 3) 'TO GET DEPTCODE
        S2 = Mid(cmb_reqno.Text, 5, 2) ' TO GET MODALITY CODE
        SEL_MODALITY_QR(0) = S1
        SEL_MODALITY_QR(1) = S2
        Set SEL_MODALITY_RS = SEL_MODALITY_QR.OpenResultset(1, 3)
        While Not SEL_MODALITY_RS.EOF
            txt_modcode.Text = Trim(SEL_MODALITY_RS("MODLITYDES"))
            SEL_MODALITY_RS.MoveNext
        Wend
        txt_deptcode.Text = S1
        txt_modcode.Text = S2
        If Opt_patient.Value Then
            SEL_detail_qr(0) = MaskEdBox1.FormattedText
            Set SEL_detail_rs = SEL_detail_qr.OpenResultset(1, 3)
            If SEL_detail_rs.RowCount = 0 Then
                SEL_detail_rs.Close
                Exit Sub
            Else
                txt_sex.Text = SEL_detail_rs("sex")
                txt_name.Text = Trim(SEL_detail_rs("first_name")) & " " & Trim(SEL_detail_rs("mid_name")) & " " & Trim(SEL_detail_rs("last_name"))
                txt_age = DateDiff("YYYY", SEL_detail_rs("BIRTH_DATE"), CStr(DATE11))
                txt_pridoc.Text = SEL_detail_rs("ud_name")
                txt_category.Text = SEL_detail_rs("pat_catg")
                SEL_detail_rs.Close
            End If
        End If
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = mod_reqno
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            If REFDR_RS.RowCount > 0 Then
                 txt_refdoc.Text = REFDR_RS("REFBYDOC")
                If Not callstatus Then
                    If REFDR_RS("COMMITFLAG") = "Y" Then
                        MsgBox "Report is Commited,You cannot Modify the Final Report"
                        cmd_modify.Enabled = False
                        disable_field
                        ADDENDUMFLAG = False
                        ADDENDUM "GET"
                    End If
                    REFDR_RS.Close
                Else
                    REFDR_RS.Close
                End If
            End If
        'TO GET PROCEDURE NAME AND CODE
        sel_proccode_qr(0) = mod_reqno
            Set sel_proccode_rs = sel_proccode_qr.OpenResultset(1, 3)
            If Not sel_proccode_rs.EOF Then
                txt_proccode.Text = sel_proccode_rs(0)
                txt_procname.Text = Trim(sel_proccode_rs(1))
                sel_proccode_rs.Close
            Else
                sel_proccode_rs.Close
            End If
            sel_bed_qr(0) = MaskEdBox1.FormattedText
            Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
        If sel_bed_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
            txt_ward.Text = sel_bed_rs("WARD_NAME")
            txt_bedno.Text = sel_bed_rs("bed_no")
            txt_status.Text = "IP"
            sel_bed_rs.Close
        Else
            txt_status.Text = "OP"
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
            sel_bed_rs.Close
        End If
        Exit Sub
'MSG:
 '       MsgBox Err.Description
End Sub

Private Sub insert1()
On Error GoTo MSG
            ins_1_qr(0) = mod_reqno
            If Len(txt_report.Text) > 0 Then
                ins_1_qr(1) = Null
                ins_1_qr(1) = Trim(txt_report.Text)
            Else
                ins_1_qr(1) = Null
                ins_1_qr(1) = " "
            End If
            If Len(txt_impression.Text) > 0 Then
                ins_1_qr(2) = Null
                ins_1_qr(2) = Trim(txt_impression.Text)
            Else
                ins_1_qr(2) = Null
                ins_1_qr(2) = " "
            End If
                ins_1_qr(3) = Trim(cmb_documentation.Text)
                ins_1_qr(4) = Trim(cmb_type.Text)
            If Len(txt_complication.Text) > 0 Then
                ins_1_qr(5) = Null
                ins_1_qr(5) = Trim(txt_complication.Text)
            Else
                ins_1_qr(5) = Null
                ins_1_qr(5) = " "
            End If
            If Len(txt_instrument.Text) > 0 Then
                ins_1_qr(6) = Null
                ins_1_qr(6) = Trim(txt_instrument.Text)
            Else
                ins_1_qr(6) = Null
                ins_1_qr(6) = " "
            End If
            ins_1_qr(7) = Trim(cmb_procdone.Text)
            If Len(txt_comments.Text) > 0 Then
                ins_1_qr(8) = Null
                ins_1_qr(8) = Trim(txt_comments.Text)
            Else
                ins_1_qr(8) = Null
                ins_1_qr(8) = " "
            End If
            ins_1_qr(9) = msk_procdate.FormattedText
            If Len(txt_sampling.Text) > 0 Then
                ins_1_qr(10) = Null
                ins_1_qr(10) = Trim(txt_sampling.Text)
            Else
                ins_1_qr(10) = Null
                ins_1_qr(10) = " "
            End If
            ins_1_qr.Execute
            update_rislib
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub RETRIVE()
'On Error GoTo MSG
    'getupperpanel
'    GET_ENTER_BY
    sel_forupd1_qr(0) = mod_reqno
    Set sel_forupd1_rs = sel_forupd1_qr.OpenResultset(1, 3)
    If Not sel_forupd1_rs.EOF Then
    While Not sel_forupd1_rs.EOF
        txt_report.Text = sel_forupd1_rs("report")
        txt_impression.Text = sel_forupd1_rs("IMPRESSION")
        cmb_documentation.Text = sel_forupd1_rs("DOCMENT")
        txt_complication.Text = sel_forupd1_rs("COMPCATION")
        txt_instrument.Text = sel_forupd1_rs("INSTRUMENT")
        cmb_type.Text = sel_forupd1_rs("DTTYPE")
        cmb_procdone.Text = sel_forupd1_rs("pperformed")
        txt_comments.Text = sel_forupd1_rs("remarks")
        'txt_sampling.Text = sel_forupd1_rs("REPORT")
        txt_sampling.Text = sel_forupd1_rs("TISSAMP")
        msk_procdate.Text = "" & sel_forupd1_rs("fld1")
        sel_forupd1_rs.MoveNext
    Wend
        'ADDENDUM "GET"
        sel_forupd1_rs.Close
    Else
        sel_forupd1_rs.Close
    End If
    GET_ENTER_BY
    Exit Sub
'MSG:
'    MsgBox Err.Description
End Sub

Private Sub disable_field()
On Error GoTo MSG
    txt_procname.Enabled = False
    txt_report.Enabled = False
    txt_impression.Enabled = False
    txt_addendum.Enabled = False
    txt_addendum.Enabled = False
    txt_addendum.Visible = False
    lbl_addendum.Visible = False
    cmb_documentation.Enabled = False
    txt_proccode.Enabled = False
    txt_sampling.Enabled = False
    cmb_type.Enabled = False
    txt_complication.Enabled = False
    txt_instrument.Enabled = False
    cmb_procdone.Enabled = False
    msk_procdate.Enabled = False
    txt_comments.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_reportby.Enabled = False
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub enable_field()
On Error GoTo MSG
    txt_procname.Enabled = True
    txt_report.Enabled = True
    txt_impression.Enabled = True
    cmb_documentation.Enabled = True
    txt_proccode.Enabled = True
    txt_sampling.Enabled = True
    cmb_type.Enabled = True
    txt_complication.Enabled = True
    txt_instrument.Enabled = True
    cmb_procdone.Enabled = True
    msk_procdate.Enabled = True
    cmb_enterby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_reportby.Enabled = True
    txt_comments.Enabled = True
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub STAFF()
On Error GoTo MSG
If new_flag Then
        SEL_dreq_staff_qr(0) = mod_reqno
        SEL_dreq_staff_qr(1) = MaskEdBox1.FormattedText
        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
        If SEL_dreq_staff_rs.RowCount > 0 Then
            txt_name = SEL_dreq_staff_rs("PNAME")
            txt_sex = SEL_dreq_staff_rs("SEX")
            txt_age = SEL_dreq_staff_rs("AGE")
            If SEL_dreq_staff_rs("RELATION") <> "SELF" Then
                txt_status = "DP"
            Else
                txt_status = "ST"
            End If
        End If
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = REQNO
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
    sel_bed_qr(0) = MaskEdBox1.FormattedText
    Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
        If sel_bed_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
            txt_ward.Text = sel_bed_rs("WARD_NAME")
            txt_bedno.Text = sel_bed_rs("bed_no")
            sel_bed_rs.Close
        Else
            txt_status.Text = "OP"
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
            sel_bed_rs.Close
        End If
        'insert1
 ElseIf VIEW_FLAG Then
        SEL_dreq_staff_qr(0) = mod_reqno
        SEL_dreq_staff_qr(1) = MaskEdBox1.Text
        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
        If SEL_dreq_staff_rs.RowCount > 0 Then
            txt_name = SEL_dreq_staff_rs("PNAME")
            txt_sex = SEL_dreq_staff_rs("SEX")
            txt_age = SEL_dreq_staff_rs("AGE")
            If SEL_dreq_staff_rs("RELATION") <> "SELF" Then
                txt_status = "DP"
            Else
                txt_status = "ST"
            End If
        End If
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = mod_reqno
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
    sel_bed_qr(0) = MaskEdBox1.FormattedText
    Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
'        If sel_bed_rs.RowCount > 0 Then
'            txt_ward.Visible = True
'            txt_bedno.Visible = True
'            lbl_ward.Visible = True
'            lbl_bed.Visible = True
'            txt_ward.Text = sel_bed_rs("WARD_NAME")
'            txt_bedno.Text = sel_bed_rs("bed_no")
'            sel_bed_rs.Close
'        Else
'            txt_status.Text = "OP"
'            txt_ward.Visible = False
'            txt_bedno.Visible = False
'            lbl_ward.Visible = False
'            lbl_bed.Visible = False
'            sel_bed_rs.Close
'        End If
        getupperpanel
        RETRIVE
 End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub MODIFY()

'STR = "UPDATE DISLIB.DEDREPORT SET REPORT=?,IMPRESSION=?,DOCMENT=?,DTTYPE=?,COMPcATION=?,INSTRUMENT=?,PPERFORMED=?,REMARKS=?,FLD1=?,TISSAMP=? WHERE REQNO=?"

On Error GoTo MSG
    If Len(txt_report.Text) > 0 Then
        upd_7_qr(0) = Null
        upd_7_qr(0) = Trim(txt_report.Text)
    Else
        upd_7_qr(0) = Null
        upd_7_qr(0) = " "
    End If
    If Len(txt_impression.Text) > 0 Then
        upd_7_qr(1) = Null
        upd_7_qr(1) = Trim(txt_impression.Text)
    Else
        upd_7_qr(1) = Null
        upd_7_qr(1) = " "
    End If
        upd_7_qr(2) = Trim(cmb_documentation.Text)
        upd_7_qr(3) = Trim(cmb_type.Text)
    If Len(txt_complication.Text) > 0 Then
        upd_7_qr(4) = Null
        upd_7_qr(4) = Trim(txt_complication.Text)
    Else
        upd_7_qr(4) = Null
        upd_7_qr(4) = " "
    End If
    If Len(txt_instrument.Text) > 0 Then
        upd_7_qr(5) = Null
        upd_7_qr(5) = Trim(txt_instrument.Text)
    Else
        upd_7_qr(5) = Null
        upd_7_qr(5) = " "
    End If
        
        upd_7_qr(6) = Trim(cmb_procdone.Text)
    
    If Len(txt_comments.Text) > 0 Then
        upd_7_qr(7) = Null
        upd_7_qr(7) = Trim(txt_comments.Text)
    Else
        upd_7_qr(7) = Null
        upd_7_qr(7) = " "
    End If
        upd_7_qr(8) = msk_procdate.FormattedText
        
    If Len(txt_sampling.Text) > 0 Then
        upd_7_qr(9) = Null
        upd_7_qr(9) = Trim(txt_sampling.Text)
    Else
        upd_7_qr(9) = Null
        upd_7_qr(9) = " "
    End If
        upd_7_qr(10) = mod_reqno
        upd_7_qr.Execute
        update_rislib
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Private Sub ADDENDUM(work As String)
'addendum are not updatable
On Error GoTo MSG
Dim K As String
If work = "INSERT" Then
    'check for tab and replace
    While InStr(txt_addendum, Chr(9)) > 0
         K = InStr(txt_addendum, Chr(9))
         S1 = Mid(txt_addendum, 1, K - 1) & " "
         S2 = Mid(txt_addendum, K + 1)
         txt_addendum = S1 & S2
    Wend
    While InStr(txt_addendum, Chr(146)) > 0
         K = InStr(txt_addendum, Chr(146))
         S1 = Mid(txt_addendum, 1, K - 1) & "`"
         S2 = Mid(txt_addendum, K + 1)
         txt_addendum = S1 & S2
    Wend
    'after commit it is allowed to add addendum
    'check for commit and enable addendum flag and check here
 If ADDENDUMFLAG Then
    sel_addendum_qr(0) = mod_reqno
    Set sel_addendum_rs = sel_addendum_qr.OpenResultset(1, 3) 'used to check whether first time or next time
    If sel_addendum_rs.RowCount > 0 Then
            INS_3_qr(0) = mod_reqno
            INS_3_qr(1) = sel_addendum_rs("ADDEND_NO") + 1
            INS_3_qr(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INS_3_qr(3) = time1
            'COMMENTED FOR TIMEBEIGN
            'INS_3_qr(4) = cmb_addendumby.Text
            INS_3_qr(4) = " "
            INS_3_qr(5) = " "
            INS_3_qr(5) = txt_addendum.Text
            INS_3_qr.Execute
        Else
            INS_3_qr(0) = mod_reqno
            INS_3_qr(1) = "1"
            INS_3_qr(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INS_3_qr(3) = time1
            'COMMENTED FOR TIMEBEIGN
            'INS_3_qr(4) = cmb_addendumby.Text
            INS_3_qr(4) = " "
            INS_3_qr(5) = " "
            INS_3_qr(5) = txt_addendum.Text
            INS_3_qr.Execute
        End If
        MsgBox "Addendum Added Successfully", vbOKOnly
        txt_addendum.Enabled = False
        btnenabl "INI"
        
        Exit Sub
 End If
ElseIf work = "GET" Then
    sel_addendum_qr(0) = mod_reqno
    Set sel_addendum_rs = sel_addendum_qr.OpenResultset(1, 3)
    S1 = ""
    If sel_addendum_rs.RowCount > 0 Then
        While Not sel_addendum_rs.EOF
            S1 = " ADDENDUM NO " & sel_addendum_rs("ADDEND_NO") & Chr(13) & sel_addendum_rs("ADD_MAX_LN") & Chr(13) & sel_addendum_rs("ADDEND_DT") & " " & sel_addendum_rs("ADDEND_TM") & " " & sel_addendum_rs("USER_ID")
        '    cmb_addendumby.AddItem SEL_ADDENDUM_RS("USER_ID")
        '    cmb_addendumby.Text = SEL_ADDENDUM_RS("USER_ID")
        sel_addendum_rs.MoveNext
        Wend
        txt_addendum.Text = Trim(S1)
    Else
        txt_addendum.Text = ""
    End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub


Private Sub GET_ENTER_BY()
On Error GoTo MSG
Dim sel_ris_qr  As rdo.rdoQuery
Dim sel_ris_rs  As rdo.rdoResultset

str = "select * from rislib.rrequisitn where reqnno=?"
Set sel_ris_qr = rdb.CreateQuery("echo_get_result", str)
    sel_ris_qr(0) = mod_reqno
    Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
    If Not sel_ris_rs.EOF Then
        While Not sel_ris_rs.EOF
            cmb_enterby.Text = sel_ris_rs("ENTERBY") & ""
            cmb_reportby.Text = sel_ris_rs("preportby") & ""
            cmb_finalby.Text = sel_ris_rs("freportby") & ""
            cmb_reportby.Enabled = False
            cmb_enterby.Enabled = False
            cmb_finalby.Enabled = False
            sel_ris_rs.MoveNext
        Wend
        sel_ris_rs.Close
        sel_ris_qr.Close
    Else
        sel_ris_rs.Close
        sel_ris_qr.Close
    End If
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub



Public Sub printendo(moda As String, mskval As String, reqval As String)
    Me.Hide
    VIEW_FLAG = True
    new_flag = False
    callstatus = True
    If IsNumeric(mskval) Then
        opt_staff.Value = True
    Else
        Opt_patient.Value = True
    End If
    txt_deptcode.Text = "CBZ"
    txt_modcode.Text = moda
    MaskEdBox1.Text = mskval
    cmb_modality_Click
    cmb_reqno.AddItem req_chglen(reqval)
    cmb_reqno.Text = req_chglen(reqval)
    cmb_reqno_Click
    cmd_print_Click
End Sub


