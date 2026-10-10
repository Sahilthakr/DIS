VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{9B8DFCFF-6726-11D6-99D4-000244193CE4}#1.0#0"; "Project1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form sppathreport 
   Caption         =   "Report - Surgical Pathology"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystalreport1 
      Left            =   960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.TextBox txt_size 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6000
      TabIndex        =   34
      Top             =   6240
      Width           =   735
   End
   Begin VB.Frame common 
      Height          =   3015
      Left            =   0
      TabIndex        =   45
      Top             =   960
      Width           =   11775
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   0
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   0
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   18
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   17
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7920
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7920
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5520
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10200
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   16711680
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10200
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   480
         TabIndex        =   87
         Top             =   2520
         Width           =   1830
      End
      Begin VB.Label lbl_reportdoc 
         AutoSize        =   -1  'True
         Caption         =   "Name of the Reporting Doctor"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6360
         TabIndex        =   66
         Top             =   2640
         Width           =   2565
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9120
         TabIndex        =   64
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   63
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6960
         TabIndex        =   62
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   61
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9720
         TabIndex        =   59
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6960
         TabIndex        =   58
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5040
         TabIndex        =   57
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1920
         TabIndex        =   56
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1680
         TabIndex        =   55
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Path No."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9360
         TabIndex        =   54
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4920
         TabIndex        =   53
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1440
         TabIndex        =   52
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9120
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   50
         Top             =   360
         Width           =   1830
      End
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame COMMAND 
      Height          =   855
      Left            =   0
      TabIndex        =   47
      Top             =   7680
      Width           =   11775
      Begin VB.CommandButton cmd_fullreport 
         Caption         =   "Full&Report"
         Height          =   375
         Left            =   10080
         MouseIcon       =   "sppathreport.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   105
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   6360
         MouseIcon       =   "sppathreport.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   86
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   8880
         MouseIcon       =   "sppathreport.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   85
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   7680
         MouseIcon       =   "sppathreport.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   84
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5160
         MouseIcon       =   "sppathreport.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3960
         MouseIcon       =   "sppathreport.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   82
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "sppathreport.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   81
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1560
         MouseIcon       =   "sppathreport.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   80
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         MouseIcon       =   "sppathreport.frx":1850
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   615
      Left            =   3000
      TabIndex        =   44
      Top             =   -120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
   End
   Begin VB.Frame REPORT 
      Height          =   3855
      Left            =   0
      TabIndex        =   46
      Top             =   3840
      Width           =   11775
      Begin VB.TextBox txt_fullreport 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   106
         Top             =   3030
         Width           =   1095
      End
      Begin VB.ComboBox CMB_PATHNO 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8760
         TabIndex        =   104
         Top             =   2760
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cmb_t 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1B5A
         Left            =   4800
         List            =   "sppathreport.frx":1B5C
         TabIndex        =   33
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox cmb_m 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1B5E
         Left            =   8280
         List            =   "sppathreport.frx":1B60
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox CMB_N 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1B62
         Left            =   7320
         List            =   "sppathreport.frx":1B64
         TabIndex        =   35
         Top             =   2400
         Width           =   570
      End
      Begin VB.ComboBox CMB_ICDH 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1B66
         Left            =   2520
         List            =   "sppathreport.frx":1B68
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox cmb_grossby 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1B6A
         Left            =   8880
         List            =   "sppathreport.frx":1B6C
         TabIndex        =   25
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox cmb_fsd 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1B6E
         Left            =   10200
         List            =   "sppathreport.frx":1B7B
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox cmb_finalby 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   3120
         Width           =   2775
      End
      Begin VB.ComboBox cmb_provby 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3480
         Width           =   2295
      End
      Begin VB.ComboBox cmb_enteryby 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3120
         Width           =   2295
      End
      Begin VB.ComboBox cmb_addendumby 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txt_laterality 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5160
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_sites 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   38
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cmb_immhische 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1B91
         Left            =   2520
         List            =   "sppathreport.frx":1B9B
         TabIndex        =   27
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txt_immhischerep 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   30
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txt_impfindiag 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6360
         TabIndex        =   31
         Top             =   2040
         Width           =   5175
      End
      Begin VB.TextBox txt_microdesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   26
         Top             =   1320
         Width           =   9015
      End
      Begin VB.TextBox txt_immhischereqno 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7800
         TabIndex        =   28
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txt_frosecrep 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   600
         Width           =   9015
      End
      Begin VB.TextBox txt_prepathreqno 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8760
         TabIndex        =   39
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txt_grossdesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   24
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txt_opprocedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8760
         TabIndex        =   22
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmb_labno 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "sppathreport.frx":1BA8
         Left            =   10800
         List            =   "sppathreport.frx":1BAA
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lbl_labno 
         AutoSize        =   -1  'True
         Caption         =   "LabNo"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   10200
         TabIndex        =   103
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Pathstg"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3960
         TabIndex        =   102
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5640
         TabIndex        =   99
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "M"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8040
         TabIndex        =   98
         Top             =   2520
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7080
         TabIndex        =   97
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4680
         TabIndex        =   96
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ICD-H"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1920
         TabIndex        =   95
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grossed By"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7800
         TabIndex        =   94
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl_fsd 
         AutoSize        =   -1  'True
         Caption         =   "F.S.Disprcy"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9240
         TabIndex        =   93
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label lbl_addendum1 
         AutoSize        =   -1  'True
         Caption         =   "Addendum "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1560
         TabIndex        =   92
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label lbl_addendum 
         AutoSize        =   -1  'True
         Caption         =   "Addendum By"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   7560
         TabIndex        =   91
         Top             =   3480
         Width           =   1170
      End
      Begin VB.Label lbl_final 
         AutoSize        =   -1  'True
         Caption         =   "Finalized By"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7560
         TabIndex        =   90
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label lbl_pro 
         AutoSize        =   -1  'True
         Caption         =   "Provisional By"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1230
         TabIndex        =   89
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1725
         TabIndex        =   88
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lbl_immhischereq 
         AutoSize        =   -1  'True
         Caption         =   "Immhistochemistry Req. No."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   77
         Top             =   1680
         Width           =   2370
      End
      Begin VB.Label lbl_Immhische 
         AutoSize        =   -1  'True
         Caption         =   "Immunohistochemistry"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   585
         TabIndex        =   76
         Top             =   1680
         Width           =   1860
      End
      Begin VB.Label Lbl_immhischerep 
         AutoSize        =   -1  'True
         Caption         =   "Immunohistochemistry Report"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -15
         TabIndex        =   75
         Top             =   2040
         Width           =   2490
      End
      Begin VB.Label lbl_impression 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   74
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label lbl_microdesc 
         AutoSize        =   -1  'True
         Caption         =   "Microscopic Description"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   73
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lbl_frosec 
         AutoSize        =   -1  'True
         Caption         =   "Frozen Secion Report"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   585
         TabIndex        =   72
         Top             =   600
         Width           =   1860
      End
      Begin VB.Label lbl_prepathreqno 
         AutoSize        =   -1  'True
         Caption         =   "Previous Path No./ Requistion No."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5745
         TabIndex        =   71
         Top             =   2880
         Width           =   2970
      End
      Begin VB.Label lbl_Grossdesc 
         AutoSize        =   -1  'True
         Caption         =   "Gross Description"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         TabIndex        =   70
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lbl_opprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Operative Procedure/s"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6735
         TabIndex        =   69
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label lbl_laterality 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4320
         TabIndex        =   68
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lbl_sites 
         AutoSize        =   -1  'True
         Caption         =   "Site/s"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1920
         TabIndex        =   67
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Size"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   101
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "T"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4200
      TabIndex        =   100
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label lbl_cap 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - SURGICAL PATHOLOGY"
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
      Height          =   330
      Left            =   3480
      TabIndex        =   78
      Top             =   480
      Width           =   4845
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9840
      TabIndex        =   49
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9840
      TabIndex        =   48
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "sppathreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel_deptqr As rdoQuery
Dim UPDSPREPORT_QR As rdoQuery
Dim sel_deptrs As rdoResultset
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim sel_casno As rdoQuery
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim selta101_qr As rdoQuery
Dim selta101_rs As rdoResultset
Dim repdr_qr As rdoQuery
Dim repdr_rs As rdoResultset
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset
Dim SELPATHDREQUISITN_QR As rdoQuery
Dim SELPATHDREQUISITN_RS As rdoResultset
Dim selma302_qr As rdoQuery
Dim selma302_rs As rdoResultset
Dim seldfsreport_qr As rdoQuery
Dim seldfsreport_rs As rdoResultset
Dim seldihreport_qr As rdoQuery
Dim seldihreport_rs As rdoResultset
Dim seldimpresion_qr As rdoQuery
Dim seldimpresion_rs As rdoResultset
Dim seldgrossdesc_qr As rdoQuery
Dim seldgrossdesc_rs As rdoResultset
Dim seldspreport_qr As rdoQuery
Dim seldspreport_rs As rdoResultset
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim w_frozen As Application, W_grossdes As Application, w_micscop As Application, w_imp As Application, w_immun As Application, w_adden As Application, w As Application
Dim d_frozen As Document, d_grossdesc As Document, d_micscop As Document, d_imp As Document, d_immun As Document, d_adden As Document, d As Document, d1 As Document
Dim SELEMPLIB_QR As rdoQuery
Dim SELEMPLIB_RS As rdoResultset
Dim I As Integer
Dim insdfsrep_qr As rdoQuery
Dim insdgrossdesc_qr As rdoQuery
Dim insdfs_qr As rdoQuery
Dim insdimpresion_qr As rdoQuery
Dim insdspreport_qr As rdoQuery
Dim insdihreport_qr As rdoQuery
Dim upddfsreport_qr As rdoQuery
Dim upddihreport_qr As rdoQuery
Dim upddimpresion_qr As rdoQuery
Dim upddspreport_qr As rdoQuery
'Dim upddimpresion_qr As rdoQuery
Dim updgrossdesc_qr As rdoQuery
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim selempbasinfo1_qr As rdoQuery
Dim selempbasinfo1_rs As rdoResultset
Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim INSDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim newflag As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer
Dim ADDENDUMFALG As Boolean


 

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
    sel_moadlityqr(0) = Trim(txt_deptcode.Text)
    sel_modalityrs.MoreResults
    sel_modalityrs.Requery
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.Clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(ind_cbodept)) = "PATHOLOGY$FZZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_immhische_Click()
If cmb_immhische.Text = "NO" Then
    txt_immhischerep.Enabled = False
End If
If cmb_immhische.Text = "YES" Then
    txt_immhischerep.Visible = True
    pathology.Show
    pathology.cmb_modality.Text = "IH"
    pathology.immunohistochemistry.Visible = True
    pathology.surgicalPathology.Visible = False
    pathology.lbl_req.Left = 3850
    pathology.lbl_req.Caption = "IMMUNOHISTOCHEMISTRY"
    pathology.Caption = "Pathology - Immunohistochemistry"
    IHFLAG = True
End If

End Sub

Private Sub cmb_modality_Click()
Dim modality As String
Dim moda As Integer
Dim l As Integer
On Error GoTo MSG:
If Trim(cmb_modality) = "" Then
    Exit Sub
End If
moda = InStr(cmb_modality, "$")
l = Len(cmb_modality.Text)
txt_modcode.Text = Right(cmb_modality.Text, (l - moda))
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmd_back_Click()
Unload Me
DIS_MAIN.Show
End Sub
Public Sub btnenabl(cap As String)

If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    cmd_fullreport.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    cmd_fullreport.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = True
    cmd_view.Enabled = True
    cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
    cmd_fullreport.Enabled = False
ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    CMD_NEW.Enabled = True
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
    cmd_fullreport.Enabled = False
ElseIf cap = "print" Then
     cmd_save.Enabled = False
     cmd_back.Enabled = True
     cmd_modify.Enabled = False
     cmd_print.Enabled = False
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False

End If
End Sub

Private Sub cmd_cancel_Click()
btnenabl ("INI")
field_clear
TXT_CLEAR
FLDENABL ("DIS")
End Sub

Private Sub cmd_commit_Click()
On Error GoTo MSG:

 rdb.BeginTrans
    updrrequisitn1_qr(0) = Left(cmb_finalby.Text, 30)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
rdb.CommitTrans
btnenabl ("INI")
MsgBox "Record is Committed Successfully."
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_fullreport_Click()
txt_fullreport_GotFocus
Set d1 = w.ActiveDocument
If Len(txt_frosecrep) > 0 Then
d1.Content = d1.Content & "FROZEN:" & Chr(13) & Trim(txt_frosecrep)
End If
If Len(txt_grossdesc.Text) > 0 Then
d1.Content = d1.Content & "GROSS DESC: " & Chr(13) & Trim(txt_grossdesc)
End If
If Len(txt_immhischerep) > 0 Then
d1.Content = d1.Content & "IMMUNOHISTO:" & Chr(13) & Trim(txt_immhischerep)
End If
If Len(txt_microdesc) > 0 Then
d1.Content = d1.Content & "MICROSCOPICDESC:" & Chr(13) & Trim(txt_microdesc)
End If
If Len(txt_impfindiag) > 0 Then
d1.Content = d1.Content & "IMPRESSION:" & Chr(13) & Trim(txt_impfindiag)
End If
If Len(txt_addendum) > 0 Then
d1.Content = d1.Content & "ADDENDUM:" & Chr(13) & Trim(txt_addendum)
End If
End Sub

Private Sub cmd_fullreport_LostFocus()
On Error GoTo MSG:

d1.Content = ""
d1.SaveAs "C:\REPORT.DOC"
d1.Close
w.Quit
'Msk_caseno.Enabled = False
'cmd_cancel.SetFocus
MSG:
End Sub

Private Sub cmd_modify_Click()
btnenabl ("NEW")
FLDENABL ("MODIFY")
flagsave = 2
If Len(txt_frosecrep.Text) > 0 Then
txt_frosecrep.Enabled = False
End If
If ADDENDUMFALG = True Then
    txt_addendum.Enabled = True
    cmb_addendumby.Enabled = True
    FLDENABL ("ADDENDUM")
    txt_addendum.SetFocus
    flagsave = 0
End If
End Sub

Private Sub cmd_new_Click()
btnenabl ("NEW")
flagsave = 1
viewflag = 0
newflag = True
FLDENABL ("ENA")
cmb_dept.SetFocus
End Sub

Private Sub cmd_print_Click()


On Error GoTo MSG:

Dim STR1 As String
Dim STR2 As String
Dim STR3 As String
 On Error GoTo prtmsg:
'txt_msg.Text = "PRINTING..."
DoEvents

Dim VAR_REQNNO As String


VAR_REQNNO = req_chglen(cmb_reqno.Text)

CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"   'AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.ReportFileName = "C:\DISREPORT\DISSP.RPT"
CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
CrystalReport1.Action = 1
'crp_rep.Formulas(0) = "MODLITYDES= '" & Trim(cmb_modality) & "'"
'crp_rep.Formulas(1) = "DEPTNAME= '" & "Department of surgical pathology" & "'"

'    STR1 = ip
'    STR2 = date11
'    STR3 = time1
'    STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
'    STR1 = USR_NAME & Space(3) & STR1
'    crp_rep.Formulas(3) = "us_id = '" & STR1 & "'"
'    crp_rep.Formulas(2) = "INOUT = '" & txt_status & "'"

'
'    crp_rep.Formulas(4) = "ward_bed = '" & txt_ward.Text & "'"
'    crp_rep.Formulas(5) = "req_placed = '" & cmb_dept & "'"
'    crp_rep.Formulas(6) = "name='" & txt_name.Text & "'"
'    crp_rep.Formulas(7) = "sex='" & txt_sex.Text & "'"
'    crp_rep.Formulas(8) = "age='" & txt_age.Text & "'"
'    crp_rep.Formulas(9) = "primarydrser='" & txt_pridoc & " '"
'    crp_rep.Formulas(10) = "refdr='" & txt_refdoc & " ' "
'    crp_rep.Formulas(11) = "provdiagby='" & txt_procedure & " ' "
    
    
    

'If addendum_flag = True Then
'cr_rep.ReportFileName = "C:\RISREPORTS\crpt_rep2.rpt"
'Else
'cr_rep.ReportFileName = "C:\RISREPORTS\crpt_rep1.rpt"
'End If

'crp_rep.ReportFileName = "C:\SPREPORT\REQ_REP.RPT"
'crp_rep.WindowState = crptMaximized

txt_msg.Text = ""
Exit Sub
prtmsg: MsgBox Error(Err)
txt_msg.Text = ""
btnenabl ("print")
Exit Sub
MSG:
MsgBox Error(Err)




End Sub

Private Sub cmd_save_Click()

On Error GoTo MSG:


Dim m As Integer
If Len(Trim(cmb_reqno)) = 0 Then
    m = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
    cmb_reqno.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_enteryby)) = 0 Then
    m = MsgBox("The Enter by field is EMPTY.", vbCritical)
    cmb_enteryby.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_finalby)) = 0 Then
    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
    cmb_finalby.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_provby)) = 0 Then
    m = MsgBox("The Provisional By field is EMPTY.", vbCritical)
    cmb_provby.SetFocus
    Exit Sub
End If

  If Len(txt_frosecrep) > 0 Then
    I = MsgBox(" ONCE YOU SAVED FROZEN SECTION YOU CAN'T MODIFY LATER", vbYesNo)
    If I = vbNo Then
        txt_frosecrep.SetFocus
        Exit Sub
    End If
  End If
        While InStr(txt_frosecrep, Chr(9)) > 0
            Dim beforst As String, afterst As String
            Dim K As Long
            K = InStr(txt_frosecrep, Chr(9))
            beforst = Mid(txt_frosecrep, 1, K - 1) & "      "
            afterst = Mid(txt_frosecrep, K + 1)
            txt_frosecrep = beforst & afterst
       Wend
        While InStr(txt_frosecrep, Chr(146)) > 0

            K = InStr(txt_frosecrep, Chr(146))
            beforst = Mid(txt_frosecrep, 1, K - 1) & "`"
            afterst = Mid(txt_frosecrep, K + 1)
            txt_frosecrep = beforst & afterst
       Wend

        While InStr(txt_grossdesc, Chr(9)) > 0

            K = InStr(txt_grossdesc, Chr(9))
            beforst = Mid(txt_grossdesc, 1, K - 1) & "      "
            afterst = Mid(txt_grossdesc, K + 1)
            txt_grossdesc = beforst & afterst
       Wend
        While InStr(txt_grossdesc, Chr(146)) > 0

            K = InStr(txt_grossdesc, Chr(146))
            beforst = Mid(txt_grossdesc, 1, K - 1) & "`"
            afterst = Mid(txt_grossdesc, K + 1)
            txt_grossdesc = beforst & afterst
       Wend

        While InStr(txt_microdesc, Chr(9)) > 0

            K = InStr(txt_microdesc, Chr(9))
            beforst = Mid(txt_microdesc, 1, K - 1) & "      "
            afterst = Mid(txt_microdesc, K + 1)
            txt_grossdesc = beforst & afterst
       Wend
        While InStr(txt_microdesc, Chr(146)) > 0

            K = InStr(txt_microdesc, Chr(146))
            beforst = Mid(txt_microdesc, 1, K - 1) & "`"
            afterst = Mid(txt_microdesc, K + 1)
            txt_microdesc = beforst & afterst
       Wend
       
       While InStr(txt_immhischerep, Chr(9)) > 0

            K = InStr(txt_immhischerep, Chr(9))
            beforst = Mid(txt_immhischerep, 1, K - 1) & "      "
            afterst = Mid(txt_immhischerep, K + 1)
            txt_immhischerep = beforst & afterst
       Wend
      While InStr(txt_immhischerep, Chr(146)) > 0

            K = InStr(txt_immhischerep, Chr(146))
            beforst = Mid(txt_immhischerep, 1, K - 1) & "`"
            afterst = Mid(txt_immhischerep, K + 1)
            txt_immhischerep = beforst & afterst
       Wend
       
       
        While InStr(txt_addendum, Chr(9)) > 0

            K = InStr(txt_addendum, Chr(9))
            beforst = Mid(txt_addendum, 1, K - 1) & "      "
            afterst = Mid(txt_addendum, K + 1)
            txt_addendum = beforst & afterst
       Wend
      While InStr(txt_addendum, Chr(146)) > 0

            K = InStr(txt_addendum, Chr(146))
            beforst = Mid(txt_addendum, 1, K - 1) & "`"
            afterst = Mid(txt_addendum, K + 1)
            txt_addendum = beforst & afterst
       
       Wend
rdb.BeginTrans
       assign
rdb.CommitTrans
       MsgBox "Record is saved Successfully."
  FLDENABL ("DIS")
  btnenabl ("SAVE")
  Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Error(Err)

End Sub

Private Sub cmd_view_Click()
btnenabl ("VIEW")
FLDENABL ("VIEW")
cmb_dept.Enabled = True
viewflag = True
cmb_dept.SetFocus
cmd_print.Enabled = True
End Sub

Private Sub Form_Load()
btnenabl ("INI")
newflag = False
viewflag = False


datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
For I = 0 To 4
    CMB_N.AddItem I
    cmb_m.AddItem I
Next I
CMB_N.AddItem "X"
cmb_m.AddItem "X"
Dim STR As String
'for dept selection
STR = "select dept_name, dept_code from PATOOLSLIB.ma106"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr11", STR)
Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
   If sel_deptrs.RowCount > 0 Then
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
        Wend
   End If
    'for icd-h retriveing
STR = "select * from patoolslib.ma302"
Set selma302_qr = rdb.CreateQuery("selma302_qr1", STR)
Set selma302_rs = selma302_qr.OpenResultset(2, 3)
        If selma302_rs.RowCount > 0 Then
            selma302_rs.MoveFirst
            While Not selma302_rs.EOF
                CMB_ICDH.AddItem Trim(selma302_rs("histdesc"))
                selma302_rs.MoveNext
            Wend
        End If
' for modality selection
STR = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr11", STR)
sel_moadlityqr(0) = "FZZ"
Set sel_modalityrs = sel_moadlityqr.OpenResultset(2, 3)
'sel_modalityrs.MoveFirst
'While Not sel_modalityrs.EOF
'    cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
'    sel_modalityrs.MoveNext
'    Wend
    'for retriving reqnos
Set req_qr = rdb.CreateQuery("req", "select * from dislib.RREQUISITN where caseno=? AND REQNSTAT='E' ORDER BY REQNDT DESC")
req_qr(0) = "aaa"
Set req_rs = req_qr.OpenResultset(1, 3)

' FOR ENTER BY, FINALIZED BY
Set SELEMPLIB_QR = rdb.CreateQuery("SELEMPLIB_QR1", "SELECT * FROM EMPLIB.EMPBASINFO WHERE TITLE='DR.'")
'SELEMPLIB_QR(0) = "AA"
Set SELEMPLIB_RS = SELEMPLIB_QR.OpenResultset(2, 3)
SELEMPLIB_RS.MoveFirst
While Not SELEMPLIB_RS.EOF
cmb_enteryby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
cmb_finalby.AddItem SELEMPLIB_RS("EmpTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
cmb_provby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
cmb_addendumby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
cmb_grossby.AddItem SELEMPLIB_RS("emptitle") & "$" & SELEMPLIB_RS("empcode")
SELEMPLIB_RS.MoveNext
Wend


' for retriveing name,age etc

Set selta101_qr = rdb.CreateQuery("selta101_qr111", "select * from PATOOLSLIB.ta101 where case_no =?")
selta101_qr(0) = "aa"
Set selta101_rs = selta101_qr.OpenResultset(2, 3)
'for view purpose
STR = "select * from emplib.empbasinfo where empcode =?"
Set selempbasinfo1_qr = rdb.CreateQuery("selempbasinfo1_qr1", STR)
selempbasinfo1_qr(0) = "00001"
Set selempbasinfo1_rs = selempbasinfo1_qr.OpenResultset(2, 3)


'' for name of  the reporting doctor
STR = "select emptitle from emplib.empbasinfo"
Set repdr_qr = rdb.CreateQuery("repdr_qr10", STR)
Set repdr_rs = repdr_qr.OpenResultset(2, 3)
repdr_rs.MoveFirst
    While Not repdr_rs.EOF
        cmb_reportdoc.AddItem repdr_rs("emptitle")
        repdr_rs.MoveNext
    Wend
    ' FOR REFEFING DOCTOR
    STR = "SELECT * FROM DISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1", STR)
    REFDR_QR(0) = "BN/00001"
    REFDR_QR(1) = "FZZSP010000"
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    ' FOR PATH NO
    STR = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1", STR)
    SELPATHDREQUISITN_QR(0) = "SSSSS"
    'SELPATHDREQUISITN_QR(1) = "FZZSP010000"
    Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(2, 3)

    ' for bedno, ward retriveing
    STR = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11", STR)
    selma104_qr(0) = "BN/00001"
    Set selma104_rs = selma104_qr.OpenResultset(2, 3)
    'for provisional diagnosis

    STR = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr11", STR)
    selta1012_qr(0) = "BN/00001"
    Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)

    STR = "select * from dislib.dfsreport where reqno=?"
    Set seldfsreport_qr = rdb.CreateQuery("seldfsreport_qr1", STR)
    seldfsreport_qr(0) = "zzzzzz"
    Set seldfsreport_rs = seldfsreport_qr.OpenResultset(2, 3)

    STR = "select * from dislib.dihreport where reqno=?"
    Set seldihreport_qr = rdb.CreateQuery("seldihreport_qr", STR)
    seldihreport_qr(0) = "zzzzzz"
    Set seldihreport_rs = seldihreport_qr.OpenResultset(2, 3)

    STR = "select * from dislib.dimpresion where reqno=?"
    Set seldimpresion_qr = rdb.CreateQuery("seldimpresion_qr11", STR)
    seldimpresion_qr(0) = "zzzzzz"
    Set seldimpresion_rs = seldimpresion_qr.OpenResultset(2, 3)

    STR = "select * from dislib.dgrossdesc where reqno=?"
    Set seldgrossdesc_qr = rdb.CreateQuery("seldgrossdesc_qr1", STR)
    seldgrossdesc_qr(0) = "zzzzzz"
    Set seldgrossdesc_rs = seldgrossdesc_qr.OpenResultset(2, 3)


    STR = "select * from dislib.dspreport wherE reqno=?"
    Set seldspreport_qr = rdb.CreateQuery("seldspreport_qr11", STR)
    seldspreport_qr(0) = "zzzzzz"
    Set seldspreport_rs = seldspreport_qr.OpenResultset(2, 3)



    ' for insertion
    STR = "insert into  dislib.dfsreport (reqno, fsreport,enterby,preportby,freportby) values (?,?,?,?,?)"
    Set insdfs_qr = rdb.CreateQuery("insdfs_qr1", STR)
    STR = "insert into dislib.dgrossdesc (reqno,grossdesc,grossedby) values (?,?,?)"
    Set insdgrossdesc_qr = rdb.CreateQuery("insdgrossdesc_qr1", STR)
    STR = "insert into dislib.dihreport (reqno,ihreport,enteryby,preportby,freportby) values (?,?,?,?,?)"
    Set insdihreport_qr = rdb.CreateQuery("insdihreport_qr1", STR)
    STR = "insert into dislib.dimpresion (reqno,impression,finaldiag,fdescry,t,tsize,n,m) values (?,?,?,?,?,?,?,?)"
    Set insdimpresion_qr = rdb.CreateQuery("insdimpresion_qr", STR)
    STR = "insert into dislib.dspreport (reqno,spreport,enterby,preportby,freportby) values(?,?,?,?,?)"
    Set insdspreport_qr = rdb.CreateQuery("insdspreport_qr1", STR)

    STR = "update dislib.dfsreport set fsreport=? where reqno=?"
    Set upddfsreport_qr = rdb.CreateQuery("upddfsreport_qr1", STR)
    
    STR = "update dislib.dihreport set ihreport=?, enteryby=?, preportby=?, freportby=? where reqno=?"
    Set upddihreport_qr = rdb.CreateQuery("upddihreport_qr11", STR)
    
    STR = "update dislib.dimpresion set impression=?, finaldiag=?, fdescry=?, t=?,tsize=?, n=?, m=? where reqno=?"
    Set upddimpresion_qr = rdb.CreateQuery("upddimpresion_qr1", STR)
    
    STR = "UPDATE DISLIB.DSPREPORT SET SPREPORT=?, ENTERBY=?, PREPORTBY=?, FREPORTBY=? WHERE REQNO=?"
   '  str = "UPDATE DISLIB.DSPREPORT SET SPREPORT=? WHERE REQNO=?"
    Set UPDSPREPORT_QR = rdb.CreateQuery("UPDSPREPORT_QR1", STR)
    
    STR = "UPDATE DISLIB.DGROSSDESC SET GROSSDESC=?,GROSSEDBY=? WHERE REQNO=?"
    Set updgrossdesc_qr = rdb.CreateQuery("updgrossdesc_qr1", STR)
    
    STR = "SELECT * FROM DISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QR11", STR)
    SELRREQUISITN_QR(0) = "BN/00000"
    SELRREQUISITN_QR(1) = "IH"
    SELRREQUISITN_QR(2) = "FZZ"
    SELRREQUISITN_QR(3) = "T"
    SELRREQUISITN_QR(4) = "F"
    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
    
    
    
   STR = "update dislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1", STR)
   
   STR = "UPDATE DISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1", STR)
    
   STR = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1", STR)
   
   STR = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11", STR)
   SELDADDENDUM_QR(0) = "FSS"
   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
   
   STR = "select * from dislib.drequisitn where caseno=?"
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr11", STR)
   seldrequisitn_qr(0) = "BN/00000"
   Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
    lbl_labno.Visible = False
    cmb_labno.Visible = False
    txt_addendum.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_dept.Enabled = False
   FLDENABL ("DIS")
   txt_fullreport.Visible = False

End Sub
Private Sub Form_Unload(cancel As Integer)
Dim RS As rdoResultset
Dim RQ As rdoQuery
For Each RS In rdb.rdoResultsets
    RS.Close
Next

For Each RQ In rdb.rdoQueries
    RQ.Close
Next
End Sub



Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
CASENO = MaskEdBox1.Text
MaskEdBox1.Text = padzero(5, "C", CASENO, 3)
 If newflag = True Then

  req_qr(0) = MaskEdBox1.FormattedText
  req_rs.MoreResults
  req_rs.Requery
  If req_rs.RowCount > 0 Then
        cmb_reqno.Clear
        While Not req_rs.EOF
        cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
        
        req_rs.MoveNext
        
        Wend
  Else
    cmb_reqno.Clear
    
    End If
End If
If viewflag = True Then
    SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
    SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
    SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
    SELRREQUISITN_QR(3) = "T"
    SELRREQUISITN_QR(4) = "F"
    SELRREQUISITN_RS.MoreResults
    SELRREQUISITN_RS.Requery
    If SELRREQUISITN_RS.RowCount > 0 Then
        cmb_reqno.Clear
        While Not SELRREQUISITN_RS.EOF
            cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
            SELRREQUISITN_RS.MoveNext
        Wend
     Else
        MsgBox " No  Requisition for this Modality"
        btnenabl ("INI")
        Exit Sub
    End If

End If
 Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
If newflag = True Then
    newflag = False
    selta101_qr(0) = MaskEdBox1.FormattedText
    selta101_rs.MoreResults
    selta101_rs.Requery
    If selta101_rs.RowCount = 0 Then
        Exit Sub
    End If

    txt_sex.Text = selta101_rs("sex")
    txt_name.Text = Trim(selta101_rs("first_name")) & " " & Trim(selta101_rs("mid_name")) & " " & Trim(selta101_rs("last_name"))
    txt_age = DateDiff("YYYY", selta101_rs("BIRTH_DATE"), CStr(DATE11))
    txt_pridoc.Text = selta101_rs("ud_name")
    txt_category.Text = selta101_rs("pat_catg")
    
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    REFDR_RS.MoreResults
    REFDR_RS.Requery

    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
  
    selma104_qr(0) = MaskEdBox1.FormattedText
    selma104_rs.MoreResults
    selma104_rs.Requery
  
  
    If selma104_rs.RowCount > 0 Then
        txt_ward.Text = selma104_rs("WARD_NAME")
        txt_bedno.Text = selma104_rs("bed_no")
        txt_status.Text = "IP"
    Else
        txt_status.Text = "OP"
    End If
  
 
'SELPATHDREQUISITN_QR(0) = MaskEdBox1.FormattedText
    SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    SELPATHDREQUISITN_RS.MoreResults
    SELPATHDREQUISITN_RS.Requery
    If SELPATHDREQUISITN_RS.RowCount > 0 Then
        TXT_PATH.Text = SELPATHDREQUISITN_RS("PATHNO")
        If IsNull(Trim(SELPATHDREQUISITN_RS("SITE"))) Then
            txt_sites.Text = ""
        Else
            txt_sites.Text = SELPATHDREQUISITN_RS("SITE")
        End If
        If IsNull(Trim(SELPATHDREQUISITN_RS("LATERALITS"))) Then
            txt_laterality.Text = ""
        Else
            txt_laterality.Text = SELPATHDREQUISITN_RS("LATERALITS")
        End If
    End If

    selta1012_qr(0) = MaskEdBox1.FormattedText
    selta1012_rs.MoreResults
    selta1012_rs.Requery
    If selta1012_rs.RowCount > 0 Then
        txt_procedure.Text = selta1012_rs("prvdiagdes")
    End If
    seldrequisitn_qr(0) = MaskEdBox1.FormattedText
    seldrequisitn_rs.MoreResults
    seldrequisitn_rs.Requery
    If seldrequisitn_rs.RowCount > 0 Then
     prpathno = ""
     While Not seldrequisitn_rs.EOF
     prpathno = prpathno & " " & seldrequisitn_rs("pathno")
     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
     seldrequisitn_rs.MoveNext
     Wend
     txt_prepathreqno.Text = prpathno
    End If

End If
If viewflag = True Then

   
    
    viewflag = False
    selta101_qr(0) = MaskEdBox1.FormattedText
    selta101_rs.MoreResults
    selta101_rs.Requery
    If selta101_rs.RowCount = 0 Then
        Exit Sub
    End If

    txt_sex.Text = selta101_rs("sex")
    txt_name.Text = Trim(selta101_rs("first_name")) & " " & Trim(selta101_rs("mid_name")) & " " & Trim(selta101_rs("last_name"))
    txt_age = DateDiff("YYYY", selta101_rs("BIRTH_DATE"), CStr(DATE11))
    txt_pridoc.Text = selta101_rs("ud_name")
    txt_category.Text = selta101_rs("pat_catg")

    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    REFDR_RS.MoreResults
    REFDR_RS.Requery

    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
        If REFDR_RS("COMMITFLAG") = "Y" Then
            MsgBox "Final Report is Generated, You cann't Modify the Final Report,You can Add ADDENDUMS Only"
            ADDENDUMFALG = True
        End If
    End If
  
    
  
    selma104_qr(0) = MaskEdBox1.FormattedText
    selma104_rs.MoreResults
    selma104_rs.Requery
  
  
    If selma104_rs.RowCount > 0 Then
        txt_ward.Text = selma104_rs("WARD_NAME")
        txt_bedno.Text = selma104_rs("bed_no")
        txt_status.Text = "IP"
    Else
        txt_status.Text = "OP"
    End If
  
 
'SELPATHDREQUISITN_QR(0) = MaskEdBox1.FormattedText
    SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    SELPATHDREQUISITN_RS.MoreResults
    SELPATHDREQUISITN_RS.Requery
    If SELPATHDREQUISITN_RS.RowCount > 0 Then
        TXT_PATH.Text = SELPATHDREQUISITN_RS("PATHNO")
        If IsNull(SELPATHDREQUISITN_RS("SITE")) Then
        Else
        txt_sites.Text = SELPATHDREQUISITN_RS("SITE")
        End If
        If IsNull(SELPATHDREQUISITN_RS("LATERALITS")) Then
        Else
        txt_laterality.Text = SELPATHDREQUISITN_RS("LATERALITS")
        End If
    End If

    selta1012_qr(0) = MaskEdBox1.FormattedText
    selta1012_rs.MoreResults
    selta1012_rs.Requery
    If selta1012_rs.RowCount > 0 Then
        txt_procedure.Text = selta1012_rs("prvdiagdes")
    End If

    seldfsreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    seldfsreport_rs.MoreResults
    seldfsreport_rs.Requery
    If seldfsreport_rs.RowCount > 0 Then
        txt_frosecrep.Text = seldfsreport_rs("fsreport")
    End If

    seldihreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    seldihreport_rs.MoreResults
    seldihreport_rs.Requery
    If seldihreport_rs.RowCount > 0 Then
        txt_immhischerep.Text = seldihreport_rs("ihreport")
    End If


    seldimpresion_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    seldimpresion_rs.MoreResults
    seldimpresion_rs.Requery
    If seldimpresion_rs.RowCount > 0 Then
            txt_impfindiag.Text = seldimpresion_rs("impression")
            cmb_fsd.AddItem seldimpresion_rs("FDESCRY")
            cmb_fsd.Text = seldimpresion_rs("FDESCRY")
            cmb_t.Text = seldimpresion_rs("T")
            txt_size.Text = seldimpresion_rs("TSIZE")
            CMB_N.Text = seldimpresion_rs("N")
            cmb_m.Text = seldimpresion_rs("M")
            CMB_ICDH.AddItem seldimpresion_rs("FINALDIAG")
            CMB_ICDH.Text = seldimpresion_rs("FINALDIAG")
    End If

    seldgrossdesc_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    seldgrossdesc_rs.MoreResults
    seldgrossdesc_rs.Requery
    If seldgrossdesc_rs.RowCount > 0 Then
            txt_grossdesc = seldgrossdesc_rs("GROSSDESC")
            cmb_grossby.Text = seldgrossdesc_rs("GROSSEDBY")
    End If


    seldspreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    seldspreport_rs.MoreResults
    seldspreport_rs.Requery
    If seldspreport_rs.RowCount > 0 Then
            txt_microdesc = seldspreport_rs("SPREPORT")
            selempbasinfo1_qr(0) = seldspreport_rs("PREPORTBY")
            selempbasinfo1_rs.MoreResults
            selempbasinfo1_rs.Requery
        If selempbasinfo1_rs.RowCount > 0 Then
            cmb_provby.Clear
            cmb_provby.AddItem selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("PREPORTBY")
            cmb_provby.Text = selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("PREPORTBY")
        End If
        selempbasinfo1_qr(0) = seldfsreport_rs("FREPORTBY")
        selempbasinfo1_rs.MoreResults
        selempbasinfo1_rs.Requery
        If selempbasinfo1_rs.RowCount > 0 Then
            cmb_finalby.Clear
            cmb_finalby.AddItem selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("FREPORTBY")
            cmb_finalby.Text = selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("FREPORTBY")
        End If
        
    seldrequisitn_qr(0) = MaskEdBox1.FormattedText
    seldrequisitn_rs.MoreResults
    seldrequisitn_rs.Requery
    If seldrequisitn_rs.RowCount > 0 Then
     prpathno = ""
     While Not seldrequisitn_rs.EOF
        prpathno = prpathno & " " & seldrequisitn_rs("pathno")
        CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
        seldrequisitn_rs.MoveNext
     Wend
        txt_prepathreqno.Text = Trim(prpathno)
    End If
        SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
        SELDADDENDUM_RS.MoreResults
        SELDADDENDUM_RS.Requery
        If SELDADDENDUM_RS.RowCount > 0 Then
            While Not SELDADDENDUM_RS.EOF
                ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
            SELDADDENDUM_RS.MoveNext
            Wend
            txt_addendum.Text = Trim(ADDENDUM)
            
        End If
    End If
    cmb_reqno.Enabled = False
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
End If



End Sub
Private Sub txt_addendum_GotFocus()
On Error GoTo m11:
w_adden.Visible = False
Set d_adden = w_adden.ActiveDocument
d_adden.Content = ""
d_adden.Content = txt_addendum
w_adden.Caption = "Addendum"
w_adden.WindowState = wdWindowStateMaximize
d_adden.Content = txt_addendum
w_adden.Visible = True
w_adden.Activate
m11:
On Error Resume Next
Set w_adden = CreateObject("Word.Application")
w_adden.Documents.Add ("C:\risreports\REPORT.DOT")
w_adden.Caption = "frozen"
w_adden.WindowState = wdWindowStateMaximize
Set d_adden = w_adden.ActiveDocument
d_adden.Content = Trim(txt_addendum.Text)
w_adden.Visible = True
w_adden.Activate
End Sub

Private Sub txt_addendum_LostFocus()
On Error GoTo m12:
Set d_adden = w_adden.ActiveDocument
txt_addendum.Text = d_adden.Content

d_adden.Content = ""
d_adden.SaveAs "C:\fronzen.DOC"
d_adden.Close
w_adden.Quit

Exit Sub
m12:
End Sub

Private Sub txt_frosecrep_GotFocus()
On Error GoTo m1:
w_frozen.Visible = False
Set d_frozen = w_frozen.ActiveDocument

d_frozen.Content = ""
d_frozen.Content = txt_frosecrep
w_frozen.Caption = "frozen"
w_frozen.WindowState = wdWindowStateMaximize
d_frozen.Content = txt_frosecrep
w_frozen.Visible = True
w_frozen.Activate
m1:
On Error Resume Next
Set w_frozen = CreateObject("Word.Application")
w_frozen.Documents.Add ("C:\Risreports\REPORT.DOT")
w_frozen.Caption = "Frozen Section Report"
w_frozen.WindowState = wdWindowStateMaximize
Set d_frozen = w_frozen.ActiveDocument
d_frozen.Content = Trim(txt_frosecrep.Text)
w_frozen.Visible = True
w_frozen.Activate
End Sub

Private Sub txt_frosecrep_LostFocus()
On Error GoTo m2:
Set d_frozen = w_frozen.ActiveDocument
txt_frosecrep.Text = d_frozen.Content
'flagimp = False
d_frozen.Content = ""
d_frozen.SaveAs "C:\fronzen.DOC"
d_frozen.Close
w_frozen.Quit
'W1.Visible = False
Exit Sub
m2:
End Sub



Private Sub txt_fullreport_GotFocus()
On Error GoTo m14:
w.Visible = False
Set d = w.ActiveDocument
d.Content = ""
d.Content = txt_fullreport
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
d.Content = txt_fullreport
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.Add ("c:\risreports\REPORT.dot")
w.Caption = "Full report"
w.WindowState = wdWindowStateMaximize
Set d = w.ActiveDocument
d.Content = txt_fullreport.Text
w.Visible = True
w.Activate
End Sub

Private Sub txt_fullreport_LostFocus()
On Error GoTo m15:
Set d = w.ActiveDocument
txt_fullreport.Text = d.Content
'flagimp = False
d.Content = ""
d.SaveAs "C:\fronzen.DOC"
d.Close
w.Quit
'W1.Visible = False
Exit Sub
m15:
End Sub

Private Sub txt_grossdesc_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_grossdesc
W_grossdes.Caption = "GrossDesc"
W_grossdes.WindowState = wdWindowStateMaximize
d_grossdesc.Content = txt_grossdesc
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.Add ("c:\risreports\REPORT.dot")
W_grossdes.Caption = "GrossDesc"
W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument

d_grossdesc.Content = txt_grossdesc.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub

Private Sub txt_grossdesc_LostFocus()
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument
txt_grossdesc.Text = d_grossdesc.Content
'flagimp = False
d_grossdesc.Content = ""
d_grossdesc.SaveAs "C:\fronzen.DOC"
d_grossdesc.Close
W_grossdes.Quit
'W1.Visible = False
Exit Sub
m4:
End Sub

Private Sub txt_immhischerep_GotFocus()
On Error GoTo m7:
w_immun.Visible = False
Set d_immun = w_immun.ActiveDocument
d_immun.Content = ""
d_immun.Content = txt_immhischerep
w_immun.Caption = "Immuno"
w_immun.WindowState = wdWindowStateMaximize
d_immun.Content = txt_immhischerep
w_immun.Visible = True
w_immun.Activate
m7:
On Error Resume Next:
Set w_immun = CreateObject("word.Application")
w_immun.Documents.Add ("c:\risreports\REPORT.dot")
w_immun.Caption = "Immuno"
w_immun.WindowState = wdWindowStateMaximize
Set d_immun = w_immun.ActiveDocument
d_immun.Content = txt_immhischerep.Text
w_immun.Visible = True
End Sub

Private Sub txt_immhischerep_LostFocus()
On Error GoTo m8:
Set d_immun = w_immun.ActiveDocument
txt_immhischerep.Text = d_immun.Content
'flagimp = False
d_immun.Content = ""
d_immun.SaveAs "C:\fronzen.DOC"
d_immun.Close
w_immun.Quit
'W1.Visible = False
Exit Sub
m8:
End Sub



Private Sub txt_impfindiag_GotFocus()
On Error GoTo m9:
w_imp.Visible = False
Set d_imp = w_imp.ActiveDocument
d_imp.Content = ""
d_imp.Content = txt_impfindiag
w_imp.Caption = "Impression"
w_imp.WindowState = wdWindowStateMaximize
d_imp.Content = txt_impfindiag
w_imp.Visible = True
w_imp.Activate
m9:
On Error Resume Next:
Set w_imp = CreateObject("word.Application")
w_imp.Documents.Add ("c:\risreports\REPORT.dot")
w_imp.Caption = "Impression"
w_imp.WindowState = wdWindowStateMaximize
Set d_imp = w_imp.ActiveDocument
d_imp.Content = txt_impfindiag
w_imp.Visible = True
w_imp.Activate
End Sub

Private Sub txt_impfindiag_LostFocus()
On Error GoTo m12:
Set d_imp = w_imp.ActiveDocument
txt_impfindiag.Text = d_imp.Content
'flagimp = False
d_imp.Content = ""
d_imp.SaveAs "C:\fronzen.DOC"
d_imp.Close
w_imp.Quit
'W1.Visible = False
Exit Sub
m12:
End Sub

Private Sub txt_microdesc_GotFocus()
On Error GoTo m5:
w_micscop.Visible = False
Set d_micscop = w_micscop.ActiveDocument
d_micscop.Content = ""
d_micscop.Content = txt_microdesc
w_micscop.Caption = "MicroDesc"
w_micscop.WindowState = wdWindowStateMaximize
d_micscop.Content = txt_microdesc
w_micscop.Visible = True
w_micscop.Activate
m5:
On Error Resume Next:
Set w_micscop = CreateObject("word.Application")
w_micscop.Documents.Add ("c:\risreports\REPORT.dot")
w_micscop.Caption = "MicroDesc"
w_micscop.WindowState = wdWindowStateMaximize
Set d_micscop = w_micscop.ActiveDocument
d_micscop.Content = txt_microdesc.Text
w_micscop.Visible = True
w_micscop.Activate
End Sub

Private Sub txt_microdesc_LostFocus()
On Error GoTo m6:
Set d_micscop = w_micscop.ActiveDocument
txt_microdesc.Text = d_micscop.Content
'flagimp = False
d_micscop.Content = ""
d_micscop.SaveAs "C:\fronzen.DOC"
d_micscop.Close
w_micscop.Quit
'W1.Visible = False
Exit Sub
m6:
End Sub

Public Sub save()

End Sub

Public Sub assign()
On Error GoTo MSG:
Dim I As Integer
Dim MICROSCOPIC As String
If flagsave = 1 Then
   flagsave = 0
If Len(txt_frosecrep.Text) > 0 Then
    insdfs_qr(0) = Trim(req_chglen(cmb_reqno))
    insdfs_qr(1) = Null
    insdfs_qr(1) = txt_frosecrep
    I = InStr(cmb_enteryby.Text, "$")
    insdfs_qr(2) = Mid(cmb_enteryby, I + 1, 6)
    I = InStr(cmb_provby.Text, "$")
    insdfs_qr(3) = Mid(cmb_provby, I + 1, 6)
    I = InStr(cmb_finalby.Text, "$")
    insdfs_qr(4) = Mid(cmb_finalby, I + 1, 6)
    insdfs_qr.Execute
End If

If Len(txt_grossdesc.Text) > 0 Then
    insdgrossdesc_qr(0) = Trim(req_chglen(cmb_reqno))
    insdgrossdesc_qr(1) = txt_grossdesc
    insdgrossdesc_qr(2) = cmb_grossby
    insdgrossdesc_qr.Execute
End If
        
 If Len(txt_immhischerep.Text) > 0 Then
    insdihreport_qr(0) = Trim(req_chglen(cmb_reqno))
    insdihreport_qr(1) = Null
    insdihreport_qr(1) = txt_immhischerep
    I = InStr(cmb_enteryby.Text, "$")
    insdihreport_qr(2) = Mid(cmb_enteryby, I + 1, 6)
    I = InStr(cmb_provby, "$")
    insdihreport_qr(3) = Mid(cmb_provby, I + 1, 6)
    I = InStr(cmb_finalby, "$")
    insdihreport_qr(4) = Mid(cmb_finalby, I + 1, 6)
    insdihreport_qr.Execute
End If
    
If Len(txt_microdesc) > 0 Then
    insdspreport_qr(0) = Trim(req_chglen(cmb_reqno))
    insdspreport_qr(1) = Null
    insdspreport_qr(1) = txt_microdesc
    I = InStr(cmb_enteryby.Text, "$")
    insdspreport_qr(2) = Mid(cmb_enteryby, I + 1, 6)
    I = InStr(cmb_provby.Text, "$")
    insdspreport_qr(3) = Mid(cmb_provby, I + 1, 6)
    I = InStr(cmb_finalby, "$")
    insdspreport_qr(4) = Mid(cmb_finalby, I + 1, 6)
    insdspreport_qr.Execute
End If
    
If Len(txt_impfindiag) > 0 Then
    insdimpresion_qr(0) = Trim(req_chglen(cmb_reqno))
    insdimpresion_qr(1) = Null
    insdimpresion_qr(1) = txt_impfindiag
    insdimpresion_qr(2) = CMB_ICDH.Text
    I = InStr(cmb_fsd.Text, "$")
    insdimpresion_qr(3) = Mid(cmb_fsd, I + 1, 6)
    insdimpresion_qr(4) = cmb_t.Text
    insdimpresion_qr(5) = txt_size.Text
    insdimpresion_qr(6) = CMB_N
    insdimpresion_qr(7) = cmb_m
    insdimpresion_qr.Execute
End If

    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provby.Text
'    updrrequisitn_qr(3) = cmb_finalby.Text
'    updrrequisitn_qr(4) = GET_REVDATE(Format(date11, "DD/MM/YYYY"))
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    End If
      
    
If flagsave = 2 Then
    flagsave = 0
If Len(txt_frosecrep.Text) > 0 Then
    
    upddfsreport_qr(0) = Null
    upddfsreport_qr(0) = Trim(txt_frosecrep.Text)
    upddfsreport_qr(1) = Trim(req_chglen(cmb_reqno))
    upddfsreport_qr.Execute
End If

If Len(txt_grossdesc.Text) > 0 Then

    updgrossdesc_qr(0) = Null
    updgrossdesc_qr(0) = txt_grossdesc
    updgrossdesc_qr(1) = cmb_grossby
    updgrossdesc_qr(2) = Trim(req_chglen(cmb_reqno))
    updgrossdesc_qr.Execute
 End If

If Len(txt_immhischerep.Text) > 0 Then
    
    upddihreport_qr(0) = Null
    upddihreport_qr(0) = Trim(txt_immhischerep.Text)
    I = InStr(cmb_enteryby.Text, "$")
    upddihreport_qr(1) = Mid(cmb_enteryby, I + 1, 6)
    I = InStr(cmb_provby, "$")
    upddihreport_qr(2) = Mid(cmb_provby, I + 1, 6)
    I = InStr(cmb_finalby, "$")
    upddihreport_qr(3) = Mid(cmb_finalby, I + 1, 6)
    upddihreport_qr(4) = Trim(req_chglen(cmb_reqno))
    upddihreport_qr.Execute
    End If

    If Len(txt_microdesc) > 0 Then
    MICROSCOPIC = RTrim(txt_microdesc.Text)
    UPDSPREPORT_QR(0) = Null
    UPDSPREPORT_QR(0) = MICROSCOPIC
    I = InStr(cmb_enteryby.Text, "$")
    UPDSPREPORT_QR(1) = Mid(cmb_enteryby.Text, I + 1, 6)
    I = InStr(cmb_provby.Text, "$")
    UPDSPREPORT_QR(2) = Mid(cmb_provby.Text, I + 1, 6)
    I = InStr(cmb_finalby.Text, "$")
    UPDSPREPORT_QR(3) = Mid(cmb_finalby.Text, I + 1, 6)
    UPDSPREPORT_QR(4) = req_chglen(Trim(cmb_reqno.Text))
    UPDSPREPORT_QR.Execute
    
    End If

    If Len(txt_impfindiag) > 0 Then
    
    upddimpresion_qr(0) = Null
    upddimpresion_qr(0) = txt_impfindiag
    
    upddimpresion_qr(1) = Trim(CMB_ICDH.Text)
    
    upddimpresion_qr(2) = Trim(cmb_fsd.Text)
    upddimpresion_qr(3) = cmb_t.Text
    upddimpresion_qr(4) = txt_size.Text
    upddimpresion_qr(5) = CMB_N
    upddimpresion_qr(6) = cmb_m
    upddimpresion_qr(7) = Trim(req_chglen(cmb_reqno))
    upddimpresion_qr.Execute
    End If
    
    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provby.Text
'    updrrequisitn_qr(3) = cmb_finalby.Text
'    updrrequisitn_qr(4) = GET_REVDATE(Format(date11, "DD/MM/YYYY"))
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    
    End If
    If ADDENDUMFALG = True Then
        ADDENDUMFALG = False
        SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
        SELDADDENDUM_RS.MoreResults
        SELDADDENDUM_RS.Requery
        If SELDADDENDUM_RS.RowCount > 0 Then
            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INSDADDENDUM_QR(3) = time1
            INSDADDENDUM_QR(4) = cmb_addendumby.Text
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = txt_addendum.Text
            INSDADDENDUM_QR.Execute
        Else
            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            INSDADDENDUM_QR(1) = "1"
            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INSDADDENDUM_QR(3) = time1
            INSDADDENDUM_QR(4) = cmb_addendumby.Text
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = txt_addendum.Text
            INSDADDENDUM_QR.Execute
        End If
    End If
    
Exit Sub
MSG:
    MsgBox Err.Description
    
    
End Sub

Public Sub field_clear()
Dim obj As Object
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
cmb_reqno.ListIndex = -1
cmb_reportdoc.ListIndex = -1
cmb_addendumby.ListIndex = -1
cmb_enteryby.ListIndex = -1
cmb_provby.ListIndex = -1
cmb_finalby.ListIndex = -1
cmb_addendumby.ListIndex = -1
cmb_immhische.ListIndex = -1
cmb_t.ListIndex = -1
cmb_t.Clear
CMB_N.Clear
CMB_N.ListIndex = -1
cmb_m.ListIndex = -1
cmb_grossby.ListIndex = -1
cmb_fsd.ListIndex = -1
CMB_ICDH.ListIndex = -1
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
MaskEdBox1.Mask = ">??/#####"
cmb_grossby.Clear

End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In sppathreport.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub

Public Sub FLDENABL(STR1 As String)
If STR1 = "ENA" Then
    txt_frosecrep.Enabled = True
    txt_grossdesc.Enabled = True
    txt_microdesc.Enabled = True
    txt_immhischerep.Enabled = True
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = True
    cmb_immhische.Enabled = True
    txt_addendum.Enabled = True
    txt_prepathreqno.Enabled = False
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provby.Enabled = True
    cmb_addendumby.Enabled = True
    cmb_t.Enabled = True
    cmb_m.Enabled = True
    CMB_N.Enabled = True
    txt_size.Enabled = True
    CMB_ICDH.Enabled = True
    cmb_grossby.Enabled = True
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_reportdoc.Enabled = True
    cmb_dept.Enabled = True
    cmb_fsd.Enabled = True
    cmb_reqno.Enabled = True
End If
If STR1 = "DIS" Then
    txt_frosecrep.Enabled = False
    txt_grossdesc.Enabled = False
    txt_microdesc.Enabled = False
    txt_immhischerep.Enabled = False
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = False
    cmb_immhische.Enabled = False
    txt_addendum.Enabled = False
    txt_prepathreqno.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_t.Enabled = False
    cmb_m.Enabled = False
    CMB_N.Enabled = False
    txt_size.Enabled = False
    CMB_ICDH.Enabled = False
    cmb_grossby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = False
End If
If STR1 = "VIEW" Then
    txt_frosecrep.Enabled = False
    txt_grossdesc.Enabled = False
    txt_microdesc.Enabled = False
    txt_immhischerep.Enabled = False
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = False
    cmb_immhische.Enabled = False
    txt_addendum.Enabled = False
    txt_prepathreqno.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_t.Enabled = False
    cmb_m.Enabled = False
    CMB_N.Enabled = False
    txt_size.Enabled = False
    CMB_ICDH.Enabled = False
    cmb_grossby.Enabled = False
    cmb_modality.Enabled = False
    
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
End If
If STR1 = "MODIFY" Then
    txt_frosecrep.Enabled = True
    txt_grossdesc.Enabled = True
    txt_microdesc.Enabled = True
    txt_immhischerep.Enabled = True
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = True
    cmb_immhische.Enabled = False
    txt_addendum.Enabled = False
    txt_prepathreqno.Enabled = False
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provby.Enabled = True
    cmb_addendumby.Enabled = False
    cmb_t.Enabled = True
    cmb_m.Enabled = True
    CMB_N.Enabled = True
    txt_size.Enabled = True
    CMB_ICDH.Enabled = True
    cmb_grossby.Enabled = True
       
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = True
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
End If
If STR1 = "ADDENDUM" Then
    txt_frosecrep.Enabled = False
    txt_grossdesc.Enabled = False
    txt_microdesc.Enabled = False
    txt_immhischerep.Enabled = False
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = False
    cmb_immhische.Enabled = False
    txt_addendum.Enabled = True
    txt_prepathreqno.Enabled = False
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = True
    cmb_t.Enabled = False
    cmb_m.Enabled = False
    CMB_N.Enabled = False
    txt_size.Enabled = False
    CMB_ICDH.Enabled = False
    cmb_grossby.Enabled = False
       
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
End If
End Sub
