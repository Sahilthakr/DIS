VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FRM_PFTTEST 
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   9645
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame common 
      Height          =   2895
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   11775
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   35
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   0
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   32
         Top             =   2160
         Width           =   3855
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
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   30
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   29
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   28
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   27
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   54
         Top             =   2520
         Width           =   1830
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   52
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   51
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   50
         Top             =   1800
         Width           =   705
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
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   47
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   46
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   45
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   44
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   43
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Path No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   41
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   40
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Width           =   1830
      End
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   11880
      TabIndex        =   15
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   11880
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7200
         MouseIcon       =   "FRM_PFTTEST.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10440
         MouseIcon       =   "FRM_PFTTEST.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8760
         MouseIcon       =   "FRM_PFTTEST.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5880
         MouseIcon       =   "FRM_PFTTEST.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "FRM_PFTTEST.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "FRM_PFTTEST.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MouseIcon       =   "FRM_PFTTEST.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         MouseIcon       =   "FRM_PFTTEST.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame12 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   11775
      Begin VB.ComboBox cmb_provisional 
         Height          =   315
         Left            =   5880
         TabIndex        =   63
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmb_enterby 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         Left            =   9600
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reported by"
         Height          =   195
         Left            =   4560
         TabIndex        =   64
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lbl_enter 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
         Height          =   195
         Left            =   7920
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame FRM_PFT 
      Caption         =   "PFT Test"
      ForeColor       =   &H000000C0&
      Height          =   3495
      Left            =   120
      TabIndex        =   55
      Top             =   3720
      Visible         =   0   'False
      Width           =   11775
      Begin VB.ComboBox cmb_pfproperby 
         Height          =   315
         Left            =   8520
         TabIndex        =   131
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txt_pftindication 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   116
         Top             =   360
         Width           =   6615
      End
      Begin VB.TextBox TXT_PFTREPORT 
         Appearance      =   0  'Flat
         Height          =   1125
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   840
         Width           =   9255
      End
      Begin VB.TextBox TXT_PFTIMPRESSION 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   2040
         Width           =   9255
      End
      Begin MSMask.MaskEdBox msk_pfexamdt 
         Height          =   375
         Left            =   9840
         TabIndex        =   124
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_pfproperby 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Performed by"
         Height          =   195
         Left            =   6360
         TabIndex        =   130
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   720
         TabIndex        =   127
         Top             =   2400
         Width           =   120
      End
      Begin VB.Label lbl_pfexamdt 
         AutoSize        =   -1  'True
         Caption         =   "Exam Date"
         Height          =   195
         Left            =   8760
         TabIndex        =   123
         Top             =   480
         Width           =   930
      End
      Begin VB.Label lbl_pftindication 
         AutoSize        =   -1  'True
         Caption         =   "Indication"
         Height          =   195
         Left            =   960
         TabIndex        =   115
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LBL_PFTREPORT 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1200
         TabIndex        =   59
         Top             =   960
         Width           =   585
      End
      Begin VB.Label LBL_PFTIMPRESSION 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   960
         TabIndex        =   58
         Top             =   2400
         Width           =   915
      End
   End
   Begin VB.Frame FRM_ECG 
      Caption         =   "ECG REPORT"
      ForeColor       =   &H000000C0&
      Height          =   3495
      Left            =   0
      TabIndex        =   90
      Top             =   3720
      Visible         =   0   'False
      Width           =   11775
      Begin VB.ComboBox cmb_edproperby 
         Height          =   315
         Left            =   8640
         TabIndex        =   129
         Top             =   1800
         Width           =   2775
      End
      Begin MSMask.MaskEdBox msk_ecexamdt 
         Height          =   375
         Left            =   8520
         TabIndex        =   120
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TXT_ECIMPRESSION 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   102
         Top             =   3000
         Width           =   9735
      End
      Begin VB.TextBox TXT_TWAVE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   101
         Top             =   2520
         Width           =   9735
      End
      Begin VB.TextBox TXT_STSEGMENT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   100
         Top             =   2160
         Width           =   9735
      End
      Begin VB.TextBox TXT_QTINTERVAL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   99
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox TXT_QRSD 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   98
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox TXT_QRSM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   97
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox TXT_PWAVE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   95
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TXT_VENRATE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   93
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TXT_ATRIALRATE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   92
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TXT_RHYTHM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   94
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TXT_PRINTERVAL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   96
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TXT_RATE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   91
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl_ecproperby 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Performed by"
         Height          =   195
         Left            =   6480
         TabIndex        =   128
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   360
         TabIndex        =   125
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label lbl_ecexamdt 
         AutoSize        =   -1  'True
         Caption         =   "Exam Date"
         Height          =   195
         Left            =   7440
         TabIndex        =   119
         Top             =   480
         Width           =   930
      End
      Begin VB.Label LBL_ECIMPRESSION 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   600
         TabIndex        =   114
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label LBL_TWAVE 
         AutoSize        =   -1  'True
         Caption         =   "T Wave"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   840
         TabIndex        =   113
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label LBL_STSEGMENT 
         AutoSize        =   -1  'True
         Caption         =   "ST Segment "
         Height          =   195
         Left            =   480
         TabIndex        =   112
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label LBL_QRSD 
         AutoSize        =   -1  'True
         Caption         =   "QRS duration"
         Height          =   195
         Left            =   3960
         TabIndex        =   111
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label LBL_QTINT 
         AutoSize        =   -1  'True
         Caption         =   "QT interval "
         Height          =   195
         Left            =   600
         TabIndex        =   110
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label LBL_QRSM 
         AutoSize        =   -1  'True
         Caption         =   "QRS Morphology"
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label LBL_PWAVE 
         AutoSize        =   -1  'True
         Caption         =   "Pwave"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   960
         TabIndex        =   108
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label LBL_VENRATE 
         AutoSize        =   -1  'True
         Caption         =   "Ventricular rate"
         Height          =   195
         Left            =   240
         TabIndex        =   107
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label LBL_ATRIALRATE 
         AutoSize        =   -1  'True
         Caption         =   "Atrial rate"
         Height          =   195
         Left            =   4200
         TabIndex        =   106
         Top             =   480
         Width           =   840
      End
      Begin VB.Label LBL_RHYTHM 
         AutoSize        =   -1  'True
         Caption         =   "Rhythm"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4440
         TabIndex        =   105
         Top             =   840
         Width           =   645
      End
      Begin VB.Label LBL_PRINTERVAL 
         AutoSize        =   -1  'True
         Caption         =   "PR Interval"
         Height          =   195
         Left            =   4080
         TabIndex        =   104
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label LBL_RATE 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
         Height          =   195
         Left            =   1080
         TabIndex        =   103
         Top             =   480
         Width           =   420
      End
   End
   Begin VB.Frame FRM_STRESSTEST 
      Caption         =   "Stress Test"
      ForeColor       =   &H000000C0&
      Height          =   3495
      Left            =   0
      TabIndex        =   65
      Top             =   3720
      Visible         =   0   'False
      Width           =   11775
      Begin VB.ComboBox cmb_stproperformed 
         Height          =   315
         Left            =   9360
         TabIndex        =   77
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox TXT_STIMPRESSION 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   76
         Top             =   2640
         Width           =   9855
      End
      Begin VB.TextBox TXT_METSA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   75
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TXT_THRACHIEVED 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10080
         TabIndex        =   71
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TXT_DOUBLEPRODUCT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   70
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TXT_STTCHANGES 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   72
         Top             =   1440
         Width           =   9855
      End
      Begin VB.TextBox TXT_ARRYTHMIA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   73
         Top             =   1800
         Width           =   9855
      End
      Begin VB.ComboBox CMB_RESFORTREAT 
         Height          =   315
         Left            =   9480
         TabIndex        =   74
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TXT_MAXHR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   68
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox CMB_RISKFACTOR 
         Height          =   315
         Left            =   1560
         TabIndex        =   67
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox TXT_MAXBP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9360
         TabIndex        =   69
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txt_stindication 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   66
         Top             =   240
         Width           =   9855
      End
      Begin MSMask.MaskEdBox msk_stexamdt 
         Height          =   375
         Left            =   1560
         TabIndex        =   122
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   240
         TabIndex        =   126
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label lbl_stexamdt 
         AutoSize        =   -1  'True
         Caption         =   "Exam Date"
         Height          =   195
         Left            =   480
         TabIndex        =   121
         Top             =   3120
         Width           =   930
      End
      Begin VB.Label lbl_stprocedureperformed 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Performed By"
         Height          =   195
         Left            =   7200
         TabIndex        =   118
         Top             =   3120
         Width           =   2070
      End
      Begin VB.Label lbl_stindication 
         Caption         =   "Indication"
         Height          =   255
         Left            =   600
         TabIndex        =   117
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LBL_STIMPRESSION 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   480
         TabIndex        =   89
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label LBL_METSACH 
         AutoSize        =   -1  'True
         Caption         =   "METS achieved"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   88
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label LBL_THRACH 
         AutoSize        =   -1  'True
         Caption         =   "% of THR achieved"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   87
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label LBL_DOUBLEPRODUCT 
         AutoSize        =   -1  'True
         Caption         =   "Double Product"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label LBL_STTCHANGES 
         AutoSize        =   -1  'True
         Caption         =   "ST /T changes"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   85
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label LBL_ARRYTHMIA 
         AutoSize        =   -1  'True
         Caption         =   "Arrythmia"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   84
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label LBL_RESFORTREATMENT 
         AutoSize        =   -1  'True
         Caption         =   "Reason for termination"
         Height          =   195
         Left            =   7440
         TabIndex        =   83
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label LBL_MAXHR 
         AutoSize        =   -1  'True
         Caption         =   "Max HR"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4560
         TabIndex        =   82
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "/Min"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6240
         TabIndex        =   81
         Top             =   840
         Width           =   405
      End
      Begin VB.Label LBL_STRISKFACTOR 
         AutoSize        =   -1  'True
         Caption         =   "Risk Factor"
         Height          =   195
         Left            =   480
         TabIndex        =   80
         Top             =   720
         Width           =   990
      End
      Begin VB.Label LBL_MAXBP 
         AutoSize        =   -1  'True
         Caption         =   "Max. BP "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8520
         TabIndex        =   79
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "mm Hg"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10800
         TabIndex        =   78
         Top             =   720
         Width           =   585
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   -480
      TabIndex        =   132
      Top             =   0
      Width           =   11655
   End
   Begin VB.Label LBL_CP 
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
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   4560
      TabIndex        =   62
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11400
      TabIndex        =   61
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11400
      TabIndex        =   60
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "FRM_PFTTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim testorder As String
Dim sel_deptqr As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SELDEPTRSFLAG1 As Boolean
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim SEL_MODALITYFLAG1 As Boolean
Dim sel_casno As rdoQuery
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim REQFLAG1 As Boolean
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim selta101flag1 As Boolean
Dim repdr_qr As rdoQuery
Dim repdr_rs As rdoResultset
Dim REPDRFLAG As Boolean
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim REFDRFLAG1 As Boolean
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset
Dim SELMA104FLAG1 As Boolean
Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset
Dim SELPATHDREQUISITNFALG1 As Boolean
Dim selma302_qr As rdoQuery
Dim selma302_rs As rdoResultset
Dim SELMA302FLAG1 As Boolean
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELTA1012FLAG1 As Boolean
'Dim SELEMPLIB_QR As rdoQuery
'Dim SELEMPLIB_RS As rdoResultset
Dim I As Integer
'Dim insdfsrep_qr As rdoQuery
'Dim insdgrossdesc_qr As rdoQuery
'Dim insdfs_qr As rdoQuery
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean
'Dim selempbasinfo1_qr As rdoQuery
'Dim selempbasinfo1_rs As rdoResultset
Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim INSDADDENDUM_QR As rdoQuery
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG1 As Boolean

Dim SELINDICATIONQR As rdoQuery
Dim SELINDICATIONRS As rdoResultset
Dim SELINDICATIONFLAG As Boolean

Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer
Dim ADDENDUMFALG As Boolean

Dim seldissecentr1_qr As rdoQuery
Dim seldissecentr1_rs As rdoResultset
Dim seldissecentrflag1 As Boolean

Dim seldissecprov1_qr As rdoQuery
Dim seldissecprov1_rs As rdoResultset
Dim seldissecprovflag1 As Boolean

Dim seldissecfin1_qr As rdoQuery
Dim seldissecfin1_rs As rdoResultset
Dim seldissecfinflag1 As Boolean

Dim SELDREQSTAFFDQR As rdoQuery
Dim SELDREQSTAFFDRS As rdoResultset
Dim SELDREQSTAFFD1FLAG As Boolean
Dim upddreqnoqr As rdoQuery

Dim selMA101QR As rdoQuery
Dim SELMA101RS As rdoResultset
Dim SELMA101FLAG As Boolean

Dim SELDMA101QR As rdoQuery
Dim SELDMA101RS As rdoResultset
Dim SELDMA101FLAG As Boolean

'Dim updrrequisitn1_qr As rdoQuery
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean

Dim INSPFREPORTQR As rdoQuery
Dim UPDPFREPORTQR As rdoQuery
Dim SELPFREPORTQR As rdoQuery
Dim SELPFREPORTRS As rdoResultset
Dim SELPFREPORTFLAG As Boolean

Dim INSSTREPORTQR As rdoQuery
Dim UPDSTREPORTQR As rdoQuery
Dim SELSTREPORTQR As rdoQuery
Dim SELSTREPORTRS As rdoResultset
Dim SELSTREPORTFLAG As Boolean
 
Dim INSECREPORTQR As rdoQuery
Dim UPDECREPORTQR As rdoQuery
Dim SELECREPORTQR As rdoQuery
Dim SELECREPORTRS As rdoResultset
Dim SELECREPORTFLAG As Boolean

Dim w_pwave As Application, w_rhythm As Application, w_stsegment As Application, w_twave As Application, w_impression As Application
Dim d_pwave As Object, d_rhythm As Object, d_stsegment As Object, d_twave As Object, d_impression As Object


Dim RNAGE As String
Dim SS As String
Dim ALKLV As Integer
Dim ALKHV As Integer
Dim LV As Integer
Dim HV As Integer

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

''******SECURITY***********



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
    If SEL_MODALITYFLAG1 = False Then
        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
        Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
        SEL_MODALITYFLAG1 = True
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
MsgBox ERROR(Err)
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(ind_cbodept)) = "GEN. MEDICINE SER.$CAZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_edproperby_LostFocus()
cmb_provisional.Text = cmb_edproperby.Text
End Sub

Private Sub cmb_modality_Click()
Dim MODALITY As String
Dim moda As Integer
Dim l As Integer
On Error GoTo MSG:
If Trim(cmb_modality) = "" Then
    Exit Sub
End If
moda = InStr(cmb_modality, "$")
l = Len(cmb_modality.Text)
txt_modcode.Text = Right(cmb_modality.Text, (l - moda))
If txt_modcode.Text = "EK" Then
    FRM_ECG.Visible = True
    FRM_PFT.Visible = False
    FRM_STRESSTEST.Visible = False
    FRM_PFTTEST.Caption = "Ecg Report"
    LBL_CP.Left = 4100
    LBL_CP.Caption = "ECG REPORT"
    TXT_ECIMPRESSION.Text = "With in Normal limit."
    If viewflag = True Then
    TXT_ECIMPRESSION.Text = ""
     End If
    If NEWFLAG = True Then
    TXT_ECIMPRESSION.Text = "With in Normal limit."
     End If
End If
If txt_modcode.Text = "ST" Then
    FRM_ECG.Visible = False
    FRM_PFT.Visible = False
    FRM_STRESSTEST.Visible = True
    FRM_PFTTEST.Caption = "Stress Test Report"
    LBL_CP.Left = 3600
    LBL_CP.Caption = "STRESS TEST REPORT"
End If
If txt_modcode.Text = "PF" Then
    FRM_ECG.Visible = False
    FRM_PFT.Visible = True
    FRM_STRESSTEST.Visible = False
    FRM_PFTTEST.Caption = "Pft Report"
    LBL_CP.Left = 4100
    LBL_CP.Caption = "PFT REPORT"
End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub



Private Sub cmb_pfproperby_LostFocus()
cmb_provisional.Text = cmb_pfproperby.Text
End Sub

Private Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
Dim entr As String
On Error GoTo MSG:
If Opt_patient.Value = True Then
        If NEWFLAG = True Then
            NEWFLAG = False
            If selta101flag1 = False Then
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101flag1 = True
            Else
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
            End If
            If SELTA101_RS.RowCount = 0 Then
                Exit Sub
            End If
        
            txt_sex.Text = SELTA101_RS("sex")
            txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
            txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            txt_pridoc.Text = SELTA101_RS("ud_name")
            txt_category.Text = SELTA101_RS("pat_catg")
            
            If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
            End If
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                lbl_ward.Visible = True
                lbl_bed.Visible = True
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_status.Text = "OP"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
            End If
          
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF Then
                    txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
                'txt_refdoc.Text = SELPATHDREQUISITN_RS("REFDOC")
       
            End If
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
            '*****
                If SELINDICATIONFLAG = False Then
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELINDICATIONRS = SELINDICATIONQR.OpenResultset(1, 3)
                    SELINDICATIONFLAG = True
                Else
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    SELINDICATIONRS.MoreResults
                    SELINDICATIONRS.Requery
                End If
                If SELINDICATIONRS.RowCount > 0 Then
                    If txt_modcode.Text = "PF" Then
                        
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND7"))
                        End If
                        msk_pfexamdt.Text = Format(DATE11, "dd/mm/yyyy")
                    End If
                    If txt_modcode.Text = "ST" Then
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                        msk_stexamdt.Text = Format(DATE11, "dd/mm/yyyy")
                    End If
                    
                End If
            '*****
            If Trim(txt_modcode.Text) = "EK" Then
                ECENABL ("ENA")
                msk_ecexamdt.Text = Format(DATE11, "dd/mm/yyyy")
                If seldissecprovflag1 = False Then
                    Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
                    seldissecprovflag1 = True
                Else
                    seldissecprov1_rs.MoreResults
                    seldissecprov1_rs.Requery
                End If
                If seldissecprov1_rs.RowCount > 0 Then
                    cmb_edproperby.clear
                   ' cmb_provisional.clear
                    While Not seldissecprov1_rs.EOF
                        cmb_edproperby.AddItem Trim(seldissecprov1_rs("CCNAME"))
                       ' cmb_provisional.AddItem Trim(seldissecprov1_rs("CCNAME"))
                        seldissecprov1_rs.MoveNext
                    Wend
                End If
                
                
            End If
            If Trim(txt_modcode.Text) = "ST" Then
                STENABL ("ENA")
                If seldissecprovflag1 = False Then
                    Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
                    seldissecprovflag1 = True
                Else
                    seldissecprov1_rs.MoreResults
                    seldissecprov1_rs.Requery
                End If
                If seldissecprov1_rs.RowCount > 0 Then
                    cmb_stproperformed.clear
                    cmb_provisional.clear
                    While Not seldissecprov1_rs.EOF
                        cmb_stproperformed.AddItem Trim(seldissecprov1_rs("CCNAME"))
                        cmb_provisional.AddItem Trim(seldissecprov1_rs("CCNAME"))
                        seldissecprov1_rs.MoveNext
                    Wend
                End If
            End If
            If Trim(txt_modcode.Text) = "PF" Then
                If seldissecprovflag1 = False Then
                    Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
                    seldissecprovflag1 = True
                Else
                    seldissecprov1_rs.MoreResults
                    seldissecprov1_rs.Requery
                End If
                If seldissecprov1_rs.RowCount > 0 Then
                    cmb_pfproperby.clear
                    cmb_provisional.clear
                    While Not seldissecprov1_rs.EOF
                        cmb_pfproperby.AddItem Trim(seldissecprov1_rs("CCNAME"))
                        cmb_provisional.AddItem Trim(seldissecprov1_rs("CCNAME"))
                        seldissecprov1_rs.MoveNext
                    Wend
                End If
            End If
        If seldissecentrflag1 = False Then
            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
            seldissecentrflag1 = True
        Else
            seldissecentr1_rs.MoreResults
            seldissecentr1_rs.Requery
        End If
        If seldissecentr1_rs.RowCount > 0 Then
        
            'entr = cmb_enterby.Text
            'cmb_enterby.clear
            While Not seldissecentr1_rs.EOF
            '    cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
                seldissecentr1_rs.MoveNext
            Wend
            seldissecentr1_rs.MovePrevious
            'cmb_enterby.Text = seldissecentr1_rs("CCNAME")
            
        End If

        If seldissecfinflag = False Then
            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
            seldissecfinflag1 = True
        Else
            seldissecfin1_rs.MoreResults
            seldissecfin1_rs.Requery
        End If
        If seldissecfin1_rs.RowCount > 0 Then
            entr = ""
            entr = cmb_finalby.Text
            'cmb_finalby.clear
            While Not seldissecfin1_rs.EOF
                'cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
                seldissecfin1_rs.MoveNext
            Wend
            cmb_finalby.Text = entr
            entr = ""
        End If
            
            
            
End If
If viewflag = True Then
     
            viewflag = False
            If selta101flag1 = False Then
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101flag1 = True
            Else
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
            End If
            If SELTA101_RS.RowCount = 0 Then
                Exit Sub
            End If
        
            txt_sex.Text = SELTA101_RS("sex")
            txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
            txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            txt_pridoc.Text = SELTA101_RS("ud_name")
            txt_category.Text = SELTA101_RS("pat_catg")
            
            If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Report is commited , You can not Modify"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                cmb_provisional.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
                End If
                
            End If
            
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_status.Text = "OP"
            End If
          
         
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF = False Then
                    'txt_relclnhis.Visible = True
                    
                    'txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
            End If
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If

'End If
        '*****
                If SELINDICATIONFLAG = False Then
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELINDICATIONRS = SELINDICATIONQR.OpenResultset(1, 3)
                    SELINDICATIONFLAG = True
                Else
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    SELINDICATIONRS.MoreResults
                    SELINDICATIONRS.Requery
                End If
                If SELINDICATIONRS.RowCount > 0 Then
                    If txt_modcode.Text = "PF" Then
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND7"))
                        End If
                    End If
                    If txt_modcode.Text = "ST" Then
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                    End If
                End If
            '*****
                If Trim(txt_modcode.Text) = "PF" Then
                    If SELPFREPORTFLAG = False Then
                        SELPFREPORTQR(0) = Trim(req_chglen(cmb_reqno.Text))
                        Set SELPFREPORTRS = SELPFREPORTQR.OpenResultset(1, 3)
                        SELPFREPORTFLAG = True
                    Else
                        SELPFREPORTQR(0) = Trim(req_chglen(cmb_reqno.Text))
                        SELPFREPORTRS.MoreResults
                        SELPFREPORTRS.Requery
                    End If
                    If SELPFREPORTRS.RowCount > 0 Then
                        If SELPFREPORTRS.EOF = False Then
                            TXT_PFTREPORT.Text = "" & SELPFREPORTRS("REPORT")
                        End If
                        If SELPFREPORTRS.EOF = False Then
                            TXT_PFTIMPRESSION.Text = "" & SELPFREPORTRS("IMPRESSION")
                        End If
                        msk_pfexamdt.Text = Format(SELPFREPORTRS("examdt"), "dd/mm/yyyy")
                        cmb_pfproperby.Text = SELPFREPORTRS("properby")
                    End If
                    TXT_PFTREPORT.Enabled = False
                    TXT_PFTIMPRESSION.Enabled = False
                    msk_ecexamdt.Enabled = False
                
                End If
                If Trim(txt_modcode.Text) = "ST" Then
                    STENABL ("DIS")
                    If SELSTREPORTFLAG = False Then
                        SELSTREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELSTREPORTRS = SELSTREPORTQR.OpenResultset(1, 3)
                        SELSTREPORTFLAG = True
                    Else
                        SELSTREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        SELSTREPORTRS.MoreResults
                        SELSTREPORTRS.Requery
                    End If
                    If SELSTREPORTRS.RowCount > 0 Then
                        CMB_RISKFACTOR.Text = Trim(SELSTREPORTRS("RISKFACTOR"))
                        TXT_MAXHR.Text = SELSTREPORTRS("MAXHR")
                        TXT_MAXBP.Text = SELSTREPORTRS("MAXBP")
                        TXT_DOUBLEPRODUCT.Text = SELSTREPORTRS("DOBPRODUCT")
                        TXT_THRACHIEVED.Text = SELSTREPORTRS("THRACHIVED")
                        If SELSTREPORTRS.EOF = False Then
                            TXT_STTCHANGES.Text = " " & SELSTREPORTRS("STTCHANGES")
                        End If
                        If SELSTREPORTRS.EOF = False Then
                            TXT_ARRYTHMIA.Text = " " & SELSTREPORTRS("ARRYTHMIA")
                        End If
                        CMB_RESFORTREAT.Text = SELSTREPORTRS("RESTERMI")
                        TXT_METSA.Text = SELSTREPORTRS("METSACHVED")
                        If SELSTREPORTRS.EOF = False Then
                        
                            TXT_STIMPRESSION.Text = " " & SELSTREPORTRS("IMPRESSION")
                        End If
                        If SELSTREPORTRS.EOF = False Then
                            cmb_stproperformed.Text = " " & SELSTREPORTRS("PROPERBY")
                        End If
                        msk_stexamdt.Text = Format(SELSTREPORTRS("fld1"), "dd/mm/yyyy")
                    End If
                End If
                If Trim(txt_modcode.Text) = "EK" Then
                    ECENABL ("DIS")
                    If SELECREPORTFLAG = False Then
                        SELECREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELECREPORTRS = SELECREPORTQR.OpenResultset(1, 3)
                        SELECREPORTFLAG = True
                    Else
                        SELECREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        SELECREPORTRS.MoreResults
                        SELECREPORTRS.Requery
                    End If
                    If SELECREPORTRS.RowCount > 0 Then
                        TXT_RATE.Text = Trim(SELECREPORTRS("RATE"))
                        TXT_ATRIALRATE.Text = Trim(SELECREPORTRS("ATRIALRATE"))
                        TXT_VENRATE.Text = Trim(SELECREPORTRS("VENRATE"))
                        If SELECREPORTRS.EOF = False Then
                            TXT_PWAVE.Text = " " & Trim(SELECREPORTRS("PWAVE"))
                       
                        End If
                        TXT_PRINTERVAL.Text = Trim(SELECREPORTRS("PRINTERVAL"))
                        If SELECREPORTRS.EOF = False Then
                            TXT_RHYTHM.Text = " " & Trim(SELECREPORTRS("RHYTHM"))
                        End If
                        TXT_QRSM.Text = Trim(SELECREPORTRS("QRSMORPH"))
                        TXT_QRSD.Text = Trim(SELECREPORTRS("QRSDUR"))
                        TXT_QTINTERVAL.Text = Trim(SELECREPORTRS("QTINTERVAL"))
                        If SELECREPORTRS.EOF = False Then
                            TXT_STSEGMENT.Text = " " & Trim(SELECREPORTRS("STSEGMENT"))
                            
                        End If
                        If SELECREPORTRS.EOF = False Then
                            TXT_TWAVE.Text = " " & Trim(SELECREPORTRS("TWAVE"))
                        End If
                        If SELECREPORTRS.EOF = False Then
                            TXT_ECIMPRESSION.Text = " " & Trim(SELECREPORTRS("IMPRESSION"))
                        End If
                        msk_ecexamdt.Text = Format(SELECREPORTRS("fld1"), "dd/mm/yyyy")
                        cmb_edproperby.Text = Trim(SELECREPORTRS("properby"))
                    End If
                End If
        
End If
End If
If opt_staff.Value = True Then
    STAFF
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub



Private Sub CMB_RESFORTREAT_GotFocus()
CMB_RESFORTREAT.clear
CMB_RESFORTREAT.AddItem "Fatigue"
CMB_RESFORTREAT.AddItem "Chest pain"
CMB_RESFORTREAT.AddItem "Dyspnea"

End Sub



Private Sub CMB_RISKFACTOR_GotFocus()
CMB_RISKFACTOR.clear
CMB_RISKFACTOR.AddItem "HT"
CMB_RISKFACTOR.AddItem "DM"
CMB_RISKFACTOR.AddItem "Smoker"
CMB_RISKFACTOR.AddItem "F/h/o IHD"
CMB_RISKFACTOR.AddItem "Hyperlipidemia"
End Sub

Private Sub cmb_stproperformed_LostFocus()
cmb_provisional.Text = cmb_stproperformed.Text
End Sub

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
btnenabl ("INI")
field_clear
FRM_ECG.Visible = False
FRM_PFT.Visible = False
FRM_STRESSTEST.Visible = False
ECCLEAR
STCLEAR
viewflag = False
NEWFLAG = False


End Sub

Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    If USRNAME_RS("SAVE") = "YES" Then
        cmd_save.Enabled = True
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
    If USRNAME_RS("MODIFY") = "YES" Then
        cmd_modify.Enabled = True
    Else
        cmd_modify.Enabled = False
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
    If USRNAME_RS("NEW") = "YES" Then
        CMD_NEW.Enabled = True
    Else
        CMD_NEW.Enabled = False
    End If
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    If USRNAME_RS("COMMIT") = "YES" Then
        cmd_commit.Enabled = True
    Else
        cmd_commit.Enabled = False
    End If
    If USRNAME_RS("VIEW") = "YES" Then
        cmd_view.Enabled = True
    Else
        cmd_view.Enabled = False
    End If
    'cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    If USRNAME_RS("NEW") = "YES" Then
        CMD_NEW.Enabled = True
    Else
        CMD_NEW.Enabled = False
    End If
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    If USRNAME_RS("VIEW") = "YES" Then
        cmd_view.Enabled = True
    Else
        cmd_view.Enabled = False
    End If
    'cmd_fullreport.Enabled = False
ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    If USRNAME_RS("NEW") = "YES" Then
        CMD_NEW.Enabled = True
    Else
        CMD_NEW.Enabled = False
    End If
    cmd_modify.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    If USRNAME_RS("VIEW") = "YES" Then
        cmd_view.Enabled = True
    Else
        cmd_view.Enabled = False
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
'
'ElseIf cap = "VIEW" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'
'ElseIf cap = "SAVE" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = True
'    cmd_view.Enabled = True
'
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'
'ElseIf cap = "COMMIT" Then
'    cmd_save.Enabled = False
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'
'ElseIf cap = "print" Then
'     cmd_save.Enabled = False
'     cmd_back.Enabled = True
'     cmd_modify.Enabled = False
'     cmd_print.Enabled = False
'     cmd_cancel.Enabled = True
'     CMD_NEW.Enabled = False
'
'End If
End Sub

Private Sub cmd_commit_Click()
Dim I As Integer
If Trim(txt_modcode.Text) = "PF" Then
    If Len(Trim(TXT_PFTIMPRESSION.Text)) = 0 Then
        MsgBox "Impression Field is Empty, You can not Commit the Report.", vbCritical
        TXT_PFTIMPRESSION.SetFocus
        Exit Sub
    End If
End If
I = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If I = vbYes Then
On Error GoTo MSG:
 rdb.BeginTrans
 
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


    updrrequisitn1_qr(0) = UCase(cmb_finalby.Text)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    MsgBox "Record is Committed Successfully."
End If
rdb.CommitTrans
btnenabl ("INI")

cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_modify_Click()
Dim str As String
Dim D1 As String
Dim D2 As String
flagsave = 2
btnenabl ("NEW")
    FLDENABL ("MODIFY")
If REFDRFLAG = False Then
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = REQNO
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    REFDRFLAG = True
Else
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = REQNO
    REFDR_RS.MoreResults
    REFDR_RS.Requery
End If

'      If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'        STR = ""
'            STR = cmb_enterby.Text
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'            cmb_enterby.Text = STR
'        End If

'        If seldissecfinflag = False Then
'            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'            seldissecfinflag1 = True
'        Else
'            seldissecfin1_rs.MoreResults
'            seldissecfin1_rs.Requery
'        End If
'        If seldissecfin1_rs.RowCount > 0 Then
'            STR = ""
'            STR = cmb_finalby.Text
'            cmb_finalby.clear
'            While Not seldissecfin1_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                seldissecfin1_rs.MoveNext
'            Wend
'            cmb_finalby.Text = STR
'        End If
cmb_enterby.Enabled = True
cmb_finalby.Enabled = True
If Trim(txt_modcode.Text) = "PF" Then
    TXT_PFTREPORT.Enabled = True
    TXT_PFTIMPRESSION.Enabled = True
    msk_pfexamdt.Enabled = True
    
                D1 = cmb_pfproperby.Text
                D2 = cmb_provisional.Text
'                If seldissecprovflag1 = False Then
'                    Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
'                    seldissecprovflag1 = True
'                Else
'                    seldissecprov1_rs.MoreResults
'                    seldissecprov1_rs.Requery
'                End If
'                If seldissecprov1_rs.RowCount > 0 Then
'                    cmb_pfproperby.clear
'                    cmb_provisional.clear
'                    While Not seldissecprov1_rs.EOF
'                        cmb_pfproperby.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                        cmb_provisional.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                        seldissecprov1_rs.MoveNext
'                    Wend
'
'                End If
                cmb_pfproperby.Text = D1
                cmb_provisional.Text = D2
    
    
    
End If
If Trim(txt_modcode.Text) = "ST" Then
    STENABL ("ENA")
    CMB_RESFORTREAT.clear
    CMB_RESFORTREAT.AddItem "Fatigue"
    CMB_RESFORTREAT.AddItem "Chest pain"
    CMB_RESFORTREAT.AddItem "Dyspnea"
    
    CMB_RISKFACTOR.clear
    CMB_RISKFACTOR.AddItem "HT"
    CMB_RISKFACTOR.AddItem "DM"
    CMB_RISKFACTOR.AddItem "Smoker"
    CMB_RISKFACTOR.AddItem "F/h/o IHD"
    CMB_RISKFACTOR.AddItem "Hyperlipidemia"
    
                D1 = cmb_stproperformed.Text
                D2 = cmb_provisional.Text
'                If seldissecprovflag1 = False Then
'                    Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
'                    seldissecprovflag1 = True
'                Else
'                    seldissecprov1_rs.MoreResults
'                    seldissecprov1_rs.Requery
'                End If
'                If seldissecprov1_rs.RowCount > 0 Then
'                    cmb_stproperformed.clear
'                    cmb_provisional.clear
'                    While Not seldissecprov1_rs.EOF
'                        cmb_stproperformed.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                        cmb_provisional.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                        seldissecprov1_rs.MoveNext
'                    Wend
'
'                End If
                cmb_stproperformed.Text = D1
                cmb_provisional.Text = D2
    
    
     
    
End If
If Trim(txt_modcode.Text) = "EK" Then
    ECENABL ("ENA")
                D1 = cmb_edproperby.Text
                D2 = cmb_provisional.Text
'                If seldissecprovflag1 = False Then
'                    Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
'                    seldissecprovflag1 = True
'                Else
'                    seldissecprov1_rs.MoreResults
'                    seldissecprov1_rs.Requery
'                End If
'                If seldissecprov1_rs.RowCount > 0 Then
'                    cmb_edproperby.clear
'                    cmb_provisional.clear
'                    While Not seldissecprov1_rs.EOF
'                        cmb_edproperby.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                        cmb_provisional.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                        seldissecprov1_rs.MoveNext
'                    Wend
'
'                End If
                cmb_edproperby.Text = D1
              cmb_provisional.Text = D2
End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG:
btnenabl ("NEW")
field_clear
TXT_CLEAR


If DATETIMEFLAG = False Then
   DATETIME
   DATETIMEFLAG = True
Else
    DATETIME
End If
txt_date.Text = DATE11
txt_time.Text = time1

'cmb_finalby.Text = cmb_finalby.List(0)
cmb_provisional.Text = cmb_provisional.List(0)
cmb_pfproperby.Text = cmb_pfproperby.List(0)
If SELDEPTRSFLAG1 = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SELDEPTRSFLAG1 = True
Else
    sel_deptrs.MoreResults
    sel_deptrs.Requery
End If
If sel_deptrs.RowCount > 0 Then
    cmb_dept.clear
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
        Wend
   End If
Opt_patient.Value = True

flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
MaskEdBox1.Enabled = True
ECCLEAR
STCLEAR
cmb_enterby.Text = cmb_enterby.List(0)

For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_finalby.List(L1)) = Trim(cmb_enterby.List(0)) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
Next L1

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_cancel_Click()
btnenabl ("INI")
field_clear
viewflag = False
NEWFLAG = False
FRM_ECG.Visible = False
FRM_PFT.Visible = False
FRM_STRESSTEST.Visible = False
ECCLEAR
STCLEAR

End Sub

Public Sub field_clear()
Dim obj As Object
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
txt_deptcode.Text = " "
cmb_reqno.ListIndex = -1
cmb_reqno.clear
If Opt_patient.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = ">??/#####"
End If
If opt_staff.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = "#######"         ' COMMENTED BY VARINDER
End If
txt_name.Text = " "
txt_sex.Text = " "
txt_age.Text = " "
txt_category.Text = " "
txt_status.Text = " "
txt_pridoc.Text = " "
txt_ward.Text = " "
txt_bedno.Text = " "
txt_refdoc.Text = " "
txt_procedure.Text = " "
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In FRM_PFTTEST.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub


Public Sub FLDENABL(EN As String)
If EN = "DIS" Then
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
        cmb_reqno.Enabled = False
        TXT_PATH.Enabled = False
        txt_name.Enabled = False
        txt_sex.Enabled = False
        txt_age.Enabled = False
        txt_category.Enabled = False
        txt_status.Enabled = False
        txt_pridoc.Enabled = False
        txt_ward.Enabled = False
        txt_bedno.Enabled = False
        txt_modcode.Enabled = False
        txt_refdoc.Enabled = False
        txt_procedure.Enabled = False
        'txt_observed.Visible = False
        'TXT_COMMENTS.Enabled = False
        cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
End If
If EN = "ENA" Then
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        TXT_PATH.Enabled = False
        txt_name.Enabled = False
        txt_sex.Enabled = False
        txt_age.Enabled = False
        txt_category.Enabled = False
        txt_status.Enabled = False
        txt_pridoc.Enabled = False
        txt_ward.Enabled = False
        txt_bedno.Enabled = False
        txt_modcode.Enabled = False
        txt_refdoc.Enabled = False
        txt_procedure.Enabled = False
        'txt_observed.Visible = False
        'TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
End If
If EN = "MOD" Then
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        TXT_PATH.Enabled = True
        txt_name.Enabled = True
        txt_sex.Enabled = True
        txt_age.Enabled = True
        txt_category.Enabled = True
        txt_status.Enabled = True
        txt_pridoc.Enabled = True
        txt_ward.Enabled = True
        txt_bedno.Enabled = True
        txt_modcode.Enabled = True
        txt_refdoc.Enabled = True
        txt_procedure.Enabled = True
        'txt_observed.Visible = True
        'TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
End If
End Sub



Private Sub cmd_print_Click()
On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
 On Error GoTo prtmsg:

DoEvents
            For I = 0 To 23
                CrystalReport1.Formulas(I) = ""
            Next I
Dim VAR_REQNNO As String


        VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If


            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
'            If SELRREQUISITN_Rs.RowCount > 0 Then
'                prepath = ""
'                While Not SELRREQUISITN_Rs.EOF
'                    If Trim(SELRREQUISITN_Rs("reqnno")) <> req_chglen(Trim(cmb_reqno.Text)) Then
'                        prepath = prepath & req_chglen(SELRREQUISITN_Rs("REQNNO")) & " "
'                    End If
'
'                    SELRREQUISITN_Rs.MoveNext
'                Wend
'             End If
'
If Opt_patient.Value = True Then
             
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

     If Trim(txt_modcode.Text) = "PF" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISPF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
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
        
     If (txt_modcode.Text) = "ST" Then
        
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISST.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            If Len(Trim(cmb_enterby.Text)) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        If Len(Trim(CMB_RISKFACTOR.Text)) > 0 Then
            CrystalReport1.Formulas(13) = "rk = '" & "Risk Factor" & "'"
        Else
            CrystalReport1.Formulas(13) = "rk = '" & "" & "'"
        End If
        If Len(Trim(TXT_MAXHR.Text)) > 0 Then
            CrystalReport1.Formulas(14) = "maxhr = '" & "Max HR" & "'"
        Else
            CrystalReport1.Formulas(14) = "maxhr = '" & "" & "'"
        End If
        If Len(Trim(TXT_MAXBP.Text)) > 0 Then
            CrystalReport1.Formulas(15) = "maxBP = '" & "Max BP" & "'"
        Else
            CrystalReport1.Formulas(15) = "maxBP = '" & "" & "'"
        End If
        If Len(Trim(TXT_DOUBLEPRODUCT.Text)) > 0 Then
            CrystalReport1.Formulas(16) = "DP = '" & "Double Product" & "'"
        Else
            CrystalReport1.Formulas(16) = "DP = '" & "" & "'"
        End If
        If Len(Trim(TXT_THRACHIEVED.Text)) > 0 Then
            CrystalReport1.Formulas(17) = "thra = '" & "% of THR Achieved" & "'"
        Else
            CrystalReport1.Formulas(17) = "thra = '" & "" & "'"
        End If
        If Len(Trim(TXT_STTCHANGES.Text)) > 0 Then
            CrystalReport1.Formulas(18) = "sttc = '" & "ST/T Changes" & "'"
        Else
            CrystalReport1.Formulas(18) = "sttc = '" & "" & "'"
        End If
        If Len(Trim(TXT_ARRYTHMIA.Text)) > 0 Then
            CrystalReport1.Formulas(19) = "arrythmia = '" & "Arrythmia" & "'"
        Else
            CrystalReport1.Formulas(19) = "arrythmia = '" & "" & "'"
        End If
        If Len(Trim(TXT_METSA.Text)) > 0 Then
            CrystalReport1.Formulas(20) = "metsa = '" & "METS Achieved" & "'"
        Else
            CrystalReport1.Formulas(20) = "metsa = '" & "" & "'"
        End If
        If Len(Trim(CMB_RESFORTREAT.Text)) > 0 Then
            CrystalReport1.Formulas(21) = "rt = '" & "Reason for termination" & "'"
        Else
            CrystalReport1.Formulas(21) = "rt = '" & "" & "'"
        End If
     End If
     If Trim(txt_modcode.Text) = "EK" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            If Len(Trim(cmb_enterby.Text)) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If Len(Trim(TXT_RHYTHM)) > 3 Then
            CrystalReport1.Formulas(13) = "rhythm = '" & "Rhythm" & "'"
        Else
            CrystalReport1.Formulas(13) = "rhythm = '" & "" & "'"
        End If
        
          If (Trim(TXT_PWAVE.Text)) <> "" Then
            CrystalReport1.Formulas(14) = "pwave = '" & "Pwave" & "'"
          Else
            CrystalReport1.Formulas(14) = "pwave = '" & "" & "'"
        End If
        If Len(Trim(TXT_TWAVE)) > 3 Then
            CrystalReport1.Formulas(15) = "twave = '" & "Twave" & "'"
        Else
            CrystalReport1.Formulas(15) = "twave = '" & "" & "'"
        End If
        If Len(Trim(TXT_STSEGMENT)) > 3 Then
            
             CrystalReport1.Formulas(16) = "stsegment = '" & "ST Segment" & "'"
        Else
             CrystalReport1.Formulas(16) = "stsegment = '" & "" & "'"
        End If
        If Len(Trim(TXT_RATE)) > 0 Then
            
             CrystalReport1.Formulas(17) = "rate = '" & "Rate" & "'"
        Else
             CrystalReport1.Formulas(17) = "rate = '" & "" & "'"
        End If
        If Len(Trim(TXT_ATRIALRATE)) > 0 Then
            
             CrystalReport1.Formulas(18) = "arate = '" & "Atrial Rate" & "'"
        Else
             CrystalReport1.Formulas(18) = "arate = '" & "" & "'"
        End If
        If Len(Trim(TXT_VENRATE)) > 0 Then
            
             CrystalReport1.Formulas(19) = "vrate = '" & "Ventricular Rate" & "'"
        Else
             CrystalReport1.Formulas(19) = "vrate = '" & "" & "'"
        End If
        If Len(Trim(TXT_PRINTERVAL)) > 0 Then
            
             CrystalReport1.Formulas(20) = "printerval = '" & "PR Interval" & "'"
        Else
             CrystalReport1.Formulas(20) = "printerval = '" & "" & "'"
        End If
        If Len(Trim(TXT_QRSM)) > 0 Then
            
             CrystalReport1.Formulas(21) = "qrsm = '" & "QRS Morphology" & "'"
        Else
             CrystalReport1.Formulas(21) = "qrsm = '" & "" & "'"
        End If
        If Len(Trim(TXT_QRSD)) > 0 Then
            
             CrystalReport1.Formulas(22) = "qrd = '" & "QRS Duration" & "'"
        Else
             CrystalReport1.Formulas(22) = "qrd = '" & "" & "'"
        End If
        If Len(Trim(TXT_QTINTERVAL)) > 0 Then
            
             CrystalReport1.Formulas(23) = "qti = '" & "QT Interval" & "'"
        Else
             CrystalReport1.Formulas(23) = "qti = '" & "" & "'"
        End If
     End If
       
End If
'**** STAFF
If opt_staff.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

     If Trim(txt_modcode.Text) = "PF" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISPFSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
            CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(3) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(4) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(3) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(4) = "ENT = '" & str1 & "'"
        End If
            CrystalReport1.Formulas(5) = "USRID = '" & str & "'"
            CrystalReport1.Formulas(6) = "NAM = '" & Trim(txt_name.Text) & "'"
            CrystalReport1.Formulas(7) = "DOB = '" & Trim(txt_age.Text) & "'"
            CrystalReport1.Formulas(8) = "STDOB = '" & Trim(txt_sex.Text) & "'"
     End If
        
     If (txt_modcode.Text) = "ST" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISSTSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
            CrystalReport1.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(2) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(2) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(3) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(4) = "ENT = '" & str1 & "'"
        End If
            CrystalReport1.Formulas(5) = "USRID = '" & str & "'"
            CrystalReport1.Formulas(6) = "NAM = '" & Trim(txt_name.Text) & "'"
            CrystalReport1.Formulas(7) = "DOB = '" & Trim(txt_age.Text) & "'"
            CrystalReport1.Formulas(8) = "STDOB = '" & Trim(txt_sex.Text) & "'"
        '****************************
        If Len(Trim(CMB_RISKFACTOR.Text)) > 0 Then
            CrystalReport1.Formulas(13) = "rk = '" & "Risk Factor" & "'"
        Else
            CrystalReport1.Formulas(13) = "rk = '" & "" & "'"
        End If
        
        If Len(Trim(TXT_MAXHR.Text)) > 0 Then
            CrystalReport1.Formulas(14) = "maxhr = '" & "Max HR" & "'"
        Else
            CrystalReport1.Formulas(14) = "maxhr = '" & "" & "'"
        End If
        
        If Len(Trim(TXT_MAXBP.Text)) > 0 Then
            CrystalReport1.Formulas(15) = "maxBP = '" & "Max BP" & "'"
        Else
            CrystalReport1.Formulas(15) = "maxBP = '" & "" & "'"
        End If
        
        If Len(Trim(TXT_DOUBLEPRODUCT.Text)) > 0 Then
            CrystalReport1.Formulas(16) = "DP = '" & "Double Product" & "'"
        Else
            CrystalReport1.Formulas(16) = "DP = '" & "" & "'"
        End If
        
        If Len(Trim(TXT_THRACHIEVED.Text)) > 0 Then
            CrystalReport1.Formulas(17) = "thra = '" & "% of THR Achieved" & "'"
        Else
            CrystalReport1.Formulas(17) = "thra = '" & "" & "'"
        End If
        
        If Len(Trim(TXT_STTCHANGES.Text)) > 0 Then
            CrystalReport1.Formulas(18) = "sttc = '" & "ST/T Changes" & "'"
        Else
            CrystalReport1.Formulas(18) = "sttc = '" & "" & "'"
        End If
        
        If Len(Trim(TXT_ARRYTHMIA.Text)) > 0 Then
            CrystalReport1.Formulas(19) = "arrythmia = '" & "Arrythmia" & "'"
        Else
            CrystalReport1.Formulas(19) = "arrythmia = '" & "" & "'"
        End If
        
        If Len(Trim(TXT_METSA.Text)) > 0 Then
            CrystalReport1.Formulas(20) = "metsa = '" & "METS Achieved" & "'"
        Else
            CrystalReport1.Formulas(20) = "metsa = '" & "" & "'"
        End If
        
        If Len(Trim(CMB_RESFORTREAT.Text)) > 0 Then
            CrystalReport1.Formulas(21) = "rt = '" & "Reason for termination" & "'"
        Else
            CrystalReport1.Formulas(21) = "rt = '" & "" & "'"
        End If
        
        '****************************
     End If
     
     If Trim(txt_modcode.Text) = "EK" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISECSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
            CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
'''''        If Len(Trim(TXT_RHYTHM)) > 0 Then
'''''            CRYSTALREPORT1.Formulas(13) = "rhythm = '" & "" & "'"
'''''        Else
'''''            CRYSTALREPORT1.Formulas(13) = "rhythm = '" & "Rhythm" & "'"
'''''        End If
'''''
'''''        If Len(Trim(TXT_PWAVE)) > 0 Then
'''''            CRYSTALREPORT1.Formulas(14) = "pwave = '" & "" & "'"
'''''        Else
'''''            CRYSTALREPORT1.Formulas(14) = "pwave = '" & "Pwave" & "'"
'''''        End If
'''''        If Len(Trim(TXT_TWAVE)) > 0 Then
'''''            CRYSTALREPORT1.Formulas(14) = "twave = '" & "" & "'"
'''''        Else
'''''            CRYSTALREPORT1.Formulas(14) = "twave = '" & "Twave" & "'"
'''''        End If
'''''        If Len(Trim(TXT_STSEGMENT)) > 0 Then
'''''            CRYSTALREPORT1.Formulas(15) = "stsegment = '" & "" & "'"
'''''        Else
'''''            CRYSTALREPORT1.Formulas(15) = "stsegment = '" & "ST Segment" & "'"
'''''        End If
        If Len(Trim(TXT_RHYTHM)) > 3 Then
            CrystalReport1.Formulas(13) = "rhythm = '" & "Rhythm" & "'"
        Else
            CrystalReport1.Formulas(13) = "rhythm = '" & "" & "'"
        End If
        
          If (Trim(TXT_PWAVE.Text)) <> "" Then
            CrystalReport1.Formulas(14) = "pwave = '" & "Pwave" & "'"
          Else
            CrystalReport1.Formulas(14) = "pwave = '" & "" & "'"
        End If
        If Len(Trim(TXT_TWAVE)) > 3 Then
            CrystalReport1.Formulas(15) = "twave = '" & "Twave" & "'"
        Else
            CrystalReport1.Formulas(15) = "twave = '" & "" & "'"
        End If
        If Len(Trim(TXT_STSEGMENT)) > 3 Then
            
             CrystalReport1.Formulas(16) = "stsegment = '" & "ST Segment" & "'"
        Else
             CrystalReport1.Formulas(16) = "stsegment = '" & "" & "'"
        End If
        If Len(Trim(TXT_RATE)) > 0 Then
            
             CrystalReport1.Formulas(17) = "rate = '" & "Rate" & "'"
        Else
             CrystalReport1.Formulas(17) = "rate = '" & "" & "'"
        End If
        If Len(Trim(TXT_ATRIALRATE)) > 0 Then
            
             CrystalReport1.Formulas(18) = "arate = '" & "Atrial Rate" & "'"
        Else
             CrystalReport1.Formulas(18) = "arate = '" & "" & "'"
        End If
        If Len(Trim(TXT_VENRATE)) > 0 Then
            
             CrystalReport1.Formulas(19) = "vrate = '" & "Ventricular Rate" & "'"
        Else
             CrystalReport1.Formulas(19) = "vrate = '" & "" & "'"
        End If
        If Len(Trim(TXT_PRINTERVAL)) > 0 Then
            
             CrystalReport1.Formulas(20) = "printerval = '" & "PR Interval" & "'"
        Else
             CrystalReport1.Formulas(20) = "printerval = '" & "" & "'"
        End If
        If Len(Trim(TXT_QRSM)) > 0 Then
            
             CrystalReport1.Formulas(21) = "qrsm = '" & "QRS Morphology" & "'"
        Else
             CrystalReport1.Formulas(21) = "qrsm = '" & "" & "'"
        End If
        If Len(Trim(TXT_QRSD)) > 0 Then
            
             CrystalReport1.Formulas(22) = "qrd = '" & "QRS Duration" & "'"
        Else
             CrystalReport1.Formulas(22) = "qrd = '" & "" & "'"
        End If
        If Len(Trim(TXT_QTINTERVAL)) > 0 Then
            
             CrystalReport1.Formulas(23) = "qti = '" & "QT Interval" & "'"
        Else
             CrystalReport1.Formulas(23) = "qti = '" & "" & "'"
        End If
    End If

End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_save_Click()
Dim pcount As Integer
    pcount = 0
Dim beforst As String, afterst As String
Dim K As Long
If Trim(txt_modcode.Text) = "EK" Then
        While InStr(TXT_PWAVE, Chr(9)) > 0
            
            K = InStr(TXT_PWAVE, Chr(9))
            beforst = Mid(TXT_PWAVE, 1, K - 1) & " "
            afterst = Mid(TXT_PWAVE, K + 1)
            TXT_PWAVE = beforst & afterst
        Wend
        While InStr(TXT_PWAVE, Chr(146)) > 0

            K = InStr(TXT_PWAVE, Chr(146))
            beforst = Mid(TXT_PWAVE, 1, K - 1) & "`"
            afterst = Mid(TXT_PWAVE, K + 1)
            TXT_PWAVE = beforst & afterst
       Wend
       
       While InStr(TXT_TWAVE, Chr(9)) > 0
            
            K = InStr(TXT_TWAVE, Chr(9))
            beforst = Mid(TXT_TWAVE, 1, K - 1) & " "
            afterst = Mid(TXT_TWAVE, K + 1)
            TXT_TWAVE = beforst & afterst
        Wend
        While InStr(TXT_TWAVE, Chr(146)) > 0

            K = InStr(TXT_TWAVE, Chr(146))
            beforst = Mid(TXT_TWAVE, 1, K - 1) & "`"
            afterst = Mid(TXT_TWAVE, K + 1)
            TXT_TWAVE = beforst & afterst
       Wend
        
       While InStr(TXT_STSEGMENT, Chr(9)) > 0
            
            K = InStr(TXT_STSEGMENT, Chr(9))
            beforst = Mid(TXT_STSEGMENT, 1, K - 1) & " "
            afterst = Mid(TXT_STSEGMENT, K + 1)
            TXT_STSEGMENT = beforst & afterst
        Wend
        While InStr(TXT_STSEGMENT, Chr(146)) > 0

            K = InStr(TXT_STSEGMENT, Chr(146))
            beforst = Mid(TXT_STSEGMENT, 1, K - 1) & "`"
            afterst = Mid(TXT_STSEGMENT, K + 1)
            TXT_STSEGMENT = beforst & afterst
       Wend
       
       While InStr(TXT_ECIMPRESSION, Chr(9)) > 0
            
            K = InStr(TXT_ECIMPRESSION, Chr(9))
            beforst = Mid(TXT_ECIMPRESSION, 1, K - 1) & " "
            afterst = Mid(TXT_ECIMPRESSION, K + 1)
            TXT_ECIMPRESSION = beforst & afterst
        Wend
        While InStr(TXT_ECIMPRESSION, Chr(146)) > 0

            K = InStr(TXT_ECIMPRESSION, Chr(146))
            beforst = Mid(TXT_ECIMPRESSION, 1, K - 1) & "`"
            afterst = Mid(TXT_ECIMPRESSION, K + 1)
            TXT_ECIMPRESSION = beforst & afterst
       Wend
       
       While InStr(TXT_RHYTHM, Chr(9)) > 0
            
            K = InStr(TXT_RHYTHM, Chr(9))
            beforst = Mid(TXT_RHYTHM, 1, K - 1) & " "
            afterst = Mid(TXT_RHYTHM, K + 1)
            TXT_RHYTHM = beforst & afterst
        Wend
        While InStr(TXT_RHYTHM, Chr(146)) > 0

            K = InStr(TXT_RHYTHM, Chr(146))
            beforst = Mid(TXT_RHYTHM, 1, K - 1) & "`"
            afterst = Mid(TXT_RHYTHM, K + 1)
            TXT_RHYTHM = beforst & afterst
       Wend
       
       If Len(Trim(TXT_PWAVE.Text)) > 200 Then
        MsgBox "Max.limit of Pwave is 200 characters."
        TXT_PWAVE.SetFocus
        Exit Sub
       End If
       
       If Len(Trim(TXT_TWAVE.Text)) > 160 Then
        MsgBox "Max.limit of Twave is 160 characters."
        TXT_TWAVE.SetFocus
        Exit Sub
       End If
       
       If Len(Trim(TXT_ECIMPRESSION.Text)) > 1000 Then
        MsgBox "Max.limit of Impression is 1000 characters."
        TXT_STIMPRESSION.SetFocus
        Exit Sub
       End If
       
       If Len(Trim(TXT_STSEGMENT.Text)) > 160 Then
        MsgBox "Max.limit of ST Segment is 160 characters."
        TXT_STSEGMENT.SetFocus
        Exit Sub
       End If
       
       If Len(Trim(TXT_RHYTHM.Text)) > 160 Then
        MsgBox "Max.limit of Rhythm is 160 characters."
        TXT_RHYTHM.SetFocus
        Exit Sub
       End If
       If Len(Trim(TXT_ECIMPRESSION.Text)) = 0 Then
        MsgBox "IMpression field is Empty.", vbCritical
        Exit Sub
       End If
       If Len(Trim(cmb_edproperby.Text)) = 0 Then
        MsgBox "Procedure Performed by  field is Empty.", vbCritical
        Exit Sub
    End If
End If

If Trim(txt_modcode.Text) = "ST" Then
    If Len(Trim(TXT_STIMPRESSION.Text)) = 0 Then
        MsgBox "Impression field is Empty.", vbCritical
        Exit Sub
    End If
    If Len(Trim(cmb_stproperformed.Text)) = 0 Then
        MsgBox "Procedure Performed by  field is Empty.", vbCritical
        Exit Sub
    End If
End If
If Trim(txt_modcode.Text) = "PF" Then
    If Len(Trim(TXT_PFTIMPRESSION.Text)) = 0 Then
        MsgBox "Impression field is Empty.", vbCritical
        Exit Sub
    End If
    If Len(Trim(cmb_pfproperby.Text)) = 0 Then
        MsgBox "Procedure Performed by  field is Empty.", vbCritical
        Exit Sub
    End If
End If

If Len(Trim(cmb_enterby.Text)) = 0 Then
    MsgBox "Enter by field is Empty.", vbCritical
    Exit Sub
End If
If Len(Trim(cmb_provisional.Text)) = 0 Then
    MsgBox "Reported by  field is Empty.", vbCritical
    Exit Sub
End If
If Len(Trim(cmb_finalby.Text)) = 0 Then
    MsgBox "To be finalised  field is Empty.", vbCritical
    Exit Sub
End If
If Trim(txt_modcode.Text) = "PF" Then
    PFASSIGN
End If
If Trim(txt_modcode.Text) = "ST" Then
    STASSIGN
End If
If Trim(txt_modcode.Text) = "EK" Then
   ECASSIGN
End If
btnenabl ("SAVE")
'cmd_print.Enabled = True
'cmd_cancel.Enabled = True


'commented by venkat
'If pcount = 0 Then
'    cmd_commit.Enabled = True
'Else
'    cmd_commit.Enabled = False
'End If
'cmd_commit.Enabled = True
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
FLDENABL ("DIS")
field_clear
TXT_CLEAR
If DATETIMEFLAG = False Then
   DATETIME
   DATETIMEFLAG = True
Else
    DATETIME
End If
txt_date.Text = DATE11
txt_time.Text = time1
If SEL_DEPTRSFLAG = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SEL_DEPTFLAG = True
Else
    sel_deptrs.MoreResults
    sel_deptrs.Requery
End If
cmb_dept.clear
If sel_deptrs.RowCount > 0 Then
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
        Wend
   End If
 
         If seldissecentrflag1 = False Then
            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
            seldissecentrflag1 = True
        Else
            seldissecentr1_rs.MoreResults
            seldissecentr1_rs.Requery
        End If
        If seldissecentr1_rs.RowCount > 0 Then
            cmb_enterby.clear
            While Not seldissecentr1_rs.EOF
                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
                seldissecentr1_rs.MoveNext
            Wend
        End If

        If seldissecfinflag = False Then
            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
            seldissecfinflag1 = True
        Else
            seldissecfin1_rs.MoreResults
            seldissecfin1_rs.Requery
        End If
        If seldissecfin1_rs.RowCount > 0 Then
            'cmb_finalby.clear
            While Not seldissecfin1_rs.EOF
                'cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
                seldissecfin1_rs.MoveNext
            Wend
        End If
 
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
Opt_patient.Value = True
cmb_dept.SetFocus
cmb_modality.Enabled = True
MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True

ECCLEAR
STCLEAR
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
''
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAMEPF1111", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.Text = LOGINUSRNAME
    cmb_enterby.AddItem LOGINUSRNAME
    'mb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Enabled = False
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'''
btnenabl ("INI")
NEWFLAG = False
viewflag = False
FLDENABL ("DIS")
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
If PFTESTFLAG = False Then
    PFTESTFLAG = True
    
  
'*********SECURITY**********************************

'''STR = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
'''Set USRNAME_QR = rdb.CreateQuery("DisUSRNAME1", STR)
'''USRNAME_QR(0) = UID
'''Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
'''If USRNAME_RS.RowCount > 0 Then
'''    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
'''    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
'''    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
'''    cmb_enterby.AddItem LOGINUSRNAME
'''    USRNAME_QR.Close
'''Else
'''    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
'''    cmd_cancel_Click
'''    Exit Sub
'''End If
    
'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNSTRPF", str)
REPWTRN_QR(0) = Trim(USRDEPT)
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)

If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    'cmb_addendumby.clear
    cmb_finalby.clear
    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 4 Then
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            
        End If
        If REPWTRNRS("USRTYPE") >= 3 Then
            cmb_provisional.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_pfproperby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        REPWTRNRS.MoveNext
        
    Wend
    cmb_finalby.Text = cmb_finalby.List(0)
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BIOPF", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

    
    
''for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr1PF", str)

'' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr1PF", str)

'    'for retriving reqnos
Set req_qr = rdb.CreateQuery("reqPF", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E'AND MODALITYCD=? ORDER BY REQNDT desc ")

'
Set SELTA101_QR = rdb.CreateQuery("selta101_qr11PF", "select * from PATOOLSLIB.ta101 where case_no =?")

'    ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1PF", str)

'    ' FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1PF", str)
''
'    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11PF", str)

'
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1PF", str)

    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1PF", str)

   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1PF", str)
'
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1PF", str)
'
'   STR = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
'   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1", STR)
'
'   STR = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
'   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11", STR)
'   SELDADDENDUM_QR(0) = "FSS"
'   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
   str = "select * from dislib.drequisitn where caseno=?"
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1BC", str)

    'commented and new query added by amarjeet on 05-11-2002 to select manjure's name in enter by
    str = "select * from dislib.dissec where enterby='YES'AND MODALITY='EC' "
 
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1PF", str)
   
   str = "select * from dislib.dissec where PROVby='YES'AND MODALITY='EC'"
   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1BC", str)
'
   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='EC'"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1PF", str)
    
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1PF", str)
    
  str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
  Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRPF", str)
    
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
  Set selMA101QR = rdb.CreateQuery("SELMA101QRPF", str)
  
'  STR = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET WHERE PRO_CODE=?"
'  Set SELDTEMPLETQR = rdb.CreateQuery("SELDTEMPLETQRPF", STR)
'
'  STR = "INSERT INTO DISLIB.DMPROCEDUR(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE) VALUES (?,?,?,?,?,?,?)"
'  Set INSDMPROCEDURQR = rdb.CreateQuery("INSDMPROCEDURQRPF", STR)
'
'  STR = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=? WHERE REQNO=? AND PRO_CODE=? "
'  Set UPDDMPROCEDURQR = rdb.CreateQuery("UPDDMPROCEDURQRPF", STR)
  
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_DET=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET=?"
  Set SELDMA101QR = rdb.CreateQuery("SELDMA101QRPF", str)
  
  str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1PF", str)
  
  
  str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrPF", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%EC%'"
  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRPF", str)
  
'  STR = "SELECT * FROM DISLIB.DBCALKPHS WHERE PROCODE=? AND AGE2=? AND SEX=? "
'  Set SELDBCALKPHSQR = rdb.CreateQuery("SELDBCALKPHSQRPF", STR)

  str = "SELECT * FROM DISLIB.INDICATION WHERE REQNO=?"
  Set SELINDICATIONQR = rdb.CreateQuery("SELINDICATIONQREC", str)

  '*** PFT
  str = "INSERT INTO DISLIB.DPFREPORT(REQNO,REPORT,IMPRESSION,examdt,properby) VALUES (?,?,?,?,?)"
  Set INSPFREPORTQR = rdb.CreateQuery("INSPFREPORTQRPF", str)
  
  str = "UPDATE DISLIB.DPFREPORT SET REPORT=?,IMPRESSION=?,examdt=?,properby=? WHERE REQNO=?"
  Set UPDPFREPORTQR = rdb.CreateQuery("UPDPFREPORTQRPF", str)
  
  str = "SELECT report,impression,examdt,properby FROM DISLIB.DPFREPORT WHERE REQNO=?"
  Set SELPFREPORTQR = rdb.CreateQuery("SELPFREPORTQRPF", str)
  '*** PFT
  '*** ST
  str = "INSERT INTO DISLIB.DSTREPORT(REQNO,RISKFACTOR,MAXHR,MAXBP,DOBPRODUCT,THRACHIVED,STTCHANGES,ARRYTHMIA,RESTERMI,METSACHVED,IMPRESSION,PROPERBY,fld1) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"
  Set INSSTREPORTQR = rdb.CreateQuery("INSSTREPORTQRPF", str)
  
  str = "UPDATE DISLIB.DSTREPORT SET RISKFACTOR=?,MAXHR=?,MAXBP=?,DOBPRODUCT=?,THRACHIVED=?,STTCHANGES=?,ARRYTHMIA=?,RESTERMI=?,METSACHVED=?,IMPRESSION=?,PROPERBY=?,fld1=? WHERE REQNO=?"
  Set UPDSTREPORTQR = rdb.CreateQuery("UPDSTREPORTQRST", str)
  
  str = "SELECT * FROM DISLIB.DSTREPORT WHERE REQNO=?"
  Set SELSTREPORTQR = rdb.CreateQuery("SELSTREPORTQRST", str)
  
  '*** ST
  '*** EC
  str = "INSERT INTO DISLIB.DECREPORT (REQNO,RATE,ATRIALRATE,VENRATE,PWAVE,PRINTERVAL,RHYTHM,QRSMORPH,QRSDUR,QTINTERVAL,STSEGMENT,TWAVE,QTCINTERVA,IMPRESSION,fld1,properby) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  Set INSECREPORTQR = rdb.CreateQuery("INSECREPORTQREC", str)
  
  str = "UPDATE DISLIB.DECREPORT SET RATE=?,ATRIALRATE=?,VENRATE=?,PWAVE=?,PRINTERVAL=?,RHYTHM=?,QRSMORPH=?,QRSDUR=?,QTINTERVAL=?,STSEGMENT=?,TWAVE=?,QTCINTERVA=?,IMPRESSION=?,fld1=?,properby=? WHERE REQNO=?"
  Set UPDECREPORTQR = rdb.CreateQuery("UPDECREPORTQREC", str)
  
  str = "SELECT RATE,ATRIALRATE,VENRATE,PWAVE,PRINTERVAL,RHYTHM,QRSMORPH,QRSDUR,QTINTERVAL,STSEGMENT,TWAVE,QTCINTERVA,IMPRESSION,fld1,properby FROM DISLIB.DECREPORT WHERE REQNO=?"
  Set SELECREPORTQR = rdb.CreateQuery("SELECREPORTQREC", str)
  
  '*** EC
  End If
  FRM_ECG.Visible = False
  FRM_STRESSTEST.Visible = False
  FRM_PFTTEST.Visible = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
       '' MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 05/12/2022
       '''' Commented PADZERO FUNCTION INSTEAD USED NEWNOPAD FUNCTION
'cno_pad = padzero(5, "C", msk_caseno.Text, 3)
    

    
  If IsNumeric(Left(Trim(MaskEdBox1.Text), 1)) Then
          Dim Ks As Integer
        Dim str1s As String
        Dim strs As String
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
         If NEWFLAG = True Then
            If REQFLAG1 = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                req_rs.MoreResults
                req_rs.Requery
          End If
          If req_rs.RowCount > 0 Then
                cmb_reqno.clear
                While Not req_rs.EOF
                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
                
                req_rs.MoveNext
                
                Wend
          Else
            cmb_reqno.clear
            
            End If
            
'          If SELDREQNOFLAG = False Then
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'            SELDREQNOFLAG = True
'          Else
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            SELDREQNORS.MoreResults
'            SELDREQNORS.Requery
'          End If
'          If SELDREQNORS.RowCount > 0 Then
'            CMB_PATHNO.clear
'            While Not SELDREQNORS.EOF
'                CMB_PATHNO.AddItem SELDREQNORS("PATHNO")
'                SELDREQNORS.MoveNext
'            Wend
'          End If
        End If
        
        If viewflag = True Then
           
            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
            If SELRREQUISITN_RS.RowCount > 0 Then
                cmb_reqno.clear
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
End If
    If opt_staff.Value = True Then
            CASENO = MaskEdBox1.Text
        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
         If NEWFLAG = True Then
            If REQFLAG1 = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                req_rs.MoreResults
                req_rs.Requery
          End If
          If req_rs.RowCount > 0 Then
                cmb_reqno.clear
                While Not req_rs.EOF
                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
                
                req_rs.MoveNext
                
                Wend
          Else
            cmb_reqno.clear
            
            End If
            
'          If SELDREQNOFLAG = False Then
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'            SELDREQNOFLAG = True
'          Else
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            SELDREQNORS.MoreResults
'            SELDREQNORS.Requery
'          End If
'          If SELDREQNORS.RowCount > 0 Then
'            CMB_PATHNO.clear
'            While Not SELDREQNORS.EOF
'                CMB_PATHNO.AddItem SELDREQNORS("PATHNO")
'                SELDREQNORS.MoveNext
'            Wend
'          End If
        End If
        
        If viewflag = True Then
            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
            If SELRREQUISITN_RS.RowCount > 0 Then
                cmb_reqno.clear
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
End If
 Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub




Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
   'MaskEdBox1.Mask = "??>/#####"
    MaskEdBox1.Enabled = True
    MaskEdBox1.SetFocus
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then

    lbl_caseno.Caption = "  CC No."
    MaskEdBox1.Mask = "########"
    MaskEdBox1.Enabled = True
    MaskEdBox1.SetFocus
End If
End Sub








Public Sub STAFF()
 If NEWFLAG = True Then
            NEWFLAG = False
            If SELDREQSTAFFD1FLAG = False Then
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
                Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
            Else
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
                SELDREQSTAFFDRS.MoreResults
                SELDREQSTAFFDRS.Requery
                
            End If
            If SELDREQSTAFFDRS.RowCount > 0 Then
                txt_sex.Text = SELDREQSTAFFDRS("sex")
                txt_name.Text = SELDREQSTAFFDRS("PNAME")
                txt_age = SELDREQSTAFFDRS("AGE")
                txt_pridoc.Text = ""
                txt_category.Text = SELDREQSTAFFDRS("RELATION")
            End If
            If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
            End If
          
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                lbl_ward.Visible = True
                lbl_bed.Visible = True
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_status.Text = "ST"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
            End If
          
         
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
        
            End If
        
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
                       
           '******
                If SELINDICATIONFLAG = False Then
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELINDICATIONRS = SELINDICATIONQR.OpenResultset(1, 3)
                    SELINDICATIONFLAG = True
                Else
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    SELINDICATIONRS.MoreResults
                    SELINDICATIONRS.Requery
                End If
                If SELINDICATIONRS.RowCount > 0 Then
                    If txt_modcode.Text = "PF" Then
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND7"))
                        End If
                    End If
                    If txt_modcode.Text = "ST" Then
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                    End If
                End If
          '*****
            
        '    seldrequisitn_qr(0) = MaskEdBox1.FormattedText
        '    seldrequisitn_rs.MoreResults
        '    seldrequisitn_rs.Requery
        '    If seldrequisitn_rs.RowCount > 0 Then
        '     prpathno = ""
        '     While Not seldrequisitn_rs.EOF
        '     prpathno = prpathno & " " & seldrequisitn_rs("pathno")
        '     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
        '     seldrequisitn_rs.MoveNext
        '     Wend
        '     txt_prepathreqno.Text = prpathno
        '    End If
        '    txt_addendum.Enabled = False
        '    cmb_addendumby.Enabled = False
            

        
        If seldissecentrflag1 = False Then
            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
            seldissecentrflag1 = True
        Else
            seldissecentr1_rs.MoreResults
            seldissecentr1_rs.Requery
        End If
        If seldissecentr1_rs.RowCount > 0 Then
            'cmb_enterby.clear
            While Not seldissecentr1_rs.EOF
               ' cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
                seldissecentr1_rs.MoveNext
            Wend
        End If

        If seldissecfinflag = False Then
            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
            seldissecfinflag1 = True
        Else
            seldissecfin1_rs.MoreResults
            seldissecfin1_rs.Requery
        End If
        If seldissecfin1_rs.RowCount > 0 Then
            'cmb_finalby.clear
            While Not seldissecfin1_rs.EOF
             '   cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
                seldissecfin1_rs.MoveNext
            Wend
        End If
            
            '******
          
STENABL ("ENA")
End If
        If viewflag = True Then
        
           
            
            viewflag = False
            If SELDREQSTAFFD1FLAG = False Then
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
                Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
            Else
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = MaskEdBox1.FormattedText
                SELDREQSTAFFDRS.MoreResults
                SELDREQSTAFFDRS.Requery
                
            End If
            If SELDREQSTAFFDRS.RowCount > 0 Then
                txt_sex.Text = SELDREQSTAFFDRS("sex")
                txt_name.Text = SELDREQSTAFFDRS("PNAME")
                txt_age = SELDREQSTAFFDRS("age")
                txt_pridoc.Text = ""
                txt_category.Text = SELDREQSTAFFDRS("RELATION")
            End If
        
            If REFDRFLAG1 = True Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
        
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                cmb_provisional.Text = REFDR_RS("PREPORTBY")
                'cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
                End If
            End If
            
            
            
          
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_status.Text = "ST"
            End If
          
         If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                'txt_ihmeterial.Text = SELPATHDREQUISITN_RS("NMETERIALS")
                txt_refdoc.Text = selpathdrequisitn_rs("REFDOC")
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
            End If
        
        
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
            '*******
            '*****
                If SELINDICATIONFLAG = False Then
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELINDICATIONRS = SELINDICATIONQR.OpenResultset(1, 3)
                    SELINDICATIONFLAG = True
                Else
                    SELINDICATIONQR(0) = req_chglen(cmb_reqno.Text)
                    SELINDICATIONRS.MoreResults
                    SELINDICATIONRS.Requery
                End If
                If SELINDICATIONRS.RowCount > 0 Then
                    If txt_modcode.Text = "PF" Then
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_pftindication.Text = txt_pftindication & " " & Trim(SELINDICATIONRS("IND7"))
                        End If
                    End If
                    If txt_modcode.Text = "ST" Then
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND1"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND2"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND3"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND4"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND5"))
                        End If
                        If SELINDICATIONRS.EOF = False Then
                            txt_stindication.Text = txt_stindication & " " & Trim(SELINDICATIONRS("IND6"))
                        End If
                    End If
                End If
            '*****
                If Trim(txt_modcode.Text) = "PF" Then
                    If SELPFREPORTFLAG = False Then
                        SELPFREPORTQR(0) = Trim(req_chglen(cmb_reqno.Text))
                        Set SELPFREPORTRS = SELPFREPORTQR.OpenResultset(1, 3)
                        SELPFREPORTFLAG = True
                    Else
                        SELPFREPORTQR(0) = Trim(req_chglen(cmb_reqno.Text))
                        SELPFREPORTRS.MoreResults
                        SELPFREPORTRS.Requery
                    End If
                    If SELPFREPORTRS.RowCount > 0 Then
                        If SELPFREPORTRS.EOF = False Then
                            TXT_PFTREPORT.Text = "" & SELPFREPORTRS("REPORT")
                            msk_pfexamdt.Text = Format(SELPFREPORTRS("examdt"), "dd/mm/yyyy")
                        End If
                        If SELPFREPORTRS.EOF = False Then
                            TXT_PFTIMPRESSION.Text = "" & SELPFREPORTRS("IMPRESSION")
                        End If
                    End If
                    TXT_PFTREPORT.Enabled = False
                    TXT_PFTIMPRESSION.Enabled = False
                
                End If
                If Trim(txt_modcode.Text) = "ST" Then
                    STENABL ("DIS")
                    If SELSTREPORTFLAG = False Then
                        SELSTREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELSTREPORTRS = SELSTREPORTQR.OpenResultset(1, 3)
                        SELSTREPORTFLAG = True
                    Else
                        SELSTREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        SELSTREPORTRS.MoreResults
                        SELSTREPORTRS.Requery
                    End If
                    If SELSTREPORTRS.RowCount > 0 Then
                        CMB_RISKFACTOR.Text = Trim(SELSTREPORTRS("RISKFACTOR"))
                        TXT_MAXHR.Text = SELSTREPORTRS("MAXHR")
                        TXT_MAXBP.Text = SELSTREPORTRS("MAXBP")
                        TXT_DOUBLEPRODUCT.Text = SELSTREPORTRS("DOBPRODUCT")
                        TXT_THRACHIEVED.Text = SELSTREPORTRS("THRACHIVED")
                        If SELSTREPORTRS.EOF = False Then
                            TXT_STTCHANGES.Text = " " & SELSTREPORTRS("STTCHANGES")
                        End If
                        If SELSTREPORTRS.EOF = False Then
                            TXT_ARRYTHMIA.Text = " " & SELSTREPORTRS("ARRYTHMIA")
                        End If
                        CMB_RESFORTREAT.Text = SELSTREPORTRS("RESTERMI")
                        TXT_METSA.Text = SELSTREPORTRS("METSACHVED")
                        If SELSTREPORTRS.EOF = False Then
                        
                            TXT_STIMPRESSION.Text = " " & SELSTREPORTRS("IMPRESSION")
                        End If
                        If SELSTREPORTRS.EOF = False Then
                            cmb_stproperformed.Text = " " & SELSTREPORTRS("PROPERBY")
                        End If
                        If SELSTREPORTRS.EOF = False Then
                            msk_stexamdt.Text = Format(SELSTREPORTRS("fld1"), "dd/mm/yyyy")
                        End If
                    End If
                End If
                If Trim(txt_modcode.Text) = "EK" Then
                    ECENABL ("DIS")
                    If SELECREPORTFLAG = False Then
                        SELECREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELECREPORTRS = SELECREPORTQR.OpenResultset(1, 3)
                        SELECREPORTFLAG = True
                    Else
                        SELECREPORTQR(0) = req_chglen(cmb_reqno.Text)
                        SELECREPORTRS.MoreResults
                        SELECREPORTRS.Requery
                    End If
                    If SELECREPORTRS.RowCount > 0 Then
                        TXT_RATE.Text = Trim(SELECREPORTRS("RATE"))
                        TXT_ATRIALRATE.Text = Trim(SELECREPORTRS("ATRIALRATE"))
                        TXT_VENRATE.Text = Trim(SELECREPORTRS("VENRATE"))
                        If SELECREPORTRS.EOF = False Then
                            TXT_PWAVE.Text = " " & Trim(SELECREPORTRS("PWAVE"))
                       
                        End If
                        TXT_PRINTERVAL.Text = Trim(SELECREPORTRS("PRINTERVAL"))
                        If SELECREPORTRS.EOF = False Then
                            TXT_RHYTHM.Text = " " & Trim(SELECREPORTRS("RHYTHM"))
                        End If
                        TXT_QRSM.Text = Trim(SELECREPORTRS("QRSMORPH"))
                        TXT_QRSD.Text = Trim(SELECREPORTRS("QRSDUR"))
                        TXT_QTINTERVAL.Text = Trim(SELECREPORTRS("QTINTERVAL"))
                        If SELECREPORTRS.EOF = False Then
                            TXT_STSEGMENT.Text = " " & Trim(SELECREPORTRS("STSEGMENT"))
                            
                        End If
                        If SELECREPORTRS.EOF = False Then
                            TXT_TWAVE.Text = " " & Trim(SELECREPORTRS("TWAVE"))
                        End If
                        If SELECREPORTRS.EOF = False Then
                            TXT_ECIMPRESSION.Text = " " & Trim(SELECREPORTRS("IMPRESSION"))
                        End If
                    End If
                End If
        
        End If
End Sub

Public Sub PFASSIGN()
On Error GoTo MSG:
rdb.BeginTrans
    If flagsave = 1 Then
        '*** Inserting into  PFT REPORT
            INSPFREPORTQR(0) = req_chglen(cmb_reqno.Text)
           If Len(Trim(TXT_PFTREPORT.Text)) > 0 Then
            INSPFREPORTQR(1) = Null
            INSPFREPORTQR(1) = TXT_PFTREPORT.Text
           Else
            INSPFREPORTQR(1) = Null
           End If
           
           If Len(Trim(TXT_PFTIMPRESSION.Text)) > 0 Then
            INSPFREPORTQR(2) = Null
            INSPFREPORTQR(2) = TXT_PFTIMPRESSION.Text
           Else
            INSPFREPORTQR(2) = Null
           End If
           INSPFREPORTQR(3) = GET_REVDATE(msk_pfexamdt.FormattedText)
           INSPFREPORTQR(4) = Trim(cmb_pfproperby.Text)
           INSPFREPORTQR.Execute
            
       '*** Inserting into PFT REPORT
                       
            
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
        
            
            
        
    
    End If
    If flagsave = 2 Then
            
            '*** updating the pftreport
            If Len(Trim(TXT_PFTREPORT.Text)) > 0 Then
             UPDPFREPORTQR(0) = Null
             UPDPFREPORTQR(0) = TXT_PFTREPORT.Text
            Else
             UPDPFREPORTQR(0) = Null
            End If
            
            If Len(Trim(TXT_PFTIMPRESSION.Text)) > 0 Then
             UPDPFREPORTQR(1) = Null
             UPDPFREPORTQR(1) = TXT_PFTIMPRESSION.Text
            Else
             UPDPFREPORTQR(1) = Null
            End If
            UPDPFREPORTQR(2) = GET_REVDATE(msk_pfexamdt.FormattedText)
            UPDPFREPORTQR(3) = Trim(cmb_pfproperby.Text)
            UPDPFREPORTQR(4) = req_chglen(cmb_reqno.Text)
            UPDPFREPORTQR.Execute
            
            '*** updating the pftreport
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
           
            
            
            
    End If
rdb.CommitTrans
MsgBox "Record is Saved"
btnenabl ("INI")

Exit Sub
MSG:
    rdb.RollbackTrans
MsgBox Err.Description
End Sub

Public Sub STASSIGN()
On Error GoTo MSG:
rdb.BeginTrans
   If flagsave = 1 Then
            'STR = "INSERT INTO DISLIB.DSTREPORT(REQNO,RISKFACTOR,MAXHR,MAXBP,DOBPRODUCT,
            'THRACHIVED,STTCHANGES,ARRYTHMIA,RESTERMI,METSAACHVED,
            'IMPRESSION,PROPERBY) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
        
            INSSTREPORTQR(0) = req_chglen(cmb_reqno.Text)
            INSSTREPORTQR(1) = Trim(CMB_RISKFACTOR.Text)
            INSSTREPORTQR(2) = Trim(TXT_MAXHR.Text)
            INSSTREPORTQR(3) = Trim(TXT_MAXBP.Text)
            INSSTREPORTQR(4) = Trim(TXT_DOUBLEPRODUCT.Text)
            INSSTREPORTQR(5) = Trim(TXT_THRACHIEVED.Text)
            If Len(Trim(TXT_STTCHANGES.Text)) > 0 Then
                INSSTREPORTQR(6) = Null
                INSSTREPORTQR(6) = Trim(TXT_STTCHANGES.Text)
            Else
                INSSTREPORTQR(6) = Null
            End If
            If Len(Trim(TXT_ARRYTHMIA.Text)) > 0 Then
                INSSTREPORTQR(7) = Null
                INSSTREPORTQR(7) = Trim(TXT_ARRYTHMIA.Text)
            Else
                INSSTREPORTQR(7) = Null
            End If
            INSSTREPORTQR(8) = Trim(CMB_RESFORTREAT)
            INSSTREPORTQR(9) = Trim(TXT_METSA.Text)
            If Len(Trim(TXT_STIMPRESSION.Text)) > 0 Then
                INSSTREPORTQR(10) = Null
                INSSTREPORTQR(10) = Trim(TXT_STIMPRESSION.Text)
            Else
                INSSTREPORTQR(10) = Null
            End If
            If Len(Trim(cmb_stproperformed.Text)) > 0 Then
                INSSTREPORTQR(11) = Null
                INSSTREPORTQR(11) = Trim(cmb_stproperformed.Text)
            Else
                INSSTREPORTQR(11) = Null
            End If
            INSSTREPORTQR(12) = GET_REVDATE(msk_stexamdt.FormattedText)
            INSSTREPORTQR.Execute
        
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
       
    End If
    If flagsave = 2 Then
     
            
            
            UPDSTREPORTQR(0) = Trim(CMB_RISKFACTOR.Text)
            UPDSTREPORTQR(1) = Trim(TXT_MAXHR.Text)
            UPDSTREPORTQR(2) = Trim(TXT_MAXBP.Text)
            UPDSTREPORTQR(3) = Trim(TXT_DOUBLEPRODUCT.Text)
            UPDSTREPORTQR(4) = Trim(TXT_THRACHIEVED.Text)
            If Len(Trim(TXT_STTCHANGES.Text)) > 0 Then
                UPDSTREPORTQR(5) = Null
                UPDSTREPORTQR(5) = Trim(TXT_STTCHANGES.Text)
            Else
                UPDSTREPORTQR(5) = Null
            End If
            If Len(Trim(TXT_ARRYTHMIA.Text)) > 0 Then
                UPDSTREPORTQR(6) = Null
                UPDSTREPORTQR(6) = Trim(TXT_ARRYTHMIA.Text)
            Else
                UPDSTREPORTQR(6) = Null
            End If
            UPDSTREPORTQR(7) = Trim(CMB_RESFORTREAT)
            UPDSTREPORTQR(8) = Trim(TXT_METSA.Text)
            If Len(Trim(TXT_STIMPRESSION.Text)) > 0 Then
                UPDSTREPORTQR(9) = Null
                UPDSTREPORTQR(9) = Trim(TXT_STIMPRESSION.Text)
            Else
                UPDSTREPORTQR(9) = Null
            End If
            If Len(Trim(cmb_stproperformed.Text)) > 0 Then
                UPDSTREPORTQR(10) = Null
                UPDSTREPORTQR(10) = Trim(cmb_stproperformed.Text)
            Else
                UPDSTREPORTQR(10) = Null
            End If
            UPDSTREPORTQR(11) = GET_REVDATE(msk_stexamdt.FormattedText)
            UPDSTREPORTQR(12) = req_chglen(cmb_reqno.Text)
            UPDSTREPORTQR.Execute
        
            
            
            
            
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
           
            
            
            
    End If
rdb.CommitTrans
MsgBox "Record is Saved"
btnenabl ("INI")
STENABL ("DIS")
Exit Sub
MSG:
    rdb.RollbackTrans
MsgBox Err.Description
End Sub

Public Sub ECASSIGN()
On Error GoTo MSG:
rdb.BeginTrans
      If flagsave = 1 Then
        
            INSECREPORTQR(0) = req_chglen(cmb_reqno.Text)
            INSECREPORTQR(1) = Trim(TXT_RATE.Text)
            INSECREPORTQR(2) = Trim(TXT_ATRIALRATE.Text)
            INSECREPORTQR(3) = Trim(TXT_VENRATE.Text)
            If Len(Trim(TXT_PWAVE.Text)) > 0 Then
                INSECREPORTQR(4) = Null
                INSECREPORTQR(4) = Trim(TXT_PWAVE.Text)
            Else
                INSECREPORTQR(4) = Null
            End If
            INSECREPORTQR(5) = Trim(TXT_PRINTERVAL.Text)
            If Len(Trim(TXT_RHYTHM.Text)) > 0 Then
                INSECREPORTQR(6) = Null
                INSECREPORTQR(6) = Trim(TXT_RHYTHM.Text)
            Else
                INSECREPORTQR(6) = Null
            End If
            INSECREPORTQR(7) = Trim(TXT_QRSM.Text)
            INSECREPORTQR(8) = Trim(TXT_QRSD.Text)
            INSECREPORTQR(9) = Trim(TXT_QTINTERVAL.Text)
            If Len(Trim(TXT_STSEGMENT.Text)) > 0 Then
                INSECREPORTQR(10) = Null
                INSECREPORTQR(10) = Trim(TXT_STSEGMENT.Text)
            Else
            
                INSECREPORTQR(10) = Null
            End If
            INSECREPORTQR(11) = Trim(TXT_TWAVE.Text)
            INSECREPORTQR(12) = Trim(TXT_QTINTERVAL.Text)
            If Len(Trim(TXT_ECIMPRESSION.Text)) > 0 Then
                INSECREPORTQR(13) = Null
                INSECREPORTQR(13) = Trim(TXT_ECIMPRESSION.Text)
            Else
                INSECREPORTQR(13) = Null
            End If
            INSECREPORTQR(14) = GET_REVDATE(msk_ecexamdt.FormattedText)
            INSECREPORTQR(15) = Trim(cmb_edproperby.Text)
            INSECREPORTQR.Execute
                                   
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
        
    End If
    If flagsave = 2 Then
     
        
            
            UPDECREPORTQR(0) = Trim(TXT_RATE.Text)
            UPDECREPORTQR(1) = Trim(TXT_ATRIALRATE.Text)
            UPDECREPORTQR(2) = Trim(TXT_VENRATE.Text)
            If Len(Trim(TXT_PWAVE.Text)) > 0 Then
                UPDECREPORTQR(3) = Null
                UPDECREPORTQR(3) = Trim(TXT_PWAVE.Text)
            Else
                UPDECREPORTQR(3) = Null
            End If
            UPDECREPORTQR(4) = Trim(TXT_PRINTERVAL.Text)
            If Len(Trim(TXT_RHYTHM.Text)) > 0 Then
                UPDECREPORTQR(5) = Null
                UPDECREPORTQR(5) = Trim(TXT_RHYTHM.Text)
            Else
                UPDECREPORTQR(5) = Null
            End If
            UPDECREPORTQR(6) = Trim(TXT_QRSM.Text)
            UPDECREPORTQR(7) = Trim(TXT_QRSD.Text)
            UPDECREPORTQR(8) = Trim(TXT_QTINTERVAL.Text)
            If Len(Trim(TXT_STSEGMENT.Text)) > 0 Then
                UPDECREPORTQR(9) = Null
                UPDECREPORTQR(9) = Trim(TXT_STSEGMENT.Text)
            Else
            
                UPDECREPORTQR(9) = Null
            End If
            UPDECREPORTQR(10) = Trim(TXT_TWAVE.Text)
            UPDECREPORTQR(11) = Trim(TXT_QTINTERVAL.Text)
            If Len(Trim(TXT_ECIMPRESSION.Text)) > 0 Then
                UPDECREPORTQR(12) = Null
                UPDECREPORTQR(12) = Trim(TXT_ECIMPRESSION.Text)
            Else
                UPDECREPORTQR(12) = Null
            End If
            UPDECREPORTQR(13) = GET_REVDATE(msk_ecexamdt.FormattedText)
            UPDECREPORTQR(14) = Trim(cmb_edproperby.Text)
            UPDECREPORTQR(15) = req_chglen(cmb_reqno.Text)
            UPDECREPORTQR.Execute
        
        
        
        
        
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
    End If
rdb.CommitTrans
MsgBox "Record is Saved"
btnenabl ("INI")
ECENABL ("DIS")
Exit Sub
MSG:
    rdb.RollbackTrans
MsgBox Err.Description
End Sub

Public Sub STENABL(a As String)
    If a = "DIS" Then
        CMB_RISKFACTOR.Enabled = False
        TXT_MAXHR.Enabled = False
        TXT_MAXBP.Enabled = False
        TXT_DOUBLEPRODUCT.Enabled = False
        TXT_THRACHIEVED.Enabled = False
        TXT_STTCHANGES.Enabled = False
        TXT_ARRYTHMIA.Enabled = False
        CMB_RESFORTREAT.Enabled = False
        TXT_METSA.Enabled = False
        TXT_STIMPRESSION.Enabled = False
        cmb_stproperformed.Enabled = False
        msk_stexamdt.Enabled = False
         cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
        cmb_provisional.Enabled = False
    
    End If
    If a = "ENA" Then
        CMB_RISKFACTOR.Enabled = True
        TXT_MAXHR.Enabled = True
        TXT_MAXBP.Enabled = True
        TXT_DOUBLEPRODUCT.Enabled = True
        TXT_THRACHIEVED.Enabled = True
        TXT_STTCHANGES.Enabled = True
        TXT_ARRYTHMIA.Enabled = True
        CMB_RESFORTREAT.Enabled = True
        TXT_METSA.Enabled = True
        TXT_STIMPRESSION.Enabled = True
        cmb_stproperformed.Enabled = True
        msk_stexamdt.Enabled = True
         cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        cmb_provisional.Enabled = True
    
    End If
End Sub

Public Sub ECENABL(a As String)
 If a = "DIS" Then
    TXT_RATE.Enabled = False
    TXT_ATRIALRATE.Enabled = False
    TXT_VENRATE.Enabled = False
    TXT_PWAVE.Enabled = False
    TXT_PRINTERVAL.Enabled = False
    TXT_RHYTHM.Enabled = False
    TXT_QRSM.Enabled = False
    TXT_QRSD.Enabled = False
    TXT_QTINTERVAL.Enabled = False
    TXT_STSEGMENT.Enabled = False
    TXT_TWAVE.Enabled = False
    TXT_ECIMPRESSION.Enabled = False
    msk_ecexamdt.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provisional.Enabled = False
    cmb_edproperby.Enabled = False
 End If
 If a = "ENA" Then
    TXT_RATE.Enabled = True
    TXT_ATRIALRATE.Enabled = True
    TXT_VENRATE.Enabled = True
    TXT_PWAVE.Enabled = True
    TXT_PRINTERVAL.Enabled = True
    TXT_RHYTHM.Enabled = True
    TXT_QRSM.Enabled = True
    TXT_QRSD.Enabled = True
    TXT_QTINTERVAL.Enabled = True
    TXT_STSEGMENT.Enabled = True
    TXT_TWAVE.Enabled = True
    TXT_ECIMPRESSION.Enabled = True
    msk_ecexamdt.Enabled = True
     cmb_enterby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provisional.Enabled = True
    cmb_edproperby.Enabled = True
    
 End If
End Sub

Public Sub ECCLEAR()
    TXT_RATE.Text = ""
    TXT_ATRIALRATE.Text = ""
    TXT_VENRATE.Text = ""
    TXT_PWAVE.Text = ""
    TXT_PRINTERVAL.Text = ""
    TXT_RHYTHM.Text = ""
    TXT_QRSM.Text = ""
    TXT_QRSD.Text = ""
    TXT_QTINTERVAL.Text = ""
    TXT_STSEGMENT.Text = ""
    TXT_TWAVE.Text = ""
    TXT_ECIMPRESSION.Text = ""
    msk_ecexamdt.Mask = ""
    msk_ecexamdt.Text = ""
    cmb_edproperby.Text = ""
    cmb_provisional.Text = ""
    cmb_enterby.Text = ""
    cmb_finalby.Text = ""
End Sub

Public Sub STCLEAR()
        CMB_RISKFACTOR.Text = ""
        TXT_MAXHR.Text = ""
        TXT_MAXBP.Text = ""
        TXT_DOUBLEPRODUCT.Text = ""
        TXT_THRACHIEVED.Text = ""
        TXT_STTCHANGES.Text = ""
        TXT_ARRYTHMIA.Text = ""
        CMB_RESFORTREAT.Text = ""
        TXT_METSA.Text = ""
        TXT_STIMPRESSION.Text = ""
        cmb_stproperformed.Text = ""
        'msk_stexamdt.Text = ""
        msk_stexamdt.Mask = "##/##/####"
        cmb_provisional.Text = ""
        cmb_enterby.Text = ""
        cmb_finalby.Text = ""
        
End Sub

Private Sub TXT_ECIMPRESSION_GotFocus()
On Error GoTo m1:
If NEWFLAG = True Then
    TXT_ECIMPRESSION.Text = "With in Normal Limits"
End If
w_impression.Visible = False
Set d_impression = w_impression.ActiveDocument
d_impression.Content = ""
d_impression.Content = TXT_ECIMPRESSION
w_impression.Caption = "Impression"
'w_frozen.WindowState = wdWindowStateMaximize
w_impression.WindowState = wdWindowStateNormal
d_impression.Content = TXT_ECIMPRESSION
w_impression.Visible = True
w_impression.Activate
m1:
On Error Resume Next
Set w_impression = CreateObject("Word.Application")
w_impression.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_impression.Caption = "Impression"
'w_frozen.WindowState = wdWindowStateMaximize
w_impression.WindowState = wdWindowStateNormal
Set d_impression = w_impression.ActiveDocument
d_impression.Content = Trim(TXT_ECIMPRESSION)
w_impression.Visible = True
w_impression.Activate
End Sub

Private Sub TXT_ECIMPRESSION_LostFocus()
On Error GoTo m2:
Set d_impression = w_impression.ActiveDocument
TXT_ECIMPRESSION.Text = d_impression.Content
'flagimp = False
d_impression.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_impression.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_impression.Close (wdDoNotSaveChanges)
w_impression.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Private Sub TXT_PWAVE_GotFocus()
On Error GoTo m1:
w_pwave.Visible = False
Setd_pwave = w_pwave.ActiveDocument
d_pwave.Content = ""
d_pwave.Content = TXT_PWAVE
w_pwave.Caption = "Pwave"
'w_frozen.WindowState = wdWindowStateMaximize
w_pwave.WindowState = wdWindowStateNormal
d_pwave.Content = TXT_PWAVE
w_pwave.Visible = True
w_pwave.Activate
m1:
On Error Resume Next
Set w_pwave = CreateObject("Word.Application")
w_pwave.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_pwave.Caption = "Pwave"
'w_frozen.WindowState = wdWindowStateMaximize
w_pwave.WindowState = wdWindowStateNormal
Set d_pwave = w_pwave.ActiveDocument
d_pwave.Content = Trim(TXT_PWAVE)
w_pwave.Visible = True
w_pwave.Activate

End Sub

Private Sub TXT_PWAVE_LostFocus()
On Error GoTo m2:
Set d_pwave = w_pwave.ActiveDocument
TXT_PWAVE.Text = d_pwave.Content
'flagimp = False
d_pwave.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_pwave.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_pwave.Close (wdDoNotSaveChanges)
w_pwave.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub





Private Sub TXT_RHYTHM_GotFocus()
On Error GoTo m1:
w_rhythm.Visible = False
Set d_rhythm = w_rhythm.ActiveDocument
d_rhythm.Content = ""
d_rhythm.Content = TXT_RHYTHM
w_rhythm.Caption = "Rhythm"
'w_frozen.WindowState = wdWindowStateMaximize
w_rhythm.WindowState = wdWindowStateNormal
d_rhythm.Content = TXT_RHYTHM
w_rhythm.Visible = True
w_rhythm.Activate
m1:
On Error Resume Next
Set w_rhythm = CreateObject("Word.Application")
w_rhythm.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_rhythm.Caption = "Rhythm"
'w_frozen.WindowState = wdWindowStateMaximize
w_rhythm.WindowState = wdWindowStateNormal
Set d_rhythm = w_rhythm.ActiveDocument
d_rhythm.Content = Trim(TXT_RHYTHM)
w_rhythm.Visible = True
w_rhythm.Activate
End Sub

Private Sub TXT_RHYTHM_LostFocus()
On Error GoTo m2:
Set d_rhythm = w_rhythm.ActiveDocument
TXT_RHYTHM.Text = d_rhythm.Content
'flagimp = False
d_rhythm.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_rhythm.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_rhythm.Close (wdDoNotSaveChanges)
w_rhythm.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub


Private Sub TXT_STSEGMENT_GotFocus()
On Error GoTo m1:
w_stsegment.Visible = False
Set d_stsegment = w_stsegment.ActiveDocument
d_stsegment.Content = ""
d_stsegment.Content = TXT_STSEGMENT
w_stsegment.Caption = "ST Segement"
'w_frozen.WindowState = wdWindowStateMaximize
w_stsegment.WindowState = wdWindowStateNormal
d_stsegment.Content = TXT_STSEGMENT
w_stsegment.Visible = True
w_stsegment.Activate
m1:
On Error Resume Next
Set w_stsegment = CreateObject("Word.Application")
w_stsegment.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_stsegment.Caption = "ST Segment"
'w_frozen.WindowState = wdWindowStateMaximize
w_stsegment.WindowState = wdWindowStateNormal
Set d_stsegment = w_stsegment.ActiveDocument
d_stsegment.Content = Trim(TXT_STSEGMENT)
w_stsegment.Visible = True
w_stsegment.Activate

End Sub

Private Sub TXT_STSEGMENT_LostFocus()
On Error GoTo m2:
Set d_stsegment = w_stsegment.ActiveDocument
TXT_STSEGMENT.Text = d_stsegment.Content
'flagimp = False
d_stsegment.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_stsegment.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_stsegment.Close (wdDoNotSaveChanges)
w_stsegment.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Private Sub TXT_TWAVE_GotFocus()
On Error GoTo m1:
w_twave.Visible = False
Set d_twave = w_twave.ActiveDocument
d_twave.Content = ""
d_twave.Content = TXT_TWAVE
w_twave.Caption = "Twave"
'w_frozen.WindowState = wdWindowStateMaximize
w_twave.WindowState = wdWindowStateNormal
d_twave.Content = TXT_TWAVE
w_twave.Visible = True
w_twave.Activate
m1:
On Error Resume Next
Set w_twave = CreateObject("Word.Application")
w_twave.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_twave.Caption = "Twave"
'w_frozen.WindowState = wdWindowStateMaximize
w_twave.WindowState = wdWindowStateNormal
Set d_twave = w_twave.ActiveDocument
d_twave.Content = Trim(TXT_TWAVE)
w_twave.Visible = True
w_twave.Activate

End Sub

Private Sub TXT_TWAVE_LostFocus()
On Error GoTo m2:
Set d_twave = w_twave.ActiveDocument
TXT_TWAVE.Text = d_twave.Content
'flagimp = False
d_twave.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_twave.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_twave.Close (wdDoNotSaveChanges)
w_twave.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub
