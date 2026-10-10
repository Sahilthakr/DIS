VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{9B8DFCFF-6726-11D6-99D4-000244193CE4}#1.0#0"; "PROJECT1.OCX"
Begin VB.Form frm_ecgtest 
   Caption         =   "ECG TEST"
   ClientHeight    =   8595
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame12 
      Height          =   735
      Left            =   0
      TabIndex        =   54
      Top             =   7200
      Width           =   11775
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         Left            =   9600
         TabIndex        =   56
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmb_enterby 
         Height          =   315
         Left            =   1320
         TabIndex        =   55
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
         Height          =   195
         Left            =   7920
         TabIndex        =   58
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl_enter 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
         Height          =   195
         Left            =   480
         TabIndex        =   57
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   0
      TabIndex        =   45
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         MouseIcon       =   "Form1.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MouseIcon       =   "Form1.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "Form1.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "Form1.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5880
         MouseIcon       =   "Form1.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8760
         MouseIcon       =   "Form1.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10440
         MouseIcon       =   "Form1.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7200
         MouseIcon       =   "Form1.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   42
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   41
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame common 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   11775
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   0
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
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
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   36
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   35
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Path No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   34
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   32
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   31
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   30
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   29
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   26
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   25
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   2520
         Width           =   1830
      End
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   495
      Left            =   2400
      TabIndex        =   39
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
   End
   Begin VB.Frame Frame3 
      Caption         =   "ECG REPORT"
      Height          =   3495
      Left            =   0
      TabIndex        =   59
      Top             =   3720
      Width           =   11775
      Begin VB.TextBox TXT_HB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   71
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   70
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   69
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   68
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   67
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   66
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   65
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   64
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   63
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   62
         Top             =   2160
         Width           =   6135
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   61
         Top             =   2520
         Width           =   6135
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   60
         Top             =   3000
         Width           =   6135
      End
      Begin VB.Label LBL_HB 
         AutoSize        =   -1  'True
         Caption         =   "Unit"
         Height          =   195
         Left            =   240
         TabIndex        =   83
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PR Interval"
         Height          =   195
         Left            =   3840
         TabIndex        =   82
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Rhythm"
         Height          =   195
         Left            =   240
         TabIndex        =   81
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Atrial rate"
         Height          =   195
         Left            =   3840
         TabIndex        =   80
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ventricular rate"
         Height          =   195
         Left            =   240
         TabIndex        =   79
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Pwave"
         Height          =   195
         Left            =   3840
         TabIndex        =   78
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "QRS Morphology"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "QT interval "
         Height          =   195
         Left            =   240
         TabIndex        =   76
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "QRS duration"
         Height          =   195
         Left            =   3840
         TabIndex        =   75
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "ST Segment "
         Height          =   195
         Left            =   240
         TabIndex        =   74
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "T Wave"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   3000
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "PFT Test"
      Height          =   3495
      Left            =   0
      TabIndex        =   109
      Top             =   3720
      Width           =   11775
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   112
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   110
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   600
         TabIndex        =   113
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   600
         TabIndex        =   111
         Top             =   720
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stress Test"
      Height          =   3495
      Left            =   0
      TabIndex        =   84
      Top             =   3720
      Width           =   11775
      Begin VB.TextBox TXT_TRAP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   96
         Top             =   3120
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   95
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   94
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   93
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   92
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   91
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5280
         TabIndex        =   90
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   89
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   88
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         TabIndex        =   87
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   86
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   85
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label LBL_TRAP 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "METS achived"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% of THR achieved"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   106
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Double Product"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ST T changes"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Arrythmia"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   103
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Reason for termination"
         Height          =   195
         Left            =   3120
         TabIndex        =   102
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Max HR"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   101
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "/Min"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   100
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Risk Factor"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Max. BP "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   98
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "mm Hg"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6720
         TabIndex        =   97
         Top             =   1080
         Width           =   585
      End
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9840
      TabIndex        =   44
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9840
      TabIndex        =   43
      Top             =   480
      Width           =   420
   End
   Begin VB.Label LBL_CP 
      AutoSize        =   -1  'True
      Caption         =   "ECG TEST REPORT"
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
      Left            =   3600
      TabIndex        =   40
      Top             =   480
      Width           =   2745
   End
End
Attribute VB_Name = "frm_ecgtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_back_Click()
DIS_MAIN.show
Me.Hide
End Sub
