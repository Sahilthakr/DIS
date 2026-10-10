VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form micro_bact 
   Caption         =   "Form2"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form2"
   ScaleHeight     =   8280
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9600
      TabIndex        =   100
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9600
      TabIndex        =   99
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   120
      TabIndex        =   89
      Top             =   8640
      Width           =   10215
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   350
         Left            =   5520
         MouseIcon       =   "micro_bact.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   98
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   350
         Left            =   7680
         MouseIcon       =   "micro_bact.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   350
         Left            =   6600
         MouseIcon       =   "micro_bact.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   4440
         MouseIcon       =   "micro_bact.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   95
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   350
         Left            =   3360
         MouseIcon       =   "micro_bact.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   94
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   350
         Left            =   2280
         MouseIcon       =   "micro_bact.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   93
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   350
         Left            =   1200
         MouseIcon       =   "micro_bact.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   92
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   350
         Left            =   120
         MouseIcon       =   "micro_bact.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   91
         Top             =   240
         Width           =   950
      End
      Begin VB.CommandButton CMD_DECOMMIT 
         Caption         =   "Sp. Serology"
         Height          =   350
         Left            =   8760
         TabIndex        =   90
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame frm_enter 
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   80
      Top             =   7680
      Width           =   10215
      Begin VB.ComboBox cmb_enteryby 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   84
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox cmb_final 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   83
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   82
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   6240
         TabIndex        =   81
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Entered By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Finalised By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   87
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lbl_addendum 
         Caption         =   "ADDENDUM"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Addendum By"
         Height          =   255
         Left            =   5040
         TabIndex        =   85
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame frm_bacteria 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   -120
      TabIndex        =   40
      Top             =   3360
      Width           =   12120
      Begin VB.TextBox txt_1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   60
         Top             =   2400
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt_comment 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         MaxLength       =   160
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   3720
         Width           =   7575
      End
      Begin VB.TextBox txt_epith 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   58
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_pus 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   57
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txt_other 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7080
         MaxLength       =   250
         TabIndex        =   56
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_ziehl 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   55
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_koh 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   54
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gpc 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "micro_bact.frx":1850
         Left            =   3960
         List            =   "micro_bact.frx":185D
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gnc 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "micro_bact.frx":187F
         Left            =   7080
         List            =   "micro_bact.frx":188C
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gpb 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "micro_bact.frx":18AE
         Left            =   9960
         List            =   "micro_bact.frx":18BB
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gnb 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "micro_bact.frx":18DD
         Left            =   1560
         List            =   "micro_bact.frx":18EA
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cmb_patho 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "micro_bact.frx":190C
         Left            =   1560
         List            =   "micro_bact.frx":1916
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txt_report 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1560
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   3120
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.TextBox txt_2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         MaxLength       =   1
         TabIndex        =   47
         Top             =   2400
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt_3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   46
         Top             =   2400
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt_organ1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   45
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox TXT_ORGAN2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4200
         TabIndex        =   44
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox TXT_ORGAN3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6720
         TabIndex        =   43
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cmb_india 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "micro_bact.frx":1923
         Left            =   7080
         List            =   "micro_bact.frx":1925
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cmb_antibiotic 
         Height          =   315
         Left            =   1800
         TabIndex        =   41
         Top             =   2400
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid grid_anti 
         Height          =   375
         Left            =   5040
         TabIndex        =   61
         Top             =   1440
         Visible         =   0   'False
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   128
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         Appearance      =   0
         FormatString    =   "   ANTIBIOTIC                               |^         1           |^          2         |^              3    "
      End
      Begin MSFlexGridLib.MSFlexGrid grid_dmbanti 
         Height          =   1455
         Left            =   1560
         TabIndex        =   62
         Top             =   2160
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   128
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         Appearance      =   0
         FormatString    =   "   ANTIBIOTIC                               |^         1           |^          2         |^              3    "
      End
      Begin VB.Label lbl_3 
         AutoSize        =   -1  'True
         Caption         =   "3"
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6600
         TabIndex        =   79
         Top             =   1800
         Width           =   120
      End
      Begin VB.Label lbl_2 
         AutoSize        =   -1  'True
         Caption         =   "2"
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   78
         Top             =   1800
         Width           =   120
      End
      Begin VB.Label lbl_1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   77
         Top             =   1800
         Width           =   120
      End
      Begin VB.Label lbl_organisms 
         AutoSize        =   -1  'True
         Caption         =   "Organisms Isolated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label lbl_bcomments 
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   75
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lbl_Epith 
         AutoSize        =   -1  'True
         Caption         =   "Epithelial Cells"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   74
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lbl_pus 
         AutoSize        =   -1  'True
         Caption         =   "Pus Cells"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3240
         TabIndex        =   73
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lbl_other 
         AutoSize        =   -1  'True
         Caption         =   "Other Findings"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6000
         TabIndex        =   72
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lbl_ziehl 
         AutoSize        =   -1  'True
         Caption         =   "Z. N. Stain"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9000
         TabIndex        =   71
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lbl_koh 
         AutoSize        =   -1  'True
         Caption         =   "K O H"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   70
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lbl_gpc 
         AutoSize        =   -1  'True
         Caption         =   "GPC"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3480
         TabIndex        =   69
         Top             =   720
         Width           =   390
      End
      Begin VB.Label lbl_gnc 
         AutoSize        =   -1  'True
         Caption         =   "GNC"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6600
         TabIndex        =   68
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lbl_gpb 
         AutoSize        =   -1  'True
         Caption         =   "GPB"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9480
         TabIndex        =   67
         Top             =   720
         Width           =   390
      End
      Begin VB.Label lbl_GNB 
         AutoSize        =   -1  'True
         Caption         =   "GNB"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1080
         TabIndex        =   66
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lbl_india 
         Caption         =   "India Ink Preparation for Cryptococcus"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3960
         TabIndex        =   65
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pathogen Grown"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label lbl_report 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   63
         Top             =   3240
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame common 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      Begin VB.TextBox txt_microno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9720
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         TabIndex        =   18
         Text            =   "cmb_reqno"
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   17
         Top             =   960
         Width           =   6255
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7920
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9720
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "micro_bact.frx":1927
         Left            =   6360
         List            =   "micro_bact.frx":1929
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9720
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10800
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox CMB_PROC 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   3
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "&EMR"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txt_spm 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   1
         Top             =   2040
         Width           =   2055
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   480
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
      Begin VB.Label lbl_LABID 
         AutoSize        =   -1  'True
         Caption         =   "Micro No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9480
         TabIndex        =   39
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   38
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8760
         TabIndex        =   37
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   36
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   35
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9240
         TabIndex        =   33
         Top             =   840
         Width           =   330
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   32
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5640
         TabIndex        =   30
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   29
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label lbl_ward 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7440
         TabIndex        =   28
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9000
         TabIndex        =   27
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5520
         TabIndex        =   26
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl_proc 
         AutoSize        =   -1  'True
         Caption         =   "Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   25
         Top             =   2040
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4680
         TabIndex        =   24
         Top             =   1680
         Width           =   2130
      End
      Begin VB.Label Label13 
         Caption         =   "Specimen"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   2040
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport CRYSTALREPORT1 
      Left            =   2040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9000
      TabIndex        =   106
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9000
      TabIndex        =   105
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_microbiology 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - MICROBIOLOGY "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   450
      Left            =   3840
      TabIndex        =   104
      Top             =   360
      Width           =   3900
   End
   Begin VB.Label lblicu 
      Caption         =   "ICU Patient"
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
      Height          =   375
      Left            =   0
      TabIndex        =   103
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "TATA MEMORIAL HOSPITAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   102
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label8 
      Caption         =   "04/12/08"
      Height          =   255
      Left            =   240
      TabIndex        =   101
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "micro_bact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim callstatus       As Boolean
Dim srno                As Integer

Dim VAR_REQNNO As String
Dim objWord             As Word.Application
Dim objDoc              As Word.Document
'for selecting antibiotics
Dim sel_antiqr          As rdo.rdoQuery
Dim sel_moadlityqr      As rdo.rdoQuery
Dim sel_deptqr          As rdo.rdoQuery
Dim req_qr              As rdo.rdoQuery
Dim SEL_detail_qr       As rdo.rdoQuery ' for selecting details of patient
Dim SELRREQUISITN_QR    As rdo.rdoQuery
Dim sel_proccode_qr     As rdo.rdoQuery ' for selecting procedure code
Dim sel_submodali_qr    As rdo.rdoQuery ' for selecting submodality
Dim REFDR_QR            As rdo.rdoQuery ' for refered by doctor not in used
Dim sel_bed_qr          As rdo.rdoQuery ' to select bed and ward
Dim SEL_MODALITY_QR     As rdo.rdoQuery
Dim sel_forupd2_qr      As rdo.rdoQuery ' for antibiotic
Dim sel_forupd1_qr      As rdo.rdoQuery ' toget data from table for bactriology
Dim sel_forupd3_qr      As rdo.rdoQuery ' to get data from table for urine
Dim sel_forupd4_qr      As rdo.rdoQuery ' to get data from table for stool
Dim sel_forupd5_qr      As rdo.rdoQuery ' to get data from table for other
Dim sel_forupd6_qr      As rdo.rdoQuery ' to get data from mycobacteriology table
Dim sel_micro_qr        As rdo.rdoQuery ' to  get micro number
Dim sel_ris_qr          As rdo.rdoQuery ' to get status of commit flag
Dim sel_FINAL_qr        As rdo.rdoQuery ' to get status of commit flag
Dim sel_ENTER_qr        As rdo.rdoQuery ' to get ENTER BY USER LIST
Dim SEL_dreq_staff_qr   As rdo.rdoQuery ' to get staff detail
Dim upd_12_qr           As rdo.rdoQuery ' to update record in dmbaanti table
Dim ins_0_qr            As rdo.rdoQuery ' insert query
Dim ins_1_qr            As rdo.rdoQuery
Dim ins_2_qr            As rdo.rdoQuery
Dim ins_4_qr            As rdo.rdoQuery
Dim ins_5_qr            As rdo.rdoQuery
Dim ins_6_qr            As rdo.rdoQuery 'insert in mycobacteriology table
'Dim ins_7_qr            As RDO.rdoQuery 'insert in DMURE TABLE AS DUMMY REQNO
'Dim ins_8_qr            As RDO.rdoQuery 'insert in DMSRE TABLE AS DUMMY REQNO
Dim upd_0_qr            As rdo.rdoQuery 'update query
Dim upd_1_qr            As rdo.rdoQuery
Dim upd_2_qr            As rdo.rdoQuery
Dim upd_4_qr            As rdo.rdoQuery
Dim upd_5_qr            As rdo.rdoQuery
Dim upd_6_qr            As rdo.rdoQuery
Dim upd_7_qr            As rdo.rdoQuery 'for updating mycobacteriology
Dim upd_8_qr            As rdo.rdoQuery 'for updating rislib about enterby
Dim upd_11_qr           As rdo.rdoQuery 'for setting requistion status
Dim del_dmbaanti_qr     As rdo.rdoQuery 'for deleting from dmbaanti table if pathogen is not grown
Dim del_dmbaanti_qr1    As rdo.rdoQuery
Dim i                   As Integer      'USE FOR SELECTING INDEX OF MICROBIOLOGY in combo
Dim sel_antirs          As rdo.rdoResultset
Dim sel_modalityrs      As rdo.rdoResultset
Dim sel_deptrs          As rdo.rdoResultset
Dim SEL_MODALITY_RS     As rdo.rdoResultset
Dim req_rs              As rdo.rdoResultset
Dim SEL_detail_rs       As rdo.rdoResultset
'******** Shraddha 26/11/08
Dim SELSPONANTIBRS As rdoResultset
Dim SELSPONANTIBQR As rdoQuery
'********
Dim SELRREQUISITN_RS    As rdo.rdoResultset
Dim sel_proccode_rs     As rdo.rdoResultset 'for selecting procedure code
Dim sel_submodali_rs    As rdo.rdoResultset 'for selecting submodality
Dim REFDR_RS            As rdo.rdoResultset 'not in used
Dim sel_bed_rs          As rdo.rdoResultset
Dim sel_forupd1_rs      As rdo.rdoResultset
Dim sel_forupd2_rs      As rdo.rdoResultset
Dim sel_forupd3_rs      As rdo.rdoResultset
Dim sel_forupd4_rs      As rdo.rdoResultset
Dim sel_forupd5_rs      As rdo.rdoResultset
Dim sel_forupd6_rs      As rdo.rdoResultset
Dim sel_micro_rs        As rdo.rdoResultset ' to  get micro number
Dim sel_ris_rs          As rdo.rdoResultset ' to  get micro number
Dim sel_ENTER_rs        As rdo.rdoResultset ' to  get ENTER BY USER LIST
Dim sel_FINAL_rs        As rdo.rdoResultset ' to  get FINAL BY USER LIST
Dim SEL_dreq_staff_rs   As rdo.rdoResultset ' to get staff detail
Dim S1                  As String ' FOR DEPTCODE
Dim S2                  As String ' FOR MODALITYCODE
Dim mod_reqno           As String 'used to store reqno by calling req_chglen method
Dim new_flag            As Boolean
Dim mod_flag            As Boolean
Dim VIEW_FLAG           As Boolean
Dim no_req              As Boolean 'to check before printing report whether req is existing or not
'here
''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean

''*******ICU**********


Dim GZZREQPATH_QR As rdoQuery
Dim GZZREQPATH_RS As rdoResultset

'for addendum
Dim INSDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUMFLAG As Boolean
Dim ADDENDUMFALG As Boolean
Dim d_grossdesc As Document
Dim W_grossdes As Application
Dim UPDRTIMEC_QR As rdoQuery
'''''''''''

' FOR CHECKING RECORDS PRESENCE DURING MODIFY
''********************************
Dim sel_formod2_qr As rdoQuery
Dim sel_formod2_rs As rdoResultset
Dim sel_formod2_flg As Boolean
''********************************

''******SECURITY***********

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String
Dim REPWTRNRS As rdoResultset
Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset
Dim REPWTRN_QR As rdoQuery
Dim txt_c_modflg As Boolean
''******SECURITY***********

'''DECOMMIT
Dim UPDRREQ_QR As rdoQuery
Dim SELDMBSEMCOM_QR As rdoQuery
Dim SELDMBSEMCOM_RS As rdoResultset
Dim SELDMBFLAG As Boolean
Dim CHECKREC_QR As rdoQuery
Dim CHECKREC_RS As rdoResultset
Dim CHECKFLAG As Boolean

''KANCHAN

Dim retantibodies_qr As rdoQuery
Dim RETRESULT_RS As rdoResultset
Dim REPORT_QR As rdoQuery
Dim REPORT_RS As rdoResultset
Dim REPORTSTR As String
Dim REPORTFLAG As Boolean
Dim UPDATERTIMELOG As rdoQuery

Dim INSINDMBANTI_QR As rdoQuery
Dim COL1 As Integer
Dim row1 As Integer
Dim MODE As String
Dim GRIDSCROLL As Boolean

Dim DEPT_CODE As String
Dim MOD_CODE As String
Dim modtxt As String
Dim ind As Integer
Dim NEWF As Boolean
Private Sub cmb_addendumby_Click()
    cmd_save.Enabled = True
    ADDENDUMFALG = True
End Sub

Private Sub cmb_antibiotic_Click()
            grid_dmbanti.TextMatrix(grid_dmbanti.row, 0) = Trim(cmb_antibiotic.Text)

End Sub
Function IsRepeated(antinm As String) As Boolean
    IsRepeated = True
    Dim i As Integer
    With grid_dmbanti
        If .Rows <= 2 Then
            IsRepeated = False
            Exit Function
        Else
            For i = 1 To .Rows - 2
                If .TextMatrix(i, 0) = Trim(antinm) And .row <> i Then
                    IsRepeated = True
                    MsgBox "This antibiotic is already entered", vbCritical
                    
                    Exit Function
                Else
                    IsRepeated = False
                End If
            Next i
        End If
    End With
End Function

Private Sub remoanti()
Dim rowcn As Integer
Dim cn As Integer
rowcn = grid_dmbanti.Rows - 1

For cn = 1 To rowcn
    ind = 0
    cmb_antibiotic.Enabled = True
    cmb_antibiotic.Text = Trim(grid_dmbanti.TextMatrix(cn, 0))
    ind = cmb_antibiotic.ListIndex
    cmb_antibiotic.RemoveItem (ind)
Next cn

End Sub

Private Sub cmb_antibiotic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmb_antibiotic.Text = "" Then
        MsgBox "Select Antibiotic", vbOKOnly
        
        cmb_antibiotic.SetFocus
    Else
        Call movecontrols(, 1)
            Call GridControls("VISIBLE", 1)
    End If
End If
End Sub
Public Sub MYDELAY(inttime As String)

Dim sttime
Dim endtime1 As Long
Dim TESTTIME As Long

endtime1 = Timer + inttime
TESTTIME = Timer

Do While Timer < endtime1
    DoEvents
Loop
End Sub

Private Sub cmb_antibiotic_LostFocus()

If Len(Trim(cmb_antibiotic.Text)) > 0 Then
    ind = 0
    ind = cmb_antibiotic.ListIndex
    If ind = -1 Then
    
    Else
        cmb_antibiotic.RemoveItem (ind)
    End If
End If
End Sub

Private Sub cmb_modality_Click()
On Error GoTo MSG
    i = InStr(cmb_modality.Text, "$")
    txt_modcode.Text = Mid(cmb_modality.Text, i + 1, 2)
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_modality_LostFocus()
 If cmb_modality.Text = "" Then
        MsgBox "Select Modality First", vbOKOnly
    End If
End Sub

Private Sub cmb_organisms_LOSTFOCUS()
On Error GoTo MSG
    If new_flag Then
        sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
        sel_antiqr(1) = Trim(cmb_organisms.Text)
        Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
        If Not sel_antirs.EOF Then
                grid_c_anti.clear
                grid_c_anti.Rows = 1
                grid_c_anti.Cols = 2
                grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
                While Not sel_antirs.EOF
                  grid_c_anti.AddItem sel_antirs("antibdesc")
                  sel_antirs.MoveNext
                Wend
                    grid_c_anti.row = 1
                    grid_c_anti.Col = 1
                    txt_c_1.Width = grid_c_anti.CellWidth
                    txt_c_1.Height = grid_c_anti.CellHeight
                    txt_c_1.Left = grid_c_anti.CellLeft + grid_c_anti.Left
                    txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
                    txt_c_1.Visible = True
                    sel_antirs.Close
        Else
            MsgBox "Use Master Screen to Enter Antimicrobial", vbOKOnly
            sel_antirs.Close
            Exit Sub
        End If
     ElseIf mod_flag Then
        If Not sel_forupd2_rs.EOF Then
                grid_c_anti.Rows = 1
                grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
                While Not sel_forupd2_rs.EOF
                    grid_c_anti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
                    sel_forupd2_rs.MoveNext
                Wend
                sel_forupd2_rs.Close
         Else
                sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
                sel_antiqr(1) = Trim(cmb_organisms.Text)
                Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
                If Not sel_antirs.EOF Then
                    grid_c_anti.Rows = 1
                    grid_c_anti.Cols = 2
                    grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
                    While Not sel_antirs.EOF
                        grid_c_anti.AddItem sel_antirs("antibdesc")
                        sel_antirs.MoveNext
                    Wend
                        grid_c_anti.row = 1
                        grid_c_anti.Col = 1
                        txt_c_1.Width = grid_c_anti.CellWidth
                        txt_c_1.Height = grid_c_anti.CellHeight
                        txt_c_1.Left = grid_c_anti.CellLeft + grid_c_anti.Left
                        txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
                        txt_c_1.Visible = True
                        sel_antirs.Close
                End If
                sel_forupd2_rs.Close
         End If
     End If
     Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_patho_click()
On Error GoTo MSG
If cmb_patho.Text = "Yes" Then
    lbl_report.Visible = False
    txt_report.Visible = False
    lbl_organisms.Visible = True
    txt_organ1.Visible = True
    TXT_ORGAN2.Visible = True
    TXT_ORGAN3.Visible = True
    lbl_1.Visible = True
    lbl_2.Visible = True
    lbl_3.Visible = True
    grid_dmbanti.Visible = True
    lbl_bcomments.Left = 600
    lbl_bcomments.Top = 3720
    txt_comment.Left = 1560
    txt_comment.Top = 3720
Else 'if selected no then
    lbl_report.Visible = True
    txt_report.Visible = True
    lbl_organisms.Visible = False
    txt_organ1.Visible = False
    TXT_ORGAN2.Visible = False
    TXT_ORGAN3.Visible = False
    txt_1.Visible = False
    txt_2.Visible = False
    txt_3.Visible = False
    lbl_1.Visible = False
    lbl_2.Visible = False
    lbl_3.Visible = False
    'grid_anti.Visible = False
    grid_dmbanti.Visible = False
    lbl_report.Left = 960
    lbl_report.Top = 2040
    txt_report.Left = 1560
    txt_report.Top = 1920
    txt_comment.Left = 1560
    txt_comment.Top = 2640
    lbl_bcomments.Left = 600
    lbl_bcomments.Top = 2640
    If mod_flag Then
        'clear previous data of dmbaanti table
        txt_organ1.Text = ""
        TXT_ORGAN2.Text = ""
        TXT_ORGAN3.Text = ""
        del_dmbaanti_qr(0) = mod_reqno
        del_dmbaanti_qr.Execute
        grid_dmbanti.Rows = 1
        grid_dmbanti.Cols = 3
    End If
End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_reqno_Click()
On Error GoTo MSG

clearlowerpanel2
'get what req is selected and depending on that change lower frame.
'FIRST GET MODALITY
' get detail of patient
If cmb_reqno.ListIndex > 0 And (cmb_reqno.Text) = "" Or Len(cmb_reqno.Text) < 1 Then
    MsgBox "Select Requisition First", vbOKOnly
    btnenabl "INI"
    Exit Sub
End If
Dim a As Integer
a = InStr(1, cmb_reqno, "$")
mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, a - 1)) 'put change value to this var so that it used later on

'GETFINALLIST
If Opt_patient.Value Then
    If new_flag Then
        ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If SELICU_FLAG = False Then
            SELICU_QR(0) = mod_reqno ' VAR_REQNNO
            Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
            SELICU_FLAG = True
        Else
            SELICU_QR(0) = mod_reqno 'VAR_REQNNO
            SELICU_RS.MoreResults
            SELICU_RS.Requery
        End If
        
        If Not SELICU_RS.EOF Then
            If Trim(SELICU_RS("ICDS2")) = "YES" Then
                lblicu.Visible = True
            Else
                lblicu.Visible = False
            End If
        End If
            ''******ICU****************
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
       
        INSERT
                 
    ElseIf VIEW_FLAG Then
       
        ''******ICU****************
        VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If SELICU_FLAG = False Then
            SELICU_QR(0) = mod_reqno   'VAR_REQNNO
            Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
            SELICU_FLAG = True
        Else
            SELICU_QR(0) = VAR_REQNNO
            SELICU_RS.MoreResults
            SELICU_RS.Requery
        End If
        If Not SELICU_RS.EOF Then
            If Trim(SELICU_RS("ICDS2")) = "YES" Then
                lblicu.Visible = True
            Else
                lblicu.Visible = False
            End If
        End If
        ''******ICU****************
        RETRIVE
            
    End If
    
ElseIf opt_staff.Value Then 'FOR STAFF MEMBER
    'STAFF
    
    If VIEW_FLAG Then
    
    STAFF
    End If
End If 'END OF TOP IF
no_req = True

Exit Sub
MSG:
    MsgBox Error(Err)
End Sub

Private Sub cmd_back_Click()
btnenabl ("INI")
clearlowerpanel1
clearlowerpanel2
clearlowerpanel3
clearlowerpanel4
clearupperpanel
set_visible "All"
DIS_MAIN.show
Me.Hide
End Sub

Private Sub cmd_cancel_Click()

DEPT_CODE = Trim(txt_deptcode.Text)
MOD_CODE = Trim(txt_modcode.Text)
modtxt = Trim(cmb_modality.Text)
btnenabl ("INI")
clearupperpanel
clearlowerpanel1
clearlowerpanel2
clearlowerpanel3
clearlowerpanel4
set_flag "ALL"
txt_1.Visible = False
txt_2.Visible = False
txt_3.Visible = False
set_visible "All"
lbl_bcomments.Left = 120
lbl_bcomments.Top = 3720
txt_comment.Left = 1080
txt_comment.Top = 3720
cmb_dept.Enabled = True
cmb_modality.Enabled = True
cmb_dept.Enabled = True
cmb_reqno.Enabled = True
MaskEdBox1.Enabled = True
txt_c_modflg = False
set_flag "Ni"
cmb_antibiotic.Visible = False
MODE = ""
End Sub

Private Sub cmd_commit_Click()
On Error GoTo MSG
Dim i As Integer
i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If i = vbNo Then
    Exit Sub
End If
If i = vbYes Then
On Error GoTo MSG:
''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_final.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
    Dim STR11 As String
    STR11 = CURUSRDET_RS("USERNAME")
        
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
            cmd_commit.SetFocus
            Exit Sub
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
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
    upd_6_qr(0) = Left(Trim(cmb_final.Text), 30)
    upd_6_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_6_qr(2) = "F"
    upd_6_qr(3) = "Y"
    upd_6_qr(4) = mod_reqno
    upd_6_qr.Execute
    rdb.CommitTrans
 
    UPDATERTIMELOG(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDATERTIMELOG(1) = Format(txt_time.Text, "HH:MM:SS")
    UPDATERTIMELOG(2) = mod_reqno
    UPDATERTIMELOG.Execute
 
    btnenabl "INI"
    MsgBox "Record is Committed Successfully."
End If
cmd_print.Enabled = True
Exit Sub

MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub CMD_DECOMMIT_Click()

Dim b As Integer
b = InStr(1, cmb_reqno, "$")
mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, b - 1)) 'put change value to this var so that it used later on

UPDRREQ_QR(0) = "T"
UPDRREQ_QR(1) = Null
UPDRREQ_QR(2) = mod_reqno
UPDRREQ_QR.Execute

cmb_reqno_Click
End Sub

Private Sub cmd_modify_Click()
btnenabl ("NEW")
MODE = "MODIFY"
cmb_antibiotic.clear
GET_ANTIBIOTIC

If Len(Trim(MaskEdBox1.Text)) = 0 Then
    MsgBox "Case Number should not be empty!", vbCritical
    MaskEdBox1.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_reqno.Text)) = 0 Then
    MsgBox "Requisition Number should not be empty!", vbCritical
    If cmb_reqno.Enabled = True Then
        cmb_reqno.SetFocus
        SendKeys "^{f4}"
    End If
    Exit Sub
End If
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
If REFDR_RS.RowCount > 0 Then
    If Trim(REFDR_RS("reqnstat")) = "T" And Trim(REFDR_RS("COMMITFLAG")) = "F" Then
    ADDENDUMFALG = True
    End If
End If
sel_ris_qr(0) = mod_reqno
Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
While Not sel_ris_rs.EOF
    If Trim(sel_ris_rs(0)) = "Y" Then
        MsgBox "Sorry Record is Already Commited", vbOKOnly
        Exit Sub
    Else
        set_flag "MOD"
    End If
    sel_ris_rs.MoveNext
Wend
enable_field
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
NEWF = True
MODE = "NEW"
btnenabl ("NEW")
    cmb_enteryby.Text = cmb_enteryby.List(0)
    cmb_enteryby.Enabled = False
    For L1 = 0 To cmb_final.ListCount
    If Trim(cmb_enteryby.List(0)) = cmb_final.List(L1) Then
        cmb_final.Text = cmb_final.List(L1)
    End If
    Next L1
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
    Wend
    For i = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(i)) = "MICROBIOLOGY$GZZ" Then
                cmb_dept.ListIndex = i
        End If
    Next i
     If Trim(cmb_dept) = "" Then
        Exit Sub
    End If
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
sel_deptrs.Close
set_flag "NEW"
Opt_patient.Value = True
cmd_modify.Enabled = False
grid_dmbanti.Rows = 2

End Sub

Private Sub cmd_print_Click()
'On Error GoTo MSG

Dim prepath As String
Dim I1 As Integer
DoEvents
'check
For I1 = 0 To 11
    CRYSTALREPORT1.Formulas(I1) = ""
Next I1
If no_req Then 'check for entered selected req no.
    sel_micro_qr(0) = mod_reqno
    Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
    While Not sel_micro_rs.EOF
        If Trim(sel_micro_rs("pathno")) <> Trim(txt_microno.Text) Then
             prepath = prepath & sel_micro_rs("pathno")
        End If
        sel_micro_rs.MoveNext
    Wend
    sel_ris_qr(0) = mod_reqno
    Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
    STR = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    If ADDENDUMFALG = False Then
        If Trim(txt_modcode.Text) = "BA" Then 'for case no. 1(bacteriology)
            If cmb_patho.Text = "Yes" Then
                CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\bacteriology_report.rpt"
                CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                CRYSTALREPORT1.Formulas(3) = "USRID = '" & STR & "'"
                CRYSTALREPORT1.Formulas(4) = "patho = '" & "Yes" & "'"
                'SELECT FOR ORGANISM ISOLATED
                If Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
                    CRYSTALREPORT1.Formulas(5) = "ORGAN1 = '" & "Yes" & "'"
                ElseIf Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) > 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
                    CRYSTALREPORT1.Formulas(5) = "ORGAN2 = '" & "Yes" & "'"
                ElseIf Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) > 0 And Len(Trim(TXT_ORGAN3.Text)) > 0 Then
                    CRYSTALREPORT1.Formulas(5) = "ORGAN3 = '" & "Yes" & "'"
                End If
           Else
                CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\patho_no_report.rpt"
                CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                CRYSTALREPORT1.Formulas(3) = "USRID = '" & STR & "'"
           End If
        
        
        Else   ' for other case
           
        End If
   
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CRYSTALREPORT1.Formulas(6) = "pathnumber='" & txt_microno.Text & "'" '17/11/08 7->6
    Else
        'If ADDENDUMFALG = False Then
        CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
        CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\DISMBADD.rpt"
        CRYSTALREPORT1.Formulas(1) = "profin = 'MICROBIOLOGY REPORT'"
        CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
        CRYSTALREPORT1.Formulas(3) = "USRID = '" & STR & "'"
        CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & mod_reqno & "'"
        CRYSTALREPORT1.Formulas(7) = "pathnumber='" & txt_microno.Text & "'"
        CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
    End If
    
    If Not Trim(txt_ward.Text) = "" Then
         If InStr(1, txt_ward.Text, "WARD") = 0 Then
             CRYSTALREPORT1.Formulas(7) = "ward='" & txt_ward.Text & " /" & txt_bedno.Text & "'"
         Else
             CRYSTALREPORT1.Formulas(7) = "ward='" & Trim(Mid(txt_ward.Text, 1, InStr(1, txt_ward.Text, "WARD") - 1)) & "'"
         End If
    Else
            CRYSTALREPORT1.Formulas(7) = "ward="
    End If
    CRYSTALREPORT1.Formulas(8) = "NAM='" & txt_name.Text & "'"
    CRYSTALREPORT1.Formulas(9) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
    If Opt_patient.Value Then
        CRYSTALREPORT1.Formulas(10) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
    Else
        CRYSTALREPORT1.Formulas(10) = "ST='Staff /" & txt_status.Text & "'"
    End If
     
End If
    
    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CRYSTALREPORT1.WindowState = crptMaximized
    CRYSTALREPORT1.Action = 1

    Exit Sub
MSG:
    MsgBox Error(Err)
End Sub

Private Sub cmd_save_Click()
'ADDED BY MILIND ON 16-10-06 FOR SEPARATING REQNO FROM LABNO.
Dim a As Integer
a = InStr(1, cmb_reqno, "$")
mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, a - 1))
'UPTO HERE

If new_flag Then    'if new record then use insert query
    
    If Len(Trim(cmb_enteryby)) = 0 Then
        MsgBox "The Enter by field is EMPTY.", vbCritical
        cmb_enteryby.SetFocus
        Exit Sub
    End If

    If Len(Trim(cmb_final)) = 0 Then
        MsgBox "The Finalized by field is EMPTY.", vbCritical
        cmb_final.SetFocus
        Exit Sub
    End If
     
    While InStr(txt_addendum, Chr(9)) > 0
            
        Dim beforst As String, afterst As String
        Dim K As Long
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
        
  
    If new_flag Then
    
      
    End If
    

    btnenabl "SAVE"
    new_flag = False
        
    'check which test is performed
    If (frm_bacteria.Visible) Then 'Bacteriology 13/11/08 Shraddha
        'insert in DMSPONANIB AND DMPRISMER
        insert_bacteriology
        Exit Sub
     End If
        
    If Len(Trim(txt_addendum.Text)) > 600 Then
        MsgBox "Addendum Max. of 600 Characters Only.", vbInformation
        txt_addendum.SetFocus
        Exit Sub
    End If
End If
    'cmd_commit.Enabled = True
    If mod_flag Then 'else if modifying a record then use update query
        If Len(Trim(cmb_enteryby)) = 0 Then
            MsgBox "The Enter by field is EMPTY.", vbCritical
            cmb_enteryby.SetFocus
        Exit Sub
        End If
    
    If Len(Trim(cmb_final)) = 0 Then
        MsgBox "The Finalized by field is EMPTY.", vbCritical
        cmb_final.SetFocus
        Exit Sub
    End If
    
    mod_flag = False
    btnenabl "SAVE"
    'check which test is performed
    
    Dim STR2 As String
    
    If frm_bacteria.Visible Then 'for Bacteriology
        'update in DMSPONANIB AND DMPRISMER
        'before update get result from database server and assign them to fields by calling getresult
        upd_1_qr(0) = txt_epith.Text
        upd_1_qr(1) = txt_pus.Text
        upd_1_qr(2) = txt_other.Text
        'upd_1_qr(3) = txt_ziehl.Text
        upd_1_qr(3) = txt_ziehl.Text
        upd_1_qr(4) = txt_koh.Text
        upd_1_qr(5) = cmb_gpc.Text
        upd_1_qr(6) = cmb_gnc.Text
        upd_1_qr(7) = cmb_gpb.Text
        upd_1_qr(8) = cmb_gnb.Text
        upd_1_qr(9) = cmb_india.Text
        If Len(txt_comment.Text) > 0 Then
            upd_1_qr(14) = Null
            upd_1_qr(14) = txt_comment.Text
        Else
            upd_1_qr(14) = Null
            upd_1_qr(14) = " "
        End If
        upd_1_qr(15) = mod_reqno
        
        If cmb_patho.Text = "No" Then     'patho
            upd_1_qr(10) = Null
            upd_1_qr(11) = Null
            upd_1_qr(12) = Null
            
            If Len(txt_report.Text) > 0 Then
                upd_1_qr(13) = Null
                upd_1_qr(13) = txt_report.Text
            Else
                upd_1_qr(13) = Null
                upd_1_qr(13) = " "
            End If
            
            upd_1_qr.Execute
        Else
            If Len(txt_organ1.Text) > 0 Then
                upd_1_qr(10) = Null
                upd_1_qr(10) = txt_organ1.Text
            Else
                upd_1_qr(10) = Null
            End If
            If Len(TXT_ORGAN2.Text) > 0 Then
                upd_1_qr(11) = Null
                upd_1_qr(11) = TXT_ORGAN2.Text
            Else
                upd_1_qr(11) = Null
            End If
            If Len(TXT_ORGAN3.Text) > 0 Then
                upd_1_qr(12) = Null
                upd_1_qr(12) = TXT_ORGAN3.Text
            Else
                upd_1_qr(12) = Null
            End If
            upd_1_qr(13) = Null
            upd_1_qr.Execute
                    
'''Done By Kanchan
            UPDATE_DMBAANTI
        End If 'of patho
        update_rislib
        MsgBox "Record Updated Successfully", vbOKOnly
        Exit Sub
  
    End If
End If
If ADDENDUMFALG = True Then
                ''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_final.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
            
    If CURUSRDET_RS.RowCount > 0 Then
    
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
                      ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! U'r not an Authorized person to write an Addendum !!", vbInformation, "Pathlogy"
            cmd_save.SetFocus
            Exit Sub
        End If
        If CURUSRDET_RS("ADDENDUM") <> "YES" Then
            MsgBox "Sorry! You are not authorised to Write Addendum of this Modality"
            cmd_save.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
    End If
        ''********SECURITY***************************
                    
    If SELDADDENDUMFLAG = False Then
        SELDADDENDUM_QR(0) = mod_reqno
        Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
        SELDADDENDUMFLAG = True
    Else
        SELDADDENDUM_QR(0) = mod_reqno
        SELDADDENDUM_RS.MoreResults
        SELDADDENDUM_RS.Requery
    End If
            
    If SELDADDENDUM_RS.RowCount > 0 Then
        INSDADDENDUM_QR(0) = mod_reqno
        INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
    Else
        INSDADDENDUM_QR(0) = mod_reqno
        INSDADDENDUM_QR(1) = "1"
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
                    
        cmd_save.Enabled = False
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
                    
        MsgBox "Record Updated Successfully!!!!!"
                    
    End If
End If
                        
'once inserted or modified then clear all field
MODE = ""
'End If
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:

btnenabl ("VIEW")

Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    'get the result
    While Not sel_deptrs.EOF
            cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
            sel_deptrs.MoveNext
    Wend
    
sel_deptrs.Close
If Trim(cmb_dept) = "" Then
        For i = 0 To (cmb_dept.ListCount - 1)
            If Trim(cmb_dept.List(i)) = "MICROBIOLOGY$GZZ" Then
                cmb_dept.ListIndex = i
            End If
        Next i
End If

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
    
set_flag "VIEW"

Opt_patient.Value = True
cmd_print.Enabled = True
cmb_dept.SetFocus
Exit Sub
MSG:
    MsgBox Err.Description
End Sub


Private Sub Command1_Click()
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    Dim VAR1
    'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
    VAR1 = MaskEdBox1.FormattedText
    VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
    AppActivate VAR
Else
    MsgBox "Please Enter The Case No."
End If
End Sub

Private Sub Command2_Click()
micro_myco.show
End Sub

Private Sub Form_Load()
NEWF = False
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
set_visible "All"
no_req = False

'addendum
txt_addendum.Enabled = False
cmb_addendumby.Enabled = False
'initialize all combo with default value
initialize_combo
cmb_india.AddItem "Negative"
cmb_india.AddItem "Positive"
cmb_india.AddItem " "

'***************************************************************************************
'                       QUERY DEFINITION
'***************************************************************************************
If MBREPORTFALG = False Then
    MBREPORTFALG = True

    '*********SECURITY**********************************

    STR = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
    Set USRNAME_QR = rdb.CreateQuery("USRNAME1MICRO", STR)
    USRNAME_QR(0) = UID
    Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
    If USRNAME_RS.RowCount > 0 Then
        LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
        LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
        LOGINUSRID = Trim(USRNAME_RS("usr_id"))
        cmb_enteryby.Text = (LOGINUSRNAME)
        cmb_enteryby.AddItem LOGINUSRNAME

        cmb_enteryby.Enabled = False
    Else
        MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
        cmd_cancel_Click
        Exit Sub
    End If

    STR = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
    Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STRFFF", STR)
    If Trim(USRDEPT) = "ALL" Then
        REPWTRN_QR(0) = "GZZ"
    Else
        REPWTRN_QR(0) = Trim(USRDEPT)
    End If
    Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
        cmb_final.clear

        REPWTRNRS.MoveFirst
        While Not REPWTRNRS.EOF
            If (REPWTRNRS("USRTYPE") >= 8 Or REPWTRNRS("USRTYPE") = 6) Then
                cmb_final.AddItem Trim(REPWTRNRS("USERNAME"))
                cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            End If
            REPWTRNRS.MoveNext
        Wend
    Else
        MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
        cmd_cancel_Click
        Exit Sub
    End If
    STR = "select * from DISLIB.Dislnkusr where username = ?"
    Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11MICRO", STR)
    CURUSRDET_QR(0) = "ABC"
    Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

    '*********SECURITY**********************************

    ''********ICU***************
    STR = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
    Set SELICU_QR = rdb.CreateQuery("SELICU_QR115", STR)

    STR = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
    Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR115", STR)

    ''********ICU***************

    STR = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
    Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQ11111", STR)
    'for addendum

    STR = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
    Set UPDATERTIMELOG = rdb.CreateQuery("UPDATERTIMELOG11", STR)
    
    STR = "INSERT INTO DISLIB.ADDENDUMMB(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
    Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QRCG", STR)
    
    STR = "SELECT  * FROM DISLIB.ADDENDUMMB WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QRCG", STR)

    'FOR ANTIBIOTICS
    STR = "select antibcode,antibdesc from dislib.dmicbantib WHERE MODALITY=? AND SUBMODALI=? ORDER BY ANTIBDESC"
    Set sel_antiqr = rdb.CreateQuery("micro_select1", STR)

    'for dept selection
    STR = "select dept_name, dept_code from PATOOLSLIB.ma106"
    Set sel_deptqr = rdb.CreateQuery("micro_select2", STR)

    'FOR MODALITY SELECTION
    STR = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
    Set sel_moadlityqr = rdb.CreateQuery("micro_SELECT3", STR)
   
    STR = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? and DEPT_CODE=? AND REQNSTAT='E' ORDER BY REQNDT DESC"
    Set req_qr = rdb.CreateQuery("micro_select4", STR)

    STR = "SELECT PATHNO FROM DISLIB.DREQUISITN WHERE REQNO=?"
    Set GZZREQPATH_QR = rdb.CreateQuery("GZZREQPATH_QR", STR)
    
    STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND  DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("micro_SELECT5", STR)

    'select other detail for patient
    STR = "select * from PATOOLSLIB.ta101 where case_no =?"
    Set SEL_detail_qr = rdb.CreateQuery("micro_select6", STR)
    
    '***** Shraddha Pandit 26/11/08
    STR = "SELECT * FrOM DISLIB.DMSPONANTI WHERE REQNO=?"
    Set SELSPONANTIBQR = rdb.CreateQuery("SELSPONANTIBQRMB", STR)
    '**************

    'for bed and ward retriving
    STR = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set sel_bed_qr = rdb.CreateQuery("micro_select7", STR)

    'for referred by dr.
    STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("micro_select8", STR)

    'FOR SELECTING WHICH PROCEDURE HE HAS SELECTED
    STR = "SELECT SER_CODE,SER_DET FROM PATOOLSLIB.MA101 WHERE SER_CODE IN(SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO= ?)"
    Set sel_proccode_qr = rdb.CreateQuery("micro_select9", STR)

    'TO GET MODALITY DESC.
    STR = "SELECT MODLITYDES,MODALITYCD FROM RISLIB.RMODALITY WHERE DEPT_CODE=? AND MODALITYCD=?"
    Set SEL_MODALITY_QR = rdb.CreateQuery("micro_SELECT10", STR)

    'to get submodality under mycobacteriology
    STR = "select DISTINCT(submodali) from dislib.dmicbantib where trim(modality)=?"
    Set sel_submodali_qr = rdb.CreateQuery("micro_select11", STR)

    'to get result before modify
    STR = "SELECT EPITHECELL,PUSCELL,OFINDING,ZNSTAIN,KOHAMT,GPC,GNC,GPB,GNB,IIPC,ORGSO1,ORGSO2,ORGSO3,REPORT,COMMENTS from  DISLIB.DMOPRISMER where reqno=?"
    Set sel_forupd1_qr = rdb.CreateQuery("micro_SELECT12", STR)
 
    'to get antibiotic with values
    STR = "SELECT * from  DISLIB.DMBAANTI where reqno=?"
    Set sel_forupd2_qr = rdb.CreateQuery("micro_SELECT13", STR)

    STR = "select * from DISLIB.DMBAANTI where reqno=? and ANTIBIOTIC=?"
    Set retantibodies_qr = rdb.CreateQuery("retantibiotic", STR)

    ' TO INSERT NEW RECORD IF RECORDS ARE NOT FOUND DURING MODIFY
    STR = "SELECT * from  DISLIB.DMBAANTI where reqno=? AND antibiotic=?"
    Set sel_formod2_qr = rdb.CreateQuery("micro_SELECTmod13", STR)

    'TO GET FOR VIEW OF URINE RESULT
    STR = "SELECT * from  DISLIB.DMURE where reqno=?"
    Set sel_forupd3_qr = rdb.CreateQuery("micro_SELECT14", STR)

    'TO GET FOR VIEW OF URINE RESULT
    STR = "SELECT * from  DISLIB.DMSRE where reqno=?"
    Set sel_forupd4_qr = rdb.CreateQuery("micro_SELECT15", STR)

    'TO GET FOR VIEW OF other RESULT
    STR = "SELECT * from  DISLIB.dmbsemcmo where reqno=?"
    Set sel_forupd5_qr = rdb.CreateQuery("micro_SELECT16", STR)

    'to get result of myco
    STR = "SELECT * from  DISLIB.dmycobacte where reqno=?"
    Set sel_forupd6_qr = rdb.CreateQuery("micro_SELECT17", STR)

    'to get micro no which is path no. and previous micro no. in report
    STR = "SELECT * from  DISLIB.Drequisitn where reqno=?"
    Set sel_micro_qr = rdb.CreateQuery("micro_SELECT18", STR)

    'to check status of commitflag
    STR = "select * from rislib.rrequisitn where reqnno=?"
    Set sel_ris_qr = rdb.CreateQuery("micro_select19", STR)

    'to get entered by name
    STR = "select * from dislib.dissec where enterby='YES'AND MODALITY='BA'"
    Set sel_ENTER_qr = rdb.CreateQuery("seldissecentr_qr0091", STR)

    'to ge fianl by list name
    STR = "select * from dislib.dissec where FINALby='YES' AND MODALITY='BA'"
    Set sel_FINAL_qr = rdb.CreateQuery("seldissecFIN_qr0091", STR)

    'for staff query
    STR = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SEL_dreq_staff_qr = rdb.CreateQuery("micro_select", STR)

    
    '###################################    INSERT QUERY    ##############################
    'insert into table for modality except 1 and 5 record
    'OVER
    STR = "insert into DISLIB.dmbsemcmo(reqno,refserv,testperfor,result)values(?,?,?,?)"
    Set ins_0_qr = rdb.CreateQuery("micro_insert0", STR)
    'insert record for madality 1
    'over
    STR = "insert into DISLIB.dmoprismer(reqno,epithecell,puscell,ofinding,znstain,kohamt,gpc,gnc,gpb,gnb,iipc,orgso1,orgso2,orgso3,REPORT,comments) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set ins_1_qr = rdb.CreateQuery("micro_insert1", STR)
    'for inserting effect of antibiotics for modality 1
    'over
    STR = "insert into DISLIB.dmbaanti(reqno,antibiotic,antibio1,antibio2,antibio3) values(?,?,?,?,?)"
    Set ins_2_qr = rdb.CreateQuery("micro_insert2", STR)

    'insert urine result
    'over
    STR = "insert into DISLIB.dmure (reqno,color,sp,ph,protein,glucose,ketones,bulirubin,urobilinog,blood,micexam,wbc,rbc,sqepcell,bacteria,casts,others,comments)values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set ins_4_qr = rdb.CreateQuery("micro_insert4", STR)
    'insert stool result
    'over
    STR = "insert into DISLIB.dmsre (reqno,color,consistenc,mucus,blood,parasties,occuttblod,ova,cysts,larvae,puscell,rbc,macophages,starch,fat,others,comments) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set ins_5_qr = rdb.CreateQuery("micro_insert5", STR)

    'insert into mycobacteriology table
    STR = "insert into DISLIB.dmycobacte (reqno,stain,organism,report,comment) values(?,?,?,?,?)"
    Set ins_6_qr = rdb.CreateQuery("micro_insert6", STR)
    
    '###################################    UPDATE QUERY    ##############################
    'update queries
    'update reqstat if he/she successfully entered data
    STR = "update rislib.rrequisitn set reqnstat='T' where caseno=? and reqnno=?"
    Set upd_11_qr = rdb.CreateQuery("micro_update0", STR)

    STR = "update dislib.dmbsemcmo set refserv=? ,testperfor=?,result=? where reqno=?"
    Set upd_0_qr = rdb.CreateQuery("micro_update1", STR)

    'UPDATE record for madality 1
    STR = "update dislib.dmoprismer set epithecell=? ,puscell=?,ofinding=?,znstain=?,kohamt=?,gpc=?,gnc=?,gpb=?,gnb=?,iipc=?,orgso1=?,orgso2=?,orgso3=?,REPORT=?,comments=? where reqno=? "
    Set upd_1_qr = rdb.CreateQuery("micro_update2", STR)

    'for UPDATE effect of antibiotics for modality 1
    STR = "update dislib.dmbaanti set antibio1=?,antibio2=?,antibio3=? where reqno=? and antibiotic=? "
    Set upd_2_qr = rdb.CreateQuery("micro_update3", STR)

    'for UPDATE effect of antibiotics for modality 1

    STR = "update dislib.dmbaanti set antibio1=?,antibio2=?,antibio3=?,antibiotic=?  where reqno=? "
    Set upd_12_qr = rdb.CreateQuery("micro_update13", STR)

    'update commit flag
    STR = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set upd_6_qr = rdb.CreateQuery("micro_update6", STR)

   
    'to update rislib for enterby and enter by
    STR = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
    Set upd_8_qr = rdb.CreateQuery("micro_update_8", STR)

    '''Done By KANCHAN

    STR = "delete from dislib.dmbaanti where reqno=?"
    Set del_dmbaanti_qr = rdb.CreateQuery("micro_delete1", STR)
    'delete for specific antibiotic
    STR = "delete from dislib.dmbaanti where reqno=? and antibiotic=?"
    Set del_dmbaanti_qr1 = rdb.CreateQuery("micro_delete2", STR)

    STR = "SELECT * FROM DISLIB.DMYCOBACTE WHERE REQNO=?"
    Set CHECKREC_QR = rdb.CreateQuery("CHECKRECORD", STR)

    STR = "SELECT REPORT FROM DISLIB.DMYCOBACTE WHERE REQNO=?"
    Set REPORT_QR = rdb.CreateQuery("REPORT", STR)

    ''DECOMMIT
    STR = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,COMMITFLAG=? WHERE REQNNO=?"
    Set UPDRREQ_QR = rdb.CreateQuery("UPDATERREQ", STR)

    STR = "SELECT * FROM DISLIB.DMBSEMCmo WHERE REQNO=?"
    Set SELDMBSEMCOM_QR = rdb.CreateQuery("SELDMBSEMCOM111", STR)

    'grid_anti.AddItem
End If
btnenabl ("INI")
new_flag = False
VIEW_FLAG = False
mod_flag = False
GRIDSCROLL = False

'set_position 'to set the position of text_field on grid.
End Sub

Private Sub set_visible(frame_name As String)
If frame_name = "All" Then
    frm_bacteria.Visible = False
ElseIf frame_name = "frm_bacteria" Then
    frm_bacteria.Visible = True
Else
    frm_bacteria.Visible = False
End If
End Sub

Private Sub grid_c_anti_Click()
If grid_c_anti.Col = 1 Or grid_c_anti.Col = 2 Then
    txt_c_1.Text = CStr(grid_c_anti.TextMatrix(grid_c_anti.row, 1))
    grid_c_anti.Col = 1
    txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
    If txt_c_modflg = True Then
        txt_c_1.Visible = True
    End If
End If
End Sub

Private Sub grid_dmbanti_Click()
GRIDSCROLL = False
If MODE = "MODIFY" Or MODE = "NEW" Then
    If (Len(Trim(txt_organ1.Text)) > 0 Or Len(Trim(TXT_ORGAN2.Text)) > 0 Or Len(Trim(TXT_ORGAN3.Text)) > 0) Then
        If grid_dmbanti.row = 0 Then
            MsgBox ("Please select proper location....!"), vbOKOnly
        Else
            If MODE = "MODIFY" Or MODE = "NEW" Then
                If grid_dmbanti.Col = 0 Then
                    cmb_antibiotic.Visible = True
                    cmb_antibiotic.Enabled = True
                    cmb_antibiotic.SetFocus
                ElseIf grid_dmbanti.Col = 1 Then
                    If Len(Trim(grid_dmbanti.TextMatrix(grid_dmbanti.row, 0))) > 0 Then
                        txt_1.Visible = True
                        txt_1.Enabled = True
                        txt_1.SetFocus
                    Else
                        MsgBox ("Please select antibiotic first....!"), vbOKOnly
                        cmb_antibiotic.Visible = True
                        cmb_antibiotic.Enabled = True
                        cmb_antibiotic.SetFocus
                        Call GridControls("VISIBLE", 0)
                        Call movecontrols(, 0)
                        Exit Sub
                    End If
                ElseIf grid_dmbanti.Col = 2 Then
                    txt_2.Visible = True
                    txt_2.Enabled = True
                    txt_2.SetFocus
                ElseIf grid_dmbanti.Col = 3 Then
                    txt_3.Visible = True
                    txt_3.Enabled = True
                    txt_3.SetFocus
                End If
                Call GridControls("VISIBLE", grid_dmbanti.Col)
                Call movecontrols(, grid_dmbanti.Col)
            Else
                GridControls ("INVISIBLE")
            End If
        End If
    Else
        MsgBox ("Atleast 1 organism is needed...!"), vbOKOnly
        txt_organ1.SetFocus
    End If ''''CHECK FOR ORGANISM
End If   ''''MODE
End Sub

Private Sub grid_dmbanti_Scroll()
    If GRIDSCROLL = False Then
        GridControls ("INVISIBLE")
        grid_dmbanti.SetFocus
    End If
End Sub

Private Sub MaskEdBox1_LOSTFOCUS()
On Error GoTo MSG:
If Trim(MaskEdBox1.Text) = "" Then 'Shraddha Pandit 17/11/08
    Exit Sub
End If
If Opt_patient.Value Then
    MaskEdBox1.Text = UCase(padzero(5, "C", MaskEdBox1.Text, 3))
    If new_flag Then
            cmb_reqno.clear
            req_qr(0) = MaskEdBox1.FormattedText
            req_qr(1) = Trim(txt_modcode.Text)
            req_qr(2) = Trim(txt_deptcode.Text)
            Set req_rs = req_qr.OpenResultset(1, 3)
        If req_rs.RowCount > 0 Then
            'FOR DISPLAYING REQNO WITH LABNO
            While Not req_rs.EOF
                GZZREQPATH_QR(0) = req_rs("REQNNO")
                Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                While Not GZZREQPATH_RS.EOF
                    cmb_reqno.AddItem req_chglen(req_rs("reqNno")) + "$" + GZZREQPATH_RS("PATHNO")
                    GZZREQPATH_RS.MoveNext
                Wend
                req_rs.MoveNext
            Wend
            req_rs.Close
            GZZREQPATH_RS.Close
            'UPTO HERE
        Else
            no_req = False
            req_rs.Close
            cmb_reqno.clear
            MsgBox "No Requisition for this Modality", vbOKOnly
            btnenabl "cancel"
            set_flag "asd"
            set_visible "All"
            Exit Sub
            
        End If 'end of rowcount
        
        SendKeys "^{f4}"
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
            cmb_reqno.clear
            'FOR DISPLAYING REQNO WITH LABNO
            While Not SELRREQUISITN_RS.EOF
                GZZREQPATH_QR(0) = SELRREQUISITN_RS("REQNNO")
                Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                While Not GZZREQPATH_RS.EOF
                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO")) + "$" + GZZREQPATH_RS("pathno")
                    GZZREQPATH_RS.MoveNext
                Wend
                SELRREQUISITN_RS.MoveNext
            Wend
            SELRREQUISITN_RS.Close
            GZZREQPATH_RS.Close
            'UPTO HERE
        Else
            no_req = False
            SELRREQUISITN_RS.Close
            'GZZREQPATH_RS.Close 'Shraddha Pandit 18/11/08
            MsgBox "No Requisition for this Modality", vbOKOnly
            btnenabl "cancel"
            set_flag "asd"
            set_visible "All"
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
                'no_req = True
                
                'ADDED BY MILIND ON 16-10-06 FOR DISPLAYING REQNO WITH LABNO
                While Not req_rs.EOF
                    GZZREQPATH_QR(0) = req_rs("REQNNO")
                    Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                    While Not GZZREQPATH_RS.EOF
                        cmb_reqno.AddItem req_chglen(req_rs("reqNno")) + "$" + GZZREQPATH_RS("pathno")
                        GZZREQPATH_RS.MoveNext
                    Wend
                        req_rs.MoveNext
                Wend
                'UPTO HERE
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
    
            If SELRREQUISITN_RS.RowCount > 0 Then
                    'ADDED BY MILIND ON 16-10-06 FOR DISPLAYING REQNO WITH LABNO
                While Not SELRREQUISITN_RS.EOF
                        GZZREQPATH_QR(0) = SELRREQUISITN_RS("REQNNO")
                        Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                    While Not GZZREQPATH_RS.EOF
                        cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO")) + "$" + GZZREQPATH_RS("pathno")
                        GZZREQPATH_RS.MoveNext
                    Wend
                        SELRREQUISITN_RS.MoveNext
                Wend
                SELRREQUISITN_RS.Close
                GZZREQPATH_RS.Close
    
             Else
                no_req = False
                MsgBox " No Requisition for this Modality"
                SELRREQUISITN_RS.Close
                GZZREQPATH_RS.Close
                btnenabl ("INI")
                Exit Sub
             End If
        End If
End If
 Exit Sub
MSG:
    MsgBox Error(Err)
End Sub

Private Sub Opt_patient_Click()
    If Opt_patient.Value = True Then
       lbl_caseno.Caption = "Case No."
       MaskEdBox1.Mask = "??>/#####"
       MaskEdBox1.SetFocus
    End If
    
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    MaskEdBox1.Mask = "#######"
    MaskEdBox1.SetFocus
End If
End Sub

Private Sub Text1_Change()
    If MSFlexGrid1.Col > 0 Then
        If Text1.Text <> "" Then
            MSFlexGrid1.CellBackColor = vbGreen
        End If
        MSFlexGrid1.Text = Text1.Text
    End If
End Sub

Private Sub txt_1_Change()
If Len(Trim(txt_1.Text)) = 1 And Not (UCase(Trim(txt_1.Text)) = "S" Or UCase(Trim(txt_1.Text)) = "I" Or UCase(Trim(txt_1.Text)) = "R") Then
    MsgBox "Please Enter S,I or R Only", vbOKOnly
    txt_1.Text = ""
    txt_1.SetFocus
Else
   grid_dmbanti.TextMatrix(grid_dmbanti.row, 1) = UCase(Trim(txt_1.Text))
End If
End Sub

Private Sub txt_2_Change()
If Len(Trim(txt_2.Text)) = 1 And Not (UCase(Trim(txt_2.Text)) = "S" Or UCase(Trim(txt_2.Text)) = "I" Or UCase(Trim(txt_2.Text)) = "R") Then
    MsgBox "Please Enter S,I or R Only", vbOKOnly
    txt_2.Text = ""
    txt_2.SetFocus
Else
    grid_dmbanti.TextMatrix(grid_dmbanti.row, 2) = UCase(Trim(txt_2.Text))
End If
End Sub

Private Sub txt_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(TXT_ORGAN3.Text) >= 1 Then
            Call GridControls("VISIBLE", 3)
            Call movecontrols(, 3)
        Else
            
            If grid_dmbanti.row < grid_dmbanti.Rows - 1 Then
                Call GridControls("VISIBLE", 0)
                Call movecontrols("DOWN", 0)
            Else
                GRIDSCROLL = True
                grid_dmbanti.Rows = grid_dmbanti.Rows + 1
                Call GridControls("VISIBLE", 0)
                Call movecontrols("DOWN", 0)
            End If
        End If
        
    End If
    
End Sub

Private Sub txt_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            If Len(TXT_ORGAN2.Text) >= 1 Then
                Call GridControls("VISIBLE", 2)
                Call movecontrols(, 2)
        Else
                If grid_dmbanti.row < grid_dmbanti.Rows - 1 Then
                    Call GridControls("VISIBLE", 0)
                    Call movecontrols("DOWN", 0)
                Else
                    'GET_ANTIBIOTIC
                    GRIDSCROLL = True
                    grid_dmbanti.Rows = grid_dmbanti.Rows + 1
                    Call GridControls("VISIBLE", 0)
                    Call movecontrols("DOWN", 0)
                End If
        End If
    End If
End Sub

Private Sub txt_3_Change()
If Len(Trim(txt_3.Text)) = 1 And Not (UCase(Trim(txt_3.Text)) = "S" Or UCase(Trim(txt_3.Text)) = "I" Or UCase(Trim(txt_3.Text)) = "R") Then
    MsgBox "Please Enter S,I or R Only", vbOKOnly
    txt_3.Text = ""
    txt_3.SetFocus
Else
    grid_dmbanti.TextMatrix(grid_dmbanti.row, 3) = UCase(Trim(txt_3.Text))
End If
End Sub

Private Sub txt_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       If grid_dmbanti.row < grid_dmbanti.Rows - 1 Then
              Call GridControls("VISIBLE", 0)
              Call movecontrols("DOWN", 0)
       Else
              GRIDSCROLL = True
              grid_dmbanti.Rows = grid_dmbanti.Rows + 1
              Call GridControls("VISIBLE", 0)
              Call movecontrols("DOWN", 0)
       End If
End If

End Sub

Private Sub txt_addendum_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_addendum
W_grossdes.Caption = "Addendum"
W_grossdes.WindowState = wdWindowStateMaximize
d_grossdesc.Content = txt_addendum
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Addendum"
W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_addendum.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub

Private Sub txt_addendum_LostFocus()
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument
txt_addendum.Text = d_grossdesc.Content
d_grossdesc.Content = ""
d_grossdesc.Close (wdDoNotSaveChanges)
W_grossdes.quit
Exit Sub
m4:
MsgBox Err.Description
End Sub

Private Sub txt_c_1_LOSTFOCUS()
txt_c_1.Text = Trim(UCase(txt_c_1.Text))
If Len(txt_c_1.Text) = 1 And Not (txt_c_1.Text = "S" Or txt_c_1.Text = "R") Then
    MsgBox "Please Enter S or R Only", vbOKOnly
    txt_c_1.Text = ""
    Exit Sub
ElseIf Len(txt_c_1.Text) = 0 Then
    grid_c_anti.TextMatrix(grid_c_anti.row, 1) = txt_c_1.Text
Else
    grid_c_anti.TextMatrix(grid_c_anti.row, 1) = txt_c_1.Text
    i = grid_c_anti.Rows - 1
    If i = grid_c_anti.row Then
        txt_c_1.Visible = False
        Exit Sub
    Else
        grid_c_anti.row = grid_c_anti.row + 1
        'check for next value and assign to textfield
        txt_c_1.Text = CStr(grid_c_anti.TextMatrix(grid_c_anti.row, 1))
        'reassign the positioning
        txt_c_1.Left = grid_c_anti.CellLeft + grid_c_anti.Left
        txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
        txt_c_1.SetFocus
     End If
End If
End Sub

Private Sub txt_organ3_lostfocus()
On Error GoTo MSG
GET_ANTIBIOTIC
If new_flag Then
  If Len(Trim(txt_organ1.Text)) >= 1 Then
        If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
            If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
                ''''THREE ORG
                grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 4
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           |^         3           "
            Else
                '''TWO ORG
                grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 3
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           "
             End If
        Else
                '''ONE ORG
                grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
        End If
  Else
  End If
  Call movecontrols(, 0)
  Call GridControls("VISIBLE", 0)
ElseIf VIEW_FLAG Then

ElseIf mod_flag Then
  If Len(Trim(txt_organ1.Text)) >= 1 Then
        If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
            If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
                ''''THREE ORG
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 4
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           |^         3           "
            Else
                '''TWO ORG
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 3
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           "
            End If
        Else
                '''ONE ORG
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
        End If
    End If

''getresult_bacteriology
        sel_forupd2_qr(0) = mod_reqno
        Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
        
        If sel_forupd2_rs.RowCount > 0 Then
                While Not sel_forupd2_rs.EOF
                    If grid_dmbanti.Cols = 2 Then
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
                                sel_forupd2_rs.MoveNext
                            Wend
                    ElseIf grid_dmbanti.Cols = 3 Then
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2")
                                sel_forupd2_rs.MoveNext
                            Wend
                    ElseIf grid_dmbanti.Cols = 4 Then
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2") & vbTab & sel_forupd2_rs("antibio3")
                                sel_forupd2_rs.MoveNext
                            Wend
                    End If
                Wend
        Else
            grid_dmbanti.Rows = grid_dmbanti.Rows + 1
            Call movecontrols(, 0)
            Call GridControls("VISIBLE", 0)
        End If
End If
 '''commented by kanchan  22/04/2008
  '  LOAD_ANTIBIOTIC
    Exit Sub
MSG:
    MsgBox Error(Err.Description)
End Sub
Private Sub GridControls(MODE As String, Optional Cl As Integer)
    
''''' Done by Kanchan for proper movement of controls on the grid

    If MODE = "INVISIBLE" Then
           
        cmb_antibiotic.Visible = False
        txt_1.Visible = False
        txt_2.Visible = False
        txt_3.Visible = False
                       
    ElseIf MODE = "VISIBLE" Then
        cmb_antibiotic.Visible = False
        txt_1.Visible = False
        txt_2.Visible = False
        txt_3.Visible = False
                
        Select Case Cl
    
        Case 0
            cmb_antibiotic.Visible = True
            cmb_antibiotic.SetFocus
        
        Case 1
            txt_1.Visible = True
            txt_1.SetFocus
        
        Case 2
            txt_2.Visible = True
            txt_2.SetFocus
        
        Case 3
            txt_3.Visible = True
            txt_3.SetFocus
        End Select
    End If
End Sub
Private Sub movecontrols(Optional Direction As String, Optional Cl As Integer)
    
'''' Done by Kanchan for proper movement of controls on the grid
        
    With grid_dmbanti
        Select Case Direction
            Case "DOWN"
                If .row < .Rows - 1 Then
                    .row = .row + 1
                End If
            Case "UP"
                If .row > 1 Then
                    .row = .row - 1
                End If
            End Select
         Select Case Cl
         Case 0
            .Col = 0
             cmb_antibiotic.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
             cmb_antibiotic.Text = .TextMatrix(.row, 0)
             cmb_antibiotic.Visible = True
             cmb_antibiotic.SetFocus
             
        Case 1
            .Col = 1
            txt_1.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            txt_1.Text = .TextMatrix(.row, 1)
            txt_1.Visible = True
            txt_1.SetFocus
        
        Case 2
            .Col = 2
            txt_2.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            txt_2.Text = .TextMatrix(.row, 2)
            txt_2.Visible = True
'            txt_2.SetFocus
            
         Case 3
            .Col = 3
            txt_3.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            txt_3.Text = .TextMatrix(.row, 3)
            txt_3.Visible = True
            txt_3.SetFocus
        End Select
      End With

End Sub

Private Sub GET_ANTIBIOTIC()
On Error GoTo MSG
cmb_antibiotic.clear
Dim RCN As Integer
Dim CHKFLG As Boolean
    If new_flag Then
            sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
                While Not sel_antirs.EOF
                    cmb_antibiotic.AddItem sel_antirs("antibdesc")
                    sel_antirs.MoveNext
                Wend
    End If
    If VIEW_FLAG Then
        CHKFLG = False
            sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
                                      
                While Not sel_antirs.EOF
                    CHKFLG = False
                    
                   For RCN = 1 To grid_dmbanti.Rows - 1
                    If Trim(grid_dmbanti.TextMatrix(RCN, 0)) = Trim(sel_antirs("ANTIBDESC")) Then
                       CHKFLG = True
                    End If
                   Next RCN
                        If CHKFLG = False Then
                        cmb_antibiotic.AddItem sel_antirs("antibdesc")
                        End If
                   sel_antirs.MoveNext
                Wend
    End If
        If mod_flag Then
        CHKFLG = False
            sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
                While Not sel_antirs.EOF
                    CHKFLG = False
                    For RCN = 1 To grid_dmbanti.Rows - 1
                    If Trim(grid_dmbanti.TextMatrix(RCN, 0)) = Trim(sel_antirs("ANTIBDESC")) Then
                       CHKFLG = True
                    End If
                   Next RCN
                        If CHKFLG = False Then
                        cmb_antibiotic.AddItem sel_antirs("antibdesc")
                        End If
                sel_antirs.MoveNext
                Wend
    End If
    Exit Sub
MSG:
    MsgBox Error(Err.Description)
End Sub
Private Sub txt_result_gotfocus()
On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    objDoc.Content = txt_result.Text
    objWord.Visible = True
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub txt_result_lostfocus()
On Error GoTo MSG
    txt_result.Text = objDoc.Content
    objDoc.Content = ""
    objDoc.Close (wdDoNotSaveChanges)
    objWord.quit
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub select_antibiotic()
On Error GoTo MSG
    If new_flag Then
         If Len(txt_organ1.Text) >= 1 Then
            If Len(TXT_ORGAN2.Text) >= 1 Then
                If (TXT_ORGAN3.Text) >= 1 Then
                    ''''three orgnism
                Else
                    ''''two organism
                End If
            Else
                 '''one organism
                 grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
           
                grid_dmbanti.row = 1
                grid_dmbanti.Col = 0
                cmb_antibiotic.Width = grid_dmbanti.CellWidth

                cmb_antibiotic.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                cmb_antibiotic.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                cmb_antibiotic.Visible = True
                
            End If
         Else
            MsgBox ("Select atleast one organism.....!"), vbOKOnly
         End If
    End If
    
    Exit Sub

MSG:
MsgBox Error(Err)
End Sub
Private Sub sel_antibiotic()
On Error GoTo MSG
    If new_flag Then
        If Len(txt_organ1.Text) >= 1 Then
            
            If Len(TXT_ORGAN2.Text) >= 1 Then
                
                If Len(TXT_ORGAN3.Text) >= 1 Then
                
                    grid_dmbanti.Rows = 1
                    grid_dmbanti.Cols = 2
                    grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
                    
                    grid_dmbanti.row = 1
                    grid_dmbanti.Col = 1
                    txt_1.Width = grid_dmbanti.CellWidth
                    txt_1.Height = grid_dmbanti.CellHeight
                    txt_1.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_1.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    grid_dmbanti.row = 1
                    grid_dmbanti.Col = 2
                    txt_2.Width = grid_dmbanti.CellWidth
                    txt_2.Height = grid_dmbanti.CellHeight
                    txt_2.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_2.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    grid_dmbanti.row = 1
                    grid_dmbanti.Col = 3
                    txt_3.Width = grid_dmbanti.CellWidth
                    txt_3.Height = grid_dmbanti.CellHeight
                    txt_3.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_3.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    txt_3.Visible = True
                    txt_2.Visible = True
                    txt_1.Visible = True
            
                Else
                'two organisms
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         "
                
                    grid_dmbanti.row = 1
                    grid_dmbanti.Col = 1
                    txt_1.Width = grid_dmbanti.CellWidth
                    txt_1.Height = grid_dmbanti.CellHeight
                    txt_1.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_1.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    grid_dmbanti.row = 1
                    grid_dmbanti.Col = 2
                    txt_2.Width = grid_dmbanti.CellWidth
                    txt_2.Height = grid_dmbanti.CellHeight
                    txt_2.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_2.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    txt_3.Visible = False
                    txt_2.Visible = True
                    txt_1.Visible = True
                    txt_1.SetFocus
                End If
            Else
            'only one organisms
            grid_dmbanti.Rows = 2
            grid_dmbanti.Cols = 2
            grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
                
                
                grid_dmbanti.row = 1
                grid_dmbanti.Col = 0
                cmb_antibiotic.Width = grid_dmbanti.CellWidth
                'cmb_antibiotic.Height = grid_dmbanti.CellHeight
                cmb_antibiotic.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                cmb_antibiotic.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                                
                grid_dmbanti.row = 1
                grid_dmbanti.Col = 1
                txt_1.Width = grid_dmbanti.CellWidth
                txt_1.Height = grid_dmbanti.CellHeight
                txt_1.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                txt_1.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                txt_1.Visible = False
                txt_2.Visible = False
                txt_3.Visible = False
                'txt_1.SetFocus
            End If
           Else
            MsgBox "Atleast 1 organisms is needed", vbOKOnly
            Exit Sub
         End If
    End If
    
    Exit Sub
    
MSG:
       MsgBox Error(Err)
End Sub
Private Sub LOAD_ANTIBIOTIC()
On Error GoTo MSG
    'check all organisms has been entered or not
    If new_flag Then
        If Len(txt_organ1.Text) >= 1 Then
            If Len(TXT_ORGAN2.Text) >= 1 Then
                If Len(TXT_ORGAN3.Text) >= 1 Then
                'three organisms
                grid_anti.Rows = 1
                grid_anti.Cols = 2
                grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
                Dim a
                    grid_anti.row = 1
                    grid_anti.Col = 1
                    txt_1.Width = grid_anti.CellWidth
                    txt_1.Height = grid_anti.CellHeight
                    txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_1.Top = grid_anti.CellTop + grid_anti.Top
                    grid_anti.row = 1
                    grid_anti.Col = 2
                    txt_2.Width = grid_anti.CellWidth
                    txt_2.Height = grid_anti.CellHeight
                    txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_2.Top = grid_anti.CellTop + grid_anti.Top
                    grid_anti.row = 1
                    grid_anti.Col = 3
                    txt_3.Width = grid_anti.CellWidth
                    txt_3.Height = grid_anti.CellHeight
                    txt_3.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_3.Top = grid_anti.CellTop + grid_anti.Top
                    txt_3.Visible = True
                    txt_2.Visible = True
                    txt_1.Visible = True
                Else
                'two organisms
                grid_anti.Rows = 1
                grid_anti.Cols = 2
                grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         "
                While Not sel_antirs.EOF
                    grid_anti.AddItem sel_antirs("antibdesc")
                    sel_antirs.MoveNext
                Wend
                    grid_anti.row = 1
                    grid_anti.Col = 1
                    txt_1.Width = grid_anti.CellWidth
                    txt_1.Height = grid_anti.CellHeight
                    txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_1.Top = grid_anti.CellTop + grid_anti.Top
                    grid_anti.row = 1
                    grid_anti.Col = 2
                    txt_2.Width = grid_anti.CellWidth
                    txt_2.Height = grid_anti.CellHeight
                    txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_2.Top = grid_anti.CellTop + grid_anti.Top
                    txt_3.Visible = False
                    txt_2.Visible = True
                    txt_1.Visible = True
                    txt_1.SetFocus
                End If
            Else
            'only one organisms
            grid_anti.Rows = 1
            grid_anti.Cols = 2
            grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           "
            While Not sel_antirs.EOF
              grid_anti.AddItem sel_antirs("antibdesc")
              sel_antirs.MoveNext
            Wend
                grid_anti.row = 1
                grid_anti.Col = 1
                txt_1.Width = grid_anti.CellWidth
                txt_1.Height = grid_anti.CellHeight
                txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                txt_1.Top = grid_anti.CellTop + grid_anti.Top
                txt_1.Visible = True
                txt_2.Visible = False
                txt_3.Visible = False
                txt_1.SetFocus
            End If
        Else
            MsgBox "Atleast 1 organisms is needed", vbOKOnly
            Exit Sub
         End If
    ElseIf VIEW_FLAG Then
    
    
            If Len(txt_organ1.Text) >= 1 Then
                 If Len(TXT_ORGAN2.Text) >= 1 Then
                     If Len(TXT_ORGAN3.Text) >= 1 Then
                        'three organisms
                        grid_anti.row = 1
                        grid_anti.Col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.Col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                        grid_anti.row = 1
                        grid_anti.Col = 3
                        txt_3.Text = grid_anti.TextMatrix(1, 3)
                        txt_3.Width = grid_anti.CellWidth
                        txt_3.Height = grid_anti.CellHeight
                        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_3.Top = grid_anti.CellTop + grid_anti.Top
                        txt_3.Visible = True
                    Else
                        'two organisms
                        grid_anti.Cols = 3
                        grid_anti.row = 1
                        grid_anti.Col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.Col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                    End If
                Else
                    'only one organisms
                    grid_anti.Cols = 2
                    grid_anti.row = 1
                        grid_anti.Col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
              End If
             txt_1.SetFocus
             End If
    
    ElseIf mod_flag Then
            
            
 If ((Len(txt_organ1.Text) >= 1 Or Len(TXT_ORGAN2.Text) >= 1 Or Len(TXT_ORGAN3.Text) >= 1)) Then
      
                    grid_anti.Rows = 1
                    grid_anti.Cols = 4
                    grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "

      
                    sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
                    sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
                    Set sel_antirs = sel_antiqr.OpenResultset(1, 3)

                    While Not sel_antirs.EOF
                        
                        retantibodies_qr(0) = mod_reqno
                        Dim anti As String
                        anti = Trim(sel_antirs("antibdesc"))
                        retantibodies_qr(1) = Trim(sel_antirs("antibdesc"))
                        'grid_anti.AddItem sel_antirs("antibdesc")
                        Set RETRESULT_RS = retantibodies_qr.OpenResultset(1, 3)
                        
                        If RETRESULT_RS.RowCount > 0 Then
                            grid_anti.AddItem RETRESULT_RS("antibiotic") & vbTab & RETRESULT_RS("antibio1") & vbTab & RETRESULT_RS("antibio2") & vbTab & RETRESULT_RS("antibio3")
                        Else
                            grid_anti.AddItem sel_antirs("antibdesc")
                        End If
                        sel_antirs.MoveNext
                    Wend
          
     If Len(txt_organ1.Text) >= 1 Then
                 If Len(TXT_ORGAN2.Text) >= 1 Then
                     If Len(TXT_ORGAN3.Text) >= 1 Then
                        'three organisms
                        grid_anti.row = 1
                        grid_anti.Col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.Col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                        grid_anti.row = 1
                        grid_anti.Col = 3
                        txt_3.Text = grid_anti.TextMatrix(1, 3)
                        txt_3.Width = grid_anti.CellWidth
                        txt_3.Height = grid_anti.CellHeight
                        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_3.Top = grid_anti.CellTop + grid_anti.Top
                        txt_3.Visible = True
                    Else
                        'two organisms
                        grid_anti.Cols = 3
                        grid_anti.row = 1
                        grid_anti.Col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.Col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                    End If
                Else
                    'only one organisms
                    grid_anti.Cols = 2
                    grid_anti.row = 1
                        grid_anti.Col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
              End If
             txt_1.SetFocus
             End If
                        
      'End With
      
            
 End If
            
    End If 'FIRST IF
    Exit Sub
MSG:
        MsgBox Error(Err)
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
    cmd_fullreport.Enabled = False
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
    cmd_fullreport.Enabled = True
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
    cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
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
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    cmd_fullreport.Enabled = False
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

Private Sub initialize_combo()
    cmb_gnb.ListIndex = 0
    cmb_gnc.ListIndex = 0
    cmb_gpb.ListIndex = 0
    cmb_gpc.ListIndex = 0
    cmb_patho.ListIndex = 0
End Sub

Private Sub clearupperpanel()
'upper panel
    lblicu.Visible = False
    MaskEdBox1.Text = ""
    cmb_reqno.clear
    cmb_modality.clear
    txt_deptcode.Text = ""
    cmb_dept.clear
    txt_microno.Text = ""
    txt_name.Text = ""
    txt_age.Text = ""
    txt_spm.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_status.Text = ""
    txt_ward.Text = ""
    txt_bedno.Text = ""
    txt_pridoc.Text = ""
    txt_modcode.Text = ""
    txt_refdoc.Text = ""
    CMB_PROC.clear
    cmb_enteryby.Text = ""
    cmb_final.Text = ""
'   grid_anti.Row = 1
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
End Sub
'for procdure bacteriology
Private Sub clearlowerpanel1()
    lblicu.Visible = False
    txt_epith.Text = ""
    txt_pus.Text = ""
    txt_other.Text = ""
    'txt_ziehl.Text = ""
    txt_koh.Text = ""
    txt_organ1.Text = ""
    TXT_ORGAN2.Text = ""
    TXT_ORGAN3.Text = ""
    txt_1.Text = ""
    txt_2.Text = ""
    txt_3.Text = ""
    txt_report.Text = ""
    txt_comment.Text = ""
    initialize_combo
    grid_anti.clear
    grid_dmbanti.clear
    cmb_final.Text = ""
    cmb_enteryby.Text = ""
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
End Sub
'for procedure clinical microbiology
Private Sub clearlowerpanel2()
On Error GoTo MSG
    lblicu.Visible = False
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
    cmb_final.Text = ""
    cmb_enteryby.Text = ""
    Exit Sub
MSG:
    MsgBox Error(Err)
End Sub
'for other procedure
Private Sub clearlowerpanel3()
On Error GoTo MSG
    lblicu.Visible = False
    cmb_final.Text = ""
    cmb_enteryby.Text = ""
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
    Exit Sub
MSG:
    MsgBox Error(Err)
End Sub
Private Sub clearlowerpanel4()
On Error GoTo MSG
    lblicu.Visible = False
    cmb_final.Text = ""
    cmb_enteryby.Text = ""
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
    txt_addendum.Text = ""
    cmb_addendumby.Text = ""
    Exit Sub
MSG:
    MsgBox Error(Err)
End Sub

'common procedure for new and view mode
Private Sub getupperpanel()
On Error GoTo MSG
        'THIS TO GET DEPTCODE AND MODALITYCODE
        sel_micro_qr(0) = mod_reqno
        Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
        While Not sel_micro_rs.EOF
            txt_microno.Text = Trim(sel_micro_rs("pathno"))
            sel_micro_rs.MoveNext
        Wend
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
            Dim rn
            rn = Trim(cmb_reqno.Text)
            Dim S As Integer
            S = InStr(1, cmb_reqno, "$")
            rn = req_chglen(Mid(Trim(cmb_reqno), 1, S - 1))
            
            SELSPONANTIBQR(0) = rn
            Set SELSPONANTIBRS = SELSPONANTIBQR.OpenResultset(1, 3)
            
            If SEL_detail_rs.RowCount = 0 Then
                SEL_detail_rs.Close
                Exit Sub
            Else
                txt_sex.Text = SEL_detail_rs("sex")
                txt_name.Text = Trim(SEL_detail_rs("first_name")) & " " & Trim(SEL_detail_rs("mid_name")) & " " & Trim(SEL_detail_rs("last_name"))
                txt_age = DateDiff("YYYY", SEL_detail_rs("BIRTH_DATE"), CStr(DATE11))
                txt_pridoc.Text = SEL_detail_rs("ud_name")
                txt_category.Text = SEL_detail_rs("pat_catg")
                txt_spm.Text = SELSPONANTIBRS("specimen")
                SEL_detail_rs.Close
            End If
        End If
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = mod_reqno
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        If REFDR_RS.RowCount > 0 Then
        'put check here also 20/12/02
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If Not callstatus Then
                If Not new_flag = True Then 'Shraddha 17/11/08
                    If REFDR_RS("COMMITFLAG") = "Y" Then
                        MsgBox "Final Report is Generated, You cannot Modify the Final Report"
                        cmd_modify.Enabled = False
                        txt_addendum.Enabled = True
                        cmb_addendumby.Enabled = True
                        cmd_save.Enabled = True
                    Else
                        cmd_modify.Enabled = True
                    End If
                End If
            End If
        End If
            sel_bed_qr(0) = MaskEdBox1.FormattedText
            Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
            
        If REFDR_RS("PATNSTAT") = "IP" Then
            
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
            sel_micro_rs.MoveFirst
            txt_ward.Text = sel_micro_rs("WARD_NAME")
            txt_bedno.Text = sel_micro_rs("bedno")
            txt_status.Text = REFDR_RS("PATNSTAT")
            sel_bed_rs.Close
        Else
            txt_status.Text = REFDR_RS("PATNSTAT")
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
            sel_bed_rs.Close
        End If
        CMB_PROC.clear
        'for procedure previously selected
        
        sel_proccode_qr(0) = mod_reqno
        Set sel_proccode_rs = sel_proccode_qr.OpenResultset(1, 3)
        If Not sel_proccode_rs.EOF Then
            While Not sel_proccode_rs.EOF
                CMB_PROC.AddItem Trim(sel_proccode_rs(1))
                sel_proccode_rs.MoveNext
            Wend
            CMB_PROC.ListIndex = 0
            sel_proccode_rs.Close
        End If
        Exit Sub
MSG:
        MsgBox Error(Err)
End Sub

'for bacteriology get result from database
Private Sub getresult_bacteriology()
On Error GoTo MSG
sel_forupd1_qr(0) = mod_reqno
Set sel_forupd1_rs = sel_forupd1_qr.OpenResultset(1, 3)
While Not sel_forupd1_rs.EOF
    txt_epith.Text = Trim(sel_forupd1_rs("epithecell"))
    txt_pus.Text = Trim(sel_forupd1_rs("puscell"))
    txt_other.Text = Trim(sel_forupd1_rs("ofinding"))
    txt_ziehl.Text = Trim(sel_forupd1_rs("ZNSTAIN"))
    For i = 0 To (cmb_gpc.ListCount - 1)
        If Trim(cmb_gpc.List(i)) = Trim(sel_forupd1_rs("gpc")) Then
                cmb_gpc.ListIndex = i
        End If
    Next i
    For i = 0 To (cmb_gnc.ListCount - 1)
        If Trim(cmb_gnc.List(i)) = Trim(sel_forupd1_rs("gnc")) Then
                cmb_gnc.ListIndex = i
        End If
    Next i
    For i = 0 To (cmb_gnb.ListCount - 1)
        If Trim(cmb_gnb.List(i)) = Trim(sel_forupd1_rs("gnb")) Then
                cmb_gnb.ListIndex = i
        End If
    Next i
    For i = 0 To (cmb_gpb.ListCount - 1)
        If Trim(cmb_gpb.List(i)) = Trim(sel_forupd1_rs("gpb")) Then
                cmb_gpb.ListIndex = i
        End If
    Next i
    For i = 0 To (cmb_india.ListCount - 1)
        If Trim(cmb_india.List(i)) = Trim(sel_forupd1_rs("iipc")) Then
                cmb_india.ListIndex = i
        End If
    Next i
    txt_koh.Text = Trim(sel_forupd1_rs("kohamt"))
    txt_organ1.Text = "" & Trim(sel_forupd1_rs("orgso1"))
    TXT_ORGAN2.Text = "" & Trim(sel_forupd1_rs("orgso2"))
    TXT_ORGAN3.Text = "" & Trim(sel_forupd1_rs("orgso3"))
    If sel_forupd1_rs.EOF = False Then
        txt_comment.Text = "" & Trim(sel_forupd1_rs("comments"))
    End If
    If Len(Trim(sel_forupd1_rs("REPORT"))) > 1 Then
        cmb_patho.clear
        cmb_patho.AddItem "Yes"
        cmb_patho.AddItem "No"
        cmb_patho.Text = "No"
        cmb_patho_click
        txt_report.Text = "" & sel_forupd1_rs("REPORT")
        Exit Sub
        
    Else
    'sel_forupd1_rs.Close
    'for loading antibiotics with values
        cmb_patho.clear
        cmb_patho.AddItem "Yes"
        cmb_patho.AddItem "No"
        cmb_patho.Text = "Yes"
        cmb_patho_click
        txt_report.Text = ""
        sel_forupd2_qr(0) = mod_reqno
        Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
        While Not sel_forupd2_rs.EOF
        grid_anti.clear
        'grid_dmbanti.clear
        
            If Len(Trim(sel_forupd1_rs("orgso1"))) > 0 Then
                If Len(Trim(sel_forupd1_rs("orgso2"))) > 0 Then
                    If Len(Trim(sel_forupd1_rs("orgso3"))) > 0 Then
                       'three organisms
                        grid_dmbanti.Cols = 4
                        grid_dmbanti.Rows = 1
                        grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
                        While Not sel_forupd2_rs.EOF
                            grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2") & vbTab & sel_forupd2_rs("antibio3")
                            sel_forupd2_rs.MoveNext
                        Wend
                    Else
                    'two organisms
                    grid_dmbanti.Cols = 3
                    grid_dmbanti.Rows = 1
                    grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         "
                    While Not sel_forupd2_rs.EOF
                            grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2")
                            sel_forupd2_rs.MoveNext
                    Wend
                    End If
                Else
                'only one organisms
                        grid_dmbanti.Cols = 2
                        grid_dmbanti.Rows = 1
                        grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
                                sel_forupd2_rs.MoveNext
                            Wend
'                     End If
                End If
            Else
                sel_forupd2_rs.Close
                Exit Sub
            End If
         Wend
         sel_forupd2_rs.Close
    End If
    sel_forupd1_rs.MoveNext
Wend
sel_forupd1_rs.Close
    Exit Sub
MSG:
    MsgBox Error(Err)
End Sub
'to get result from dmure table for view of urine report
Private Sub getresult_urine()
On Error GoTo MSG
    sel_forupd3_qr(0) = mod_reqno
    Set sel_forupd3_rs = sel_forupd3_qr.OpenResultset(1, 3)
    While Not sel_forupd3_rs.EOF
        sel_forupd3_rs.MoveNext
    Wend
    sel_forupd3_rs.Close
    Exit Sub
MSG:
    MsgBox Error(Err)
End Sub

Private Sub getresult_stool()
On Error GoTo MSG
    sel_forupd4_qr(0) = mod_reqno
    Set sel_forupd4_rs = sel_forupd4_qr.OpenResultset(1, 3)
    While Not sel_forupd4_rs.EOF
        'HERE ALL STOOL FIELD
        sel_forupd4_rs.MoveNext
    Wend
    sel_forupd4_rs.Close
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub getresult_other()
On Error GoTo MSG
    sel_forupd5_qr(0) = mod_reqno
    Set sel_forupd5_rs = sel_forupd5_qr.OpenResultset(1, 3)
    While Not sel_forupd5_rs.EOF
        sel_forupd5_rs.MoveNext
    Wend
    Exit Sub
MSG:
    MsgBox Error(Err)
End Sub
Private Sub getresult_myco()
On Error GoTo MSG
    sel_forupd6_qr(0) = mod_reqno
    Set sel_forupd6_rs = sel_forupd6_qr.OpenResultset(1, 3)
    
    sel_forupd6_rs.Close
    
    sel_forupd2_qr(0) = mod_reqno
    Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
        
    Exit Sub
MSG:
    MsgBox Err.Description
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
End If
End Sub

Private Sub set_position() 'set top position for all three textboxes
On Error GoTo MSG
MSG:
    MsgBox Error(Err)
End Sub
    
Private Sub enable_field()
On Error GoTo MSG
'called when modifying records so disable upper panel
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    cmb_dept.Enabled = False
    cmb_reqno.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_enteryby.Enabled = True
    cmb_final.Enabled = True
'for panel1
    txt_epith.Enabled = True
    txt_pus.Enabled = True
    txt_other.Enabled = True
    txt_koh.Enabled = True
    txt_organ1.Enabled = True
    TXT_ORGAN2.Enabled = True
    TXT_ORGAN3.Enabled = True
    txt_1.Enabled = True
    txt_2.Enabled = True
    txt_3.Enabled = True
    txt_comment.Enabled = True
    txt_report.Enabled = True
'for panel2
       
'for panel3
    
'for all combo
    cmb_gnb.Enabled = True
    cmb_gnc.Enabled = True
    cmb_gpb.Enabled = True
    cmb_gpc.Enabled = True
    cmb_india.Enabled = True
    txt_ziehl.Enabled = True
    cmb_patho.Enabled = True
'for panel4
     txt_c_modflg = True
    
    'for addendum
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
        cmb_addendumby.Enabled = True
    Else
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
    
    Exit Sub
    
MSG:
    MsgBox Err.Description
End Sub
Private Sub disable_field()
'for panel1(bacteriology)
    txt_epith.Enabled = False
    txt_pus.Enabled = False
    txt_other.Enabled = False
    txt_koh.Enabled = False
    txt_organ1.Enabled = False
    TXT_ORGAN2.Enabled = False
    TXT_ORGAN3.Enabled = False
    txt_1.Enabled = False
    txt_2.Enabled = False
    
    txt_3.Enabled = False
    txt_comment.Enabled = False
    txt_report.Enabled = False
'for panel2(urine and stool)
    
'for all combo
    cmb_gnb.Enabled = False
    cmb_gnc.Enabled = False
    cmb_gpb.Enabled = False
    cmb_gpc.Enabled = False
    cmb_india.Enabled = False
    txt_ziehl.Enabled = False
    cmb_patho.Enabled = False
 'for panel4(myco)
     
End Sub

Private Sub get_myco()
On Error GoTo MSG
    Dim ST1 As String
    ST1 = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
    
    sel_submodali_qr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
    Set sel_submodali_rs = sel_submodali_qr.OpenResultset(1, 3)
    If Not sel_submodali_rs.EOF Then
        
        While Not sel_submodali_rs.EOF
            
            sel_submodali_rs.MoveNext
        Wend
        sel_submodali_rs.Close
     Else
        MsgBox "Sorry No Organism Isolated is Available,Use Master Screen For Entring  ", vbOKOnly
        sel_submodali_rs.Close
        Exit Sub
     End If
     Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub insert_urine()
    'insert in urine table DMURE
            ins_4_qr(0) = mod_reqno
            ins_4_qr(1) = Trim(txt_urinecolour.Text)
            ins_4_qr(2) = Trim(txt_gravity.Text)
            ins_4_qr(3) = Trim(cmb_ph.Text) 'Trim(txt_ph.Text)
            ins_4_qr(4) = Trim(cmb_proteins.Text) 'Trim(txt_protins.Text)
            ins_4_qr(5) = Trim(cmb_glucose.Text) 'Trim(txt_glucose.Text)
            ins_4_qr(6) = Trim(cmb_ketones.Text) 'Trim(txt_ketones.Text)
            ins_4_qr(7) = Trim(cmb_bilirubin.Text)  'Trim(txt_bilirubin.Text)
            ins_4_qr(8) = Trim(cmb_uroblngn.Text) 'Trim(txt_uro.Text)
            ins_4_qr(9) = Trim(cmb_blood.Text) 'Trim(txt_blood.Text)
            'ins_4_qr(10) = Trim(txt_MicroExam.Text)
            ins_4_qr(10) = Null
            ins_4_qr(11) = Trim(txt_wbc.Text)
            ins_4_qr(12) = Trim(TXT_RBC.Text)
            ins_4_qr(13) = Trim(txt_squamous.Text)
            ins_4_qr(14) = Trim(txt_bacteria.Text)
            ins_4_qr(15) = Trim(txt_casts.Text)
            If Len(txt_urinother.Text) > 0 Then
                ins_4_qr(16) = Null
                ins_4_qr(16) = Trim(txt_urinother.Text)
            Else
                ins_4_qr(16) = Null
                ins_4_qr(16) = " "
            End If
            If Len(txt_urinecomment.Text) > 0 Then
                ins_4_qr(17) = Null
                ins_4_qr(17) = Trim(txt_urinecomment.Text)
            Else
                ins_4_qr(17) = Null
                ins_4_qr(17) = " "
            End If
            ins_4_qr.Execute
        update_rislib
        update_rrequisitn
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub insert_stool()
On Error GoTo MSG
            ins_5_qr(0) = mod_reqno
            ins_5_qr(1) = Trim(txt_stoolcolour.Text)
            ins_5_qr(2) = Trim(txt_consistency.Text)
            ins_5_qr(3) = Trim(txt_mucus.Text)
            ins_5_qr(4) = Trim(txt_sblood.Text)
            ins_5_qr(5) = Trim(txt_parasties.Text)
            ins_5_qr(6) = Trim(txt_occult.Text)
            ins_5_qr(7) = Trim(txt_Ova.Text)
            ins_5_qr(8) = Trim(txt_cysts.Text)
            ins_5_qr(9) = Trim(txt_larvae.Text)
            ins_5_qr(10) = Trim(txt_puscells.Text)
            ins_5_qr(11) = Trim(txt_stoolrbc.Text)
            ins_5_qr(12) = Trim(txt_microphages.Text)
            ins_5_qr(13) = Trim(txt_starch.Text)
            ins_5_qr(14) = Trim(txt_fat.Text)
            ins_5_qr(15) = Trim(txt_stoolother.Text)
            If Len(txt_stoolcomments.Text) > 0 Then
                ins_5_qr(16) = Null
                ins_5_qr(16) = txt_stoolcomments.Text
            Else
                ins_5_qr(16) = Null
                ins_5_qr(16) = " "
            End If
            ins_5_qr.Execute
            
            update_rislib
    update_rrequisitn
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub insert_myco()
On Error GoTo MSG
            ins_6_qr(0) = mod_reqno
            ins_6_qr(1) = cmb_stain.Text
            ins_6_qr(2) = cmb_organisms.Text
            If Len(txt_c_report.Text) > 0 Then
                ins_6_qr(3) = Null
                ins_6_qr(3) = txt_c_report.Text
            Else
                ins_6_qr(3) = Null
                ins_6_qr(3) = " "
            End If
            If Len(txt_c_comment.Text) > 0 Then
                ins_6_qr(4) = Null
                ins_6_qr(4) = txt_c_comment.Text
            Else
                ins_6_qr(4) = Null
                ins_6_qr(4) = " "
            End If
            ins_6_qr.Execute
            i = grid_c_anti.Rows
            srno = 1
            If Not i = 1 Then 'srno = 1 Then
                If Not (Trim(grid_c_anti.TextMatrix(srno, 1)) = "") Then
                    While Not srno = i
                        ins_2_qr(0) = mod_reqno
                        ins_2_qr(1) = Trim(grid_c_anti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                        ins_2_qr(2) = grid_c_anti.TextMatrix(srno, 1) 'get first value and insert
                        ins_2_qr(3) = Null
                        ins_2_qr(4) = Null
                        ins_2_qr.Execute
                        srno = srno + 1
                    Wend
                 Else
                        ins_2_qr(0) = mod_reqno
                        ins_2_qr(1) = Null 'first get name of antibiotic and insert
                        ins_2_qr(2) = Null 'get first value and insert
                        ins_2_qr(3) = Null
                        ins_2_qr(4) = Null
                        ins_2_qr.Execute
                 End If
             Else
                ''''INSERT SINGLE RECORD IN DATABASE FOR CRYSTAL REPORT JOIN
                ins_2_qr(0) = mod_reqno
                ins_2_qr(1) = Null
                ins_2_qr(2) = Null
                ins_2_qr(3) = Null
                ins_2_qr(4) = Null
                ins_2_qr.Execute
            End If
                update_rislib
                update_rrequisitn
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub insert_other()
On Error GoTo MSG
ins_0_qr(0) = mod_reqno

                ins_0_qr(1) = Null
            If Len(txt_test_perform.Text) > 0 Then
                ins_0_qr(2) = Null
                ins_0_qr(2) = Trim(txt_test_perform.Text)
            Else
                ins_0_qr(2) = Null
                ins_0_qr(2) = " "
            End If
            If Len(txt_result.Text) > 0 Then
                ins_0_qr(3) = Null
                ins_0_qr(3) = txt_result.Text
            Else
                ins_0_qr(3) = Null
                ins_0_qr(3) = " "
            End If
            ins_0_qr.Execute
            update_rislib
update_rrequisitn

Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub UPDATE_DMBAANTI()
                     
''''' Done By Kanchan
 ''''' Check for data in the database
 
                    sel_forupd2_qr(0) = mod_reqno
                    Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
                                     
                    If Not sel_forupd2_rs.EOF Then
                    
                        row1 = 0
                        row1 = grid_dmbanti.Rows - 1
    
                            If Len(Trim(grid_dmbanti.TextMatrix(row1, 0))) > 0 Then
                                row1 = grid_dmbanti.Rows - 1
                            Else
                                row1 = grid_dmbanti.Rows - 2
                            End If
                        
                        srno = 1
                        For i = 1 To row1
                                If Len(Trim(txt_organ1.Text)) >= 1 Then
                                    If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
                                        If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
                                            ''''''' three org
                                            If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then '''''CHECK FOR ANTIBIOTIC
                                                If ((Not Trim(grid_dmbanti.TextMatrix(srno, 1)) = "") Or (Not Trim(grid_dmbanti.TextMatrix(srno, 2)) = "") Or (Not Trim(grid_dmbanti.TextMatrix(srno, 3)) = "")) Then
                                                    
                                                        sel_formod2_qr(0) = mod_reqno
                                                        sel_formod2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                        If sel_formod2_flg = False Then
                                                            Set sel_formod2_rs = sel_formod2_qr.OpenResultset(1, 3)
                                                            sel_formod2_flg = True
                                                        Else
                                                            sel_formod2_rs.MoreResults
                                                            sel_formod2_rs.Requery
                                                        End If
                                    
                                                        If sel_formod2_rs.RowCount > 0 Then
                                                            upd_2_qr(0) = Trim(grid_dmbanti.TextMatrix(srno, 1)) 'get first value and insert
                                                            upd_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 2)) 'get second value and insert"
                                                            upd_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 3)) 'get third value and insert
                                                            upd_2_qr(3) = mod_reqno
                                                            upd_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                                                            upd_2_qr.Execute
                                                        Else
                                    
                                                        'reqno,antibiotic,antibio1,antibio2,antibio3
                                            
                                                            ins_2_qr(0) = Trim(mod_reqno)
                                                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                                                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                                                            ins_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 3))
                                                            ins_2_qr.Execute
                                                        End If
                                                   
                                                Else
                                                    ''''delete from table
                                                    
                                                      del_dmbaanti_qr1(0) = mod_reqno
                                                      del_dmbaanti_qr1(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                      del_dmbaanti_qr1.Execute
                                      
                                                End If
                                             End If ''''
                                        Else
                                            ''''''' two org
                                                If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then '''''CHECK FOR ANTIBIOTIC
                                                    If ((Not Trim(grid_dmbanti.TextMatrix(srno, 1)) = "") Or (Not Trim(grid_dmbanti.TextMatrix(srno, 2)) = "")) Then
                                                    
                                                        sel_formod2_qr(0) = mod_reqno
                                                        sel_formod2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                        If sel_formod2_flg = False Then
                                                            Set sel_formod2_rs = sel_formod2_qr.OpenResultset(1, 3)
                                                            sel_formod2_flg = True
                                                        Else
                                                            sel_formod2_rs.MoreResults
                                                            sel_formod2_rs.Requery
                                                        End If
                                    
                                                        If sel_formod2_rs.RowCount > 0 Then
                                                            upd_2_qr(0) = Trim(grid_dmbanti.TextMatrix(srno, 1)) 'get first value and insert
                                                            upd_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 2)) 'get second value and insert"
                                                            upd_2_qr(2) = Null 'get third value and insert
                                                            upd_2_qr(3) = mod_reqno
                                                            upd_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                                                            upd_2_qr.Execute
                                                        Else
                                    
                                                        'reqno,antibiotic,antibio1,antibio2,antibio3
                                            
                                                            ins_2_qr(0) = Trim(mod_reqno)
                                                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                                                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                                                            ins_2_qr(4) = Null
                                                            ins_2_qr.Execute
                                                        End If
                                        
                                                Else
                                                    ''''delete from table
                                                      del_dmbaanti_qr1(0) = mod_reqno
                                                      del_dmbaanti_qr1(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                      del_dmbaanti_qr1.Execute
                                      
                                                End If
                                            
                                            End If '''''
                                        End If
                                    Else
                                            '''''''one org
                                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then '''''CHECK FOR ANTIBIOTIC
                                            If ((Not Trim(grid_dmbanti.TextMatrix(srno, 1)) = "")) Then
                                                    
                                                        sel_formod2_qr(0) = mod_reqno
                                                        sel_formod2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                        If sel_formod2_flg = False Then
                                                            Set sel_formod2_rs = sel_formod2_qr.OpenResultset(1, 3)
                                                            sel_formod2_flg = True
                                                        Else
                                                            sel_formod2_rs.MoreResults
                                                            sel_formod2_rs.Requery
                                                        End If
                                    
                                                        If sel_formod2_rs.RowCount > 0 Then
                                                            upd_2_qr(0) = Trim(grid_dmbanti.TextMatrix(srno, 1)) 'get first value and insert
                                                            upd_2_qr(1) = Null 'get second value and insert"
                                                            upd_2_qr(2) = Null 'get third value and insert
                                                            upd_2_qr(3) = mod_reqno
                                                            upd_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                                                            upd_2_qr.Execute
                                                        Else
                                    
                                                        'reqno,antibiotic,antibio1,antibio2,antibio3
                                            
                                                            ins_2_qr(0) = Trim(mod_reqno)
                                                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                                                            ins_2_qr(3) = Null
                                                            ins_2_qr(4) = Null
                                                            ins_2_qr.Execute
                                                        End If
                                        
                                                Else
                                                    ''''delete from table
                                                      del_dmbaanti_qr1(0) = mod_reqno
                                                      del_dmbaanti_qr1(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                      del_dmbaanti_qr1.Execute
                                      
                                                End If
                                           End If '''''
                                    End If
                                End If
                             
                            srno = srno + 1
                            
                        Next i
                        
                    Else
                        INSERT_DMBANTI
                    End If

End Sub

Private Sub INSERT_DMBANTI()

'''' Done by KANCHAN

If Len(txt_organ1.Text) >= 1 Or Len(TXT_ORGAN2.Text) >= 1 Or Len(TXT_ORGAN3.Text) >= 1 Then

    i = grid_dmbanti.Rows - 1
    
    If Len(Trim(grid_dmbanti.TextMatrix(i, 0))) > 0 Then
        i = grid_dmbanti.Rows
    Else
        i = grid_dmbanti.Rows - 1
    End If
    
'    i = grid_dmbanti.Rows
    srno = 1
    
    If i = 1 Then
        
            ins_2_qr(0) = Trim(mod_reqno)
            ins_2_qr(1) = Null
            ins_2_qr(2) = Null
            ins_2_qr(3) = Null
            ins_2_qr(4) = Null
            ins_2_qr.Execute
    
    End If
    While Not srno = i
    
        If Len(txt_organ1.Text) >= 1 Then
            If Len(TXT_ORGAN2.Text) >= 1 Then
                If Len(TXT_ORGAN3.Text) >= 1 Then
                    '''THREE ORG
                    If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then
                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 1))) > 0 Or Len(Trim(grid_dmbanti.TextMatrix(srno, 2))) > 0 Or Len(Trim(grid_dmbanti.TextMatrix(srno, 3))) > 0) Then
                    
                            ins_2_qr(0) = Trim(mod_reqno)
                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                            ins_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 3))
                            ins_2_qr.Execute
                        End If
                    
                    End If
                Else
                    ''''TWO ORG
                    If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then
                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 1))) > 0 Or Len(Trim(grid_dmbanti.TextMatrix(srno, 2))) > 0) Then
                            ins_2_qr(0) = Trim(mod_reqno)
                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                            ins_2_qr(4) = Null
                            ins_2_qr.Execute
                    
                        End If
                    End If
                 End If
            Else
                
                   ''ONE ORG
                    If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then
                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 1))) > 0) Then
                            ins_2_qr(0) = Trim(mod_reqno)
                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                            ins_2_qr(3) = Null
                            ins_2_qr(4) = Null
                            ins_2_qr.Execute
                        End If
                    End If
            End If
        End If
    
        srno = srno + 1
    Wend

Else
        
End If

End Sub

Private Sub insert_bacteriology()
On Error GoTo MSG
            ins_1_qr(0) = mod_reqno
            ins_1_qr(1) = Trim(txt_epith.Text)
            ins_1_qr(2) = Trim(txt_pus.Text)
            ins_1_qr(3) = Trim(txt_other.Text)

            ins_1_qr(4) = Trim(txt_ziehl.Text)
            ins_1_qr(5) = Trim(txt_koh.Text)
            ins_1_qr(6) = Trim(cmb_gpc.Text)
            ins_1_qr(7) = Trim(cmb_gnc.Text)
            ins_1_qr(8) = Trim(cmb_gpb.Text)
            ins_1_qr(9) = Trim(cmb_gnb.Text)
            ins_1_qr(10) = Trim(cmb_india.Text)
            If Len(txt_organ1.Text) > 0 Then
                ins_1_qr(11) = Null
                ins_1_qr(11) = Trim(txt_organ1.Text)
            Else
                ins_1_qr(11) = Null
                ins_1_qr(11) = " "
            End If
            If Len(TXT_ORGAN2.Text) > 0 Then
                ins_1_qr(12) = Null
                ins_1_qr(12) = Trim(TXT_ORGAN2.Text)
            Else
                ins_1_qr(12) = Null
                ins_1_qr(12) = " "
            End If
            If Len(TXT_ORGAN3.Text) > 0 Then
                ins_1_qr(13) = Null
                ins_1_qr(13) = Trim(TXT_ORGAN3.Text)
            Else
                ins_1_qr(13) = Null
                ins_1_qr(13) = " "
            End If
            If cmb_patho.Text = "Yes" Then
                ins_1_qr(14) = Null
            Else
                If Len(txt_report.Text) > 0 Then
                    ins_1_qr(14) = Null
                    ins_1_qr(14) = Trim(txt_report.Text)
                Else
                    ins_1_qr(14) = Null
                    ins_1_qr(14) = " "
                End If
            End If
            If Len(txt_comment.Text) > 0 Then

                ins_1_qr(15) = Trim(txt_comment.Text)
            Else
                ins_1_qr(15) = Null
                ins_1_qr(15) = " "
            End If
            ins_1_qr.Execute
            MsgBox "Record successfully saved"
            
   INSERT_DMBANTI
    update_rislib
    update_rrequisitn
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub update_rrequisitn()
On Error GoTo MSG
    upd_11_qr(0) = MaskEdBox1.FormattedText
    upd_11_qr(1) = mod_reqno
    upd_11_qr.Execute
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub update_rislib()
On Error GoTo MSG
    upd_8_qr(0) = cmb_enteryby.Text
    upd_8_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_8_qr(2) = cmb_final.Text
    upd_8_qr(3) = "T"
    upd_8_qr(4) = cmb_final.Text
    upd_8_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_8_qr(6) = Trim(mod_reqno)
    upd_8_qr.Execute
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub GET_ENTER_BY()
Dim ENTR_BY As String
On Error GoTo MSG
    sel_ris_qr(0) = mod_reqno
    Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
    If Not sel_ris_rs.EOF Then
        While Not sel_ris_rs.EOF

            cmb_enteryby.Text = sel_ris_rs("ENTERBY") & ""
            cmb_final.Text = sel_ris_rs("freportby") & ""

            cmb_enteryby.Enabled = False
            cmb_final.Enabled = False
        sel_ris_rs.MoveNext
        Wend
        sel_ris_rs.Close
    Else
        sel_ris_rs.Close
    End If
    Exit Sub
MSG:
    MsgBox Err.Description

End Sub
Private Sub UPDATE_URINE()
On Error GoTo MSG
    upd_4_qr(0) = txt_urinecolour.Text
    upd_4_qr(1) = txt_gravity.Text
    upd_4_qr(2) = cmb_ph.Text 'txt_ph.Text
    upd_4_qr(3) = cmb_proteins.Text 'txt_protins.Text
    upd_4_qr(4) = cmb_glucose.Text 'txt_glucose.Text
    upd_4_qr(5) = cmb_ketones.Text 'txt_ketones.Text
    upd_4_qr(6) = cmb_bilirubin.Text 'txt_bilirubin.Text
    upd_4_qr(7) = cmb_uroblngn.Text 'txt_uro.Text
    upd_4_qr(8) = cmb_blood.Text 'txt_blood.Text
    'upd_4_qr(9) = txt_MicroExam.Text
    upd_4_qr(9) = Null
    upd_4_qr(10) = txt_wbc.Text
    upd_4_qr(11) = TXT_RBC.Text
    upd_4_qr(12) = txt_squamous.Text
    upd_4_qr(13) = txt_bacteria.Text
    upd_4_qr(14) = txt_casts.Text
    If Len(txt_urinother.Text) > 0 Then
        upd_4_qr(15) = Null
        upd_4_qr(15) = txt_urinother.Text
    Else
        upd_4_qr(15) = Null
        upd_4_qr(15) = " "
    End If
    If Len(txt_urinecomment.Text) > 0 Then
        upd_4_qr(16) = Null
        upd_4_qr(16) = txt_urinecomment.Text
    Else
        upd_4_qr(16) = Null
        upd_4_qr(16) = " "
    End If
    upd_4_qr(17) = mod_reqno
    upd_4_qr.Execute
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub UPDATE_STOOL()
On Error GoTo MSG
    upd_5_qr(0) = txt_stoolcolour.Text
    upd_5_qr(1) = txt_consistency.Text
    upd_5_qr(2) = txt_mucus.Text
    upd_5_qr(3) = txt_sblood.Text
    upd_5_qr(4) = txt_parasties.Text
    upd_5_qr(5) = txt_occult.Text
    upd_5_qr(6) = txt_Ova.Text
    upd_5_qr(7) = txt_cysts.Text
    upd_5_qr(8) = txt_larvae.Text
    upd_5_qr(9) = txt_puscells.Text
    upd_5_qr(10) = txt_stoolrbc.Text
    upd_5_qr(11) = txt_microphages.Text
    upd_5_qr(12) = txt_starch.Text
    upd_5_qr(13) = txt_fat.Text
    upd_5_qr(14) = txt_stoolother.Text
    If Len(txt_stoolcomments.Text) > 0 Then
        upd_5_qr(15) = Null
        upd_5_qr(15) = txt_stoolcomments.Text
    Else
        upd_5_qr(15) = Null
        upd_5_qr(15) = " "
    End If
    upd_5_qr(16) = mod_reqno
    upd_5_qr.Execute
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub STAFF()
On Error GoTo MSG
If new_flag Then
            ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = mod_reqno ' VAR_REQNNO
               Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
               SELICU_FLAG = True
            Else
               SELICU_QR(0) = mod_reqno 'VAR_REQNNO
               SELICU_RS.MoreResults
               SELICU_RS.Requery
            End If
            If Not SELICU_RS.EOF Then
               If Trim(SELICU_RS("ICDS2")) = "YES" Then
                  lblicu.Visible = True
               Else
                  lblicu.Visible = False
               End If
            End If
            ''******ICU****************
    
        SEL_dreq_staff_qr(0) = mod_reqno
        SEL_dreq_staff_qr(1) = MaskEdBox1.FormattedText
        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
        If SEL_dreq_staff_rs.RowCount > 0 Then
            txt_name = SEL_dreq_staff_rs("PNAME")
            txt_sex = SEL_dreq_staff_rs("SEX")
            txt_age = SEL_dreq_staff_rs("AGE")
            If SEL_dreq_staff_rs("RELATION") <> "STAFF" Then
                txt_status = "DP"
            Else
                txt_status = "ST"
            End If
        End If
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = reqno
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
        INSERT
 ElseIf VIEW_FLAG Then
            ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = mod_reqno ' VAR_REQNNO
               Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
               SELICU_FLAG = True
            Else
               SELICU_QR(0) = mod_reqno 'VAR_REQNNO
               SELICU_RS.MoreResults
               SELICU_RS.Requery
            End If
            If Not SELICU_RS.EOF Then
               If Trim(SELICU_RS("ICDS2")) = "YES" Then
                  lblicu.Visible = True
               Else
                  lblicu.Visible = False
               End If
            End If
            ''******ICU****************
       
        SEL_dreq_staff_qr(0) = mod_reqno
        SEL_dreq_staff_qr(1) = MaskEdBox1.Text
        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
        If SEL_dreq_staff_rs.RowCount > 0 Then
            txt_name = SEL_dreq_staff_rs("PNAME")
            txt_sex = SEL_dreq_staff_rs("SEX")
            txt_age = SEL_dreq_staff_rs("AGE")
            If SEL_dreq_staff_rs("RELATION") <> "STAFF" Then
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
        RETRIVE
 End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub INSERT()
On Error GoTo MSG
enable_field
        getupperpanel
        If Trim(txt_modcode.Text) = "BA" Then 'for case no. 1(bacteriology)
            set_visible "frm_bacteria"
        ElseIf Trim(txt_modcode.Text) = "CM" Then ' for case no. 5(clinical micro)
            If Trim(CMB_PROC.Text) = "PREGNANCY TEST (URINARY B-HCG)" Then
                set_visible "frm_others"
                Exit Sub
            End If
            If Trim(CMB_PROC.ListCount) > 1 Then
                set_visible "frm_stoolexam"
            ElseIf Trim(CMB_PROC.Text) = "URINE EXAMINATION" Then
                    set_visible "frm_urineexam"
            Else
                    set_visible "frm_stoolexam"
            End If
        ElseIf Trim(txt_modcode.Text) = "MB" Then
                
                If Trim(CMB_PROC.Text) = "ADA LEVEL" Then
                    set_visible "frm_others"
                    Exit Sub
                Else
                    set_visible "frm_myco"
                    get_myco
                End If
        Else
        ' for other case
                     
                txt_addendum.Enabled = False
                cmb_addendumby.Enabled = False
        End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub RETRIVE()
On Error GoTo MSG
Dim ADDENDUM As String
getupperpanel
        GET_ENTER_BY
        If S2 = "BA" Then
            set_visible "frm_bacteria"
            getresult_bacteriology
        
        
        Else
            set_visible "frm_others"
            getresult_other
        End If
                
            If SELDADDENDUMFLAG = False Then
                SELDADDENDUM_QR(0) = mod_reqno
                Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
                SELDADDENDUMFLAG = True
            Else
                SELDADDENDUM_QR(0) = mod_reqno
                SELDADDENDUM_RS.MoreResults
                SELDADDENDUM_RS.Requery
            End If
            
            If SELDADDENDUM_RS.RowCount > 0 Then
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
        
      disable_field 'disable all fields
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub txt_urinecomment_GOTFOCUS()
On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    objDoc.Content = txt_urinecomment.Text
    objWord.Visible = True
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub txt_urinecomment_LOSTFOCUS()
On Error GoTo MSG
    txt_urinecomment.Text = objDoc.Content
    objDoc.Content = ""
    objDoc.Close (wdDoNotSaveChanges)

    objWord.quit
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If

End Sub

Private Sub txt_urinother_GOTFOCUS()

On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    objDoc.Content = txt_urinother.Text
    objWord.Visible = True
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub

Private Sub txt_urinother_LOSTFOCUS()
On Error GoTo MSG
    txt_urinother.Text = objDoc.Content
    objDoc.Content = ""
    ' Modified By Pankaj Baggal 03/09/2004
    'objDoc.SaveAs "c:\windows\system\frozen.doc"
    objDoc.Close (wdDoNotSaveChanges)
    objWord.quit
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Public Sub printpage(moda As String, mskval As String, reqval As String)
Me.Hide
'If moda = "BA" Then
    'cmd_view_Click
    VIEW_FLAG = True
    new_flag = False
    callstatus = True
    If IsNumeric(mskval) Then
        opt_staff.Value = True
    Else
        Opt_patient.Value = True
    End If
    txt_deptcode.Text = "GZZ"
    txt_modcode.Text = moda
    MaskEdBox1.Text = mskval
    cmb_modality_Click
    
    GZZREQPATH_QR(0) = req_chglen(req_sta)
    Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
    
    cmb_reqno.AddItem req_chglen(reqval) + "$" + GZZREQPATH_RS("pathno")
    
    cmb_reqno.Text = req_chglen(reqval) + "$" + GZZREQPATH_RS("pathno")
    cmb_reqno_Click
    cmd_print_Click

End Sub

