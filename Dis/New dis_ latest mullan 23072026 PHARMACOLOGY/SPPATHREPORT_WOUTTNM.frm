VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form SPPATHREPORT_WOUTTNM 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame common 
      Height          =   2535
      Left            =   240
      TabIndex        =   77
      Top             =   960
      Width           =   11895
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   95
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
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
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
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
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   94
         Top             =   0
         Width           =   0
      End
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "SPPATHREPORT_WOUTTNM.frx":0000
         Left            =   6240
         List            =   "SPPATHREPORT_WOUTTNM.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   8280
         TabIndex        =   92
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   2400
         TabIndex        =   91
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   10440
         TabIndex        =   90
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   7920
         TabIndex        =   89
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   4920
         TabIndex        =   88
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   87
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   10440
         TabIndex        =   86
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   7920
         TabIndex        =   85
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   5520
         TabIndex        =   84
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   2400
         TabIndex        =   83
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   2400
         TabIndex        =   82
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   10200
         TabIndex        =   81
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
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
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   10200
         TabIndex        =   79
         Top             =   240
         Width           =   1335
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   240
         Width           =   2415
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   98
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
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
         Height          =   195
         Left            =   6360
         TabIndex        =   115
         Top             =   2160
         Width           =   1830
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
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
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
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
         Height          =   195
         Left            =   9120
         TabIndex        =   113
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   112
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
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
         Height          =   195
         Left            =   6960
         TabIndex        =   111
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_ward 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   110
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
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
         Height          =   195
         Left            =   240
         TabIndex        =   109
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Height          =   195
         Left            =   9720
         TabIndex        =   108
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
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
         Height          =   195
         Left            =   6960
         TabIndex        =   107
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
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
         Height          =   195
         Left            =   5040
         TabIndex        =   106
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
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
         Height          =   195
         Left            =   1920
         TabIndex        =   105
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Height          =   195
         Left            =   1680
         TabIndex        =   104
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Path No."
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
         Height          =   195
         Left            =   9360
         TabIndex        =   103
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
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
         Height          =   195
         Left            =   4920
         TabIndex        =   102
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
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
         Height          =   195
         Left            =   1440
         TabIndex        =   101
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
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
         Height          =   195
         Left            =   9120
         TabIndex        =   100
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
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
         Height          =   195
         Left            =   360
         TabIndex        =   99
         Top             =   360
         Width           =   1830
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   855
      Left            =   120
      TabIndex        =   67
      Top             =   7680
      Width           =   11775
      Begin VB.CommandButton cmd_fullreport 
         Caption         =   "Full&Report"
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
         Left            =   10080
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":0004
         MousePointer    =   99  'Custom
         TabIndex        =   76
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
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
         Left            =   6360
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":030E
         MousePointer    =   99  'Custom
         TabIndex        =   75
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
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
         Left            =   8880
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":0618
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
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
         Left            =   7680
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":0922
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
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
         Left            =   5160
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":0C2C
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
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
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":0F36
         MousePointer    =   99  'Custom
         TabIndex        =   71
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
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
         Left            =   2760
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":1240
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
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
         Left            =   1560
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":154A
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
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
         Left            =   360
         MouseIcon       =   "SPPATHREPORT_WOUTTNM.frx":1854
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10800
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10800
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmb_enteryby 
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
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   6960
      Width           =   2895
   End
   Begin VB.ComboBox cmb_provby 
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
      ItemData        =   "SPPATHREPORT_WOUTTNM.frx":1B5E
      Left            =   2040
      List            =   "SPPATHREPORT_WOUTTNM.frx":1B60
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   7320
      Width           =   2895
   End
   Begin VB.TextBox txt_fullreport 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8655
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   7320
      Width           =   1095
   End
   Begin VB.ComboBox cmb_addendumby 
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
      Left            =   8640
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   7320
      Width           =   3135
   End
   Begin VB.ComboBox cmb_finalby 
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
      Left            =   8640
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9120
      Top             =   120
   End
   Begin VB.CommandButton cmd_emr 
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
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CheckBox chk_Word 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Word Doc"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock WIN_CLIENT 
      Left            =   9120
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   9
      Top             =   3600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Screen I"
      TabPicture(0)   =   "SPPATHREPORT_WOUTTNM.frx":1B62
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_prepathreqno"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_opprocedure"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Grossdesc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_frosec"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_microdesc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl_immhischerep"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_Immhische"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_immhischereq"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_natmeterial"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LBL_FNACREP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LBL_VPP"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CMB_PATHNO"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_opprocedure"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_grossdesc"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_frosecrep"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_immhischereqno"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_microdesc"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_immhischerep"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmb_immhische"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmb_grossby"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_natmeterial"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_fnacreport"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Screen II"
      TabPicture(1)   =   "SPPATHREPORT_WOUTTNM.frx":1B7E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_register"
      Tab(1).Control(1)=   "Shape1"
      Tab(1).Control(2)=   "LBL_CM"
      Tab(1).Control(3)=   "lbl_Grade"
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(7)=   "Label4"
      Tab(1).Control(8)=   "Label3"
      Tab(1).Control(9)=   "Label2"
      Tab(1).Control(10)=   "lbl_addendum1"
      Tab(1).Control(11)=   "lbl_impression"
      Tab(1).Control(12)=   "lbl_laterality"
      Tab(1).Control(13)=   "lbl_sites"
      Tab(1).Control(14)=   "Label7"
      Tab(1).Control(15)=   "Label8"
      Tab(1).Control(16)=   "Label10"
      Tab(1).Control(17)=   "lbl_fsd"
      Tab(1).Control(18)=   "cmb_reg"
      Tab(1).Control(19)=   "txt_size"
      Tab(1).Control(20)=   "cmb_t"
      Tab(1).Control(21)=   "cmb_m"
      Tab(1).Control(22)=   "CMB_N"
      Tab(1).Control(23)=   "CMB_ICDH"
      Tab(1).Control(24)=   "txt_addendum"
      Tab(1).Control(25)=   "txt_impfindiag"
      Tab(1).Control(26)=   "cmb_grade"
      Tab(1).Control(27)=   "txt_laterality"
      Tab(1).Control(28)=   "txt_sites"
      Tab(1).Control(29)=   "cmb_sites"
      Tab(1).Control(30)=   "cmb_laterality"
      Tab(1).Control(31)=   "txt_prepathreqno"
      Tab(1).Control(32)=   "cmb_pendingremarks"
      Tab(1).Control(33)=   "TXT_FSDIS"
      Tab(1).Control(34)=   "cmb_fsd"
      Tab(1).ControlCount=   35
      Begin VB.TextBox txt_fnacreport 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   37
         Top             =   1680
         Width           =   7095
      End
      Begin VB.TextBox txt_natmeterial 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   36
         Top             =   960
         Width           =   3735
      End
      Begin VB.ComboBox cmb_grossby 
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
         Left            =   2520
         TabIndex        =   35
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ComboBox cmb_immhische 
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
         Left            =   2535
         TabIndex        =   34
         Top             =   2790
         Width           =   735
      End
      Begin VB.TextBox txt_immhischerep 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   33
         Top             =   2760
         Width           =   6495
      End
      Begin VB.TextBox txt_microdesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7290
         TabIndex        =   32
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox txt_immhischereqno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_frosecrep 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   30
         Top             =   1320
         Width           =   7095
      End
      Begin VB.TextBox txt_grossdesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   29
         Top             =   2040
         Width           =   7095
      End
      Begin VB.TextBox txt_opprocedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   28
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cmb_reg 
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
         Left            =   -73200
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txt_size 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67695
         TabIndex        =   26
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox cmb_t 
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
         Left            =   -68535
         TabIndex        =   25
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox cmb_m 
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
         Left            =   -65655
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox CMB_N 
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
         Left            =   -66615
         TabIndex        =   23
         Top             =   1800
         Width           =   585
      End
      Begin VB.ComboBox CMB_ICDH 
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
         Left            =   -73200
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -73200
         TabIndex        =   21
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox txt_impfindiag 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73200
         TabIndex        =   20
         Top             =   1320
         Width           =   8055
      End
      Begin VB.ComboBox cmb_grade 
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
         Left            =   -73200
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txt_laterality 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -65400
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txt_sites 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cmb_sites 
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
         Left            =   -73200
         TabIndex        =   16
         Top             =   480
         Width           =   6015
      End
      Begin VB.ComboBox cmb_laterality 
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
         Left            =   -65880
         TabIndex        =   15
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txt_prepathreqno 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -70320
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CMB_PATHNO 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2520
         TabIndex        =   13
         Top             =   480
         Width           =   7095
      End
      Begin VB.ComboBox cmb_pendingremarks 
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
         Left            =   -67320
         TabIndex        =   12
         Top             =   2760
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox TXT_FSDIS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71760
         TabIndex        =   11
         Top             =   2280
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cmb_fsd 
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
         ItemData        =   "SPPATHREPORT_WOUTTNM.frx":1B9A
         Left            =   -73200
         List            =   "SPPATHREPORT_WOUTTNM.frx":1BA4
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label LBL_VPP 
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label LBL_FNACREP 
         AutoSize        =   -1  'True
         Caption         =   "FNAC Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1200
         TabIndex        =   65
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label lbl_natmeterial 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Material Received"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5400
         TabIndex        =   64
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grossed By"
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
         Height          =   195
         Left            =   1320
         TabIndex        =   63
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lbl_immhischereq 
         AutoSize        =   -1  'True
         Caption         =   "Immhistchem Req. No."
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
         Height          =   195
         Left            =   5280
         TabIndex        =   62
         Top             =   1320
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label lbl_Immhische 
         AutoSize        =   -1  'True
         Caption         =   "Immunohistochemistry"
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
         Height          =   195
         Left            =   435
         TabIndex        =   61
         Top             =   2760
         Width           =   1860
      End
      Begin VB.Label Lbl_immhischerep 
         AutoSize        =   -1  'True
         Caption         =   "Immhistchem Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3360
         TabIndex        =   60
         Top             =   2760
         Width           =   1725
      End
      Begin VB.Label lbl_microdesc 
         AutoSize        =   -1  'True
         Caption         =   "Microscopic Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5160
         TabIndex        =   59
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lbl_frosec 
         AutoSize        =   -1  'True
         Caption         =   "Frozen Secion Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   480
         TabIndex        =   58
         Top             =   1320
         Width           =   1860
      End
      Begin VB.Label lbl_Grossdesc 
         AutoSize        =   -1  'True
         Caption         =   "Gross Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   840
         TabIndex        =   57
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label lbl_opprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Operative Procedure/s"
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
         Height          =   195
         Left            =   360
         TabIndex        =   56
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label lbl_register 
         AutoSize        =   -1  'True
         Caption         =   "Reg/Res Doc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74520
         TabIndex        =   55
         Top             =   2280
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   -68775
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label LBL_CM 
         Caption         =   "cm"
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
         Height          =   255
         Left            =   -67095
         TabIndex        =   54
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lbl_Grade 
         AutoSize        =   -1  'True
         Caption         =   "Grade"
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
         Height          =   195
         Left            =   -73800
         TabIndex        =   53
         Top             =   1800
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Pathology Stage"
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
         Height          =   195
         Left            =   -70200
         TabIndex        =   52
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Size"
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
         Height          =   195
         Left            =   -68055
         TabIndex        =   51
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "M"
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
         Height          =   195
         Left            =   -65895
         TabIndex        =   50
         Top             =   1920
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N"
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
         Height          =   195
         Left            =   -66735
         TabIndex        =   49
         Top             =   1920
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T"
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
         Height          =   195
         Left            =   -68655
         TabIndex        =   48
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Histology Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74520
         TabIndex        =   47
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lbl_addendum1 
         AutoSize        =   -1  'True
         Caption         =   "Addendum "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74280
         TabIndex        =   46
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label lbl_impression 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74160
         TabIndex        =   45
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lbl_laterality 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
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
         Height          =   195
         Left            =   -66840
         TabIndex        =   44
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lbl_sites 
         AutoSize        =   -1  'True
         Caption         =   "Final Site Diag."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74520
         TabIndex        =   43
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label7 
         Caption         =   "*"
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   42
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "*"
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   840
         Width           =   135
      End
      Begin VB.Label lbl_prepathreqno 
         AutoSize        =   -1  'True
         Caption         =   "Prev Path No./ Req No."
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
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   2070
      End
      Begin VB.Label Label10 
         Caption         =   "Pending Remarks"
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
         Left            =   -68880
         TabIndex        =   39
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbl_fsd 
         AutoSize        =   -1  'True
         Caption         =   "F.S.Disprcy"
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
         Height          =   195
         Left            =   -74280
         TabIndex        =   38
         Top             =   2280
         Width           =   990
      End
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
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   3600
      TabIndex        =   125
      Top             =   600
      Width           =   4845
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
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
      Height          =   195
      Left            =   10200
      TabIndex        =   124
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Height          =   195
      Left            =   10200
      TabIndex        =   123
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl_enterby 
      AutoSize        =   -1  'True
      Caption         =   "Entered  By"
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
      Height          =   195
      Left            =   120
      TabIndex        =   122
      Top             =   6960
      Width           =   1005
   End
   Begin VB.Label lbl_pro 
      AutoSize        =   -1  'True
      Caption         =   "Resident Pathologist"
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
      Height          =   195
      Left            =   120
      TabIndex        =   121
      Top             =   7320
      Width           =   1770
   End
   Begin VB.Label lbl_final 
      AutoSize        =   -1  'True
      Caption         =   "Reported By"
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
      Height          =   195
      Left            =   7440
      TabIndex        =   120
      Top             =   6960
      Width           =   1065
   End
   Begin VB.Label lbl_addendum 
      AutoSize        =   -1  'True
      Caption         =   "Addendum By"
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
      Height          =   195
      Index           =   0
      Left            =   7320
      TabIndex        =   119
      Top             =   7320
      Width           =   1170
   End
   Begin VB.Label Label11 
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
      Left            =   120
      TabIndex        =   118
      Top             =   0
      Width           =   11655
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Height          =   195
      Left            =   10200
      TabIndex        =   117
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label13 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Click here to open word on modify)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   360
      TabIndex        =   116
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "SPPATHREPORT_WOUTTNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim MODADD As String
Dim str As String
Dim FSFLAG As Boolean
Dim sel_deptqr As rdoQuery
Dim UPDSPREPORT_QR As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SEL_DEPTFLAG As Boolean
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim SEL_MODALITYFLAG As Boolean
Dim sel_casno As rdoQuery
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim REQFLAG As Boolean
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim SELTA101FLAG As Boolean
Dim repdr_qr As rdoQuery
Dim repdr_rs As rdoResultset
Dim REPDRFLAG As Boolean
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim REFDRFLAG As Boolean
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset
Dim selma104flag As Boolean
Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset
Dim SELPATHDREQUISITNFLAG As Boolean
Dim selma302_qr As rdoQuery
Dim selma302_rs As rdoResultset
Dim SELMA302FLAG As Boolean
Dim seldfsreport_qr As rdoQuery
Dim seldfsreport_rs As rdoResultset
Dim SELDFSREPORTFLAG As Boolean
Dim seldihreport_qr As rdoQuery
Dim seldihreport_rs As rdoResultset
Dim seldihreportflag As Boolean
Dim seldimpresion_qr As rdoQuery
Dim seldimpresion_rs As rdoResultset
Dim SELDIMPRESIONFLAG As Boolean
Dim seldgrossdesc_qr As rdoQuery
Dim seldgrossdesc_rs As rdoResultset
Dim SELDGROSSDESCFLAG As Boolean
Dim seldspreport_qr As rdoQuery
Dim seldspreport_rs As rdoResultset
Dim SELDSPREPORTFLAG As Boolean
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELTA1012FLAG As Boolean
Dim w_frozen As Application, W_grossdes As Application, w_micscop As Application, w_imp As Application, w_immun As Application, w_adden As Application, w As Application
Dim d_frozen As Document, d_grossdesc As Document, d_micscop As Document, d_imp As Document, d_immun As Document, d_adden As Document, d As Document, D1 As Document
Dim w_meterial As Application
Dim d_meterial As Document
Dim WFNAC As Application
Dim DFNAC As Document

Dim SELEMPLIB_QR As rdoQuery
Dim SELEMPLIB_RS As rdoResultset

Dim UPDTADNDMTIME_QR As rdoQuery

Dim i As Integer
Dim insdfsrep_qr As rdoQuery
Dim insdgrossdesc_qr As rdoQuery
Dim insdfs_qr As rdoQuery
Dim insdimpresion_qr As rdoQuery
Dim insdspreport_qr As rdoQuery
Dim INSDIHREPORT_QR As rdoQuery
Dim upddfsreport_qr As rdoQuery
Dim upddihreport_qr As rdoQuery
Dim upddimpresion_qr As rdoQuery
Dim upddspreport_qr As rdoQuery
'Dim upddimpresion_qr As rdoQuery
Dim updgrossdesc_qr As rdoQuery
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim selempbasinfo1_qr As rdoQuery
Dim selempbasinfo1_rs As rdoResultset
Dim SELEMPBASINFO1FLAG As Boolean
Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim UPDRREQUISITN1FLAG As Boolean
Dim INSDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUMFLAG As Boolean
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG As Boolean
Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer
Dim ADDENDUMFALG As Boolean
Dim SELMA301QR As rdoQuery
Dim SELMA301RS As rdoResultset

Dim SELMA301FLAG As Boolean
Dim INSFREPORT_QR As rdoQuery
Dim UPDFREPORT_QR As rdoQuery
Dim SELRREQNOQR As rdoQuery
Dim SELRREQNORS As rdoResultset
Dim SELRREQNOFLAG As Boolean
'Dim SELDFREPORT_QR As rdoQuery
'Dim SELDFREPORT_RS As rdoResultset
'Dim seldfreportflag As Boolean
Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset
Dim seldfreportflag As Boolean

Dim SELDFREPORT_IHQR As rdoQuery
Dim SELDFREPORT_IHRS As rdoResultset
Dim SELDFREPORT_IHFLAG As Boolean
Dim upddrequisitn_qr As rdoQuery

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset

Dim seldgradeqr As rdoQuery
Dim seldgraders As rdoResultset
Dim SELDGRADEFLAG As Boolean
Dim seldgrade_qr As rdoQuery
Dim seldgrade_rs As rdoResultset
Dim SELDGRADE_FLAG As Boolean
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
'Dim As rdoResultset
Dim SELDREQNOFLAG As Boolean
Dim selihreportqr As rdoQuery
Dim selihreportrs As rdoResultset
Dim SELIHREPORTFLAG As Boolean

Dim seldissecentr_qr As rdoQuery
Dim seldissecentr_rs As rdoResultset
Dim seldissecentrflag As Boolean

Dim seldissecprov_qr As rdoQuery
Dim seldissecprov_rs As rdoResultset
Dim seldissecprovflag As Boolean

Dim seldissecfin_qr As rdoQuery
Dim seldissecfin_rs As rdoResultset
Dim seldissecfinflag As Boolean

Dim seldissecadd_qr As rdoQuery
Dim seldissecadd_rs As rdoResultset
Dim seldissecaddrflag As Boolean

Dim seldissecgross_qr As rdoQuery
Dim seldissecgross_rs As rdoResultset
Dim seldissecgrossflag As Boolean

Dim SELDREQSTAFFD1QR1 As rdoQuery
Dim SELDREQSTAFFD1RS1 As rdoResultset
Dim SELDREQSTAFFD1FLAG1 As Boolean

Dim viewsitedupqr As rdoQuery
Dim viewsiteduprs As rdoResultset
Dim VIEWSITEDUPRSFLAG As Boolean
Dim INSFNREPORTQR As rdoQuery
Dim UPDFNREPORTQR As rdoQuery
Dim INSDSPFNREQQR As rdoQuery

Dim SELDSPFN_QR As rdoQuery
Dim SELDSPFN_RS As rdoResultset
Dim SELDSPFNFLAG As Boolean

Dim SELDFNREPORTQR As rdoQuery
Dim SELDFNREPORTRS As rdoResultset
Dim SELDFNREPORTFLAG As Boolean
Dim SPFNFLAG As Boolean

Dim INSDISIMP_QR As rdoQuery
Dim UPDDISIMP_QR As rdoQuery

Dim UPDTCMT_QR As rdoQuery

Dim SELDSITE_QR As rdoQuery
Dim SELDSITE_RS As rdoResultset

Dim SELMA10122_QR As rdoQuery
Dim SELMA10122_RS As rdoResultset

Dim INSSITEDETAIL_QR As rdoQuery

Dim UPDSITEDETAIL_QR As rdoQuery

Dim SELSITEDETAIL_QR As rdoQuery
Dim SELSITEDETAIL_RS As rdoResultset
Dim SELSITEDETAIL_FLAG As Boolean

Dim SELSITEDETAIL_QR1 As rdoQuery
Dim SELSITEDETAIL_RS1 As rdoResultset
Dim SELSITEDETAIL_FLAG1 As Boolean

Dim seldrequisitn_qr10 As rdoQuery
Dim seldrequisitn_rs10 As rdoResultset
Dim seldrequisitn_flag10 As Boolean

Dim seldrequisitn_qr11 As rdoQuery
Dim seldrequisitn_rs11 As rdoResultset
Dim seldrequisitn_flag11 As Boolean

Dim MODALITY1
Dim DEPT1
Dim REQNO1
Dim MODDESC

Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset

Dim selsite_qr As rdoQuery
Dim selsite_rs As rdoResultset
Dim SELSITE_FLAG As Boolean

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
'' H7QUERIES
Dim SELPACSLOG_RQ As rdoQuery
Dim UPDPACSLOGPREP_RQ As rdoQuery
Dim UPDPACSLOGFREP_RQ As rdoQuery
Dim UPDPACSLOGADD_RQ As rdoQuery
Dim INSHL7TAB3_RQ As rdoQuery
Dim SELPACSLOG_RS As rdoResultset

'' HL7QUERIES
            '***for HL7*****24122003
                Dim insas400_hl7_rq As rdoQuery
                Dim caseno_hl7 As String
                Dim reqnno_hl7 As String
             
                Dim insas400_hl7_rs As rdoResultset
                Dim DTTIM As rdoResultset
                Dim msgtype_hl7 As String
                
                Dim str_hl7 As String
                Dim uniqueid_hl7 As String
              '*******hl7 declarations till here

Private Sub Check1_Click()

End Sub

Private Sub chk_frozen_Click()
'Dim upds1 As String
'If chk_frozen.Value = 1 Then
'upds1 = "update dislib.drequisitn set FROZENSEC='YES' where reqno='" & cmb_reqno.Text & "'"
'Else
'upds1 = "update dislib.drequisitn set FROZENSEC=null where reqno='" & cmb_reqno.Text & "'"
'End If
'rdb.Execute upds1
'txt_frosecrep.Enabled = True
'txt_frosecrep.SetFocus
End Sub

''******SECURITY***********





Public Sub cmb_dept_Click()
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

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "FZZ" Then
           cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub





Private Sub cmb_dept_LostFocus()
cmb_modality.ListIndex = 0

End Sub

Private Sub cmb_fsd_Click()

If Trim(cmb_fsd.Text) = "YES" Then
    TXT_FSDIS.Visible = True
Else
    TXT_FSDIS.Visible = False
End If

End Sub

Private Sub CMB_ICDH_Click()
    If Trim(CMB_ICDH.Text) <> "" Then
        If Trim(CMB_ICDH.Text) = "UNKNOWN" Then
         '   txt_impfindiag.Text = ""
        End If
    End If
'If Len(Trim(txt_impfindiag.Text)) > 0 Then
'    txt_impfindiag.Text = txt_impfindiag & " " & Trim(CMB_ICDH.Text)
'Else
'    txt_impfindiag.Text = Trim(CMB_ICDH.Text)
'End If
End Sub

Private Sub cmb_immhische_Click()
On Error GoTo MSG:
 If (cmb_immhische.Text) = "YES" Then
    IMHrequisation.show 1
    txt_immhischerep.Visible = True
    Lbl_immhischerep.Visible = True
    txt_immhischerep.Enabled = True
 End If
 If (cmb_immhische.Text) = "NO" Then
     txt_immhischerep.Visible = False
     Lbl_immhischerep.Visible = False
 End If
'*** commented by chandu 3/01/2003
''If (cmb_immhische.Text = "NO") Or Trim(Len(cmb_immhische.Text) < 1) Then
''    txt_immhischerep.Enabled = False
''End If
''If cmb_immhische.Text = "YES" Then
''
''    pathology.show
''
'''    pathology.cmb_modality.Text = "IMMUNOHISTOCHEMISTRY&IH"
''    'IMHrequisation.Show
''    pathology.immunohistochemistry.Visible = True
''    pathology.surgicalPathology.Visible = False
''    pathology.lbl_req.Left = 3850
''    pathology.lbl_req.Caption = "IMMUNOHISTOCHEMISTRY"
''    pathology.Caption = "Pathology - Immunohistochemistry"
''    IHFLAG = True
''    pathology.txt_sppathreqno = cmb_reqno
''    pathology.txt_sppathreqno.Enabled = False
''    txt_immhischerep.Enabled = True
''    'pathology.btnenabl ("NEW")
''End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_laterality_GotFocus()
Set SELMA10122_RS = SELMA10122_QR.OpenResultset(1, 3)
        If Not SELMA10122_RS.EOF Then
            cmb_laterality.clear
            Do While Not SELMA10122_RS.EOF
                cmb_laterality.AddItem SELMA10122_RS(0)
                SELMA10122_RS.MoveNext
            Loop
        End If
End Sub

Private Sub cmb_laterality_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub cmb_m_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_m) = "" Then
    For ind_cbodept = 0 To (cmb_m.ListCount - 1)
        If Trim(cmb_m.List(ind_cbodept)) = " " Then
            cmb_m.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Public Sub cmb_modality_Click()
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
'If txt_modcode.Text = "IH" Then
'    txt_deptcode.Text = "FDZ"
'End If
'        If seldissecentrflag = False Then
'            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
'            seldissecentrflag = True
'        Else
'            seldissecentr_rs.MoreResults
'            seldissecentr_rs.Requery
'        End If
'        If seldissecentr_rs.RowCount > 0 Then
'            cmb_enteryby.clear
'            While Not seldissecentr_rs.EOF
'                cmb_enteryby.AddItem Trim(seldissecentr_rs("CCNAME"))
'                seldissecentr_rs.MoveNext
'            Wend
'        End If
'        If seldissecprovflag = False Then
'            Set seldissecprov_rs = seldissecprov_qr.OpenResultset(1, 3)
'            seldissecprovflag = True
'        Else
'            seldissecprov_rs.MoreResults
'            seldissecprov_rs.Requery
'        End If
'        If seldissecprov_rs.RowCount > 0 Then
'            cmb_provby.clear
'            While Not seldissecprov_rs.EOF
'                cmb_provby.AddItem Trim(seldissecprov_rs("CCNAME"))
'                seldissecprov_rs.MoveNext
'            Wend
'        End If
'        If seldissecfinflag = False Then
'            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
'            seldissecfinflag = True
'        Else
'            seldissecfin_rs.MoreResults
'            seldissecfin_rs.Requery
'        End If
'        If seldissecfin_rs.RowCount > 0 Then
'            cmb_finalby.clear
'            cmb_reg.clear
'            While Not seldissecfin_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
'                cmb_reg.AddItem Trim(seldissecfin_rs("ccname"))
'                seldissecfin_rs.MoveNext
'            Wend
'        End If
'        If seldissecgrossflag = False Then
'            Set seldissecgross_rs = seldissecgross_qr.OpenResultset(1, 3)
'            seldissecgrossflag = True
'        Else
'            seldissecgross_rs.MoreResults
'            seldissecgross_rs.Requery
'        End If
'        If seldissecgross_rs.RowCount > 0 Then
'            cmb_grossby.clear
'            While Not seldissecgross_rs.EOF
'                cmb_grossby.AddItem Trim(seldissecgross_rs("CCNAME"))
'                seldissecgross_rs.MoveNext
'            Wend
'        End If




Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "SURGICAL PATHOLOGY$SP" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub CMB_N_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(CMB_N) = "" Then
    For ind_cbodept = 0 To (CMB_N.ListCount - 1)
        If Trim(CMB_N.List(ind_cbodept)) = "X" Then
            CMB_N.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub CMB_PATHNO_Click()

X = InStr(12, Trim(CMB_PATHNO.Text), "$")
REQNO1 = Mid(Trim(CMB_PATHNO.Text), 1, X - 1)
DEPT1 = Mid(Trim(REQNO1), 1, 3)
MODALITY1 = Mid(Trim(REQNO1), 4, 2)
If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
    SPREPORT
ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
    FAREPORT
ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
    CGCNREPORT
ElseIf DEPT1 = "LZZ" And (MODALITY1 = "MM" Or MODALITY1 = "MR" Or MODALITY1 = "CR" Or MODALITY1 = "CT" Or MODALITY1 = "US") Then
    RISREPORT
End If

End Sub


Private Sub cmb_sites_GotFocus()
Set SELDSITE_RS = SELDSITE_QR.OpenResultset(1, 3)
If Not SELDSITE_RS.EOF Then
  cmb_sites.clear
  Do While Not SELDSITE_RS.EOF
      cmb_sites.AddItem SELDSITE_RS("SDESC")
      SELDSITE_RS.MoveNext
  Loop
End If
cmb_sites.AddItem "Not Specified"
End Sub

Private Sub cmb_sites_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub cmb_t_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_t) = "" Then
    For ind_cbodept = 0 To (cmb_t.ListCount - 1)
        If Trim(cmb_t.List(ind_cbodept)) = 0 Then
            cmb_t.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmd_back_Click()
'btnenabl ("INI")
BTNENABL1 ("INI")
'FLDENABL ("DIS")
FLDENABL1 ("DIS")
TXT_CLEAR
field_clear

If WORKLIST = True Then
    'Unload Me
    Me.Hide
    disw.show
Else
    'Unload Me
    Me.Hide
    DIS_MAIN.show
End If

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
    cmd_print.Enabled = True
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
    cmd_print.Enabled = True
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
'btnenabl ("INI")
BTNENABL1 ("INI")
field_clear
TXT_CLEAR
ADDENDUMFALG = False
FSFLAG = False
'FLDENABL ("DIS")
FLDENABL1 ("DIS")
txt_immhischerep.Visible = False
Lbl_immhischerep.Visible = False
End Sub

Public Sub cmd_commit_Click()
Dim i As Integer
On Error GoTo MSG:
i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If i = vbNo Then
    Exit Sub
    cmd_commit.Enabled = False
End If
If i = vbYes Then
                caseno_hl7 = MaskEdBox1.FormattedText
                reqnno_hl7 = req_chglen(cmb_reqno.Text)
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
        Else
            If CURUSRDET_RS("usrtype") <= 6 Then
                MsgBox "Sorry! You are not authorised to Commit the Report of this Modality"
                cmd_commit.SetFocus
                Exit Sub
            End If
        End If
    Else
        MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
    End If
''********SECURITY***************************
    rdb.BeginTrans
    updrrequisitn1_qr(0) = UCase(Left(cmb_finalby.Text, 30))
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = REQNO
    updrrequisitn1_qr.Execute
    If (cmb_immhische.Text) = "YES" And Len(Trim(txt_immhischereqno.Text)) > 0 Then
        updrrequisitn1_qr(0) = UCase(Left(cmb_finalby.Text, 30))
        updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        updrrequisitn1_qr(2) = "F"
        updrrequisitn1_qr(3) = "Y"
        updrrequisitn1_qr(4) = req_chglen(txt_immhischereqno.Text)
        updrrequisitn1_qr.Execute
   End If
   
   'Format(txt_date.Text, "YYYY-MM-DD")
   'UPDTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
   UPDTCMT_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
   'GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
   UPDTCMT_QR(1) = Format(txt_time.Text, "HH:MM:SS")
   UPDTCMT_QR(2) = req_chglen(cmb_reqno.Text)
   UPDTCMT_QR.Execute
   ' hl7 code
    SELPACSLOG_RQ(0) = REQNO
    Set SELPACSLOG_RS = SELPACSLOG_RQ.OpenResultset(1, 1)
    If SELPACSLOG_RS.RowCount > 0 Then
     '********Inserting into HL7 Table****24122003****
'            Dim insas400_hl7_rq As rdoQuery
'            Dim caseno_hl7 As String
'            Dim reqnno_hl7 As String
'            caseno_hl7 = msk_caseno.FormattedText
'            reqnno_hl7 = req_chglen(cbo_reqno.Text)
'            Dim insas400_hl7_rs As rdoResultset
'
'            Dim msgtype_hl7 As String
'
'            Dim str_hl7 As String
'            Dim uniqueid_hl7 As String
            msgtype_hl7 = "Final Report" 'Form1.lbl_type.Tag
            
            str_hl7 = WIN_CLIENT.LocalIP
            '''
              Set DTTIM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE  , CURRENT TIME  FROM PATOOLSLIB.MA107") '***2/2/99
 
            
             uniqueid_hl7 = str_hl7 & "-" & Format(DTTIM(0), "dd/mm/yyyy") & "-" & Format(DTTIM(1), "hh:mm:ss")
            
            
           
            Set insas400_hl7_rq = rdb.CreateQuery("ins as400_fr", "insert into hl7lib.hl7tab3(UNQMSGID,MSGTYPID,CASENO,CLNTIP,CLNTDATE,CLNTTIME,REQNNO,PROCBRDGYN) values(?,?,?,?,?,?,?,?)")
            insas400_hl7_rq(0) = uniqueid_hl7
            insas400_hl7_rq(1) = msgtype_hl7
            insas400_hl7_rq(2) = caseno_hl7 ' msk_caseno.FormattedText
            insas400_hl7_rq(3) = WIN_CLIENT.LocalIP
            insas400_hl7_rq(4) = Format(DTTIM(0), "dd/mm/yyyy")
            insas400_hl7_rq(5) = Format(DTTIM(1), "hh:mm:ss")
            insas400_hl7_rq(6) = reqnno_hl7
            insas400_hl7_rq(7) = "N"
            insas400_hl7_rq.Execute
      
    
     '  UPDATING PACSLOG
          UPDPACSLOGFREP_RQ(0) = "Y"
          UPDPACSLOGFREP_RQ(1) = Format(DTTIM(0), "DD/MM/YYYY")
          UPDPACSLOGFREP_RQ(2) = Format(DTTIM(1), "HH:MM:SS")
          UPDPACSLOGFREP_RQ(3) = REQNO
          UPDPACSLOGFREP_RQ.Execute
        '  UPDATING PACSLOG over
    
  '*************Till Here HL7**********************************
   ' hl7 code over
  End If ' FOR PACS LOG READ
  
End If
 'MsgBox "Final Report Data successfully sent to AS/400 - HL7 Table", vbInformation
rdb.CommitTrans

'btnenabl ("INI")
BTNENABL1 ("INI")
MsgBox "Record is Committed Successfully."
If ADDENDUMFALG = True Then
sendRMSG REQNO, MaskEdBox1.FormattedText, "A"
Else
sendRMSG REQNO, MaskEdBox1.FormattedText, "F"
End If
cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub CMD_EMR_Click()

'If Len(Trim(MaskEdBox1.Text)) > 0 Then
'    Dim VAR1
'    'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
'    VAR1 = MaskEdBox1.FormattedText
'    'VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
'    VAR = Shell("C:\EMR1.EXE " & Trim(MaskEdBox1.FormattedText), vbNormalFocus)
'    AppActivate VAR
'Else
'    MsgBox "Please Enter the Case No."
'End If
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://10.1.4.13/emrweb/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If

End Sub

Private Sub cmd_fullreport_Click()
On Error GoTo MSG:
txt_fullreport_GotFocus
Set D1 = w.ActiveDocument
If Len(txt_frosecrep) > 0 Then
D1.Content = D1.Content & "FROZEN:" & Chr(13) & Trim(txt_frosecrep)
End If
If Len(txt_grossdesc.Text) > 0 Then
D1.Content = D1.Content & "GROSS DESC: " & Chr(13) & Trim(txt_grossdesc)
End If
If Len(txt_immhischerep) > 0 Then
D1.Content = D1.Content & "IMMUNOHISTO:" & Chr(13) & Trim(txt_immhischerep)
End If
If Len(txt_microdesc) > 0 Then
D1.Content = D1.Content & "MICROSCOPICDESC:" & Chr(13) & Trim(txt_microdesc)
End If
If Len(txt_impfindiag) > 0 Then
D1.Content = D1.Content & "IMPRESSION:" & Chr(13) & Trim(txt_impfindiag)
End If
If Len(txt_addendum) > 0 Then
D1.Content = D1.Content & Chr(13) & "ADDENDUM:" & Chr(13) & Trim(txt_addendum)
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_fullreport_LostFocus()
On Error GoTo MSG:

D1.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'D1.SaveAs "C:\DISREPORTSEC\REPORT.DOC"
D1.Close (wdDoNotSaveChanges)
w.quit
'Msk_caseno.Enabled = False
'cmd_cancel.SetFocus
MSG:
End Sub

Private Sub cmd_modify_Click()
''********SECURITY***************************
CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
'CURUSRDET_QR(0) = Trim(cmb_enteryby.Text) 'HERE1
CURUSRDET_RS.MoreResults
CURUSRDET_RS.Requery
'Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 3)
'CURUSRDET_RS.MoreResults
'CURUSRDET_RS.Requery
If CURUSRDET_RS.RowCount > 0 Then
    If LOGINTYP > CURUSRDET_RS("usrtype") Then
      ''   no action as he can commit
    ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
      '  MsgBox "Sorry ! This Report needs to be Modified by a Higher Level User."
       '' cmd_modify.SetFocus
       ' Exit Sub
    'ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
     '   If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
            '  If loginusrid <> curusrdetrs("usr_id") Then
      '      MsgBox "Sorry! This Report is not to be Finalized by You"
       '     cmd_commit.SetFocus
        '    Exit Sub
        'End If
    End If
    If CURUSRDET_RS("MODIFY") <> "YES" Then
        MsgBox "Sorry! You are not authorised to Modify the Report of this Modality"
        cmd_commit.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
    'cmd_commit.SetFocus
    Exit Sub
End If

''********SECURITY***************************
''********SECURITY***************************
If MODADD <> "" Then
Else
'***** Added by Shraddha on 29/01/2013 *****
'For modification in Frozen, Gross and Microscopic fields.
If chk_Word.Value = False Then
frm_SubPath.show


frm_SubPath.txt_gross.Text = Trim(SPPATHREPORT.txt_grossdesc.Text)
frm_SubPath.txt_microscopic.Text = SPPATHREPORT.txt_microdesc.Text
frm_SubPath.txt_Frozen.Text = SPPATHREPORT.txt_frosecrep.Text
frm_SubPath.txt_Impres.Text = SPPATHREPORT.txt_impfindiag.Text
frm_SubPath.Label5.Caption = "Case No : " & Trim(SPPATHREPORT.MaskEdBox1.FormattedText) & " Req No. : " & Trim(SPPATHREPORT.cmb_reqno.Text) & vbCrLf & " Patient Name : " & Trim(SPPATHREPORT.txt_name.Text)
frm_SubPath.txt_nmat.Text = txt_natmeterial.Text
SPPATHREPORT.FLDENABL1 ("VIEW")
End If
End If
'*******************************
'btnenabl ("NEW")
BTNENABL1 ("NEW")
'FLDENABL ("MODIFY")
FLDENABL1 ("MODIFY")
 flagsave = 2
If (cmb_immhische.Text = "NO") Or Trim(Len(cmb_immhische.Text) < 1) Then
    txt_immhischerep.Enabled = False
Else
    txt_immhischerep.Enabled = True
End If
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
If REFDR_RS.RowCount > 0 Then
'If REFDR_RS("ser_code") = "F104" Or REFDR_RS("SER_CODE") = "F105" Then
'    txt_grossdesc.Enabled = False
'    cmb_grossby.Enabled = False
'End If
End If
If REFDR_RS.RowCount > 0 Then
    If Trim(REFDR_RS("reqnstat")) = "F" And Trim(REFDR_RS("COMMITFLAG")) = "Y" Then
    ADDENDUMFALG = True
    End If
    
End If

'        If seldissecentrflag = False Then
'            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
'            seldissecentrflag = True
'        Else
'            seldissecentr_rs.MoreResults
'            seldissecentr_rs.Requery
'            'seldissecadd_rs.Requery
'        End If
'        If seldissecentr_rs.RowCount > 0 Then
'            cmb_enteryby.clear
'            While Not seldissecentr_rs.EOF
'                cmb_enteryby.AddItem Trim(seldissecentr_rs("CCNAME"))
'                seldissecentr_rs.MoveNext
'            Wend
'        End If
'        If seldissecprovflag = False Then
'            Set seldissecprov_rs = seldissecprov_qr.OpenResultset(1, 3)
'            seldissecprovflag = True
'        Else
'            seldissecprov_rs.MoreResults
'            seldissecprov_rs.Requery
'        End If
'        If seldissecprov_rs.RowCount > 0 Then
'            cmb_provby.clear
'            While Not seldissecprov_rs.EOF
'                cmb_provby.AddItem Trim(seldissecprov_rs("CCNAME"))
'                seldissecprov_rs.MoveNext
'            Wend
'        End If
'        If seldissecfinflag = False Then
'            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
'            seldissecfinflag = True
'        Else
'            seldissecfin_rs.MoreResults
'            seldissecfin_rs.Requery
'        End If
'        If seldissecfin_rs.RowCount > 0 Then
'            cmb_finalby.clear
'            While Not seldissecfin_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
'                seldissecfin_rs.MoveNext
'            Wend
'        End If
'        If seldissecgrossflag = False Then
'            Set seldissecgross_rs = seldissecgross_qr.OpenResultset(1, 3)
'            seldissecgrossflag = True
'        Else
'            seldissecgross_rs.MoreResults
'            seldissecgross_rs.Requery
'        End If
'        If seldissecgross_rs.RowCount > 0 Then
'            cmb_grossby.clear
'            While Not seldissecgross_rs.EOF
'                cmb_grossby.AddItem Trim(seldissecgross_rs("CCNAME"))
'                seldissecgross_rs.MoveNext
'            Wend
'        End If
If SELPATHDREQUISITNFLAG = False Then
    SELPATHDREQUISITN_QR(0) = REQNO
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
    SELPATHDREQUISITNFLAG = True
Else
    SELPATHDREQUISITN_QR(0) = REQNO
    selpathdrequisitn_rs.MoreResults
    selpathdrequisitn_rs.Requery

End If
If selpathdrequisitn_rs.RowCount > 0 Then
    If ADDENDUMFALG = False Then
        
        txt_fnacreport.Enabled = True
    End If
End If
If FSFLAG = True Then
    'txt_frosecrep.Enabled = False
    txt_frosecrep.Enabled = True
End If
If Trim(selpathdrequisitn_rs("FROZENSEC")) = "NO" Or Trim(selpathdrequisitn_rs("FROZENSEC")) = "" Then
        'txt_frosecrep.Enabled = False
        txt_frosecrep.Enabled = True
End If
If Len(txt_frosecrep.Text) > 1 Then
    If Trim(selpathdrequisitn_rs("FROZENSEC")) = "NO" Or Trim(selpathdrequisitn_rs("FROZENSEC")) = "" Then
        'txt_frosecrep.Enabled = False
        txt_frosecrep.Enabled = True
    End If
If Trim(cmb_immhische.Text) = "YES" Then
    cmb_immhische.Enabled = False
Else
    cmb_immhische.Enabled = True
End If
End If
If ADDENDUMFALG = True Then
    txt_addendum.Enabled = True
    cmb_addendumby.Enabled = True
'    If seldissecaddrflag = False Then
'            Set seldissecadd_rs = seldissecadd_qr.OpenResultset(1, 3)
'            seldissecaddrflag = True
'        Else
'            seldissecadd_rs.MoreResults
'            seldissecadd_rs.Requery
'        End If
'        If seldissecadd_rs.RowCount > 0 Then
'            cmb_addendumby.clear
'            While Not seldissecadd_rs.EOF
'                cmb_addendumby.AddItem Trim(seldissecadd_rs("CCNAME"))
'                seldissecadd_rs.MoveNext
'            Wend
'    End If
    
    'FLDENABL ("ADDENDUM")
    FLDENABL1 ("ADDENDUM")
    'txt_addendum.SetFocus
    flagsave = 0
End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
'On Error GoTo MSG:
'btnenabl ("NEW")
BTNENABL1 ("NEW")

field_clear
TXT_CLEAR
If DATETIMEFLAG = False Then
   datetime
   DATETIMEFLAG = True
Else
    datetime
End If
txt_date.Text = DATE11
txt_time.Text = time1
cmb_enteryby.Text = cmb_enteryby.List(0)
cmb_enteryby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enteryby.List(0)) = Trim(cmb_finalby.List(L1)) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
Next L1
'cmb_finalby.Text = cmb_finalby.List(0)
'cmb_provby.Text = cmb_provby.List(0)
If SEL_DEPTRSFLAG = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SEL_DEPTFLAG = True
Else
    sel_deptrs.MoreResults
    sel_deptrs.Requery
End If
If sel_deptrs.RowCount > 0 Then
    sel_deptrs.MoveFirst
    cmb_dept.clear
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
    Wend
End If

SELMA302FLAG = False
If SELMA302FLAG = False Then
   Set selma302_rs = selma302_qr.OpenResultset(1, 3)
   SELMA302FLAG = True
Else
    selma302_rs.MoreResults
    selma302_rs.Requery
End If

If selma302_rs.RowCount > 0 Then
    selma302_rs.MoveFirst
    CMB_ICDH.clear
    While Not selma302_rs.EOF
        CMB_ICDH.AddItem Trim(selma302_rs("histdesc"))
        selma302_rs.MoveNext
    Wend
End If
CMB_ICDH.AddItem "Not Specified"
flagsave = 1
viewflag = 0
NEWFLAG = True
'FLDENABL ("ENA")
FLDENABL1 ("ENA")
cmb_dept.SetFocus
Opt_patient.Value = True
MaskEdBox1.Enabled = True
Lbl_immhischerep.Visible = False
txt_immhischerep.Visible = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub cmd_print_Click()
'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim i As Integer
 'On Error GoTo prtmsg:

DoEvents

Dim VAR_REQNNO As String
VAR_REQNNO = REQNO

If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = REQNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = REQNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 300
        If seldrequisitn_flag11 = False Then
            seldrequisitn_qr11(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            Set seldrequisitn_rs11 = seldrequisitn_qr11.OpenResultset(1, 3)
            seldrequisitn_flag11 = True
        Else
            seldrequisitn_qr11(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
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
'    If SELDREQUISITNFLAG = False Then
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'            SELDREQUISITNFLAG = True
'        Else
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            seldrequisitn_rs.MoreResults
'            seldrequisitn_rs.Requery
'        End If
'        If seldrequisitn_rs.RowCount > 0 Then
'            prepath = ""
'
'            While Not seldrequisitn_rs.EOF
'                If Trim(seldrequisitn_rs("pathno")) <> Trim(TXT_PATH.Text) Then
'                    prepath = prepath & seldrequisitn_rs("pathno")
'                End If
'                seldrequisitn_rs.MoveNext
'            Wend
'
'        End If
 'STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
If Opt_patient.Value = True Then
    If ADDENDUMFALG = False Then
          '  CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSPFINAL.RPT"
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
            If Trim(prpathno) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Trim(TXT_PATH.Text) & "'"
                'CrystalReport1.Formulas(1) = "PREPATH = ''"
            End If
            CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
            If Len(Trim(cmb_sites.Text)) > 0 Then
                CrystalReport1.Formulas(3) = "SIT = 'Site :" & Trim(cmb_sites.Text) & "'"
            Else
                CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
            End If
             CrystalReport1.Formulas(4) = "OPPRO = '" & txt_opprocedure & "'"
                   
            
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY  REPORT" & "'"
                str1 = ""
                'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
                 str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
    
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY  REPORT" & "'"
                str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            If Len(Trim(txt_frosecrep.Text)) > 3 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(txt_immhischerep.Text)) > 3 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
    '            If Trim(txt_grossdesc.Text) = Chr(13) Then
    '                While InStr(txt_grossdesc, Chr(13)) > 0
    '
    '                    K = InStr(txt_grossdesc, Chr(13))
    '                    beforst = Mid(txt_grossdesc, 1, K - 1) & ""
    '                    afterst = Mid(txt_grossdesc, K + 1)
    '                    txt_grossdesc = beforst & afterst
    '                Wend
    '            End If
            If Len(Trim(txt_grossdesc.Text)) > 3 Then
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "Gross Description" & "'"
            Else
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(txt_microdesc.Text)) > 3 Then
                CrystalReport1.Formulas(10) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(10) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(txt_impfindiag.Text)) > 3 Then
                CrystalReport1.Formulas(11) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(11) = "IMPN = '" & "" & "'"
            End If
      
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            If Len(Trim(CMB_ICDH.Text)) = 0 Or Trim(CMB_ICDH.Text) = "NULL" Then
                CrystalReport1.Formulas(13) = "FFINDIAG = '" & "" & "'"
            Else
                
                CrystalReport1.Formulas(13) = "FFINDIAG = '" & "Histology Type " & CMB_ICDH & "'"
            End If
    '        CrystalReport1.Formulas(14) = "FDIMPT = '" & cmb_t & "'"
    '        CrystalReport1.Formulas(15) = "FDIMPTSIZE = '" & txt_size & "'"
    '        CrystalReport1.Formulas(16) = "FDIMPN = '" & CMB_N & "'"
    '        CrystalReport1.Formulas(17) = "FDIMPM = '" & cmb_m & "'"
         If Trim(cmb_t.Text) = "" Then
               CrystalReport1.Formulas(14) = "FDIMPT = '" & "" & "'"
            Else
                CrystalReport1.Formulas(14) = "FDIMPT = '" & "T " & cmb_t & "'"
            End If
            If Trim(txt_size.Text) = "" Then
                CrystalReport1.Formulas(15) = "FDIMPTSIZE = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "FDIMPTSIZE = '" & "TSize " & txt_size & "'"
            End If
            If Trim(CMB_N) = "" Then
                CrystalReport1.Formulas(16) = "FDIMPN = '" & "" & "'"
            Else
                CrystalReport1.Formulas(16) = "FDIMPN = '" & "N " & CMB_N & "'"
            End If
            If Trim(cmb_m) = "" Then
                CrystalReport1.Formulas(17) = "FDIMPM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(17) = "FDIMPM = '" & "M" & cmb_m & "'"
            End If
    
            i = InStr(cmb_grade.Text, "$")
            If i = 0 Then
                CrystalReport1.Formulas(18) = "FDIMPGRADE = '" & cmb_grade & "'"
            Else
                CrystalReport1.Formulas(18) = "FDIMPGRADE = '" & "Grade" & Mid(cmb_grade, 1, i - 1) & "'"
            End If
            If Len(Trim(txt_fnacreport.Text)) = 0 Then
                CrystalReport1.Formulas(19) = "FNA = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "FNA = '" & "FNAC" & "'"
            End If
    Else
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
            'Commented and added on 20/09/10
            CrystalReport1.SelectionFormula = "{FROZENADLF.REQNNO} = '" & req_chglen(Trim(cmb_reqno.Text)) & "'"
            CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\FROZENADD.RPT"
            
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(Trim(cmb_reqno.Text)) & "'"
            If Trim(prpathno) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Trim(TXT_PATH.Text) & "'"
            End If
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(txt_procedure.Text) & "'"
            If Len(Trim(cmb_sites.Text)) > 0 Then
                CrystalReport1.Formulas(3) = "SIT = ' Site: " & Trim(cmb_sites.Text) & "'"
            Else
                CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
            End If
            CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(txt_opprocedure) & "'"
               
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY  REPORT" & "'"
                str1 = ""
                'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
                 str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(18) = "ENT = '" & str1 & "'"
             Else
    
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY  REPORT" & "'"
                str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(18) = "ENT = '" & str1 & "'"
            End If
            
            If Len(Trim(txt_frosecrep.Text)) > 1 Then
                CrystalReport1.Formulas(5) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(5) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(txt_immhischerep.Text)) > 1 Then
                CrystalReport1.Formulas(6) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(6) = "IHREP = '" & "" & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
            
            If Len(Trim(txt_grossdesc.Text)) > 1 Then
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "Gross Description" & "'"
            Else
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(txt_microdesc.Text)) > 1 Then
                CrystalReport1.Formulas(9) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(9) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(txt_impfindiag.Text)) > 1 Then
                CrystalReport1.Formulas(10) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(10) = "IMPN = '" & "" & "'"
            End If
            
            If Len(Trim(CMB_ICDH.Text)) = 0 Then
                CrystalReport1.Formulas(11) = "FFINDIAG = '" & "" & "'"
            Else
                CrystalReport1.Formulas(11) = "FFINDIAG = '" & "Histology Type " & CMB_ICDH & "'"
            End If
            If Trim(cmb_t.Text) = "" Then
               CrystalReport1.Formulas(12) = "FDIMPT = '" & "" & "'"
            Else
                CrystalReport1.Formulas(12) = "FDIMPT = '" & "T " & cmb_t & "'"
            End If
            If Trim(txt_size.Text) = "" Then
                CrystalReport1.Formulas(13) = "FDIMPTSIZE = '" & "" & "'"
            Else
                CrystalReport1.Formulas(13) = "FDIMPTSIZE = '" & "TSize " & txt_size & "'"
            End If
            If Trim(CMB_N) = "" Then
                CrystalReport1.Formulas(14) = "FDIMPN = '" & "" & "'"
            Else
                CrystalReport1.Formulas(14) = "FDIMPN = '" & "N " & CMB_N & "'"
            End If
            If Trim(cmb_m) = "" Then
                CrystalReport1.Formulas(15) = "FDIMPM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "FDIMPM = '" & "M" & cmb_m & "'"
            End If
            i = InStr(cmb_grade.Text, "$")
            If i = 0 Then
                CrystalReport1.Formulas(16) = "FDIMPGRADE = '" & cmb_grade & "'"
            Else
                CrystalReport1.Formulas(16) = "FDIMPGRADE = '" & Mid(cmb_grade, 1, i - 1) & "'"
            End If
            
            If Len(Trim(txt_fnacreport.Text)) = 0 Then
               CrystalReport1.Formulas(17) = "FNA = '" & "" & "'"
            Else
               CrystalReport1.Formulas(17) = "FNA = '" & "FNAC" & "'"
            End If
            
    End If
End If
If opt_staff.Value = True Then
    If ADDENDUMFALG = False Then
          '  CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSPFINAL.RPT"
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENS.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Trim(TXT_PATH.Text) & "'"
            End If
            CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
            If Len(Trim(cmb_sites.Text)) > 0 Then
                CrystalReport1.Formulas(3) = "SIT = 'Site" & Trim(cmb_sites.Text) & "'"
            Else
                CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
            End If
             CrystalReport1.Formulas(4) = "OPPRO = '" & txt_opprocedure & "'"
                   
            
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY  REPORT" & "'"
                CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY  REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                str1 = ""
                str1 = "Entered By " & Trim(cmb_enteryby.Text)
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY  REPORT" & "'"
                CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY  REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            If Len(Trim(txt_frosecrep.Text)) > 3 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(txt_immhischerep.Text)) > 3 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
    '            If Trim(txt_grossdesc.Text) = Chr(13) Then
    '                While InStr(txt_grossdesc, Chr(13)) > 0
    '
    '                    K = InStr(txt_grossdesc, Chr(13))
    '                    beforst = Mid(txt_grossdesc, 1, K - 1) & ""
    '                    afterst = Mid(txt_grossdesc, K + 1)
    '                    txt_grossdesc = beforst & afterst
    '                Wend
    '            End If
            If Len(Trim(txt_grossdesc.Text)) > 3 Then
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "Gross Description" & "'"
            Else
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(txt_microdesc.Text)) > 3 Then
                CrystalReport1.Formulas(10) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(10) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(txt_impfindiag.Text)) > 3 Then
                CrystalReport1.Formulas(11) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(11) = "IMPN = '" & "" & "'"
            End If
      
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            
            If Len(Trim(CMB_ICDH.Text)) = 0 Or Trim(CMB_ICDH.Text) = "NULL" Then
                CrystalReport1.Formulas(13) = "FFINDIAG = '" & "" & "'"
            Else
                CrystalReport1.Formulas(13) = "FFINDIAG = '" & "Histology Type " & CMB_ICDH & "'"
            End If
    '        CrystalReport1.Formulas(14) = "FDIMPT = '" & cmb_t & "'"
    '        CrystalReport1.Formulas(15) = "FDIMPTSIZE = '" & txt_size & "'"
    '        CrystalReport1.Formulas(16) = "FDIMPN = '" & CMB_N & "'"
    '        CrystalReport1.Formulas(17) = "FDIMPM = '" & cmb_m & "'"
            If Trim(cmb_t.Text) = "" Then
               CrystalReport1.Formulas(14) = "FDIMPT = '" & "" & "'"
            Else
                CrystalReport1.Formulas(14) = "FDIMPT = '" & "T " & cmb_t & "'"
            End If
            If Trim(txt_size.Text) = "" Then
                CrystalReport1.Formulas(15) = "FDIMPTSIZE = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "FDIMPTSIZE = '" & "TSize " & txt_size & "'"
            End If
            If Trim(CMB_N) = "" Then
                CrystalReport1.Formulas(16) = "FDIMPN = '" & "" & "'"
            Else
                CrystalReport1.Formulas(16) = "FDIMPN = '" & "N " & CMB_N & "'"
            End If
            If Trim(cmb_m) = "" Then
                CrystalReport1.Formulas(17) = "FDIMPM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(17) = "FDIMPM = '" & "M" & cmb_m & "'"
            End If
    
            i = InStr(cmb_grade.Text, "$")
            If i = 0 Then
                CrystalReport1.Formulas(18) = "FDIMPGRADE = '" & cmb_grade & "'"
            Else
                CrystalReport1.Formulas(18) = "FDIMPGRADE = '" & "Grade" & Mid(cmb_grade, 1, i - 1) & "'"
            End If
            If Len(Trim(txt_fnacreport.Text)) = 0 Then
                CrystalReport1.Formulas(19) = "FNA = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "FNA = '" & "FNAC" & "'"
            End If
    Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADDS.RPT"
            
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
            If Trim(prpath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Trim(TXT_PATH.Text) & "'"
            End If
            CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
            If Len(Trim(cmb_sites.Text)) > 0 Then
                CrystalReport1.Formulas(3) = "SIT = 'Site " & Trim(cmb_sites.Text) & "'"
            Else
                CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
            End If
            CrystalReport1.Formulas(4) = "OPPRO = '" & txt_opprocedure & "'"
            
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY  REPORT" & "'"
                str1 = ""
                'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
                 str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(18) = "ENT = '" & str1 & "'"
             Else
    
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY  REPORT" & "'"
                str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(18) = "ENT = '" & str1 & "'"
            End If
            
            If Len(Trim(txt_frosecrep.Text)) > 1 Then
                CrystalReport1.Formulas(5) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(5) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(txt_immhischerep.Text)) > 1 Then
                CrystalReport1.Formulas(6) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(6) = "IHREP = '" & "" & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
            
            If Len(Trim(txt_grossdesc.Text)) > 1 Then
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "Gross Description" & "'"
            Else
                CrystalReport1.Formulas(8) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(txt_microdesc.Text)) > 1 Then
                CrystalReport1.Formulas(9) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(9) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(txt_impfindiag.Text)) > 1 Then
                CrystalReport1.Formulas(10) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(10) = "IMPN = '" & "" & "'"
            End If
            
            If Len(Trim(CMB_ICDH.Text)) = 0 Then
                CrystalReport1.Formulas(11) = "FFINDIAG = '" & "" & "'"
            Else
                CrystalReport1.Formulas(11) = "FFINDIAG = '" & "Histology Type " & CMB_ICDH & "'"
            End If
            If Trim(cmb_t.Text) = "" Then
               CrystalReport1.Formulas(12) = "FDIMPT = '" & "" & "'"
            Else
                CrystalReport1.Formulas(12) = "FDIMPT = '" & "T " & cmb_t & "'"
            End If
            If Trim(txt_size.Text) = "" Then
                CrystalReport1.Formulas(13) = "FDIMPTSIZE = '" & "" & "'"
            Else
                CrystalReport1.Formulas(13) = "FDIMPTSIZE = '" & "TSize " & txt_size & "'"
            End If
            If Trim(CMB_N) = "" Then
                CrystalReport1.Formulas(14) = "FDIMPN = '" & "" & "'"
            Else
                CrystalReport1.Formulas(14) = "FDIMPN = '" & "N " & CMB_N & "'"
            End If
            If Trim(cmb_m) = "" Then
                CrystalReport1.Formulas(15) = "FDIMPM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "FDIMPM = '" & "M" & cmb_m & "'"
            End If
            i = InStr(cmb_grade.Text, "$")
            If i = 0 Then
                CrystalReport1.Formulas(16) = "FDIMPGRADE = '" & cmb_grade & "'"
            Else
                CrystalReport1.Formulas(16) = "FDIMPGRADE = '" & Mid(cmb_grade, 1, i - 1) & "'"
            End If
            
            If Len(Trim(txt_fnacreport.Text)) = 0 Then
                CrystalReport1.Formulas(19) = "FNA = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "FNA = '" & "FNAC" & "'"
            End If
            
    End If
End If
CrystalReport1.WindowState = crptMaximized
'CRYSTALREPORT1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox Error(Err)
'txt_msg.Text = ""
'btnenabl ("print")
BTNENABL1 ("print")
Exit Sub
MSG:
MsgBox Error(Err)
End Sub


'Private Sub cmd_print_Click()
'
'
'On Error GoTo MSG:
'
'Dim STR1 As String
'Dim STR2 As String
'Dim STR3 As String
' On Error GoTo prtmsg:
'
'DoEvents
'
'Dim VAR_REQNNO As String
'
'
'VAR_REQNNO = req_chglen(cmb_reqno.Text)
'If SELRREQNOFLAG = False Then
'    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
'    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
'    SELRREQNOFLAG = True
'Else
'    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
'    SELRREQNORS.MoreResults
'    SELRREQNORS.Requery
'End If
' STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"   'AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'If ADDENDUMFALG = False Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSP.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & txt_prepathreqno & "'"
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        CRYSTALREPORT1.Formulas(3) = "SIT = '" & txt_sites & "'"
'        CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & txt_opprocedure & "'"
'
'
'        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
'        CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PROVISIONAL REPORT" & "'"
'        STR1 = ""
'        STR1 = "Entered By " & Trim(cmb_enteryby.Text)
'        CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        Else
'        CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINAL REPORT" & "'"
'        CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        If Len(Trim(txt_frosecrep.Text)) > 1 Then
'            CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "FROZEN SECTION" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'        End If
'        If Len(Trim(txt_immhischerep.Text)) > 1 Then
'            CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
'        End If
'        CRYSTALREPORT1.Formulas(8) = "USRID = '" & STR & "'"
'Else
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSPADD.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & txt_prepathreqno & "'"
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        CRYSTALREPORT1.Formulas(3) = "SIT = '" & txt_sites & "'"
'        CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & txt_opprocedure & "'"
'
'
'        If Len(Trim(txt_frosecrep.Text)) > 1 Then
'            CRYSTALREPORT1.Formulas(5) = "FRSEC = '" & "FROZEN SECTION" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(5) = "FRSEC = '" & "" & "'"
'        End If
'        If Len(Trim(txt_immhischerep.Text)) > 1 Then
'            CRYSTALREPORT1.Formulas(6) = "IHREP = '" & "Immunohistochemistry" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(6) = "IHREP = '" & "" & "'"
'        End If
'        CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'End If
'CRYSTALREPORT1.Action = 1
'Exit Sub
'prtmsg: MsgBox Error(Err)
''txt_msg.Text = ""
'btnenabl ("print")
'Exit Sub
'msg:
'MsgBox Error(Err)
'
'
'
'
'End Sub



Public Sub cmd_save_Click()
 
'On Error GoTo MSG:

'
'If txt_grossdesc = vbNullString Then
'txt_grossdesc = Chr(9)
'End If

Dim m As Integer

If Len(Trim(cmb_sites.Text)) = 0 Then
    m = MsgBox("The Site field is EMPTY.", vbCritical)
    cmb_sites.SetFocus
    Exit Sub
End If
If Len(Trim(CMB_ICDH.Text)) = 0 Then
    m = MsgBox("The Histology field is EMPTY.", vbCritical)
    CMB_ICDH.Enabled = True
    CMB_ICDH.SetFocus
    Exit Sub
End If
If Len(Trim(cmb_reqno.Text)) = 0 Then
    m = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
    cmb_reqno.SetFocus
    Exit Sub
End If
'
If Len(Trim(cmb_enteryby.Text)) = 0 Then
    m = MsgBox("The Enter by field is EMPTY.", vbCritical)
    cmb_enteryby.SetFocus
    Exit Sub
End If
'
If Len(Trim(cmb_finalby.Text)) = 0 Then
    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
    cmb_finalby.SetFocus
    Exit Sub
End If
'
'If Len(Trim(cmb_addendumby.Text)) = 0 Then
'    m = MsgBox("The Addendum by field is EMPTY.", vbCritical)
'    cmb_addendumby.SetFocus
'    Exit Sub
'End If
'
If Len(Trim(cmb_provby.Text)) = 0 Then
    m = MsgBox("The Reported By field is EMPTY.", vbCritical)
    cmb_provby.SetFocus
    Exit Sub
End If
If Len(Trim(txt_natmeterial.Text)) = 0 Then
    m = MsgBox("The Nature of Material field is EMPTY.", vbCritical)
    txt_natmeterial.SetFocus
    Exit Sub
End If

If ADDENDUMFALG = True Then
    
    If Len(Trim(cmb_addendumby)) = 0 Then
        
        m = MsgBox("The Addendum By field is EMPTY.", vbCritical)
        cmb_addendumby.SetFocus
        Exit Sub
    
    End If
    
End If

 If Len(Trim(txt_frosecrep)) > 1 Then
    
    If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
    If seldrequisitn_rs.RowCount > 0 Then
        If seldrequisitn_rs("FROZENSEC") = "YES" Then
        i = MsgBox(" ONCE YOU SAVED FROZEND SECTION YOU CAN'T MODIFY LATER", vbYesNo)
            If i = vbNo Then
                txt_frosecrep.SetFocus
                Exit Sub
            End If
        End If
    End If
  End If
'  If Len(txt_impfindiag.Text) = 0 Then
'    MsgBox "Impression Field is Empty"
'    txt_impfindiag.SetFocus
'    Exit Sub
'  End If
  
  
        While InStr(txt_frosecrep, Chr(9)) > 0
            Dim beforst As String, afterst As String
            Dim K As Long
            K = InStr(txt_frosecrep, Chr(9))
            beforst = Mid(txt_frosecrep, 1, K - 1) & " "
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
            beforst = Mid(txt_grossdesc, 1, K - 1) & " "
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
            beforst = Mid(txt_microdesc, 1, K - 1) & " "
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
            beforst = Mid(txt_immhischerep, 1, K - 1) & " "
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
            beforst = Mid(txt_addendum, 1, K - 1) & " "
            afterst = Mid(txt_addendum, K + 1)
            txt_addendum = beforst & afterst
       Wend
      While InStr(txt_addendum, Chr(146)) > 0

            K = InStr(txt_addendum, Chr(146))
            beforst = Mid(txt_addendum, 1, K - 1) & "`"
            afterst = Mid(txt_addendum, K + 1)
            txt_addendum = beforst & afterst
       
       Wend
       
       While InStr(txt_natmeterial, Chr(9)) > 0
           
            
            K = InStr(txt_natmeterial, Chr(9))
            beforst = Mid(txt_natmeterial, 1, K - 1) & " "
            afterst = Mid(txt_natmeterial, K + 1)
            txt_natmeterial = beforst & afterst
       Wend
        While InStr(txt_natmeterial, Chr(146)) > 0

            K = InStr(txt_natmeterial, Chr(146))
            beforst = Mid(txt_natmeterial, 1, K - 1) & "`"
            afterst = Mid(txt_natmeterial, K + 1)
            txt_natmeterial = beforst & afterst
       Wend
        While InStr(txt_fnacreport, Chr(9)) > 0
           
            K = InStr(txt_fnacreport, Chr(9))
            beforst = Mid(txt_fnacreport, 1, K - 1) & " "
            afterst = Mid(txt_fnacreport, K + 1)
            txt_fnacreport = beforst & afterst
       Wend
        While InStr(txt_fnacreport, Chr(146)) > 0

            K = InStr(txt_fnacreport, Chr(146))
            beforst = Mid(txt_fnacreport, 1, K - 1) & "`"
            afterst = Mid(txt_fnacreport, K + 1)
            txt_fnacreport = beforst & afterst
       Wend
       
'        While InStr(txt_natmeterial, Chr(13)) > 0
'
'            K = InStr(txt_natmeterial, Chr(13))
'            beforst = Mid(txt_natmeterial, 1, K - 1) & " "
'            afterst = Mid(txt_natmeterial, K + 1)
'            txt_natmeterial = beforst & afterst
'       Wend
       If Len(Trim(txt_grossdesc.Text)) >= 8000 Then
            MsgBox "The GrosssDescription Max. length is 8000 Characters Only"
            txt_grossdesc.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_frosecrep.Text)) >= 2000 Then
            MsgBox "The Frozen Report Max. length is 2000 Characters Only"
            txt_frosecrep.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_microdesc.Text)) >= 7000 Then
            MsgBox "The Microscopic Description Max. length is 7000 Characters Only"
            txt_microdesc.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_immhischerep.Text)) >= 7000 Then
            MsgBox "The Immuniohistochemsitry Report Max. length is 7000 Characters Only"
            txt_immhischerep.SetFocus
            Exit Sub
       End If
        If Len(Trim(txt_impfindiag.Text)) >= 2048 Then
            MsgBox "The Impression  Max. length is 2048 Characters Only"
            txt_impfindiag.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_fnacreport.Text)) > 3000 Then
            MsgBox "The Impression  Max. length is 3000 Characters Only"
            txt_fnacreport.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_impfindiag.Text)) = 0 Then
            MsgBox "Please Enter the Impression ", vbInformation
            txt_impfindiag.SetFocus
            Exit Sub
       End If
       
'rdb.BeginTrans
       assign
'rdb.CommitTrans
       
  'FLDENABL ("DIS")
  FLDENABL1 ("DIS")
  'btnenabl ("SAVE")
  BTNENABL1 ("SAVE")
  Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Error(Err)
End Sub

Public Sub cmd_view_Click()

On Error GoTo MSG:
'btnenabl ("VIEW")
BTNENABL1 ("VIEW")
'FLDENABL ("VIEW")
FLDENABL1 ("VIEW")
field_clear
TXT_CLEAR

If DATETIMEFLAG = False Then
   datetime
   DATETIMEFLAG = True
Else
    datetime
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

If sel_deptrs.RowCount > 0 Then
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
    Wend
End If

SELMA302FLAG = False
If SELMA302FLAG = False Then
   Set selma302_rs = selma302_qr.OpenResultset(2, 3)
   SELMA302FLAG = True
Else
    selma302_rs.MoreResults
    selma302_rs.Requery
End If
        
        If selma302_rs.RowCount > 0 Then
            selma302_rs.MoveFirst
            CMB_ICDH.clear
            While Not selma302_rs.EOF
                CMB_ICDH.AddItem Trim(selma302_rs("histdesc"))
                selma302_rs.MoveNext
            Wend
        End If
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
Opt_patient.Value = True
cmb_dept.SetFocus
cmd_print.Enabled = True

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
'On Error GoTo MSG:

NEWFLAG = False
viewflag = False

'If WORKLIST = False Then
'    MAKEFALSE
'End If

MaskEdBox1.Enabled = False
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
For i = 0 To 4
    CMB_N.AddItem i
    cmb_m.AddItem i
    cmb_t.AddItem i
Next i
CMB_N.AddItem "X"
cmb_m.AddItem "X"
CMB_N.AddItem " "
cmb_m.AddItem " "

cmb_immhische.AddItem "YES"
cmb_immhische.AddItem "NO"

SPREPORTFLAG = False

If SPREPORTFLAG = False Then
    SPREPORTFLAG = True

''**********SECURITY**********************
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enteryby.Text = (LOGINUSRNAME)
    cmb_enteryby.AddItem LOGINUSRNAME
    cmb_enteryby.Text = cmb_enteryby.List(0)
    cmb_enteryby.Enabled = False
    cmb_addendumby.clear
    cmb_addendumby.AddItem LOGINUSRNAME
    
Else
    MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
    'cmd_cancel_Click
    Exit Sub
End If

'If (USRNAME_RS("USRTYPE") = 4) Or (USRNAME_RS("USR_ID") = UID) Then
'    cmb_provby.clear
'    cmb_provby.Text = Trim(USRNAME_RS("USERNAME"))
'End If

If (USRNAME_RS("USRTYPE") = 8) Or (USRNAME_RS("USR_ID") = UID) Then
    'cmb_provby.clear
    'cmb_provby.Text = Trim(USRNAME_RS("USERNAME"))
    cmb_finalby.clear
    cmb_finalby.Text = Trim(USRNAME_RS("USERNAME"))
End If
'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)

Dim REPWTRNG_QR As rdoQuery
Dim REPWTRNRSG As rdoResultset

'STR = "SELECT DISTINCT USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A ,EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND (A.USR_ID LIKE 'A1%' OR A.USR_ID LIKE 'A3%') AND B.DEPTCODE='FZZ'  AND B.CESSYN='4' "
str = "select EMPCODE,EMPTITLE from emplib.elempinfo where (empcode like '1%' or empcode like '3%') and deptcode='FZZ' and cessyn='4'"
Set REPWTRNG_QR = rdb.CreateQuery("REPWTRNDDG_STR", str)
Set REPWTRNRSG = REPWTRNG_QR.OpenResultset(1, 1)
If REPWTRNRSG.RowCount > 0 Then
    cmb_grossby.clear
    REPWTRNRSG.MoveFirst
    While Not REPWTRNRSG.EOF
    
        
        'If REPWTRNRS("USRTYPE") = 4 Then
            cmb_grossby.AddItem Trim(REPWTRNRSG(1))
         
         If Mid(REPWTRNRSG(0), 1, 1) = 3 Then
     
            cmb_provby.AddItem Trim(REPWTRNRSG(1))
         End If
        'End If
        REPWTRNRSG.MoveNext
    Wend
End If




Dim REPWTRNF_QR As rdoQuery
Dim REPWTRNRSF As rdoResultset

str = "SELECT DISTINCT USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A ,EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND (A.USR_ID LIKE 'A1%' OR A.USR_ID LIKE 'A3%') AND (B.DEPTCODE='FZZ' OR B.DEPTCODE='ACFZ') AND B.CESSYN='4' "


'STR = "SELECT DISTINCT USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A ,EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND (A.USR_ID LIKE 'A1%' or A.USR_ID LIKE 'A2%') AND (B.DEPTCODE='FZZ' OR B.DEPTCODE='XFZZ') AND B.CESSYN='4' "
Set REPWTRNF_QR = rdb.CreateQuery("REPWTRNDDF_STR", str)
Set REPWTRNRSF = REPWTRNF_QR.OpenResultset(1, 1)
If REPWTRNRSF.RowCount > 0 Then
    'cmb_finalby.clear
    REPWTRNRSF.MoveFirst
    While Not REPWTRNRSF.EOF
        If REPWTRNRSF("USRTYPE") >= 8 Then
            cmb_finalby.AddItem Trim(REPWTRNRSF("USERNAME"))
        End If
        REPWTRNRSF.MoveNext
    Wend
End If

Dim REPWTRNA_QR As rdoQuery
Dim REPWTRNRSA As rdoResultset

str = "SELECT DISTINCT USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A ,EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND A.USR_ID LIKE 'A10%' AND B.DEPTCODE='FZZ' AND B.CESSYN='4' "
Set REPWTRNA_QR = rdb.CreateQuery("REPWTRNDDA_STR", str)
Set REPWTRNRSA = REPWTRNA_QR.OpenResultset(1, 1)
If REPWTRNRSA.RowCount > 0 Then
    'cmb_addendumby.clear
    REPWTRNRSA.MoveFirst
    While Not REPWTRNRSA.EOF
        If REPWTRNRSA("USRTYPE") = 9 Then 'FROM 5 TO 7
            cmb_addendumby.AddItem Trim(REPWTRNRSA("USERNAME"))
        End If
        REPWTRNRSA.MoveNext
    Wend
            
End If

'Dim REPWTRNR_QR As rdoQuery
'Dim REPWTRNRSR As rdoResultset
'
'STR = "SELECT DISTINCT USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A ,EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND (A.USR_ID LIKE 'A1%' OR A.USR_ID LIKE 'A3%') AND (B.DEPTCODE='FZZ' OR B.DEPTCODE='XFZZ') AND B.CESSYN='4' "
'Set REPWTRNR_QR = rdb.CreateQuery("REPWTRNDDR_STR", STR)
'Set REPWTRNRSR = REPWTRNR_QR.OpenResultset(1, 1)
'
'If REPWTRNRSR.RowCount > 0 Then
'
'    'cmb_provby.clear
'    REPWTRNRSR.MoveFirst
'
'    While Not REPWTRNRSR.EOF
'
'        If ((REPWTRNRSR("USRTYPE") <> 6) And (REPWTRNRSR("USRTYPE") >= 4)) Then
'            cmb_provby.AddItem Trim(REPWTRNRSR("USERNAME"))
'        End If
'
'            REPWTRNRSR.MoveNext
'    Wend
'
'End If

str = "SELECT DISTINCT USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A ,EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND (A.USR_ID LIKE 'A1%' OR A.USR_ID LIKE 'A3%') AND (B.DEPTCODE='FZZ' OR B.DEPTCODE='XFZZ') AND B.CESSYN='4'"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNDD_STR", str)
'REPWTRN_QR(0) = Trim(USRDEPT)
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
''cmb_enteryby.Text = usr_name
If REPWTRNRS.RowCount > 0 Then
'    'cmb_enteryby.clear
'    cmb_addendumby.clear
'    cmb_finalby.clear
'    cmb_provby.clear
    cmb_reg.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
'       ' If REPWTRNRS("USRTYPE") >= 1 Then
'         'If REPWTRNRS("USRTYPE") > 1 Then
'         If REPWTRNRS("USRTYPE") >= 8 Then 'FROM 5 TO 7
'            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
'            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
'        End If
'        If ((REPWTRNRS("USRTYPE") <> 6) And (REPWTRNRS("USRTYPE") >= 4)) Then
'           cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
'
'           'cmb_grossby.AddItem Trim(REPWTRNRS("USERNAME"))
'        End If
           'If REPWTRNRS("USRTYPE") = 3 Or REPWTRNRS("USRTYPE") = 4 Then
            cmb_reg.AddItem Trim(REPWTRNRS("USERNAME"))
          'End If
'          If REPWTRNRS("USRTYPE") = 4 Then    'MADE 4  FROM  1
'            cmb_grossby.AddItem Trim(REPWTRNRS("USERNAME"))
'          End If
        REPWTRNRS.MoveNext
    Wend
'   ' cmb_grossby.Text = cmb_grossby.List(0)
Else
    MsgBox "The Security Details of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq1", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=? AND DEPT_CODE='FZZ'"
'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ? AND DEPT_CODE='FZZ'"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR", STR)
'SELFINALBY_QR(0) = 8    'MADE 8 FROM 5
''SELFINALBY_QR(1) = FZZ
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If
'''''''''CRYSTALCODE
str = "SELECT SDP FROM DISLIB.BLOODPRODU WHERE REQNO=?"
Set AMQR = rdb.CreateQuery("AM_STR", str)

''**********SECURITY**********************


str = "SELECT * FROM DISLIB.SITEDETAIL WHERE REQNO=?"
Set selsite_qr = rdb.CreateQuery("SELSITE_STR", str)

str = "SELECT * FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
Set SELRMODALITY_QR = rdb.CreateQuery("", str)
SELRMODALITY_QR(0) = "CT"
Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(2, 3)
     
'''SITE SELECTION
str = "SELECT * FROM DISLIB.DSITE ORDER BY SDESC"
Set SELDSITE_QR = rdb.CreateQuery("SELDSITE_STR", str)

''LATERALITY SELECTION
str = "SELECT * FROM PATOOLSLIB.MA10122"
Set SELMA10122_QR = rdb.CreateQuery("SELMA10122_QR101", str)

''INSERT INTO SITEDETAIL
str = "INSERT INTO DISLIB.SITEDETAIL VALUES(?,?,?,?,?,?,?,?,?,?)"
Set INSSITEDETAIL_QR = rdb.CreateQuery("INSSITEDETAIL_STR", str)

''SELECT SITEDETAIL INDORMATION
str = "SELECT * FROM DISLIB.DSITE WHERE SDESC=?"
Set SELSITEDETAIL_QR = rdb.CreateQuery("SELSITEDETAIL_STR", str)

str = "UPDATE RISLIB.RTIMELOG SET ADRREPDATE=?, ADRREPTIME=? WHERE REQNNO=?"
Set UPDTADNDMTIME_QR = rdb.CreateQuery("UPDTADNDMTIME_QR", str)

''UPDATE SITEDETAIL
str = "UPDATE DISLIB.SITEDETAIL SET SCODE=?,SDESC=?,NCODE=?,LATERALITY=?,RESDOC=? WHERE REQNO=?"
Set UPDSITEDETAIL_QR = rdb.CreateQuery("UPDSITEDETAIL_STR", str)

str = "SELECT * FROM DISLIB.SITEDETAIL WHERE REQNO=?"
Set SELSITEDETAIL_QR1 = rdb.CreateQuery("SELSITEDETAIL_STR1", str)

'for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ') "
Set sel_deptqr = rdb.CreateQuery("sel1_deptqr11", str)
'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)

str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?, COMREPTIME=? WHERE REQNNO=?"
Set UPDTCMT_QR = rdb.CreateQuery("UPDTCMT_QR", str)

    'for icd-h retriveing
str = "select * from patoolslib.ma302 order by histdesc "
Set selma302_qr = rdb.CreateQuery("selma302_qr1", str)
'Set selma302_rs = selma302_qr.OpenResultset(2, 3)
        
str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301QR = rdb.CreateQuery("SELMA3011QR1", str)
        
        
' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqr11", str)

    'for retriving reqnos
'Set req_qr = rdb.CreateQuery("1req", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT desc ")
'Commented and added by Shraddha on 13/09/2011
Set req_qr = rdb.CreateQuery("1req", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND RECPTNO<>'VPP' AND MODALITYCD=? ORDER BY REQNDT desc ")
'req_qr(0) = "aaa"
'req_qr(1) = "CC"
'Set req_rs = req_qr.OpenResultset(1, 3)

' FOR ENTER BY, FINALIZED BY commented by chandra sekhar
''Set SELEMPLIB_QR = rdb.CreateQuery("SELEMPLIB_QR1", "SELECT * FROM EMPLIB.EMPBASINFO WHERE TITLE='DR.'")
'''SELEMPLIB_QR(0) = "AA"
''Set SELEMPLIB_RS = SELEMPLIB_QR.OpenResultset(2, 3)
''SELEMPLIB_RS.MoveFirst
''While Not SELEMPLIB_RS.EOF
''cmb_enteryby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
''cmb_finalby.AddItem SELEMPLIB_RS("EmpTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
''cmb_provby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
''cmb_addendumby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
''cmb_grossby.AddItem SELEMPLIB_RS("emptitle") & "$" & SELEMPLIB_RS("empcode")
''SELEMPLIB_RS.MoveNext
''Wend


' for retriveing name,age etc

Set SELTA101_QR = rdb.CreateQuery("selta101_qr11", "select * from PATOOLSLIB.ta101 where case_no =?")
'selta101_qr(0) = "aa"
'Set selta101_rs = selta101_qr.OpenResultset(2, 3)
'for view purpose
'STR = "select * from emplib.empbasinfo where empcode =?"
'Set selempbasinfo1_qr = rdb.CreateQuery("selempbasinfo1_qr1", STR)
'selempbasinfo1_qr(0) = "00001"
'Set selempbasinfo1_rs = selempbasinfo1_qr.OpenResultset(2, 3)


'' for name of  the reporting doctor
'STR = "select emptitle from emplib.empbasinfo"
'Set repdr_qr = rdb.CreateQuery("repdr_qr10", STR)
'Set repdr_rs = repdr_qr.OpenResultset(2, 3)
'repdr_rs.MoveFirst
'    While Not repdr_rs.EOF
'        cmb_reportdoc.AddItem repdr_rs("emptitle")
'        repdr_rs.MoveNext
'    Wend
    ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR1_QR1", str)
'    REFDR_QR(0) = "BN/00001"
'    REFDR_QR(1) = "FZZSP010000"
'    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    ' FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1_QR1", str)
    'SELPATHDREQUISITN_QR(0) = "SSSSS"
    'SELPATHDREQUISITN_QR(1) = "FZZSP010000"
    'Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(2, 3)

    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma1041_qr11", str)
    'Selma104_qr(0) = "BN/00001"
    'Set selma104_rs = selma104_qr.OpenResultset(2, 3)
    'for provisional diagnosis

    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("se1lta10121_qr1", str)
    'selta1012_qr(0) = "BN/00001"
    'Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)

'    STR = "select * from dislib.dfsreport where reqno=?"
'    Set seldfsreport_qr = rdb.CreateQuery("seldfsreport_qr1", STR)
'    seldfsreport_qr(0) = "zzzzzz"
'    Set seldfsreport_rs = seldfsreport_qr.OpenResultset(2, 3)

'    STR = "select * from dislib.dihreport where reqno=?"
'    Set seldihreport_qr = rdb.CreateQuery("seldihreport_qr", STR)
'    seldihreport_qr(0) = "zzzzzz"
'    Set seldihreport_rs = seldihreport_qr.OpenResultset(2, 3)

    str = "select * from dislib.dimpresion where reqno=?"
    Set seldimpresion_qr = rdb.CreateQuery("seldimpresion_qr11", str)
'    seldimpresion_qr(0) = "zzzzzz"
'    Set seldimpresion_rs = seldimpresion_qr.OpenResultset(2, 3)

    str = "select * from dislib.dgrossdesc where reqno=?"
    Set seldgrossdesc_qr = rdb.CreateQuery("seldgrossdesc_qr1", str)
'    seldgrossdesc_qr(0) = "zzzzzz"
'    Set seldgrossdesc_rs = seldgrossdesc_qr.OpenResultset(2, 3)


   ' for insertion
    str = "insert into  dislib.dfsreport (reqno, fsreport,enterby,preportby,freportby) values (?,?,?,?,?)"
    Set insdfs_qr = rdb.CreateQuery("insdfs_qr1", str)
    str = "insert into dislib.dgrossdesc (reqno,grossdesc,grossedby) values (?,?,?)"
    Set insdgrossdesc_qr = rdb.CreateQuery("insdgrossdesc_qr1", str)
    str = "insert into dislib.dihreport (reqno,ihreport,enteryby,preportby,freportby) values (?,?,?,?,?)"
    Set INSDIHREPORT_QR = rdb.CreateQuery("insdihreport_qr1", str)
    str = "insert into dislib.dimpresion (reqno,finaldiag,fdescry,t,tsize,n,m,gradecode) values (?,?,?,?,?,?,?,?)"
    Set insdimpresion_qr = rdb.CreateQuery("1insdimpresion_qr", str)
    str = "insert into dislib.dspreport (reqno,spreport,enterby,preportby,freportby) values(?,?,?,?,?)"
    Set insdspreport_qr = rdb.CreateQuery("insdspreport_qr1", str)

    str = "update dislib.dfsreport set fsreport=? where reqno=?"
    Set upddfsreport_qr = rdb.CreateQuery("upddfsreport_qr1", str)
    
    str = "update dislib.dihreport set ihreport=?, enteryby=?, preportby=?, freportby=? where reqno=?"
    Set upddihreport_qr = rdb.CreateQuery("upddihreport1_qr11", str)
    
    str = "update dislib.dimpresion set  finaldiag=?, fdescry=?, t=?,tsize=?, n=?, m=?,GRADECODE=? where reqno=?"
    Set upddimpresion_qr = rdb.CreateQuery("upddimpresion_qr1", str)
    
    str = "UPDATE DISLIB.DSPREPORT SET SPREPORT=?, ENTERBY=?, PREPORTBY=?, FREPORTBY=? WHERE REQNO=?"
   '  str = "UPDATE DISLIB.DSPREPORT SET SPREPORT=? WHERE REQNO=?"
    Set UPDSPREPORT_QR = rdb.CreateQuery("UPDSPREPORT_QR1", str)
    
    str = "UPDATE DISLIB.DGROSSDESC SET GROSSDESC=?,GROSSEDBY=? WHERE REQNO=?"
    Set updgrossdesc_qr = rdb.CreateQuery("updgrossdesc_qr1", str)
    
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR11", str)
'    SELRREQUISITN_QR(0) = "BN/00000"
'    SELRREQUISITN_QR(1) = "IH"
'    SELRREQUISITN_QR(2) = "FZZ"
'    SELRREQUISITN_QR(3) = "T"
'    SELRREQUISITN_QR(4) = "F"
'    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
    
    
    
   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1", str)
   
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qr1", str)
    
   str = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1", str)
   
   str = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11", str)
   
   SELDADDENDUM_QR(0) = "FSS"
   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
   
   str = "select * from dislib.drequisitn where caseno=? and reqno like 'FZZSP%'"
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_1qr1", str)
   
   '''for sp prev. path report
   
   'STR = "SELECT REQNNO,REQNDT,REQNSTAT,RISLIB.RREQUISITN.CASENO,PREPORTDT,PATHNO FROM RISLIB.RREQUISITN LEFT OUTER JOIN DISLIB.DREQUISITN ON RISLIB.RREQUISITN.REQNNO=DISLIB.DREQUISITN.REQNO WHERE RISLIB.RREQUISITN.CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')" 'ORDER BY RREQUISITN.REQNDT DESC"
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='F' or REQNSTAT='T') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%') ORDER BY PREPORTDT DESC"   'Order by  DSPLF.REQNDT,DSPLF.REQNNO DESC
    Set seldrequisitn_qr10 = rdb.CreateQuery("seldrequisitn_1qr110", str)
   
    seldrequisitn_qr10(0) = "BN/00000"
    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
   
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr11 = rdb.CreateQuery("seldrequisitn11_1qrspREP", str)
    
    'seldrequisitn_qr11(0) = "BN/00000"
    'Set seldrequisitn_rs11 = seldrequisitn_qr11.OpenResultset(1, 3)
    
        
'   seldrequisitn_qr(0) = "BN/00000"
'   Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
    
   str = "INSERT INTO DISLIB.DFREPORT(REQNO,FSPREPORT,IHREPORT,IMPRESSION,GROSSDESC,GROSSEDBY,SPREPORT,REPORTIND,FNAC) VALUES (?,?,?,?,?,?,?,?,?)"
   Set INSFREPORT_QR = rdb.CreateQuery("INSFREPORT_QR1", str)

   str = "UPDATE DISLIB.DFREPORT SET FSPREPORT=?,IHREPORT=?,IMPRESSION=?,GROSSDESC=?,SPREPORT=?,REPORTIND=?,FNAC=? WHERE REQNO=?"
   Set upddfsreport_qr = rdb.CreateQuery("upddfsreport_qr11", str)
    
   str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
   Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQR1", str)
'   SELRREQNOQR(0) = "FZZSP0200000"
'   Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    
    str = "SELECT REQNO,FSPREPORT,IHREPORT,IMPRESSION,GROSSDESC,GROSSEDBY,SPREPORT,REPORTIND,FNAC FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QR1", str)
    
   ' SELDFREPORT_QR(0) = reqno
   ' Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
    
    str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_IHQR = rdb.CreateQuery("SELDFREPORT_IHQR1", str)
'    SELDFREPORT_IHQR(0) = "FZZSP0200001"
'    Set SELDFREPORT_IHRS = SELDFREPORT_IHQR.OpenResultset(1, 3)
    
    str = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?,NMETERIALS=? WHERE REQNO=?"
    'STR = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?,NMETERIALS=?,IMMHISCHE=? WHERE REQNO=?"
    Set upddrequisitn_qr = rdb.CreateQuery("UPDSPREPORT1_QR1", str)
    
    str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO11_QR091", str)
'    SELEMPBASINFO_QR(0) = "00000"
    'Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    
    str = "select * from dislib.dgrade"
    Set seldgradeqr = rdb.CreateQuery("seldgradeqr1", str)
        
    str = "select * from dislib.dgrade where gradecode=?"
    Set seldgrade_qr = rdb.CreateQuery("seldgrade_qr1", str)
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
    Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQR1", str)
    
   str = "select * from dislib.dihreport where reqno=?"
   Set selihreportqr = rdb.CreateQuery("selihreportqr1", str)
   
   str = "select * from dislib.dissec where enterby='YES'AND MODALITY='SP'"
   Set seldissecentr_qr = rdb.CreateQuery("seldissecentr_qr1", str)
   
   str = "select * from dislib.dissec where PROVby='YES' AND MODALITY='SP'"
   Set seldissecprov_qr = rdb.CreateQuery("seldissecPROV_qr1", str)
   
   str = "select * from dislib.dissec where FINALby='YES' AND MODALITY='SP'"
   Set seldissecfin_qr = rdb.CreateQuery("seldissecFIN_qr1", str)
   
   str = "select * from dislib.dissec where ADDENDUMby='YES'AND MODALITY='SP'"
   Set seldissecadd_qr = rdb.CreateQuery("seldissecADD_qr1", str)
   
   str = "select * from dislib.dissec where GROSSby='YES'AND MODALITY='SP'"
   Set seldissecgross_qr = rdb.CreateQuery("seldissecGROSS_qr1", str)
   
   str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
   Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1QR11", str)
   
   str = "select sitecd,sitedesc from dislib.sitedup where sitecd = ?"
   Set viewsitedupqr = rdb.CreateQuery("viewsitedupqr1", str)
   
   str = "INSERT INTO DISLIB.DSPFN (SPREQNO,FNREQNO) VALUES (?,?)"
   Set INSDSPFNREQQR = rdb.CreateQuery("INSDSPFNREQQR1", str)
  
  str = "INSERT INTO DISLIB.DFNREPORT (REQNO,MICREPORT) VALUES(?,?)"
  Set INSFNREPORTQR = rdb.CreateQuery("INSFNREPORTQR1", str)
  
  str = "UPDATE DISLIB.DFNREPORT SET MICREPORT=? WHERE REQNO=?"
  Set UPDFNREPORTQR = rdb.CreateQuery("UPDFNREPORTQR1", str)
  
  str = "SELECT * FROM DISLIB.DSPFN WHERE SPREQNO=?"
  Set SELDSPFN_QR = rdb.CreateQuery("SLEDSPFN_QR1", str)
  
  str = "SELECT * FROM DISLIB.DFNREPORT WHERE REQNO=?"
  Set SELDFNREPORTQR = rdb.CreateQuery("SELDFNREPORTQR1", str)
  
  str = "INSERT INTO RISLIB.DISIMP(REQNO,IMP)VALUES(?,?)"
  Set INSDISIMP_QR = rdb.CreateQuery("INSDISIMP_QRSP", str)
  
  str = "UPDATE RISLIB.DISIMP SET IMP=? WHERE REQNO=?"
  Set UPDDISIMP_QR = rdb.CreateQuery("UPDDISIMP_QRSP", str)
    
    
        
Set seldgraders = seldgradeqr.OpenResultset(1, 3)
If seldgraders.RowCount > 0 Then
    cmb_grade.clear
    While Not seldgraders.EOF
        cmb_grade.AddItem Trim(seldgraders("gradedesc")) & "$" & Trim(seldgraders("gradecode"))
        seldgraders.MoveNext
    Wend
End If
        
    lbl_Grade.Visible = True
    cmb_grade.Visible = True
    txt_addendum.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_dept.Enabled = False
   'FLDENABL ("DIS")
   txt_fullreport.Visible = False

End If
BTNENABL1 ("INI")
txt_immhischerep.Visible = False
Lbl_immhischerep.Visible = False
'' HL7QUERIS
str = "SELECT * FROM DISLIB.PACSLOG WHERE REQNO = ?"
Set SELPACSLOG_RQ = rdb.CreateQuery("SELPACSLOG1", str)
str = "UPDATE DISLIB.PACSLOG SET REPSEND = ?, R_DATE = ?, R_TIME = ? WHERE REQNO = ?"
Set UPDPACSLOGPREP_RQ = rdb.CreateQuery("UPDPACSPREP", str)
str = "UPDATE DISLIB.PACSLOG SET FREPSEND = ?, FR_DATE = ?, FR_TIME = ? WHERE REQNO = ?"
Set UPDPACSLOGFREP_RQ = rdb.CreateQuery("UPDPACSLOGFREP", str)
str = "UPDATE DISLIB.PACSLOG SET ADDENDUM = ?, ADD_DATE = ?, ADD_TIME = ? WHERE REQNO = ?"
Set UPDPACSLOGADD_RQ = rdb.CreateQuery("UPDPACSLOFADD", str)


'' HL7QUERIS
Exit Sub
MSG:
MsgBox Err.Description
    
End Sub
Private Sub Form_Unload(cancel As Integer)

Dim RS As rdoResultset
Dim RQ As rdoQuery

If WORKLIST = True Then

    For Each RS In rdb.rdoResultsets
        RS.Close
    Next

    For Each RQ In rdb.rdoQueries
        RQ.Close
    Next

End If

End Sub

Public Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:

If Len(Trim(MaskEdBox1.Text)) = "" Then
    Exit Sub
End If

If Opt_patient.Value = True Then
    CASENO = MaskEdBox1.Text
    MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
    If NEWFLAG = True Then
        If REQFLAG = False Then
            req_qr(0) = MaskEdBox1.FormattedText
            req_qr(1) = Trim(txt_modcode.Text)
            Set req_rs = req_qr.OpenResultset(1, 3)
            REQFLAG = True
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
    End If
    If viewflag = True Then
        If LOGINTYP = "0" Then
            If SELRREQUISITNFLAG = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "F"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "F"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
        Else
            If SELRREQUISITNFLAG = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
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
    If NEWFLAG = True Then
    
        
          If REQFLAG = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
                REQFLAG = True
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
        End If
        If viewflag = True Then
            If LOGINTYP = "0" Then
                If SELRREQUISITNFLAG = False Then
                
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "F"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "F"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
        Else
           If SELRREQUISITNFLAG = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
        End If
            If SELRREQUISITN_RS.RowCount > 0 Then
                cmb_reqno.clear
                While Not SELRREQUISITN_RS.EOF
                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                    SELRREQUISITN_RS.MoveNext
                Wend
             Else
                MsgBox " No Requisition for this Modality"
                btnenabl ("INI")
                Exit Sub
            End If
        
        End If

End If

Exit Sub

MSG:
MsgBox Error(Err)
End Sub

Public Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
'On Error GoTo MSG:
REQNO = req_chglen(cmb_reqno.Text)
SELDFREPORT_QR(0) = REQNO
Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
    
If Opt_patient.Value = True Then
    If NEWFLAG = True Then
        NEWFLAG = False
        If SELTA101FALG = False Then
        
           
            SELTA101_QR(0) = MaskEdBox1.FormattedText
             Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            SELTA101FALG = True
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
        
        If REFDRFLAG = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = REQNO
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = REQNO
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
        If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
'            If Trim(REFDR_RS("ser_code")) = "F104" Or Trim(REFDR_RS("SER_CODE")) = "F105" Then
'                txt_grossdesc.Enabled = False
'                cmb_grossby.Enabled = False
'            Else
                txt_grossdesc.Enabled = True
                cmb_grossby.Enabled = True
'            End If
            If REFDR_RS("RECORDNO") = 1 Then
                LBL_VPP.Caption = "VPP"
                If (Trim(REFDR_RS("RECPTNO")) = "VPP" And Trim(REFDR_RS("REQNSTAT")) <> "E") Then
                    MsgBox "Payment is not done,You are not Authorised to enter the Report"
                    btnenabl ("INI")
                    field_clear
                    Exit Sub
                End If
            End If
            If REFDR_RS("RECORDNO") = 2 Then
                LBL_VPP.Caption = "LR"
            End If
            If REFDR_RS("RECORDNO") = 3 Then
                LBL_VPP.Caption = "BAR"
            End If
            If REFDR_RS("RECORDNO") = 4 Then
                LBL_VPP.Caption = "DrOp"
            End If
        End If
        
'        If REFDR_RS.RowCount > 0 Then
'            txt_refdoc.Text = REFDR_RS("REFBYDOC")
'        End If
        If selma104flag = False Then
            selma104_qr(0) = MaskEdBox1.FormattedText
            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
            selma104flag = True
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
      
     
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = REQNO
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = REQNO
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
'            If IsNull(Trim(selpathdrequisitn_rs("SITE"))) Then
'                txt_sites.Text = ""
'            Else
''                If SELMA301FLAG = False Then
''                    SELMA301QR(0) = selpathdrequisitn_rs("SITE")
''                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''                    SELMA301FLAG = True
''                Else
''                    SELMA301QR(0) = selpathdrequisitn_rs("SITE")
''                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''                    SELMA301RS.MoreResults
''                    SELMA301RS.Requery
''
''                End If
''                If SELMA301RS.RowCount > 0 Then
'''                    If Trim(SELMA301RS("SITEDESC")) = "site not specified/recorded" Then
'''                        txt_sites.Text = ""
'''                    Else
'''                        txt_sites.Text = SELMA301RS("SITEDESC")
'''                    End If
''                    Else
'''                            If VIEWSITEDUPRSFLAG = False Then
'''                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
'''                                 Set viewsiteduprs = viewsitedupqr.OpenResultset(1, 3)
'''                                 VIEWSITEDUPRSFLAG = True
'''                            Else
'''                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
'''                                 viewsiteduprs.MoreResults
'''                                 viewsiteduprs.Requery
'''
'''                            End If
'''                            If viewsiteduprs.RowCount > 0 Then
'''                                txt_sites.Text = viewsiteduprs("SITEDESC")
'''                            End If
''                End If
'            End If
'            If IsNull(Trim(selpathdrequisitn_rs("LATERALITS"))) Then
'                txt_laterality.Text = ""
'            Else
'                txt_laterality.Text = selpathdrequisitn_rs("LATERALITS")
'            End If
        
            If Trim(selpathdrequisitn_rs("FROZENSEC")) = "NO" Or Trim(selpathdrequisitn_rs("FROZENSEC")) = "" Then
               'txt_frosecrep.Enabled = False
               txt_frosecrep.Enabled = True
            End If
            If selpathdrequisitn_rs.EOF = False Then
                txt_natmeterial.Text = " " & Trim(selpathdrequisitn_rs("NMETERIALS"))
            Else
                txt_natmeterial.Text = " "
            End If
            txt_opprocedure.Text = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
            txt_refdoc.Text = Trim(selpathdrequisitn_rs("refdoc"))
       
            If IsNull(selpathdrequisitn_rs("IMMHISCHE")) Then
                txt_immhischerep.Enabled = False
                txt_immhischereqno.Enabled = False
                
            End If
            If Trim(selpathdrequisitn_rs("labno")) = "YES" Then
                txt_fnacreport.Enabled = True
            Else
                txt_fnacreport.Enabled = False
            End If
        End If
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = MaskEdBox1.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = MaskEdBox1.FormattedText
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
        If selta1012_rs.RowCount > 0 Then
            txt_procedure.Text = selta1012_rs("prvdiagdes")
        End If
'        If SELDREQUISITNFLAG = False Then
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'            SELDREQUISITNFLAG = True
'        Else
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            seldrequisitn_rs.MoreResults
'            seldrequisitn_rs.Requery
'        End If
        'seldrequisitn_qr10.QueryTimeout = 150
        'If seldrequisitn_flag10 = False Then
        seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
        seldrequisitn_rs10.MoreResults
        seldrequisitn_rs10.Requery
        If seldrequisitn_rs10.RowCount > 0 Then
            prpathno = ""
            CMB_PATHNO.clear
            While Not seldrequisitn_rs10.EOF
                MODALITY1 = Mid(seldrequisitn_rs10("REQNNO"), 4, 2)
                If Trim(MODALITY1) = "SP" Then
                    MODDESC = "SURGICAL PATHOLOGY"
                ElseIf Trim(MODALITY1) = "CN" Then
                    MODDESC = "CYTOLOGY NON GYNAEC"
                ElseIf Trim(MODALITY1) = "CG" Then
                    MODDESC = "CYTOLOGY GYNAEC"
                ElseIf Trim(MODALITY1) = "FA" Then
                    MODDESC = "FINE NEEDLE ASPIRATION CYTOLOGY"
                ElseIf Trim(MODALITY1) = "CT" Then
                    MODDESC = "CT SCAN"
                ElseIf Trim(MODALITY1) = "CR" Then
                    MODDESC = "CONVENTIONAL RADIOGRAPHY"
                ElseIf Trim(MODALITY1) = "MM" Then
                    MODDESC = "MAMMOGRAPHY"
                ElseIf Trim(MODALITY1) = "MR" Then
                    MODDESC = "MRI SCAN"
                ElseIf Trim(MODALITY1) = "US" Then
                    MODDESC = "ULTRASONOGRAPHY"
                End If
                If Trim(TXT_PATH.Text) <> seldrequisitn_rs10("pathno") Or IsNull(seldrequisitn_rs10("pathno")) Then
                prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
                'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " $ " & seldrequisitn_rs10("PREPORTDT") & " " & "$" & " " & req_chglen(seldrequisitn_rs10("REQNNO")) & " " & "$" & " " & MODDESC
                CMB_PATHNO.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
                'CMB_PATHNO.ListIndex = 0
                
             End If
            seldrequisitn_rs10.MoveNext
         Wend
         CMB_PATHNO.Text = CMB_PATHNO.List(0)
         txt_prepathreqno.Text = prpathno
        ' cmb_grossby.Text = cmb_grossby.List(0)
    End If
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
        
'        If seldissecentrflag = False Then
'            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
'            seldissecentrflag = True
'        Else
'            seldissecentr_rs.MoreResults
'            seldissecentr_rs.Requery
'        End If
'        If seldissecentr_rs.RowCount > 0 Then
'            cmb_enteryby.clear
'            While Not seldissecentr_rs.EOF
'                cmb_enteryby.AddItem Trim(seldissecentr_rs("CCNAME"))
'                seldissecentr_rs.MoveNext
'            Wend
'        End If
'        If seldissecprovflag = False Then
'            Set seldissecprov_rs = seldissecprov_qr.OpenResultset(1, 3)
'            seldissecprovflag = True
'        Else
'            seldissecprov_rs.MoreResults
'            seldissecprov_rs.Requery
'        End If
'        If seldissecprov_rs.RowCount > 0 Then
'            cmb_provby.clear
'            While Not seldissecprov_rs.EOF
'                cmb_provby.AddItem Trim(seldissecprov_rs("CCNAME"))
'                seldissecprov_rs.MoveNext
'            Wend
'        End If
'        If seldissecfinflag = False Then
'            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
'            seldissecfinflag = True
'        Else
'            seldissecfin_rs.MoreResults
'            seldissecfin_rs.Requery
'        End If
'        If seldissecfin_rs.RowCount > 0 Then
'            cmb_finalby.clear
'            cmb_reg.clear
'            While Not seldissecfin_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
'                cmb_reg.AddItem Trim(seldissecfin_rs("ccname"))
'                seldissecfin_rs.MoveNext
'            Wend
'        End If
'        If seldissecgrossflag = False Then
'            Set seldissecgross_rs = seldissecgross_qr.OpenResultset(1, 3)
'            seldissecgrossflag = True
'        Else
'            seldissecgross_rs.MoreResults
'            seldissecgross_rs.Requery
'        End If
'        If seldissecgross_rs.RowCount > 0 Then
'            cmb_grossby.clear
'            While Not seldissecgross_rs.EOF
'                cmb_grossby.AddItem Trim(seldissecgross_rs("CCNAME"))
'                seldissecgross_rs.MoveNext
'            Wend
'        End If
    SELPACSLOG_RQ(0) = req_chglen(cmb_reqno)
    Set SELPACSLOG_RS = SELPACSLOG_RQ.OpenResultset(1, 1)
        If SELPACSLOG_RS.RowCount > 0 Then
         MsgBox "This Requisition Number for this Case is in Pacslog, Ensure Weather Picture is Captured & Then Proceed to Write Report"
         
        End If
        
    End If
    If viewflag = True Then
    
       
        
        viewflag = False
        
        If SELSITE_FLAG = False Then
            selsite_qr(0) = REQNO ' req_chglen(Trim(cmb_reqno.Text))
            Set selsite_rs = selsite_qr.OpenResultset(1, 3)
            SELSITE_FLAG = True
        Else
            selsite_qr(0) = REQNO 'req_chglen(Trim(cmb_reqno.Text))
            selsite_rs.MoreResults
            selsite_rs.Requery
        End If
        If selsite_rs.EOF Then
            cmb_sites.Text = "NULL"
        End If
        If SELTA101FALG = False Then
            SELTA101_QR(0) = MaskEdBox1.FormattedText
            Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            SELTA101FALG = True
        Else
            SELTA101_QR(0) = MaskEdBox1.FormattedText
            Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
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
    
        If REFDRFLAG = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = REQNO
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = REQNO
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
    
        If REFDR_RS.RowCount > 0 Then
        MODADD = ""
            'txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If REFDR_RS("COMMITFLAG") = "Y" Then
                MsgBox "Final Report is Generated, You can't Modify the Final Report,You can Add ADDENDUMS Only"
                cmd_modify.Enabled = True
               MODADD = "ADD"
                'txt_addendum.Enabled = True
                'cmd_save.Enabled = True
                'ADDENDUMFALG = True
            End If
            If REFDR_RS("RECORDNO") = 1 Then
                LBL_VPP.Caption = "VPP"
            End If
            If REFDR_RS("RECORDNO") = 2 Then
                LBL_VPP.Caption = "LR"
            End If
            If REFDR_RS("RECORDNO") = 3 Then
                LBL_VPP.Caption = "BAR"
            End If
            If REFDR_RS("RECORDNO") = 4 Then
                LBL_VPP.Caption = "DrOp"
            End If
        End If
        If SELSITEDETAIL_FLAG1 = False Then
            SELSITEDETAIL_QR1(0) = REQNO 'req_chglen(cmb_reqno.Text)
            Set SELSITEDETAIL_RS1 = SELSITEDETAIL_QR1.OpenResultset(1, 3)
            SELSITEDETAIL_FLAG1 = True
        Else
            SELSITEDETAIL_QR1(0) = REQNO 'req_chglen(cmb_reqno.Text)
            SELSITEDETAIL_RS1.MoreResults
            SELSITEDETAIL_RS1.Requery
        End If
        If Not SELSITEDETAIL_RS1.EOF Then
            cmb_sites.Text = "" & Trim(SELSITEDETAIL_RS1("SDESC"))
            cmb_laterality.Text = "" & Trim(SELSITEDETAIL_RS1("LATERALITY"))
            If SELSITEDETAIL_RS1("RESDOC") <> "NO" Then
                TXT_FSDIS.Visible = True
                TXT_FSDIS.Text = "" & Trim(SELSITEDETAIL_RS1("RESDOC"))
            End If
        End If
        
        If selma104flag = False Then
            selma104_qr(0) = MaskEdBox1.FormattedText
            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
            selma104flag = True
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
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
            txt_status.Text = "OP"
        End If
      
     
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = REQNO
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = REQNO
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
            txt_refdoc.Text = selpathdrequisitn_rs("REFDOC")
            'If IsNull(selpathdrequisitn_rs("SITE")) Then
            '    txt_sites.Text = ""
            'Else
'                If SELMA301FLAG = False Then
'                    SELMA301QR(0) = selpathdrequisitn_rs("SITE")
'                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
'                    SELMA301FLAG = True
'                Else
'                    SELMA301QR(0) = selpathdrequisitn_rs("SITE")
'                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
'                    SELMA301RS.MoreResults
'                    SELMA301RS.Requery
'
'                End If
                'If SELMA301RS.RowCount > 0 Then
'                    If Trim(SELMA301RS("sitedesc")) = "site not specified/recorded" Then
'                        txt_sites.Text = ""
'                    Else
'                        txt_sites.Text = SELMA301RS("SITEDESC")
'
'                    End If
                 '   Else
'                            If VIEWSITEDUPRSFLAG = False Then
'                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
'                                 Set viewsiteduprs = viewsitedupqr.OpenResultset(1, 3)
'                                 VIEWSITEDUPRSFLAG = True
'                            Else
'                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
'                                 viewsiteduprs.MoreResults
'                                 viewsiteduprs.Requery
'
'                            End If
'                            If viewsiteduprs.RowCount > 0 Then
'                                txt_sites.Text = viewsiteduprs("SITEDESC")
'                            End If
                'End If
            'End If
            'If IsNull(selpathdrequisitn_rs("LATERALITS")) Then
            'Else
             '   txt_laterality.Text = selpathdrequisitn_rs("LATERALITS")
            'End If
            If selpathdrequisitn_rs.EOF = False Then
                txt_natmeterial.Text = " " & (selpathdrequisitn_rs("NMETERIALS"))
            Else
                txt_natmeterial.Text = ""
            End If
            
            If IsNull(Trim(selpathdrequisitn_rs("OPPROCEDUR"))) Then
                txt_opprocedure.Text = ""
            Else
                txt_opprocedure.Text = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
            End If
            
            If Trim(selpathdrequisitn_rs("FROZENSEC")) = "NO" Then
                FSFLAG = True
            End If
            
            SELDREQNOQR(0) = REQNO
            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
                If SELDREQNORS.RowCount > 0 Then
                    txt_immhischereqno.Text = req_chglen((SELDREQNORS("REQNO")))
                    cmb_immhische.Text = "YES"
                    txt_immhischerep.Visible = True
                    Lbl_immhischerep.Visible = True
            Else
                    txt_immhischerep.Text = ""
                    cmb_immhische.Text = "NO"
                
            End If
        End If
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = MaskEdBox1.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = MaskEdBox1.FormattedText
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
        
        If selta1012_rs.RowCount > 0 Then
            txt_procedure.Text = selta1012_rs("prvdiagdes")
        End If
        
'''        If seldfreportflag = False Then
'''            SELDFREPORT_QR(0) = reqno
'''            Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
'''            seldfreportflag = True
'''        Else
'''            SELDFREPORT_QR(0) = reqno
'''            SELDFREPORT_RS.MoreResults
'''            SELDFREPORT_RS.Requery
'''        End If
        
''        If SELDFREPORT_RS.RowCount > 0 Then
''            If SELDFREPORT_RS.EOF = False Then
''                txt_frosecrep.Text = " " & SELDFREPORT_RS("FSPREPORT").Value
''            Else
''                txt_frosecrep.Text = ""
''            End If
''            If SELDFREPORT_RS.EOF = False Then
''                txt_immhischerep.Text = " " & SELDFREPORT_RS("IHREPORT").Value
''            Else
''                txt_immhischerep.Text = ""
''            End If
''            If SELDFREPORT_RS.EOF = False Then
''                txt_impfindiag.Text = " " & SELDFREPORT_RS("IMPRESSION").Value
''            Else
''                txt_impfindiag.Text = ""
''            End If
''            If SELDFREPORT_RS.EOF = False Then
''                txt_grossdesc.Text = " " & SELDFREPORT_RS("GROSSDESC").Value
''            Else
''                txt_grossdesc.Text = ""
''            End If
''            If SELDFREPORT_RS.EOF = False Then
''                txt_microdesc.Text = " " & SELDFREPORT_RS("SPREPORT").Value
''            Else
''                txt_SPREPORT.Text = ""
''            End If
''            If SELDFREPORT_RS.EOF = False Then
''                txt_fnacreport.Text = " " & SELDFREPORT_RS("FNAC").Value
''                CMB_FNAC.Text = "YES"
''            Else
''                txt_fnacreport.Text = ""
''                CMB_FNAC.Text = "NO"
''            End If
''        End If
        
        If SELDFREPORT_RS.RowCount > 0 Then
            
            If SELDFREPORT_RS.EOF = False Then
                If IsNull(SELDFREPORT_RS("FSPREPORT")) Then
                    txt_frosecrep.Text = ""
                Else
                    txt_frosecrep.Text = SELDFREPORT_RS("FSPREPORT").Value
                End If
            End If
        Else
            txt_frosecrep.Text = ""
        End If

        If SELDFREPORT_RS.RowCount > 0 Then
               If SELDFREPORT_RS.EOF = False Then
                    If IsNull(SELDFREPORT_RS("ihreport")) Then
                        txt_immhischerep.Text = ""
                        cmb_immhische.Text = "NO"
                    Else
                       txt_immhischerep.Text = SELDFREPORT_RS("ihreport").Value
                       cmb_immhische.Text = "YES"
                    End If
                End If
        Else
               txt_immhischerep.Text = ""
               cmb_immhische.Text = "NO"
        End If

        If SELDFREPORT_RS.RowCount > 0 Then
               If SELDFREPORT_RS.EOF = False Then
                       If IsNull(SELDFREPORT_RS("impression")) Then
                           txt_impfindiag = ""
                       Else
                           txt_impfindiag.Text = SELDFREPORT_RS("impression").Value
                       End If
               Else
                 txt_impfindiag.Text = ""
                End If
        End If
              
        If SELDFREPORT_RS.RowCount > 0 Then
            cmb_grossby.Text = Trim(SELDFREPORT_RS("GROSSEDBY").Value)
            If SELDFREPORT_RS.EOF = False Then
               If IsNull(SELDFREPORT_RS("GROSSDESC")) Then
                      txt_grossdesc = ""
               Else
                      txt_grossdesc = SELDFREPORT_RS("GROSSDESC").Value
               End If
            End If
        Else
           txt_grossdesc = ""
        End If
        
       If SELDFREPORT_RS.RowCount > 0 Then
           If SELDFREPORT_RS.EOF = False Then
                If IsNull(SELDFREPORT_RS("SPREPORT")) Then
                    txt_microdesc = ""
                Else
                    txt_microdesc = Trim(SELDFREPORT_RS("SPREPORT").Value)
                End If
           End If
        Else
           txt_microdesc = ""
        End If
             
        If SELDFREPORT_RS.RowCount > 0 Then
              If SELDFREPORT_RS.EOF = False Then
                     If IsNull(SELDFREPORT_RS("GROSSEDBY")) Then
                        cmb_grossby.Text = ""
                     Else
                        cmb_grossby.Text = SELDFREPORT_RS("GROSSEDBY")
                     End If
              End If
             Else
                cmb_grossby.Text = ""
        End If

        If SELDFREPORT_RS.RowCount > 0 Then
             If SELDFREPORT_RS.EOF = False Then
                If IsNull(SELDFREPORT_RS("FNAC").Value) Then
                        txt_fnacreport.Text = ""
                Else
                        txt_fnacreport.Text = SELDFREPORT_RS("FNAC").Value
                End If
            Else
                txt_fnacreport.Text = ""
            End If
        End If
             
    '*****
        If SELDSPFNFLAG = False Then
            SELDSPFN_QR(0) = REQNO
            Set SELDSPFN_RS = SELDSPFN_QR.OpenResultset(1, 3)
            SELDSPFNFLAG = True
        Else
            SELDSPFN_QR(0) = REQNO
            SELDSPFN_RS.MoreResults
            SELDSPFN_RS.Requery
        End If
'        If SELDSPFN_RS.RowCount > 0 Then
'            If SELDFNREPORTFLAG = False Then
'                SELDFNREPORTQR(0) = SELDSPFN_RS("FNREQNO")
'                Set SELDFNREPORTRS = SELDFNREPORTQR.OpenResultset(1, 3)
'                SELDFNREPORTFLAG = True
'            Else
'                SELDFNREPORTQR(0) = SELDSPFN_RS("FNREQNO")
'                SELDFNREPORTRS.MoreResults
'                SELDFNREPORTRS.Requery
'            End If
'            If SELDFNREPORTRS.RowCount > 0 Then
'                If SELDFNREPORTRS.EOF = False Then
'                    txt_fnacreport.Text = SELDFNREPORTRS("FNAC").Value
'                    cmb_fnac.Text = "YES"
'                    SPFNFLAG = True
'                 Else
'                    txt_fnacreport.Text = ""
'                    cmb_fnac.Text = "NO"
'                End If
'            End If
'        End If
        '*****
             
    ''    End If
    '    seldfsreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    '    seldfsreport_rs.MoreResults
    '    seldfsreport_rs.Requery
    '    If seldfsreport_rs.RowCount > 0 Then
    '        txt_frosecrep.Text = seldfsreport_rs("fsreport")
    '    End If
    
    '    seldihreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    '    seldihreport_rs.MoreResults
    '    seldihreport_rs.Requery
    '    If seldihreport_rs.RowCount > 0 Then
    '        txt_immhischerep.Text = seldihreport_rs("ihreport")
    '    End If
    
        If REFDR_RS.RowCount > 0 Then 'MODIFIED
         If REFDR_RS.EOF = False Then 'REPLACE WITH SELDFREPORT_RS
            cmb_enteryby.Text = Trim(REFDR_RS("ENTERBY"))
            cmb_provby.Text = Trim(REFDR_RS("PREPORTBY"))
               If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                cmb_finalby.Text = ""
                Else
                cmb_finalby.Text = Trim(REFDR_RS("FREPORTBY"))
                End If
         End If
        Else
             cmb_finalby.Text = ""
        End If
        If SELDIMPRESIONFLAG = False Then
            seldimpresion_qr(0) = REQNO
            Set seldimpresion_rs = seldimpresion_qr.OpenResultset(1, 3)
            SELDIMPRESIONFLAG = True
        Else
            seldimpresion_qr(0) = REQNO
            seldimpresion_rs.MoreResults
            seldimpresion_rs.Requery
        End If
        If seldimpresion_rs.RowCount > 0 Then
               ' txt_impfindiag.Text = seldimpresion_rs("impression")
                cmb_fsd.AddItem seldimpresion_rs("FDESCRY")
                cmb_fsd.Text = seldimpresion_rs("FDESCRY")
                cmb_t.Text = seldimpresion_rs("T")
                txt_size.Text = seldimpresion_rs("TSIZE")
                CMB_N.Text = seldimpresion_rs("N")
                cmb_m.AddItem seldimpresion_rs("M")
                cmb_m.Text = seldimpresion_rs("M")
                CMB_ICDH.AddItem seldimpresion_rs("FINALDIAG")
                CMB_ICDH.Text = seldimpresion_rs("FINALDIAG")
                seldgrade_qr(0) = seldimpresion_rs("gradecode")
                Set seldgrade_rs = seldgrade_qr.OpenResultset(1, 3)
                If seldgrade_rs.RowCount > 0 Then
                     
                    cmb_grade.AddItem seldgrade_rs("gradedesc")
                    cmb_grade.Text = seldgrade_rs("gradedesc")
                End If
        End If
    
    '    seldgrossdesc_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    '    seldgrossdesc_rs.MoreResults
    '    seldgrossdesc_rs.Requery
    '    If seldgrossdesc_rs.RowCount > 0 Then
    '            txt_grossdesc = seldgrossdesc_rs("GROSSDESC")
    '            cmb_grossby.Text = seldgrossdesc_rs("GROSSEDBY")
    '    End If
    
    
    '    seldspreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    '    seldspreport_rs.MoreResults
    '    seldspreport_rs.Requery
    '    If seldspreport_rs.RowCount > 0 Then
    '            txt_microdesc = seldspreport_rs("SPREPORT")
    '            selempbasinfo1_qr(0) = seldspreport_rs("PREPORTBY")
    '            selempbasinfo1_rs.MoreResults
    '            selempbasinfo1_rs.Requery
    '        If selempbasinfo1_rs.RowCount > 0 Then
    '            cmb_provby.Clear
    '            cmb_provby.AddItem selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("PREPORTBY")
    '            cmb_provby.Text = selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("PREPORTBY")
    '        End If
    '        selempbasinfo1_qr(0) = seldspreport_rs("FREPORTBY")
    '        selempbasinfo1_rs.MoreResults
    '        selempbasinfo1_rs.Requery
    '        If selempbasinfo1_rs.RowCount > 0 Then
    '            cmb_finalby.Clear
    '            cmb_finalby.AddItem selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("FREPORTBY")
    '            cmb_finalby.Text = selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("FREPORTBY")
    '        End If
        'seldrequisitn_qr10.QueryTimeout = 140
        'If seldrequisitn_flag10 = False Then
            seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
         '   Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
          '  seldrequisitn_flag10 = True
        'Else
         '   seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
            seldrequisitn_rs10.MoreResults
            seldrequisitn_rs10.Requery
        'End If
        If seldrequisitn_rs10.RowCount > 0 Then
         prpathno = ""
         CMB_PATHNO.clear
         While Not seldrequisitn_rs10.EOF
            MODALITY1 = Mid(seldrequisitn_rs10("REQNNO"), 4, 2)
            If Trim(MODALITY1) = "SP" Then
                MODDESC = "SURGICAL PATHOLOGY"
            ElseIf Trim(MODALITY1) = "CN" Then
                MODDESC = "CYTOLOGY NON GYNAEC"
            ElseIf Trim(MODALITY1) = "CG" Then
                MODDESC = "CYTOLOGY GYNAEC"
            ElseIf Trim(MODALITY1) = "FA" Then
                MODDESC = "FINE NEEDLE ASPIRATION CYTOLOGY"
            ElseIf Trim(MODALITY1) = "CT" Then
                MODDESC = "CT SCAN"
            ElseIf Trim(MODALITY1) = "CR" Then
                MODDESC = "CONVENTIONAL RADIOGRAPHY"
            ElseIf Trim(MODALITY1) = "MM" Then
                MODDESC = "MAMMOGRAPHY"
            ElseIf Trim(MODALITY1) = "MR" Then
                MODDESC = "MRI SCAN"
            ElseIf Trim(MODALITY1) = "US" Then
                MODDESC = "ULTRASONOGRAPHY"
            End If
            If Trim(TXT_PATH.Text) <> Trim(seldrequisitn_rs10("pathno")) Or IsNull(seldrequisitn_rs10("pathno")) Then
                prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
                'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & seldrequisitn_rs10("PREPORTDT") & " " & "$" & " " & req_chglen(seldrequisitn_rs10("REQNNO")) & " " & "$" & " " & MODDESC
                CMB_PATHNO.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
                
            End If
            seldrequisitn_rs10.MoveNext
         Wend
            CMB_PATHNO.Text = CMB_PATHNO.List(0)
            txt_prepathreqno.Text = Trim(prpathno)
        End If
            SELDADDENDUM_QR(0) = REQNO
            SELDADDENDUM_RS.MoreResults
            SELDADDENDUM_RS.Requery
            If SELDADDENDUM_RS.RowCount > 0 Then
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = ADDENDUM + Chr(13) & "ADDENDUM NO " & Trim(SELDADDENDUM_RS("ADDEND_NO")) & Chr(13) & Trim(SELDADDENDUM_RS("ADD_MAX_LN")) & Chr(13) & Trim(SELDADDENDUM_RS("ADDEND_DT")) & " " & Trim(SELDADDENDUM_RS("ADDEND_TM")) & " " & Trim(SELDADDENDUM_RS("USER_ID"))
                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
        'End If
        cmb_reqno.Enabled = False
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
    End If

End If
If opt_staff.Value = True Then
    STAFF
End If
Exit Sub
'MSG:
'MsgBox Err.Description
End Sub

Private Sub Opt_patient_Click()
    If Opt_patient.Value = True Then
       lbl_caseno.Caption = "Case No."
        MaskEdBox1.Mask = "??>/#####"
    End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    MaskEdBox1.Mask = "#######"
End If
End Sub

Private Sub Timer1_Timer()
txt_time.Text = Format(Time, "HH:MM:SS")
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
w_adden.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_adden.Caption = "Addendum "
w_adden.WindowState = wdWindowStateMaximize
Set d_adden = w_adden.ActiveDocument
txt_addendum.Text = ""
d_adden.Content = Trim(txt_addendum.Text)
If ADDENDUMFALG = True Then
    d_adden.Content = ""
End If
w_adden.Visible = True
w_adden.Activate
End Sub

Private Sub txt_addendum_LostFocus()
On Error GoTo m12:
Set d_adden = w_adden.ActiveDocument
txt_addendum.Text = d_adden.Content

d_adden.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_adden.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_adden.Close (wdDoNotSaveChanges)
w_adden.quit

Exit Sub
m12:
MsgBox Err.Description
End Sub

Private Sub txt_fnacreport_GotFocus()
On Error GoTo m9:
WFNAC.Visible = False
Set DFNAC = WFNAC.ActiveDocument
DFNAC.Content = ""
DFNAC.Content = txt_fnacreport
WFNAC.Caption = "FNAC "
'w_imp.WindowState = wdWindowStateMaximize
WFNAC.WindowState = wdWindowStateNormal
DFNAC.Content = txt_fnacreport
WFNAC.Visible = True
WFNAC.Activate
m9:
On Error Resume Next:
Set WFNAC = CreateObject("word.Application")
WFNAC.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
WFNAC.Caption = "Impression"
'w_imp.WindowState = wdWindowStateMaximize
WFNAC.WindowState = wdWindowStateNormal
Set DFNAC = WFNAC.ActiveDocument
DFNAC.Content = txt_fnacreport
WFNAC.Visible = True
WFNAC.Activate
End Sub

Private Sub txt_fnacreport_LostFocus()
On Error GoTo m2:
Set DFNAC = WFNAC.ActiveDocument
txt_fnacreport.Text = DFNAC.Content
'flagimp = False
DFNAC.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DFNAC.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DFNAC.Close (wdDoNotSaveChanges)
WFNAC.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Private Sub txt_frosecrep_GotFocus()
On Error GoTo m1:
w_frozen.Visible = False
Set d_frozen = w_frozen.ActiveDocument
d_frozen.Content = ""
d_frozen.Content = txt_frosecrep
w_frozen.Caption = "frozen"
'w_frozen.WindowState = wdWindowStateMaximize
w_frozen.WindowState = wdWindowStateNormal
d_frozen.Content = txt_frosecrep
w_frozen.Visible = True
w_frozen.Activate
m1:
On Error Resume Next
Set w_frozen = CreateObject("Word.Application")
w_frozen.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_frozen.Caption = "frozen"
'w_frozen.WindowState = wdWindowStateMaximize
w_frozen.WindowState = wdWindowStateNormal
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
' Modified By Pankaj Baggal 03/09/2004
'd_frozen.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_frozen.Close (wdDoNotSaveChanges)
w_frozen.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
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
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
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
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub txt_grossdesc_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_grossdesc
W_grossdes.Caption = "GrossDesc"
'W_grossdes.WindowState = wdWindowStateMaximize
W_grossdes.WindowState = wdWindowStateNormal
d_grossdesc.Content = txt_grossdesc
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "GrossDesc"
W_grossdes.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
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
' Modified By Pankaj Baggal 03/09/2004
'd_grossdesc.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_grossdesc.Close (wdDoNotSaveChanges)
W_grossdes.quit
'W1.Visible = False
Exit Sub
m4:
MsgBox Err.Description
End Sub

Private Sub txt_immhischerep_GotFocus()
Dim ihrep As String
selihreportqr(0) = req_chglen(txt_immhischereqno.Text)
Set selihreportrs = selihreportqr.OpenResultset(1, 3)
If selihreportrs.RowCount > 0 Then
    While Not selihreportrs.EOF
        ihrep = ihrep & selihreportrs("testorder")
        selihreportrs.MoveNext
    Wend
    txt_immhischerep.Text = Trim(ihrep)
End If

On Error GoTo m7:
w_immun.Visible = False
Set d_immun = w_immun.ActiveDocument
d_immun.Content = ""
d_immun.Content = txt_immhischerep
w_immun.Caption = "Immuno"
w_immun.WindowState = wdWindowStateNormal
'w_immun.WindowState = wdWindowStateMaximize
d_immun.Content = txt_immhischerep
w_immun.Visible = True
w_immun.Activate
m7:
On Error Resume Next:
Set w_immun = CreateObject("word.Application")
w_immun.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w_immun.Caption = "Immuno"
w_immun.WindowState = wdWindowStateMaximize
w_immun.WindowState = wdWindowStateNormal
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
' Modified By Pankaj Baggal 03/09/2004
'd_immun.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_immun.Close (wdDoNotSaveChanges)
w_immun.quit
'W1.Visible = False
Exit Sub
m8:
MsgBox Err.Description
End Sub



Private Sub txt_impfindiag_GotFocus()
On Error GoTo m9:
w_imp.Visible = False
Set d_imp = w_imp.ActiveDocument
d_imp.Content = ""
d_imp.Content = txt_impfindiag
w_imp.Caption = "Impression"
'w_imp.WindowState = wdWindowStateMaximize
w_imp.WindowState = wdWindowStateNormal
d_imp.Content = txt_impfindiag
w_imp.Visible = True
w_imp.Activate
m9:
On Error Resume Next:
Set w_imp = CreateObject("word.Application")
w_imp.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w_imp.Caption = "Impression"
'w_imp.WindowState = wdWindowStateMaximize
w_imp.WindowState = wdWindowStateNormal
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
' Modified By Pankaj Baggal 03/09/2004
'd_imp.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_imp.Close (wdDoNotSaveChanges)
w_imp.quit
'W1.Visible = False
Exit Sub
m12:
MsgBox Err.Description
End Sub

Private Sub txt_microdesc_GotFocus()
On Error GoTo m5:
w_micscop.Visible = False
Set d_micscop = w_micscop.ActiveDocument
d_micscop.Content = ""
d_micscop.Content = txt_microdesc
w_micscop.Caption = "MicroDesc"
'w_micscop.WindowState = wdWindowStateMaximize
w_micscop.WindowState = wdWindowStateNormal
d_micscop.Content = txt_microdesc
w_micscop.Visible = True
w_micscop.Activate
m5:
On Error Resume Next:
Set w_micscop = CreateObject("word.Application")
w_micscop.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w_micscop.Caption = "Microsocpic Description"
w_micscop.WindowState = wdWindowStateNormal
'w_micscop.WindowState = wdWindowStateMaximize
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
' Modified By Pankaj Baggal 03/09/2004
'd_micscop.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_micscop.Close (wdDoNotSaveChanges)
w_micscop.quit
'W1.Visible = False
Exit Sub
m6:
MsgBox Err.Description
End Sub

Public Sub save()

End Sub

Public Sub assign()
''msgtype_hl7 = "PrelimRepModify"
'On Error GoTo MSG:
Dim i As Integer
Dim MICROSCOPIC As String
                caseno_hl7 = MaskEdBox1.FormattedText
                reqnno_hl7 = req_chglen(cmb_reqno.Text)
' DECLARATIONS SHIFTED TO GENERAL DECLARATIONS

'            '***for HL7*****24122003
'                Dim insas400_hl7_rq As rdoQuery
'                Dim caseno_hl7 As String
'                Dim reqnno_hl7 As String
'                caseno_hl7 = msk_caseno.FormattedText
'                reqnno_hl7 = req_chglen(cbo_reqno.Text)
'                Dim insas400_hl7_rs As rdoResultset
'                Dim DTTIM As rdoResultset
'                Dim msgtype_hl7 As String
'
'                Dim str_hl7 As String
'                Dim uniqueid_hl7 As String
'              '*******hl7 declarations till here
rdb.BeginTrans
If flagsave = 1 Then
   flagsave = 0
If SELSITEDETAIL_FLAG = False Then
    SELSITEDETAIL_QR(0) = Trim(cmb_sites.Text)
    Set SELSITEDETAIL_RS = SELSITEDETAIL_QR.OpenResultset(1, 3)
    SELSITEDETAIL_FLAG = True
Else
    SELSITEDETAIL_QR(0) = Trim(cmb_sites.Text)
    SELSITEDETAIL_RS.MoreResults
    SELSITEDETAIL_RS.Requery
End If
If Not SELSITEDETAIL_RS.EOF Then
    INSSITEDETAIL_QR(0) = Trim(REQNO)
    INSSITEDETAIL_QR(1) = Trim(SELSITEDETAIL_RS("SCODE"))
    INSSITEDETAIL_QR(2) = Trim(SELSITEDETAIL_RS("SDESC"))
    INSSITEDETAIL_QR(3) = Trim(SELSITEDETAIL_RS("NCODE"))
    INSSITEDETAIL_QR(4) = Trim(cmb_laterality.Text)
    If Trim(TXT_FSDIS.Text) = "NO" Then
        INSSITEDETAIL_QR(5) = "NO"
    Else
        INSSITEDETAIL_QR(5) = Trim(TXT_FSDIS.Text)
    End If
    INSSITEDETAIL_QR(6) = Trim(cmb_pendingremarks.Text)
    INSSITEDETAIL_QR(7) = ""
    INSSITEDETAIL_QR(8) = ""
    INSSITEDETAIL_QR(9) = ""
    INSSITEDETAIL_QR.Execute
End If
'If Len(txt_impfindiag) > 0 Then
    insdimpresion_qr(0) = Trim(REQNO)
''    insdimpresion_qr(1) = Null
''    insdimpresion_qr(1) = txt_impfindiag
 If Len(Trim(CMB_ICDH)) = 0 Then
    insdimpresion_qr(1) = "NULL"
    Else
    insdimpresion_qr(1) = CMB_ICDH.Text
 End If
    'I = InStr(cmb_fsd.Text, "$")
    'insdimpresion_qr(2) = Mid(cmb_fsd, I + 1, 6)
    insdimpresion_qr(2) = cmb_fsd
    insdimpresion_qr(3) = cmb_t.Text
    insdimpresion_qr(4) = txt_size.Text
    insdimpresion_qr(5) = CMB_N
    insdimpresion_qr(6) = cmb_m
    i = InStr(cmb_grade, "$")
    If i = 0 Then
        insdimpresion_qr(7) = cmb_grade.Text
    Else
        insdimpresion_qr(7) = Mid(cmb_grade, i + 1, 2)
    End If
    insdimpresion_qr.Execute
'End If
''    '**** INSERTING INTO FREPORT
        INSFREPORT_QR(0) = Trim(REQNO)
        If Len(txt_frosecrep.Text) > 0 Then
            INSFREPORT_QR(1).Value = Null
            INSFREPORT_QR(1) = Trim(txt_frosecrep.Text)
        Else
            INSFREPORT_QR(1).Value = Null
            INSFREPORT_QR(1) = " "
        End If
        If Len(txt_immhischerep.Text) > 0 Then
            INSFREPORT_QR(2).Value = Null
            INSFREPORT_QR(2) = Trim(txt_immhischerep.Text)
        Else
            INSFREPORT_QR(2).Value = Null
            INSFREPORT_QR(2) = " "
        End If
        If Len(txt_impfindiag.Text) > 0 Then
            INSFREPORT_QR(3).Value = Null
            INSFREPORT_QR(3) = Trim(txt_impfindiag.Text)
        Else
            INSFREPORT_QR(3).Value = Null
            INSFREPORT_QR(3) = " "

        End If
        If Len(txt_grossdesc.Text) > 0 Then
            INSFREPORT_QR(4).Value = Null
            INSFREPORT_QR(4) = Trim(txt_grossdesc.Text)
            INSFREPORT_QR(5) = cmb_grossby.Text
        Else
            INSFREPORT_QR(4).Value = Null
            INSFREPORT_QR(4) = ""
            INSFREPORT_QR(5) = cmb_grossby.Text
        End If
        If Len(txt_microdesc.Text) > 0 Then
            INSFREPORT_QR(6).Value = Null
            INSFREPORT_QR(6) = Trim(txt_microdesc.Text)

        Else
            INSFREPORT_QR(6).Value = Null
            INSFREPORT_QR(6) = " "

        End If
            INSFREPORT_QR(7) = "1"
        If Len(txt_fnacreport.Text) > 0 Then
            INSFREPORT_QR(8).Value = Null
            INSFREPORT_QR(8) = Trim(txt_fnacreport.Text)

        Else
            INSFREPORT_QR(8).Value = Null
            INSFREPORT_QR(8) = " "

        End If
            INSFREPORT_QR.Execute
    '**** INSERTING INTOFREPORT
    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provby.Text

    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(REQNO)
    updrrequisitn_qr.Execute
    
    upddrequisitn_qr(0) = Trim(txt_prepathreqno.Text)
    If Len(Trim(txt_natmeterial.Text)) > 0 Then
    upddrequisitn_qr(1) = Null
    upddrequisitn_qr(1) = Trim(txt_natmeterial.Text)
    Else
    upddrequisitn_qr(1) = Null
    upddrequisitn_qr(1) = " "
    End If
    'upddrequisitn_qr(2) = Trim(cmb_immhische.Text)
    upddrequisitn_qr(2) = Trim(REQNO)
    upddrequisitn_qr.Execute
    
'    ' INSERTING INTO DFNREPORT
'    If CMB_FNAC.Text = "YES" Then
'
'            INSFNREPORTQR(0) = req_chglen(FNACREQNO)
'        If Len(Trim(txt_fnacreport.Text)) > 0 Then
'            INSFNREPORTQR(1) = Null
'            INSFNREPORTQR(1) = txt_fnacreport.Text
'        Else
'            INSFNREPORTQR(1) = Null
'            INSFNREPORTQR(1) = " "
'        End If
'            INSFNREPORTQR.Execute
'         INSDSPFNREQQR(0) = Trim(REQNO)
'         INSDSPFNREQQR(1) = req_chglen(FNACREQNO)
'         INSDSPFNREQQR.Execute
'    End If
'    ' INSERTING INTO DFNREPORT
    '**** INSERTING INTO DISIMP
    If Len(Trim(txt_impfindiag.Text)) > 0 Then
        INSDISIMP_QR(0) = Trim(REQNO)
        INSDISIMP_QR(1) = Null
        INSDISIMP_QR(1) = Trim(txt_impfindiag.Text)
        INSDISIMP_QR.Execute
    End If
    '**** INSERTING INTO DISIMP
        '' HL7INS
      SELPACSLOG_RQ(0) = REQNO
      Set SELPACSLOG_RS = SELPACSLOG_RQ.OpenResultset(1, 1)
      If SELPACSLOG_RS.RowCount > 0 Then
                caseno_hl7 = MaskEdBox1.FormattedText
                reqnno_hl7 = req_chglen(cmb_reqno.Text)

                msgtype_hl7 = "PrelimRepOrder" 'Form1.lbl_type.Tag
                str_hl7 = WIN_CLIENT.LocalIP
                ''
                 Set DTTIM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE  , CURRENT TIME  FROM PATOOLSLIB.MA107") '***2/2/99
 

             uniqueid_hl7 = str_hl7 & "-" & Format(DTTIM(0), "dd/mm/yyyy") & "-" & Format(DTTIM(1), "hh:mm:ss")
      
                
       
                Set insas400_hl7_rq = rdb.CreateQuery("ins as4001_pr111", "insert into hl7lib.hl7tab3(UNQMSGID,MSGTYPID,CASENO,CLNTIP,CLNTDATE,CLNTTIME,REQNNO,PROCBRDGYN) values(?,?,?,?,?,?,?,?)")
                insas400_hl7_rq(0) = uniqueid_hl7
                insas400_hl7_rq(1) = msgtype_hl7
                insas400_hl7_rq(2) = caseno_hl7 '  msk_caseno.FormattedText
                insas400_hl7_rq(3) = WIN_CLIENT.LocalIP
                
                insas400_hl7_rq(4) = Format(DTTIM(0), "dd/mm/yyyy")
                insas400_hl7_rq(5) = Format(DTTIM(1), "hh:mm:ss")
                
                insas400_hl7_rq(6) = reqnno_hl7
                insas400_hl7_rq(7) = "N"
                insas400_hl7_rq.Execute
                   delay (2)
                 
                caseno_hl7 = MaskEdBox1.FormattedText
                reqnno_hl7 = req_chglen(cmb_reqno.Text)
                msgtype_hl7 = "Preliminary Rep" 'Form1.lbl_type.Tag
                str_hl7 = WIN_CLIENT.LocalIP
                '''
                   str_hl7 = WIN_CLIENT.LocalIP
                ''
                 Set DTTIM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE  , CURRENT TIME  FROM PATOOLSLIB.MA107") '***2/2/99
 
            
             uniqueid_hl7 = str_hl7 & "-" & Format(DTTIM(0), "dd/mm/yyyy") & "-" & Format(DTTIM(1), "hh:mm:ss")
              Set insas400_hl7_rq = rdb.CreateQuery("ins as4001_pr1", "insert into hl7lib.hl7tab3(UNQMSGID,MSGTYPID,CASENO,CLNTIP,CLNTDATE,CLNTTIME,REQNNO,PROCBRDGYN) values(?,?,?,?,?,?,?,?)")
                insas400_hl7_rq(0) = uniqueid_hl7
                insas400_hl7_rq(1) = msgtype_hl7
                insas400_hl7_rq(2) = MaskEdBox1.FormattedText
                insas400_hl7_rq(3) = WIN_CLIENT.LocalIP
                insas400_hl7_rq(4) = Format(DTTIM(0), "dd/mm/yyyy")
                insas400_hl7_rq(5) = Format(DTTIM(1), "hh:mm:ss")
                
                
                insas400_hl7_rq(6) = reqnno_hl7
                insas400_hl7_rq(7) = "N"
                insas400_hl7_rq.Execute
                 
       '  UPDATING PACSLOG
          UPDPACSLOGPREP_RQ(0) = "Y"
          UPDPACSLOGPREP_RQ(1) = Format(DTTIM(0), "DD/MM/YYYY")
          UPDPACSLOGPREP_RQ(2) = Format(DTTIM(1), "HH:MM:SS")
          UPDPACSLOGPREP_RQ(3) = REQNO
          UPDPACSLOGPREP_RQ.Execute
        '  UPDATING PACSLOG over
        
          
      
      End If '' HL7 ENDIF
      
    '' HL7INS
    End If
      

If flagsave = 2 Then
    flagsave = 0
'''If Len(txt_frosecrep.Text) > 0 Then
'''
'''    upddfsreport_qr(0) = Null
'''    upddfsreport_qr(0) = Trim(txt_frosecrep.Text)
'''    upddfsreport_qr(1) = Trim(req_chglen(cmb_reqno))
'''    upddfsreport_qr.Execute
'''End If
'''
'''If Len(txt_grossdesc.Text) > 0 Then
'''
'''    updgrossdesc_qr(0) = Null
'''    updgrossdesc_qr(0) = txt_grossdesc
'''    updgrossdesc_qr(1) = cmb_grossby
'''    updgrossdesc_qr(2) = Trim(req_chglen(cmb_reqno))
'''    updgrossdesc_qr.Execute
''' End If
'''
'''If Len(txt_immhischerep.Text) > 0 Then
'''
'''    upddihreport_qr(0) = Null
'''    upddihreport_qr(0) = Trim(txt_immhischerep.Text)
'''    I = InStr(cmb_enteryby.Text, "$")
'''    upddihreport_qr(1) = Mid(cmb_enteryby, I + 1, 6)
'''    I = InStr(cmb_provby, "$")
'''    upddihreport_qr(2) = Mid(cmb_provby, I + 1, 6)
'''    I = InStr(cmb_finalby, "$")
'''    upddihreport_qr(3) = Mid(cmb_finalby, I + 1, 6)
'''    upddihreport_qr(4) = Trim(req_chglen(cmb_reqno))
'''    upddihreport_qr.Execute
'''    End If
'''
'''    If Len(txt_microdesc) > 0 Then
'''    MICROSCOPIC = RTrim(txt_microdesc.Text)
'''    UPDSPREPORT_QR(0) = Null
'''    UPDSPREPORT_QR(0) = MICROSCOPIC
'''    I = InStr(cmb_enteryby.Text, "$")
'''    UPDSPREPORT_QR(1) = Mid(cmb_enteryby.Text, I + 1, 6)
'''    I = InStr(cmb_provby.Text, "$")
'''    UPDSPREPORT_QR(2) = Mid(cmb_provby.Text, I + 1, 6)
'''    I = InStr(cmb_finalby.Text, "$")
'''    UPDSPREPORT_QR(3) = Mid(cmb_finalby.Text, I + 1, 6)
'''    UPDSPREPORT_QR(4) = req_chglen(Trim(cmb_reqno.Text))
'''    UPDSPREPORT_QR.Execute
'''
'''    End If
'''
    If SELSITEDETAIL_FLAG = False Then
        SELSITEDETAIL_QR(0) = Trim(cmb_sites.Text)
        Set SELSITEDETAIL_RS = SELSITEDETAIL_QR.OpenResultset(1, 3)
        SELSITEDETAIL_FLAG = True
    Else
        SELSITEDETAIL_QR(0) = Trim(cmb_sites.Text)
        SELSITEDETAIL_RS.MoreResults
        SELSITEDETAIL_RS.Requery
    End If
    If Not SELSITEDETAIL_RS.EOF Then
        UPDSITEDETAIL_QR(0) = SELSITEDETAIL_RS("SCODE")
        UPDSITEDETAIL_QR(1) = Trim(cmb_sites.Text)
        UPDSITEDETAIL_QR(2) = SELSITEDETAIL_RS("NCODE")
        UPDSITEDETAIL_QR(3) = Trim(cmb_laterality.Text)
        UPDSITEDETAIL_QR(4) = Trim(TXT_FSDIS.Text)
        UPDSITEDETAIL_QR(5) = Trim(REQNO)
        UPDSITEDETAIL_QR.Execute
    End If
    
    
    If Len(Trim(txt_impfindiag)) > 0 Then

    
        
        
        upddimpresion_qr(0) = Trim(CMB_ICDH.Text)
        upddimpresion_qr(1) = Trim(cmb_fsd.Text)
        upddimpresion_qr(2) = cmb_t.Text
        upddimpresion_qr(3) = txt_size.Text
        upddimpresion_qr(4) = CMB_N
        upddimpresion_qr(5) = cmb_m
        i = InStr(cmb_grade, "$")
        upddimpresion_qr(6) = Mid(cmb_grade, i + 1, 2) '
        
        upddimpresion_qr(7) = Trim(REQNO)
        upddimpresion_qr.Execute
    End If
    
    updrrequisitn_qr(0) = Trim(cmb_enteryby.Text)
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = Trim(cmb_provby.Text)
'    updrrequisitn_qr(3) = cmb_finalby.Text
'    updrrequisitn_qr(4) = GET_REVDATE(Format(date11, "DD/MM/YYYY"))
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = Trim(cmb_finalby.Text)
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(REQNO)
    updrrequisitn_qr.Execute
    
    If (cmb_immhische.Text) = "YES" And Len(Trim(txt_immhischereqno.Text)) > 0 Then
     updrrequisitn_qr(0) = cmb_enteryby.Text
     updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
     updrrequisitn_qr(2) = cmb_provby.Text
'    updrrequisitn_qr(3) = cmb_finalby.Text
'    updrrequisitn_qr(4) = GET_REVDATE(Format(date11, "DD/MM/YYYY"))
     updrrequisitn_qr(3) = "T"
     updrrequisitn_qr(4) = Trim(req_chglen(txt_immhischereqno.Text))
     updrrequisitn_qr.Execute
   End If
    
    
     '**** updating INTO FREPORT
        
        If Len(txt_frosecrep.Text) > 0 Then
            upddfsreport_qr(0).Value = Null
            upddfsreport_qr(0) = Trim(txt_frosecrep.Text)
        Else
            upddfsreport_qr(0).Value = Null
            upddfsreport_qr(0) = " "
        End If
        If Len(txt_immhischerep.Text) > 0 Then
            upddfsreport_qr(1).Value = Null
            upddfsreport_qr(1) = Trim(txt_immhischerep.Text)
        Else
            upddfsreport_qr(1).Value = Null
            upddfsreport_qr(1) = " "
        End If
        If Len(txt_impfindiag.Text) > 0 Then
            upddfsreport_qr(2).Value = Null
            upddfsreport_qr(2) = Trim(txt_impfindiag.Text)
        Else
            upddfsreport_qr(2).Value = Null
            upddfsreport_qr(2) = " "

        End If
        If Len(txt_grossdesc.Text) > 0 Then
            upddfsreport_qr(3).Value = Null
            upddfsreport_qr(3) = Trim(txt_grossdesc.Text)
            
        Else
            upddfsreport_qr(3).Value = Null
            upddfsreport_qr(3) = " "
            
        End If
        If Len(txt_microdesc.Text) > 0 Then
            upddfsreport_qr(4).Value = Null
            upddfsreport_qr(4) = Trim(txt_microdesc.Text)

        Else
            upddfsreport_qr(4).Value = Null
           upddfsreport_qr(4) = " "

        End If
            upddfsreport_qr(5) = "1"
       If Len(txt_fnacreport.Text) > 0 Then
            upddfsreport_qr(6).Value = Null
            upddfsreport_qr(6) = Trim(txt_fnacreport.Text)

        Else
            upddfsreport_qr(6).Value = Null
           upddfsreport_qr(6) = " "

        End If
            upddfsreport_qr(7) = Trim(REQNO)
            upddfsreport_qr.Execute
    '**** updating INTO FREPORT
    
    upddrequisitn_qr(0) = Trim(txt_prepathreqno.Text)
    If Len(Trim(txt_natmeterial.Text)) > 0 Then
        upddrequisitn_qr(1) = Null
        upddrequisitn_qr(1) = Trim(txt_natmeterial.Text)
    Else
        upddrequisitn_qr(1) = Null
        upddrequisitn_qr(1) = " "
    End If
    upddrequisitn_qr(2) = Trim(REQNO)
    upddrequisitn_qr.Execute
    
'    If CMB_FNAC.Text = "YES" Then
'         INSDSPFNREQQR(0) = REQNO
'         INSDSPFNREQQR(1) = req_chglen(FNACREQNO)
'         INSDSPFNREQQR.Execute


'        If SPFNFLAG = False Then
'                INSFNREPORTQR(0) = req_chglen(FNACREQNO)
'            If Len(Trim(txt_fnacreport.Text)) > 0 Then
'                INSFNREPORTQR(1) = Null
'                INSFNREPORTQR(1) = txt_fnacreport.Text
'            Else
'                INSFNREPORTQR(1) = Null
'                INSFNREPORTQR(1) = " "
'            End If
'                INSFNREPORTQR.Execute
'                INSDSPFNREQQR(0) = req_chglen(cmb_reqno.Text)
'                INSDSPFNREQQR(1) = req_chglen(FNACREQNO)
'                INSDSPFNREQQR.Execute
'        Else
'            If Len(Trim(txt_fnacreport.Text)) > 0 Then
'                UPDFNREPORTQR(0) = Null
'                UPDFNREPORTQR(0) = txt_fnacreport
'            Else
'                UPDFNREPORTQR(0) = Null
'                UPDFNREPORTQR(0) = " "
'            End If
'            UPDFNREPORTQR(1) = req_chglen(FNACREQNO)
'            UPDFNREPORTQR.Execute
'        End If
'    End If
    '**** UPDATE THE DISIMP
    If Len(Trim(txt_impfindiag.Text)) > 0 Then
        UPDDISIMP_QR(0) = Null
        UPDDISIMP_QR(0) = Trim(txt_impfindiag.Text)
        UPDDISIMP_QR(1) = Trim(REQNO)
        UPDDISIMP_QR.Execute
    End If
    '*** UPDATE THE DISIMP
        '' HL7INSMOD
      SELPACSLOG_RQ(0) = REQNO
      Set SELPACSLOG_RS = SELPACSLOG_RQ.OpenResultset(1, 1)
      If SELPACSLOG_RS.RowCount > 0 Then
             caseno_hl7 = MaskEdBox1.FormattedText
                reqnno_hl7 = req_chglen(cmb_reqno.Text)
                
                msgtype_hl7 = "PrelimRepModify" 'Form1.lbl_type.Tag
                str_hl7 = WIN_CLIENT.LocalIP
                '''
                   str_hl7 = WIN_CLIENT.LocalIP
                ''
                 Set DTTIM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE  , CURRENT TIME  FROM PATOOLSLIB.MA107") '***2/2/99
 
            
             uniqueid_hl7 = str_hl7 & "-" & Format(DTTIM(0), "dd/mm/yyyy") & "-" & Format(DTTIM(1), "hh:mm:ss")
              Set insas400_hl7_rq = rdb.CreateQuery("ins as4001_pr11", "insert into hl7lib.hl7tab3(UNQMSGID,MSGTYPID,CASENO,CLNTIP,CLNTDATE,CLNTTIME,REQNNO,PROCBRDGYN) values(?,?,?,?,?,?,?,?)")
                insas400_hl7_rq(0) = uniqueid_hl7
                insas400_hl7_rq(1) = msgtype_hl7
                insas400_hl7_rq(2) = caseno_hl7 'msk_caseno.FormattedText
                insas400_hl7_rq(3) = WIN_CLIENT.LocalIP
                insas400_hl7_rq(4) = Format(DTTIM(0), "dd/mm/yyyy")
                insas400_hl7_rq(5) = Format(DTTIM(1), "hh:mm:ss")
                
                
                insas400_hl7_rq(6) = reqnno_hl7
                insas400_hl7_rq(7) = "N"
                insas400_hl7_rq.Execute
          '  UPDATING PACSLOG
          UPDPACSLOGPREP_RQ(0) = "Y"
          UPDPACSLOGPREP_RQ(1) = Format(DTTIM(0), "DD/MM/YYYY")
          UPDPACSLOGPREP_RQ(2) = Format(DTTIM(1), "HH:MM:SS")
          UPDPACSLOGPREP_RQ(3) = REQNO
          UPDPACSLOGPREP_RQ.Execute
        '  UPDATING PACSLOG over
      End If 'HL7MOD
      
      
       
    '' HL7INSMOD
End If
If ADDENDUMFALG = True Then
    ''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
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
    If Len(Trim(cmb_addendumby.Text)) = 0 Then
        MsgBox "The Addendum by field is EMPTY.", vbCritical
        cmb_addendumby.SetFocus
        Exit Sub
    End If
''********SECURITY***************************
    SELDADDENDUM_QR(0) = REQNO
    SELDADDENDUM_RS.MoreResults
    SELDADDENDUM_RS.Requery
    If SELDADDENDUM_RS.RowCount > 0 Then
       SELDADDENDUM_RS.MoveLast
        INSDADDENDUM_QR(0) = REQNO
        INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
        
        UPDTADNDMTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDTADNDMTIME_QR(1) = time1
        UPDTADNDMTIME_QR(2) = REQNO
        UPDTADNDMTIME_QR.Execute
    Else
        INSDADDENDUM_QR(0) = REQNO
        INSDADDENDUM_QR(1) = "1"
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
        
        UPDTADNDMTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDTADNDMTIME_QR(1) = time1
        UPDTADNDMTIME_QR(2) = REQNO
        UPDTADNDMTIME_QR.Execute
        
    End If
    ' hl7 dis code
      SELPACSLOG_RQ(0) = REQNO
      Set SELPACSLOG_RS = SELPACSLOG_RQ.OpenResultset(1, 1)
      If SELPACSLOG_RS.RowCount > 0 Then
                   '*******Inserting into HL7 Table**24122003
            
'
                caseno_hl7 = MaskEdBox1.FormattedText
                reqnno_hl7 = req_chglen(cmb_reqno.Text)

                msgtype_hl7 = "Addendum" 'Form1.lbl_type.Tag
                
                str_hl7 = WIN_CLIENT.LocalIP
                '''
                     Set DTTIM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE  , CURRENT TIME  FROM PATOOLSLIB.MA107") '***2/2/99
 
            'uniqueid_hl7 = str_hl7 & "-" & Format(Date, "dd/mm/yyyy") & "-" & Format(Time, "hh:mm:ss")
             uniqueid_hl7 = str_hl7 & "-" & Format(DTTIM(0), "dd/mm/yyyy") & "-" & Format(DTTIM(1), "hh:mm:ss")
                
                
                
            
                Set insas400_hl7_rq = rdb.CreateQuery("ins as400_ad", "insert into hl7lib.hl7tab3(UNQMSGID,MSGTYPID,CASENO,CLNTIP,CLNTDATE,CLNTTIME,REQNNO,PROCBRDGYN) values(?,?,?,?,?,?,?,?)")
                insas400_hl7_rq(0) = uniqueid_hl7
                insas400_hl7_rq(1) = msgtype_hl7
                insas400_hl7_rq(2) = caseno_hl7  'msk_caseno.FormattedText
                insas400_hl7_rq(3) = WIN_CLIENT.LocalIP
                insas400_hl7_rq(4) = Format(DTTIM(0), "dd/mm/yyyy")
                insas400_hl7_rq(5) = Format(DTTIM(1), "hh:mm:ss")
                insas400_hl7_rq(6) = reqnno_hl7
                insas400_hl7_rq(7) = "N"
                insas400_hl7_rq.Execute
               
       
                      
            '***********Till Here HL7*************
      '  UPDATING PACSLOG
          UPDPACSLOGADD_RQ(0) = "Y"
          UPDPACSLOGADD_RQ(1) = Format(DTTIM(0), "DD/MM/YYYY")
          UPDPACSLOGADD_RQ(2) = Format(DTTIM(1), "HH:MM:SS")
          UPDPACSLOGADD_RQ(3) = REQNO
          UPDPACSLOGADD_RQ.Execute
        '  UPDATING PACSLOG over
       
End If  ' FOR PACSLOG READ

 MsgBox "Addendum Data successfully sent to AS/400 - HL7 Table", vbInformation
 
' sendRMSG REQNO, MaskEdBox1.FormattedText, "A"
' hl7 dis code over

End If

rdb.CommitTrans
MsgBox "Record is saved Successfully."
Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
    
    
End Sub

Public Sub field_clear()
Dim obj As Object
cmb_laterality.Text = ""
'cmb_laterality.ListIndex = -1
cmb_sites.Text = ""
'cmb_sites.ListIndex = -1
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
cmb_reqno.ListIndex = -1
cmb_reqno.clear
cmb_reportdoc.ListIndex = -1
cmb_addendumby.ListIndex = -1
cmb_addendumby.Text = ""
cmb_enteryby.ListIndex = -1
'cmb_enteryby.Text = ""
cmb_provby.ListIndex = -1
cmb_provby.Text = ""
cmb_finalby.ListIndex = -1
cmb_finalby.Text = ""
cmb_reg.Text = ""
cmb_grade.ListIndex = -1
'CMB_FNAC.Text = ""
cmb_immhische.ListIndex = -1
cmb_immhische.Text = ""
LBL_VPP.Caption = ""
'cmb_immhische.clear
cmb_t.ListIndex = -1
cmb_t.clear
CMB_N.clear
CMB_N.ListIndex = -1
cmb_m.ListIndex = -1
cmb_grossby.ListIndex = -1
cmb_fsd.ListIndex = -1
CMB_ICDH.ListIndex = -1
CMB_ICDH.Text = ""
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
MaskEdBox1.Mask = ">??/#####"
cmb_grossby.ListIndex = -1
CMB_PATHNO.clear
CMB_PATHNO.ListIndex = -1
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In SPPATHREPORT.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub
Public Function delay(DLYTIME As String)
Dim sttime
sttime = Timer

Do While Timer < sttime + DLYTIME
  DoEvents
Loop





End Function
Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    cmb_reg.Enabled = True
    cmb_sites.Enabled = True
    cmb_laterality.Enabled = True
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
    cmb_grade.Enabled = True
    CMB_PATHNO.Enabled = True
    txt_natmeterial.Enabled = True
    'CMB_FNAC.Enabled = True
    txt_fnacreport.Enabled = False
End If
If str1 = "DIS" Then
    cmb_reg.Enabled = False
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
    cmb_grade.Enabled = False
    'CMB_PATHNO.Enabled = False
    txt_natmeterial.Enabled = False
    'CMB_FNAC.Enabled = False
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = False
    cmb_laterality.Enabled = False
    cmb_reg.Enabled = False
End If
If str1 = "VIEW" Then
    cmb_reg.Enabled = False
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
    cmb_grade.Enabled = False
    'CMB_PATHNO.Enabled = False
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_natmeterial.Enabled = False
    'CMB_FNAC.Enabled = False
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = False
    cmb_laterality.Enabled = False
End If
If str1 = "MODIFY" Then
    cmb_reg.Enabled = True
    txt_frosecrep.Enabled = True
    txt_grossdesc.Enabled = True
    txt_microdesc.Enabled = True
    txt_immhischerep.Enabled = True
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = True
    cmb_immhische.Enabled = True
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
    cmb_grade.Enabled = True
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = True
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    CMB_PATHNO.Enabled = True
    txt_natmeterial.Enabled = True
    'CMB_FNAC.Enabled = True
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = True
    cmb_laterality.Enabled = True
End If
If str1 = "ADDENDUM" Then
    cmb_reg.Enabled = False
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
    'CMB_PATHNO.Enabled = False
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    cmb_grade.Enabled = False
    txt_natmeterial.Enabled = False
    'CMB_FNAC.Enabled = False
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = False
    cmb_laterality.Enabled = False
End If
End Sub

Public Sub STAFF()
'On Error GoTo MSG:
If NEWFLAG = True Then
    NEWFLAG = False
    
    If SELDREQSTAFFD1FLAG1 = False Then
        SELDREQSTAFFD1QR1(0) = REQNO
        SELDREQSTAFFD1QR1(1) = MaskEdBox1.Text
        Set SELDREQSTAFFD1RS1 = SELDREQSTAFFD1QR1.OpenResultset(1, 3)
    Else
        SELDREQSTAFFD1QR1(0) = REQNO
        SELDREQSTAFFD1QR1(1) = MaskEdBox1.Text
        SELDREQSTAFFD1RS1.MoreResults
        SELDREQSTAFFD1RS1.Requery
    End If
    If SELDREQSTAFFD1RS1.RowCount > 0 Then
        txt_name = SELDREQSTAFFD1RS1("PNAME")
        txt_sex = SELDREQSTAFFD1RS1("SEX")
        txt_age = DateDiff("YYYY", SELDREQSTAFFD1RS1("AGE"), Format(DATE11, "DD/MM/YYYY"))
        If SELDREQSTAFFD1RS1("RELATION") <> "STAFF" Then
            txt_status = "DP"
        Else
            txt_status = "ST"
        End If
    End If
'      SELEMPBASINFO_QR(0) = MaskEdBox1.Text
'      SELEMPBASINFO_RS.MoreResults
'      SELEMPBASINFO_RS.Requery
'      If SELEMPBASINFO_RS.RowCount > 0 Then
'        txt_sex.Text = SELEMPBASINFO_RS("GENDER")
'        txt_name.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
'        txt_age = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), CStr(DATE11))
'
'      End If

    
    If REFDRFLAG = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = REQNO
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = REQNO
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If

    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
  
        If selma104flag = False Then
            selma104_qr(0) = MaskEdBox1.FormattedText
            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
            selma104flag = True
        Else
            selma104_qr(0) = MaskEdBox1.FormattedText
            selma104_rs.MoreResults
            selma104_rs.Requery
        End If
  
  
    If selma104_rs.RowCount > 0 Then
        txt_ward.Text = selma104_rs("WARD_NAME")
        txt_bedno.Text = selma104_rs("bed_no")
        txt_status.Text = "IP"
    Else
        txt_status.Text = "OP"
    End If
  
 
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = REQNO
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = REQNO
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
    If selpathdrequisitn_rs.RowCount > 0 Then
        TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
        If IsNull(Trim(selpathdrequisitn_rs("SITE"))) Then
            txt_sites.Text = ""
        Else
            If SELMA301FLAG = False Then
                SELMA301QR(0) = selpathdrequisitn_rs("SITE")
                Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                SELMA301FLAG = True
            Else
                SELMA301QR(0) = selpathdrequisitn_rs("SITE")
                Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                SELMA301RS.MoreResults
                SELMA301RS.Requery
        
            End If
            If SELMA301RS.RowCount > 0 Then
                txt_sites.Text = SELMA301RS("SITEDESC")
            End If
        End If
        If IsNull(Trim(selpathdrequisitn_rs("LATERALITS"))) Then
            txt_laterality.Text = ""
        Else
            txt_laterality.Text = selpathdrequisitn_rs("LATERALITS")
        End If
    
        If Trim(selpathdrequisitn_rs("FROZENSEC")) = "NO" Then
'           txt_frosecrep.Enabled = False
txt_frosecrep.Enabled = True
        End If
        txt_natmeterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        txt_opprocedure.Text = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
        
   
        If IsNull(selpathdrequisitn_rs("IMMHISCHE")) Then
            txt_immhischerep.Enabled = False
            txt_immhischereqno.Enabled = False
            
        End If
    End If
    If SELTA1012FLAG = False Then
            selta1012_qr(0) = MaskEdBox1.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = MaskEdBox1.FormattedText
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
    If selta1012_rs.RowCount > 0 Then
        txt_procedure.Text = selta1012_rs("prvdiagdes")
    End If
'    If SELDREQUISITNFLAG = False Then
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'            SELDREQUISITNFLAG = True
'        Else
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            seldrequisitn_rs.MoreResults
'            seldrequisitn_rs.Requery
'        End If
'    If seldrequisitn_rs.RowCount > 0 Then
'     prpathno = ""
'     While Not seldrequisitn_rs.EOF
'     prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
'     seldrequisitn_rs.MoveNext
'     Wend
'     txt_prepathreqno.Text = prpathno
'    End If
    seldrequisitn_qr10.QueryTimeout = 300
        If seldrequisitn_flag10 = False Then
            seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
            seldrequisitn_flag10 = True
        Else
            seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            seldrequisitn_rs10.MoreResults
            seldrequisitn_rs10.Requery
        End If
        If seldrequisitn_rs10.RowCount > 0 Then
            prpathno = ""
            CMB_PATHNO.clear
                While Not seldrequisitn_rs10.EOF
                    'MODALITY1 = Mid(seldrequisitn_rs10("REQNO"), 4, 2)
                    'Commented and added by Shraddha on 22/11/10
                    MODALITY1 = Mid(seldrequisitn_rs10("REQNNO"), 4, 2)
                    If Trim(MODALITY1) = "SP" Then
                        MODDESC = "SURGICAL PATHOLOGY"
                    ElseIf Trim(MODALITY1) = "CN" Then
                        MODDESC = "CYTOLOGY NON GYNAEC"
                    ElseIf Trim(MODALITY1) = "CG" Then
                        MODDESC = "CYTOLOGY GYNAEC"
                    ElseIf Trim(MODALITY1) = "FA" Then
                        MODDESC = "FNAC"
                    ElseIf Trim(MODALITY1) = "CT" Then
                        MODDESC = "CT SCAN"
                    ElseIf Trim(MODALITY1) = "CR" Then
                        MODDESC = "CONVENTIONAL RADIOGRAPHY"
                    ElseIf Trim(MODALITY1) = "MM" Then
                        MODDESC = "MAMMOGRAPHY"
                    ElseIf Trim(MODALITY1) = "MR" Then
                        MODDESC = "MRI SCAN"
                    ElseIf Trim(MODALITY1) = "US" Then
                        MODDESC = "ULTRASONOGRAPHY"
                    End If
                    prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
                    'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & seldrequisitn_rs10("PREPORTDT") & " " & "$" & " " & seldrequisitn_rs10("REQNO") & " " & "$" & " " & MODDESC
                     CMB_PATHNO.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
                    
                    seldrequisitn_rs10.MoveNext
                Wend
                 CMB_PATHNO.Text = CMB_PATHNO.List(0)
                txt_prepathreqno.Text = prpathno
        End If

    txt_addendum.Enabled = False
    cmb_addendumby.Enabled = False
'        If seldissecentrflag = False Then
'            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
'            seldissecentrflag = True
'        Else
'            seldissecentr_rs.MoreResults
'            seldissecentr_rs.Requery
'        End If
'        If seldissecentr_rs.RowCount > 0 Then
'            cmb_enteryby.clear
'            While Not seldissecentr_rs.EOF
'                cmb_enteryby.AddItem Trim(seldissecentr_rs("CCNAME"))
'                seldissecentr_rs.MoveNext
'            Wend
'        End If
'        If seldissecprovflag = False Then
'            Set seldissecprov_rs = seldissecprov_qr.OpenResultset(1, 3)
'            seldissecprovflag = True
'        Else
'            seldissecprov_rs.MoreResults
'            seldissecprov_rs.Requery
'        End If
'        If seldissecprov_rs.RowCount > 0 Then
'            cmb_provby.clear
'            While Not seldissecprov_rs.EOF
'                cmb_provby.AddItem Trim(seldissecprov_rs("CCNAME"))
'                seldissecprov_rs.MoveNext
'            Wend
'        End If
'        If seldissecfinflag = False Then
'            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
'            seldissecfinflag = True
'        Else
'            seldissecfin_rs.MoreResults
'            seldissecfin_rs.Requery
'        End If
'        If seldissecfin_rs.RowCount > 0 Then
'            cmb_finalby.clear
'            cmb_reg.clear
'            While Not seldissecfin_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
'                cmb_reg.AddItem Trim(seldissecfin_rs("ccname"))
'                seldissecfin_rs.MoveNext
'            Wend
'        End If
'        If seldissecgrossflag = False Then
'            Set seldissecgross_rs = seldissecgross_qr.OpenResultset(1, 3)
'            seldissecgrossflag = True
'        Else
'            seldissecgross_rs.MoreResults
'            seldissecgross_rs.Requery
'        End If
'        If seldissecgross_rs.RowCount > 0 Then
'            cmb_grossby.clear
'            While Not seldissecgross_rs.EOF
'                cmb_grossby.AddItem Trim(seldissecgross_rs("CCNAME"))
'                seldissecgross_rs.MoveNext
'            Wend
'        End If
    
    
    

End If
If viewflag = True Then

   
    
    viewflag = False
    If SELDREQSTAFFD1FLAG1 = False Then
        SELDREQSTAFFD1QR1(0) = Trim(REQNO)
        SELDREQSTAFFD1QR1(1) = Trim(MaskEdBox1.Text)
        Set SELDREQSTAFFD1RS1 = SELDREQSTAFFD1QR1.OpenResultset(1, 3)
    Else
        SELDREQSTAFFD1QR1(0) = Trim(REQNO)
        SELDREQSTAFFD1QR1(1) = Trim(MaskEdBox1.Text)
        SELDREQSTAFFD1RS1.MoreResults
        SELDREQSTAFFD1RS1.Requery
    End If
    If SELDREQSTAFFD1RS1.RowCount > 0 Then
        txt_name = SELDREQSTAFFD1RS1("PNAME")
        txt_sex = SELDREQSTAFFD1RS1("SEX")
        txt_age = SELDREQSTAFFD1RS1("AGE")
        If Trim(SELDREQSTAFFD1RS1("RELATION")) <> "SELF" Then
            txt_status = "DP"
        Else
            txt_status = "ST"
        End If
    End If
''        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
''      SELEMPBASINFO_RS.MoreResults
''      SELEMPBASINFO_RS.Requery
''      If SELEMPBASINFO_RS.RowCount > 0 Then
''        txt_sex.Text = SELEMPBASINFO_RS("GENDER")
''        txt_name.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
''        txt_age = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), CStr(DATE11))
''
''      End If

    If REFDRFLAG = False Then
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = Trim(REQNO)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = Trim(REQNO)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If

      
    If REFDR_RS.RowCount > 0 Then
        '***************************************
  
            If REFDR_RS("RECORDNO") = 1 Then
                LBL_VPP.Caption = "VPP"
            End If
            If REFDR_RS("RECORDNO") = 2 Then
                LBL_VPP.Caption = "LR"
            End If
            If REFDR_RS("RECORDNO") = 3 Then
                LBL_VPP.Caption = "BAR"
            End If
            If REFDR_RS("RECORDNO") = 4 Then
                LBL_VPP.Caption = "DrOp"
            End If
        
      '***************************************
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
        If REFDR_RS("COMMITFLAG") = "Y" Then
            MsgBox "Final Report is Generated, You cann't Modify the Final Report,You can Add ADDENDUMS Only"
            cmd_modify.Enabled = False
            ADDENDUMFALG = True
        End If
    End If
  
    
  
'    If selma104flag = False Then
'            selma104_qr(0) = MaskEdBox1.FormattedText
'            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'            selma104flag = True
'        Else
'            selma104_qr(0) = MaskEdBox1.FormattedText
'            selma104_rs.MoreResults
'            selma104_rs.Requery
'        End If
  
  
'    If selma104_rs.RowCount > 0 Then
'        txt_ward.Text = selma104_rs("WARD_NAME")
'        txt_bedno.Text = selma104_rs("bed_no")
'        txt_status.Text = "IP"
'    Else
'        txt_status.Text = "OP"
'    End If
  
 
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = Trim(REQNO)
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = Trim(REQNO)
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
    If selpathdrequisitn_rs.RowCount > 0 Then
        TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
        If IsNull(selpathdrequisitn_rs("SITE")) Then
        Else
        ' COMMENTED AS VALUE HAS TO BE TAKEN FROM SITEDETAIL RATHER THAN FROM MA301
        ' 04-03-2004
''''            If SELMA301FLAG = False Then
''''                SELMA301QR(0) = Trim(selpathdrequisitn_rs("SITE"))
''''                Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''''                SELMA301FLAG = True
''''            Else
''''                SELMA301QR(0) = Trim(selpathdrequisitn_rs("SITE"))
''''                Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''''                SELMA301RS.MoreResults
''''                SELMA301RS.Requery
''''            End If
''''            If SELMA301RS.RowCount > 0 Then
''''                txt_sites.Text = Trim(SELMA301RS("SITEDESC"))
''''            End If
' COMMENTED AS VALUE HAS TO BE TAKEN FROM SITEDETAIL RATHER THAN FROM MA301
        ' 04-03-2004

            '****************************************
            If SELSITEDETAIL_FLAG1 = False Then
                SELSITEDETAIL_QR1(0) = REQNO 'req_chglen(cmb_reqno.Text)
                Set SELSITEDETAIL_RS1 = SELSITEDETAIL_QR1.OpenResultset(1, 3)
                SELSITEDETAIL_FLAG1 = True
            Else
                SELSITEDETAIL_QR1(0) = REQNO 'req_chglen(cmb_reqno.Text)
                SELSITEDETAIL_RS1.MoreResults
                SELSITEDETAIL_RS1.Requery
            End If
            If Not SELSITEDETAIL_RS1.EOF Then
                cmb_sites.Text = "" & Trim(SELSITEDETAIL_RS1("SDESC"))
                cmb_laterality.Text = "" & Trim(SELSITEDETAIL_RS1("LATERALITY"))
                TXT_FSDIS.Text = "" & Trim(SELSITEDETAIL_RS1("RESDOC"))
            End If
            '****************************************
        End If
        ' COMMENTED AS VALUE HAS TO BE TAKEN FROM SITEDETAIL RATHER THAN FROM MA301
        ' 04-03-2004

''''''        If IsNull(selpathdrequisitn_rs("LATERALITS")) Then
''''''        Else
''''''        txt_laterality.Text = selpathdrequisitn_rs("LATERALITS")
''''''        End If
        ' COMMENTED AS VALUE HAS TO BE TAKEN FROM SITEDETAIL RATHER THAN FROM MA301
        ' 04-03-2004

        txt_natmeterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        txt_opprocedure.Text = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
    End If

    If SELTA1012FLAG = False Then
            selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
    If selta1012_rs.RowCount > 0 Then
        txt_procedure.Text = selta1012_rs("prvdiagdes")
    End If

    If seldfreportflag = False Then
            SELDFREPORT_QR(0) = REQNO
            Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
            seldfreportflag = True
        Else
            SELDFREPORT_QR(0) = REQNO
            SELDFREPORT_RS.MoreResults
            SELDFREPORT_RS.Requery
        End If
    
    If SELDFREPORT_RS.RowCount > 0 Then
        If SELDFREPORT_RS.EOF = False Then
              txt_frosecrep.Text = SELDFREPORT_RS("FSPREPORT").Value
         End If
     Else
        txt_frosecrep.Text = ""
     End If
         
    
       If SELDFREPORT_RS.RowCount > 0 Then
            If SELDFREPORT_RS.EOF = False Then
                txt_immhischerep.Text = SELDFREPORT_RS("ihreport").Value
                txt_immhischerep.Visible = True
                cmb_immhische.Text = "YES"
                
                
            End If
        Else
            txt_immhischerep.Text = ""
            cmb_immhische.Text = "NO"
            txt_immhischerep.Visible = False
            
        End If
        
         If SELDFREPORT_RS.RowCount > 0 Then
            If SELDFREPORT_RS.EOF = False Then
                 txt_impfindiag.Text = SELDFREPORT_RS("impression").Value
            End If
         Else
            txt_impfindiag.Text = ""
         End If
          If SELDFREPORT_RS.RowCount > 0 Then
            If SELDFREPORT_RS.EOF = False Then
                txt_grossdesc = " " & SELDFREPORT_RS("GROSSDESC").Value
             End If
         Else
            txt_grossdesc = ""
         End If

         
          If SELDFREPORT_RS.RowCount > 0 Then
            If SELDFREPORT_RS.EOF = False Then
               txt_microdesc = Trim(SELDFREPORT_RS("SPREPORT").Value)
             Else
                'txt_microdesc = SELDFREPORT_RS("SPREPORT").Value
            End If
         Else
            txt_microdesc = ""
         End If
         If SELDFREPORT_RS.EOF = False Then
            cmb_grossby.Text = SELDFREPORT_RS("GROSSEDBY")
         Else
            cmb_grossby.Text = ""
         End If
         
        If SELDFREPORT_RS.EOF = False Then
            txt_fnacreport.Text = "" & Trim(SELDFREPORT_RS("FNAC"))
        Else
            txt_fnacreport.Text = ""
        End If
''    End If
'    seldfsreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
'    seldfsreport_rs.MoreResults
'    seldfsreport_rs.Requery
'    If seldfsreport_rs.RowCount > 0 Then
'        txt_frosecrep.Text = seldfsreport_rs("fsreport")
'    End If

'    seldihreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
'    seldihreport_rs.MoreResults
'    seldihreport_rs.Requery
'    If seldihreport_rs.RowCount > 0 Then
'        txt_immhischerep.Text = seldihreport_rs("ihreport")
'    End If

    If REFDR_RS.RowCount > 0 Then
        cmb_enteryby.Text = Trim(REFDR_RS("ENTERBY"))
        cmb_provby.Text = Trim(REFDR_RS("PREPORTBY"))
        If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
            cmb_finalby.Text = ""
        Else
            cmb_finalby.Text = Trim(REFDR_RS("FREPORTBY"))
        End If
    End If

    If SELDIMPRESIONFLAG = False Then
            seldimpresion_qr(0) = REQNO
            Set seldimpresion_rs = seldimpresion_qr.OpenResultset(1, 3)
            
            SELDIMPRESIONFLAG = True
        Else
            seldimpresion_qr(0) = REQNO
            seldimpresion_rs.MoreResults
            seldimpresion_rs.Requery
        End If
    If seldimpresion_rs.RowCount > 0 Then
           ' txt_impfindiag.Text = seldimpresion_rs("impression")
            cmb_fsd.AddItem seldimpresion_rs("FDESCRY")
            cmb_fsd.Text = seldimpresion_rs("FDESCRY")
            cmb_t.Text = seldimpresion_rs("T")
            txt_size.Text = seldimpresion_rs("TSIZE")
            CMB_N.Text = seldimpresion_rs("N")
            cmb_m.Text = seldimpresion_rs("M")
            CMB_ICDH.AddItem seldimpresion_rs("FINALDIAG")
            CMB_ICDH.Text = seldimpresion_rs("FINALDIAG")
    End If

'    seldgrossdesc_qr(0) = req_chglen(Trim(cmb_reqno.Text))
'    seldgrossdesc_rs.MoreResults
'    seldgrossdesc_rs.Requery
'    If seldgrossdesc_rs.RowCount > 0 Then
'            txt_grossdesc = seldgrossdesc_rs("GROSSDESC")
'            cmb_grossby.Text = seldgrossdesc_rs("GROSSEDBY")
'    End If


'    seldspreport_qr(0) = req_chglen(Trim(cmb_reqno.Text))
'    seldspreport_rs.MoreResults
'    seldspreport_rs.Requery
'    If seldspreport_rs.RowCount > 0 Then
'            txt_microdesc = seldspreport_rs("SPREPORT")
'            selempbasinfo1_qr(0) = seldspreport_rs("PREPORTBY")
'            selempbasinfo1_rs.MoreResults
'            selempbasinfo1_rs.Requery
'        If selempbasinfo1_rs.RowCount > 0 Then
'            cmb_provby.Clear
'            cmb_provby.AddItem selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("PREPORTBY")
'            cmb_provby.Text = selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("PREPORTBY")
'        End If
'        selempbasinfo1_qr(0) = seldspreport_rs("FREPORTBY")
'        selempbasinfo1_rs.MoreResults
'        selempbasinfo1_rs.Requery
'        If selempbasinfo1_rs.RowCount > 0 Then
'            cmb_finalby.Clear
'            cmb_finalby.AddItem selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("FREPORTBY")
'            cmb_finalby.Text = selempbasinfo1_rs("EMPTITLE") & "$" & seldspreport_rs("FREPORTBY")
'        End If
        
'        If SELDREQUISITNFLAG = False Then
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'            SELDREQUISITNFLAG = True
'        Else
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            seldrequisitn_rs.MoreResults
'            seldrequisitn_rs.Requery
'        End If
'    If seldrequisitn_rs.RowCount > 0 Then
'     prpathno = ""
'     While Not seldrequisitn_rs.EOF
'        prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'        CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
'        seldrequisitn_rs.MoveNext
'     Wend
'        txt_prepathreqno.Text = Trim(prpathno)
'    End If

    seldrequisitn_qr10.QueryTimeout = 300
        If seldrequisitn_flag10 = False Then
            seldrequisitn_qr10(0) = Trim(MaskEdBox1.FormattedText)
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
            seldrequisitn_flag10 = True
        Else
            seldrequisitn_qr10(0) = Trim(MaskEdBox1.FormattedText)
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            seldrequisitn_rs10.MoreResults
            seldrequisitn_rs10.Requery
        End If
        If seldrequisitn_rs10.RowCount > 0 Then
            prpathno = ""
            CMB_PATHNO.clear
                While Not seldrequisitn_rs10.EOF
                    prpathno = prpathno & " " & Trim(seldrequisitn_rs10("pathno"))
                    'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & seldrequisitn_rs10("REQNO") 'without date sun
                    CMB_PATHNO.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
                    seldrequisitn_rs10.MoveNext
                Wend
                CMB_PATHNO.Text = CMB_PATHNO.List(0)
                txt_prepathreqno.Text = prpathno
        End If
        SELDADDENDUM_QR(0) = REQNO
        SELDADDENDUM_RS.MoreResults
        SELDADDENDUM_RS.Requery
        If SELDADDENDUM_RS.RowCount > 0 Then
            While Not SELDADDENDUM_RS.EOF
                ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
            SELDADDENDUM_RS.MoveNext
            Wend
            txt_addendum.Text = Trim(ADDENDUM)
            
        End If
    'End If
    cmb_reqno.Enabled = False
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub txt_natmeterial_GotFocus()
On Error GoTo m1:
w_meterial.Visible = False
Set d_frozen = w_frozen.ActiveDocument
d_meterial.Content = ""
d_meterial.Content = txt_natmeterial
w_meterial.Caption = "Nature of Meterial"
'w_meterial.WindowState = wdWindowStateMaximize
w_meterial.WindowState = wdWindowStateNormal
d_meterial.Content = txt_natmeterial
w_meterial.Visible = True
w_meterial.Activate
m1:
On Error Resume Next
Set w_meterial = CreateObject("Word.Application")
w_meterial.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_meterial.Caption = "Nature of Meterial"
'w_meterial.WindowState = wdWindowStateMaximize
w_meterial.WindowState = wdWindowStateNormal
Set d_meterial = w_meterial.ActiveDocument
d_meterial.Content = Trim(txt_natmeterial.Text)
w_meterial.Visible = True
w_meterial.Activate
End Sub

Private Sub txt_natmeterial_LostFocus()
On Error GoTo m2:
Set d_meterial = w_meterial.ActiveDocument
txt_natmeterial.Text = d_meterial.Content
If Len(Trim(txt_natmeterial.Text)) > 1000 Then
    MsgBox "Nature of Meterial Max. of 1000 Characters"
    txt_natmeterial.SetFocus
    Exit Sub
End If
'flagimp = False
d_meterial.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_meterial.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_meterial.Close (wdDoNotSaveChanges)
w_meterial.quit
'W1.Visible = False
If SELPATHDREQUISITNFLAG = False Then
    SELPATHDREQUISITN_QR(0) = REQNO
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
    SELPATHDREQUISITNFLAG = True
Else
    SELPATHDREQUISITN_QR(0) = REQNO
    selpathdrequisitn_rs.MoreResults
    selpathdrequisitn_rs.Requery
End If

If Trim(selpathdrequisitn_rs("FROZENSEC")) = "NO" Or Trim(selpathdrequisitn_rs("FROZENSEC")) = "" Then
   'txt_frosecrep.Enabled = False
   txt_frosecrep.Enabled = True
End If

Exit Sub
m2:
MsgBox Err.Description
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
    SELRREQNOQR(0) = Trim(REQNO1) 'VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = Trim(REQNO1) 'VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 150
If seldrequisitn_flag11 = False Then
    seldrequisitn_qr11(0) = MaskEdBox1.FormattedText
    Set seldrequisitn_rs11 = seldrequisitn_qr11.OpenResultset(1, 3)
    seldrequisitn_flag11 = True
Else
    seldrequisitn_qr11(0) = MaskEdBox1.FormattedText
    seldrequisitn_rs11.MoreResults
    seldrequisitn_rs11.Requery
End If
If seldrequisitn_rs11.RowCount > 0 Then
    prpathno = ""
    'CMB_PATHNO.clear
    While Not seldrequisitn_rs11.EOF
        prpathno = prpathno & " " & Trim(seldrequisitn_rs11("pathno"))
        'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
        seldrequisitn_rs11.MoveNext
    Wend
    'txt_prepathreqno.Text = prpathno
End If
SELDADDENDUM_QR(0) = Trim(REQNO1) 'VAR_REQNNO
SELDADDENDUM_RS.MoreResults
SELDADDENDUM_RS.Requery

For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i


str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Trim(REQNO1) & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"
If Opt_patient.Value = True Then
        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
                'Commented and added on 20/09/10
                CrystalReport1.SelectionFormula = "{FROZENADLF.REQNNO} = '" & Trim(REQNO1) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\FROZENADD.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = Trim(REQNO1) 'VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = 'Prev. Path NO. " & Trim(prpathno) & "'"
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
        SELDFREPORT_QR(0) = Trim(REQNO1) 'VAR_REQNNO
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
        'CRYSTALREPORT1.WindowState = crptMaximized
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
    SELRREQNOQR(0) = Trim(REQNO1) '   VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = Trim(REQNO1) ' VAR_REQNNO
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
SELDADDENDUM_QR(0) = Trim(REQNO1)   '          VAR_REQNNO
SELDADDENDUM_RS.MoreResults
SELDADDENDUM_RS.Requery

For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Trim(REQNO1) & "'"
'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"

If Opt_patient.Value = True Then
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
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
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
VAR_REQNNO = REQNO1
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = Trim(REQNO1) 'req_chglen(Trim(REQNO1))
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = Trim(REQNO1) 'req_chglen(Trim(REQNO1))
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
SELDADDENDUM_QR(0) = Trim(REQNO1)   'req_chglen(Trim(VAR_REQNNO))
SELDADDENDUM_RS.MoreResults
SELDADDENDUM_RS.Requery

For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Trim(REQNO1) & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"

If Opt_patient.Value = True Then
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
        CrystalReport1.Formulas(12) = "NOSLIDES=" & "" & selpathdrequisitn_rs("LABNO")
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
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"
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
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"
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
Private Sub RISREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = Trim(REQNO1)  '  VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = Trim(REQNO1)   'VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'rq_grid_det(0) = Trim(VAR_REQNNO)
'rs_grid_det.MoreResults
'rs_grid_det.Requery

For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i

If SELRREQNORS.RowCount > 0 And (Trim(SELRREQNORS("REQNSTAT")) = "T" Or Trim(SELRREQNORS("REQNSTAT")) = "F") Then
    If Opt_patient.Value = True Then
        CrystalReport1.Formulas(0) = "provtyp= '" & Trim(SELRREQNORS("ENTERBY")) & "'"
        'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
        'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
        'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
    
        SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
        SELRMODALITY_RS.MoreResults
        SELRMODALITY_RS.Requery
        If SELRMODALITY_RS.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
        Else
            CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
        End If
        CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
        CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
        str1 = ip
        str2 = DATE11
        STR3 = time1
        str1 = str1 & Space(5) & str2 & Space(5) & STR3
        str1 = str1
        CrystalReport1.Formulas(3) = "us_id = '" & str1 & "'"
        CrystalReport1.Formulas(4) = "INOUT = '" & Trim(SELRREQNORS("PATNSTAT")) & "'"
        CrystalReport1.Formulas(5) = "reqno = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Or Trim(SELRREQNORS("REQNSTAT")) = "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1.rpt"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2.rpt"
        End If
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
        
        CrystalReport1.Formulas(0) = "provtyp= '" & Trim(SELRREQNORS("ENTERBY")) & "'"
        SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
        SELRMODALITY_RS.MoreResults
        SELRMODALITY_RS.Requery
        If SELRMODALITY_RS.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
        Else
            CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
        End If
        CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
        CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & Trim(REQNO1) & "' AND {RREQUISITN.REQNNO} = '" & Trim(REQNO1) & "'"
        str1 = ip
        str2 = DATE11
        STR3 = time1
        str1 = str1 & Space(5) & str2 & Space(5) & STR3
        str1 = Space(3) & str1
        CrystalReport1.Formulas(3) = "us_id = '" & str1 & "'"
        CrystalReport1.Formulas(4) = "INOUT = '" & Trim(SELRREQNORS("PATNSTAT")) & "'"
        CrystalReport1.Formulas(5) = "reqno = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(6) = "PATNAME = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(7) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(5) = "PATAGE = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Or Trim(SELRREQNORS("REQNSTAT")) = "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1STAFF.rpt"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2STAFF.rpt"
        End If
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
    Exit Sub
Else
    Exit Sub
End If
End Sub
Private Sub BTNENABL1(cap As String)

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
    cmd_fullreport.Enabled = False
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
    cmd_fullreport.Enabled = True
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
    cmd_fullreport.Enabled = False
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
    cmd_fullreport.Enabled = False
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
Public Sub FLDENABL1(str1 As String)
If str1 = "ENA" Then
    cmb_reg.Enabled = True
    cmb_sites.Enabled = True
    cmb_laterality.Enabled = True
    txt_frosecrep.Enabled = True
    txt_grossdesc.Enabled = True
    txt_microdesc.Enabled = True
    txt_immhischerep.Enabled = True
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = True
    cmb_immhische.Enabled = True
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
    Else
        txt_addendum.Enabled = False
    End If
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
    cmb_grade.Enabled = True
    CMB_PATHNO.Enabled = True
    txt_natmeterial.Enabled = True
    'CMB_FNAC.Enabled = True
    txt_fnacreport.Enabled = False
End If
If str1 = "DIS" Then
    cmb_reg.Enabled = False
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
    cmb_grade.Enabled = False
    'CMB_PATHNO.Enabled = False
    txt_natmeterial.Enabled = False
    'CMB_FNAC.Enabled = False
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = False
    cmb_laterality.Enabled = False
    cmb_reg.Enabled = False
End If
If str1 = "VIEW" Then
    cmb_reg.Enabled = False
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
    cmb_grade.Enabled = False
    'CMB_PATHNO.Enabled = False
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_natmeterial.Enabled = False
    'CMB_FNAC.Enabled = False
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = False
    cmb_laterality.Enabled = False
End If
If str1 = "MODIFY" Then
    cmb_reg.Enabled = True
    txt_frosecrep.Enabled = True
    txt_grossdesc.Enabled = True
    txt_microdesc.Enabled = True
    txt_immhischerep.Enabled = True
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = True
    cmb_immhische.Enabled = True
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
    cmb_grade.Enabled = True
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = True
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    CMB_PATHNO.Enabled = True
    txt_natmeterial.Enabled = True
    'CMB_FNAC.Enabled = True
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = True
    cmb_laterality.Enabled = True
End If
If str1 = "ADDENDUM" Then
    cmb_reg.Enabled = False
    txt_frosecrep.Enabled = False
    txt_grossdesc.Enabled = False
    txt_microdesc.Enabled = False
    txt_immhischerep.Enabled = False
    txt_immhischereqno.Enabled = False
    txt_impfindiag.Enabled = False
    cmb_immhische.Enabled = False
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
        cmb_addendumby.Enabled = True
    Else
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
    txt_prepathreqno.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_t.Enabled = False
    cmb_m.Enabled = False
    CMB_N.Enabled = False
    txt_size.Enabled = False
    CMB_ICDH.Enabled = False
    cmb_grossby.Enabled = False
    'CMB_PATHNO.Enabled = False
    cmb_reportdoc.Enabled = False
    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    cmb_grade.Enabled = False
    txt_natmeterial.Enabled = False
    'CMB_FNAC.Enabled = False
    txt_fnacreport.Enabled = False
    cmb_sites.Enabled = False
    cmb_laterality.Enabled = False
End If
End Sub

Public Sub MAKEFALSE()
    
    SELTA1012FLAG = False
    SELRREQUISITNFLAG = False
    FSFLAG = False
    SEL_DEPTFLAG = False
    SEL_MODALITYFLAG = False
    REQFLAG = False
    SELTA101FLAG = False
    REPDRFLAG = False
    REFDRFLAG = False
    selma104flag = False
    SELPATHDREQUISITNFLAG = False
    SELMA302FLAG = False
    SELDFSREPORTFLAG = False
    seldihreportflag = False
    SELDIMPRESIONFLAG = False
    SELDGROSSDESCFLAG = False
    SELDSPREPORTFLAG = False
    SELDADDENDUMFLAG = False
    SELDREQUISITNFLAG = False
    NEWFLAG = False
    viewflag = False
    ADDENDUMFALG = False
    SELMA301FLAG = False
    SELRREQNOFLAG = False
    seldfreportflag = False
    SELDFREPORT_IHFLAG = False
    SELDGRADEFLAG = False
    SELDGRADE_FLAG = False
    SELDREQNOFLAG = False
    SELIHREPORTFLAG = False
    seldissecprovflag = False
    seldissecaddrflag = False
    seldissecgrossflag = False
    SELDREQSTAFFD1FLAG1 = False
    VIEWSITEDUPRSFLAG = False
    SELDSPFNFLAG = False
    SELDFNREPORTFLAG = False
    SELDSPFNFLAG = False
    SELDFNREPORTFLAG = False
    SELSITEDETAIL_FLAG = False
    SELSITEDETAIL_FLAG1 = False
    seldrequisitn_flag10 = False
    seldrequisitn_flag11 = False
    SELSITE_FLAG = False

End Sub


