VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form cytology_report 
   Caption         =   "Cytology"
   ClientHeight    =   10830
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   14940
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame COMMAND 
      Height          =   795
      Left            =   0
      TabIndex        =   53
      Top             =   7800
      Width           =   11775
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
         MouseIcon       =   "cytology_report.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   0
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
         Left            =   1680
         MouseIcon       =   "cytology_report.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   56
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
         Left            =   3000
         MouseIcon       =   "cytology_report.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   55
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
         Left            =   4320
         MouseIcon       =   "cytology_report.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   14
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
         Left            =   5640
         MouseIcon       =   "cytology_report.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   54
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
         Left            =   8400
         MouseIcon       =   "cytology_report.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Width           =   1095
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
         Left            =   9720
         MouseIcon       =   "cytology_report.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   20
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
         Left            =   6960
         MouseIcon       =   "cytology_report.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10680
      TabIndex        =   52
      Top             =   360
      Width           =   1095
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10680
      TabIndex        =   51
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame common 
      Height          =   3015
      Left            =   0
      TabIndex        =   17
      Top             =   960
      Width           =   11775
      Begin MSMask.MaskEdBox msk_coll_tm 
         Height          =   255
         Left            =   9960
         TabIndex        =   86
         Top             =   2640
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_coll_dt 
         Height          =   375
         Left            =   9960
         TabIndex        =   85
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.CommandButton CMD_EMR 
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
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   2160
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   70
         Top             =   600
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
            Left            =   240
            TabIndex        =   72
            Top             =   600
            Width           =   975
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
            Left            =   240
            TabIndex        =   71
            Top             =   240
            Width           =   975
         End
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
         TabIndex        =   1
         Top             =   240
         Width           =   2895
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
         TabIndex        =   33
         Top             =   240
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
         Left            =   6225
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txt_cytology 
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
         TabIndex        =   32
         Top             =   600
         Width           =   1335
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
         TabIndex        =   31
         Top             =   1080
         Width           =   9135
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
         TabIndex        =   30
         Top             =   1440
         Width           =   855
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
         TabIndex        =   29
         Top             =   1440
         Width           =   735
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
         TabIndex        =   28
         Top             =   1440
         Width           =   975
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
         TabIndex        =   27
         Top             =   1440
         Width           =   1095
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   26
         Top             =   1800
         Width           =   1575
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
         TabIndex        =   25
         Top             =   1800
         Width           =   1335
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
         TabIndex        =   24
         Top             =   1800
         Width           =   975
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
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
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
         TabIndex        =   22
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_clndiag 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   2520
         Width           =   3855
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
         ItemData        =   "cytology_report.frx":1850
         Left            =   6240
         List            =   "cytology_report.frx":1852
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
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
      Begin VB.Label Label5 
         Caption         =   "24 Hrs"
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
         Left            =   10800
         TabIndex        =   87
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lbl_colltm 
         Caption         =   "Collection Tm : "
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
         Left            =   8640
         TabIndex        =   84
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lbl_colldt 
         Caption         =   "Collection Dt : "
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
         Left            =   8640
         TabIndex        =   83
         Top             =   2280
         Width           =   1335
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
         TabIndex        =   50
         Top             =   360
         Width           =   1830
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
         TabIndex        =   49
         Top             =   240
         Width           =   975
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
         Left            =   1560
         TabIndex        =   48
         Top             =   720
         Width           =   795
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
         TabIndex        =   47
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_cytologyno 
         AutoSize        =   -1  'True
         Caption         =   "Cytology No."
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
         TabIndex        =   46
         Top             =   600
         Width           =   1095
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
         TabIndex        =   45
         Top             =   1080
         Width           =   495
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
         TabIndex        =   44
         Top             =   1440
         Width           =   330
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
         TabIndex        =   43
         Top             =   1440
         Width           =   345
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
         TabIndex        =   42
         Top             =   1440
         Width           =   765
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
         TabIndex        =   41
         Top             =   1440
         Width           =   555
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
         TabIndex        =   40
         Top             =   1800
         Width           =   1995
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
         TabIndex        =   39
         Top             =   1800
         Width           =   495
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
         TabIndex        =   38
         Top             =   1800
         Width           =   705
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
         TabIndex        =   37
         Top             =   360
         Width           =   720
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
         TabIndex        =   36
         Top             =   1800
         Width           =   1215
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
         TabIndex        =   35
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_clndis 
         AutoSize        =   -1  'True
         Caption         =   "Clinical Diagnosis"
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
         Left            =   480
         TabIndex        =   34
         Top             =   2520
         Width           =   1515
      End
   End
   Begin VB.Frame REPORT 
      Height          =   3825
      Left            =   0
      TabIndex        =   57
      Top             =   4080
      Width           =   11775
      Begin VB.ComboBox CMB_PATHNO 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2400
         TabIndex        =   81
         Top             =   240
         Width           =   6375
      End
      Begin VB.ComboBox cmb_rpt_class 
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
         ItemData        =   "cytology_report.frx":1854
         Left            =   2400
         List            =   "cytology_report.frx":1856
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txt_noslides 
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
         Left            =   10440
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TXT_CSITE 
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
         Left            =   5760
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txt_typespeciman 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   2640
         Width           =   4335
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "cytology_report.frx":1858
         Left            =   8595
         List            =   "cytology_report.frx":185A
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   3120
         Width           =   2895
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3120
         Width           =   2295
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "cytology_report.frx":185C
         Left            =   8640
         List            =   "cytology_report.frx":185E
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2160
         Width           =   2775
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "cytology_report.frx":1860
         Left            =   8595
         List            =   "cytology_report.frx":1862
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txt_report 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   9135
      End
      Begin VB.TextBox txt_impdiag 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1680
         Width           =   9135
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
         TabIndex        =   80
         Top             =   240
         Width           =   2070
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   360
         TabIndex        =   77
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Report Classification"
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
         Height          =   375
         Left            =   480
         TabIndex        =   76
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Height          =   225
         Left            =   9000
         TabIndex        =   75
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "No. of Slides"
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
         Left            =   9240
         TabIndex        =   74
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LBL_VPP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   6000
         TabIndex        =   73
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl_sitespeciman 
         AutoSize        =   -1  'True
         Caption         =   "Site"
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
         TabIndex        =   69
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lbl_typespecimen 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Material"
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
         Left            =   720
         TabIndex        =   68
         Top             =   720
         Width           =   1545
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
         Left            =   1440
         TabIndex        =   67
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Entered  by"
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
         TabIndex        =   66
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label lbl_pro 
         AutoSize        =   -1  'True
         Caption         =   "Provisional By"
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
         Left            =   7305
         TabIndex        =   65
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lbl_final 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized By"
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
         Left            =   6915
         TabIndex        =   64
         Top             =   2640
         Width           =   1590
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
         Left            =   7395
         TabIndex        =   63
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Label lbl_report 
         AutoSize        =   -1  'True
         Caption         =   " Report"
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
         Left            =   1680
         TabIndex        =   59
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lbl_impression 
         AutoSize        =   -1  'True
         Caption         =   "Impression/ Diagnosis"
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
         Top             =   1680
         Width           =   1890
      End
   End
   Begin MSWinsockLib.Winsock WIN_CLIENT 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Height          =   255
      Left            =   1440
      TabIndex        =   79
      Top             =   120
      Width           =   8655
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
      Left            =   120
      TabIndex        =   78
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   10080
      TabIndex        =   62
      Top             =   480
      Width           =   420
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
      Left            =   10080
      TabIndex        =   61
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lbl_cyt 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - CYTOLOGY"
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
      Left            =   4680
      TabIndex        =   60
      Top             =   600
      Width           =   3105
   End
End
Attribute VB_Name = "cytology_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   'Option Explicit
Dim str As String

'SAVE THE DATE AND TIME INTO RTIMELOG
Dim UPDRTIME_QR As rdoQuery
Dim UPD_RS As rdoResultset

Dim seldrequisitn_qrprv As rdoQuery
Dim seldrequisitn_rsprv As rdoResultset
Dim seldrequisitn_flagprv As Boolean
Dim MODALITY1
Dim MODDESC
Dim DEPT1
Dim X
Dim REQNO1

Dim seldrequisitn_qr11 As rdoQuery
Dim seldrequisitn_rs11 As rdoResultset
Dim seldrequisitn_flag11 As Boolean

Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset

Dim UPDRTIMECADD_QR As rdoQuery

Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEC_RS As rdoResultset

Dim UPDRTIMEP_QR As rdoQuery
Dim UPDRTIMEM_QR As rdoQuery

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
'Dim SELTA101FLAG As Boolean
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
Dim WREPORT As Application
Dim DREPORT As Document
Dim WIMPRE As Application
Dim DIMPRE As Document
''MONO ON 15/04/03
Dim d_grossdesc As Document
Dim W_grossdes As Application

Dim WMETERIAL As Application
Dim DMETERIAL As Document

Dim SELEMPLIB_QR As rdoQuery
Dim SELEMPLIB_RS As rdoResultset
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
Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset
Dim seldfreportflag As Boolean
Dim SELDFREPORT_IHQR As rdoQuery
Dim SELDFREPORT_IHRS As rdoResultset
Dim SELDFREPORT_IHFLAG As Boolean
Dim upddrequisitn_qr As rdoQuery

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFOFLAG As Boolean
Dim seldgradeqr As rdoQuery
Dim seldgraders As rdoResultset
Dim SELDGRADEFLAG As Boolean
Dim seldgrade_qr As rdoQuery
Dim seldgrade_rs As rdoResultset
Dim SELDGRADE_FLAG As Boolean
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
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

Dim SELCLIN_QR As rdoQuery
Dim SELCLIN_RS As rdoResultset
Dim SELCLINFLAG As Boolean

Dim INSDISIMP_QR As rdoQuery
Dim UPDDISIMP_QR As rdoQuery

Dim UPDDREQ_QR As rdoQuery

''REPORT CLASSIFICATION

Dim UPDIMMHISCHE_QR As rdoQuery

Dim SELIMMHISCHE_QR As rdoQuery
Dim SELIMMHISCHE_RS As rdoResultset
Dim SELIMMHISCHE_FLAG As Boolean

''REPORT CLASSIFICATION

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

''******SECURITY***********

''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean
Dim VAR_REQNNO As String
''*******ICU**********
Dim NEWF As Boolean
Dim upd_rtime_qr As rdoQuery 'Shraddha 21/03/11
Dim sel_rtimelog_qr As rdoQuery
Dim sel_rtimelog_rs As rdoResultset



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
'            While Not seldissecgross_rs.EOF
'                seldissecgross_rs.MoveNext
'            Wend
'        End If
Exit Sub
MSG:
'MsgBox Error(Err)
End Sub

Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "CYTOLOGY GYNAEC$CG" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If

Exit Sub
MSG:
MsgBox Error(Err)
End Sub



Private Sub CMB_PATHNO_Click()
    'ADDED BY MILIND
    X = InStr(12, Trim(CMB_PATHNO.Text), "$")
    REQNO1 = Mid(Trim(CMB_PATHNO.Text), 1, X - 1)
    DEPT1 = Mid(Trim(REQNO1), 1, 3)
    MODALITY1 = Mid(Trim(REQNO1), 4, 2)
    
    If DEPT1 = "FZZ" And MODALITY1 = "FA" Then
        FAREPORT
    ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
        CGCNREPORT
    End If
    
End Sub


Private Sub cmb_rpt_class_Click()
    If Trim(cmb_rpt_class.Text) <> "" Then
       If Trim(cmb_rpt_class.Text) = "Negative" Then
       'FOLLOWING
       'str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ?  AND DEPT_CODE = 'FZZ'"
       'FOLLOWING QUERY CHANGED BY SHWETA 04-12-2012 'AS REQUEST FROM FZZ DEPT FOR 303762
       'str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ?  AND DEPT_CODE = 'FZZ' AND USR_id NOT LIKE 'A3%'"
       'str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ?  AND DEPT_CODE = 'FZZ' AND (USR_id NOT LIKE 'A3%' )"          ''COMMENT BY RAHUL AT SANGRUR ON 05.04.2016
       str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE usrtype >= 9 or ( USRTYPE >= ?  AND DEPT_CODE = 'FZZ' AND (USR_id NOT LIKE 'A3%' OR MODMR='A'))" '' ADD BY RAHUL AT SANGRUR ON 05.04.2016
       Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1", str)
            cmb_finalby.clear
            SELFINALBY_QR(0) = 8 ' MADE 6 FROM 5
            Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
            If Not SELFINALBY_RS.EOF Then
               Do While Not SELFINALBY_RS.EOF
                cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
                SELFINALBY_RS.MoveNext
               Loop
            End If
            SELFINALBY_QR.Close
        End If
        
        If Trim(cmb_rpt_class.Text) = "Positive" Then
        
        'str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ? AND DEPT_CODE = 'FZZ'"
        'str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ? AND DEPT_CODE = 'FZZ' AND (USR_id NOT LIKE 'A3%' "
        'FOLLOWING QUERY CHANGED BY SHWETA 04-12-2012 'AS REQUEST FROM FZZ DEPT FOR 303762
       'str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ?  AND DEPT_CODE = 'FZZ' AND USR_id NOT LIKE 'A3%'"
       str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE usrtype >= 9 or ( USRTYPE >= ?  AND DEPT_CODE = 'FZZ' AND (USR_id NOT LIKE 'A3%' OR MODMR='A'))"
        Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1", str)
             cmb_finalby.clear
            SELFINALBY_QR(0) = 8 'MADE 8 FROM 5
            Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
            If Not SELFINALBY_RS.EOF Then
                Do While Not SELFINALBY_RS.EOF
                    cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
                    SELFINALBY_RS.MoveNext
                Loop
            End If
            SELFINALBY_QR.Close
        End If
    End If
End Sub

Private Sub cmd_back_Click()
'btnenabl ("INI")
BTNENABL1 ("INI")
FLDENABL ("DIS")
field_clear
TXT_CLEAR
DIS_MAIN.show
Me.Hide
viewflag = False
NEWFLAG = False
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
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
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
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = True
    cmd_view.Enabled = True
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    CMD_NEW.Enabled = True
    cmd_modify.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
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
FLDENABL ("DIS")
NEWF = False
End Sub

Private Sub cmd_commit_Click()

Dim i As Integer

''********SECURITY***************************

'******* add from here for give authority only for usrtype >=10 sangrur 18.07.2017
If ADDENDUMFALG = True Then
If LOGINTYP < 10 Then
 MsgBox "Sorry! You are not authorised to commit the addendum of Report "
 Exit Sub
End If
End If
'******* add till here for give authority only for usrtype >=10 sangrur 18.07.2017

CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
CURUSRDET_RS.MoreResults
CURUSRDET_RS.Requery
If CURUSRDET_RS.RowCount > 0 Then
    If LOGINTYP > CURUSRDET_RS("usrtype") Then
      ''   no action as he can commit
    ElseIf Val(LOGINTYP) < Val(CURUSRDET_RS("usrtype")) Then
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
    MsgBox "Sorry! The Security Details of this user is not available check up in DISLNKUSR"
    cmd_commit.SetFocus
    Exit Sub
End If
If cmb_rpt_class.Text = "Positive" Then
      If CURUSRDET_RS("usrtype") <= 6 Then 'MADE 6 FROM 2
          MsgBox "Postive Reports cannot be commited by Cyto tecnhologist"
         Exit Sub
      End If
End If

''********SECURITY***************************

i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If i = vbYes Then
On Error GoTo MSG:

 rdb.BeginTrans
    updrrequisitn1_qr(0) = Left(cmb_finalby.Text, 30)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
End If
rdb.CommitTrans
'btnenabl ("INI")

    
    'rtimelog
    UPDRTIMEC_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute
BTNENABL1 ("INI")

MsgBox "Record is Committed Successfully."

If ADDENDUMFALG = True Then
sendRMSG req_chglen(cmb_reqno.Text), MaskEdBox1.FormattedText, "A", WIN_CLIENT.LocalIP
Else
sendRMSG req_chglen(cmb_reqno.Text), MaskEdBox1.FormattedText, "F", WIN_CLIENT.LocalIP
End If

cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_fullreport_Click()
On Error GoTo MSG:
txt_fullreport_GotFocus
Set D1 = w.ActiveDocument
If Len(txt_addendum) > 0 Then
D1.Content = D1.Content & "ADDENDUM:" & Chr(13) & Trim(txt_addendum)
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_fullreport_LostFocus()
On Error GoTo MSG:
D1.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
'D1.SaveAs "C:\REPORT.DOC"
D1.Close (wdDoNotSaveChanges)
w.quit
'Msk_caseno.Enabled = False
'cmd_cancel.SetFocus
MSG:
End Sub

Private Sub CMD_EMR_Click()
'If Len(Trim(MaskEdBox1.Text)) > 0 Then
'    Dim VAR1
'    'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
'    VAR1 = MaskEdBox1.FormattedText
'    VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
'    AppActivate VAR
'Else
'    MsgBox "Please Enter The Case No."
'End If
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://10.80.1.9/emrtmhact/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
End Sub

Private Sub cmd_modify_Click()

''********SECURITY***************************
CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
CURUSRDET_RS.MoreResults
CURUSRDET_RS.Requery
If CURUSRDET_RS.RowCount > 0 Then
    If LOGINTYP > CURUSRDET_RS("usrtype") Then
      ''   no action as he can commit
    ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
     '   MsgBox "Sorry! This Report needs is to Modified by a Higher Level than  that of Yours"
     '   cmd_modify.SetFocus
     '   Exit Sub
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
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provby.Enabled = True
    Exit Sub
End If

''********SECURITY***************************
'btnenabl ("NEW")
BTNENABL1 ("NEW")
FLDENABL ("MODIFY")
flagsave = 2
If REFDRFLAG = False Then
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    REFDRFLAG = True
Else
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    REFDR_RS.MoreResults
    REFDR_RS.Requery
End If
If REFDR_RS.RowCount > 0 Then
    'If Trim(REFDR_RS("reqnstat")) = "T" And Trim(REFDR_RS("COMMITFLAG")) = "F" Then
    If Trim(REFDR_RS("reqnstat")) = "F" And Trim(REFDR_RS("COMMITFLAG")) = "Y" Then
        ADDENDUMFALG = True
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
    FLDENABL ("clindiag")
    FLDENABL ("ADDENDUM")
    flagsave = 2
End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
'On Error GoTo MSG:
'btnenabl ("NEW")
BTNENABL1 ("NEW")
field_clear
TXT_CLEAR
NEWF = True
datetime

txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
cmb_enteryby.Text = cmb_enteryby.List(0)
cmb_enteryby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enteryby.List(0)) = cmb_finalby.List(L1) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
Next L1
cmb_provby.Text = cmb_provby.List(0)
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
 
flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
Opt_patient.Value = True
'MaskEdBox1.Enabled = True

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim i As Integer
'On Error GoTo prtmsg:

DoEvents

''' added for print emr report for sangrur by rahul 22/01/2016 as Dr Sankalp wanted to direct link to emr for printing report

VAR_REQNNO = req_chglen(cmb_reqno.Text)
 strLink = "http://10.80.1.9/EMRREPORTS/allreports.aspx?reqno=" & VAR_REQNNO & "  ~GUEST"
 
 
    'strLink = "http://10.80.1.9/EMRTMHACT/DEFAULT.ASPX?QPNO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink, vbMaximizedFocus
Exit Sub
'' till there by rahul 22/01/2015

'If Trim(cmb_rpt_class.Text) = "Pos" Then
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
    
    'If Trim(SELRREQNORS("COMMITFLAG")) <> "Y" Or (IsNull(SELRREQNORS("COMMITFLAG") = True)) Then
    'MsgBox ("Only Commited Records Can be Printed"), vbInformation
    'Exit Sub
    'End If
    'End If
 
VAR_REQNNO = req_chglen(cmb_reqno.Text)

        
        ''*******icu****************
'   If Opt_patient.Value = True Then
'        If SELICU_FLAG = False Then
'            SELICU_QR(0) = VAR_REQNNO
'            Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
'            SELICU_FLAG = True
'        Else
'            SELICU_QR(0) = VAR_REQNNO
'            SELICU_RS.MoreResults
'            SELICU_RS.Requery
'        End If
'        If Not SELICU_RS.EOF Then
'            If SELMEMO_FLAG = False Then
'                SELMEMO_QR(0) = VAR_REQNNO
'                Set SELMEMO_RS = SELMEMO_QR.OpenResultset(1, 3)
'                SELMEMO_FLAG = True
'            Else
'                SELMEMO_QR(0) = VAR_REQNNO
'                SELMEMO_RS.MoreResults
'                SELMEMO_RS.Requery
'            End If
'        End If
'        If Not SELMEMO_RS.EOF Then
'            If IsNull(SELMEMO_RS("MEMONO")) And Trim(SELMEMO_RS("PAT_CATG")) <> "NC" Then
'                MsgBox " Please Prepare the Memo !!!", vbInformation, "DIS"
'                Exit Sub
'            End If
'        End If
'    End If
        ''*******icu****************

If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
For i = 0 To 17
    CrystalReport1.Formulas(i) = ""
Next
'CrystalReport1
 str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
If Opt_patient.Value = True Then
    If ADDENDUMFALG = False Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(TXT_CSITE.Text) = "" Then
            CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
        End If
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(15) = "noslides = '" & Trim(txt_noslides.Text) & "'"
        
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
        ' CrystalReport1.Formulas(15) = "signname= '" & Trim(cmb_provby.Text) & "'"
        CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT WITH ADDENDUM" & "'"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(4) = "noslides = '" & Trim(txt_noslides.Text) & "'"
    End If
End If

If opt_staff.Value = True Then
    If ADDENDUMFALG = False Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(TXT_CSITE.Text) = "" Then
            CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "STL = '" & "Site  " & Trim(TXT_CSITE.Text) & "'"
        End If
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOPATHOLOGY REPORT" & "'" 'Shraddha 28/12/10
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDS.RPT"
        CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT WITH ADDENDUM" & "'"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    End If
End If
If Trim(cmb_provby.Text) <> "" And Trim(cmb_finalby.Text) <> "" And UCase(Trim(cmb_provby.Text)) <> UCase(Trim(cmb_finalby.Text)) Then
 'If Trim(cmb_rpt_class.Text) = "Pos" Then
 If Trim(cmb_provby.Text) <> Trim(cmb_finalby.Text) Then
    CrystalReport1.Formulas(11) = "PRFN= '" & Trim(cmb_provby.Text) & "'"
    CrystalReport1.Formulas(12) = "ENT = '" & Trim(cmb_enteryby.Text) & "'"
    CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
    'CrystalReport1.Formulas(14) = "P = '" & Trim(cmb_provby.Text) & "'"
    CrystalReport1.Formulas(16) = "FCAP = '" & "Pathologist" & "'"
    CrystalReport1.Formulas(17) = "F = '" & Trim(cmb_finalby.Text) & "'"
 '   CrystalReport1.Formulas(15) = "signname= '" & Trim(cmb_finalby.Text) & "'"
    'CrystalReport1.Formulas(12) = "F= '" & Trim(cmb_finalby.Text) & "'"
 Else
    CrystalReport1.Formulas(11) = "P= '" & Trim(cmb_provby.Text) & "'"
     CrystalReport1.Formulas(15) = "signname= '" & Trim(cmb_finalby.Text) & "'"
    CrystalReport1.Formulas(12) = "F= '" & Trim(cmb_finalby.Text) & "'" 'commented out by shweta
    CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
    'CRYSTALREPORT1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
 End If
 
End If
If Trim(cmb_provby.Text) <> "" And Trim(cmb_finalby.Text) = "" Then
    CrystalReport1.Formulas(11) = "PONLY= '" & Trim(cmb_provby.Text) & "'"
    CrystalReport1.Formulas(15) = "signname= '" & Trim(cmb_provby.Text) & "'"
    'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
    CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
    
End If
If Trim(cmb_provby.Text) <> "" And Trim(cmb_finalby.Text) <> "" And UCase(Trim(cmb_provby.Text)) = UCase(Trim(cmb_finalby.Text)) Then '
    CrystalReport1.Formulas(11) = "PONLY= '" & Trim(cmb_provby.Text) & "'" '29/12/09
    CrystalReport1.Formulas(15) = "signname= '" & Trim(cmb_provby.Text) & "'"
    CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '" ''29/12/09
    'No such FF in report. So ADDED by Shraddha on 29/12/09
End If

'If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
' CrystalReport1.Formulas(15) = "SIGNNAME= '" & "" & " '"
'End If

'If Trim(cmb_finalby.Text) <> "" Then
'            'CrystalReport1.Formulas(10) = "PATH1 = '" & "Pathologist" & "'"
'            CrystalReport1.Formulas(11) = "PATH = '" & Trim(cmb_finalby.Text) & "'"
'Else
'            'CrystalReport1.Formulas(10) = "path1 = '" & "" & "'"
'            CrystalReport1.Formulas(11) = "PATH = '" & Trim(cmb_finalby.Text) & "'"
'End If
CrystalReport1.WindowState = crptMaximized
'CRYSTALREPORT1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1

'rtime log
UPDRTIMEP_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIMEP_QR(1) = Format(txt_time.Text, "HH:MM:SS")
UPDRTIMEP_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIMEP_QR.Execute


Exit Sub
'prtmsg: MsgBox Error(Err)
'txt_msg.Text = ""
'btnenabl ("print")
BTNENABL1 ("print")
Exit Sub
'MSG:
'MsgBox Error(Err)
End Sub

Private Sub cmd_save_Click()
'On Error GoTo MSG:

Dim m As Integer
If Trim(cmb_finalby.Text) = "" Then
    MsgBox "Finalized By Field is Empty !!", vbInformation, "Cytology"
    cmb_finalby.SetFocus
    Exit Sub
End If
If Trim(cmb_provby.Text) = "" Then
    MsgBox "Provisional By Field is Empty !!", vbInformation, "Cytology"
    cmb_provby.SetFocus
    Exit Sub
End If
If Trim(cmb_enteryby.Text) = "" Then
    MsgBox "Entered By Field is Empty !!", vbInformation, "Cytology"
    cmb_enteryby.SetFocus
    Exit Sub
End If
If Trim(txt_noslides.Text) = "" Then
    MsgBox "Please Enter the No. of Slides !!", vbInformation, "Cytology"
    txt_noslides.SetFocus
    Exit Sub
End If
If Len(Trim(cmb_reqno)) = 0 Then
    m = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
    cmb_reqno.SetFocus
    Exit Sub
End If
If Trim(cmb_rpt_class.Text) = "" Then
    MsgBox " Please Enter Report Classification Field Data !!", vbInformation, "Cytology"
    cmb_rpt_class.SetFocus
    Exit Sub
End If

'***** Shraddha 21/03/11 *****
'Check whether collection date and time has been entered or not.
If NEWF = True Then

    Dim a As Integer
    Dim cur_date As Date
    If (IsDate(msk_coll_dt.Text)) = False Then
        MsgBox "Wrong Collection date"
        Exit Sub
    End If
    
    If msk_coll_tm.Text = "__:__" Then
        MsgBox "Please enter collection time"
        Exit Sub
    End If
    
    cur_date = Format(DATE11, "yyyy-mm-dd")
    
    If msk_coll_dt.Text > cur_date Then
        MsgBox "Collection date can not be greater than todays date"
        Exit Sub
    End If
    
    If Mid(Trim(msk_coll_tm.Text), 1, 2) > 24 Then
        MsgBox ("Please enter time in 24 hours format..."), vbOKOnly
        msk_coll_tm.Text = "__:__"
        msk_coll_tm.SetFocus
        Exit Sub
    End If
    
    If Mid(Trim(msk_coll_tm.Text), 4, 2) >= 60 Then
        MsgBox ("Minutes can not exceeds 60..."), vbOKOnly
        msk_coll_tm.Text = "__:__"
        msk_coll_tm.SetFocus
        Exit Sub
    End If
    
    'a = InStr(1, cmb_reqno, "$")
    mod_reqno = req_chglen(Trim(cmb_reqno))
    'change here 05-04-2012 shweta swati
    
    upd_rtime_qr(0) = Format(msk_coll_dt.Text, "YYYY-MM-DD")
    upd_rtime_qr(1) = Format(msk_coll_tm.Text, "HH:MM:SS")
    'upd_rtime_qr(1) = Format(time1, "HH:MM:SS")
    upd_rtime_qr(2) = mod_reqno
    upd_rtime_qr.Execute

End If

'***** 21/03/11 *****

    If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
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
       While InStr(txt_report, Chr(9)) > 0
        
            K = InStr(txt_report, Chr(9))
            beforst = Mid(txt_report, 1, K - 1) & "      "
            afterst = Mid(txt_report, K + 1)
            txt_report = beforst & afterst
       Wend
      While InStr(txt_report, Chr(146)) > 0

            K = InStr(txt_report, Chr(146))
            beforst = Mid(txt_report, 1, K - 1) & "`"
            afterst = Mid(txt_report, K + 1)
            txt_report = beforst & afterst
       
       Wend
       While InStr(txt_impdiag, Chr(9)) > 0
        
            K = InStr(txt_impdiag, Chr(9))
            beforst = Mid(txt_impdiag, 1, K - 1) & "      "
            afterst = Mid(txt_impdiag, K + 1)
            txt_impdiag = beforst & afterst
       Wend
      While InStr(txt_impdiag, Chr(146)) > 0

            K = InStr(txt_impdiag, Chr(146))
            beforst = Mid(txt_impdiag, 1, K - 1) & " "
            afterst = Mid(txt_impdiag, K + 1)
            txt_impdiag = beforst & afterst
       
       Wend
       While InStr(txt_typespeciman, Chr(9)) > 0
       
            K = InStr(txt_typespeciman, Chr(9))
            beforst = Mid(txt_typespeciman, 1, K - 1) & "      "
            afterst = Mid(txt_typespeciman, K + 1)
            txt_typespeciman = beforst & afterst
       Wend
      While InStr(txt_typespeciman, Chr(146)) > 0

            K = InStr(txt_typespeciman, Chr(146))
            beforst = Mid(txt_typespeciman, 1, K - 1) & "`"
            afterst = Mid(txt_typespeciman, K + 1)
            txt_typespeciman = beforst & afterst
       
       Wend
       If Len(Trim(txt_report.Text)) > 8000 Then
            MsgBox "Report Max. of 8000 Characters Only.", vbInformation
            txt_report.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_impdiag.Text)) > 8000 Then
            MsgBox "Impression Max. of 8000 Characters Only.", vbInformation
            txt_impdiag.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_impdiag.Text)) = 0 Then
            MsgBox "Please Enter the Impression ", vbInformation
            txt_impdiag.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_typespeciman.Text)) > 1000 Then
            MsgBox "Nature of Meterial Max. of 1000 Characters Only.", vbInformation
            txt_typespeciman.SetFocus
            Exit Sub
       End If
       If Len(Trim(txt_addendum.Text)) > 600 Then
            MsgBox "Addendum Max. of 600 Characters Only.", vbInformation
            txt_addendum.SetFocus
            Exit Sub
       End If
'rdb.BeginTrans
       assign
      'rtimelog
UPDRTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                
UPDRTIME_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIME_QR.Execute
       
'rdb.CommitTrans
       'MsgBox "Record is saved Successfully."
  FLDENABL ("DIS")
  'btnenabl ("SAVE")
  BTNENABL1 ("SAVE")
  Exit Sub
'MSG:
'rdb.RollbackTrans
'MsgBox Error(Err)
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
'btnenabl ("VIEW")
BTNENABL1 ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR
NEWF = False

'&&&&&&&&&&&&&&&&&&&&77 not found hence declared
Dim SEL_DEPTRSFLAG As Boolean

If SEL_DEPTRSFLAG = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SEL_DEPTRSFLAG = True
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
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
'cmd_commit.Enabled = True
cmb_dept.SetFocus
cmd_print.Enabled = True
Opt_patient.Value = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
'On Error GoTo MSG:
If CYREPORTFLAG = False Then
    CYREPORTFLAG = True
'btnenabl ("INI")

NEWFLAG = False
viewflag = False
MaskEdBox1.Enabled = False


''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR117", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR117", str)


''********ICU***************

str = "SELECT * FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
Set SELRMODALITY_QR = rdb.CreateQuery("SELRMODALITY1_QRXX1", str)
SELRMODALITY_QR(0) = "CT"
Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(2, 3)


str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
Set seldrequisitn_qr11 = rdb.CreateQuery("", str)
    
'DISPLAYING PREVIOUS INVESTIGATIONS
str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='F' or REQNSTAT='T') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' ) ORDER BY PREPORTDT DESC"   'Order by  DSPLF.REQNDT,DSPLF.REQNNO DESC
Set seldrequisitn_qrprv = rdb.CreateQuery("seldrequisitn_qrprv", str)
 ' seldrequisitn_qrprv(0) = "BN/00000"
seldrequisitn_qrprv(0) = ""
Set seldrequisitn_rsprv = seldrequisitn_qrprv.OpenResultset(1, 3)
'UPTO HERE

'THIS IS RTIMLOG
 str = "update rislib.rtimelog set ADPREPDATE=?,ADPREPTIME=? where reqnno=?"
 Set UPDRTIMECADD_QR = rdb.CreateQuery("updatetimecadcy", str)
 
 str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=? WHERE REQNNO=?"
 Set UPDRTIME_QR = rdb.CreateQuery("QWEQWEQWEcy", str)

 str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEccy", str)
 
 
 str = "UPDATE RISLIB.RTIMELOG SET PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEP_QR = rdb.CreateQuery("QWEQWEQWEPcy", str)

 str = "UPDATE RISLIB.RTIMELOG SET MODREPDATE=?,MODREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEM_QR = rdb.CreateQuery("QWEQWEQWEMcy", str)

'***** Shraddha 21/03/11 *****
'following line commented and new query added by shweta
'reason : in report time of the collection is s
str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?,SAMCOLTIME=? WHERE REQNNO=?"
'str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?,SAMCOLTIME=? WHERE REQNNO=? and SAMCOLDATE=NULL and SAMCOLTIME=NULL"
Set upd_rtime_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set sel_rtimelog_qr = rdb.CreateQuery("", str)
'*****************************

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enteryby.AddItem LOGINUSRNAME
    cmb_enteryby.Text = Trim(LOGINUSRNAME)
    
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STRAA", str)
If Trim(USRDEPT) = "ALL" Then
    REPWTRN_QR(0) = "FZZ"
Else
    REPWTRN_QR(0) = Trim(USRDEPT)
End If
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    cmb_addendumby.clear
    cmb_finalby.clear
    cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 9 Then   'MADE 9 TO 7
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        If REPWTRNRS("USRTYPE") >= 3 And REPWTRNRS("USRTYPE") <> 4 Then
            cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
                REPWTRNRS.MoveNext
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"

Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

'str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE > ? AND DEPT_CODE = 'FZZ' AND USR_ID NOT LIKE 'A3%'"
'FOLLOWING QUERY CHANGED BY SHWETA 04-12-2012 'AS REQUEST FROM FZZ DEPT FOR 303762
       'str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ?  AND DEPT_CODE = 'FZZ' AND USR_id NOT LIKE 'A3%'"
       str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE usrtype >=9 or ( USRTYPE >= ?  AND DEPT_CODE = 'FZZ' OR  DEPT_CODE = 'FEC' AND (USR_id NOT LIKE 'A3%' OR MODMR='A'))"
Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1", str)
SELFINALBY_QR(0) = 8 ' made 8 to 5
Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
If Not SELFINALBY_RS.EOF Then
    Do While Not SELFINALBY_RS.EOF
       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
       SELFINALBY_RS.MoveNext
    Loop
End If

  SELFINALBY_QR.Close
'*********SECURITY**********************************

''SELECT REPORT CLASSIFICATION FROM DREQUISITION TABLE
str = "SELECT IMMHISCHE FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELIMMHISCHE_QR = rdb.CreateQuery("SELIMMHISCHE_QR", str)

''UPDATE REPORT CLASSIFICATION OF DREQUISITION TABLE
str = "UPDATE DISLIB.DREQUISITN SET IMMHISCHE=? WHERE REQNO=?"
Set UPDIMMHISCHE_QR = rdb.CreateQuery("UPDIMMHISCHE_STR", str)

''update drequisitn table by slides no.
str = "UPDATE DISLIB.DREQUISITN SET LABNO=? WHERE REQNO=?"
Set UPDDREQ_QR = rdb.CreateQuery("UPDDREQ_STR", str)
   
    'To display current date
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
    
    'for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' OR DEPT_CODE='IZZ')"
Set sel_deptqr = rdb.CreateQuery("selC1_deptqrcg", str)
'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
   
           
str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301QR = rdb.CreateQuery("SELMA3011QRCG", str)

    'FOR GETTING CLINICAL DIAGNOSIS
str = "select CLINHIST  from rislib.CLINHIST WHERE CASENO=? "
Set SELCLIN_QR = rdb.CreateQuery("SELCLIN_QRCG", str)
        
    ' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqrcg", str)

    'for retriving reqnos
Set req_qr = rdb.CreateQuery("1reqCG", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT DESC")

    ' for retriveing name,age etc
Set SELTA101_QR = rdb.CreateQuery("selta101_qrcg", "select * from PATOOLSLIB.ta101 where case_no =?")

    ' FOR REFEFING DOCTOR
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
Set REFDR_QR = rdb.CreateQuery("REFDR1_QRCG", str)

    'FOR PATH NO
str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1_QRCG", str)

    'for bedno, ward retriveing
str = "select * from PATOOLSLIB.ma104 where case_no =?"
Set selma104_qr = rdb.CreateQuery("selma1041_qrcg", str)
    
    'for provisional diagnosis
str = "select * from patoolslib.ta1012 where case_no =?"
Set selta1012_qr = rdb.CreateQuery("se1lta10121_qrcg", str)
    
     
'STR = "select * from dislib.dimpresion where reqno=?"
'Set seldimpresion_qr = rdb.CreateQuery("seldimpresion_qrcg", STR)


    ' for insertion
'    STR = "insert into  dislib.dcyreport (reqno, cyreport, impression ) values (?,?,?)"
'    Set insdfs_qr = rdb.CreateQuery("insdfs_qrcg", STR)
'    STR = "insert into dislib.dgrossdesc (reqno,grossdesc,grossedby) values (?,?,?)"
'    Set insdgrossdesc_qr = rdb.CreateQuery("insdgrossdesc_qrcg", STR)
'    STR = "insert into dislib.dihreport (reqno,ihreport,enteryby,preportby,freportby) values (?,?,?,?,?)"
'    Set INSDIHREPORT_QR = rdb.CreateQuery("insdihreport_qrcg", STR)
'    STR = "insert into dislib.dimpresion (reqno,finaldiag,fdescry,t,tsize,n,m,gradecode) values (?,?,?,?,?,?,?,?)"
'    Set insdimpresion_qr = rdb.CreateQuery("1insdimpresion_qrcg", STR)
'    STR = "insert into dislib.dspreport (reqno,spreport,enterby,preportby,freportby) values(?,?,?,?,?)"
'    Set insdspreport_qr = rdb.CreateQuery("insdspreport_qrcg", STR)

    ' for update report
    str = "update dislib.dcyreport set cyreport=?,impression=? where reqno=?"
    Set upddfsreport_qr = rdb.CreateQuery("upddcyreport_qrcg", str)
    
   
    ' for update impresion - diagnosis
'    STR = "update dislib.dimpresion set  finaldiag=?, fdescry=?, t=?,tsize=?, n=?, m=?,GRADECODE=? where reqno=?"
'    Set upddimpresion_qr = rdb.CreateQuery("upddimpresion_qrcg", STR)
    
       
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QRCG", str)
    
    ' for update enterby, provreportby, finalby
    str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?, FREPORTBY=?, FREPORTDT=? WHERE REQNNO=?"
    Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qrcg", str)
   
    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrcg", str)
    
    'for insert addendum report
    str = "INSERT INTO DISLIB.ADDENDUMCY(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
    Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QRCG", str)
   
    str = "SELECT  * FROM DISLIB.ADDENDUMCY WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QRCG", str)
   
    str = "select * from dislib.drequisitn where caseno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_1qrcg", str)

    txt_addendum.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_dept.Enabled = False
    FLDENABL ("DIS")

   str = "insert into  dislib.dcyreport(reqno, cyreport, impression) values (?,?,?)"
   Set INSFREPORT_QR = rdb.CreateQuery("INSFREPORT_QRCG", str)

   ' for update report and impression
   str = "UPDATE DISLIB.dcyreport SET cyreport=?,IMPRESSION=? WHERE REQNO=?"
   Set upddfsreport_qr = rdb.CreateQuery("upddfsreportcg_qrcg", str)
    
   str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
   Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQRCG", str)
   
   str = "SELECT * FROM DISLIB.Dcyreport WHERE REQNO=?"
   Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QRCG", str)
    
   str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
   Set SELDFREPORT_IHQR = rdb.CreateQuery("SELDFREPORT_IHQRCG", str)
    
   str = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?, NMETERIALS=? WHERE REQNO=?"
   Set upddrequisitn_qr = rdb.CreateQuery("UPDSPREPORT1_QRCG", str)
    
'    STR = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
'    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO11_QR1", STR)
'    SELEMPBASINFO_QR(0) = "00000"
'    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    
    str = "select * from dislib.dgrade"
    Set seldgradeqr = rdb.CreateQuery("seldgradeqrcg", str)
    Set seldgraders = seldgradeqr.OpenResultset(1, 3)
    If seldgraders.RowCount > 0 Then
    'cmb_grade.clear
    While Not seldgraders.EOF
        'cmb_grade.AddItem Trim(seldgraders("gradedesc")) & "$" & Trim(seldgraders("gradecode"))
        seldgraders.MoveNext
    Wend
    End If
    
'    STR = "select * from dislib.dgrade where gradecode=?"
'    Set seldgrade_qr = rdb.CreateQuery("seldgrade_qrcg", STR)
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
    Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRCG", str)
    
   str = "select * from dislib.dihreport where reqno=?"
   Set selihreportqr = rdb.CreateQuery("selihreportqrcg", str)
   
   str = "select * from dislib.dissec where enterby='YES'AND MODALITY='CG' "
   Set seldissecentr_qr = rdb.CreateQuery("seldissecentr_qrcg", str)
   
   str = "select * from dislib.dissec where PROVby='YES' AND (MODALITY='CG'OR MODALITY='SP') "
   Set seldissecprov_qr = rdb.CreateQuery("seldissecPROV_qrcg", str)
   
   str = "select * from dislib.dissec where FINALby='YES' AND (MODALITY='CG'OR MODALITY='SP') "
   Set seldissecfin_qr = rdb.CreateQuery("seldissecFIN_qrcg", str)
   
   str = "select * from dislib.dissec where ADDENDUMby='YES' AND (MODALITY='CG'OR MODALITY='SP') "
   Set seldissecadd_qr = rdb.CreateQuery("seldissecADD_qrcg", str)
   
'   STR = "select * from dislib.dissec where GROSSby='YES'"
'   Set seldissecgross_qr = rdb.CreateQuery("seldissecgROSS_qrcg", STR)
   
   str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
   Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1QRCG", str)
   
   str = "select sitecd,sitedesc from dislib.sitedup where sitecd = ?"
   Set viewsitedupqr = rdb.CreateQuery("viewsitedupqrcg", str)
   
    str = "INSERT INTO RISLIB.DISIMP(REQNO,IMP)VALUES(?,?)"
   Set INSDISIMP_QR = rdb.CreateQuery("INSDISIMP_QRCG", str)
  
   str = "UPDATE RISLIB.DISIMP SET IMP=? WHERE REQNO=?"
   Set UPDDISIMP_QR = rdb.CreateQuery("UPDDISIMP_QRCG", str)
   
End If


    



'cmb_rpt_class.AddItem "N"
BTNENABL1 ("INI")
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Unload(cancel As Integer)
'Dim RS As rdoResultset
'Dim RQ As rdoQuery
'For Each RS In rdb.rdoResultsets
'    RS.Close
'Next
'
'For Each RQ In rdb.rdoQueries
'    RQ.Close
'Next
End Sub



Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
        ''MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) commented by Monika on 02/12/2022
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
                'btnenabl ("INI")
                BTNENABL1 ("INI")
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
                MsgBox " No  Requisition for this Modality"
                'btnenabl ("INI")
                BTNENABL1 ("INI")
                Exit Sub
            End If
        
        End If

End If
cmb_reqno.SetFocus
 Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_reqno_Click()

'***** Shraddha 22/03/11 *****
 Dim rqno As String
 Dim j As Integer
 Dim TM As String
         
'j = InStr(1, cmb_reqno, "$")
'rqno = req_chglen(Mid(Trim(cmb_reqno.Text), 1, j - 1))
rqno = req_chglen(Trim(cmb_reqno.Text))

 sel_rtimelog_qr(0) = rqno
 Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)
 
 If Not sel_rtimelog_rs.EOF Then
     If IsNull(sel_rtimelog_rs("SAMCOLDATE")) = False And Trim(sel_rtimelog_rs("SAMCOLDATE")) <> "" Then
         msk_coll_dt.Text = sel_rtimelog_rs("SAMCOLDATE")
         msk_coll_dt.Enabled = False
     Else
        'If NEWF = True Then
        msk_coll_dt.Enabled = True
        'End If
     End If
     If IsNull(sel_rtimelog_rs("SAMCOLTIME")) = False And Trim(sel_rtimelog_rs("SAMCOLTIME")) <> "" Then
     
        If InStr(1, sel_rtimelog_rs("SAMCOLTIME"), ":") = 2 Then
        'MsgBox (Format(sel_rtimelog_rs("SAMCOLTIME"), "HH:MM:SS"))
            'TM = "0" & Mid(sel_rtimelog_rs("SAMCOLTIME"), 1, 4)
            'SHWETA 11-04-2012 TIME WAS SHOWING ONLY AM
             TM = Format(sel_rtimelog_rs("SAMCOLTIME"), "HH:MM")
            msk_coll_tm.Text = TM
            msk_coll_tm.Enabled = False
        ElseIf InStr(1, sel_rtimelog_rs("SAMCOLTIME"), ":") = 3 Then
            ' msk_coll_tm.Text = Mid(sel_rtimelog_rs("SAMCOLTIME"), 1, 5)
           'SHWETA 11-04-2012 TIME WAS SHOWING ONLY AM
             TM = Format(sel_rtimelog_rs("SAMCOLTIME"), "HH:MM")
             msk_coll_tm.Text = TM
            msk_coll_tm.Enabled = False
        End If
     Else
         msk_coll_tm.Enabled = True
     End If
 End If
'***** 22/03/11 ****************



Dim prpathno As String
Dim ADDENDUM As String
'On Error GoTo MSG:

If Trim(txt_modcode.Text) = "FA" Or Trim(txt_modcode.Text) = "CN" Then
    cmb_rpt_class.clear
    cmb_rpt_class.AddItem "Inadequate"
    cmb_rpt_class.AddItem "Negative"
    cmb_rpt_class.AddItem "Atypia"
    cmb_rpt_class.AddItem "Positive"
Else
    cmb_rpt_class.clear
    cmb_rpt_class.AddItem "Unsuitable"
    cmb_rpt_class.AddItem "Negative"
    cmb_rpt_class.AddItem "Organisms"
    cmb_rpt_class.AddItem "Reactive"
    cmb_rpt_class.AddItem "ASCUS"
    cmb_rpt_class.AddItem "ASC-H"
    cmb_rpt_class.AddItem "LSIL"
    cmb_rpt_class.AddItem "HSIL"
    cmb_rpt_class.AddItem "Squamous Carcinoma"
    cmb_rpt_class.AddItem "AGC"
    cmb_rpt_class.AddItem "AIS"
    cmb_rpt_class.AddItem "Adenocarcinoma"
    cmb_rpt_class.AddItem "Carcinoma NOS"
    cmb_rpt_class.AddItem "Other Malignancies"
    cmb_rpt_class.AddItem "Recurrent Squamous Carcinoma"
    cmb_rpt_class.AddItem "Recurrent Adenocarcinoma"
    cmb_rpt_class.AddItem "Residual Carcinoma"
End If
If Opt_patient.Value = True Then
    If NEWFLAG = True Then
        NEWFLAG = False
               
        ''******ICU****************
        VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If SELICU_FLAG = False Then
            SELICU_QR(0) = VAR_REQNNO
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
            
        'ADDED BY MILIND ON 091006
        seldrequisitn_qrprv(0) = MaskEdBox1.FormattedText
        seldrequisitn_rsprv.MoreResults
        seldrequisitn_rsprv.Requery
        If seldrequisitn_rsprv.RowCount > 0 Then
            'prpathno = ""
            CMB_PATHNO.clear
            While Not seldrequisitn_rsprv.EOF
                MODALITY1 = Mid(seldrequisitn_rsprv("REQNNO"), 4, 2)
                If Trim(MODALITY1) = "SP" Then
                    MODDESC = "SURGICAL PATHOLOGY"
                ElseIf Trim(MODALITY1) = "CN" Then
                    MODDESC = "CYTOPATHOLOGY NON GYNAEC"
                ElseIf Trim(MODALITY1) = "CG" Then
                    MODDESC = "CYTOPATHOLOGY GYNAEC"
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
                'If Trim(TXT_PATH.Text) <> seldrequisitn_rsprv("pathno") Or IsNull(seldrequisitn_rsprv("pathno")) Then
                'prpathno = prpathno & " " & seldrequisitn_rsprv("pathno")
                'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " $ " & seldrequisitn_rs10("PREPORTDT") & " " & "$" & " " & req_chglen(seldrequisitn_rs10("REQNNO")) & " " & "$" & " " & MODDESC
                CMB_PATHNO.AddItem Trim(seldrequisitn_rsprv("REQNNO")) & " $ " & Trim(seldrequisitn_rsprv("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rsprv("PATHNO"))
                'CMB_PATHNO.ListIndex = 0
                'End If
                seldrequisitn_rsprv.MoveNext
            Wend
            CMB_PATHNO.Text = CMB_PATHNO.List(0)
        End If
        'UPTO HERE
        
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
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
        If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If REFDR_RS("recordno") = "1" Then
                LBL_VPP.Caption = "VPP"
                If Trim(REFDR_RS("RECPTNO")) = "VPP" Then
                    MsgBox "Payment is not Done"
                    'btnenabl ("INI")
                    BTNENABL1 ("INI")
                    field_clear
                    Exit Sub
                End If
            End If
            If REFDR_RS("recordno") = "2" Then
                LBL_VPP.Caption = "LR"
            End If
            If REFDR_RS("recordno") = "3" Then
                LBL_VPP.Caption = "BAR"
            End If
            If REFDR_RS("recordno") = "4" Then
                LBL_VPP.Caption = "DrOp"
            End If
        End If
        '            If REFDR_RS.RowCount > 0 Then
        '            txt_refdoc.Text = REFDR_RS("REFBYDOC")
        '            End If
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
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_cytology.Text = selpathdrequisitn_rs("PATHNO")
            If selpathdrequisitn_rs.EOF = False Then
                txt_typespeciman.Text = " " & selpathdrequisitn_rs("NMETERIALS")
            Else
                txt_typespeciman.Text = ""
            End If
            If IsNull(Trim(selpathdrequisitn_rs("SITE"))) Then
                'txt_sites.Text = ""
            Else
                If SELMA301FLAG = False Then
                    SELMA301QR(0) = Trim(selpathdrequisitn_rs("SITE"))
                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                    SELMA301FLAG = True
                Else
                    SELMA301QR(0) = Trim(selpathdrequisitn_rs("SITE"))
                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                    SELMA301RS.MoreResults
                    SELMA301RS.Requery
                End If
                If SELMA301RS.RowCount > 0 Then
                    TXT_CSITE.Text = Trim(SELMA301RS("SITEDESC"))
                    'txt_noofslides.Text = SELPATHDREQUISITN_RS("OUTSLIDES")
                Else
                If Trim(selpathdrequisitn_rs("SITE")) = "" Then
                Else
                
                    If VIEWSITEDUPRSFLAG = False Then
                       viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
                        Set viewsiteduprs = viewsitedupqr.OpenResultset(1, 3)
                        VIEWSITEDUPRSFLAG = True
                    Else
                        viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
                        viewsiteduprs.MoreResults
                        viewsiteduprs.Requery
                    End If
                    If viewsiteduprs.RowCount > 0 Then
                        txt_sites.Text = viewsiteduprs("SITEDESC")
                    End If
                    End If
                End If
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
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
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
            prpathno = ""
            While Not seldrequisitn_rs.EOF
                prpathno = prpathno & " " & seldrequisitn_rs("pathno")
                'CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
                seldrequisitn_rs.MoveNext
            Wend
            '         txt_prepathreqno.Text = prpathno
        End If
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
        
        '        If seldissecentrflag = False Then
        '            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
        '            seldissecentrflag = True
        '        Else
        '                seldissecentr_rs.MoreResults
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
        '            End If
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
        '
    End If
    '
    If viewflag = True Then
        viewflag = False
        ''******ICU****************
        VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If SELICU_FLAG = False Then
            SELICU_QR(0) = VAR_REQNNO
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
            
        seldrequisitn_qrprv(0) = MaskEdBox1.FormattedText
        seldrequisitn_rsprv.MoreResults
        seldrequisitn_rsprv.Requery
        If seldrequisitn_rsprv.RowCount > 0 Then
            'prpathno = ""
            CMB_PATHNO.clear
            While Not seldrequisitn_rsprv.EOF
                MODALITY1 = Mid(seldrequisitn_rsprv("REQNNO"), 4, 2)
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
                'If Trim(TXT_PATH.Text) <> seldrequisitn_rsprv("pathno") Or IsNull(seldrequisitn_rsprv("pathno")) Then
                'prpathno = prpathno & " " & seldrequisitn_rsprv("pathno")
                'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " $ " & seldrequisitn_rs10("PREPORTDT") & " " & "$" & " " & req_chglen(seldrequisitn_rs10("REQNNO")) & " " & "$" & " " & MODDESC
                CMB_PATHNO.AddItem Trim(seldrequisitn_rsprv("REQNNO")) & " $ " & Trim(seldrequisitn_rsprv("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rsprv("PATHNO"))
                'CMB_PATHNO.ListIndex = 0
                'End If
                seldrequisitn_rsprv.MoveNext
            Wend
            CMB_PATHNO.Text = CMB_PATHNO.List(0)
        End If

        If SELIMMHISCHE_FLAG = False Then
            SELIMMHISCHE_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELIMMHISCHE_RS = SELIMMHISCHE_QR.OpenResultset(1, 3)
            SELIMMHISCHE_FLAG = True
        Else
            SELIMMHISCHE_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELIMMHISCHE_RS.MoreResults
            SELIMMHISCHE_RS.Requery
        End If
        
        If Not SELIMMHISCHE_RS.EOF Then
                
            cmb_rpt_class.clear
            If Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "P" Then
                cmb_rpt_class.AddItem "Positive"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "N" Then
                cmb_rpt_class.AddItem "Negative"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "I" Then
                cmb_rpt_class.AddItem "Inadequate"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "A" Then
                cmb_rpt_class.AddItem "Atypia"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "1" Then
                cmb_rpt_class.AddItem "Unsuitable"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "2" Then
                cmb_rpt_class.AddItem "Negative"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "3" Then
                cmb_rpt_class.AddItem "Organisms"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "4" Then
                cmb_rpt_class.AddItem "Reactive"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "5" Then
                cmb_rpt_class.AddItem "ASCUS"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "6" Then
                cmb_rpt_class.AddItem "ASC-H"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "7" Then
                cmb_rpt_class.AddItem "LSIL"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "8" Then
                cmb_rpt_class.AddItem "HSIL"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "9" Then
                cmb_rpt_class.AddItem "Squamous Carcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "B" Then
                cmb_rpt_class.AddItem "AGC"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "C" Then
                cmb_rpt_class.AddItem "AIS"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "D" Then
                cmb_rpt_class.AddItem "Adenocarcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "E" Then
                cmb_rpt_class.AddItem "Carcinoma NOS"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "F" Then
                cmb_rpt_class.AddItem "Other Malignancies"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "G" Then
                cmb_rpt_class.AddItem "Recurrent Squamous Carcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "H" Then
                cmb_rpt_class.AddItem "Recurrent Adenocarcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "J" Then
                cmb_rpt_class.AddItem "Residual Carcinoma"
            End If
                If cmb_rpt_class.ListCount > 0 Then
                cmb_rpt_class.Text = cmb_rpt_class.List(0)
                End If
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
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
    
        If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If REFDR_RS("COMMITFLAG") = "Y" Then
                MsgBox "Final Report is Generated, You cann't Modify the Final Report,You can Add ADDENDUMS Only"
                cmd_modify.Enabled = True
                ADDENDUMFALG = True
            End If
            If REFDR_RS("recordno") = "1" Then
            LBL_VPP.Caption = "VPP"
            End If
            If REFDR_RS("recordno") = "2" Then
            LBL_VPP.Caption = "LR"
            End If
            If REFDR_RS("recordno") = "3" Then
            LBL_VPP.Caption = "BAR"
            End If
            If REFDR_RS("recordno") = "4" Then
            LBL_VPP.Caption = "DrOp"
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
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_noslides.Text = "" & selpathdrequisitn_rs("LABNO")
            'TXT_PATH.Text = SELPATHDREQUISITN_RS("PATHNO")
            If IsNull(selpathdrequisitn_rs("SITE")) Then
            Else
                If SELMA301FLAG = False Then
                    SELMA301QR(0) = Trim(selpathdrequisitn_rs("SITE"))
                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                    SELMA301FLAG = True
                Else
                    SELMA301QR(0) = Trim(selpathdrequisitn_rs("SITE"))
                    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                    SELMA301RS.MoreResults
                    SELMA301RS.Requery
            
                End If
                If SELMA301RS.RowCount > 0 Then
                    TXT_CSITE.Text = SELMA301RS("SITEDESC")
                Else
                     TXT_CSITE.Text = ""
                End If
            End If
            
            If selpathdrequisitn_rs.EOF = False Then
                
                txt_typespeciman.Text = " " & selpathdrequisitn_rs("NMETERIALS").Value
            Else
                txt_typespeciman.Text = ""
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
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
        If seldfreportflag = False Then
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
            seldfreportflag = True
        Else
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDFREPORT_RS.MoreResults
            SELDFREPORT_RS.Requery
        End If
                    
        
           If SELDFREPORT_RS.RowCount > 0 Then
               If SELDFREPORT_RS.EOF = False Then
               If IsNull(SELDFREPORT_RS("CYREPORT")) Then
                       txt_report.Text = ""
                       Else
                       SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                        SELDFREPORT_RS.MoreResults
                        SELDFREPORT_RS.Requery
                       txt_report.Text = SELDFREPORT_RS("CYREPORT").Value
                       End If
'
'                   txt_report.Text = SELDFREPORT_RS("CYREPORT").Value
                Else
                   txt_report.Text = ""
                End If
                
                If SELDFREPORT_RS.EOF = False Then
                   txt_impdiag.Text = SELDFREPORT_RS("IMPRESSION").Value
                Else
                   txt_impdiag.Text = ""
                End If
            Else
                'txt_immhischerep.Text = ""
            End If
            
        If REFDR_RS.RowCount > 0 Then
         If SELDFREPORT_RS.EOF = False Then
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
         prpathno = ""
'         CMB_PATHNO.clear
         While Not seldrequisitn_rs.EOF
            prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'            CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
            seldrequisitn_rs.MoveNext
         Wend
        End If
            If SELDADDENDUMFLAG = False Then
                SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
                SELDADDENDUMFLAG = True
            Else
                SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
                SELDADDENDUM_RS.MoreResults
                SELDADDENDUM_RS.Requery
            End If
            If SELDADDENDUM_RS.RowCount > 0 Then
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
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
'        txt_clindiag.Enabled = False1
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
       ' MaskEdBox1.Mask = "??>/#####"
        MaskEdBox1.Enabled = True
        'MaskEdBox1.SetFocus
    End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    MaskEdBox1.Mask = "#######"
    MaskEdBox1.Enabled = True
    'MaskEdBox1.SetFocus
End If
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

Public Sub assign()
'On Error GoTo MSG:
Dim i As Integer
Dim MICROSCOPIC As String
rdb.BeginTrans
If flagsave = 1 Then
   flagsave = 0

''    '**** INSERTING INTO FREPORT
        INSFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
        If Len(Trim(txt_report.Text)) > 0 Then
            INSFREPORT_QR(1) = Null
            INSFREPORT_QR(1) = txt_report.Text
        Else
            INSFREPORT_QR(1) = Null
        End If
        If Len(Trim(txt_impdiag.Text)) > 0 Then
            INSFREPORT_QR(2) = Null
            INSFREPORT_QR(2) = txt_impdiag.Text
        Else
            INSFREPORT_QR(2) = Null
        End If
        INSFREPORT_QR.Execute
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    UPDDREQ_QR(0) = Trim(txt_noslides.Text)
    UPDDREQ_QR(1) = Trim(req_chglen(cmb_reqno.Text))
    UPDDREQ_QR.Execute
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    '**** INSERTING INTO FREPORT
    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    
    upddrequisitn_qr(0) = ""
    If Len(Trim(txt_typespeciman.Text)) > 0 Then
        upddrequisitn_qr(1) = Null
        upddrequisitn_qr(1) = txt_typespeciman.Text
    Else
        upddrequisitn_qr(1) = Null
        upddrequisitn_qr(1) = " "
    End If
    upddrequisitn_qr(2) = req_chglen(cmb_reqno.Text)
    upddrequisitn_qr.Execute
    
     '**** INSERTING INTO DISIMP
        INSDISIMP_QR(0) = req_chglen(cmb_reqno.Text)
        INSDISIMP_QR(1) = Null
        INSDISIMP_QR(1) = Trim(txt_impdiag.Text)
        INSDISIMP_QR.Execute
    
  '**** INSERTING INTO DISIMP
''*********UPDATE REPORT CLASSIFICATION DREQUISITN TABLE(IMMHISCHE)**********
    
        
    If Trim(txt_modcode.Text) = "FA" Or Trim(txt_modcode.Text) = "CN" Then
    
        UPDIMMHISCHE_QR(0) = Trim(cmb_rpt_class.Text)
        UPDIMMHISCHE_QR(1) = req_chglen(cmb_reqno.Text)
        UPDIMMHISCHE_QR.Execute
    
    Else
    
        If Trim(cmb_rpt_class.Text) = "Unsuitable" Then
            UPDIMMHISCHE_QR(0) = "1"
        ElseIf Trim(cmb_rpt_class.Text) = "Negative" Then
            UPDIMMHISCHE_QR(0) = "2"
        ElseIf Trim(cmb_rpt_class.Text) = "Organisms" Then
            UPDIMMHISCHE_QR(0) = "3"
        ElseIf Trim(cmb_rpt_class.Text) = "Reactive" Then
            UPDIMMHISCHE_QR(0) = "4"
        ElseIf Trim(cmb_rpt_class.Text) = "ASCUS" Then
            UPDIMMHISCHE_QR(0) = "5"
        ElseIf Trim(cmb_rpt_class.Text) = "ASC-H" Then
            UPDIMMHISCHE_QR(0) = "6"
        ElseIf Trim(cmb_rpt_class.Text) = "LSIL" Then
            UPDIMMHISCHE_QR(0) = "7"
        ElseIf Trim(cmb_rpt_class.Text) = "HSIL" Then
            UPDIMMHISCHE_QR(0) = "8"
        ElseIf Trim(cmb_rpt_class.Text) = "Squamous Carcinoma" Then
            UPDIMMHISCHE_QR(0) = "9"
        ElseIf Trim(cmb_rpt_class.Text) = "AGC" Then
            UPDIMMHISCHE_QR(0) = "B"
        ElseIf Trim(cmb_rpt_class.Text) = "AIS" Then
            UPDIMMHISCHE_QR(0) = "C"
        ElseIf Trim(cmb_rpt_class.Text) = "Adenocarcinoma" Then
            UPDIMMHISCHE_QR(0) = "D"
        ElseIf Trim(cmb_rpt_class.Text) = "Carcinoma NOS" Then
            UPDIMMHISCHE_QR(0) = "E"
        ElseIf Trim(cmb_rpt_class.Text) = "Other Malignancies" Then
            UPDIMMHISCHE_QR(0) = "F"
        ElseIf Trim(cmb_rpt_class.Text) = "Recurrent Squamous Carcinoma" Then
            UPDIMMHISCHE_QR(0) = "G"
        ElseIf Trim(cmb_rpt_class.Text) = "Recurrent Adenocarcinoma" Then
            UPDIMMHISCHE_QR(0) = "H"
        ElseIf Trim(cmb_rpt_class.Text) = "Residual Carcinoma" Then
            UPDIMMHISCHE_QR(0) = "J"
        End If
        UPDIMMHISCHE_QR(1) = req_chglen(cmb_reqno.Text)
        UPDIMMHISCHE_QR.Execute
    End If
    
''*********UPDATE REPORT CLASSIFICATION DREQUISITN TABLE(IMMHISCHE)**********
End If
  

If flagsave = 2 Then
    flagsave = 0
    
    ''*********UPDATE DREQUISITN TABLE(IMMHISCHE)**********
    
    If Trim(txt_modcode.Text) = "FA" Or Trim(txt_modcode.Text) = "CN" Then
    
        UPDIMMHISCHE_QR(0) = Trim(cmb_rpt_class.Text)
        UPDIMMHISCHE_QR(1) = req_chglen(cmb_reqno.Text)
        UPDIMMHISCHE_QR.Execute
    
    Else
    
        If Trim(cmb_rpt_class.Text) = "Unsuitable" Then
            UPDIMMHISCHE_QR(0) = "1"
        ElseIf Trim(cmb_rpt_class.Text) = "Negative" Then
            UPDIMMHISCHE_QR(0) = "2"
        ElseIf Trim(cmb_rpt_class.Text) = "Organisms" Then
            UPDIMMHISCHE_QR(0) = "3"
        ElseIf Trim(cmb_rpt_class.Text) = "Reactive" Then
            UPDIMMHISCHE_QR(0) = "4"
        ElseIf Trim(cmb_rpt_class.Text) = "ASCUS" Then
            UPDIMMHISCHE_QR(0) = "5"
        ElseIf Trim(cmb_rpt_class.Text) = "ASC-H" Then
            UPDIMMHISCHE_QR(0) = "6"
        ElseIf Trim(cmb_rpt_class.Text) = "LSIL" Then
            UPDIMMHISCHE_QR(0) = "7"
        ElseIf Trim(cmb_rpt_class.Text) = "HSIL" Then
            UPDIMMHISCHE_QR(0) = "8"
        ElseIf Trim(cmb_rpt_class.Text) = "Squamous Carcinoma" Then
            UPDIMMHISCHE_QR(0) = "9"
        ElseIf Trim(cmb_rpt_class.Text) = "AGC" Then
            UPDIMMHISCHE_QR(0) = "B"
        ElseIf Trim(cmb_rpt_class.Text) = "AIS" Then
            UPDIMMHISCHE_QR(0) = "C"
        ElseIf Trim(cmb_rpt_class.Text) = "Adenocarcinoma" Then
            UPDIMMHISCHE_QR(0) = "D"
        ElseIf Trim(cmb_rpt_class.Text) = "Carcinoma NOS" Then
            UPDIMMHISCHE_QR(0) = "E"
        ElseIf Trim(cmb_rpt_class.Text) = "Other Malignancies" Then
            UPDIMMHISCHE_QR(0) = "F"
        ElseIf Trim(cmb_rpt_class.Text) = "Recurrent Squamous Carcinoma" Then
            UPDIMMHISCHE_QR(0) = "G"
        ElseIf Trim(cmb_rpt_class.Text) = "Recurrent Adenocarcinoma" Then
            UPDIMMHISCHE_QR(0) = "H"
        ElseIf Trim(cmb_rpt_class.Text) = "Residual Carcinoma" Then
            UPDIMMHISCHE_QR(0) = "J"
        End If
        UPDIMMHISCHE_QR(1) = req_chglen(cmb_reqno.Text)
        UPDIMMHISCHE_QR.Execute
    End If

    
'    UPDIMMHISCHE_QR(0) = Trim(cmb_rpt_class.Text)
'    UPDIMMHISCHE_QR(1) = req_chglen(cmb_reqno.Text)
'    UPDIMMHISCHE_QR.Execute
''*********UPDATE DREQUISITN TABLE(IMMHISCHE)**********
    
    upddfsreport_qr(0) = Null
    upddfsreport_qr(0) = Trim(txt_report.Text)
    upddfsreport_qr(1).Value = Null
    upddfsreport_qr(1) = Trim(txt_impdiag.Text)
    upddfsreport_qr(2) = Trim(req_chglen(cmb_reqno.Text))
    upddfsreport_qr.Execute
    
    upddrequisitn_qr(0) = ""
    If Len(Trim(txt_typespeciman.Text)) > 0 Then
        upddrequisitn_qr(1) = Null
        upddrequisitn_qr(1) = txt_typespeciman.Text
    Else
        upddrequisitn_qr(1) = Null
        upddrequisitn_qr(1) = " "
    End If
    upddrequisitn_qr(2) = req_chglen(cmb_reqno.Text)
    upddrequisitn_qr.Execute
    
    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
   
    '**** UPDATE THE DISIMP
        If Len(Trim(txt_impdiag.Text)) > 0 Then
        UPDDISIMP_QR(0) = Null
        UPDDISIMP_QR(0) = Trim(txt_impdiag.Text)
        UPDDISIMP_QR(1) = Trim(req_chglen(cmb_reqno.Text))
        UPDDISIMP_QR.Execute
        End If
        
    '*** UPDATE THE DISIMP
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    UPDDREQ_QR(0) = Trim(txt_noslides.Text)
    UPDDREQ_QR(1) = Trim(req_chglen(cmb_reqno.Text))
    UPDDREQ_QR.Execute
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    
    '**** UPDATING INTO FREPORT
    
    If ADDENDUMFALG = True Then
        ''********SECURITY***************************
        CURUSRDET_QR(0) = Trim(cmb_finalby.Text) ' comment for Sangrur as dr type=10  on  15.07.2017
        CURUSRDET_QR(0) = Trim(cmb_addendumby.Text)
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
            MsgBox "Sorry! The Security Details of this user is not available check up in DISLNKUSR"
            cmd_commit.SetFocus
            Exit Sub
        End If
''********SECURITY***************************
        
        
        If SELDADDENDUMFLAG = False Then
            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
            SELDADDENDUMFLAG = True
        Else
            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            SELDADDENDUM_RS.MoreResults
            SELDADDENDUM_RS.Requery
        End If
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
    
  End If
  rdb.CommitTrans
  MsgBox "Record is Successfully Saved."
  Exit Sub
'MSG:
 '   rdb.RollbackTrans
 '   MsgBox Err.Description
End Sub

Public Sub field_clear()
Dim obj As Object
cmb_rpt_class.ListIndex = -1
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
cmb_reqno.ListIndex = -1
cmb_reqno.clear
cmb_addendumby.ListIndex = -1
cmb_addendumby.Text = ""
cmb_enteryby.ListIndex = -1
cmb_enteryby.Text = ""
cmb_provby.ListIndex = -1
cmb_provby.Text = ""
cmb_finalby.ListIndex = -1
cmb_finalby.Text = ""
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
CMB_PATHNO.clear
CMB_PATHNO.ListIndex = -1
If Opt_patient.Value = True Then
   ' MaskEdBox1.Mask = ">??/#####"
ElseIf opt_staff.Value = True Then
    MaskEdBox1.Mask = ">######"
End If
LBL_VPP.Caption = ""
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In cytology_report.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
msk_coll_dt.Text = "__/__/____"
msk_coll_tm.Text = "__:__"
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    cmb_rpt_class.Enabled = True
    txt_addendum.Enabled = True
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provby.Enabled = True
    cmb_addendumby.Enabled = True
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_dept.Enabled = True
    cmb_reqno.Enabled = True
    txt_typespeciman.Enabled = True
    txt_impdiag.Enabled = True
    txt_report.Enabled = True
    txt_noslides.Enabled = True
    txt_addendum.Enabled = False
    TXT_CSITE.Enabled = False
    CMB_PATHNO.Enabled = True
End If
If str1 = "DIS" Then
    cmb_rpt_class.Enabled = False
    txt_noslides.Enabled = False
    txt_addendum.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    txt_typespeciman.Enabled = False
    txt_report.Enabled = False
    txt_addendum.Enabled = False
    txt_impdiag.Enabled = False
    TXT_CSITE.Enabled = False
End If
If str1 = "VIEW" Then
    cmb_rpt_class.Enabled = False
    txt_noslides.Enabled = False
    txt_addendum.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_typespeciman.Enabled = False
    txt_report.Enabled = False
    txt_addendum.Enabled = False
    txt_impdiag.Enabled = False
    TXT_CSITE.Enabled = False
End If
If str1 = "MODIFY" Then
    cmb_rpt_class.Enabled = True
    txt_noslides.Enabled = True
    txt_addendum.Enabled = False
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provby.Enabled = True
    cmb_addendumby.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_typespeciman.Enabled = True
    txt_report.Enabled = True
    txt_addendum.Enabled = False
    txt_impdiag.Enabled = True
    TXT_CSITE.Enabled = False
    CMB_PATHNO.Enabled = True
End If
If str1 = "ADDENDUM" Then
    cmb_rpt_class.Enabled = False
    txt_addendum.Enabled = True
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
'    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    txt_typespeciman.Enabled = False
    txt_report.Enabled = False
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
        cmb_addendumby.Enabled = True
    Else
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
    txt_impdiag.Enabled = False
    TXT_CSITE.Enabled = False
    txt_noslides.Enabled = False
End If
End Sub

Public Sub STAFF()
'On Error GoTo MSG:
If NEWFLAG = True Then
    NEWFLAG = False
    ''******ICU****************
            VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = VAR_REQNNO
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
            
    
    
        seldrequisitn_qrprv.QueryTimeout = 300
        If seldrequisitn_flagprv = False Then
            seldrequisitn_qrprv(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            Set seldrequisitn_rsprv = seldrequisitn_qrprv.OpenResultset(1, 3)
            seldrequisitn_flagprv = True
        Else
            seldrequisitn_qrprv(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            seldrequisitn_rsprv.MoreResults
            seldrequisitn_rsprv.Requery
        End If
        If seldrequisitn_rsprv.RowCount > 0 Then
            'prpathno = ""
            CMB_PATHNO.clear
                While Not seldrequisitn_rsprv.EOF
                    MODALITY1 = Mid(seldrequisitn_rsprv("REQNNO"), 4, 2)
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
                    'prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
                    'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & seldrequisitn_rs10("PREPORTDT") & " " & "$" & " " & seldrequisitn_rs10("REQNO") & " " & "$" & " " & MODDESC
                     CMB_PATHNO.AddItem Trim(seldrequisitn_rsprv("REQNNO")) & " $ " & Trim(seldrequisitn_rsprv("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rsprv("PATHNO"))
                    
                    seldrequisitn_rsprv.MoveNext
                Wend
                 CMB_PATHNO.Text = CMB_PATHNO.List(0)
                'txt_prepathreqno.Text = prpathno
        End If

    If SELDREQSTAFFD1FLAG1 = False Then
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
        SELDREQSTAFFD1QR1(1) = MaskEdBox1.Text
        Set SELDREQSTAFFD1RS1 = SELDREQSTAFFD1QR1.OpenResultset(1, 3)
    Else
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
        SELDREQSTAFFD1QR1(1) = MaskEdBox1.Text
        SELDREQSTAFFD1RS1.MoreResults
        SELDREQSTAFFD1RS1.Requery
    End If
    If SELDREQSTAFFD1RS1.RowCount > 0 Then
        txt_name = SELDREQSTAFFD1RS1("PNAME")
        txt_sex = SELDREQSTAFFD1RS1("SEX")
        txt_age = SELDREQSTAFFD1RS1("AGE")
        If SELDREQSTAFFD1RS1("RELATION") <> "SELF" Then
            txt_status = "DP"
        Else
            txt_status = "ST"
        End If
    End If

     If REFDRFLAG = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
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
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
    '****
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_cytology.Text = selpathdrequisitn_rs("PATHNO")
            If selpathdrequisitn_rs.EOF = False Then
                txt_typespeciman.Text = " " & selpathdrequisitn_rs("NMETERIALS")
            Else
                txt_typespeciman.Text = ""
            End If
            If IsNull(Trim(selpathdrequisitn_rs("SITE"))) Then
                'txt_sites.Text = ""
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
                    TXT_CSITE.Text = SELMA301RS("SITEDESC")
                    'txt_noofslides.Text = SELPATHDREQUISITN_RS("OUTSLIDES")
                    Else
                            If VIEWSITEDUPRSFLAG = False Then
                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
                                 Set viewsiteduprs = viewsitedupqr.OpenResultset(1, 3)
                                 VIEWSITEDUPRSFLAG = True
                            Else
                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
                                 viewsiteduprs.MoreResults
                                 viewsiteduprs.Requery
                            End If
                            If viewsiteduprs.RowCount > 0 Then
                                txt_sites.Text = viewsiteduprs("SITEDESC")
                            End If
                    End If
            End If
        End If

    '****
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
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
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
         prpathno = ""
         While Not seldrequisitn_rs.EOF
         prpathno = prpathno & " " & seldrequisitn_rs("pathno")
         'CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
         seldrequisitn_rs.MoveNext
         Wend
'         txt_prepathreqno.Text = prpathno
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
'            While Not seldissecfin_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
'                seldissecfin_rs.MoveNext
'            Wend
'        End If

End If
If viewflag = True Then

    viewflag = False
    ''******ICU****************
            VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = VAR_REQNNO
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
            
            seldrequisitn_qrprv.QueryTimeout = 300
        If seldrequisitn_flagprv = False Then
            seldrequisitn_qrprv(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            Set seldrequisitn_rsprv = seldrequisitn_qrprv.OpenResultset(1, 3)
            seldrequisitn_flagprv = True
        Else
            seldrequisitn_qrprv(0) = MaskEdBox1.FormattedText
            'seldrequisitn_qr10(1) = req_chglen(Trim(cmb_reqno.Text))
            seldrequisitn_rsprv.MoreResults
            seldrequisitn_rsprv.Requery
        End If
        If seldrequisitn_rsprv.RowCount > 0 Then
            'prpathno = ""
            CMB_PATHNO.clear
                While Not seldrequisitn_rsprv.EOF
                    MODALITY1 = Mid(seldrequisitn_rsprv("REQNNO"), 4, 2)
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
                    'prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
                    'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & seldrequisitn_rs10("PREPORTDT") & " " & "$" & " " & seldrequisitn_rs10("REQNO") & " " & "$" & " " & MODDESC
                     CMB_PATHNO.AddItem Trim(seldrequisitn_rsprv("REQNNO")) & " $ " & Trim(seldrequisitn_rsprv("PREPORTDT")) & " $ " & Trim(MODDESC) & " $ " & Trim(seldrequisitn_rsprv("PATHNO"))
                    
                    seldrequisitn_rsprv.MoveNext
                Wend
                 CMB_PATHNO.Text = CMB_PATHNO.List(0)
                'txt_prepathreqno.Text = prpathno
        End If

    If SELIMMHISCHE_FLAG = False Then
        SELIMMHISCHE_QR(0) = VAR_REQNNO
        Set SELIMMHISCHE_RS = SELIMMHISCHE_QR.OpenResultset(1, 3)
        SELIMMHISCHE_FLAG = True
    Else
        SELIMMHISCHE_QR(0) = VAR_REQNNO
        SELIMMHISCHE_RS.MoreResults
        SELIMMHISCHE_RS.Requery
    End If
    If Not SELIMMHISCHE_RS.EOF Then
        If Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "P" Then
                cmb_rpt_class.Text = "Positive"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "N" Then
                cmb_rpt_class.Text = "Negative"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "I" Then
                cmb_rpt_class.Text = "Inadequate"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "A" Then
                cmb_rpt_class.Text = "Atypia"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "1" Then
                cmb_rpt_class.Text = "Unsuitable"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "2" Then
                cmb_rpt_class.Text = "Negative"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "3" Then
                cmb_rpt_class.Text = "Organisms"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "4" Then
                cmb_rpt_class.Text = "Reactive"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "5" Then
                cmb_rpt_class.Text = "ASCUS"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "6" Then
                cmb_rpt_class.Text = "ASC-H"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "7" Then
                cmb_rpt_class.Text = "LSIL"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "8" Then
                cmb_rpt_class.Text = "HSIL"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "9" Then
                cmb_rpt_class.Text = "Squamous Carcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "B" Then
                cmb_rpt_class.Text = "AGC"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "C" Then
                cmb_rpt_class.Text = "AIS"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "D" Then
                cmb_rpt_class.Text = "Adenocarcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "E" Then
                cmb_rpt_class.Text = "Carcinoma NOS"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "F" Then
                cmb_rpt_class.Text = "Other Malignancies"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "G" Then
                cmb_rpt_class.Text = "Recurrent Squamous Carcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "H" Then
                cmb_rpt_class.Text = "Recurrent Adenocarcinoma"
            ElseIf Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "J" Then
                cmb_rpt_class.Text = "Residual Carcinoma"
            End If
    End If
    If SELDREQSTAFFD1FLAG1 = False Then
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
        SELDREQSTAFFD1QR1(1) = Trim(MaskEdBox1.Text)
        Set SELDREQSTAFFD1RS1 = SELDREQSTAFFD1QR1.OpenResultset(1, 3)
    Else
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
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

    If REFDRFLAG = False Then
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            REFDRFLAG = True
        Else
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
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
        txt_refdoc.Text = Trim(REFDR_RS("REFBYDOC"))
        If REFDR_RS("COMMITFLAG") = "Y" Then
            MsgBox "Final Report is Generated, You cann't Modify the Final Report,You can Add ADDENDUMS Only"
            cmd_modify.Enabled = False
            ADDENDUMFALG = True
        End If
    End If
  
    If selma104flag = False Then
            selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
            selma104flag = True
        Else
            selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
            selma104_rs.MoreResults
            selma104_rs.Requery
        End If
  
  
'    If selma104_rs.RowCount > 0 Then
'        txt_ward.Text = selma104_rs("WARD_NAME")
'        txt_bedno.Text = selma104_rs("bed_no")
'        txt_status.Text = "IP"
'    Else
'        txt_status.Text = "OP"
'    End If
  If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_noslides.Text = "" & selpathdrequisitn_rs("LABNO")
            'TXT_PATH.Text = SELPATHDREQUISITN_RS("PATHNO")
            If IsNull(selpathdrequisitn_rs("SITE")) Then
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
                    TXT_CSITE.Text = SELMA301RS("SITEDESC")
                Else
                     TXT_CSITE.Text = ""
                End If
            End If
            
            If selpathdrequisitn_rs.EOF = False Then
                
                txt_typespeciman.Text = " " & selpathdrequisitn_rs("NMETERIALS").Value
            Else
                txt_typespeciman.Text = ""
            End If
           
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
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
        If seldfreportflag = False Then
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
            seldfreportflag = True
        Else
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDFREPORT_RS.MoreResults
            SELDFREPORT_RS.Requery
        End If
                    
        
           If SELDFREPORT_RS.RowCount > 0 Then
               If SELDFREPORT_RS.EOF = False Then
               If IsNull(SELDFREPORT_RS("CYREPORT")) Then
                       txt_report.Text = ""
                       Else
                       SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                        SELDFREPORT_RS.MoreResults
                        SELDFREPORT_RS.Requery
                       txt_report.Text = SELDFREPORT_RS("CYREPORT").Value
                       End If
'
'                   txt_report.Text = SELDFREPORT_RS("CYREPORT").Value
                Else
                   txt_report.Text = ""
                End If
                
                If SELDFREPORT_RS.EOF = False Then
                   txt_impdiag.Text = SELDFREPORT_RS("IMPRESSION").Value
                Else
                   txt_impdiag.Text = ""
                End If
            Else
                'txt_immhischerep.Text = ""
            End If
            
        If REFDR_RS.RowCount > 0 Then
         If SELDFREPORT_RS.EOF = False Then
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
                
        If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = Trim(MaskEdBox1.FormattedText)
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = Trim(MaskEdBox1.FormattedText)
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
        If seldrequisitn_rs.RowCount > 0 Then
         prpathno = ""
'         CMB_PATHNO.clear
         While Not seldrequisitn_rs.EOF
            prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'            CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
            seldrequisitn_rs.MoveNext
         Wend
        End If
            If SELDADDENDUMFLAG = False Then
                SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
                SELDADDENDUMFLAG = True
            Else
                SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
                SELDADDENDUM_RS.MoreResults
                SELDADDENDUM_RS.Requery
            End If
            If SELDADDENDUM_RS.RowCount > 0 Then
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
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
'        txt_clindiag.Enabled = False1
    End If

 
    
    
Exit Sub
'MSG:
'MsgBox Err.Description
End Sub

Private Sub Timer1_Timer()
    txt_date.Text = Format(DATE11, "DD/MM/YYYY")
    txt_time.Text = Format(time1, "HH:MM:SS")
End Sub

Private Sub txt_addendum_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_grossdesc
W_grossdes.Caption = "Addendum"
W_grossdes.WindowState = wdWindowStateMaximize
d_grossdesc.Content = txt_grossdesc
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




Private Sub txt_impdiag_GotFocus()
On Error GoTo m3:
WIMPRE.Visible = False
Set DIMPRE = WIMPRE.ActiveDocument
DIMPRE.Content = ""
DIMPRE.Content = txt_impdiag
WIMPRE.Caption = "Diagnosis"
WIMPRE.WindowState = wdWindowStateMaximize
DIMPRE.Content = txt_impdiag
WIMPRE.Visible = True
WIMPRE.Activate
m3:
On Error Resume Next:
Set WIMPRE = CreateObject("word.Application")
WIMPRE.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
WIMPRE.Caption = "Diagnosis"
WIMPRE.WindowState = wdWindowStateMaximize
Set DIMPRE = WIMPRE.ActiveDocument
DIMPRE.Content = txt_impdiag.Text
WIMPRE.Visible = True
WIMPRE.Activate
End Sub
Private Sub txt_impdiag_LostFocus()
On Error GoTo m4:
Set DIMPRE = WIMPRE.ActiveDocument
txt_impdiag.Text = DIMPRE.Content
'flagimp = False
DIMPRE.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DIMPRE.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DIMPRE.Close (wdDoNotSaveChanges)
WIMPRE.quit
'W1.Visible = False
Exit Sub
m4:
MsgBox Err.Description
End Sub

Private Sub txt_report_GotFocus()
On Error GoTo m3:
WREPORT.Visible = False
Set DREPORT = WREPORT.ActiveDocument
DREPORT.Content = ""
DREPORT.Content = txt_report
WREPORT.Caption = "Report"
WREPORT.WindowState = wdWindowStateMaximize
DREPORT.Content = txt_report
WREPORT.Visible = True
WREPORT.Activate
m3:
On Error Resume Next:
Set WREPORT = CreateObject("word.Application")
WREPORT.Documents.add ("c:\DISREPORTSEC\help.doC")
WREPORT.Caption = "Report"
WREPORT.WindowState = wdWindowStateMaximize
Set DREPORT = WREPORT.ActiveDocument
DREPORT.Content = txt_report.Text
WREPORT.Visible = True
WREPORT.Activate
End Sub
Private Sub txt_report_LostFocus()
On Error GoTo m4:
Set DREPORT = WREPORT.ActiveDocument
txt_report.Text = DREPORT.Content
DREPORT.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DREPORT.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DREPORT.Close (wdDoNotSaveChanges)
WREPORT.quit
Exit Sub
m4:
MsgBox Err.Description
End Sub



Private Sub txt_typespeciman_GotFocus()
On Error GoTo m3:
WMETERIAL.Visible = False
Set DMETERIAL = WMETERIAL.ActiveDocument
DMETERIAL.Content = ""
DMETERIAL.Content = txt_typespeciman
WMETERIAL.Caption = "Report"
WMETERIAL.WindowState = wdWindowStateMaximize
DMETERIAL.Content = txt_typespeciman
WMETERIAL.Visible = True
WMETERIAL.Activate
m3:
On Error Resume Next:
Set WMETERIAL = CreateObject("word.Application")
WMETERIAL.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
WMETERIAL.Caption = "Report"
WMETERIAL.WindowState = wdWindowStateMaximize
Set DMETERIAL = WMETERIAL.ActiveDocument
DMETERIAL.Content = txt_typespeciman.Text
WMETERIAL.Visible = True
WMETERIAL.Activate
End Sub

Private Sub txt_typespeciman_LostFocus()
On Error GoTo m4:
Set DMETERIAL = WMETERIAL.ActiveDocument
txt_typespeciman.Text = DMETERIAL.Content
'flagimp = False
DMETERIAL.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DMETERIAL.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DMETERIAL.Close (wdDoNotSaveChanges)
WMETERIAL.quit
'W1.Visible = False
Exit Sub
m4:
MsgBox Err.Description
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
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
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
