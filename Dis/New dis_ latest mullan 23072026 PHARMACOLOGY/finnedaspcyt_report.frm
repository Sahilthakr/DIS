VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form finnedaspcyt_report 
   Caption         =   "Fine Needle Aspiration Cytology"
   ClientHeight    =   10845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame common 
      ForeColor       =   &H8000000A&
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   11775
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2280
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   72
         Top             =   720
         Width           =   1335
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
            TabIndex        =   74
            Top             =   240
            Width           =   1095
         End
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
            TabIndex        =   73
            Top             =   600
            Width           =   855
         End
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
         TabIndex        =   12
         Top             =   600
         Width           =   1335
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
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_provdiag 
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
         TabIndex        =   26
         Top             =   2520
         Width           =   3855
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   1080
         Width           =   9135
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
         TabIndex        =   5
         Top             =   600
         Width           =   2175
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
         TabIndex        =   3
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
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
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
         TabIndex        =   71
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl_clndis 
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
         Left            =   360
         TabIndex        =   62
         Top             =   2520
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   1080
         Width           =   495
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
         Top             =   360
         Width           =   1830
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10680
      TabIndex        =   46
      Top             =   600
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10680
      TabIndex        =   45
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame COMMAND 
      Height          =   855
      Left            =   0
      TabIndex        =   44
      Top             =   8160
      Width           =   11775
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
         Left            =   7200
         MouseIcon       =   "finnedaspcyt_report.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   35
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
         Left            =   10440
         MouseIcon       =   "finnedaspcyt_report.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   37
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
         Left            =   8760
         MouseIcon       =   "finnedaspcyt_report.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   36
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
         Left            =   5880
         MouseIcon       =   "finnedaspcyt_report.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   34
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
         Left            =   4440
         MouseIcon       =   "finnedaspcyt_report.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   33
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
         Left            =   3120
         MouseIcon       =   "finnedaspcyt_report.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   32
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
         Left            =   1800
         MouseIcon       =   "finnedaspcyt_report.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   31
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
         MouseIcon       =   "finnedaspcyt_report.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame REPORT 
      Height          =   3855
      Left            =   0
      TabIndex        =   27
      Top             =   3960
      Width           =   11775
      Begin VB.TextBox txt_noslideswet 
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
         Left            =   6360
         TabIndex        =   79
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txt_noofslidesdry 
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
         TabIndex        =   77
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cmb_cellblock 
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
         Left            =   9120
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txt_nmeterial 
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
         TabIndex        =   6
         Top             =   600
         Width           =   3855
      End
      Begin VB.ComboBox cmb_prereqno 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7440
         TabIndex        =   29
         Top             =   3240
         Width           =   4095
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   2880
         Width           =   9135
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
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   3240
         Width           =   2775
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
         Left            =   2355
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   3015
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
         Left            =   8640
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2040
         Width           =   2895
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
         Left            =   8595
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txt_sitefnac 
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
         TabIndex        =   28
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txt_impdiag 
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
         TabIndex        =   9
         Top             =   1680
         Width           =   9135
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
         Top             =   1320
         Width           =   9135
      End
      Begin VB.TextBox txt_noofslides 
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
         Left            =   9600
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbl_vpp 
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
         Left            =   6480
         TabIndex        =   81
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_noslideswet 
         AutoSize        =   -1  'True
         Caption         =   "No of Slides Wet"
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
         Left            =   4800
         TabIndex        =   80
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label LBL_SLIDESDRY 
         AutoSize        =   -1  'True
         Caption         =   "No of Slides Dry"
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
         Left            =   840
         TabIndex        =   78
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label LBL_CELLBLOCK 
         AutoSize        =   -1  'True
         Caption         =   "Cell Block"
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
         Left            =   8160
         TabIndex        =   76
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lbl_nmeterial 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Meterial"
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
         TabIndex        =   75
         Top             =   600
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
         Left            =   1200
         TabIndex        =   70
         Top             =   2880
         Width           =   960
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
         Left            =   1080
         TabIndex        =   69
         Top             =   2040
         Width           =   1005
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
         TabIndex        =   68
         Top             =   2040
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
         TabIndex        =   67
         Top             =   2520
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
         Left            =   1035
         TabIndex        =   66
         Top             =   3240
         Width           =   1170
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
         TabIndex        =   43
         Top             =   1800
         Width           =   1890
      End
      Begin VB.Label lbl_report 
         AutoSize        =   -1  'True
         Caption         =   "Microscopy Report"
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
         TabIndex        =   42
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lbl_noofslides 
         AutoSize        =   -1  'True
         Caption         =   "No of Slides"
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
         Left            =   8415
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lbl_prevreqno 
         AutoSize        =   -1  'True
         Caption         =   "Previous RequisitionNos."
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
         TabIndex        =   40
         Top             =   3360
         Width           =   2145
      End
      Begin VB.Label lbl_sitefnac 
         AutoSize        =   -1  'True
         Caption         =   "Site of FNAC"
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
         Left            =   1200
         TabIndex        =   39
         Top             =   240
         Width           =   1110
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
      Left            =   1200
      TabIndex        =   83
      Top             =   0
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
      TabIndex        =   82
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl_finnedaspcyt 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - FINE NEEDLE ASPIRATION CYTOLOGY"
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
      Left            =   2400
      TabIndex        =   65
      Top             =   600
      Width           =   6990
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
      TabIndex        =   64
      Top             =   600
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
      Left            =   10080
      TabIndex        =   63
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "finnedaspcyt_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Option Explicit
Dim str As String
Dim FSFLAG As Boolean
Dim sel_deptqr As rdoQuery
Dim SEL_DEPTRSFLAG As Boolean
Dim UPDSPREPORT_QR As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SEL_DEPTFLAG As Boolean
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim SEL_MODALITYFLAG As Boolean
Dim sel_casno As rdoQuery
Dim req_qr As rdoQuery
Dim REQ_RQD As rdoQuery
Dim REQ_RQD_RS As rdoResultset
Dim req_rs As rdoResultset
Dim REQFLAG As Boolean
Dim disflg As Boolean

Dim selstfname_qr As rdoQuery
Dim selstfname_rs As rdoResultset

Dim updrtimefa_qr As rdoQuery

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
Dim d_frozen As Document, d_imp As Document, d_immun As Document, d_adden As Document ', d As Document , D1 As Document
Dim D1 As Object
Dim d As Object
Dim d_micscop As Object
'Dim d_grossdesc As Document
Dim d_grossdesc As Object
Dim SELEMPLIB_QR As rdoQuery
Dim SELEMPLIB_RS As rdoResultset
Dim I As Integer
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

''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean
Dim VAR_REQNNO As String
''*******ICU**********
''*********FNE QRY************
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
''*********FNE QRY************
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

Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset
Dim SELRMODALITYFLAG As Boolean
''******SECURITY***********

Dim upd_rtime_qr As rdoQuery '01/11/10


Private Sub cmb_cellblock_GotFocus()
cmb_cellblock.clear
cmb_cellblock.AddItem "YES"
cmb_cellblock.AddItem "NO"
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_cellblock) = "" Then
    For ind_cbodept = 0 To (cmb_cellblock.ListCount - 1)
        If Trim(cmb_cellblock.List(ind_cbodept)) = "NO" Then
            cmb_cellblock.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

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
        If sel_modalityrs("modalitycd") = "FA" Then ''******* Add IF Condition on 30.05.2016 at sangrur by Rahul
        
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            End If
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
        If Trim(cmb_dept.List(ind_cbodept)) = "PATHOLOGY$FZZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
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
'            'cmb_grossby.clear
'            While Not seldissecgross_rs.EOF
'                'cmb_grossby.AddItem Trim(seldissecgross_rs("CCNAME"))
'                seldissecgross_rs.MoveNext
'            Wend
'        End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "FNAC$FA" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_prereqno_Click()
X = InStr(1, Trim(cmb_prereqno.Text), "$")
REQNO1 = Mid(Trim(cmb_prereqno.Text), 1, X - 1)
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

Private Sub cmd_back_Click()
'btnenabl ("INI")
BTNENABL1 ("INI")
FLDENABL ("DIS")
TXT_CLEAR

If WORKLIST = True Then
    Me.Hide
    disw.show
Else
    'Unload Me
    Me.Hide
    DIS_MAIN.show
    
End If

'DIS_MAIN.show
'Me.Hide
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
End Sub

Private Sub cmd_commit_Click()
Dim I As Integer
I = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If I = vbYes Then
    On Error GoTo MSG:
    ''********SECURITY***************************
    
    '******* add from here for give authority only for usrtype >=10 sangrur 18.07.2017 at sangrur
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
        MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
    End If
''********SECURITY***************************

    rdb.BeginTrans
    updrrequisitn1_qr(0) = Left(cmb_finalby.Text, 30)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    
    updrtimefa_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrtimefa_qr(1) = time1
    updrtimefa_qr(2) = req_chglen(cmb_reqno.Text)
    updrtimefa_qr.Execute
    

rdb.CommitTrans
'btnenabl ("INI")
BTNENABL1 ("INI")
MsgBox "Record is Committed Successfully."
cmd_print.Enabled = True
Else
'MsgBox "Record is NOT Committed Successfully."
Exit Sub
End If

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
D1.Content = D1.Content.Text & "ADDENDUM:" & Chr(13) & Trim(txt_addendum)
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
'btnenabl ("NEW")
''********SECURITY***************************
CURUSRDET_QR(0) = cmb_enteryby.List(0) 'Trim(cmb_enteryby.Text)'SHRADDHA 12/04/11
CURUSRDET_RS.MoreResults
CURUSRDET_RS.Requery
If CURUSRDET_RS.RowCount > 0 Then
    If LOGINTYP > CURUSRDET_RS("usrtype") Then
      ''   no action as he can commit
    ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
        'MsgBox "Sorry! This Report needs is to Modified by a Higher Level than  that of Yours"
        'cmd_modify.SetFocus
        'Exit Sub
    ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
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
    MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
    'cmd_commit.SetFocus
    Exit Sub
End If

''********SECURITY***************************

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
    
    FLDENABL ("ADDENDUM")
    txt_addendum.SetFocus
    'flagsave = 0
End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
'On Error GoTo MSG:
'btnenabl ("NEW")
BTNENABL1 ("NEW")
field_clear
TXT_CLEAR
txt_date.Text = DATE11
txt_time.Text = time1
cmb_enteryby.Text = cmb_enteryby.List(0)
cmb_enteryby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enteryby.List(0)) = cmb_finalby.List(L1) Then
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
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
        Wend
   End If
If SELMA302FLAG = False Then
   Set selma302_rs = selma302_qr.OpenResultset(2, 3)
   SELMA302FLAG = True
Else
    selma302_rs.MoreResults
    selma302_rs.Requery
End If

If selma302_rs.RowCount > 0 Then
            selma302_rs.MoveFirst
            While Not selma302_rs.EOF
                'CMB_ICDH.AddItem Trim(selma302_rs("histdesc"))
                selma302_rs.MoveNext
            Wend
        End If
flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
Opt_patient.Value = True
MaskEdBox1.Enabled = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
 On Error GoTo prtmsg:

DoEvents




VAR_REQNNO = req_chglen(cmb_reqno.Text)

  ''' added for print emr report for sangrur by rahul 22/01/2016 as Dr Sankalp wanted to direct link to emr for printing report


 strLink = "http://10.80.1.9/EMRREPORTS/allreports.aspx?reqno=" & VAR_REQNNO & "  ~GUEST"
 
 
    'strLink = "http://10.80.1.9/EMRTMHACT/DEFAULT.ASPX?QPNO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink, vbMaximizedFocus
Exit Sub
'' till there by rahul 22/01/2015


        
        ''*******icu****************
'   If opt_patient.Value = True Then
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
 str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
If Opt_patient.Value = True Then
    If ADDENDUMFALG = False Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        If Trim(txt_report.Text) = "" Then
            CrystalReport1.Formulas(8) = "fn = '" & "" & "'"
        Else
            CrystalReport1.Formulas(8) = "fn = '" & "12" & "'"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
            If Trim(SELRREQNORS("REQNSTAT")) = "F" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT WITH ADDENDUM " & "'"
                str1 = ""
                'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
                 str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             'Else
    
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY  REPORT" & "'"
                'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
                'CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
            End If
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        
        If Trim(txt_report.Text) = "" Then
            CrystalReport1.Formulas(8) = "fn = '" & "" & "'"
        Else
            CrystalReport1.Formulas(8) = "fn = '" & "12" & "'"
        End If
    End If
Else
    If SELEMPBASINFOFLAG = False Then
        SELEMPBASINFO_QR(0) = MaskEdBox1.FormattedText
        Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
        SELEMPBASINFOFLAG = True
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.FormattedText
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
    End If
    If ADDENDUMFALG = False Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        If Trim(txt_report.Text) = "" Then
            CrystalReport1.Formulas(8) = "fn = '" & "" & "'"
        Else
            CrystalReport1.Formulas(8) = "fn = '" & "12" & "'"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "F" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT WITH ADDENDUM" & "'"
                str1 = ""
                'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
                 str1 = "Entered By " & Trim(cmb_enteryby.Text)
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             'Else
    
                'CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY  REPORT" & "'"
                'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
                'CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
            End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    End If
    
    selstfname_qr(0) = req_chglen(cmb_reqno.Text)
    Set selstfname_rs = selstfname_qr.OpenResultset(2, 3)
    CrystalReport1.Formulas(10) = "NAM = '" & Trim(selstfname_rs("PNAME")) & "'"
    'CrystalReport1.Formulas(10) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
    CrystalReport1.Formulas(11) = "STDOB = '" & selstfname_rs("sex") & "'"
    'CrystalReport1.Formulas(11) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
    CrystalReport1.Formulas(12) = "DOB = '" & selstfname_rs("age") & "'"
    'CrystalReport1.Formulas(12) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
End If
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=TMH;PWD= TMH;DSN=DISDSN"
CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("print")
BTNENABL1 ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_save_Click()
'On Error GoTo MSG:
 Dim beforst As String, afterst As String
Dim m As Integer
If Len(Trim(cmb_reqno)) = 0 Then
    m = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
    cmb_reqno.SetFocus
    Exit Sub
End If
    
If Len(Trim(cmb_enteryby.Text)) = 0 Then
    m = MsgBox("The Entered By field is EMPTY.", vbCritical)
    cmb_enteryby.SetFocus
    Exit Sub
End If
If Len(Trim(cmb_provby.Text)) = 0 Then
    m = MsgBox("The Provisional By field is EMPTY.", vbCritical)
    cmb_provby.SetFocus
    Exit Sub
End If
If Len(Trim(cmb_finalby.Text)) = 0 Then
    m = MsgBox("The To be Finalized field is EMPTY.", vbCritical)
    cmb_finalby.SetFocus
    Exit Sub
End If
    
        While InStr(txt_report, Chr(9)) > 0
       
            Dim K As Long
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
        
            'Dim K As Long
            K = InStr(txt_impdiag, Chr(9))
            beforst = Mid(txt_impdiag, 1, K - 1) & "      "
            afterst = Mid(txt_impdiag, K + 1)
            txt_impdiag = beforst & afterst
       Wend
      While InStr(txt_impdiag, Chr(146)) > 0

            K = InStr(txt_impdiag, Chr(146))
            beforst = Mid(txt_impdiag, 1, K - 1) & "`"
            afterst = Mid(txt_impdiag, K + 1)
            txt_impdiag = beforst & afterst
       
       Wend
       While InStr(txt_addendum, Chr(9)) > 0
        
            'Dim K As Long
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
  If Len(Trim(txt_nmeterial.Text)) > 1000 Then
    MsgBox "Nature of Material Max. of 1000 Characters Only.", vbInformation
    txt_nmeterial.SetFocus
    Exit Sub
  End If
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
    MsgBox "Please Enter Impression", vbInformation
    txt_impdiag.SetFocus
    Exit Sub
  End If
  If Len(Trim(txt_addendum.Text)) > 600 Then
    MsgBox "Addendum Max. of 600 Characters Only.", vbInformation
    txt_addendum.SetFocus
    Exit Sub
  End If
      
    If ADDENDUMFALG = True Then
        If Len(Trim(cmb_addendumby.Text)) = 0 Then
              MsgBox "The Addendum by field is EMPTY.", vbCritical
              cmb_addendumby.Enabled = True
              cmb_addendumby.SetFocus
              Exit Sub
        End If
    End If
'rdb.BeginTrans
       assign
'rdb.CommitTrans
       'MsgBox "Record is saved Successfully."
  FLDENABL ("DIS")
  'btnenabl ("SAVE")
  BTNENABL1 ("SAVE")
  Exit Sub
MSG:
rdb.RollbackTrans
MsgBox ERROR(Err)
End Sub

Public Sub cmd_view_Click()
On Error GoTo MSG:
'btnenabl ("VIEW")
BTNENABL1 ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR
txt_date.Text = DATE11
txt_time.Text = time1
'&&&&&&&&&&&&&&&&&&&&77 not found hence declared

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
   If SELMA302FLAG = False Then
   Set selma302_rs = selma302_qr.OpenResultset(2, 3)
   SELMA302FLAG = True
Else
    selma302_rs.MoreResults
    selma302_rs.Requery
End If
        If selma302_rs.RowCount > 0 Then
            selma302_rs.MoveFirst
            While Not selma302_rs.EOF
                selma302_rs.MoveNext
            Wend
        End If
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
cmb_dept.SetFocus
cmd_print.Enabled = True
Opt_patient.Value = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
'btnenabl ("INI")
'BTNENABL1 ("INI")
NEWFLAG = False
viewflag = False
MaskEdBox1.Enabled = False

If WORKLIST = False Then
    MAKEFALSE
End If

If FAREPORTFLAG = False Then
    FAREPORTFLAG = True
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1FNAC", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    'cmb_enteryby.Text = (LOGINUSRNAME)
    cmb_enteryby.clear
    cmb_enteryby.AddItem LOGINUSRNAME
    cmb_enteryby.Text = cmb_enteryby.List(0)
    cmb_addendumby.clear
    cmb_addendumby.AddItem LOGINUSRNAME
Else
    MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNYY_STR", str)
If Trim(USRDEPT) = "ALL" Then
    REPWTRN_QR(0) = "FZZ"
Else
    REPWTRN_QR(0) = Trim(USRDEPT)
End If
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    'cmb_addendumby.clear
    cmb_finalby.clear
    cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 9 Then 'mADE 8 FROM 1
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        If REPWTRNRS("USRTYPE") > 3 And (REPWTRNRS("USRTYPE") <> 6) Then
            cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        REPWTRNRS.MoveNext
    Wend
    
Else
    MsgBox "The Security Details of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"

Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11FNAC", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

'str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ? and dept_code='FZZ'"   ''Commented for Sangrur Rahul 15/02/2016
str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE usrtype >= 9  or (USRTYPE >= ? and dept_code='FZZ')"
Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1FNAC", str)
SELFINALBY_QR(0) = 8    'MADE 8 FROM 5
Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'MsgBox SELFINALBY_RS.RowCount
cmb_finalby.clear
If Not SELFINALBY_RS.EOF Then
    Do While Not SELFINALBY_RS.EOF
       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
       SELFINALBY_RS.MoveNext
    Loop
End If


'*********SECURITY**********************************

''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR1161", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC%' OR PAT_CATG<>'SF%')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR1161", str)


''********ICU***************

str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr11 = rdb.CreateQuery("seldrequisitn11_1qrfnac", str)
     

'*****FNE QRY***********
str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%') ORDER BY PREPORTDT"
Set seldrequisitn_qr10 = rdb.CreateQuery("seldrequisitn11_1qr110f", str)
'*****FNE QRY***********

'*********SECURITY**********************************

'''''STR = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
'''''Set USRNAME_QR = rdb.CreateQuery("USRNAME11113", STR)
'''''USRNAME_QR(0) = UID
'''''Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
'''''If USRNAME_RS.RowCount > 0 Then
'''''    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
'''''    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
'''''    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
'''''    cmb_enteryby.AddItem LOGINUSRNAME
'''''Else
'''''    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
'''''    cmd_cancel_Click
'''''    Exit Sub
'''''End If
'''''
''''''Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'''''Set REPWTRNRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE = 'FZZ' ORDER BY USERNAME", 1, 1)
'''''
'''''If REPWTRNRS.RowCount > 0 Then
'''''    'cmb_enteryby.clear
'''''    cmb_addendumby.clear
'''''    cmb_finalby.clear
'''''    cmb_provby.clear
'''''    REPWTRNRS.MoveFirst
'''''    While Not REPWTRNRS.EOF
'''''        If REPWTRNRS("USRTYPE") > 1 Then
'''''            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
'''''            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
'''''            'AGAIN
'''''             cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
'''''            cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
'''''        End If
'''''        REPWTRNRS.MoveNext
'''''    Wend
'''''Else
'''''    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
'''''    cmd_cancel_Click
'''''    Exit Sub
'''''End If
'''''STR = "select * from DISLIB.Dislnkusr where username = ?"
'''''Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq1107", STR)
'''''CURUSRDET_QR(0) = "ABC"
'''''Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'''''
'''''STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'''''Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR2", STR)
'''''SELFINALBY_QR(0) = 6
'''''Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'''''If Not SELFINALBY_RS.EOF Then
'''''    Do While Not SELFINALBY_RS.EOF
'''''       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'''''       SELFINALBY_RS.MoveNext
'''''    Loop
'''''End If

'*********SECURITY**********************************


'for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel1_deptqr11FA", str)
'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
   
    'for icd-h retriveing
str = "select * from patoolslib.ma302"
Set selma302_qr = rdb.CreateQuery("selma302_qr1FA", str)
'Set selma302_rs = selma302_qr.OpenResultset(2, 3)
        
str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301QR = rdb.CreateQuery("SELMA3011QR1FA", str)
'FOR GETTING CLINICAL DIAGNOSIS
str = "select CLINHIST  from rislib.CLINHIST WHERE CASENO=? "
Set SELCLIN_QR = rdb.CreateQuery("SELCLIN_QRFA1", str)
        
' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqr11FA", str)

str = "select * from rislib.rmodality where dept_code=?"
Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr11f", str)
    
'STR = "select * from rislib.rmodality where dept_code=?"
'Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr11", STR)
    
str = "select * from dislib.dreqstaffd where reqno=?"
Set selstfname_qr = rdb.CreateQuery("selstfname_qr", str)


    'for retriving reqnos
Set req_qr = rdb.CreateQuery("1reqFA", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT DESC")
Set REQ_RQD = rdb.CreateQuery("REQ_rqd", "select reqnno,reqndt from rislib.rrequisitn, dislib.drequisitn where RREQUISITN.REQNNO = DREQUISITN.REQNO AND RREQUISITN.CASENO = ? AND RREQUISITN.MODALITYCD = ? and DREQUISITN.icds2 = 'YES' and RREQUISITN.REQNSTAT = 'M'  ORDER BY REQNDT DESC")

' SELECT RREQUISITN.REQNNO, RREQUISITN.REQNDT FROM RISLIB/RREQUISITN,
' DISLIB/DREQUISITN WHERE RREQUISITN.REQNNO = DREQUISITN.REQNO AND
' RREQUISITN.CASENO = 'TT/00500' AND RREQUISITN.MODALITYCD = 'FA' AND
' RREQUISITN.REQNSTAT = 'M' AND RREQUISITN.REQNNO = 'FZZFA03008063'
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

Set SELTA101_QR = rdb.CreateQuery("selta101_qr11FA", "select * from PATOOLSLIB.ta101 where case_no =?")
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
     Set REFDR_QR = rdb.CreateQuery("REFDR1_QR1FA", str)
'    REFDR_QR(0) = "BN/00001"
'    REFDR_QR(1) = "FZZSP010000"
'    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
'    FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1_QR1FA", str)
    'SELPATHDREQUISITN_QR(0) = "SSSSS"
    'SELPATHDREQUISITN_QR(1) = "FZZSP010000"
    'Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(2, 3)

    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma1041_qr11FA", str)
    'Selma104_qr(0) = "BN/00001"
    'Set selma104_rs = selma104_qr.OpenResultset(2, 3)
    'for provisional diagnosis

    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("se1lta10121_qr1FA", str)
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
    Set seldimpresion_qr = rdb.CreateQuery("seldimpresion_qr11FA", str)
'    seldimpresion_qr(0) = "zzzzzz"
'    Set seldimpresion_rs = seldimpresion_qr.OpenResultset(2, 3)

    str = "select * from dislib.dgrossdesc where reqno=?"
    Set seldgrossdesc_qr = rdb.CreateQuery("seldgrossdesc_qr1FA", str)
'    seldgrossdesc_qr(0) = "zzzzzz"
'    Set seldgrossdesc_rs = seldgrossdesc_qr.OpenResultset(2, 3)


'    STR = "select * from dislib.dspreport wherE reqno=?"
'    Set seldspreport_qr = rdb.CreateQuery("seldspreport_qr11", STR)
'    seldspreport_qr(0) = "zzzzzz"
'    Set seldspreport_rs = seldspreport_qr.OpenResultset(2, 3)



    ' for insertion
'    STR = "insert into  dislib.dfsreport (reqno, fsreport,enterby,preportby,freportby) values (?,?,?,?,?)"
'    Set insdfs_qr = rdb.CreateQuery("insdfs_qr1FA", STR)
'    STR = "insert into dislib.dgrossdesc (reqno,grossdesc,grossedby) values (?,?,?)"
'    Set insdgrossdesc_qr = rdb.CreateQuery("insdgrossdesc_qr1FA", STR)
'    STR = "insert into dislib.dihreport (reqno,ihreport,enteryby,preportby,freportby) values (?,?,?,?,?)"
'    Set INSDIHREPORT_QR = rdb.CreateQuery("insdihreport_qr1FA", STR)
'    STR = "insert into dislib.dimpresion (reqno,finaldiag,fdescry,t,tsize,n,m,gradecode) values (?,?,?,?,?,?,?,?)"
'    Set insdimpresion_qr = rdb.CreateQuery("1insdimpresion_qrFA", STR)
'    STR = "insert into dislib.dspreport (reqno,spreport,enterby,preportby,freportby) values(?,?,?,?,?)"
'    Set insdspreport_qr = rdb.CreateQuery("insdspreport_qr1FA", STR)

'    STR = "update dislib.DFNREPORT set MICREPORT=? IMPRESSION=? where reqno=?"
'    Set upddfsreport_qr = rdb.CreateQuery("upddfsreport_qr1FA", STR)
    
'    STR = "update dislib.dihreport set ihreport=?, enteryby=?, preportby=?, freportby=? where reqno=?"
'    Set upddihreport_qr = rdb.CreateQuery("upddihreport1_qr11FA", STR)
'
'    STR = "update dislib.dimpresion set  finaldiag=?, fdescry=?, t=?,tsize=?, n=?, m=?,GRADECODE=? where reqno=?"
'    Set upddimpresion_qr = rdb.CreateQuery("upddimpresion_qr1FA", STR)
'
'    STR = "UPDATE DISLIB.DSPREPORT SET SPREPORT=?, ENTERBY=?, PREPORTBY=?, FREPORTBY=? WHERE REQNO=?"
'     STR = "UPDATE DISLIB.DSPREPORT SET SPREPORT=? WHERE REQNO=?"
'    Set UPDSPREPORT_QR = rdb.CreateQuery("UPDSPREPORT_QR1FA", STR)
'
'    STR = "UPDATE DISLIB.DGROSSDESC SET GROSSDESC=?,GROSSEDBY=? WHERE REQNO=?"
'    Set updgrossdesc_qr = rdb.CreateQuery("updgrossdesc_qr1FA", STR)
'
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR11FA", str)

    str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
    Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1FA", str)
   
    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qr1FA", str)
    
    str = "INSERT INTO DISLIB.ADDENDUMFA(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
    Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1FA", str)
   
    str = "SELECT  * FROM DISLIB.ADDENDUMFA WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11FA", str)
    SELDADDENDUM_QR(0) = "FSS"
    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
   
    str = "select * from dislib.drequisitn where caseno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_1qr1FA", str)

'   Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)

    txt_addendum.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_dept.Enabled = False
    FLDENABL ("DIS")


   str = "INSERT INTO DISLIB.DFNREPORT(REQNO,micreport,IMPRESSION) VALUES (?,?,?)"
   Set INSFREPORT_QR = rdb.CreateQuery("INSFREPORT_QR1FA", str)

   str = "UPDATE DISLIB.dfnreport SET micreport=?,IMPRESSION=? WHERE REQNO=?"
   Set upddfsreport_qr = rdb.CreateQuery("upddfsreport_qr11FA", str)
    
   str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
   Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQR1FA", str)
'   SELRREQNOQR(0) = "FZZSP0200000"
'   Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    
    str = "SELECT * FROM DISLIB.DFNREPORT WHERE REQNO=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QR1FA", str)
'    SELDFREPORT_QR(0) = "FZZSP020000"
'    Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
    
    str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_IHQR = rdb.CreateQuery("SELDFREPORT_IHQR1FA", str)
'    SELDFREPORT_IHQR(0) = "FZZSP0200001"
'    Set SELDFREPORT_IHRS = SELDFREPORT_IHQR.OpenResultset(1, 3)
    
    str = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?,FROZENSEC=?,NMETERIALS=? WHERE REQNO=?"
    Set upddrequisitn_qr = rdb.CreateQuery("UPDSPREPORT1_QR1FA", str)
    
    str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO11_QR1101", str)

'    SELEMPBASINFO_QR(0) = "00000"
'    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    
'    STR = "select * from dislib.dgrade"
'    Set seldgradeqr = rdb.CreateQuery("seldgradeqr1", STR)
'    Set seldgraders = seldgradeqr.OpenResultset(1, 3)
'    If seldgraders.RowCount > 0 Then
'    'cmb_grade.clear
'    While Not seldgraders.EOF
'        'cmb_grade.AddItem Trim(seldgraders("gradedesc")) & "$" & Trim(seldgraders("gradecode"))
'        seldgraders.MoveNext
'    Wend
'    End If
    
    str = "select * from dislib.dgrade where gradecode=?"
    Set seldgrade_qr = rdb.CreateQuery("seldgrade_qr1FA", str)
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
    Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQR1FA", str)
    
   str = "select * from dislib.dihreport where reqno=?"
   Set selihreportqr = rdb.CreateQuery("selihreportqr1FA", str)
   
   str = "select * from dislib.dissec where enterby='YES'AND modality='SP'"
   Set seldissecentr_qr = rdb.CreateQuery("seldissecentr_qr1FA", str)
   
   
   str = "update rislib.rtimelog set comrepdate=?, comreptime=? where reqnno=?"
   Set updrtimefa_qr = rdb.CreateQuery("updrtimefa_qr", str)
   
   '***** Shraddha 01/11/10 *****
   str = "update rislib.rtimelog set SAVREPDATE=?, SAVREPTIME=? where reqnno=?"
   Set upd_rtime_qr = rdb.CreateQuery("", str)
   '*****************************
   
   str = "select * from dislib.dissec where PROVby='YES'AND MODALITY='SP'"
   Set seldissecprov_qr = rdb.CreateQuery("seldissecPROV_qr1FA", str)
   
   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='SP'"
   Set seldissecfin_qr = rdb.CreateQuery("seldissecFIN_qr1FA", str)
   
   str = "select * from dislib.dissec where ADDENDUMby='YES'AND MODALITY='SP'"
   Set seldissecadd_qr = rdb.CreateQuery("seldissecADD_qr1FA", str)
   
   str = "select * from dislib.dissec where GROSSby='YES'"
   Set seldissecgross_qr = rdb.CreateQuery("seldissecGROSS_qr1FA", str)
   
   str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
   Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1QR11FA", str)
   
   str = "select sitecd,sitedesc from dislib.sitedup where sitecd = ?"
  Set viewsitedupqr = rdb.CreateQuery("viewsitedupqr1FA", str)
  
   str = "INSERT INTO RISLIB.DISIMP(REQNO,IMP)VALUES(?,?)"
  Set INSDISIMP_QR = rdb.CreateQuery("INSDISIMP_QRFA", str)
  
  str = "UPDATE RISLIB.DISIMP SET IMP=? WHERE REQNO=?"
  Set UPDDISIMP_QR = rdb.CreateQuery("UPDDISIMP_QRFA", str)
  
   
 End If
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



Public Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
 
 CASENO = MaskEdBox1.Text
 If Opt_patient.Value = True Then
 CASENO = MaskEdBox1.Text
 ''MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 02/12/2022
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
 
 End If
 Call FNEQuery
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
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
                  req_rs.MoveFirst
                  While Not req_rs.EOF
                  cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
                  
                  req_rs.MoveNext
                  
                  Wend
            Else
               cmb_reqno.clear
              
            End If
                 '''
           ' If disflg = False Then
          '      REQ_RQD(0) = MaskEdBox1.FormattedText
            '    REQ_RQD(1) = Trim(txt_modcode.Text)
             '   Set REQ_RQD_RS = req_qr.OpenResultset(1, 3)
             '   REQFLAG = True
            'Else
             '   REQ_RQD(0) = MaskEdBox1.FormattedText
             '   REQ_RQD(1) = Trim(txt_modcode.Text)
             '   REQ_RQD_RS.MoreResults
             '   REQ_RQD_RS.Requery
            'End If
             '    If REQ_RQD_RS.RowCount > 0 Then
              '    'cmb_reqno.clear
               '   REQ_RQD_RS.MoveFirst
                '  While Not REQ_RQD_RS.EOF
                '  cmb_reqno.AddItem req_chglen(REQ_RQD_RS("reqNno"))
                '  REQ_RQD_RS.MoveNext
                  
                '  Wend
                ' End If
                 
                 '''
                      
       
        'If (req_rs.RowCount = 0) And (REQ_RQD_RS.RowCount = 0) Then
        '     cmb_reqno.clear
        'End If
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
 Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Public Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
'On Error GoTo MSG:
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
            lbl_vpp.Caption = "VPP"
                If Trim(REFDR_RS("recptno")) = "VPP" Then
                    MsgBox "Payment is not done,You are not authorised to generate report."
                    'btnenabl ("INI")
                    BTNENABL1 ("INI")
                    field_clear
                    Exit Sub
                End If
            End If
            If REFDR_RS("recordno") = "2" Then
            lbl_vpp.Caption = "LR"
            End If
            If REFDR_RS("recordno") = "3" Then
            lbl_vpp.Caption = "BAR"
            End If
            If REFDR_RS("recordno") = "4" Then
            lbl_vpp.Caption = "DrOp"
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
                txt_nmeterial.Text = " " & selpathdrequisitn_rs("nmeterials").Value
            Else
                txt_nmeterial.Text = ""
            End If
            If selpathdrequisitn_rs.EOF = False Then
                cmb_cellblock.Text = " " & selpathdrequisitn_rs("frozensec")
            Else
                cmb_cellblock.Text = " "
            End If
            txt_noslideswet.Text = selpathdrequisitn_rs("outblock")
            txt_noofslidesdry.Text = selpathdrequisitn_rs("outsspecim")
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
                    txt_sitefnac.Text = SELMA301RS("SITEDESC")
                    'txt_noofslides.Text = SELMA301RS("OUTSLIDES")
                    Else
                     txt_sitefnac.Text = ""
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
            txt_provdiag.Text = selta1012_rs("prvdiagdes")
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
         seldrequisitn_rs.MoveNext
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
'

txt_sitefnac.Enabled = True
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
                'ADDENDUMFALG = True
            End If
            If REFDR_RS("recordno") = "1" Then
            lbl_vpp.Caption = "VPP"
            End If
            If REFDR_RS("recordno") = "2" Then
            lbl_vpp.Caption = "LR"
            End If
            If REFDR_RS("recordno") = "3" Then
            lbl_vpp.Caption = "BAR"
            End If
            If REFDR_RS("recordno") = "4" Then
            lbl_vpp.Caption = "DrOp"
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
                txt_cytology.Text = Trim(selpathdrequisitn_rs("PATHNO"))
                If selpathdrequisitn_rs.EOF = False Then
                    txt_nmeterial.Text = " " & selpathdrequisitn_rs("nmeterials")
                Else
                    txt_nmeterial.Text = ""
                End If
                cmb_cellblock.Text = selpathdrequisitn_rs("frozensec")
                txt_noslideswet.Text = selpathdrequisitn_rs("outblock")
                txt_noofslidesdry.Text = selpathdrequisitn_rs("outsspecim")
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
                    txt_sitefnac.Text = SELMA301RS("SITEDESC")
                    'venkat******************
                    Else
                      txt_sitefnac.Text = ""
                End If
            End If
            
            If IsNull(Trim(selpathdrequisitn_rs("NMETERIALS"))) Then
                txt_nmeterial.Text = ""
            Else
              txt_nmeterial.Text = Trim(selpathdrequisitn_rs("nmeterials"))
            End If
            
            
'            SELDREQNOQR(0) = req_chglen(cmb_reqno.Text)
'            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'            If SELDREQNORS.RowCount > 0 Then
'
'                    'txt_immhischereqno.Text = req_chglen((SELDREQNORS("REQNO")))
'                    'cmb_immhische.Text = "YES"
'
'                Else
'                    'txt_immhischerep.Text = ""
'                    'cmb_immhische.Text = "NO"
'
'            End If
            
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
            txt_provdiag.Text = selta1012_rs("prvdiagdes")
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
                txt_report.Text = SELDFREPORT_RS("MICREPORT").Value
             Else
                txt_report.Text = ""
             End If
             If SELDFREPORT_RS.EOF = False Then
                txt_impdiag.Text = SELDFREPORT_RS("IMPRESSION").Value
             Else
                txt_impdiag.Text = ""
             End If
         Else
            'txt_frosecrep.Text = ""
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
'            txt_prepathreqno.Text = Trim(prpathno)
        End If
            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
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
    ' MaskEdBox1.Mask = "#######"            ' COMMENTED BY VARINDER
End If
End Sub



Private Sub txt_fullreport_GotFocus()
On Error GoTo m14:
w.Visible = False
Set d = w.ActiveDocument
d.Content.Text = ""
d.Content.Text = txt_fullreport.Text
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
d.Content.Text = txt_fullreport.Text
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
d.Content.Text = txt_fullreport.Text
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

Public Sub save()

End Sub

Public Sub FNEQuery()
      'STR = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%') ORDER BY PREPORTDT"
      
      
If seldrequisitn_flag10 = False Then
    seldrequisitn_qr10(0) = UCase(MaskEdBox1.FormattedText)
    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
    seldrequisitn_FLAG = True
Else
    seldrequisitn_qr10(0) = UCase(MaskEdBox1.FormattedText)
    seldrequisitn_rs10.MoreResults
    seldrequisitn_rs10.Requery
End If
If seldrequisitn_rs10.RowCount > 0 Then
    cmb_prereqno.clear
    While Not seldrequisitn_rs10.EOF
        'CMB_PREINVEST.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & req_chglen(seldrequisitn_rs10("REQNNO"))
        cmb_prereqno.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
                                '" $ " & Trim(MODDESC) &
        seldrequisitn_rs10.MoveNext
    Wend
    cmb_prereqno.Text = cmb_prereqno.List(0)
End If

End Sub
Public Sub assign()
'On Error GoTo MSG:
Dim I As Integer
Dim MICROSCOPIC As String
rdb.BeginTrans
If flagsave = 1 Then
   flagsave = 0
   
   INSFREPORT_QR(0) = Trim(req_chglen(cmb_reqno.Text))
   If Len(Trim(txt_report.Text)) > 0 Then
    INSFREPORT_QR(1) = Null
    INSFREPORT_QR(1) = Trim(txt_report.Text)
   Else
    INSFREPORT_QR(1) = Null
   INSFREPORT_QR(1) = " "
   End If
   If Len(Trim(txt_impdiag.Text)) > 0 Then
    INSFREPORT_QR(2) = Null
    INSFREPORT_QR(2) = Trim(txt_impdiag.Text)
   Else
    INSFREPORT_QR(2) = Null
    INSFREPORT_QR(2) = " "
   End If
   INSFREPORT_QR.Execute

    '**** INSERTING INTOFREPORT
    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    
    upddrequisitn_qr(0) = ""
    upddrequisitn_qr(1) = cmb_cellblock.Text
    If Len(Trim(txt_nmeterial.Text)) > 0 Then
        upddrequisitn_qr(2) = Null
        upddrequisitn_qr(2) = txt_nmeterial.Text
    Else
        upddrequisitn_qr(2) = Null
        upddrequisitn_qr(2) = " "
    End If
    upddrequisitn_qr(3) = req_chglen(cmb_reqno.Text)
    upddrequisitn_qr.Execute
    
    '***** Update RTIMELOG for savedt, savetm
    DATETIME
    upd_rtime_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_rtime_qr(1) = time1
    upd_rtime_qr(2) = req_chglen(cmb_reqno.Text)
    upd_rtime_qr.Execute
    '*******************************************
    
    
    '**** INSERTING INTO DISIMP
        If Len(Trim(txt_impdiag.Text)) > 0 Then
        INSDISIMP_QR(0) = req_chglen(cmb_reqno.Text)
        INSDISIMP_QR(1) = Null
        INSDISIMP_QR(1) = Trim(txt_impdiag.Text)
        INSDISIMP_QR.Execute
        End If
  '**** INSERTING INTO DISIMP
    
End If
      

If flagsave = 2 Then
    flagsave = 0
    If Len(Trim(txt_report.Text)) > 0 Then
        upddfsreport_qr(0) = Null
        upddfsreport_qr(0) = txt_report.Text
    Else
        upddfsreport_qr(0) = Null
        upddfsreport_qr(0) = " "
    End If
    If Len(Trim(txt_impdiag.Text)) > 0 Then
        upddfsreport_qr(1) = Null
        upddfsreport_qr(1) = txt_impdiag.Text
    Else
        upddfsreport_qr(1) = Null
        upddfsreport_qr(1) = " "
    End If
    upddfsreport_qr(2) = Trim(req_chglen(cmb_reqno.Text))
    upddfsreport_qr.Execute

    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provby.Text
    
    If ADDENDUMFALG = True Then
        updrrequisitn_qr(3) = "F"
    Else
        updrrequisitn_qr(3) = "T"
    End If
    
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute

    upddrequisitn_qr(0) = ""
    upddrequisitn_qr(1) = cmb_cellblock.Text
    If Len(Trim(txt_nmeterial.Text)) > 0 Then
        upddrequisitn_qr(2) = Null
        upddrequisitn_qr(2) = txt_nmeterial.Text
    Else
        upddrequisitn_qr(2) = Null
        upddrequisitn_qr(2) = " "
    End If
    upddrequisitn_qr(3) = req_chglen(cmb_reqno.Text)
    upddrequisitn_qr.Execute

    '***** Update RTIMELOG for savedt, savetm
    DATETIME
    upd_rtime_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_rtime_qr(1) = time1
    upd_rtime_qr(2) = req_chglen(cmb_reqno.Text)
    upd_rtime_qr.Execute
    '*******************************************

    '**** updating INTO FREPORT
    
    '**** UPDATE THE DISIMP
    If Len(Trim(txt_impdiag.Text)) > 0 Then
        UPDDISIMP_QR(0) = Null
        UPDDISIMP_QR(0) = Trim(txt_impdiag.Text)
        UPDDISIMP_QR(1) = req_chglen(cmb_reqno.Text)
        UPDDISIMP_QR.Execute
    End If
        
    '*** UPDATE THE DISIMP
    
    
    
    
    If ADDENDUMFALG = True Then
            ''********SECURITY***************************
        'CURUSRDET_QR(0) = Trim(cmb_finalby.Text)  'comment for sangrur as dr type=10  on 15.07.2016
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
            MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
            cmd_commit.SetFocus
            Exit Sub
        End If
''********SECURITY***************************
        
'        If Len(Trim(cmb_addendumby.Text)) = 0 Then
'            MsgBox "The Addendum by field is EMPTY.", vbCritical
'            cmb_addendumby.Enabled = True
'            cmb_addendumby.SetFocus
'            Exit Sub
'        End If
        
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
lblicu.Visible = False
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
cmb_reqno.ListIndex = -1
cmb_reqno.clear
'cmb_reportdoc.ListIndex = -1
cmb_addendumby.ListIndex = -1
cmb_addendumby.Text = ""
cmb_enteryby.ListIndex = -1
cmb_enteryby.Text = ""
cmb_provby.ListIndex = -1
cmb_provby.Text = ""
cmb_finalby.ListIndex = -1
cmb_finalby.Text = ""
'cmb_grade.ListIndex = -1
cmb_cellblock.Text = ""
lbl_vpp.Caption = ""
'cmb_immhische.ListIndex = -1
'cmb_immhische.Text = ""
'cmb_immhische.clear
'cmb_t.ListIndex = -1
'cmb_t.clear
'CMB_N.clear
'CMB_N.ListIndex = -1
'cmb_m.ListIndex = -1
'cmb_grossby.ListIndex = -1
'cmb_fsd.ListIndex = -1
'CMB_ICDH.ListIndex = -1

MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = ">??/#####"
'cmb_grossby.clear

End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In finnedaspcyt_report.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
   
    txt_addendum.Enabled = True
    
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provby.Enabled = True
    cmb_addendumby.Enabled = True
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    txt_report.Enabled = True
    txt_impdiag.Enabled = True
    cmb_dept.Enabled = True
    cmb_reqno.Enabled = True
    txt_provdiag.Enabled = False
    txt_sitefnac.Enabled = False
    txt_noofslides.Enabled = False
    txt_nmeterial.Enabled = True
    txt_noofslidesdry.Enabled = False
    txt_noslideswet.Enabled = False
    cmb_cellblock.Enabled = True
End If
If str1 = "DIS" Then
    txt_addendum.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_reqno.Enabled = False
    txt_report.Enabled = False
    txt_impdiag.Enabled = False
    txt_provdiag.Enabled = False
    txt_sitefnac.Enabled = False
    txt_noofslides.Enabled = False
    txt_nmeterial.Enabled = False
    txt_noofslidesdry.Enabled = False
    txt_noslideswet.Enabled = False
    cmb_cellblock.Enabled = False
End If
If str1 = "VIEW" Then

    txt_addendum.Enabled = True
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_report.Enabled = False
    txt_impdiag.Enabled = False
    txt_provdiag.Enabled = False
    txt_sitefnac.Enabled = False
    txt_noofslides.Enabled = False
    txt_nmeterial.Enabled = False
    txt_noofslidesdry.Enabled = False
    txt_noslideswet.Enabled = False
    cmb_cellblock.Enabled = False
End If
If str1 = "MODIFY" Then

    txt_addendum.Enabled = True

    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_provby.Enabled = True
    cmb_addendumby.Enabled = True
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_report.Enabled = True
    txt_impdiag.Enabled = True
    txt_provdiag.Enabled = False
    txt_sitefnac.Enabled = False
    txt_noofslides.Enabled = False
    txt_nmeterial.Enabled = True
    txt_noofslidesdry.Enabled = False
    txt_noslideswet.Enabled = False
    cmb_cellblock.Enabled = True
End If
If str1 = "ADDENDUM" Then

    txt_addendum.Enabled = True
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
    cmb_addendumby.Enabled = True
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    txt_report.Enabled = False
    txt_impdiag.Enabled = False
    txt_provdiag.Enabled = False
    txt_sitefnac.Enabled = False
    txt_noofslides.Enabled = False
    txt_nmeterial.Enabled = False
    txt_noofslidesdry.Enabled = False
    txt_noslideswet.Enabled = False
    cmb_cellblock.Enabled = False
End If
End Sub

Public Sub STAFF()
On Error GoTo MSG:
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
        If Trim(SELDREQSTAFFD1RS1("RELATION")) <> "SELF" Then
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
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
  'check by
'        If selma104flag = False Then
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
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
    If selpathdrequisitn_rs.RowCount > 0 Then
        'TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
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
                'txt_sites.Text = SELMA301RS("SITEDESC")
            End If
        End If
        If IsNull(Trim(selpathdrequisitn_rs("LATERALITS"))) Then
            'txt_laterality.Text = ""
        Else
            'txt_laterality.Text = selpathdrequisitn_rs("LATERALITS")
        End If
    
        If Trim(selpathdrequisitn_rs("FROZENSEC")) = "NO" Then
'           txt_frosecrep.Enabled = False
        End If
        'txt_natmeterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        'txt_opprocedure.Text = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
        
   
        If IsNull(selpathdrequisitn_rs("IMMHISCHE")) Then
'            txt_immhischerep.Enabled = False
            'txt_immhischereqno.Enabled = False
            
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
        'txt_procedure.Text = selta1012_rs("prvdiagdes")
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
      txt_cytology.Text = seldrequisitn_rs("PATHNO")
'     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
     seldrequisitn_rs.MoveNext
     Wend
'     txt_prepathreqno.Text = prpathno
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
'        If seldissecgrossflag = False Then
'            Set seldissecgross_rs = seldissecgross_qr.OpenResultset(1, 3)
'            seldissecgrossflag = True
'        Else
'            seldissecgross_rs.MoreResults
'            seldissecgross_rs.Requery
'        End If
'        If seldissecgross_rs.RowCount > 0 Then
''            cmb_grossby.clear
'            While Not seldissecgross_rs.EOF
''                cmb_grossby.AddItem Trim(seldissecgross_rs("CCNAME"))
'                seldissecgross_rs.MoveNext
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
        txt_age = DateDiff("YYYY", SELDREQSTAFFD1RS1("AGE"), Format(DATE11, "DD/MM/YYYY"))
        If Trim(SELDREQSTAFFD1RS1("RELATION")) <> "SELF" Then
            txt_status = "DP"
        Else
            txt_status = "ST"
        End If
    End If
''      SELEMPBASINFO_QR(0) = MaskEdBox1.Text
''      SELEMPBASINFO_RS.MoreResults
''      SELEMPBASINFO_RS.Requery
''      If SELEMPBASINFO_RS.RowCount > 0 Then
''        txt_sex.Text = SELEMPBASINFO_RS("GENDER")
''        txt_name.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
''        txt_age = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), CStr(DATE11))
''      End If

    If REFDRFLAG = False Then
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If

    If REFDR_RS.RowCount > 0 Then
    '***************************************
            If REFDR_RS("RECORDNO") = 1 Then
                lbl_vpp.Caption = "VPP"
            End If
            If REFDR_RS("RECORDNO") = 2 Then
                lbl_vpp.Caption = "LR"
            End If
            If REFDR_RS("RECORDNO") = 3 Then
                lbl_vpp.Caption = "BAR"
            End If
            If REFDR_RS("RECORDNO") = 4 Then
                lbl_vpp.Caption = "DrOp"
            End If
        
      '***************************************
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
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
    If selpathdrequisitn_rs.RowCount > 0 Then
        txt_cytology.Text = Trim(selpathdrequisitn_rs("PATHNO"))
        If selpathdrequisitn_rs.EOF = False Then
            txt_nmeterial.Text = " " & selpathdrequisitn_rs("nmeterials")
        Else
            txt_nmeterial.Text = ""
        End If
        cmb_cellblock.Text = selpathdrequisitn_rs("frozensec")
        txt_noslideswet.Text = selpathdrequisitn_rs("outblock")
        txt_noofslidesdry.Text = selpathdrequisitn_rs("outsspecim")
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
                txt_sitefnac.Text = SELMA301RS("SITEDESC")
                    'Sundeep *****************8
                    Else
                      txt_sitefnac.Text = ""
              End If
        End If
''        If IsNull(selpathdrequisitn_rs("LATERALITS")) Then
''        Else
''            txt_laterality.Text = selpathdrequisitn_rs("LATERALITS")
''        End If
        If IsNull(Trim(selpathdrequisitn_rs("NMETERIALS"))) Then
            txt_nmeterial.Text = ""
        Else
            txt_nmeterial.Text = Trim(selpathdrequisitn_rs("nmeterials"))
        End If
            
        'txt_natmeterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        'txt_opprocedure.Text = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
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
        txt_provdiag.Text = selta1012_rs("prvdiagdes")
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
            txt_report.Text = SELDFREPORT_RS("MICREPORT").Value
        Else
            txt_report.Text = ""
        End If
        If SELDFREPORT_RS.EOF = False Then
            txt_impdiag.Text = SELDFREPORT_RS("IMPRESSION").Value
        Else
            txt_impdiag.Text = ""
        End If
    Else
        'txt_frosecrep.Text = ""
    End If

              
         If SELDFREPORT_RS.RowCount > 0 Then
            If SELDFREPORT_RS.EOF = False Then
                 'txt_impfindiag.Text = SELDFREPORT_RS("impression").Value
            End If
        Else
            'txt_impfindiag.Text = ""
        End If
        
        If SELDFREPORT_RS.RowCount > 0 Then
            If SELDFREPORT_RS.EOF = False Then
'                txt_grossdesc = SELDFREPORT_RS("GROSSDESC").Value
            End If
        Else
'            txt_grossdesc = ""
        End If

         
        If SELDFREPORT_RS.RowCount > 0 Then
           If SELDFREPORT_RS.EOF = False Then
'               txt_microdesc = Trim(SELDFREPORT_RS("SPREPORT").Value)
            Else
                'txt_microdesc = SELDFREPORT_RS("SPREPORT").Value
           End If
        Else
'            txt_microdesc = ""
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
            seldimpresion_qr(0) = req_chglen(Trim(cmb_reqno.Text))
            Set seldimpresion_rs = seldimpresion_qr.OpenResultset(1, 3)
            SELDIMPRESIONFLAG = True
        Else
            seldimpresion_qr(0) = req_chglen(Trim(cmb_reqno.Text))
            seldimpresion_rs.MoreResults
            seldimpresion_rs.Requery
        End If
    If seldimpresion_rs.RowCount > 0 Then
           ' txt_impfindiag.Text = seldimpresion_rs("impression")
'            cmb_fsd.AddItem seldimpresion_rs("FDESCRY")
'            cmb_fsd.Text = seldimpresion_rs("FDESCRY")
            'cmb_t.Text = seldimpresion_rs("T")
            'txt_size.Text = seldimpresion_rs("TSIZE")
'            CMB_N.Text = seldimpresion_rs("N")
            'cmb_m.Text = seldimpresion_rs("M")
            'CMB_ICDH.AddItem seldimpresion_rs("FINALDIAG")
            'CMB_ICDH.Text = seldimpresion_rs("FINALDIAG")
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
'        CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
        seldrequisitn_rs.MoveNext
     Wend
'        txt_prepathreqno.Text = Trim(prpathno)
    End If
        SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
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

Private Sub txt_addendum_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content.Text = ""
d_grossdesc.Content.Text = txt_grossdesc.Text
W_grossdes.Caption = "Addendum"
W_grossdes.WindowState = wdWindowStateMaximize
d_grossdesc.Content.Text = txt_grossdesc.Text
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Addendum"
W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content.Text = ""
'd_grossdesc.Content = txt_addendum.Text
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
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_impdiag
W_grossdes.Caption = "Diagnosis"
W_grossdes.WindowState = wdWindowStateMaximize
d_grossdesc.Content = txt_impdiag
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Diagnosis"
W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_impdiag.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub

Private Sub txt_impdiag_LostFocus()
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument
txt_impdiag.Text = d_grossdesc.Content
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

Private Sub txt_nmeterial_GotFocus()
On Error GoTo m3:
w_micscop.Visible = False
Set d_micscop = w_micscop.ActiveDocument
d_micscop.Content.Text = ""
d_micscop.Content.Text = txt_nmeterial.Text
w_micscop.Caption = "Nature of Material"
w_micscop.WindowState = wdWindowStateMaximize
d_micscop.Content.Text = txt_nmeterial.Text
w_micscop.Visible = True
w_micscop.Activate
m3:
On Error Resume Next:
Set w_micscop = CreateObject("word.Application")
w_micscop.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w_micscop.Caption = "Nature of Material"
w_micscop.WindowState = wdWindowStateMaximize
Set d_micscop = w_micscop.ActiveDocument
d_micscop.Content.Text = txt_nmeterial.Text
w_micscop.Visible = True
w_micscop.Activate

End Sub

Private Sub txt_nmeterial_LostFocus()
On Error GoTo m4:
Set d_micscop = w_micscop.ActiveDocument
txt_nmeterial.Text = d_micscop.Content
'flagimp = False
d_micscop.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_micscop.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_micscop.Close (wdDoNotSaveChanges)
w_micscop.quit
'W1.Visible = False
Exit Sub
m4:
MsgBox Err.Description
End Sub

Private Sub txt_report_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_report
W_grossdes.Caption = "Microscopy Report"
W_grossdes.WindowState = wdWindowStateMaximize
d_grossdesc.Content = txt_report
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Microscopy Report"
W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_report.Text
W_grossdes.Visible = True
W_grossdes.Activate

End Sub

Private Sub txt_report_LostFocus()
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument
txt_report.Text = d_grossdesc.Content
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

Private Sub CGCNREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
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

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"

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
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
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

Private Sub RISREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'rq_grid_det(0) = Trim(VAR_REQNNO)
'rs_grid_det.MoreResults
'rs_grid_det.Requery

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I

If SELRREQNORS.RowCount > 0 And (Trim(SELRREQNORS("REQNSTAT")) = "T" Or Trim(SELRREQNORS("REQNSTAT")) = "F") Then
    If Opt_patient.Value = True Then
        CrystalReport1.Formulas(0) = "provtyp= '" & Trim(SELRREQNORS("ENTERBY")) & "'"
        'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
        'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
        'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
        SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
        If SELRMODALITYFLAG = False Then
            Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(1, 3)
            SELRMODALITYFLAG = True
        Else
            
            SELRMODALITY_RS.MoreResults
            SELRMODALITY_RS.Requery
        End If
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
        str1 = Space(3) & str1
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
        CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
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
    'CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
    Exit Sub
Else
    Exit Sub
End If
End Sub


Private Sub FAREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
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

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
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

Private Sub SPREPORT()
Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
On Error GoTo prtmsg:
DoEvents
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
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

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I


str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
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
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = 'Prev. Path No." & Trim(prpathno) & "'"
        Else
             CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
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
        selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
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
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
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


Public Sub MAKEFALSE()
    
    FAREPORTFLAG = False
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
