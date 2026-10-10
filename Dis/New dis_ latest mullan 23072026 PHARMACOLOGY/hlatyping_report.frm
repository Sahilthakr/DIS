VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form hlatyping_report 
   Caption         =   "ICU Requisition"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   77
      Top             =   8640
      Width           =   11775
      Begin VB.TextBox txt_comments 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   89
         Top             =   240
         Visible         =   0   'False
         Width           =   10455
      End
      Begin VB.ComboBox cmb_enterby 
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
         Left            =   1080
         TabIndex        =   18
         Top             =   600
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
         Height          =   315
         Left            =   9000
         TabIndex        =   78
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label LBL_COMMENTS 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
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
         TabIndex        =   81
         Top             =   360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Entered by"
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
         Left            =   0
         TabIndex        =   80
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lbl_finalby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
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
         Left            =   7320
         TabIndex        =   79
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame common 
      Height          =   3015
      Left            =   0
      TabIndex        =   41
      Top             =   1080
      Width           =   11775
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   0
         TabIndex        =   86
         Top             =   600
         Width           =   1215
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
            TabIndex        =   87
            Top             =   600
            Width           =   735
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
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txt_cytology 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   10200
         TabIndex        =   84
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_clndiag 
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
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox txt_relpat 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   9120
         TabIndex        =   75
         Top             =   2520
         Width           =   2415
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10200
         TabIndex        =   54
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
         Left            =   6240
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_nameofdonar 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   9120
         TabIndex        =   53
         Top             =   2160
         Width           =   2415
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   2160
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
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   2175
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DTM No."
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
         Left            =   9360
         TabIndex        =   85
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label1 
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
         Left            =   720
         TabIndex        =   83
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label lblrelpat 
         AutoSize        =   -1  'True
         Caption         =   "Relationship with the patient"
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
         Left            =   6480
         TabIndex        =   74
         Top             =   2640
         Width           =   2445
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         Left            =   1440
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_namofdonar 
         AutoSize        =   -1  'True
         Caption         =   "Name of Donor"
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
         Left            =   7680
         TabIndex        =   66
         Top             =   2160
         Width           =   1290
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   2160
         Width           =   2130
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
      Left            =   10680
      TabIndex        =   40
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10680
      TabIndex        =   39
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame COMMAND 
      Height          =   855
      Left            =   0
      TabIndex        =   37
      Top             =   9600
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
         MouseIcon       =   "hlatyping_report.frx":0000
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
         Left            =   1800
         MouseIcon       =   "hlatyping_report.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   52
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
         MouseIcon       =   "hlatyping_report.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   20
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
         MouseIcon       =   "hlatyping_report.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   19
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
         MouseIcon       =   "hlatyping_report.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   38
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
         MouseIcon       =   "hlatyping_report.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   22
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
         Left            =   10440
         MouseIcon       =   "hlatyping_report.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   23
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
         Left            =   7200
         MouseIcon       =   "hlatyping_report.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame REPORT 
      Height          =   4455
      Left            =   0
      TabIndex        =   24
      Top             =   4200
      Width           =   11820
      Begin VB.CommandButton Command1 
         Caption         =   "Show All Reqn."
         Height          =   255
         Left            =   9000
         TabIndex        =   105
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton CMD_LSTREQ 
         Caption         =   "Remove from print"
         Height          =   255
         Left            =   9000
         TabIndex        =   104
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   9000
         TabIndex        =   103
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txt_dnadqb1_SERO 
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
         TabIndex        =   101
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txt_dnadrb1_SERO 
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
         TabIndex        =   100
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txt_dnaDRB345_SERO 
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
         TabIndex        =   99
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txt_dnaDRB345 
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
         Left            =   3960
         TabIndex        =   98
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txt_hlac_SERO 
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
         TabIndex        =   97
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txt_hlab_SERO 
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
         TabIndex        =   96
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txt_hlaa_SERO 
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
         TabIndex        =   95
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox TXT_IMPRESSION 
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
         Left            =   2400
         TabIndex        =   17
         Top             =   3960
         Width           =   9015
      End
      Begin VB.TextBox txt_dnadqb1 
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
         Left            =   3960
         TabIndex        =   10
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txt_hlabw4 
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
         Left            =   3960
         TabIndex        =   6
         Top             =   2400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_hladr 
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
         Left            =   7680
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_hladqlocus 
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
         Left            =   7680
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_hlaa 
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
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txt_dnadrb1 
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
         Left            =   3960
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txt_dnadqa1 
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
         Left            =   3960
         TabIndex        =   15
         Top             =   3120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_hlab 
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
         Left            =   3960
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txt_hlac 
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
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txt_hladr52 
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
         Left            =   3960
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_hlab27 
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
         Left            =   7680
         TabIndex        =   14
         Top             =   2760
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_dnadrb1sub 
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
         Left            =   3960
         TabIndex        =   16
         Top             =   3480
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DRB3 / DRB4 / DRB5"
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
         TabIndex        =   102
         Top             =   2040
         Width           =   1920
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   1320
         TabIndex        =   88
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label lcl_dnadqb1 
         AutoSize        =   -1  'True
         Caption         =   "HLA DQB1 locus"
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
         Left            =   2400
         TabIndex        =   76
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label lbl_hlaa 
         AutoSize        =   -1  'True
         Caption         =   "HLA A Locus"
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
         Left            =   2760
         TabIndex        =   36
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lbl_dna 
         AutoSize        =   -1  'True
         Caption         =   "HLA-DRB1 locus"
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
         Left            =   2400
         TabIndex        =   35
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label lbl_hlabw4 
         AutoSize        =   -1  'True
         Caption         =   "HLA BW4/BW6"
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
         Left            =   2400
         TabIndex        =   34
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl_hladr 
         AutoSize        =   -1  'True
         Caption         =   " HLA DR"
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
         Left            =   6720
         TabIndex        =   33
         Top             =   2400
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_hladqlocus 
         AutoSize        =   -1  'True
         Caption         =   "HLA DQ Locus"
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
         Left            =   6240
         TabIndex        =   32
         Top             =   3120
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lbl_dnadqa1 
         AutoSize        =   -1  'True
         Caption         =   "HLA  DQA1 locus"
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
         Left            =   2280
         TabIndex        =   31
         Top             =   3120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lbl_hlac 
         AutoSize        =   -1  'True
         Caption         =   "HLA C Locus"
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
         Left            =   2640
         TabIndex        =   30
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lbl_hlab27 
         AutoSize        =   -1  'True
         Caption         =   "HLA B 27"
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
         Left            =   6720
         TabIndex        =   28
         Top             =   2760
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lbl_dnadrb1sub 
         AutoSize        =   -1  'True
         Caption         =   "DNA-DRB1 Subtyping of identified alleles"
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
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label lbl_hlab 
         AutoSize        =   -1  'True
         Caption         =   "HLA B Locus"
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
         Left            =   2760
         TabIndex        =   26
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lbl_hladr52 
         AutoSize        =   -1  'True
         Caption         =   "HLA AB-DR (Molecular Typing)"
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
         TabIndex        =   25
         Top             =   2760
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "HLA-DR52/DR53"
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
         Left            =   960
         TabIndex        =   91
         Top             =   2760
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.Label lbl_icureq 
      Caption         =   "ICU Requisition"
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
      Left            =   1920
      TabIndex        =   94
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl_rwd 
      Caption         =   "RECOVERY WARD"
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
      Height          =   255
      Left            =   120
      TabIndex        =   93
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "05/11/21"
      Height          =   255
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Width           =   975
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
      Height          =   375
      Left            =   1200
      TabIndex        =   90
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label lbl_transmedhla 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - HLA TYPING"
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
      Left            =   4080
      TabIndex        =   73
      Top             =   600
      Width           =   3225
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
      TabIndex        =   72
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
      TabIndex        =   71
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "hlatyping_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
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
'Dim DIMPRE As Document
Dim DIMPRE As Object
Dim WMETERIAL As Application
Dim DMETERIAL As Document

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
Dim updRTIMELOG_qr As rdoQuery
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

Dim INSHLAREPORTQR As rdoQuery
Dim UPDHLAREPORTQR As rdoQuery
Dim SELHLAREPORTQR As rdoQuery
Dim SELHLAREPORTRS As rdoResultset
Dim SELHLAREPORTFLAG As Boolean

''DTMUSER
Dim SELDTMUSER_QR As rdoQuery
Dim SELDTMUSER_RS As rdoResultset
Dim SELDTMUSER_FLAG As Boolean

Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_FLAG As Boolean

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
cmb_modality.SetFocus
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "NZZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
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
    If seldissecentrflag = False Then
            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
            seldissecentrflag = True
        Else
            seldissecentr_rs.MoreResults
            seldissecentr_rs.Requery
        End If
        If seldissecentr_rs.RowCount > 0 Then
            cmb_enteryby.clear
            While Not seldissecentr_rs.EOF
                cmb_enteryby.AddItem Trim(seldissecentr_rs("CCNAME"))
                seldissecentr_rs.MoveNext
            Wend
        End If
        If seldissecprovflag = False Then
            Set seldissecprov_rs = seldissecprov_qr.OpenResultset(1, 3)
            seldissecprovflag = True
        Else
            seldissecprov_rs.MoreResults
            seldissecprov_rs.Requery
        End If
        If seldissecprov_rs.RowCount > 0 Then
            cmb_provby.clear
            While Not seldissecprov_rs.EOF
                cmb_provby.AddItem Trim(seldissecprov_rs("CCNAME"))
                seldissecprov_rs.MoveNext
            Wend
        End If
        If seldissecfinflag = False Then
            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
            seldissecfinflag = True
        Else
            seldissecfin_rs.MoreResults
            seldissecfin_rs.Requery
        End If
        If seldissecfin_rs.RowCount > 0 Then
            cmb_finalby.clear
            While Not seldissecfin_rs.EOF
                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
                seldissecfin_rs.MoveNext
            Wend
        End If
        If seldissecgrossflag = False Then
            Set seldissecgross_rs = seldissecgross_qr.OpenResultset(1, 3)
            seldissecgrossflag = True
        Else
            seldissecgross_rs.MoreResults
            seldissecgross_rs.Requery
        End If
        If seldissecgross_rs.RowCount > 0 Then
            While Not seldissecgross_rs.EOF
                seldissecgross_rs.MoveNext
            Wend
        End If
Exit Sub
MSG:
'MsgBox Error(Err)
End Sub

Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "HLA TYPING$HT" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If

Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_back_Click()
btnenabl ("INI")
FLDENABL ("DIS")
TXT_CLEAR
DIS_MAIN.show
Me.Hide
End Sub
Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    If USRNAME_RS("SAVE") = "YES" Then
        cmd_save.Enabled = True
    Else
        cmd_save.Enabled = False
    End If
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
    cmd_view.Enabled = True
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
    cmd_view.Enabled = True
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
btnenabl ("INI")
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
'**********************************************
''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(LOGINUSRNAME)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
             'no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            If LOGINTYP = 8 Or LOGINTYP = 9 Then
            
            Else
            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
            cmd_commit.SetFocus
            Exit Sub
            End If
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
            If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
                '  If loginusrid <> curusrdetrs("usr_id") Then
                'MsgBox "Sorry! This Report is not to be Finalized by You"
                'cmd_commit.SetFocus
                'Exit Sub
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
'**********************************************

 rdb.BeginTrans
    updrrequisitn1_qr(0) = Left(cmb_enterby.Text, 30)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    
    
    updRTIMELOG_qr(0) = req_chglen(cmb_reqno.Text)
    updRTIMELOG_qr.Execute
    
    
End If
rdb.CommitTrans
btnenabl ("INI")
MsgBox "Record is Committed Successfully."
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
' Modified By Pankaj Baggal 03/09/2004
'D1.SaveAs "C:\REPORT.DOC"
D1.Close (wdDoNotSaveChanges)
w.quit
'Msk_caseno.Enabled = False
'cmd_cancel.SetFocus
MSG:
End Sub

Private Sub CMD_LSTREQ_Click()
   Dim KM As Integer
If List1.ListCount > 0 Then
    If List1.ListIndex < 0 Then
    Else
         
            For KM = 0 To List1.ListCount - 1
            If KM > List1.ListCount - 1 Then
            Else
                If List1.Selected(KM) = True Then
                List1.RemoveItem (KM)
                End If
                End If
            Next
    
    End If
    
End If
End Sub
Sub addListReq()
 List1.clear
  Dim rs_RRlIST  As rdoResultset
    Set rs_RRlIST = rdb.OpenResultset("SELECT * FROM RISLIB.RREQUISITN WHERE CASENO='" & Trim(MaskEdBox1.FormattedText) & "' AND DEPT_CODE='NZZ' AND MODALITYCD='HT'  AND REQNSTAT IN ('T','F') ", 1, 1)
    If Not rs_RRlIST.EOF Then
    While Not rs_RRlIST.EOF
        If Trim(rs_RRlIST("REQNNO")) <> "" Then
              List1.AddItem (Trim(rs_RRlIST("REQNNO")))
          End If
          rs_RRlIST.MoveNext
    Wend
    End If
End Sub
Private Sub cmd_modify_Click()
btnenabl ("NEW")
FLDENABL ("MODIFY")
flagsave = 2
'TextEnable
'If REFDRFLAG = False Then
'    REFDR_QR(0) = MaskEdBox1.FormattedText
'    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
'    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
'    REFDRFLAG = True
'Else
'    REFDR_QR(0) = MaskEdBox1.FormattedText
'    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
'    REFDR_RS.MoreResults
'    REFDR_RS.Requery
'End If
'If REFDR_RS.RowCount > 0 Then
'    If Trim(REFDR_RS("reqnstat")) = "T" And Trim(REFDR_RS("COMMITFLAG")) = "F" Then
'    ADDENDUMFALG = True
'    End If
'End If
'If ADDENDUMFALG = True Then
'    'txt_addendum.Enabled = True
'    'cmb_addendumby.Enabled = True
'    If seldissecaddrflag = False Then
'            Set seldissecadd_rs = seldissecadd_qr.OpenResultset(1, 3)
'            seldissecaddrflag = True
'        Else
'            seldissecadd_rs.MoreResults
'            seldissecadd_rs.Requery
'        End If
'        If seldissecadd_rs.RowCount > 0 Then
'            'cmb_addendumby.clear
'            While Not seldissecadd_rs.EOF
'                cmb_addendumby.AddItem Trim(seldissecadd_rs("CCNAME"))
'                seldissecadd_rs.MoveNext
'            Wend
'    End If
'    FLDENABL ("clindiag")
'    FLDENABL ("ADDENDUM")
'    flagsave = 2
'End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG:
btnenabl ("NEW")
field_clear
TXT_CLEAR

DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
cmb_enterby.Text = cmb_enterby.List(0)
cmb_enterby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enterby.List(0)) = cmb_finalby.List(L1) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
Next L1
'cmb_finalby.Text = cmb_finalby.List(0)

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
Opt_patient.Value = True
MaskEdBox1.Enabled = True
cmb_dept.SetFocus
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

Dim VAR_REQNNO As String
For I = 0 To 16
    CrystalReport1.Formulas(I) = ""
Next
VAR_REQNNO = req_chglen(cmb_reqno.Text)
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
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
If SELTA101FALG = False Then
    'selta101
    SELTA101_QR(0) = MaskEdBox1.FormattedText
    Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
    SELTA101FALG = True
Else
    'selta101
    SELTA101_QR(0) = MaskEdBox1.FormattedText
    Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
    SELTA101_RS.MoreResults
    SELTA101_RS.Requery
End If
If SELEMPBASINFOFLAG = False Then
    SELEMPBASINFO_QR(0) = MaskEdBox1.FormattedText
    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    SELEMPBASINFOFLAG = True
Else
    SELEMPBASINFO_QR(0) = MaskEdBox1.FormattedText
    SELEMPBASINFO_RS.MoreResults
    SELEMPBASINFO_RS.Requery
End If
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
'If ADDENDUMFALG = False Then
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLA.RPT"
'        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(TXT_CSITE.Text) = "" Then
'            CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
'        End If
'        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(cmb_enteryby.Text)
'            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
'            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CrystalReport1.Formulas(7) = "USRID = '" & STR & "'"
'Else
        'CrystalReport1.SelectionFormula = "{NZZHLARPT.REQNNO} = '" & VAR_REQNNO & "'"
        'commented and added by shweta 17-03-2015
        'CrystalReport1.SelectionFormula = "{NZZHLARPT.CASENO} = '" & MaskEdBox1.FormattedText & "'"
        If List1.ListCount = 0 Then
        CrystalReport1.SelectionFormula = "{NZZHLARPT.CASENO} = '" & MaskEdBox1.FormattedText & "'"
        Else
        CrystalReport1.SelectionFormula = "{NZZHLARPT.CASENO} = '" & MaskEdBox1.FormattedText & "'" & " AND {NZZHLARPT.REQNNO} in ['NZZHT05000011'], ['NZZHT06000032']"
        End If
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(1) = "USRID = '" & str & "'"
        If Len(Trim(selpathdrequisitn_rs("LABNO"))) > 0 Then
          '  CrystalReport1.Formulas(2) = "PNAME_D ='(Donor)      " & selpathdrequisitn_rs("NMETERIALS") & "'"
           ' CrystalReport1.Formulas(3) = "RELATION = 'Relation to Patient ' "
           '' CrystalReport1.Formulas(4) = "STDOB ='" & Trim(selpathdrequisitn_rs("OUTSSPECIM")) & "/" & "'"
           ' CrystalReport1.Formulas(5) = "AGE11 ='" & Trim(selpathdrequisitn_rs("LATERALITM")) & "'"
           ' 'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLAdon.RPT"
            CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISHLANEW.RPT"
           ' CrystalReport1.Formulas(6) = "PNAME = '(Patient)  " & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PNAME_D = '" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
            CrystalReport1.Formulas(3) = "RELATION ='" & " " & "'"
            'CrystalReport1.Formulas(4) = "STDOB ='" & Trim(SELTA101_RS("SEX")) & "/" & "'"
            CrystalReport1.Formulas(5) = "AGE11 ='" & DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), SELRREQNORS("REQNDT")) & "'"
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLA.RPT"
            CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISHLANEW.RPT"
        End If
     '   CrystalReport1.Formulas(5) = "DTMNO ='" & Trim(txt_cytology.Text) & "'"
CrystalReport1.Formulas(8) = "ent ='" & Trim(cmb_enterby.Text) & "'"
        
'End If
Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLAstaff.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        If Len(Trim(selpathdrequisitn_rs("LABNO"))) > 0 Then
            CrystalReport1.Formulas(8) = "NAM = '" & selpathdrequisitn_rs("NMETERIALS") & "'"
            CrystalReport1.Formulas(9) = "RELATION = 'Relation' "
            CrystalReport1.Formulas(10) = "STDOB ='" & Trim(selpathdrequisitn_rs("OUTSSPECIM")) & "'"
            CrystalReport1.Formulas(11) = "DOB ='" & Trim(selpathdrequisitn_rs("LATERALITM")) & "'"
        Else
            CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(9) = "RELATION ='" & " " & "'"
            CrystalReport1.Formulas(10) = "STDOB ='" & SELEMPBASINFO_RS("GENDER") & "'"
            CrystalReport1.Formulas(11) = "DOB ='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT")) & "'"
        End If
        'CRYSTALREPORT1.Formulas(10) = "STDOB ='" & SELEMPBASINFO_RS("GENDER") & "'"
        'CRYSTALREPORT1.Formulas(11) = "DOB ='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT")) & "'"
        CrystalReport1.Formulas(12) = "ent ='" & Trim(cmb_enterby.Text) & "'"

End If
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
btnenabl ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_save_Click()
On Error GoTo MSG:

Dim m As Integer
If Len(Trim(cmb_reqno)) = 0 Then
    m = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
    cmb_reqno.SetFocus
    Exit Sub
End If
If Trim(cmb_enterby.Text) = "" Then
    MsgBox " Entered By Field is Empty !!", vbInformation, "HLA Typing"
    cmb_enterby.SetFocus
    Exit Sub
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
            beforst = Mid(txt_impdiag, 1, K - 1) & "`"
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
'       If Len(Trim(txt_report.Text)) > 8000 Then
'            MsgBox "Report Max. of 8000 Characters Only.", vbInformation
'            txt_report.SetFocus
'            Exit Sub
'       End If
'       If Len(Trim(txt_impdiag.Text)) > 8000 Then
'            MsgBox "Impression Max. of 8000 Characters Only.", vbInformation
'            txt_impdiag.SetFocus
'            Exit Sub
'       End If
'       If Len(Trim(txt_impdiag.Text)) = 0 Then
'            MsgBox "Please Enter the Impression ", vbInformation
'            txt_impdiag.SetFocus
'            Exit Sub
'       End If
'       If Len(Trim(txt_typespeciman.Text)) > 1000 Then
'            MsgBox "Nature of Meterial Max. of 1000 Characters Only.", vbInformation
'            txt_typespeciman.SetFocus
'            Exit Sub
'       End If
'       If Len(Trim(txt_addendum.Text)) > 600 Then
'            MsgBox "Addendum Max. of 600 Characters Only.", vbInformation
'            txt_addendum.SetFocus
'            Exit Sub
'       End If
'rdb.BeginTrans
       assign
'rdb.CommitTrans
       'MsgBox "Record is saved Successfully."
  FLDENABL ("DIS")
  btnenabl ("SAVE")


  Exit Sub
MSG:
'rdb.RollbackTrans
MsgBox ERROR(Err)
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR
'&&&&&&&&&&&&&&&&&&&&77 not found hence declared
Dim SEL_DEPTRSFLAG As Boolean
cmb_enterby.Enabled = False
btnenabl ("VIEW")
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
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
cmb_dept.SetFocus
If LOGINTYP >= 3 Then
    cmd_commit.Enabled = True
Else
    cmd_commit.Enabled = False
End If
cmd_print.Enabled = True
Opt_patient.Value = True

'TextEnable
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
addListReq
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
If CYREPORTFLAG = False Then
    CYREPORTFLAG = True
'btnenabl ("INI")
NEWFLAG = False
viewflag = False
MaskEdBox1.Enabled = False

   
    'To display current date
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
    
'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1HLA", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 3)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    'cmb_enterby.Text = (LOGINUSRNAME)
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Enabled = False
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'***************** Shraddha Pandit 04/03/09 ******************
If ipadd = "150.150.200.62" Or ipadd = "150.150.15.51" Then
    lbl_icureq.Visible = True
End If
If ipadd = "10.3.1.56" Then
    lbl_rwd.Visible = True
End If
'*************************************************************

btnenabl ("INI")
'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'STR = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
'Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STREEE", STR)
'REPWTRN_QR(0) = Trim(USRDEPT)
'Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
'If REPWTRNRS.RowCount > 0 Then
'    'cmb_enteryby.clear
'    cmb_addendumby.clear
'    cmb_finalby.clear
'    cmb_provby.clear
'    REPWTRNRS.MoveFirst
'    While Not REPWTRNRS.EOF
'        If REPWTRNRS("USRTYPE") > 1 Then
'            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
'            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
'            cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
'        End If
'        REPWTRNRS.MoveNext
'    Wend
'Else
'    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
'    cmd_cancel_Click
'    Exit Sub
'End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11HLA", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'
str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE>=? and dept_code='NZZ' AND MODMR='6'"
Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1HLA", str)
SELFINALBY_QR(0) = 3
Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
cmb_finalby.clear
If Not SELFINALBY_RS.EOF Then
    Do While Not SELFINALBY_RS.EOF
       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
       SELFINALBY_RS.MoveNext
    Loop
End If

'*********SECURITY**********************************

''FOR DTM USER
str = "SELECT * FROM DISLIB.DTMUSER"
Set SELDTMUSER_QR = rdb.CreateQuery("SELDTMUSER_STR11", str)
    
    
    'for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' OR DEPT_CODE='IZZ')"
Set sel_deptqr = rdb.CreateQuery("selC1_deptqrht", str)
'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
   
           
str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301QR = rdb.CreateQuery("SELMA3011QRht", str)

    'FOR GETTING CLINICAL DIAGNOSIS
str = "select CLINHIST  from rislib.CLINHIST WHERE CASENO=? "
Set SELCLIN_QR = rdb.CreateQuery("SELCLIN_QRht", str)
        
    ' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqrht", str)

    'for retriving reqnos
str = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? AND REQNSTAT='E' AND DEPT_CODE='NZZ' ORDER BY REQNDT DESC"
Set req_qr = rdb.CreateQuery("1reqht", str)

    ' for retriveing name,age etc
Set SELTA101_QR = rdb.CreateQuery("selta101_qrht", "select * from PATOOLSLIB.ta101 where case_no =?")

    ' FOR REFEFING DOCTOR
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
Set REFDR_QR = rdb.CreateQuery("REFDR1_QRht", str)

    'FOR PATH NO
str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1_QRht", str)

    'for bedno, ward retriveing
str = "select * from PATOOLSLIB.ma104 where case_no =?"
Set selma104_qr = rdb.CreateQuery("selma1041_qrht", str)
    
    'for provisional diagnosis
str = "select * from patoolslib.ta1012 where case_no =?"
Set selta1012_qr = rdb.CreateQuery("se1lta10121_qrht", str)
  
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QRht", str)
    
    ' for update enterby, provreportby, finalby
    str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?, FREPORTBY=?, FREPORTDT=? WHERE REQNNO=?"
    Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qrht", str)
   
    str = "UPDATE RISLIB.RTIMELOG SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set updRTIMELOG_qr = rdb.CreateQuery("", str)
    
    str = "UPDATE RISLIB.RREQUISITN SET COMREPDATE=CURRENT DATE, COMREPTIME=CURRENT TIME WHERE REQNNO=?"
    Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrht", str)
    
    'for insert addendum report
    str = "INSERT INTO DISLIB.ADDENDUMCY(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
    Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QRht", str)
   
    str = "SELECT  * FROM DISLIB.ADDENDUMCY WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QRht", str)
   
    str = "select * from dislib.drequisitn where caseno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_1qrht", str)

    str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
    Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR1111", str)
'    txt_addendum.Enabled = False
'    cmb_addendumby.Enabled = False
    cmb_dept.Enabled = False
    FLDENABL ("DIS")

   str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
   Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQRht", str)
   
   str = "SELECT * FROM DISLIB.Dcyreport WHERE REQNO=?"
   Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QRht", str)
    
   str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
   Set SELDFREPORT_IHQR = rdb.CreateQuery("SELDFREPORT_IHQRht", str)
    
   str = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?, NMETERIALS=? WHERE REQNO=?"
   Set upddrequisitn_qr = rdb.CreateQuery("UPDSPREPORT1_QRht", str)
    
    str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO1190_QR1", str)
'    SELEMPBASINFO_QR(0) = "00000"
'    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    
    
    
'    STR = "select * from dislib.dgrade where gradecode=?"
'    Set seldgrade_qr = rdb.CreateQuery("seldgrade_qrcg", STR)
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
    Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRht", str)
    
   str = "select * from dislib.dihreport where reqno=?"
   Set selihreportqr = rdb.CreateQuery("selihreportqrht", str)
   
   str = "select * from dislib.dissec where enterby='YES'AND MODALITY='CG' "
   Set seldissecentr_qr = rdb.CreateQuery("seldissecentr_qrht", str)
   
   str = "select * from dislib.dissec where PROVby='YES' AND (MODALITY='CG'OR MODALITY='SP') "
   Set seldissecprov_qr = rdb.CreateQuery("seldissecPROV_qrht", str)
   
   str = "select * from dislib.dissec where FINALby='YES' AND (MODALITY='CG'OR MODALITY='SP') "
   Set seldissecfin_qr = rdb.CreateQuery("seldissecFIN_qrht", str)
   
   str = "select * from dislib.dissec where ADDENDUMby='YES' AND (MODALITY='CG'OR MODALITY='SP') "
   Set seldissecadd_qr = rdb.CreateQuery("seldissecADD_qrht", str)
   
'   STR = "select * from dislib.dissec where GROSSby='YES'"
'   Set seldissecgross_qr = rdb.CreateQuery("seldissecgROSS_qrcg", STR)
   
   str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
   Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1QRht", str)
   
   str = "select sitecd,sitedesc from dislib.sitedup where sitecd = ?"
   Set viewsitedupqr = rdb.CreateQuery("viewsitedupqrht", str)
   
    str = "INSERT INTO RISLIB.DISIMP(REQNO,IMP)VALUES(?,?)"
   Set INSDISIMP_QR = rdb.CreateQuery("INSDISIMP_QRht", str)
  
'   STR = "UPDATE RISLIB.DISIMP SET IMP=? WHERE REQNO=?"
'   Set UPDDISIMP_QR = rdb.CreateQuery("UPDDISIMP_QRht", STR)

    str = "INSERT INTO DISLIB.HLAREPORT VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSHLAREPORTQR = rdb.CreateQuery("INSHLAREPORTQRHT", str)
    
    str = "UPDATE DISLIB.HLAREPORT SET HLAALOCUS=?,HLABLOCUS=?,HLABW4BW6=?,HLACLOCUS=?,HLADR=?,HLADR5253=?,HLADQLOCUS=?,HLAB27=?,DNADRB1LR=?,DNADQA1=?,DNADQB1=?,DNADRB1SIA=?,IMPRESSION=?,REMARKS=?,DRB345=?,SERO1=?,SERO2=?,SERO3=?,SERO4=?,SERO5=?,SERO6=? WHERE REQNO=?"
    Set UPDHLAREPORTQR = rdb.CreateQuery("UPDHLAREPORTQRHT", str)
   
   'str = "SELECT REQNO,HLAALOCUS,HLABLOCUS,HLABW4BW6,HLACLOCUS,HLADR,HLADR5253, HLADQLOCUS,HLAB27,DNADRB1LR, DNADQA1,DNADQB1,DNADRB1SIA,IMPRESSION,REMARKS FROM DISLIB.HLAREPORT WHERE REQNO=?"
   str = "SELECT REQNO,HLAALOCUS,HLABLOCUS,HLABW4BW6,HLACLOCUS,HLADR,HLADR5253, HLADQLOCUS,HLAB27,DNADRB1LR, DNADQA1,DNADQB1,DNADRB1SIA,IMPRESSION,REMARKS,DRB345,SERO1,SERO2,SERO3,SERO4,SERO5,SERO6 FROM DISLIB.HLAREPORT WHERE REQNO=?"
   Set SELHLAREPORTQR = rdb.CreateQuery("SELHLAREPORTQRHT", str)
 End If
 
' If SELDTMUSER_FLAG = False Then
'    Set SELDTMUSER_RS = SELDTMUSER_QR.OpenResultset(1, 3)
'    SELDTMUSER_FLAG = False
'Else
'    SELDTMUSER_RS.MoreResults
'    SELDTMUSER_RS.Requery
'End If
'If Not SELDTMUSER_RS.EOF Then
'    Do While Not SELDTMUSER_RS.EOF
'        cmb_enterby.AddItem SELDTMUSER_RS("USERNAME")
'        SELDTMUSER_RS.MoveNext
'    Loop
'End If
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
cmd_view.Enabled = False
cmd_modify.Enabled = True
cmd_print.Enabled = True
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
        ''MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 05/12/2022
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
            If SELRREQUISITN_RS.RowCount > 0 Then
                cmb_reqno.clear
                While Not SELRREQUISITN_RS.EOF
                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                    SELRREQUISITN_RS.MoveNext
                Wend
             Else
               ' MsgBox " No  Requisition for this Modality"
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
  addListReq
cmb_reqno.SetFocus

 Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
On Error GoTo MSG:
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
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
            If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = "" & REFDR_RS("REFBYDOC")
            
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
        If Len(selpathdrequisitn_rs("PREPATHNO")) > 0 Then
            txt_relpat.Text = selpathdrequisitn_rs("PREPATHNO")
        End If
        If Len(selpathdrequisitn_rs("NMETERIALS")) > 0 Then
            txt_nameofdonar.Text = selpathdrequisitn_rs("NMETERIALS")
        End If
        If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            SELDTMNO_RS.MoreResults
            SELDTMNO_RS.Requery
        End If
        If SELDTMNO_RS.RowCount > 0 Then
            txt_cytology.Text = "" & SELDTMNO_RS("DTMNO")
        End If
'
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
        
        
        
'        If seldissecentrflag = False Then
'            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
'            seldissecentrflag = True
'        Else
'            seldissecentr_rs.MoreResults
'            seldissecentr_rs.Requery
'        End If
'        If seldissecentr_rs.RowCount > 0 Then
'            cmb_enterby.clear
'            While Not seldissecentr_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr_rs("CCNAME"))
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
        If seldissecfinflag = False Then
            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
            seldissecfinflag = True
        Else
            seldissecfin_rs.MoreResults
            seldissecfin_rs.Requery
        End If
        If seldissecfin_rs.RowCount > 0 Then
            cmb_finalby.clear
            While Not seldissecfin_rs.EOF
                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
                seldissecfin_rs.MoveNext
            Wend
        End If
        
    End If
    
    If viewflag = True Then
        viewflag = False
        
        '''***DONOR NAME & RELATION
        
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If Len(selpathdrequisitn_rs("PREPATHNO")) > 0 Then
            txt_relpat.Text = selpathdrequisitn_rs("PREPATHNO")
        End If
        If Len(selpathdrequisitn_rs("NMETERIALS")) > 0 Then
            txt_nameofdonar.Text = selpathdrequisitn_rs("NMETERIALS")
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
                MsgBox "Final Report is Generated, You cann't Modify the Final Report"
                cmd_modify.Enabled = False
                'ADDENDUMFALG = True
            End If
            
        End If
      
        
'        If selma104flag = False Then
'            selma104_qr(0) = MaskEdBox1.FormattedText
'            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'            selma104flag = True
'        Else
'            selma104_qr(0) = MaskEdBox1.FormattedText
'            selma104_rs.MoreResults
'            selma104_rs.Requery
'        End If
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
      
        If REFDR_RS.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
            txt_ward.Text = "" & selpathdrequisitn_rs("WARD_NAME")
            txt_bedno.Text = "" & selpathdrequisitn_rs("bedno")
            txt_status.Text = "" & REFDR_RS("PATNSTAT")
        Else
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
            txt_status.Text = REFDR_RS("PATNSTAT")
        End If
      
     
        
        If selpathdrequisitn_rs.RowCount > 0 Then
            'TXT_PATH.Text = SELPATHDREQUISITN_RS("PATHNO")
         
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
            txt_clndiag.Text = "" & selta1012_rs("prvdiagdes")
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
        
        If SELHLAREPORTFLAG = False Then
            SELHLAREPORTQR(0) = req_chglen(cmb_reqno.Text)
            Set SELHLAREPORTRS = SELHLAREPORTQR.OpenResultset(1, 3)
            SELHLAREPORTFLAG = True
        Else
            SELHLAREPORTQR(0) = req_chglen(cmb_reqno.Text)
            SELHLAREPORTRS.MoreResults
            SELHLAREPORTRS.Requery
            
        End If
        If SELHLAREPORTRS.RowCount > 0 Then
            txt_hlaa.Text = " " & SELHLAREPORTRS("HLAALOCUS")
            txt_hlab.Text = " " & SELHLAREPORTRS("HLABLOCUS")
            txt_hlabw4.Text = " " & SELHLAREPORTRS("HLABW4BW6")
            txt_hlac.Text = " " & SELHLAREPORTRS("HLACLOCUS")
            txt_hladr.Text = " " & SELHLAREPORTRS("HLADR")
            txt_hladqlocus.Text = " " & SELHLAREPORTRS("HLADQLOCUS")
            txt_hlab27.Text = " " & SELHLAREPORTRS("HLAB27")
            txt_hladr52.Text = " " & SELHLAREPORTRS("HLADR5253")
            txt_dnadrb1.Text = " " & SELHLAREPORTRS("DNADRB1LR")
            txt_dnadqa1.Text = " " & SELHLAREPORTRS("DNADQA1")
            txt_dnadqb1.Text = " " & SELHLAREPORTRS("DNADQB1")
            txt_dnadrb1sub.Text = " " & SELHLAREPORTRS("DNADRB1SIA")
            TXT_IMPRESSION.Text = " " & Trim(SELHLAREPORTRS("IMPRESSION"))
            txt_comments.Text = " " & Trim(SELHLAREPORTRS("REMARKS"))
            
            
                txt_dnaDRB345.Text = " " & SELHLAREPORTRS("DRB345")
            
             txt_hlaa_SERO.Text = " " & SELHLAREPORTRS("SERO1")
            txt_hlab_SERO.Text = " " & SELHLAREPORTRS("SERO2")
            txt_hlac_SERO.Text = " " & SELHLAREPORTRS("SERO3")
            txt_dnadrb1_SERO.Text = " " & SELHLAREPORTRS("SERO4")
            txt_dnadqb1_SERO.Text = " " & SELHLAREPORTRS("SERO5")
            txt_dnaDRB345_SERO.Text = " " & SELHLAREPORTRS("SERO6")
            
        End If
        cmb_reqno.Enabled = False
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
'        txt_clindiag.Enabled = False1
        
        If SELRREQNOFLAG = False Then
            SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
            Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
            SELRREQNOFLAG = True
        Else
            SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
            SELRREQNORS.MoreResults
            SELRREQNORS.Requery
        End If
        
        cmb_enterby.Text = "" & SELRREQNORS("ENTERBY")
        cmb_finalby.Text = "" & SELRREQNORS("PREPORTBY")
        If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            SELDTMNO_RS.MoreResults
            SELDTMNO_RS.Requery
        End If
        If SELDTMNO_RS.RowCount > 0 Then
            txt_cytology.Text = "" & SELDTMNO_RS("DTMNO")
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

Private Sub Opt_patient_Click()
    If Opt_patient.Value = True Then
       lbl_caseno.Caption = "Case No."
       ' MaskEdBox1.Mask = "??>/#####"
        MaskEdBox1.Enabled = True
        MaskEdBox1.SetFocus
    End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    ' MaskEdBox1.Mask = "#######"      ' COMMENTED BY VARINDER
    MaskEdBox1.Enabled = True
    MaskEdBox1.SetFocus
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
On Error GoTo MSG:
Dim I As Integer
Dim MICROSCOPIC As String
rdb.BeginTrans
If flagsave = 1 Then
   flagsave = 0

''    '**** INSERTING INTO HLAREPORT

        INSHLAREPORTQR(0) = req_chglen(cmb_reqno.Text)
        If Len(Trim(txt_hlaa.Text)) > 0 Then
            INSHLAREPORTQR(1) = Null
            INSHLAREPORTQR(1) = Trim(txt_hlaa.Text)
        Else
            INSHLAREPORTQR(1) = Null
        End If
        If Len(Trim(txt_hlab.Text)) > 0 Then
            INSHLAREPORTQR(2) = Null
            INSHLAREPORTQR(2) = Trim(txt_hlab.Text)
        Else
            INSHLAREPORTQR(2) = Null
        End If
        If Len(Trim(txt_hlabw4.Text)) > 0 Then
            INSHLAREPORTQR(3) = Null
            INSHLAREPORTQR(3) = Trim(txt_hlabw4.Text)
        Else
            INSHLAREPORTQR(3) = Null
        End If
        If Len(Trim(txt_hlac.Text)) > 0 Then
            INSHLAREPORTQR(4) = Null
            INSHLAREPORTQR(4) = Trim(txt_hlac.Text)
        Else
            INSHLAREPORTQR(4) = Null
        End If
        If Len(Trim(txt_hladr.Text)) > 0 Then
            INSHLAREPORTQR(5) = Null
            INSHLAREPORTQR(5) = Trim(txt_hladr.Text)
        Else
            INSHLAREPORTQR(5) = Null
        End If
        If Len(Trim(txt_hladr52.Text)) > 0 Then
            INSHLAREPORTQR(6) = Null
            INSHLAREPORTQR(6) = Trim(txt_hladr52.Text)
        Else
            INSHLAREPORTQR(6) = Null
        End If
        If Len(Trim(txt_hladqlocus.Text)) > 0 Then
            INSHLAREPORTQR(7) = Null
            INSHLAREPORTQR(7) = Trim(txt_hladqlocus.Text)
        Else
            INSHLAREPORTQR(7) = Null
        End If
        If Len(Trim(txt_hlab27.Text)) > 0 Then
            INSHLAREPORTQR(8) = Null
            INSHLAREPORTQR(8) = Trim(txt_hlab27.Text)
        Else
            INSHLAREPORTQR(8) = Null
        End If
        If Len(Trim(txt_dnadrb1.Text)) > 0 Then
            INSHLAREPORTQR(9) = Null
            INSHLAREPORTQR(9) = Trim(txt_dnadrb1.Text)
        Else
            INSHLAREPORTQR(9) = Null
        End If
        If Len(Trim(txt_dnadqa1.Text)) > 0 Then
            INSHLAREPORTQR(10) = Null
            INSHLAREPORTQR(10) = Trim(txt_dnadqa1.Text)
        Else
            INSHLAREPORTQR(10) = Null
        End If
        If Len(Trim(txt_dnadqb1.Text)) > 0 Then
            INSHLAREPORTQR(11) = Null
            INSHLAREPORTQR(11) = Trim(txt_dnadqb1.Text)
        Else
            INSHLAREPORTQR(11) = Null
        End If
        If Len(Trim(txt_dnadrb1sub.Text)) > 0 Then
            INSHLAREPORTQR(12) = Null
            INSHLAREPORTQR(12) = Trim(txt_dnadrb1sub.Text)
        Else
            INSHLAREPORTQR(12) = Null
        End If
        If Len(Trim(TXT_IMPRESSION.Text)) > 0 Then
            INSHLAREPORTQR(13) = Null
            INSHLAREPORTQR(13) = Trim(TXT_IMPRESSION.Text)
        Else
           INSHLAREPORTQR(13) = Null
        End If
        If Len(Trim(txt_comments.Text)) > 0 Then
            INSHLAREPORTQR(14) = Null
            INSHLAREPORTQR(14) = Trim(txt_comments.Text)
        Else
            INSHLAREPORTQR(14) = Null
        End If
         INSHLAREPORTQR(15) = ""
         
         
         INSHLAREPORTQR(16) = ""
         
               
       If Len(Trim(txt_dnaDRB345.Text)) > 0 Then
            INSHLAREPORTQR(15) = Null
            INSHLAREPORTQR(15) = Trim(txt_dnaDRB345.Text)
        Else
           INSHLAREPORTQR(15) = Null
        End If
         
         
         
         'ADDED BY SHWETA 18-02-2014
         
       If Len(Trim(txt_hlaa_SERO.Text)) > 0 Then
            INSHLAREPORTQR(17) = Null
            INSHLAREPORTQR(17) = Trim(txt_hlaa_SERO.Text)
        Else
           INSHLAREPORTQR(17) = Null
        End If
         
         If Len(Trim(txt_hlab_SERO.Text)) > 0 Then
            INSHLAREPORTQR(18) = Null
            INSHLAREPORTQR(18) = Trim(txt_hlab_SERO.Text)
        Else
           INSHLAREPORTQR(18) = Null
        End If
         
              If Len(Trim(txt_hlac_SERO.Text)) > 0 Then
            INSHLAREPORTQR(19) = Null
            INSHLAREPORTQR(19) = Trim(txt_hlac_SERO.Text)
        Else
           INSHLAREPORTQR(19) = Null
        End If
         
          If Len(Trim(txt_dnadrb1_SERO.Text)) > 0 Then
            INSHLAREPORTQR(20) = Null
            INSHLAREPORTQR(20) = Trim(txt_dnadrb1_SERO.Text)
        Else
           INSHLAREPORTQR(20) = Null
        End If
         
           If Len(Trim(txt_dnadqb1_SERO.Text)) > 0 Then
            INSHLAREPORTQR(21) = Null
            INSHLAREPORTQR(21) = Trim(txt_dnadqb1_SERO.Text)
        Else
           INSHLAREPORTQR(21) = Null
        End If
         
         
       If Len(Trim(txt_dnaDRB345_SERO.Text)) > 0 Then
            INSHLAREPORTQR(22) = Null
            INSHLAREPORTQR(22) = Trim(txt_dnaDRB345_SERO.Text)
        Else
           INSHLAREPORTQR(22) = Null
        End If
         
         
        INSHLAREPORTQR.Execute
    
    
    '**** INSERTING INTO HLAREPORT
    updrrequisitn_qr(0) = UCase(cmb_enterby.Text)
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = " " 'UCase(cmb_provby.Text)
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = " " 'cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    
    
    
     
    End If
  

If flagsave = 2 Then
    flagsave = 0
    
    ''    '**** UPDATING INTO HLAREPORT

        
        If Len(Trim(txt_hlaa.Text)) > 0 Then
            UPDHLAREPORTQR(0) = Null
            UPDHLAREPORTQR(0) = Trim(txt_hlaa.Text)
        Else
            UPDHLAREPORTQR(0) = Null
        End If
        If Len(Trim(txt_hlab.Text)) > 0 Then
            UPDHLAREPORTQR(1) = Null
            UPDHLAREPORTQR(1) = Trim(txt_hlab.Text)
        Else
            UPDHLAREPORTQR(1) = Null
        End If
        If Len(Trim(txt_hlabw4.Text)) > 0 Then
            UPDHLAREPORTQR(2) = Null
            UPDHLAREPORTQR(2) = Trim(txt_hlabw4.Text)
        Else
            UPDHLAREPORTQR(2) = Null
        End If
        If Len(Trim(txt_hlac.Text)) > 0 Then
            UPDHLAREPORTQR(3) = Null
            UPDHLAREPORTQR(3) = Trim(txt_hlac.Text)
        Else
            UPDHLAREPORTQR(3) = Null
        End If
        If Len(Trim(txt_hladr.Text)) > 0 Then
            UPDHLAREPORTQR(4) = Null
            UPDHLAREPORTQR(4) = Trim(txt_hladr.Text)
        Else
            UPDHLAREPORTQR(4) = Null
        End If
        If Len(Trim(txt_hladr52.Text)) > 0 Then
            UPDHLAREPORTQR(5) = Null
            UPDHLAREPORTQR(5) = Trim(txt_hladr52.Text)
        Else
            UPDHLAREPORTQR(5) = Null
        End If
        If Len(Trim(txt_hladqlocus.Text)) > 0 Then
            UPDHLAREPORTQR(6) = Null
            UPDHLAREPORTQR(6) = Trim(txt_hladqlocus.Text)
        Else
            UPDHLAREPORTQR(6) = Null
        End If
        If Len(Trim(txt_hlab27.Text)) > 0 Then
            UPDHLAREPORTQR(7) = Null
            UPDHLAREPORTQR(7) = Trim(txt_hlab27.Text)
        Else
            UPDHLAREPORTQR(7) = Null
        End If
        If Len(Trim(txt_dnadrb1.Text)) > 0 Then
            UPDHLAREPORTQR(8) = Null
            UPDHLAREPORTQR(8) = Trim(txt_dnadrb1.Text)
        Else
            UPDHLAREPORTQR(8) = Null
        End If
        If Len(Trim(txt_dnadqa1.Text)) > 0 Then
            UPDHLAREPORTQR(9) = Null
            UPDHLAREPORTQR(9) = Trim(txt_dnadqa1.Text)
        Else
            UPDHLAREPORTQR(9) = Null
        End If
        If Len(Trim(txt_dnadqb1.Text)) > 0 Then
            UPDHLAREPORTQR(10) = Null
            UPDHLAREPORTQR(10) = Trim(txt_dnadqb1.Text)
        Else
            UPDHLAREPORTQR(10) = Null
        End If
        If Len(Trim(txt_dnadrb1sub.Text)) > 0 Then
            UPDHLAREPORTQR(11) = Null
            UPDHLAREPORTQR(11) = Trim(txt_dnadrb1sub.Text)
        Else
            UPDHLAREPORTQR(11) = Null
        End If
        If Len(Trim(TXT_IMPRESSION.Text)) > 0 Then
            UPDHLAREPORTQR(12) = Null
            UPDHLAREPORTQR(12) = Trim(TXT_IMPRESSION.Text)
        Else
           UPDHLAREPORTQR(12) = Null
        End If
        If Len(Trim(txt_comments.Text)) > 0 Then
            UPDHLAREPORTQR(13) = Null
            UPDHLAREPORTQR(13) = Trim(txt_comments.Text)
        Else
            UPDHLAREPORTQR(13) = Null
        End If
       
    '---mrunal
    
         
             
               
       If Len(Trim(txt_dnaDRB345.Text)) > 0 Then
            UPDHLAREPORTQR(14) = Null
            UPDHLAREPORTQR(14) = Trim(txt_dnaDRB345.Text)
        Else
           UPDHLAREPORTQR(14) = Null
        End If
         
         
     If Len(Trim(txt_hlaa_SERO.Text)) > 0 Then
            UPDHLAREPORTQR(15) = Null
            UPDHLAREPORTQR(15) = Trim(txt_hlaa_SERO.Text)
        Else
           UPDHLAREPORTQR(15) = Null
        End If
         
         If Len(Trim(txt_hlab_SERO.Text)) > 0 Then
            UPDHLAREPORTQR(16) = Null
            UPDHLAREPORTQR(16) = Trim(txt_hlab_SERO.Text)
        Else
           UPDHLAREPORTQR(16) = Null
        End If
         
              If Len(Trim(txt_hlac_SERO.Text)) > 0 Then
            UPDHLAREPORTQR(17) = Null
            UPDHLAREPORTQR(17) = Trim(txt_hlac_SERO.Text)
        Else
           UPDHLAREPORTQR(17) = Null
        End If
         
          If Len(Trim(txt_dnadrb1_SERO.Text)) > 0 Then
            UPDHLAREPORTQR(18) = Null
            UPDHLAREPORTQR(18) = Trim(txt_dnadrb1_SERO.Text)
        Else
           UPDHLAREPORTQR(18) = Null
        End If
         
           If Len(Trim(txt_dnadqb1_SERO.Text)) > 0 Then
            UPDHLAREPORTQR(19) = Null
            UPDHLAREPORTQR(19) = Trim(txt_dnadqb1_SERO.Text)
        Else
           UPDHLAREPORTQR(19) = Null
        End If
         
         
       If Len(Trim(txt_dnaDRB345_SERO.Text)) > 0 Then
            UPDHLAREPORTQR(20) = Null
            UPDHLAREPORTQR(20) = Trim(txt_dnaDRB345_SERO.Text)
        Else
           UPDHLAREPORTQR(20) = Null
        End If
         UPDHLAREPORTQR(21) = req_chglen(cmb_reqno.Text)
        UPDHLAREPORTQR.Execute
        '----Mrunal
        '---------------
    '**** INSERTING INTO HLAREPORT
    
    
    
    
    updrrequisitn_qr(0) = cmb_enterby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_finalby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = " " 'cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
   
''    '**** UPDATE THE DISIMP
''        If Len(Trim(txt_impdiag.Text)) > 0 Then
''        UPDDISIMP_QR(0) = Null
''        UPDDISIMP_QR(0) = Trim(txt_impdiag.Text)
''        UPDDISIMP_QR(1) = Trim(req_chglen(cmb_reqno.Text))
''        UPDDISIMP_QR.Execute
''        End If
''
''    '*** UPDATE THE DISIMP
    
    '**** UPDATING INTO FREPORT
    
''    If ADDENDUMFALG = True Then
''        If SELDADDENDUMFLAG = False Then
''            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
''            Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
''            SELDADDENDUMFLAG = True
''        Else
''            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
''            SELDADDENDUM_RS.MoreResults
''            SELDADDENDUM_RS.Requery
''        End If
''        If SELDADDENDUM_RS.RowCount > 0 Then
''            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
''            INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
''            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''            INSDADDENDUM_QR(3) = time1
''            INSDADDENDUM_QR(4) = cmb_addendumby.Text
''            INSDADDENDUM_QR(5) = Null
''            INSDADDENDUM_QR(5) = txt_addendum.Text
''            INSDADDENDUM_QR.Execute
''        Else
''            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
''            INSDADDENDUM_QR(1) = "1"
''            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''            INSDADDENDUM_QR(3) = time1
''            INSDADDENDUM_QR(4) = cmb_addendumby.Text
''            INSDADDENDUM_QR(5) = Null
''            INSDADDENDUM_QR(5) = txt_addendum.Text
''            INSDADDENDUM_QR.Execute
''        End If
''    End If
    
  End If
  rdb.CommitTrans
  MsgBox "Record is Successfully Saved."
  cmd_commit.Enabled = False
  Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
End Sub

Public Sub field_clear()
Dim obj As Object
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
cmb_reqno.ListIndex = -1
cmb_reqno.clear
'cmb_addendumby.ListIndex = -1
'cmb_addendumby.Text = ""
cmb_enterby.ListIndex = -1
cmb_enterby.Text = ""
'cmb_provby.ListIndex = -1
'cmb_provby.Text = ""
cmb_finalby.ListIndex = -1
'cmb_finalby.Text = ""
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = ">??/#####"
txt_hlaa.Text = ""
txt_hlab.Text = ""
txt_hlabw4.Text = ""
txt_hlac.Text = ""
txt_hladr.Text = ""
txt_hladr52.Text = ""
txt_hladqlocus.Text = ""
txt_hlab27.Text = ""
txt_dnadrb1.Text = ""
txt_dnadqa1.Text = ""
txt_dnadqb1.Text = ""
txt_dnadrb1sub.Text = ""
TXT_IMPRESSION.Text = ""
txt_comments.Text = ""
txt_name.Text = ""
txt_age.Text = ""
txt_sex.Text = ""
txt_category.Text = ""
txt_status.Text = ""
txt_modcode.Text = ""
txt_pridoc.Text = ""
txt_refdoc.Text = ""
txt_clndiag.Text = ""
txt_relpat.Text = ""
txt_nameofdonar.Text = ""
txt_cytology.Text = ""

    txt_dnaDRB345.Text = ""
            
             txt_hlaa_SERO.Text = ""
            txt_hlab_SERO.Text = ""
            txt_hlac_SERO.Text = ""
            txt_dnadrb1_SERO.Text = ""
            txt_dnadqb1_SERO.Text = ""
            txt_dnaDRB345_SERO.Text = ""


End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In hlatyping_report.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    'txt_addendum.Enabled = True
    cmb_enterby.Enabled = True
    cmb_finalby.Enabled = True
   ' cmb_provby.Enabled = True
    'cmb_addendumby.Enabled = True
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_dept.Enabled = True
    cmb_reqno.Enabled = True
    txt_hlaa.Enabled = True
    txt_hlab.Enabled = True
    txt_hlabw4.Enabled = True
    txt_hlac.Enabled = True
    txt_hladr.Enabled = True
    txt_hladqlocus.Enabled = True
    txt_hlab27.Enabled = True
    txt_hladr52.Enabled = True
    txt_dnadrb1.Enabled = True
    txt_dnadqa1.Enabled = True
    txt_dnadqb1.Enabled = True
    txt_dnadrb1sub.Enabled = True
    TXT_IMPRESSION.Enabled = True
    txt_comments.Enabled = True
      txt_dnaDRB345.Enabled = True

    txt_hlaa_SERO.Enabled = True
    txt_hlab_SERO.Enabled = True
    txt_hlac_SERO.Enabled = True
    txt_dnadrb1_SERO.Enabled = True
    txt_dnadqb1_SERO.Enabled = True
    txt_dnaDRB345_SERO.Enabled = True
    
    
End If

If str1 = "DIS" Then
    'txt_addendum.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
    'cmb_provby.Enabled = False
    'cmb_addendumby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    
    txt_hlaa.Enabled = False
    txt_hlab.Enabled = False
    txt_hlabw4.Enabled = False
    txt_hlac.Enabled = False
    txt_hladr.Enabled = False
    txt_hladqlocus.Enabled = False
    txt_hlab27.Enabled = False
    txt_hladr52.Enabled = False
    txt_dnadrb1.Enabled = False
    txt_dnadqa1.Enabled = False
    txt_dnadqb1.Enabled = False
    txt_dnadrb1sub.Enabled = False
    TXT_IMPRESSION.Enabled = False
    txt_comments.Enabled = False
    
    
     txt_dnaDRB345.Enabled = False

    txt_hlaa_SERO.Enabled = False
    txt_hlab_SERO.Enabled = False
    txt_hlac_SERO.Enabled = False
    txt_dnadrb1_SERO.Enabled = False
    txt_dnadqb1_SERO.Enabled = False
    txt_dnaDRB345_SERO.Enabled = False
    
    
End If
If str1 = "VIEW" Then
    'txt_addendum.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
    'cmb_provby.Enabled = False
    'cmb_addendumby.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    
    txt_hlaa.Enabled = False
    txt_hlab.Enabled = False
    txt_hlabw4.Enabled = False
    txt_hlac.Enabled = False
    txt_hladr.Enabled = False
    txt_hladqlocus.Enabled = False
    txt_hlab27.Enabled = False
    txt_hladr52.Enabled = False
    txt_dnadrb1.Enabled = False
    txt_dnadqa1.Enabled = False
    txt_dnadqb1.Enabled = False
    txt_dnadrb1sub.Enabled = False
    TXT_IMPRESSION.Enabled = False
    txt_comments.Enabled = False
    
    
     txt_dnaDRB345.Enabled = False

    txt_hlaa_SERO.Enabled = False
    txt_hlab_SERO.Enabled = False
    txt_hlac_SERO.Enabled = False
    txt_dnadrb1_SERO.Enabled = False
    txt_dnadqb1_SERO.Enabled = False
    txt_dnaDRB345_SERO.Enabled = False
End If
If str1 = "MODIFY" Then
    'txt_addendum.Enabled = False
    cmb_enterby.Enabled = True
    cmb_finalby.Enabled = True
    'cmb_provby.Enabled = True
    'cmb_addendumby.Enabled = False
    'cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    
    txt_hlaa.Enabled = True
    txt_hlab.Enabled = True
    txt_hlabw4.Enabled = True
    txt_hlac.Enabled = True
    txt_hladr.Enabled = True
    txt_hladqlocus.Enabled = True
    txt_hlab27.Enabled = True
    txt_hladr52.Enabled = True
    txt_dnadrb1.Enabled = True
    txt_dnadqa1.Enabled = True
    txt_dnadqb1.Enabled = True
    txt_dnadrb1sub.Enabled = True
    TXT_IMPRESSION.Enabled = True
    txt_comments.Enabled = True
    
    
     txt_dnaDRB345.Enabled = True

    txt_hlaa_SERO.Enabled = True
    txt_hlab_SERO.Enabled = True
    txt_hlac_SERO.Enabled = True
    txt_dnadrb1_SERO.Enabled = True
    txt_dnadqb1_SERO.Enabled = True
    txt_dnaDRB345_SERO.Enabled = True
    
End If
If str1 = "ADDENDUM" Then
   ' txt_addendum.Enabled = True
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = False
    cmb_provby.Enabled = False
   ' cmb_addendumby.Enabled = True
'    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    txt_hlaa.Enabled = False
    txt_hlab.Enabled = False
    txt_hlabw4.Enabled = False
    txt_hlac.Enabled = False
    txt_hladr.Enabled = False
    txt_hladqlocus.Enabled = False
    txt_hlab27.Enabled = False
    txt_hladr52.Enabled = False
    txt_dnadrb1.Enabled = False
    txt_dnadqa1.Enabled = False
    txt_dnadqb1.Enabled = False
    txt_dnadrb1sub.Enabled = False
    TXT_IMPRESSION.Enabled = False
    txt_comments.Enabled = False
    
    
    
       txt_dnaDRB345.Enabled = False

    txt_hlaa_SERO.Enabled = False
    txt_hlab_SERO.Enabled = False
    txt_hlac_SERO.Enabled = False
    txt_dnadrb1_SERO.Enabled = False
    txt_dnadqb1_SERO.Enabled = False
    txt_dnaDRB345_SERO.Enabled = False
End If
End Sub

Public Sub STAFF()
On Error GoTo MSG:
If NEWFLAG = True Then
    NEWFLAG = False
    
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
        If SELDREQSTAFFD1RS1("RELATION") <> "STAFF" Then
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
        txt_refdoc.Text = Trim(REFDR_RS("REFBYDOC"))
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
        txt_ward.Text = Trim(selma104_rs("WARD_NAME"))
        txt_bedno.Text = Trim(selma104_rs("bed_no"))
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
    If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            SELDTMNO_RS.MoreResults
            SELDTMNO_RS.Requery
        End If
        If SELDTMNO_RS.RowCount > 0 Then
            txt_cytology.Text = "" & SELDTMNO_RS("DTMNO")
        End If
    If selpathdrequisitn_rs.RowCount > 0 Then
        'txt_cytology.Text = Trim(selpathdrequisitn_rs("PATHNO"))
        'TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
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
'     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
     seldrequisitn_rs.MoveNext
     Wend
'     txt_prepathreqno.Text = prpathno
    End If
    'txt_addendum.Enabled = False
    'cmb_addendumby.Enabled = False
'    If seldissecentrflag = False Then
'            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
'            seldissecentrflag = True
'        Else
'            seldissecentr_rs.MoreResults
'            seldissecentr_rs.Requery
'        End If
'        If seldissecentr_rs.RowCount > 0 Then
'            cmb_enterby.clear
'            While Not seldissecentr_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr_rs("CCNAME"))
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
        'If seldissecprov_rs.RowCount > 0 Then
            'cmb_provby.clear
            'While Not seldissecprov_rs.EOF
                'cmb_provby.AddItem Trim(seldissecprov_rs("CCNAME"))
                'seldissecprov_rs.MoveNext
            'Wend
        'End If
'        If seldissecfinflag = False Then
'            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
'            seldissecfinflag = True
'        Else
'            seldissecfin_rs.MoreResults
'            seldissecfin_rs.Requery
'        End If
        'If seldissecfin_rs.RowCount > 0 Then
            'cmb_finalby.clear
            'While Not seldissecfin_rs.EOF
                'cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
                'seldissecfin_rs.MoveNext
            'Wend
        'End If
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

    '''***DONOR NAME & RELATION
        
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If Len(Trim(selpathdrequisitn_rs("PREPATHNO"))) > 0 Then
            txt_relpat.Text = Trim(selpathdrequisitn_rs("PREPATHNO"))
        End If
        If Len(Trim(selpathdrequisitn_rs("NMETERIALS"))) > 0 Then
            txt_nameofdonar.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        End If
End If
If viewflag = True Then

    viewflag = False
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
        If SELDREQSTAFFD1RS1("RELATION") <> "STAFF" Then
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
        txt_refdoc.Text = Trim(REFDR_RS("REFBYDOC"))
        If REFDR_RS("COMMITFLAG") = "Y" Then
            MsgBox "Final Report is Generated, You cann't Modify the Final Report,You can Add ADDENDUMS Only"
            cmd_modify.Enabled = False
            ADDENDUMFALG = True
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
        txt_ward.Text = Trim(selma104_rs("WARD_NAME"))
        txt_bedno.Text = Trim(selma104_rs("bed_no"))
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
    If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
            SELDTMNO_QR(0) = MaskEdBox1.FormattedText
            SELDTMNO_RS.MoreResults
            SELDTMNO_RS.Requery
        End If
        If SELDTMNO_RS.RowCount > 0 Then
            txt_cytology.Text = "" & SELDTMNO_RS("DTMNO")
        End If
    If selpathdrequisitn_rs.RowCount > 0 Then
        'TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
        'txt_cytology.Text = selpathdrequisitn_rs("PATHNO")
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
                'txt_sites.Text = SELMA301RS("SITEDESC")
            End If
        End If
        If IsNull(selpathdrequisitn_rs("LATERALITS")) Then
        Else
        'txt_laterality.Text = selpathdrequisitn_rs("LATERALITS")
        End If
        'txt_natmeterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        'txt_opprocedure.Text = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
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
    
    '''***DONOR NAME & RELATION
        
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If Len(selpathdrequisitn_rs("PREPATHNO")) > 0 Then
            txt_relpat.Text = Trim(selpathdrequisitn_rs("PREPATHNO"))
        End If
        If Len(selpathdrequisitn_rs("NMETERIALS")) > 0 Then
            txt_nameofdonar.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
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
'              txt_frosecrep.Text = SELDFREPORT_RS("FSPREPORT").Value
         End If
     Else
'        txt_frosecrep.Text = ""
     End If
         
    
       If SELDFREPORT_RS.RowCount > 0 Then
            If SELDFREPORT_RS.EOF = False Then
'                txt_immhischerep.Text = SELDFREPORT_RS("ihreport").Value
            End If
        Else
'            txt_immhischerep.Text = ""
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
         If SELDFREPORT_RS.EOF = False Then
'            cmb_grossby.Text = SELDFREPORT_RS("GROSSEDBY")
         Else
'            cmb_grossby.Text = ""
         End If


    If REFDR_RS.RowCount > 0 Then
        cmb_enterby.Text = "" & Trim(REFDR_RS("ENTERBY"))
        'cmb_provby.Text = Trim(REFDR_RS("PREPORTBY"))
        If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
            'cmb_finalby.Text = ""
        Else
            'cmb_finalby.Text = Trim(REFDR_RS("FREPORTBY"))
        End If
    End If

'    If SELDIMPRESIONFLAG = False Then
'            Set seldimpresion_rs = seldimpresion_qr.OpenResultset(2, 3)
'            seldimpresion_qr(0) = req_chglen(Trim(cmb_reqno.Text))
'            SELDIMPRESIONFLAG = True
'        Else
'            seldimpresion_qr(0) = req_chglen(Trim(cmb_reqno.Text))
'            seldimpresion_rs.MoreResults
'            seldimpresion_rs.Requery
'        End If
'    If seldimpresion_rs.RowCount > 0 Then
'           ' txt_impfindiag.Text = seldimpresion_rs("impression")
''            cmb_fsd.AddItem seldimpresion_rs("FDESCRY")
''            cmb_fsd.Text = seldimpresion_rs("FDESCRY")
'            'cmb_t.Text = seldimpresion_rs("T")
'            'txt_size.Text = seldimpresion_rs("TSIZE")
''            CMB_N.Text = seldimpresion_rs("N")
'            'cmb_m.Text = seldimpresion_rs("M")
'            'CMB_ICDH.AddItem seldimpresion_rs("FINALDIAG")
'            'CMB_ICDH.Text = seldimpresion_rs("FINALDIAG")
'    End If
       
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
'        SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'        SELDADDENDUM_RS.MoreResults
'        SELDADDENDUM_RS.Requery
'        If SELDADDENDUM_RS.RowCount > 0 Then
'            While Not SELDADDENDUM_RS.EOF
'                ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
'                cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
'                cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
'            SELDADDENDUM_RS.MoveNext
'            Wend
'            txt_addendum.Text = Trim(ADDENDUM)
'        End If
'    'End If

    If SELHLAREPORTFLAG = False Then
            SELHLAREPORTQR(0) = req_chglen(cmb_reqno.Text)
            Set SELHLAREPORTRS = SELHLAREPORTQR.OpenResultset(1, 3)
            SELHLAREPORTFLAG = True
        Else
            SELHLAREPORTQR(0) = req_chglen(cmb_reqno.Text)
            SELHLAREPORTRS.MoreResults
            SELHLAREPORTRS.Requery
            
        End If
        If SELHLAREPORTRS.RowCount > 0 Then
            txt_hlaa.Text = " " & SELHLAREPORTRS("HLAALOCUS")
            txt_hlab.Text = " " & SELHLAREPORTRS("HLABLOCUS")
            txt_hlabw4.Text = " " & SELHLAREPORTRS("HLABW4BW6")
            txt_hlac.Text = " " & SELHLAREPORTRS("HLACLOCUS")
            txt_hladr.Text = " " & SELHLAREPORTRS("HLADR")
            txt_hladqlocus.Text = " " & SELHLAREPORTRS("HLADQLOCUS")
            txt_hlab27.Text = " " & SELHLAREPORTRS("HLAB27")
            txt_hladr52.Text = " " & SELHLAREPORTRS("HLADR5253")
            txt_dnadrb1.Text = " " & SELHLAREPORTRS("DNADRB1LR")
            txt_dnadqa1.Text = " " & SELHLAREPORTRS("DNADQA1")
            txt_dnadqb1.Text = " " & SELHLAREPORTRS("DNADQB1")
            txt_dnadrb1sub.Text = " " & SELHLAREPORTRS("DNADRB1SIA")
            TXT_IMPRESSION.Text = " " & SELHLAREPORTRS("IMPRESSION")
            txt_comments.Text = " " & SELHLAREPORTRS("REMARKS")
             txt_dnaDRB345.Text = " " & SELHLAREPORTRS("DRB345")
            
             txt_hlaa_SERO.Text = " " & SELHLAREPORTRS("SERO1")
            txt_hlab_SERO.Text = " " & SELHLAREPORTRS("SERO2")
            txt_hlac_SERO.Text = " " & SELHLAREPORTRS("SERO3")
            txt_dnadrb1_SERO.Text = " " & SELHLAREPORTRS("SERO4")
            txt_dnadqb1_SERO.Text = " " & SELHLAREPORTRS("SERO5")
            txt_dnaDRB345_SERO.Text = " " & SELHLAREPORTRS("SERO6")
            
        End If
        
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
DIMPRE.Content.Text = ""
DIMPRE.Content.Text = txt_impdiag.Text
WIMPRE.Caption = "Diagnosis"
WIMPRE.WindowState = wdWindowStateMaximize
DIMPRE.Content.Text = txt_impdiag.Text
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
''Private Sub txt_report_GotFocus()
''On Error GoTo m3:
''WREPORT.Visible = False
''Set DREPORT = WREPORT.ActiveDocument
''DREPORT.Content = ""
''DREPORT.Content = txt_report
''WREPORT.Caption = "Report"
''WREPORT.WindowState = wdWindowStateMaximize
''DREPORT.Content = txt_report
''WREPORT.Visible = True
''WREPORT.Activate
''m3:
''On Error Resume Next:
''Set WREPORT = CreateObject("word.Application")
''WREPORT.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
''WREPORT.Caption = "Report"
''WREPORT.WindowState = wdWindowStateMaximize
''Set DREPORT = WREPORT.ActiveDocument
''DREPORT.Content = txt_report.Text
''WREPORT.Visible = True
''WREPORT.Activate
''End Sub
''Private Sub txt_report_LostFocus()
''On Error GoTo m4:
''Set DREPORT = WREPORT.ActiveDocument
''txt_report.Text = DREPORT.Content
''DREPORT.Content = ""
''DREPORT.SaveAs "C:\fronzen.DOC"
''DREPORT.Close
''WREPORT.Quit
''Exit Sub
''m4:
''MsgBox Err.Description
''End Sub
Private Sub TextEnable()
txt_hlaa.Enabled = True
    txt_hlab.Enabled = True
    txt_hlabw4.Enabled = True
    txt_hlac.Enabled = True
    txt_hladr.Enabled = True
    txt_hladqlocus.Enabled = True
    txt_hlab27.Enabled = True
    txt_hladr52.Enabled = True
    txt_dnadrb1.Enabled = True
    txt_dnadqa1.Enabled = True
    txt_dnadqb1.Enabled = True
    txt_dnadrb1sub.Enabled = True
    TXT_IMPRESSION.Enabled = True
    txt_comments.Enabled = True
End Sub





