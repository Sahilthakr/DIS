VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form cytogenetic_report 
   Caption         =   "Cancer Cytogentics Laboratory"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   10905
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmAn2 
      Height          =   1815
      Left            =   120
      TabIndex        =   120
      Top             =   7440
      Width           =   11535
      Begin VB.ComboBox txtAn2Hd1 
         Height          =   315
         Left            =   120
         TabIndex        =   126
         Top             =   120
         Width           =   6735
      End
      Begin VB.TextBox txtData2 
         Height          =   285
         Left            =   960
         TabIndex        =   122
         Top             =   720
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox txtAn2Hd2 
         Height          =   285
         Left            =   6960
         TabIndex        =   121
         Top             =   120
         Width           =   4335
      End
      Begin MSFlexGridLib.MSFlexGrid msfAn2 
         Height          =   1335
         Left            =   120
         TabIndex        =   123
         Top             =   480
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2355
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
      End
   End
   Begin VB.Frame frmAn1 
      Height          =   1815
      Left            =   120
      TabIndex        =   116
      Top             =   7440
      Width           =   11535
      Begin VB.ComboBox txtAn1Hd1 
         Height          =   315
         Left            =   120
         TabIndex        =   125
         Top             =   120
         Width           =   6735
      End
      Begin VB.TextBox txtAn1Hd2 
         Height          =   285
         Left            =   6960
         TabIndex        =   118
         Top             =   120
         Width           =   4215
      End
      Begin VB.TextBox txtData 
         Height          =   280
         Left            =   1800
         TabIndex        =   117
         Top             =   720
         Width           =   2500
      End
      Begin MSFlexGridLib.MSFlexGrid msfAn1 
         Height          =   1335
         Left            =   120
         TabIndex        =   119
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2355
         _Version        =   393216
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
      End
   End
   Begin VB.Frame frmAn3 
      Height          =   1815
      Left            =   120
      TabIndex        =   112
      Top             =   7440
      Width           =   11535
      Begin VB.ComboBox txtAn3Hd1 
         Height          =   315
         Left            =   120
         TabIndex        =   124
         Top             =   120
         Width           =   6735
      End
      Begin VB.TextBox txtAn3Hd2 
         Height          =   285
         Left            =   6960
         TabIndex        =   115
         Top             =   120
         Width           =   4335
      End
      Begin VB.TextBox txtData3 
         Height          =   285
         Left            =   1080
         TabIndex        =   113
         Top             =   720
         Width           =   3615
      End
      Begin MSFlexGridLib.MSFlexGrid msfAn3 
         Height          =   1335
         Left            =   120
         TabIndex        =   114
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2355
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
      End
   End
   Begin VB.Frame Fish_modality 
      Caption         =   "Fluorescence In Situ Hybridization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   0
      TabIndex        =   73
      Top             =   3720
      Width           =   11745
      Begin VB.ComboBox cmb_methodFI 
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
         Left            =   840
         TabIndex        =   102
         Top             =   240
         Width           =   10575
      End
      Begin VB.CheckBox fish_method3 
         Caption         =   "Ploidy Analysis by Chromosome Counting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   96
         Top             =   600
         Width           =   11055
      End
      Begin VB.TextBox txt_fresult 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2355
         TabIndex        =   95
         Top             =   1650
         Width           =   9090
      End
      Begin VB.ComboBox cmb_probe 
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
         Left            =   2355
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   3090
         Width           =   9135
      End
      Begin VB.TextBox txt_test 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2355
         TabIndex        =   11
         Top             =   2760
         Width           =   9105
      End
      Begin VB.CheckBox Fish_method2 
         Caption         =   "PHA stimulated lymphocyte culture, Fluorescence in situ hybridization on interphase and metaphase cells"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   7
         Top             =   915
         Width           =   11055
      End
      Begin VB.TextBox txt_no_of_cell 
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
         Height          =   315
         Left            =   2355
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1320
         Width           =   645
      End
      Begin VB.TextBox FIsh_txt_inter 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2355
         TabIndex        =   9
         Top             =   2040
         Width           =   9090
      End
      Begin VB.TextBox fish_txt_comments 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2355
         TabIndex        =   10
         Top             =   2385
         Width           =   9090
      End
      Begin VB.CheckBox fish_method1 
         Caption         =   $"cytogenetic_report.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11760
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   10965
      End
      Begin VB.Label Label12 
         Caption         =   "Method"
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
         Left            =   120
         TabIndex        =   101
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Result"
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
         Left            =   1650
         TabIndex        =   94
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label Label5 
         Caption         =   "Test"
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
         Left            =   1815
         TabIndex        =   78
         Top             =   2895
         Width           =   420
      End
      Begin VB.Label lbl_meta 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Cells Analyzed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   77
         Top             =   1440
         Width           =   2205
      End
      Begin VB.Label lbl_Inter 
         Caption         =   "Probes/Probe panel used"
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
         Left            =   120
         TabIndex        =   76
         Top             =   3240
         Width           =   2250
      End
      Begin VB.Label lbl_commets 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1305
         TabIndex        =   75
         Top             =   2550
         Width           =   915
      End
      Begin VB.Label lbl_Interpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Interpretation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   74
         Top             =   2040
         Width           =   2115
      End
   End
   Begin VB.ComboBox cmb_cyspreinv 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5160
      TabIndex        =   90
      Top             =   9840
      Width           =   4830
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
      ItemData        =   "cytogenetic_report.frx":008F
      Left            =   11520
      List            =   "cytogenetic_report.frx":0091
      TabIndex        =   82
      Top             =   9840
      Width           =   2670
   End
   Begin VB.ComboBox cmb_enteryby 
      Height          =   315
      Left            =   1080
      TabIndex        =   81
      Top             =   9840
      Width           =   2430
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame common 
      ForeColor       =   &H8000000A&
      Height          =   3075
      Left            =   45
      TabIndex        =   0
      Top             =   600
      Width           =   11730
      Begin VB.TextBox txt_clinicaldetail 
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
         Left            =   7725
         TabIndex        =   89
         Top             =   2400
         Width           =   3810
      End
      Begin VB.TextBox txt_specimentype 
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
         Height          =   300
         Left            =   2400
         TabIndex        =   86
         Top             =   2760
         Width           =   9135
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   85
         Top             =   1680
         Width           =   585
      End
      Begin VB.TextBox txt_rec_on 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8940
         TabIndex        =   59
         Top             =   2040
         Width           =   2595
      End
      Begin VB.TextBox txt_specimen 
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
         Left            =   8940
         TabIndex        =   57
         Top             =   1680
         Width           =   2595
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   360
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
            TabIndex        =   54
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
            TabIndex        =   53
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
         TabIndex        =   13
         Top             =   480
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
         ItemData        =   "cytogenetic_report.frx":0093
         Left            =   6240
         List            =   "cytogenetic_report.frx":009D
         TabIndex        =   2
         Top             =   120
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
         TabIndex        =   22
         Top             =   2400
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
         TabIndex        =   21
         Top             =   2040
         Width           =   3855
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
         TabIndex        =   20
         Top             =   1680
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
         TabIndex        =   19
         Top             =   1680
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
         TabIndex        =   18
         Top             =   1320
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
         TabIndex        =   17
         Top             =   1320
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
         TabIndex        =   16
         Top             =   1320
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
         TabIndex        =   15
         Top             =   1320
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
         TabIndex        =   14
         Top             =   960
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
         Top             =   480
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
         Top             =   120
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
         ItemData        =   "cytogenetic_report.frx":00C2
         Left            =   2400
         List            =   "cytogenetic_report.frx":00C4
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   480
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
      Begin VB.Label Label6 
         Caption         =   "Clinical Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6360
         TabIndex        =   88
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label Label4 
         Caption         =   "Specimen Type"
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
         Left            =   855
         TabIndex        =   87
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label lbl_bed 
         Caption         =   "Bed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6360
         TabIndex        =   84
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lbl_rec_on 
         Caption         =   "Received On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7680
         TabIndex        =   58
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label lbl_specimen 
         Caption         =   "Specimen Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7440
         TabIndex        =   56
         Top             =   1680
         Width           =   1560
      End
      Begin VB.Label lbl_cytologyno 
         AutoSize        =   -1  'True
         Caption         =   "Lab No."
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
         TabIndex        =   51
         Top             =   480
         Width           =   690
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
         TabIndex        =   47
         Top             =   2400
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
         TabIndex        =   46
         Top             =   2160
         Width           =   2130
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
         TabIndex        =   45
         Top             =   120
         Width           =   720
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
         Left            =   4440
         TabIndex        =   44
         Top             =   1680
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
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   2115
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
         TabIndex        =   42
         Top             =   1320
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
         TabIndex        =   41
         Top             =   1320
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
         TabIndex        =   40
         Top             =   1320
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
         Left            =   1800
         TabIndex        =   39
         Top             =   1320
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
         TabIndex        =   38
         Top             =   960
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
         TabIndex        =   37
         Top             =   480
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
         TabIndex        =   36
         Top             =   600
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
         TabIndex        =   35
         Top             =   120
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
         TabIndex        =   34
         Top             =   120
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
      Left            =   10440
      TabIndex        =   33
      Top             =   240
      Width           =   1125
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
      Left            =   10440
      TabIndex        =   32
      Top             =   0
      Width           =   1125
   End
   Begin VB.Frame COMMAND 
      Height          =   495
      Left            =   240
      TabIndex        =   31
      Top             =   10200
      Width           =   11715
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
         Height          =   255
         Left            =   7320
         MouseIcon       =   "cytogenetic_report.frx":00C6
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   120
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
         Height          =   255
         Left            =   10320
         MouseIcon       =   "cytogenetic_report.frx":03D0
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   120
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
         Height          =   255
         Left            =   8880
         MouseIcon       =   "cytogenetic_report.frx":06DA
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   120
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
         Height          =   255
         Left            =   5880
         MouseIcon       =   "cytogenetic_report.frx":09E4
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   120
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
         Height          =   255
         Left            =   4440
         MouseIcon       =   "cytogenetic_report.frx":0CEE
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   120
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
         Height          =   255
         Left            =   3000
         MouseIcon       =   "cytogenetic_report.frx":0FF8
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   120
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
         Height          =   255
         Left            =   1560
         MouseIcon       =   "cytogenetic_report.frx":1302
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   120
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
         Height          =   255
         Left            =   120
         MouseIcon       =   "cytogenetic_report.frx":160C
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame frm_Conventional 
      Caption         =   "Conventional Cytogenetics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   0
      TabIndex        =   60
      Top             =   3720
      Width           =   11760
      Begin VB.ComboBox cmb_CCProbes 
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
         Left            =   3960
         TabIndex        =   108
         Top             =   3120
         Width           =   7575
      End
      Begin VB.ComboBox cmb_CCTest 
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
         Left            =   600
         TabIndex        =   106
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ComboBox cmb_methodCC 
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
         Left            =   960
         TabIndex        =   104
         Top             =   480
         Width           =   10695
      End
      Begin VB.TextBox txt_result 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3960
         TabIndex        =   93
         Top             =   2040
         Width           =   7575
      End
      Begin VB.OptionButton opt_chromosome 
         Caption         =   "Chromosomal Breakage(Fragility) Analysis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7800
         TabIndex        =   63
         Top             =   100
         Width           =   3900
      End
      Begin VB.TextBox txt_no_of_mata 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3930
         MaxLength       =   3
         TabIndex        =   67
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txt_Interpre 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3960
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   2400
         Width           =   7575
      End
      Begin VB.TextBox txt_comments 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         TabIndex        =   69
         Top             =   2760
         Width           =   7575
      End
      Begin VB.OptionButton opt_chromo 
         Caption         =   "Chromosome Analysis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   61
         Top             =   120
         Width           =   2610
      End
      Begin VB.CheckBox Chk_method1 
         Caption         =   "Direct Harvesting of Bone Marrow Aspirate/Peripheral Blood, GTG-Banding, Chromosome Analysis using PSI/Vysis Karyotype Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11760
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   11040
      End
      Begin VB.CheckBox chk_method3 
         Caption         =   $"cytogenetic_report.frx":1916
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   65
         Top             =   840
         Width           =   11415
      End
      Begin VB.OptionButton opt_karyotype 
         Caption         =   "Karyotype Analysis"
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
         Left            =   5400
         TabIndex        =   62
         Top             =   120
         Width           =   2040
      End
      Begin VB.CheckBox chk_method2 
         Caption         =   "PHA stimulated lymphocyte culture, GTG-Banding, Chromosome Analysis using PSI Karyotype Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   66
         Top             =   1320
         Width           =   11055
      End
      Begin VB.Label Label15 
         Caption         =   "Probes"
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
         Left            =   3240
         TabIndex        =   107
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Test"
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
         Left            =   120
         TabIndex        =   105
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Method"
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
         Left            =   240
         TabIndex        =   103
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   92
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2955
         TabIndex        =   72
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Interpretation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2685
         TabIndex        =   71
         Top             =   2400
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "No. of metaphase cells analyzed/Karyotyped"
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
         Left            =   75
         TabIndex        =   70
         Top             =   1680
         Width           =   3825
      End
   End
   Begin VB.Frame frame_addendum 
      Height          =   495
      Left            =   240
      TabIndex        =   98
      Top             =   9240
      Width           =   11775
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   9000
         TabIndex        =   111
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   110
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label lbl_addendum 
         Caption         =   "ADDENDUM"
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
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbl_addendumby 
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
         Height          =   255
         Left            =   7800
         TabIndex        =   99
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2055
      Left            =   0
      TabIndex        =   109
      Top             =   7200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3625
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Analysis 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Analysis 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Analysis 3"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "09/11/2009"
      Height          =   255
      Left            =   720
      TabIndex        =   97
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Previous Investigation "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3960
      TabIndex        =   91
      Top             =   9720
      Width           =   1170
   End
   Begin VB.Label lblicu 
      Caption         =   "ICU"
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
      Left            =   930
      TabIndex        =   83
      Top             =   450
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl_Analysed_by 
      Caption         =   "Analyzed and Reported by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10200
      TabIndex        =   80
      Top             =   9840
      Width           =   1365
   End
   Begin VB.Label lbl_enterby 
      Caption         =   "Enterd by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   79
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2370
      TabIndex        =   55
      Top             =   0
      Width           =   6990
   End
   Begin VB.Label lbl_finnedaspcyt 
      AutoSize        =   -1  'True
      Caption         =   "CANCER CYTOGENETICS LABORATORY REPORT"
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
      Left            =   2460
      TabIndex        =   50
      Top             =   240
      Width           =   7020
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
      Left            =   9840
      TabIndex        =   49
      Top             =   240
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
      Left            =   9840
      TabIndex        =   48
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "cytogenetic_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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
Dim selreqmoda_qr As rdoQuery
Dim selreqmoda_RS As rdoResultset
Dim seldrequisitn_FLAG As Boolean
Dim selreqnno_qr As rdoQuery
Dim selreqnno_rs As rdoResultset
Dim selreqnno_flg As Boolean

Dim SELPROBE_RS As rdoResultset
Dim SELPROBE_QR As rdoQuery

Dim INSREPORT_QR As rdoQuery
Dim UPDREPORT_QR As rdoQuery
Dim SELREPORT_QR As rdoQuery
Dim SELREPORT_RS As rdoResultset
Dim SELREPORT_FLAG As Boolean

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
Dim d_frozen As Document, d_grossdesc As Document, d_micscop As Document, d_imp As Document, d_immun As Object, d_adden As Object, D As Object, D1 As Object
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

Dim CYTOGCLDTA_QR As rdoQuery
Dim CYTOGCLDTA_RS As rdoResultset
Dim CYTOGCLDTA_FLG As Boolean

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

Dim UPDATERTIMELOG As rdoQuery 'Shraddha Pandit 16/01/2009
Dim UPDTADNDMTIME_QR As rdoQuery 'Shraddha Pandit 25/02/2009
Dim C_S As String 'Shraddha Pandit 25/08/2009
Dim S_A As String 'Shraddha Pandit 25/08/2009

'***** Shraddha Pandit 06/11/09 *****
Dim INS_CYTOGENET_QR As rdoQuery
Dim sel_rreq_qr As rdoQuery
Dim sel_rreq_rs As rdoResultset

Dim sel_rtime_qr As rdoQuery
Dim sel_rtime_rs As rdoResultset

Dim UPD_CYTOGENET_QR As rdoQuery
Dim sel_cytogenet_qr As rdoQuery
Dim sel_cytogenet_rs As rdoResultset

Dim sel_dreq_qr As rdoQuery
Dim sel_dreq_rs As rdoResultset

Dim UPD_CYTOGENETmod_QR As rdoQuery
'******* Shraddha 02/12/09 *******
Dim sel_cytogmethod_qr As rdoQuery
Dim sel_cytogmethod_rs As rdoResultset

'******* Shraddha 07/12/09 *******
Dim sel_cytogtest_qr As rdoQuery
Dim sel_cytogtest_rs As rdoResultset

Dim sel_CYTOGPROBS_qr As rdoQuery
Dim sel_CYTOGPROBS_rs As rdoResultset
'*********************************
Dim chk_Adfld As Boolean
Dim InsCytog_ETO_qr As rdoQuery
Dim InsCytog_CHIC_qr As rdoQuery
Dim InsCytog_SP_qr As rdoQuery
Dim delCytog As rdoQuery
Dim selCytog As rdoQuery
Dim rsSelCytog As rdoResultset

Dim selHead As rdoQuery


''******SECURITY***********
'Private Sub cmb_cellblock_GotFocus()
'cmb_cellblock.clear
'cmb_cellblock.AddItem "YES"
'cmb_cellblock.AddItem "NO"
'On Error GoTo MSG:
'Dim ind_cbodept As Integer
'If Trim(cmb_cellblock) = "" Then
'    For ind_cbodept = 0 To (cmb_cellblock.ListCount - 1)
'        If Trim(cmb_cellblock.List(ind_cbodept)) = "NO" Then
'            cmb_cellblock.ListIndex = ind_cbodept
'        End If
'    Next ind_cbodept
'    End If
'Exit Sub
'MSG:
'MsgBox Error(Err)
'End Sub

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
MsgBox ERROR(Err)
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
cmb_dept.ListIndex = 0
'Dim ind_cbodept As Integer
'If Trim(cmb_dept) = "" Then
'    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
'        If Trim(cmb_dept.List(ind_cbodept)) = "PATHOLOGY$FZZ" Then
'            cmb_dept.ListIndex = ind_cbodept
'        End If
'    Next ind_cbodept
'    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_modality_Click()

If NEWFLAG = True Then
    If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
        frm_Conventional.Visible = True
        Fish_modality.Visible = False
    End If
    If Trim(Right(cmb_modality.Text, 2)) = "FI" Then
        frm_Conventional.Visible = False
        Fish_modality.Visible = True
        txt_no_of_cell.Text = 200
    End If
End If
Dim MODALITY As String
Dim moda As Integer
Dim l As Integer
On Error GoTo MSG:
If Trim(cmb_modality) = "" Then
    Exit Sub
End If

'moda = InStr(cmb_modality, "$")
'l = Len(cmb_modality.Text)
'txt_modcode.Text = Right(cmb_modality.Text, (l - moda))
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
'
Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
cmb_modality.ListIndex = 0
'Dim ind_cbodept As Integer
'If Trim(cmb_modality) = "" Then
'    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
'        If Trim(cmb_modality.List(ind_cbodept)) = "FNAC$FA" Then
'            cmb_modality.ListIndex = ind_cbodept
'        End If
'    Next ind_cbodept
'    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub


Private Sub cmd_back_Click()
'btnenabl ("INI")
BTNENABL1 ("INI")
FLDENABL ("DIS")
TXT_CLEAR
field_clear
DIS_MAIN.show
Me.Hide
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
'*************** Shraddha Pandit 25/02/09 **********
    frame_addendum.Visible = False
    lbl_addendum.Visible = False
    txt_addendum.Visible = False
    lbl_addendumby.Visible = False
    cmb_addendumby.Visible = False
    ADDENDUMFALG = False
    '****************************************************
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
    'On Error GoTo MSG:
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
    updrrequisitn1_qr(0) = Left(cmb_finalby.Text, 30)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    rdb.CommitTrans
    '*********** Added by Shraddha Pandit on 16/01/2009 ***************
    'COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
    UPDATERTIMELOG(0) = Format(txt_date.Text, "YYYY-MM-DD") 'COMREPDATE
    UPDATERTIMELOG(1) = Format(txt_time.Text, "HH:MM:SS") 'COMREPTIME
    UPDATERTIMELOG(2) = req_chglen(cmb_reqno.Text)
    UPDATERTIMELOG.Execute
    '************* Shraddha 06/11/09 *********
    'PREPORTBY FREPORTBY SAMCOLDATE SAMCOLTIME SAMRECDATE SAMRECTIME COMREPDATE COMREPTIME 8
'    INS_CYTOGENET_QR(13) = sel_rreq_rs("COMREPDATE") 'COMREPDATE
'    INS_CYTOGENET_QR(14) = sel_rreq_rs("COMREPTIME") 'COMREPTIME
    '******************************************************************
    
    '***** Shraddha Pandit 09/11/09 *****
    sel_rtime_qr(0) = Trim(req_chglen(cmb_reqno.Text))
    Set sel_rtime_rs = sel_rtime_qr.OpenResultset(1, 3)
    
    str = "UPDATE DISLIB.CYTOGENET SET SAMCOLDATE=?,SAMCOLTIME=?,SAMRECDATE=?,SAMRECTIME=?," & _
     "COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
    'Commented and added by Shraddha on 29/08/2011
     str = "UPDATE DISLIB.CYTOGENET SET SAMCOLDATE=?,SAMCOLTIME=?,SAMRECDATE=?,SAMRECTIME=?," & _
     "COMREPDATE=?,COMREPTIME=?, FLD1=? WHERE REQNNO=?"
    Set UPD_CYTOGENET_QR = rdb.CreateQuery("", str)
    
    If sel_rtime_rs("SAMCOLDATE") <> "" Or sel_rtime_rs("SAMCOLDATE") <> Null Then
        UPD_CYTOGENET_QR(0) = sel_rtime_rs("SAMCOLDATE") 'SAMCOLDATE
    Else
        UPD_CYTOGENET_QR(0) = Null
    End If
    If sel_rtime_rs("SAMCOLTIME") <> "" Or sel_rtime_rs("SAMCOLTIME") <> Null Then
        UPD_CYTOGENET_QR(1) = sel_rtime_rs("SAMCOLTIME") 'SAMCOLTIME
    Else
        UPD_CYTOGENET_QR(1) = Null
    End If
    If sel_rtime_rs("SAMRECDATE") <> "" Or sel_rtime_rs("SAMRECDATE") <> Null Then
        UPD_CYTOGENET_QR(2) = sel_rtime_rs("SAMRECDATE") 'SAMRECDATE
    Else
        UPD_CYTOGENET_QR(2) = Null
    End If
    If sel_rtime_rs("SAMRECTIME") <> "" Or sel_rtime_rs("SAMRECTIME") <> Null Then
        UPD_CYTOGENET_QR(3) = sel_rtime_rs("SAMRECTIME") 'SAMRECTIME
    Else
        UPD_CYTOGENET_QR(3) = Null
    End If
    
    If sel_rtime_rs("COMREPDATE") <> "" Or sel_rtime_rs("COMREPDATE") <> Null Then
        UPD_CYTOGENET_QR(4) = sel_rtime_rs("COMREPDATE") 'COMREPDATE
    Else
        UPD_CYTOGENET_QR(4) = Null
    End If
    
    If sel_rtime_rs("COMREPTIME") <> "" Or sel_rtime_rs("COMREPTIME") <> Null Then
        UPD_CYTOGENET_QR(5) = sel_rtime_rs("COMREPTIME") 'COMREPTIME
    Else
        UPD_CYTOGENET_QR(5) = Null
    End If
    
    
    UPD_CYTOGENET_QR(5) = ipadd
    
    UPD_CYTOGENET_QR(6) = Trim(req_chglen(cmb_reqno.Text))
    
    
    UPD_CYTOGENET_QR.Execute
    UPD_CYTOGENET_QR.Close
    sel_rtime_rs.Close
    
    '************************************
    
End If
rdb.CommitTrans
'btnenabl ("INI")
BTNENABL1 ("INI")
MsgBox "Record is Committed Successfully."
cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
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

Private Sub cmd_modify_Click()
'btnenabl ("NEW")
''********SECURITY***************************
CURUSRDET_QR(0) = Trim(cmb_enteryby.Text)
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

If ADDENDUMFALG = True Then
    cmd_modify.Enabled = False
    txt_no_of_mata.Enabled = False
    txt_result.Enabled = False
    txt_Interpre.Enabled = False
    txt_comments.Enabled = False
    
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    
    txt_no_of_cell.Enabled = False
    txt_fresult.Enabled = False
    FIsh_txt_inter.Enabled = False
    fish_txt_comments.Enabled = False
    txt_test.Enabled = False
    cmb_probe.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    
Else
    FLDENABL ("MODIFY")
End If
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

sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)

If sel_rreq_rs.RowCount > 0 Then
    If Trim(sel_rreq_rs("reqnstat")) = "F" And Trim(sel_rreq_rs("COMMITFLAG")) = "Y" Then
    ADDENDUMFALG = True
    End If
End If

'If ADDENDUMFALG = True Then
'    txt_addendum.Enabled = True
'    cmb_addendumby.Enabled = True
    
    
    
    
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
If ADDENDUMFALG = True Then
    FLDENABL ("ADDENDUM")
End If
    'txt_addendum.SetFocus
    'flagsave = 0
'End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG:
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
chk_Adfld = False
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

'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
Dim flgEto As Integer
Dim flgCHIC As Integer
Dim flgSP As Integer

flgEto = 0
flgCHIC = 0
flgSP = 0
 'On Error GoTo prtmsg:

CrystalReport1.Reset


DoEvents

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
Dim STRSQL As String
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'CrystalReport1.SelectionFormula = "{RPTCYTOGLF.REQNO} = '" & VAR_REQNNO & "'"
'R_NO = req_chglen(cmb_reqno.Text)
'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & R_NO & "'"
'If Opt_patient.Value = True Then
    'If txt_addendum.Text = "" Then
        'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCY.RPT"
        
        '***** Shraddha 09/11/09 *****
        sel_cytogenet_qr(0) = req_chglen(Trim(cmb_reqno.Text))
        Set sel_cytogenet_rs = sel_cytogenet_qr.OpenResultset(1, 3)
        If Not sel_cytogenet_rs.EOF Then
            'CrystalReport1.ReplaceSelectionFormula ("{CYTOGENET.REQNO} = '" & VAR_REQNNO & "' AND {CYTOGENET.CASENO} = '" & Trim(MaskEdBox1.FormattedText) & "'") ' and {CYTOG_ETO.SRNO} = 1.00 and {CYTOG_SP.SRNO} = 1.00 and {CYTOG_CHIC.SRNO} = 1.00"
            STRSQL = "select * from dislib.cytog_eto where reqno='" & VAR_REQNNO & "' and srno=1"
            Set rsSelCytog = rdb.OpenResultset(STRSQL, 1, 1)
            If rsSelCytog.EOF Then
                flgEto = 0
            Else
                If Trim(rsSelCytog!eto5) <> "" And Trim(rsSelCytog!eto4) <> "" And Trim(rsSelCytog!NOCELLS) <> "" And Trim(rsSelCytog!ANALYSIS) <> "" Then
                    flgEto = 7
                ElseIf Trim(rsSelCytog!eto5) <> "" And Trim(rsSelCytog!eto4) <> "" And Trim(rsSelCytog!NOCELLS) <> "" Then
                    flgEto = 6
                ElseIf Trim(rsSelCytog!eto5) <> "" And Trim(rsSelCytog!eto4) <> "" Then
                    flgEto = 5
                ElseIf Trim(rsSelCytog!eto4) <> "" Then
                    flgEto = 4
                ElseIf Trim(rsSelCytog!ETOAML1) <> "" Then
                    flgEto = 3
                End If
            End If
            
            rsSelCytog.Close
            
            'CrystalReport1.SelectionFormula = "{CYTOGENET.REQNO} = '" & VAR_REQNNO & "'" ' and {CYTOG_ETO.SRNO} = 1.00 and {CYTOG_SP.SRNO} = 1.00 and {CYTOG_CHIC.SRNO} = 1.00"
            STRSQL = "select * from dislib.cytog_chic where reqno='" & VAR_REQNNO & "' and srno=1"
            Set rsSelCytog = rdb.OpenResultset(STRSQL, 1, 1)
            If rsSelCytog.EOF Then
                flgCHIC = 0
            Else
                If Trim(rsSelCytog!CHIC5) <> "" And Trim(rsSelCytog!CHIC4) <> "" And Trim(rsSelCytog!NOCELLS) <> "" And Trim(rsSelCytog!ANALYSIS) <> "" Then
                    flgCHIC = 7
                ElseIf Trim(rsSelCytog!CHIC5) <> "" And Trim(rsSelCytog!CHIC4) <> "" And Trim(rsSelCytog!NOCELLS) <> "" Then
                    flgCHIC = 6
                ElseIf Trim(rsSelCytog!CHIC5) <> "" And Trim(rsSelCytog!CHIC4) <> "" Then
                    flgCHIC = 5
                ElseIf Trim(rsSelCytog!CHIC4) <> "" Then
                    flgCHIC = 4
                ElseIf Trim(rsSelCytog!CHIC3) <> "" Then
                    flgCHIC = 3
                End If
            End If
            rsSelCytog.Close
            
            'CrystalReport1.SelectionFormula = "{CYTOGENET.REQNO} = '" & VAR_REQNNO & "'" ' and {CYTOG_ETO.SRNO} = 1.00 and {CYTOG_SP.SRNO} = 1.00 and {CYTOG_CHIC.SRNO} = 1.00"
            STRSQL = "select * from dislib.cytog_sp where reqno='" & VAR_REQNNO & "' and srno=1"
            Set rsSelCytog = rdb.OpenResultset(STRSQL, 1, 1)
            If rsSelCytog.EOF Then
                flgSP = 0
            Else
                If Trim(rsSelCytog!CHIC5) <> "" And Trim(rsSelCytog!CHIC4) <> "" And Trim(rsSelCytog!NOCELLS) <> "" And Trim(rsSelCytog!ANALYSIS) <> "" Then
                    flgSP = 7
                ElseIf Trim(rsSelCytog!CHIC5) <> "" And Trim(rsSelCytog!CHIC4) <> "" And Trim(rsSelCytog!NOCELLS) <> "" Then
                    flgSP = 6
                ElseIf Trim(rsSelCytog!CHIC5) <> "" And Trim(rsSelCytog!CHIC4) <> "" Then
                    flgSP = 5
                ElseIf Trim(rsSelCytog!CHIC4) <> "" Then
                    flgSP = 4
                ElseIf Trim(rsSelCytog!CHIC3) <> "" Then
                    flgSP = 3
                End If
            End If
            rsSelCytog.Close
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCYGENET.rpt"
            
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIS_CYTOGENET_rev.rpt"
            CrystalReport1.ReplaceSelectionFormula ("{CYTOGEN_lf.REQNO} = '" & VAR_REQNNO & "' AND {CYTOGEN_lf.CASENO} = '" & Trim(MaskEdBox1.FormattedText) & "'") ' and {CYTOG_ETO.SRNO} = 1.00 and {CYTOG_SP.SRNO} = 1.00 and {CYTOG_CHIC.SRNO} = 1.00"
            CrystalReport1.Formulas(16) = "flgeto = '" & flgEto & "'"
            CrystalReport1.Formulas(17) = "flgchic = '" & flgCHIC & "'"
            CrystalReport1.Formulas(18) = "flgsp = '" & flgSP & "'"
            If txt_addendum.Text <> "" Then
            CrystalReport1.Formulas(19) = "flgadd ='1' "
            Else
            CrystalReport1.Formulas(19) = "flgadd ='0' "
            End If
            
            'CrystalReport1.SelectionFormula = ""
            CrystalReport1.ReplaceSelectionFormula ("{CYTOGEN_lf.REQNO} = '" & VAR_REQNNO & "' AND {CYTOGEN_lf.CASENO} = '" & Trim(MaskEdBox1.FormattedText) & "'")
    
            'CrystalReport2.ReportFileName = "C:\DISREPORTSEC\DIS_CYTOGENET_V1.rpt"
            'CrystalReport2.ReportFileName = "C:\DISREPORTSEC\VIDYA.rpt"
            
            'CrystalReport2.SelectionFormula = "{CYTOGENET.REQNO} = '" & VAR_REQNNO & "' AND {CYTOGENET.CASENO} = '" & Trim(MaskEdBox1.FormattedText) & "'"
        Else
            CrystalReport1.SelectionFormula = "{RPTCYTOGLF.REQNO} = '" & VAR_REQNNO & "'"
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCYOLD.RPT"
        End If
        '*****************************
        'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCYOLD.RPT"
        
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        
        sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
        Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
        
        If Trim(sel_rreq_rs("REQNSTAT")) = "T" Then
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(1) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(1) = "ENT = '" & str1 & "'"
        End If
        '********** Shraddha **********
        CrystalReport1.Formulas(12) = "NAM = '" & txt_name.Text & "'"
        
        
        '********** Shraddha 25/08/09 **********
        S_A = Trim(txt_sex.Text) & "/" & Trim(txt_age.Text)
        
        CrystalReport1.Formulas(13) = "SexAge = '" & S_A & "'"
                
        C_S = Trim(txt_category.Text) & "/" & Trim(txt_status.Text)
        
        CrystalReport1.Formulas(14) = "CAT_STS = '" & C_S & "'"
        
        CrystalReport1.Formulas(15) = "LABNO = '" & Trim(txt_cytology.Text) & "'"
        
        'txt_age
        'txt_sex
        'txt_category
        'txt_status
        'txt_cytology PathNo
        
        '***********************************
        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
        
        If sel_rreq_rs("MODALITYCD") = "FI" Then
           If Trim(FIsh_txt_inter.Text) <> "" Then
            CrystalReport1.Formulas(10) = "inter = 'abc'" '& STR1 & "'"
           Else
            CrystalReport1.Formulas(10) = "inter = ''" '& STR1 & "'"
           End If

           If Trim(fish_txt_comments.Text) <> "" Then
            CrystalReport1.Formulas(11) = "comments = 'abc'" '& STR1 & "'"
           Else
            CrystalReport1.Formulas(11) = "comments = ''" '& STR1 & "'"
           End If

           CrystalReport1.Formulas(3) = "PROFIN = '" & "MOLECULAR CYTOGENETICS (FISH) REPORT" & "'"

        Else
            If Trim(txt_Interpre.Text) <> "" Then
                CrystalReport1.Formulas(10) = "inter = 'abc'" '& STR1 & "'"
            Else
                CrystalReport1.Formulas(10) = "inter = ''" '& STR1 & "'"
            End If

            If Trim(txt_comments.Text) <> "" Then
                CrystalReport1.Formulas(11) = "comments = 'abc'" '& STR1 & "'"
            Else
                CrystalReport1.Formulas(11) = "comments = ''" '& STR1 & "'"
            End If
            'CrystalReport1.Formulas(3) = "PROFIN = '" & "CONVENTIONAL CYTOGENETICS REPORT" & "'"
            'Commented and added by Shraddha on 26/11/09
            CrystalReport1.Formulas(3) = "PROFIN = '" & "CONVENTIONAL/MOLECULAR CYTOGENETICS REPORT" & "'"
        End If
        CrystalReport1.Formulas(4) = "cdetails = '" & txt_clinicaldetail.Text & "'"
        CrystalReport1.Formulas(5) = "spstatus = '" & txt_specimen.Text & "'"
        CrystalReport1.Formulas(6) = "PROVDIA = '" & txt_provdiag.Text & "'"
    'Else
    '******************* Shraddha Pandit 26/02/09 ***********************
''    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIS_CYTOGENET.rpt"
''        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
''        If Trim(sel_rreq_rs("REQNSTAT")) = "T" Then
''            str1 = ""
''            str1 = "Entered By " & Trim(cmb_enteryby.Text)
''            CrystalReport1.Formulas(1) = "ENT = '" & str1 & "'"
''        Else
''            CrystalReport1.Formulas(1) = "ENT = '" & str1 & "'"
''        End If
''        '******************** Shraddha ******************************
''        CrystalReport1.Formulas(12) = "NAM = '" & txt_name.Text & "'"
''        '************************************************************
''        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
''        If sel_rreq_rs("MODALITYCD") = "FI" Then
''           If Trim(FIsh_txt_inter.Text) <> "" Then
''            CrystalReport1.Formulas(10) = "inter = 'abc'" '& STR1 & "'"
''           Else
''            CrystalReport1.Formulas(10) = "inter = ''" '& STR1 & "'"
''           End If
''
''           If Trim(fish_txt_comments.Text) <> "" Then
''            CrystalReport1.Formulas(11) = "comments = 'abc'" '& STR1 & "'"
''           Else
''            CrystalReport1.Formulas(11) = "comments = ''" '& STR1 & "'"
''           End If
''
''           CrystalReport1.Formulas(3) = "PROFIN = '" & "MOLECULAR CYTOGENETICS (FISH) REPORT" & "'"
''
''        Else
''            If Trim(txt_Interpre.Text) <> "" Then
''                CrystalReport1.Formulas(10) = "inter = 'abc'" '& STR1 & "'"
''            Else
''                CrystalReport1.Formulas(10) = "inter = ''" '& STR1 & "'"
''            End If
''
''            If Trim(TXT_COMMENTS.Text) <> "" Then
''                CrystalReport1.Formulas(11) = "comments = 'abc'" '& STR1 & "'"
''            Else
''                CrystalReport1.Formulas(11) = "comments = ''" '& STR1 & "'"
''            End If
''            'CrystalReport1.Formulas(3) = "PROFIN = '" & "CONVENTIONAL CYTOGENETICS REPORT" & "'"
''        End If
''        CrystalReport1.Formulas(4) = "cdetails = '" & txt_clinicaldetail.Text & "'"
''        CrystalReport1.Formulas(5) = "spstatus = '" & txt_specimen.Text & "'"
''        CrystalReport1.Formulas(6) = "PROVDIA = '" & txt_provdiag.Text & "'"
''    '********************************************************************
''    CrystalReport1.SelectionFormula = "{CYTOGENET.REQNO} = '" & VAR_REQNNO & "'"  'SHWETA 29-08-2012
    'End If
'Else
'    If SELEMPBASINFOFLAG = False Then
'        SELEMPBASINFO_QR(0) = MaskEdBox1.FormattedText
'        Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
'        SELEMPBASINFOFLAG = True
'    Else
'        SELEMPBASINFO_QR(0) = MaskEdBox1.FormattedText
'        SELEMPBASINFO_RS.MoreResults
'        SELEMPBASINFO_RS.Requery
'    End If
'    If ADDENDUMFALG = False Then
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
'        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'            str1 = ""
'            str1 = "Entered By " & Trim(cmb_enteryby.Text)
'            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
'         Else
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
'        End If
'        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
'        If Trim(txt_report.Text) = "" Then
'            CrystalReport1.Formulas(8) = "fn = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(8) = "fn = '" & "12" & "'"
'        End If
'    Else
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
'        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
'    End If
'    CrystalReport1.Formulas(10) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'    CrystalReport1.Formulas(11) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'    CrystalReport1.Formulas(12) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
'End If
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
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
'If Len(Trim(cmb_provby.Text)) = 0 Then
'    m = MsgBox("The Provisional By field is EMPTY.", vbCritical)
'    cmb_provby.SetFocus
'    Exit Sub
'End If
If Len(Trim(cmb_finalby.Text)) = 0 Then
    m = MsgBox("The To be Finalized field is EMPTY.", vbCritical)
    cmb_finalby.SetFocus
    Exit Sub
End If

       
'  If Len(Trim(txt_addendum.Text)) > 600 Then
'    MsgBox "Addendum Max. of 600 Characters Only.", vbInformation
'    txt_addendum.SetFocus
'    Exit Sub
'  End If

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

Private Sub cmd_view_Click()
On Error GoTo MSG:
'btnenabl ("VIEW")
BTNENABL1 ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR
txt_date.Text = DATE11
txt_time.Text = time1
'&&&&&&&&&&&&&&&&&&&&77 not found hence declared
chk_Adfld = False
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



Private Sub fish_txt_comments_GotFocus()
On Error GoTo m14:
w.Visible = False
Set D = w.ActiveDocument
D.Content = ""
D.Content = fish_txt_comments
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
D.Content = fish_txt_comments
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set D = w.ActiveDocument
D.Content = fish_txt_comments.Text
w.Visible = True
w.Activate
End Sub

Private Sub fish_txt_comments_LostFocus()
On Error GoTo m15:
Set D = w.ActiveDocument
fish_txt_comments.Text = D.Content
'flagimp = False
D.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
D.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub FIsh_txt_inter_GotFocus()
On Error GoTo m14:
w.Visible = False
Set D = w.ActiveDocument
D.Content = ""
D.Content = FIsh_txt_inter
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
D.Content = FIsh_txt_inter.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set D = w.ActiveDocument
D.Content = FIsh_txt_inter.Text
w.Visible = True
w.Activate
End Sub

Private Sub FIsh_txt_inter_LostFocus()
On Error GoTo m15:
Set D = w.ActiveDocument
FIsh_txt_inter.Text = D.Content
'flagimp = False
D.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
D.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description

End Sub

Private Sub Form_Load()
'On Error GoTo MSG:
'*************** Shraddha Pandit 25/02/09 **********
frame_addendum.Visible = False
lbl_addendum.Visible = False
txt_addendum.Visible = False
lbl_addendumby.Visible = False
cmb_addendumby.Visible = False

str = "UPDATE RISLIB.RTIMELOG SET ADRREPDATE=?, ADRREPTIME=? WHERE REQNNO=?"
Set UPDTADNDMTIME_QR = rdb.CreateQuery("UPDTADNDMTIME_QR", str)

'**************** Shraddha Pandit 02/03/09 ******************
str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1Q", str)
'***************** 03/03/09 ******************************
str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QR1sprepXXCBC", str)
'************************************************************
chk_Adfld = False
field_clear
TXT_CLEAR
frm_Conventional.Visible = False
Fish_modality.Visible = False

'btnenabl ("INI")
'BTNENABL1 ("INI")
NEWFLAG = False
viewflag = False
MaskEdBox1.Enabled = False

str = "select distinct head1 from dislib.signal_pat order by head1"
Set rsSelCytog = rdb.OpenResultset(str, 1, 1)
While Not rsSelCytog.EOF
    txtAn1Hd1.AddItem (rsSelCytog(0))
    txtAn2Hd1.AddItem (rsSelCytog(0))
    txtAn3Hd1.AddItem (rsSelCytog(0))
    rsSelCytog.MoveNext
Wend

If FAREPORTFLAG = False Then
    FAREPORTFLAG = True
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1FNCY", str)
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
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNYY_STRCY", str)
If Trim(USRDEPT) = "ALL" Then
    REPWTRN_QR(0) = "FZZ"
Else
    REPWTRN_QR(0) = Trim(USRDEPT)
End If
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    'cmb_addendumby.clear
    'cmb_finalby.clear
'    cmb_provby.clear
'    REPWTRNRS.MoveFirst
'    While Not REPWTRNRS.EOF
'        If REPWTRNRS("USRTYPE") >= 8 Then 'mADE 8 FROM 1
'            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
'            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
'        End If
'        If REPWTRNRS("USRTYPE") > 3 And (REPWTRNRS("USRTYPE") <> 6) Then
'            cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
'        End If
'        REPWTRNRS.MoveNext
'    Wend

Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11CY", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ? and dept_code='VZZ'"
Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1FNCY", str)
SELFINALBY_QR(0) = 8    'MADE 8 FROM 5
Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'MsgBox SELFINALBY_RS.RowCount
'cmb_finalby.clear

If Not SELFINALBY_RS.EOF Then
    Do While Not SELFINALBY_RS.EOF
       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
       cmb_addendumby.AddItem Trim(SELFINALBY_RS("USERNAME"))
       SELFINALBY_RS.MoveNext
    Loop
End If

'*********SECURITY**********************************
''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR1161CY", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC%' OR PAT_CATG<>'SF%')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR1161CY", str)

''********ICU***************

str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
Set seldrequisitn_qr11 = rdb.CreateQuery("seldrequisitn11_1qrfnCY", str)


'****** Method : Added by Shraddha on 02/12/09 *****

str = "SELECT * FROM DISLIB.CYTOGMETHO"
Set sel_cytogmethod_qr = rdb.CreateQuery("", str)

Set sel_cytogmethod_rs = sel_cytogmethod_qr.OpenResultset(1, 3)

cmb_methodFI.clear
While Not sel_cytogmethod_rs.EOF
    cmb_methodFI.AddItem Trim(sel_cytogmethod_rs("method"))
    cmb_methodCC.AddItem Trim(sel_cytogmethod_rs("method"))
    sel_cytogmethod_rs.MoveNext
Wend
sel_cytogmethod_rs.Close

'***************************************************

'***** Shraddha 07/12/09 *****
str = "SELECT * FROM DISLIB.CYTOGTEST"
Set sel_cytogtest_qr = rdb.CreateQuery("", str)

Set sel_cytogtest_rs = sel_cytogtest_qr.OpenResultset(1, 3)

cmb_CCTest.clear

While Not sel_cytogtest_rs.EOF
    cmb_CCTest.AddItem Trim(sel_cytogtest_rs("test"))
    sel_cytogtest_rs.MoveNext
Wend
sel_cytogtest_rs.Close

'-----------------------------
str = "SELECT * FROM DISLIB.CYTOGPROBS"
Set sel_CYTOGPROBS_qr = rdb.CreateQuery("", str)

Set sel_CYTOGPROBS_rs = sel_CYTOGPROBS_qr.OpenResultset(1, 3)

cmb_CCProbes.clear

While Not sel_CYTOGPROBS_rs.EOF
    cmb_CCProbes.AddItem Trim(sel_CYTOGPROBS_rs("probes"))
    sel_CYTOGPROBS_rs.MoveNext
Wend
sel_CYTOGPROBS_rs.Close

'*****************************

'*****FNE QRY***********
str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%') ORDER BY PREPORTDT"
Set seldrequisitn_qr10 = rdb.CreateQuery("seldrequisitn11_1qr110CY", str)

str = "SELECT * FROM DISLIB.PROBE where probcode like ?"
Set SELPROBE_QR = rdb.CreateQuery("selprobe_qrCY", str)
SELPROBE_QR(0) = "t%"
Set SELPROBE_RS = SELPROBE_QR.OpenResultset(2, 3)

  str = "select b.modlitydes from rislib.rrequisitn a,rislib.rmodality b  where a.modalitycd=b.modalitycd and a.reqnno=?"
  Set selreqmoda_qr = rdb.CreateQuery("selmodareqcy", str)
    
  str = "Select reqnno from rislib.rrequisitn where (dept_code='UZZ' or dept_code='FZZ' or dept_code='GZZ' or dept_code='VZZ') and caseno=?"
  Set selreqnno_qr = rdb.CreateQuery("select_qrcy", str)
  
  
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
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE DEPT_CODE='VZZ'"
Set sel_deptqr = rdb.CreateQuery("sel1_deptqr11CY", str)
'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)

    'for icd-h retriveing
str = "select * from patoolslib.ma302"
Set selma302_qr = rdb.CreateQuery("selma302_qr1CY", str)
'Set selma302_rs = selma302_qr.OpenResultset(2, 3)

str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301QR = rdb.CreateQuery("SELMA3011QR1CY", str)
'FOR GETTING CLINICAL DIAGNOSIS
str = "select CLINHIST  from rislib.CLINHIST WHERE CASENO=? "
Set SELCLIN_QR = rdb.CreateQuery("SELCLIN_QRCY1", str)

' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqr11CY", str)

str = "select * from rislib.rmodality where dept_code=?"
Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr1CY", str)

'STR = "select * from rislib.rmodality where dept_code=?"
'Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr11", STR)


    'for retriving reqnos
Set req_qr = rdb.CreateQuery("1reqCY", "select * from RISLIB.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT DESC")
Set REQ_RQD = rdb.CreateQuery("REQ_rqdCY", "select reqnno,reqndt from rislib.rrequisitn, dislib.drequisitn where RREQUISITN.REQNNO = DREQUISITN.REQNO AND RREQUISITN.CASENO = ? AND RREQUISITN.MODALITYCD = ? and DREQUISITN.icds2 = 'YES' and RREQUISITN.REQNSTAT = 'M'  ORDER BY REQNDT DESC")

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

Set SELTA101_QR = rdb.CreateQuery("selta101_qr11CY", "select * from PATOOLSLIB.ta101 where case_no =?")



' FOR REFEFING DOCTOR
     'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
     'Commented and added by Shraddha on 03/01/2012
     str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO=?"
     Set REFDR_QR = rdb.CreateQuery("REFDR1_QR1CY", str)
'
'    FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1_QR1CY", str)
   
    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma1041_qr11CY", str)
   
    'for provisional diagnosis

    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("se1lta10121_qr1CY", str)
  
    str = "select * from dislib.dimpresion where reqno=?"
    Set seldimpresion_qr = rdb.CreateQuery("seldimpresion_qr11CY", str)
    
    str = "insert into dislib.cytogreprt(reqno,caseno,modalitycd,fsitudir,fsitupha,noca,opprocedur,probdesc,intrpre,comments,ca,ka,cb,ccnoca,ccintrpre,cccomments,ccpha,ccdir,ccphas,labno,enterby,reportby,result,commitby,fsituphab)values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSREPORT_QR = rdb.CreateQuery("test", str)
    
    
    '******* Shraddha 06/11/09 *******
    str = "insert into dislib.CYTOGENET(reqno,CASENO,REQNDT,REQNTM,REQNSTAT,PATNSTAT,PAT_CATG," & _
    "PREPORTBY,FREPORTBY,SAMCOLDATE,SAMCOLTIME,SAMRECDATE,SAMRECTIME,COMREPDATE,COMREPTIME," & _
    "FSITUDIR,FSITUPHA,FSITUPHAB,NOCA,OPPROCEDUR,Result,INTRPRE,Comments,Ca,Ka,CB,CCNOCA," & _
    "CCINTRPRE,CCCOMMENTS,CCPHA,CCDIR,CCPHAS,LABNO,ENTERBY,REPORTBY,COMMITBY,AT_DIAG,PR_TREAT," & _
    "PS_INDU,RELAPSE,REMISSION,FOLLUP,POSTBMT,B_SEX,SP_TYPE,CGDIAG1,CGDIAG2,FAX,PHONE,MOBILE," & _
    "CALDATE,CALTIME,NCONTROL,NCON_AGE,NCON_SEX,REMARKS1,REMARKS2,PROBDESC,LATERALITS,PREPORTDT,FREPORTDT,REFBYDOC,FLD1)values " & _
    "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    
    Set INS_CYTOGENET_QR = rdb.CreateQuery("", str)
    
    'reqno CASENO REQNDT REQNTM REQNSTAT PATNSTAT PAT_CATG 7
    'PREPORTBY FREPORTBY SAMCOLDATE SAMCOLTIME SAMRECDATE SAMRECTIME COMREPDATE COMREPTIME 8
    'FSITUDIR FSITUPHA FSITUPHAB NOCA OPPROCEDUR Result INTRPRE Comments Ca Ka CB CCNOCA 12
    'CCINTRPRE CCCOMMENTS CCPHA CCDIR CCPHAS LABNO ENTERBY REPORTBY COMMITBY AT_DIAG PR_TREAT 11
    'PS_INDU RELAPSE REMISSION FOLLUP POSTBMT B_SEX SP_TYPE CGDIAG1 CGDIAG2 FAX PHONE MOBILE 12
    'CALDATE CALTIME NCONTROL NCON_AGE NCON_SEX REMARKS1 REMARKS2 7

    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set sel_rreq_qr = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM RISLIB.RTIMELOG WHERE REQNNO=?"
    Set sel_rtime_qr = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM dislib.CYTOGENET WHERE REQNO=?"
    Set sel_cytogenet_qr = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set sel_dreq_qr = rdb.CreateQuery("", str)
    
    str = "update dislib.CYTOGENET set fsitudir =?,fsitupha =?,noca =?,probdesc =?,intrpre =?,comments =?,ca =?,ka =?,cb =?,ccnoca =?,ccintrpre =?,cccomments =?,ccpha =?,ccdir =?,ccphas =?,labno =?,enterby =?,reportby =?,commitby =?,result =?,OPPROCEDUR=?,fsituphab =? where reqno =?"
    Set UPD_CYTOGENETmod_QR = rdb.CreateQuery("", str)
    '*********************************
    
    'STR = "UPDATE DISLIB.CYTOGREPRT SET MODALITYCD=?,FSITUDIR=?,FSITUPHA=?,NOCA=?,PROBDESC=?,INTRPRE=?,COMMENTS=?,CA=?,KA=?,CB=?,CCNOCA=?,CCINTRPRE=?,CCCOMMENTS=?,CCPHA=?,CCDIR=?,CCPHAS=?,LABNO=?,ENTERBY=?,REPORTBY=?,COMMITBY=?,RESULT=?  WHERE REQNO=?"
    str = "update dislib.cytogreprt set modalitycd =?,fsitudir =?,fsitupha =?,noca =?,probdesc =?,intrpre =?,comments =?,ca =?,ka =?,cb =?,ccnoca =?,ccintrpre =?,cccomments =?,ccpha =?,ccdir =?,ccphas =?,labno =?,enterby =?,reportby =?,commitby =?,result =?,OPPROCEDUR=?,fsituphab =? where reqno =?"
    Set UPDREPORT_QR = rdb.CreateQuery("finalt", str)
    
    'STR = "SELECT * FROM DISLIB.CYTOGREPRT WHERE REQNO=?"
    str = "select *  from dislib.cytogreprt where REQNO = ?"
    Set SELREPORT_QR = rdb.CreateQuery("selectport_qr", str)
    SELREPORT_FLAG = False
       
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR11CY", str)
                                        
    '******************** Shraddha Pandit 16/01/09 ****************************
    str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
    Set UPDATERTIMELOG = rdb.CreateQuery("", str)
    '**************************************************************************
    str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
    Set UPDATERTIMELOG = rdb.CreateQuery("UPDATERTIMELOG11", str)
                                        
    str = "update RISLIB.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
    Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qrCY", str)

    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?,FREPORTDT=?,REQNSTAT=?,COMMITFLAG=? WHERE REQNNO=?"
    Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qr1CY", str)

'    STR = "INSERT INTO DISLIB.ADDENDUMFN(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
'    Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1CY", STR)
    str = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
    Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1", str)

    str = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11CY", str)
    SELDADDENDUM_QR(0) = "FSS"
    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)

    str = "select * from DISLIB.drequisitn where caseno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_1qr1CY", str)

'   Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)

'    txt_addendum.Enabled = False
'    cmb_addendumby.Enabled = False
    cmb_dept.Enabled = False
    FLDENABL ("DIS")

   str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
   Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQR1CY", str)
'   SELRREQNOQR(0) = "FZZSP0200000"
'   Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)

    

    str = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?,FROZENSEC=?,NMETERIALS=? WHERE REQNO=?"
    Set upddrequisitn_qr = rdb.CreateQuery("UPDSPREPORT1_QR1CY", str)

    str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO11_QR1CY", str)

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
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQR1CY", str)

   
  str = "INSERT INTO RISLIB.DISIMP(REQNO,IMP)VALUES(?,?)"
  Set INSDISIMP_QR = rdb.CreateQuery("INSDISIMP_QRCY", str)

  str = "UPDATE RISLIB.DISIMP SET IMP=? WHERE REQNO=?"
  Set UPDDISIMP_QR = rdb.CreateQuery("UPDDISIMP_QRCY", str)
  
  str = "Insert into dislib.CYTOG_ETO values(?,?,?,?,?,?,?,?,?,?,?)"
  Set InsCytog_ETO_qr = rdb.CreateQuery("InsCytogETO", str)
  
  str = "Insert into dislib.CYTOG_chic values(?,?,?,?,?,?,?,?,?,?,?)"
  Set InsCytog_CHIC_qr = rdb.CreateQuery("InsCytogchic", str)
  
  str = "Insert into dislib.CYTOG_sp values(?,?,?,?,?,?,?,?,?,?,?)"
  Set InsCytog_SP_qr = rdb.CreateQuery("InsCytogsp", str)
  
  
  
  'STR = "select * from dislib.cytogcldta where reqno=?"
  str = "SELECT * FROM DISLIB.CYTOGCLDTA WHERE REQNO=?"
  Set CYTOGCLDTA_QR = rdb.CreateQuery("CYTOGCLDTA_QRRcy", str)
  CYTOGCLDTA_FLG = False
  
  str = "select * from dislib.signal_pat where head1=?"
  Set selHead = rdb.CreateQuery("qrselHead", str)

 End If
 
BTNENABL1 ("INI")
Exit Sub
MSG:
MsgBox Err.Description

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
Private Sub cmb_reqno_Click()
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
        If SELTA101FLAG = False Then
            SELTA101_QR(0) = MaskEdBox1.FormattedText
             Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            SELTA101FLAG = True
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
            REFDR_QR(1) = VAR_REQNNO
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = VAR_REQNNO
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
        If REFDR_RS.RowCount > 0 Then
            If REFDR_RS("REFDOC") <> "" Then
                txt_refdoc.Text = REFDR_RS("REFDOC") ' REFDR_RS("REFBYDOC")
            Else
                txt_refdoc.Text = ""
            End If
            
            
            'SHWETA 13-FEB-2012
            Dim rq_DATE As rdoQuery
            Dim RS_DATE As rdoResultset
            Set rq_DATE = rdb.CreateQuery("", "select REQNDT from RISLIB.RREQUISITN where REQNNO=? ")
            rq_DATE(0) = Trim(Replace(cmb_reqno.Text, "/", ""))
            Set RS_DATE = rq_DATE.OpenResultset(1, 3)
         '   MsgBox RS_DATE(0)
            
            If RS_DATE(0) <> "" Or RS_DATE(0) <> Null Then
                txt_rec_on.Text = RS_DATE(0)
            Else
                txt_rec_on.Text = ""
            End If
            RS_DATE.Close
            Set RS_DATE = Nothing
'            If REFDR_RS("recordno") = "1" Then
'            LBL_VPP.Caption = "VPP"
'                If Trim(REFDR_RS("recptno")) = "VPP" Then
'                    MsgBox "Payment is not done,You are not authorised to generate report."
'                    'btnenabl ("INI")
'                    BTNENABL1 ("INI")
'                    field_clear
'                    Exit Sub
'                End If
'            End If
'            ALLOWING CHANGE OF PROBE & TEST FOR T108, T229 & T230

'         If REFDR_RS("SER_CODE") = "T108" Or REFDR_RS("SER_CODE") = "T229" Or REFDR_RS("SER_CODE") = "T230" Then
'          txt_test.Locked = False
'          cmb_probe.Locked = False
'
'         Else
'           txt_test.Locked = True
'           cmb_probe.Locked = True
'
'         End If
         
        End If

        If Left(Trim(MaskEdBox1.FormattedText), 1) = "R" And Left(Trim(MaskEdBox1.FormattedText), 2) <> "RA" Then
            Dim rs_refby As rdoResultset
            Set rs_refby = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.ta1011 WHERE CASE_NO='" & Trim(MaskEdBox1.FormattedText) & "' ", 1, 1)
                If Not rs_refby.EOF Then
                txt_refdoc.Text = Trim(rs_refby("REF_BY_DOC"))
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
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_cytology.Text = selpathdrequisitn_rs("PATHNO")
            'OR CONDITION ADDED BY SHWETA 13-FEB-2012 BCOZ OK WAS NOT SHOWING IN TEXTBOX
            If Trim(selpathdrequisitn_rs("nmeterials").Value) <> Null Or Trim(selpathdrequisitn_rs("nmeterials").Value) <> "" Then
                txt_specimen.Text = selpathdrequisitn_rs("nmeterials").Value
            Else
                txt_specimen.Text = ""
            End If
            
'            If selpathdrequisitn_rs.EOF = False Then
'                cmb_cellblock.Text = " " & selpathdrequisitn_rs("frozensec")
'            Else
'                cmb_cellblock.Text = " "
'            End If
'            txt_noslideswet.Text = selpathdrequisitn_rs("outblock")

            Dim code As String
            txt_test.Text = selpathdrequisitn_rs("OPPROCEDUR")
            
            sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
            Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
            
            code = Right(sel_rreq_rs!ser_code, 3) & "%"
            SELPROBE_QR(0) = Trim(code)
            SELPROBE_RS.MoreResults
            SELPROBE_RS.Requery
            cmb_probe.clear
            While Not SELPROBE_RS.EOF
               cmb_probe.AddItem Trim(SELPROBE_RS("PROBDESC"))
               SELPROBE_RS.MoveNext
               cmb_probe.Text = cmb_probe.List(0)
            Wend
            
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
        If SELTA101FLAG = False Then
            SELTA101_QR(0) = MaskEdBox1.FormattedText
            Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            SELTA101FLAG = True
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
        'If viewflag = True Then
        If REFDR_RS.RowCount > 0 Then
            'txt_refdoc.Text = REFDR_RS("REFBYDOC")
            txt_refdoc.Text = REFDR_RS("REFDOC")
            
            sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
            Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
            If sel_rreq_rs("APNTDT") <> "" Then
                txt_rec_on.Text = sel_rreq_rs("APNTDT")
            Else
                txt_rec_on.Text = ""
            End If
            If Trim(sel_rreq_rs("ENTERBY")) <> "" Then
                cmb_enteryby.Text = Trim(sel_rreq_rs("ENTERBY"))
            Else
                cmb_enteryby.Text = ""
            End If
            
            If Trim(sel_rreq_rs("PREPORTBY")) <> "" Then
                cmb_finalby.Text = Trim(sel_rreq_rs("PREPORTBY"))
            Else
                cmb_finalby.Text = ""
            End If
            
            '******* View CYTOG_ETO added by v.s.k. on 03.04.2012
            str = "select * from dislib.cytog_eto where reqno=?"
            Set selCytog = rdb.CreateQuery("", str)
            selCytog(0) = req_chglen(cmb_reqno.Text)
            Set rsSelCytog = selCytog.OpenResultset(1, 3)
            
            I = 0
            If Not rsSelCytog.EOF Then
                txtAn1Hd1.Text = rsSelCytog(2)
                txtAn1Hd2.Text = rsSelCytog(3)
                
                While Not rsSelCytog.EOF
                    msfAn1.TextMatrix(I, 0) = rsSelCytog(4)
                    msfAn1.TextMatrix(I, 1) = rsSelCytog(5)
                    msfAn1.TextMatrix(I, 2) = rsSelCytog(6)
                    msfAn1.TextMatrix(I, 3) = rsSelCytog(7)
                    msfAn1.TextMatrix(I, 4) = rsSelCytog(8)
                    msfAn1.TextMatrix(I, 5) = rsSelCytog(9)
                    msfAn1.TextMatrix(I, 6) = rsSelCytog(10)
                    rsSelCytog.MoveNext
                    I = I + 1
                Wend
            End If
            
            str = "select * from dislib.cytog_chic where reqno=?"
            Set selCytog = rdb.CreateQuery("", str)
            selCytog(0) = req_chglen(cmb_reqno.Text)
            Set rsSelCytog = selCytog.OpenResultset(1, 3)
            
            I = 0
            If Not rsSelCytog.EOF Then
                txtAn2Hd1.Text = rsSelCytog(2)
                txtAn2Hd2.Text = rsSelCytog(3)
                
                While Not rsSelCytog.EOF
                    msfAn2.TextMatrix(I, 0) = rsSelCytog(4)
                    msfAn2.TextMatrix(I, 1) = rsSelCytog(5)
                    msfAn2.TextMatrix(I, 2) = rsSelCytog(6)
                    msfAn2.TextMatrix(I, 3) = rsSelCytog(7)
                    msfAn2.TextMatrix(I, 4) = rsSelCytog(8)
                    msfAn2.TextMatrix(I, 5) = rsSelCytog(9)
                    msfAn2.TextMatrix(I, 6) = rsSelCytog(10)
                    rsSelCytog.MoveNext
                    I = I + 1
                Wend
            End If
            
            str = "select * from dislib.cytog_sp where reqno=?"
            Set selCytog = rdb.CreateQuery("", str)
            selCytog(0) = req_chglen(cmb_reqno.Text)
            Set rsSelCytog = selCytog.OpenResultset(1, 3)
            
            I = 0
            If Not rsSelCytog.EOF Then
                txtAn3Hd1.Text = rsSelCytog(2)
                txtAn3Hd2.Text = rsSelCytog(3)
                
                While Not rsSelCytog.EOF
                    msfAn3.TextMatrix(I, 0) = rsSelCytog(4)
                    msfAn3.TextMatrix(I, 1) = rsSelCytog(5)
                    msfAn3.TextMatrix(I, 2) = rsSelCytog(6)
                    msfAn3.TextMatrix(I, 3) = rsSelCytog(7)
                    msfAn3.TextMatrix(I, 4) = rsSelCytog(8)
                    msfAn3.TextMatrix(I, 5) = rsSelCytog(9)
                    msfAn3.TextMatrix(I, 6) = rsSelCytog(10)
                    rsSelCytog.MoveNext
                    I = I + 1
                Wend
            End If
            selCytog.Close
            
            '********
            
            If sel_rreq_rs("COMMITFLAG") = "Y" Then
                'MsgBox "Final Report is Generated, You can't Modify the Final Report,You can Add ADDENDUMS Only"
                'cmd_modify.Enabled = True
                MsgBox "Final Report is Generated, You can't Modify the Final Report ,You can Add ADDENDUMS Only"
                cmd_modify.Enabled = True
                'cmd_modify.Enabled = False
                cmd_commit.Enabled = False
                
                  sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
                Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
                
               If IsNull(Trim(sel_rreq_rs("FREPORTBY"))) Then
                   cmb_finalby.Text = ""
                Else
                   cmb_finalby.Text = Trim(sel_rreq_rs("FREPORTBY"))
               End If
                '************** Shraddha Pandit 25/02/09 ***************
                ADDENDUMFALG = True
                frame_addendum.Visible = True
                lbl_addendum.Visible = True
                txt_addendum.Visible = True
                txt_addendum.Enabled = False
                lbl_addendumby.Visible = True
                cmb_addendumby.Visible = True
                cmb_addendumby.Enabled = False
                
                txt_no_of_cell.Enabled = False
                txt_fresult.Enabled = False
                FIsh_txt_inter.Enabled = False
                fish_txt_comments.Enabled = False
                txt_test.Enabled = False
                cmb_probe.Enabled = False
                
                txt_no_of_mata.Enabled = False
                txt_result.Enabled = False
                txt_Interpre.Enabled = False
                txt_comments.Enabled = False
                '*******************************************************
            End If
          ' ALLOWING CHANGE OF PROBE & TEST FOR T108, T229 & T230
            
         If sel_rreq_rs("SER_CODE") = "T108" Or sel_rreq_rs("SER_CODE") = "T229" Or sel_rreq_rs("SER_CODE") = "T230" Then
          txt_test.Locked = False
          cmb_probe.Locked = False
          
         Else
           txt_test.Locked = True
           cmb_probe.Locked = True
          
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
        
        SELREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
        If SELREPORT_FLAG = False Then
           Set SELREPORT_RS = SELREPORT_QR.OpenResultset(2, 3)
           SELREPORT_FLAG = True
        Else
          SELREPORT_RS.MoreResults
          SELREPORT_RS.Requery
        End If
        If SELREPORT_RS.RowCount > 0 Then
                If Trim(SELREPORT_RS!modalitycd) = "CC" Then
                   frm_Conventional.Visible = True
                   Fish_modality.Visible = False
                End If
                If Trim(SELREPORT_RS!modalitycd) = "FI" Then
                    frm_Conventional.Visible = False
                    Fish_modality.Visible = True
                End If
                
                cmb_modality.Text = Trim(SELREPORT_RS!modalitycd)
                If Trim(SELREPORT_RS!FSITUDIR) = "" Then
                  fish_method1.Visible = True
                  fish_method1.Value = 0
                Else
                  fish_method1.Visible = True
                  fish_method1.Value = 1
                End If
                
                '***** Shraddha 02/12/09 *****
                If Trim(SELREPORT_RS!FSITUDIR) = "" Then
                  fish_method1.Visible = True
                  fish_method1.Value = 0
                  cmb_methodFI.Text = ""
                Else
                  fish_method1.Visible = True
                  fish_method1.Value = 1
                  cmb_methodFI.Text = Trim(SELREPORT_RS!FSITUDIR)
                End If
                'cmb_methodFI
                'In either modality methods should be available to select
                'So lets not disturb those combos (cmb_methodCC,cmb_methodFI)
                'Under view
                '*****************************
                
                '***** Shraddha 07/12/09 *****
                If Trim(SELREPORT_RS!OPPROCEDUR) <> "" Then
                    cmb_CCTest.Text = Trim(SELREPORT_RS!OPPROCEDUR)
                End If
                If Trim(SELREPORT_RS!PROBDESC) <> "" Then
                    cmb_CCProbes.Text = Trim(SELREPORT_RS!PROBDESC)
                End If
                '*****************************
                             
                If Trim(SELREPORT_RS!fsitupha) = "" Then
                 Fish_method2.Visible = True
                 Fish_method2.Value = 0
                Else
                 Fish_method2.Visible = True
                 Fish_method2.Value = 1
                End If
                
                If Trim(SELREPORT_RS!FSITUPHAB) = "" Then
                 fish_method3.Visible = True
                 fish_method3.Value = 0
                Else
                 fish_method3.Visible = True
                 fish_method3.Value = 1
                End If
                
                txt_no_of_cell.Text = Trim(SELREPORT_RS!NOCA)
                txt_test.Text = Trim(SELREPORT_RS!OPPROCEDUR)
                cmb_probe.Text = Trim(SELREPORT_RS!PROBDESC)
                FIsh_txt_inter.Text = Trim(SELREPORT_RS!INTRPRE)
                fish_txt_comments.Text = Trim(SELREPORT_RS!Comments)
                txt_no_of_mata.Text = Trim(SELREPORT_RS!CCNOCA)
                txt_Interpre.Text = Trim(SELREPORT_RS!CCINTRPRE)
                txt_comments.Text = Trim(SELREPORT_RS!CCCOMMENTS)
                If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
                 txt_result.Text = Trim(SELREPORT_RS!Result)
                Else
                 txt_fresult.Text = Trim(SELREPORT_RS!Result)
                End If
                
                If Trim(SELREPORT_RS!CCPHA) = "" Then
                  Chk_method1.Visible = True
                  Chk_method1.Value = 0
                Else
                  Chk_method1.Value = 1
                  Chk_method1.Visible = True
                End If
                
                '***** Shraddha 02/12/09 *****
                
                If Trim(SELREPORT_RS!CCPHA) = "" Then
                  Chk_method1.Visible = True
                  Chk_method1.Value = 0
                  cmb_methodCC.Text = ""
                Else
                  Chk_method1.Value = 1
                  Chk_method1.Visible = True
                  cmb_methodCC.Text = Trim(SELREPORT_RS!CCPHA)
                End If
                
                'In either modality methods should be available to select
                'So lets not disturb those combos (cmb_methodCC,cmb_methodFI)
                'Under view
                '*****************************
                
                
                
                If Trim(SELREPORT_RS!CCDIR) = "" Then
                  chk_method2.Visible = True
                  chk_method2.Value = 0
                Else
                  chk_method2.Value = 1
                  chk_method2.Visible = True
                End If
                
                If Trim(SELREPORT_RS!CCPHAS) = "" Then
                  chk_method3.Visible = True
                  chk_method3.Value = 0
                Else
                  chk_method3.Visible = True
                  chk_method3.Value = 1
                End If
                txt_cytology.Text = Trim(SELREPORT_RS!LABNO)
                cmb_enteryby.Text = Trim(SELREPORT_RS!ENTERBY)
                cmb_finalby.Text = Trim(SELREPORT_RS!REPORTBY)
            End If

'if
      
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        
        'following code commented on 25 june 2012 by v.s.k.
'        If selpathdrequisitn_rs.RowCount > 0 Then
'            txt_cytology.Text = Trim(selpathdrequisitn_rs("PATHNO"))
''            If Trim(SELREPORT_RS!MODALITYCD) = "FI" Then
'            If Trim(selpathdrequisitn_rs("NMETERIALS")) <> "" Or Trim(selpathdrequisitn_rs("nmeterials").Value) <> "" Then
'               txt_specimen.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
'            Else
'                txt_specimen.Text = ""
'            End If
''            Else
''               txt_Interpre.Text = Trim(selpathdrequisitn_rs("nmeterials"))
''            End If
'
'            txt_test.Text = selpathdrequisitn_rs("OPPROCEDUR")
'
'            sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
'            Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
'
'            code = Right(sel_rreq_rs!ser_code, 3) & "%"
'            SELPROBE_QR(0) = Trim(code)
'            SELPROBE_RS.MoreResults
'            SELPROBE_RS.Requery
'            cmb_probe.clear
'            While Not SELPROBE_RS.EOF
'               cmb_probe.AddItem Trim(SELPROBE_RS("PROBDESC"))
'               SELPROBE_RS.MoveNext
'               cmb_probe.Text = cmb_probe.List(0)
'            Wend
'
'
'        End If

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
'        SELDADDENDUM_QR(0) = reqno
'            SELDADDENDUM_RS.MoreResults
'            SELDADDENDUM_RS.Requery
'            If SELDADDENDUM_RS.RowCount > 0 Then
'                While Not SELDADDENDUM_RS.EOF
'                    ADDENDUM = ADDENDUM + Chr(13) & "ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
'                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
'                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
'                SELDADDENDUM_RS.MoveNext
'                Wend
'                txt_addendum.Text = Trim(ADDENDUM)
'                ADDENDUMFALG = True
'            End If
'********************** Shraddha Pandit 25/02/09 *************
            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            SELDADDENDUM_RS.MoreResults
            SELDADDENDUM_RS.Requery
            If SELDADDENDUM_RS.RowCount > 0 Then
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
'**************************************************************
        'End If
        cmb_reqno.Enabled = False
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
    End If

End If 'End if patient
Dim STATUS As String
CYTOGCLDTA_QR(0) = req_chglen(cmb_reqno.Text)
If CYTOGCLDTA_FLG = False Then
 Set CYTOGCLDTA_RS = CYTOGCLDTA_QR.OpenResultset(2, 3)
 CYTOGCLDTA_FLG = True
Else
 CYTOGCLDTA_RS.MoreResults
 CYTOGCLDTA_RS.Requery
End If
If CYTOGCLDTA_RS.RowCount > 0 Then
  If CYTOGCLDTA_RS!AT_DIaG <> 0 Then
   STATUS = "At Diagnosis"
  End If
  If CYTOGCLDTA_RS!pr_treat <> 0 Then
    STATUS = STATUS & "  " & "Partially Treated"
  End If
  If CYTOGCLDTA_RS!ps_indu <> 0 Then
  STATUS = STATUS & "  " & "Post Induction"
  End If
  If CYTOGCLDTA_RS!remission <> 0 Then
  STATUS = STATUS & "  " & "Remission"
  End If
  If CYTOGCLDTA_RS!relapse <> 0 Then
  STATUS = STATUS & "  " & "Relapse"
  End If
  If CYTOGCLDTA_RS!postbmt <> 0 Then
   STATUS = STATUS & "  " & "Post BMT"
  End If
   If Trim(CYTOGCLDTA_RS!B_SEX) = "M" Then
      STATUS = STATUS & "  " & "Male"
   ElseIf Trim(CYTOGCLDTA_RS!B_SEX) = "F" Then
      STATUS = STATUS & "  " & "Female"
   End If
   txt_clinicaldetail.Text = Trim(STATUS) & "  " & Trim(CYTOGCLDTA_RS!FOLLUP)
   txt_specimentype.Text = Trim(CYTOGCLDTA_RS!SP_TYPE)
End If
If opt_staff.Value = True Then
    STAFF
End If
'SHWETA 30-SEPT-2010
If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
frm_Conventional.Visible = True
Fish_modality.Visible = False
End If
If Trim(Right(cmb_modality.Text, 2)) = "FI" Then
frm_Conventional.Visible = False
Fish_modality.Visible = True
End If

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub msfAn1_EnterCell()
If msfAn1.col = 0 Then
    txtData.Top = msfAn1.Top + msfAn1.CellTop
    txtData.Left = msfAn1.Left + msfAn1.CellLeft
    txtData.Text = msfAn1.TextMatrix(msfAn1.row, msfAn1.col)
    txtData.Visible = True
    txtData.SetFocus
    txtData.Width = msfAn1.CellWidth
ElseIf msfAn1.col = 1 Then
    txtData.Top = msfAn1.Top + msfAn1.CellTop
    txtData.Left = msfAn1.Left + msfAn1.CellLeft
    txtData.Text = msfAn1.TextMatrix(msfAn1.row, msfAn1.col)
    txtData.Visible = True
    txtData.SetFocus
    txtData.Width = msfAn1.CellWidth
ElseIf msfAn1.col = 2 Then
    txtData.Top = msfAn1.Top + msfAn1.CellTop
    txtData.Left = msfAn1.Left + msfAn1.CellLeft
    txtData.Text = msfAn1.TextMatrix(msfAn1.row, msfAn1.col)
    txtData.Visible = True
    txtData.SetFocus
    txtData.Width = msfAn1.CellWidth
ElseIf msfAn1.col = 3 Then
'    If msfAn1.row = 0 Then
'
'    Else
        txtData.Top = msfAn1.Top + msfAn1.CellTop
        txtData.Left = msfAn1.Left + msfAn1.CellLeft
        txtData.Visible = True
        txtData.SetFocus
        txtData.Text = msfAn1.TextMatrix(msfAn1.row, msfAn1.col)
        txtData.Width = msfAn1.CellWidth
'    End If
    
ElseIf msfAn1.col = 4 Then
'    If msfAn1.row = 0 Then
'    Else
        txtData.Top = msfAn1.Top + msfAn1.CellTop
        txtData.Left = msfAn1.Left + msfAn1.CellLeft
        txtData.Text = msfAn1.TextMatrix(msfAn1.row, msfAn1.col)
        txtData.Visible = True
        txtData.SetFocus
        txtData.Width = msfAn1.CellWidth
'    End If
ElseIf msfAn1.col = 5 Then
    txtData.Top = msfAn1.Top + msfAn1.CellTop
    txtData.Left = msfAn1.Left + msfAn1.CellLeft
    txtData.Text = msfAn1.TextMatrix(msfAn1.row, msfAn1.col)
    txtData.Visible = True
    txtData.SetFocus
    txtData.Width = msfAn1.CellWidth
ElseIf msfAn1.col = 6 Then
    txtData.Top = msfAn1.Top + msfAn1.CellTop
    txtData.Left = msfAn1.Left + msfAn1.CellLeft
    txtData.Text = msfAn1.TextMatrix(msfAn1.row, msfAn1.col)
    txtData.Visible = True
    txtData.SetFocus
    txtData.Width = msfAn1.CellWidth
End If
End Sub


Private Sub msfAn2_EnterCell()
If msfAn2.col = 0 Then
    txtData2.Visible = True
    txtData2.Top = msfAn2.Top + msfAn2.CellTop
    txtData2.Left = msfAn2.Left + msfAn2.CellLeft
    txtData2.Text = msfAn2.TextMatrix(msfAn2.row, msfAn2.col)
    txtData2.SetFocus
    txtData2.Width = msfAn2.CellWidth
ElseIf msfAn2.col = 1 Then
    txtData2.Visible = True
    txtData2.Top = msfAn2.Top + msfAn2.CellTop
    txtData2.Left = msfAn2.Left + msfAn2.CellLeft
    txtData2.Text = msfAn2.TextMatrix(msfAn2.row, msfAn2.col)
    txtData2.SetFocus
    txtData2.Width = msfAn3.CellWidth
ElseIf msfAn2.col = 2 Then
    txtData2.Visible = True
    txtData2.Top = msfAn2.Top + msfAn2.CellTop
    txtData2.Left = msfAn2.Left + msfAn2.CellLeft
    txtData2.Text = msfAn2.TextMatrix(msfAn2.row, msfAn2.col)
    txtData2.SetFocus
    txtData2.Width = msfAn2.CellWidth
ElseIf msfAn2.col = 3 Then
'    If msfAn2.row = 0 Then
'
'    Else
    txtData2.Visible = True
    txtData2.Top = msfAn2.Top + msfAn2.CellTop
        txtData2.Left = msfAn2.Left + msfAn2.CellLeft
        txtData2.SetFocus
        txtData2.Text = msfAn2.TextMatrix(msfAn2.row, msfAn2.col)
        txtData2.Width = msfAn2.CellWidth
'    End If
    
ElseIf msfAn2.col = 4 Then
'    If msfAn2.row = 0 Then
'    Else
    txtData2.Visible = True
    txtData2.Top = msfAn2.Top + msfAn2.CellTop
        txtData2.Left = msfAn2.Left + msfAn2.CellLeft
        txtData2.Text = msfAn2.TextMatrix(msfAn2.row, msfAn2.col)
        txtData2.SetFocus
        txtData2.Width = msfAn2.CellWidth
'    End If
ElseIf msfAn2.col = 5 Then
    txtData2.Visible = True
    txtData2.Top = msfAn2.Top + msfAn2.CellTop
    txtData2.Left = msfAn2.Left + msfAn2.CellLeft
    txtData2.Text = msfAn2.TextMatrix(msfAn2.row, msfAn2.col)
    txtData2.SetFocus
    txtData2.Width = msfAn2.CellWidth
ElseIf msfAn2.col = 6 Then
    txtData2.Visible = True
    txtData2.Top = msfAn2.Top + msfAn2.CellTop
    txtData2.Left = msfAn2.Left + msfAn2.CellLeft
    txtData2.Text = msfAn2.TextMatrix(msfAn2.row, msfAn2.col)
    txtData2.SetFocus
    txtData2.Width = msfAn2.CellWidth
End If

End Sub

Private Sub msfAn3_EnterCell()
If msfAn3.col = 0 Then
    txtData3.Top = msfAn3.Top + msfAn3.CellTop
    txtData3.Left = msfAn3.Left + msfAn3.CellLeft
    txtData3.Text = msfAn3.TextMatrix(msfAn3.row, msfAn3.col)
    txtData3.Visible = True
    txtData3.SetFocus
    txtData3.Width = msfAn3.CellWidth
ElseIf msfAn3.col = 1 Then
    txtData3.Top = msfAn3.Top + msfAn3.CellTop
    txtData3.Left = msfAn3.Left + msfAn3.CellLeft
    txtData3.Text = msfAn3.TextMatrix(msfAn3.row, msfAn3.col)
    txtData3.Visible = True
    txtData3.SetFocus
    txtData3.Width = msfAn3.CellWidth
ElseIf msfAn3.col = 2 Then
    txtData3.Top = msfAn3.Top + msfAn3.CellTop
    txtData3.Left = msfAn3.Left + msfAn3.CellLeft
    txtData3.Text = msfAn3.TextMatrix(msfAn3.row, msfAn3.col)
    txtData3.Visible = True
    txtData3.SetFocus
    txtData3.Width = msfAn3.CellWidth
ElseIf msfAn3.col = 3 Then
'    If msfAn3.row = 0 Then
'
'    Else
        txtData3.Top = msfAn3.Top + msfAn3.CellTop
        txtData3.Left = msfAn3.Left + msfAn3.CellLeft
        txtData3.Visible = True
        txtData3.SetFocus
        txtData3.Text = msfAn3.TextMatrix(msfAn3.row, msfAn3.col)
        txtData3.Width = msfAn3.CellWidth
'    End If
    
ElseIf msfAn3.col = 4 Then
'    If msfAn3.row = 0 Then
'    Else
        txtData3.Top = msfAn3.Top + msfAn3.CellTop
        txtData3.Left = msfAn3.Left + msfAn3.CellLeft
        txtData3.Text = msfAn3.TextMatrix(msfAn3.row, msfAn3.col)
        txtData3.Visible = True
        txtData3.SetFocus
        txtData3.Width = msfAn3.CellWidth
'    End If
ElseIf msfAn3.col = 5 Then
    txtData3.Top = msfAn3.Top + msfAn3.CellTop
    txtData3.Left = msfAn3.Left + msfAn3.CellLeft
    txtData3.Text = msfAn3.TextMatrix(msfAn3.row, msfAn3.col)
    txtData3.Visible = True
    txtData3.SetFocus
    txtData3.Width = msfAn3.CellWidth
ElseIf msfAn3.col = 6 Then
    txtData3.Top = msfAn3.Top + msfAn3.CellTop
    txtData3.Left = msfAn3.Left + msfAn3.CellLeft
    txtData3.Text = msfAn3.TextMatrix(msfAn3.row, msfAn3.col)
    txtData3.Visible = True
    txtData3.SetFocus
    txtData3.Width = msfAn3.CellWidth

End If

End Sub

Private Sub Opt_patient_Click()
    If Opt_patient.Value = True Then
       lbl_caseno.Caption = "Case No."
        ''MaskEdBox1.Mask = "??>/#####" ''commented by Monika on 02/12/2022
    End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    'MaskEdBox1.Mask = "#######"            ' COMMENTED BY VARINDER
End If
End Sub

'Public Sub FNEQuery()
'      'STR = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%') ORDER BY PREPORTDT"
'
'
'If seldrequisitn_flag10 = False Then
'    seldrequisitn_qr10(0) = UCase(MaskEdBox1.FormattedText)
'    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
'    seldrequisitn_FLAG = True
'Else
'    seldrequisitn_qr10(0) = UCase(MaskEdBox1.FormattedText)
'    seldrequisitn_rs10.MoreResults
'    seldrequisitn_rs10.Requery
'End If
'If seldrequisitn_rs10.RowCount > 0 Then
'    cmb_prereqno.clear
'    While Not seldrequisitn_rs10.EOF
'        'CMB_PREINVEST.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & req_chglen(seldrequisitn_rs10("REQNNO"))
'        cmb_prereqno.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
'                                '" $ " & Trim(MODDESC) &
'        seldrequisitn_rs10.MoveNext
'    Wend
'    cmb_prereqno.Text = cmb_prereqno.List(0)
'End If
'
'End Sub
Public Sub assign()
'On Error GoTo MSG:
Dim I As Integer
Dim MICROSCOPIC As String
Dim flg_eto As Boolean
Dim flg_chic As Boolean
Dim flg_sp As Boolean
Dim srno As Integer
rdb.BeginTrans
'Next statement commented and immediate next is added by Shraddha on 26/02/09
'If flagsave = 1 Then
If flagsave = 1 And ADDENDUMFALG = False Then
    flagsave = 0
    
    '************** Shraddha 06/11/09 ************
    'reqno CASENO REQNDT REQNTM REQNSTAT PATNSTAT PAT_CATG 7
    'PREPORTBY FREPORTBY SAMCOLDATE SAMCOLTIME SAMRECDATE SAMRECTIME COMREPDATE COMREPTIME 8
    'FSITUDIR FSITUPHA FSITUPHAB NOCA OPPROCEDUR Result INTRPRE Comments Ca Ka CB CCNOCA 12
    'CCINTRPRE CCCOMMENTS CCPHA CCDIR CCPHAS LABNO ENTERBY REPORTBY COMMITBY AT_DIAG PR_TREAT 11
    'PS_INDU RELAPSE REMISSION FOLLUP POSTBMT B_SEX SP_TYPE CGDIAG1 CGDIAG2 FAX PHONE MOBILE 12
    'CALDATE CALTIME NCONTROL NCON_AGE NCON_SEX REMARKS1  REMARKS2 7
    
    sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
    Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
    
    INS_CYTOGENET_QR(0) = Trim(req_chglen(cmb_reqno.Text)) 'reqno
    INS_CYTOGENET_QR(1) = Trim(MaskEdBox1.FormattedText) 'CASENO
    INS_CYTOGENET_QR(2) = sel_rreq_rs("REQNDT") 'REQNDT
    INS_CYTOGENET_QR(3) = sel_rreq_rs("REQNTM") 'REQNTM
    'INS_CYTOGENET_QR(4) = sel_rreq_rs("REQNSTAT") 'REQNSTAT
    INS_CYTOGENET_QR(5) = sel_rreq_rs("PATNSTAT") 'PATNSTAT
    INS_CYTOGENET_QR(6) = sel_rreq_rs("PAT_CATG") 'PAT_CATG
    INS_CYTOGENET_QR(61) = Trim(txt_refdoc.Text) 'REFBYDOC
    
 
    
    '*********************************************
    'reqno,caseno,modalitycd,fsitudir,fsitupha,noca,opprocedur,probdesc,intrpre,
    'comments,ca,ka,cb,ccnoca,ccintrpre,cccomments,ccpha,ccdir,ccphas,labno,
    'enterby,reportby,result,commitby,fsituphab
    
    INSREPORT_QR(0) = Trim(req_chglen(cmb_reqno.Text)) 'reqno
    INSREPORT_QR(1) = Trim(MaskEdBox1.FormattedText) 'caseno
    INSREPORT_QR(2) = Trim(Right(cmb_modality.Text, 2)) 'modalitycd
   
   'FSITUDIR FSITUPHA FSITUPHAB NOCA OPPROCEDUR Result INTRPRE Comments Ca Ka CB CCNOCA 12
'   If fish_method1.Value = 1 Then
'        INSREPORT_QR(3) = Trim(fish_method1.Caption) 'fsitudir
'        INS_CYTOGENET_QR(15) = Trim(fish_method1.Caption) 'FSITUDIR
'   Else
'        INSREPORT_QR(3) = "" 'fsitudir
'        INS_CYTOGENET_QR(15) = "" 'FSITUDIR
'   End If
   
   ''Change needed 02/12/09
   ''Commented above code Shraddha
    If Trim(cmb_methodFI.Text) <> "" Then
        INSREPORT_QR(3) = Trim(cmb_methodFI.Text) 'fsitudir
        INS_CYTOGENET_QR(15) = Trim(cmb_methodFI.Text) 'FSITUDIR
    Else
        INSREPORT_QR(3) = "" 'fsitudir
        INS_CYTOGENET_QR(15) = "" 'FSITUDIR
    End If
   ''****************************
   
   
    If Fish_method2.Value = 1 Then
        INSREPORT_QR(4) = Trim(Fish_method2.Caption) 'fsitupha
        INS_CYTOGENET_QR(16) = Trim(Fish_method2.Caption) 'FSITUPHA
    Else
        INSREPORT_QR(4) = ""
        INS_CYTOGENET_QR(16) = ""
    End If
   
   
    If Trim(txt_no_of_cell.Text) <> "" Then
        INSREPORT_QR(5) = Trim(txt_no_of_cell.Text) 'noca
    Else
        INSREPORT_QR(5) = 0
    End If
   
    If Trim(txt_no_of_cell.Text) <> "" Then
        INS_CYTOGENET_QR(18) = txt_no_of_cell.Text 'NOCA
    Else
        INS_CYTOGENET_QR(18) = 0
    End If
    
    If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
        INSREPORT_QR(6) = Trim(cmb_CCTest.Text) 'opprocedur
        INS_CYTOGENET_QR(19) = Trim(cmb_CCTest.Text) 'OPPROCEDUR
    ElseIf Trim(Right(cmb_modality.Text, 2)) = "FI" Then
        INSREPORT_QR(6) = Trim(txt_test.Text) 'opprocedur
        INS_CYTOGENET_QR(19) = Trim(txt_test.Text) 'OPPROCEDUR
    End If
   
    If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
        INSREPORT_QR(7) = Trim(cmb_CCProbes.Text) 'probdesc
        INS_CYTOGENET_QR(57) = Trim(cmb_CCProbes.Text) 'probdesc
    ElseIf Trim(Right(cmb_modality.Text, 2)) = "FI" Then
        INSREPORT_QR(7) = Trim(cmb_probe.Text) 'probdesc
        INS_CYTOGENET_QR(57) = Trim(cmb_probe.Text) 'probdesc
    End If
    
   
    INSREPORT_QR(8) = Trim(FIsh_txt_inter.Text) 'intrpre
   
    INS_CYTOGENET_QR(21) = Trim(FIsh_txt_inter.Text) 'INTRPRE
   
    INSREPORT_QR(9) = Trim(fish_txt_comments.Text) 'comments
   
    INS_CYTOGENET_QR(22) = Trim(fish_txt_comments.Text) 'comments
     
    If opt_chromo.Value = True Then
        INSREPORT_QR(10) = 1 'ca
        INS_CYTOGENET_QR(23) = 1 'ca
    Else
        INSREPORT_QR(10) = 0
        INS_CYTOGENET_QR(23) = 0
    End If
   
    If opt_karyotype.Value = True Then
        INSREPORT_QR(11) = 1 'ka
        INS_CYTOGENET_QR(24) = 1
    Else
        INSREPORT_QR(11) = 0
        INS_CYTOGENET_QR(24) = 0
    End If
    If opt_chromosome.Value = True Then
        INSREPORT_QR(12) = 1 'cb
        INS_CYTOGENET_QR(25) = 1
    Else
        INSREPORT_QR(12) = 0
        INS_CYTOGENET_QR(25) = 0
    End If
   
   'comments,ca,ka,cb,ccnoca,ccintrpre,cccomments,ccpha,ccdir,ccphas,labno,
   
    If Trim(txt_no_of_mata.Text) <> "" Then
        INSREPORT_QR(13) = Trim(txt_no_of_mata.Text) 'ccnoca
        INS_CYTOGENET_QR(26) = Trim(txt_no_of_mata.Text) 'CCNOCA
    Else
        INSREPORT_QR(13) = 0
        INS_CYTOGENET_QR(26) = 0
    End If
   
    'CCINTRPRE CCCOMMENTS CCPHA CCDIR CCPHAS LABNO ENTERBY REPORTBY COMMITBY AT_DIAG PR_TREAT 11
    INSREPORT_QR(14) = Trim(txt_Interpre.Text) 'ccintrpre
   
    INS_CYTOGENET_QR(27) = Trim(txt_Interpre.Text) 'CCINTRPRE
   
    INSREPORT_QR(15) = Trim(txt_comments.Text) 'cccomments
      
    INS_CYTOGENET_QR(28) = Trim(txt_comments.Text) 'CCCOMMENTS
   
   
'   If Chk_method1.Value = 1 Then
'     INSREPORT_QR(16) = Trim(Chk_method1.Caption) 'ccpha
'     INS_CYTOGENET_QR(29) = Trim(Chk_method1.Caption) 'CCPHA
'   Else
'     INSREPORT_QR(16) = ""
'     INS_CYTOGENET_QR(29) = ""
'   End If
   'Commenting/uncommenting above code, doesn't make any difference
   'Still commented :)
   '***** Shraddha 02/12/09 *****
    If Trim(cmb_methodCC.Text) <> "" Then
        INSREPORT_QR(16) = Trim(cmb_methodCC.Text) 'ccpha
        INS_CYTOGENET_QR(29) = Trim(cmb_methodCC.Text) 'CCPHA
    Else
        INSREPORT_QR(16) = ""
        INS_CYTOGENET_QR(29) = ""
    End If
   '*****************************
   
   
    If chk_method2.Value = 1 Then
     INSREPORT_QR(17) = Trim(chk_method2.Caption) 'ccdir
     INS_CYTOGENET_QR(30) = Trim(chk_method2.Caption) 'CCDIR
    Else
     INSREPORT_QR(17) = ""
     INS_CYTOGENET_QR(30) = ""
    End If
   
    If chk_method3.Value = 1 Then
     INSREPORT_QR(18) = Trim(chk_method3.Caption) 'ccphas
     INS_CYTOGENET_QR(31) = Trim(chk_method3.Caption) 'CCPHAS
    Else
     INSREPORT_QR(18) = ""
     INS_CYTOGENET_QR(31) = ""
    End If
    'enterby,reportby,result,commitby,fsituphab
    INSREPORT_QR(19) = Trim(txt_cytology.Text) 'labno
    INS_CYTOGENET_QR(32) = Trim(txt_cytology.Text) 'LABNO
    
    INSREPORT_QR(20) = Trim(cmb_enteryby.Text) 'enterby
    INS_CYTOGENET_QR(33) = Trim(cmb_enteryby.Text) 'ENTERBY
    
    INSREPORT_QR(21) = Trim(cmb_finalby.Text) 'reportby
    INS_CYTOGENET_QR(34) = Trim(cmb_finalby.Text) 'REPORTBY
    
    If Trim(txt_fresult.Text) <> "" Then
    INSREPORT_QR(22) = Trim(txt_fresult.Text) 'result
    INS_CYTOGENET_QR(20) = Trim(txt_fresult.Text)  'result
    Else
    INSREPORT_QR(22) = Trim(txt_result.Text)
    INS_CYTOGENET_QR(20) = Trim(txt_result.Text) 'result
    End If
    
      
    INSREPORT_QR(23) = "" 'commitby
    INS_CYTOGENET_QR(35) = "" 'COMMITBY
   
   
   '***** Shraddha 06/11/09 *****
    CYTOGCLDTA_QR(0) = Trim(req_chglen(cmb_reqno.Text))
    Set CYTOGCLDTA_RS = CYTOGCLDTA_QR.OpenResultset(2, 3)
    
    If Not CYTOGCLDTA_RS.EOF Then
        If CYTOGCLDTA_RS("AT_DIAG") <> "" Then
            INS_CYTOGENET_QR(36) = CYTOGCLDTA_RS!AT_DIaG 'AT_DIAG
        Else
            INS_CYTOGENET_QR(36) = ""
        End If
        If CYTOGCLDTA_RS("PR_TREAT") <> "" Then
            INS_CYTOGENET_QR(37) = CYTOGCLDTA_RS!pr_treat 'PR_TREAT
        Else
            INS_CYTOGENET_QR(37) = ""
        End If
        If CYTOGCLDTA_RS("PS_INDU") <> "" Then
            INS_CYTOGENET_QR(38) = CYTOGCLDTA_RS!ps_indu 'PS_INDU
        Else
            INS_CYTOGENET_QR(38) = ""
        End If
        If CYTOGCLDTA_RS("RELAPSE") <> "" Then
            INS_CYTOGENET_QR(39) = CYTOGCLDTA_RS!relapse 'RELAPSE
        Else
            INS_CYTOGENET_QR(39) = CYTOGCLDTA_RS!relapse 'RELAPSE
        End If
        If CYTOGCLDTA_RS("REMISSION") <> "" Then
            INS_CYTOGENET_QR(40) = CYTOGCLDTA_RS!remission 'REMISSION
        Else
            INS_CYTOGENET_QR(40) = ""
        End If
        If CYTOGCLDTA_RS("FOLLUP") <> "" Then
            INS_CYTOGENET_QR(41) = CYTOGCLDTA_RS!FOLLUP 'FOLLUP
        Else
            INS_CYTOGENET_QR(41) = ""
        End If
        If CYTOGCLDTA_RS("POSTBMT") <> "" Then
            INS_CYTOGENET_QR(42) = CYTOGCLDTA_RS!postbmt 'POSTBMT
        Else
            INS_CYTOGENET_QR(42) = ""
        End If
        If CYTOGCLDTA_RS("B_SEX") <> "" Then
            INS_CYTOGENET_QR(43) = CYTOGCLDTA_RS!B_SEX 'B_SEX
        Else
            INS_CYTOGENET_QR(43) = ""
        End If
        If CYTOGCLDTA_RS("SP_TYPE") <> "" Then
            INS_CYTOGENET_QR(44) = CYTOGCLDTA_RS!SP_TYPE 'SP_TYPE
        Else
            INS_CYTOGENET_QR(44) = ""
        End If
    'INS_CYTOGENET_QR(36) = CYTOGCLDTA_RS!AT_DIAG 'AT_DIAG
    'INS_CYTOGENET_QR(37) = CYTOGCLDTA_RS!PR_TREAT 'PR_TREAT
    'INS_CYTOGENET_QR(38) = CYTOGCLDTA_RS!PS_INDU 'PS_INDU
    'INS_CYTOGENET_QR(39) = CYTOGCLDTA_RS!RELAPSE 'RELAPSE
    'INS_CYTOGENET_QR(40) = CYTOGCLDTA_RS!REMISSION 'REMISSION
    'INS_CYTOGENET_QR(41) = CYTOGCLDTA_RS!FOLLUP 'FOLLUP
    'INS_CYTOGENET_QR(42) = CYTOGCLDTA_RS!POSTBMT 'POSTBMT
    'INS_CYTOGENET_QR(43) = CYTOGCLDTA_RS!B_SEX 'B_SEX
    'INS_CYTOGENET_QR(44) = CYTOGCLDTA_RS!SP_TYPE 'SP_TYPE
   '*****************************
   
   'PS_INDU RELAPSE REMISSION FOLLUP POSTBMT B_SEX SP_TYPE CGDIAG1 CGDIAG2 FAX PHONE MOBILE 12
   
   'Read from CYTOGCLDTA and insert
   
    '***** Shraddha Pandit 09/11/09 *****
'    CYTOGCLDTA_QR(0) = req_chglen(cmb_reqno.Text)
'    If CYTOGCLDTA_FLG = False Then
'        Set CYTOGCLDTA_RS = CYTOGCLDTA_QR.OpenResultset(2, 3)
'        CYTOGCLDTA_FLG = True
'    Else
'        CYTOGCLDTA_RS.MoreResults
'        CYTOGCLDTA_RS.Requery
'    End If
'
'    If Not CYTOGCLDTA_RS.EOF Then
    
' reqno  CASEN0 AT_DIAG PR_TREAT  PS_INDU  RELAPSE REMISSION  FOLLUP POSTBMT  B_SEX
' SP_TYPE CGDIAG1 CGDIAG2  FAX PHONE  CALDATE  CALTIME  NCONTROL NCON_AGE NCON_SEX
' REMARK1 NCON_AGE NCON_SEX REMARK1 REMARK2

        If CYTOGCLDTA_RS("CGDIAG1") <> "" Then
            INS_CYTOGENET_QR(45) = CYTOGCLDTA_RS("CGDIAG1")  'CGDIAG1
        Else
            INS_CYTOGENET_QR(45) = "" 'CGDIAG1
        End If
    
        If CYTOGCLDTA_RS("CGDIAG2") <> "" Then
            INS_CYTOGENET_QR(46) = CYTOGCLDTA_RS("CGDIAG2") 'CGDIAG2
        Else
            INS_CYTOGENET_QR(46) = "" 'CGDIAG2
        End If
        
        If CYTOGCLDTA_RS("FAX") <> "" Then
            INS_CYTOGENET_QR(47) = CYTOGCLDTA_RS("FAX") 'FAX
        Else
            INS_CYTOGENET_QR(47) = "" 'FAX
        End If
        
        If CYTOGCLDTA_RS("PHONE") <> "" Then
            INS_CYTOGENET_QR(48) = CYTOGCLDTA_RS("PHONE") 'PHONE
        Else
            INS_CYTOGENET_QR(48) = "" 'PHONE
        End If
        
        INS_CYTOGENET_QR(49) = "" 'MOBILE
        
        If CYTOGCLDTA_RS("CALDATE") <> Null Then
            INS_CYTOGENET_QR(50) = CYTOGCLDTA_RS("CALDATE")  'CALDATE
        Else
            INS_CYTOGENET_QR(50) = Null 'CALDATE
        End If
        
        If CYTOGCLDTA_RS("CALTIME") <> Null Then
            INS_CYTOGENET_QR(51) = CYTOGCLDTA_RS("CALTIME") 'CALTIME
        Else
            INS_CYTOGENET_QR(51) = Null 'CALTIME
        End If
        
        If CYTOGCLDTA_RS("NCONTROL") <> "" Then
            INS_CYTOGENET_QR(52) = CYTOGCLDTA_RS("NCONTROL") 'NCONTROL
        Else
            INS_CYTOGENET_QR(52) = "" 'NCONTROL
        End If
    
        If CYTOGCLDTA_RS("NCON_AGE") <> Null Then
            INS_CYTOGENET_QR(53) = CYTOGCLDTA_RS("NCON_AGE") 'NCON_AGE
        Else
            INS_CYTOGENET_QR(53) = Null 'NCON_AGE
        End If
        
        If CYTOGCLDTA_RS("NCON_SEX") <> "" Then
            INS_CYTOGENET_QR(54) = CYTOGCLDTA_RS("NCON_SEX") 'NCON_SEX
        Else
            INS_CYTOGENET_QR(54) = "" 'NCON_SEX
        End If
        
        If CYTOGCLDTA_RS("REMARK1") <> "" Then
            INS_CYTOGENET_QR(55) = CYTOGCLDTA_RS("REMARK1") 'REMARKS1
        Else
            INS_CYTOGENET_QR(55) = "" 'REMARKS1
        End If
        
        If CYTOGCLDTA_RS("REMARK2") <> "" Then
            INS_CYTOGENET_QR(56) = CYTOGCLDTA_RS("REMARK2") 'REMARKS2
        Else
            INS_CYTOGENET_QR(56) = "" 'REMARKS2
        End If
        
    Else
    'If there no entry in CYTOGCLDTA
        INS_CYTOGENET_QR(36) = "" 'AT_DIAG
        INS_CYTOGENET_QR(37) = "" 'PR_TREAT
        INS_CYTOGENET_QR(38) = "" 'PS_INDU
        INS_CYTOGENET_QR(39) = "" 'RELAPSE
        INS_CYTOGENET_QR(40) = "" 'REMISSION
        INS_CYTOGENET_QR(41) = "" 'FOLLUP
        INS_CYTOGENET_QR(42) = "" 'POSTBMT
        INS_CYTOGENET_QR(43) = "" 'B_SEX
        INS_CYTOGENET_QR(44) = "" 'SP_TYPE
        INS_CYTOGENET_QR(45) = "" 'CGDIAG1
        INS_CYTOGENET_QR(46) = "" 'CGDIAG2
        INS_CYTOGENET_QR(47) = "" 'FAX
        INS_CYTOGENET_QR(48) = "" 'PHONE
        INS_CYTOGENET_QR(49) = "" 'MOBILE
        INS_CYTOGENET_QR(50) = Null 'CALDATE
        INS_CYTOGENET_QR(51) = Null 'CALTIME
        INS_CYTOGENET_QR(52) = "" 'NCONTROL
        INS_CYTOGENET_QR(53) = Null 'NCON_AGE
        INS_CYTOGENET_QR(54) = "" 'NCON_SEX
        INS_CYTOGENET_QR(55) = "" 'REMARKS1
        INS_CYTOGENET_QR(56) = "" 'REMARKS2
    End If
   '************************************
   
   
    sel_dreq_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    Set sel_dreq_rs = sel_dreq_qr.OpenResultset(1, 3)
    If Not sel_dreq_rs.EOF Then
        If sel_dreq_rs("LATERALITS") <> Null Or sel_dreq_rs("LATERALITS") <> "" Then
            INS_CYTOGENET_QR(58) = sel_dreq_rs("LATERALITS") 'LATERALITS
        Else
            INS_CYTOGENET_QR(58) = Null
        End If
    Else
        INS_CYTOGENET_QR(58) = "" 'LATERALITS
    End If
   
   'CALDATE CALTIME NCONTROL NCON_AGE NCON_SEX REMARKS1 REMARKS2 7
   
    If fish_method3.Value = 1 Then
        INSREPORT_QR(24) = Trim(fish_method3.Caption) 'fsituphab
        INS_CYTOGENET_QR(17) = Trim(fish_method3.Caption) 'FSITUPHAB
    Else
        INSREPORT_QR(24) = ""
        INS_CYTOGENET_QR(17) = "" 'FSITUPHAB
    End If
   
   
   
    INSREPORT_QR.Execute

    '**** INSERTING INTOFREPORT
    
    'ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
    
    updrrequisitn_qr(0) = Trim(cmb_enteryby.Text) 'ENTERBY
    
    updrrequisitn_qr(1) = GET_REVDATE(CStr(DATE11)) 'PREPORTDT
    INS_CYTOGENET_QR(59) = GET_REVDATE(CStr(DATE11))  'PREPORTDT
    
    updrrequisitn_qr(2) = Trim(cmb_finalby.Text) 'PREPORTBY
    updrrequisitn_qr(3) = "T" 'REQNSTAT
    INS_CYTOGENET_QR(4) = "T"
    updrrequisitn_qr(4) = Trim(cmb_finalby.Text) 'freportby
    
    updrrequisitn_qr(5) = GET_REVDATE(CStr(DATE11)) 'freportdt
    INS_CYTOGENET_QR(60) = GET_REVDATE(CStr(DATE11)) '
    
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text)) 'REQNNO
    updrrequisitn_qr.Execute
    
    'PREPORTDT FREPORTDT
    '***** Shraddha 06/11/09 *****
    'PREPORTBY FREPORTBY SAMCOLDATE SAMCOLTIME SAMRECDATE SAMRECTIME COMREPDATE COMREPTIME 8
    INS_CYTOGENET_QR(7) = Trim(cmb_finalby.Text)  'PREPORTBY
    INS_CYTOGENET_QR(8) = Trim(cmb_finalby.Text)  'FREPORTBY
    
    sel_rtime_qr(0) = Trim(req_chglen(cmb_reqno.Text))
    Set sel_rtime_rs = sel_rtime_qr.OpenResultset(1, 3)
    If sel_rtime_rs.EOF Then
    
    End If
    If sel_rtime_rs("SAMCOLDATE") <> "" Or sel_rtime_rs("SAMCOLDATE") <> Null Then
        INS_CYTOGENET_QR(9) = sel_rtime_rs("SAMCOLDATE") 'SAMCOLDATE
    Else
        INS_CYTOGENET_QR(9) = Null
    End If
    If sel_rtime_rs("SAMCOLTIME") <> "" Or sel_rtime_rs("SAMCOLTIME") <> Null Then
        INS_CYTOGENET_QR(10) = sel_rtime_rs("SAMCOLTIME") 'SAMCOLTIME
    Else
        INS_CYTOGENET_QR(10) = Null
    End If
    If sel_rtime_rs("SAMRECDATE") <> "" Or sel_rtime_rs("SAMRECDATE") <> Null Then
        INS_CYTOGENET_QR(11) = sel_rtime_rs("SAMRECDATE") 'SAMRECDATE
    Else
        INS_CYTOGENET_QR(11) = Null
    End If
    If sel_rtime_rs("SAMRECTIME") <> "" Or sel_rtime_rs("SAMRECTIME") <> Null Then
        INS_CYTOGENET_QR(12) = sel_rtime_rs("SAMRECTIME") 'SAMRECTIME
    Else
        INS_CYTOGENET_QR(12) = Null
    End If
    
    sel_rtime_rs.Close
    
    INS_CYTOGENET_QR(13) = Null 'COMREPDATE
    INS_CYTOGENET_QR(14) = Null 'COMREPTIME
    
    INS_CYTOGENET_QR(62) = Trim(ipadd)
    
    INS_CYTOGENET_QR.Execute
    '*****************************
'    upddrequisitn_qr(0) = ""
'    upddrequisitn_qr(1) = "" 'cmb_cellblock.Text
'    If Len(Trim(txt_Interpre.Text)) > 0 Then
'        upddrequisitn_qr(2) = Null
'        upddrequisitn_qr(2) = txt_Interpre.Text
'    ElseIf Len(Trim(FIsh_txt_inter.Text)) Then
'        upddrequisitn_qr(2) = Null
'        upddrequisitn_qr(2) = Trim(FIsh_txt_inter.Text)
'    Else
'        upddrequisitn_qr(2) = Null
'        upddrequisitn_qr(2) = " "
'    End If
'    upddrequisitn_qr(3) = req_chglen(cmb_reqno.Text)
'    upddrequisitn_qr.Execute


 '   **** INSERTING INTO DISIMP
    If Len(Trim(txt_Interpre.Text)) > 0 Or Len(Trim(FIsh_txt_inter.Text)) > 0 Then
    INSDISIMP_QR(0) = req_chglen(cmb_reqno.Text)
    INSDISIMP_QR(1) = Null
    If Trim(txt_Interpre.Text) <> "" Then
        INSDISIMP_QR(1) = Trim(txt_Interpre.Text)
    Else
        INSDISIMP_QR(1) = Trim(FIsh_txt_inter.Text)
    End If
    INSDISIMP_QR.Execute
    End If
  '**** INSERTING INTO DISIMP
  
    
    '   **** inserting into CYTOG_ETO added by v.s.k. on 29.03.2012
    For I = 0 To 9
        If msfAn1.TextMatrix(I, 0) <> "" Or msfAn1.TextMatrix(I, 1) <> "" Or msfAn1.TextMatrix(I, 2) <> "" Or msfAn1.TextMatrix(I, 3) <> "" Or msfAn1.TextMatrix(I, 4) <> "" Or msfAn1.TextMatrix(I, 5) <> "" Or msfAn1.TextMatrix(I, 6) <> "" Then Exit For
    Next
    
    If I < 10 Then
        srno = 1
        For I = 0 To 9
            If msfAn1.TextMatrix(I, 0) <> "" Or msfAn1.TextMatrix(I, 1) <> "" Or msfAn1.TextMatrix(I, 2) <> "" Or msfAn1.TextMatrix(I, 3) <> "" Or msfAn1.TextMatrix(I, 4) <> "" Or msfAn1.TextMatrix(I, 5) <> "" Or msfAn1.TextMatrix(I, 6) <> "" Then
                InsCytog_ETO_qr(0) = req_chglen(cmb_reqno.Text)
                InsCytog_ETO_qr(1) = srno
                InsCytog_ETO_qr(2) = Trim(txtAn1Hd1.Text)
                InsCytog_ETO_qr(3) = Trim(txtAn1Hd2.Text)
                InsCytog_ETO_qr(4) = Trim(msfAn1.TextMatrix(I, 0))
                InsCytog_ETO_qr(5) = Trim(msfAn1.TextMatrix(I, 1))
                InsCytog_ETO_qr(6) = Trim(msfAn1.TextMatrix(I, 2))
                InsCytog_ETO_qr(7) = Trim(msfAn1.TextMatrix(I, 3))
                InsCytog_ETO_qr(8) = Trim(msfAn1.TextMatrix(I, 4))
                InsCytog_ETO_qr(9) = Trim(msfAn1.TextMatrix(I, 5))
                InsCytog_ETO_qr(10) = Trim(msfAn1.TextMatrix(I, 6))
                InsCytog_ETO_qr.Execute
                srno = srno + 1
            End If
        Next
    End If
    
    '   **** inserting into CYTOG_chic
    For I = 0 To 9
        If msfAn2.TextMatrix(I, 0) <> "" Or msfAn2.TextMatrix(I, 1) <> "" Or msfAn2.TextMatrix(I, 2) <> "" Or msfAn2.TextMatrix(I, 3) <> "" Or msfAn2.TextMatrix(I, 4) <> "" Or msfAn1.TextMatrix(I, 5) <> "" Or msfAn1.TextMatrix(I, 6) <> "" Then Exit For
    Next
    
    If I < 10 Then
        srno = 1
        For I = 0 To 9
            If msfAn2.TextMatrix(I, 0) <> "" Or msfAn2.TextMatrix(I, 1) <> "" Or msfAn2.TextMatrix(I, 2) <> "" Or msfAn2.TextMatrix(I, 3) <> "" Or msfAn2.TextMatrix(I, 4) <> "" Or msfAn1.TextMatrix(I, 5) <> "" Or msfAn1.TextMatrix(I, 6) <> "" Then
                InsCytog_CHIC_qr(0) = req_chglen(cmb_reqno.Text)
                InsCytog_CHIC_qr(1) = srno
                InsCytog_CHIC_qr(2) = Trim(txtAn2Hd1.Text)
                InsCytog_CHIC_qr(3) = Trim(txtAn2Hd2.Text)
                InsCytog_CHIC_qr(4) = Trim(msfAn2.TextMatrix(I, 0))
                InsCytog_CHIC_qr(5) = Trim(msfAn2.TextMatrix(I, 1))
                InsCytog_CHIC_qr(6) = Trim(msfAn2.TextMatrix(I, 2))
                InsCytog_CHIC_qr(7) = Trim(msfAn2.TextMatrix(I, 3))
                InsCytog_CHIC_qr(8) = Trim(msfAn2.TextMatrix(I, 4))
                InsCytog_CHIC_qr(9) = Trim(msfAn2.TextMatrix(I, 5))
                InsCytog_CHIC_qr(10) = Trim(msfAn2.TextMatrix(I, 6))
                InsCytog_CHIC_qr.Execute
                srno = srno + 1
            End If
        Next
    End If
    
    '   **** inserting into CYTOG_sp
    For I = 0 To 9
        If msfAn3.TextMatrix(I, 0) <> "" Or msfAn3.TextMatrix(I, 1) <> "" Or msfAn3.TextMatrix(I, 2) <> "" Or msfAn3.TextMatrix(I, 3) <> "" Or msfAn3.TextMatrix(I, 4) <> "" Or msfAn1.TextMatrix(I, 5) <> "" Or msfAn1.TextMatrix(I, 6) <> "" Then Exit For
    Next
    
    If I < 10 Then
        srno = 1
        For I = 0 To 9
            If msfAn3.TextMatrix(I, 0) <> "" Or msfAn3.TextMatrix(I, 1) <> "" Or msfAn3.TextMatrix(I, 2) <> "" Or msfAn3.TextMatrix(I, 3) <> "" Or msfAn3.TextMatrix(I, 4) <> "" Or msfAn1.TextMatrix(I, 5) <> "" Or msfAn1.TextMatrix(I, 6) <> "" Then
                InsCytog_SP_qr(0) = req_chglen(cmb_reqno.Text)
                InsCytog_SP_qr(1) = srno
                InsCytog_SP_qr(2) = Trim(txtAn3Hd1.Text)
                InsCytog_SP_qr(3) = Trim(txtAn3Hd2.Text)
                InsCytog_SP_qr(4) = Trim(msfAn3.TextMatrix(I, 0))
                InsCytog_SP_qr(5) = Trim(msfAn3.TextMatrix(I, 1))
                InsCytog_SP_qr(6) = Trim(msfAn3.TextMatrix(I, 2))
                InsCytog_SP_qr(7) = Trim(msfAn3.TextMatrix(I, 3))
                InsCytog_SP_qr(8) = Trim(msfAn3.TextMatrix(I, 4))
                InsCytog_SP_qr(9) = Trim(msfAn3.TextMatrix(I, 5))
                InsCytog_SP_qr(10) = Trim(msfAn3.TextMatrix(I, 6))
                
                InsCytog_SP_qr.Execute
                srno = srno + 1
            End If
        Next
    End If
   
    '   *** end insert
    MsgBox "Record is saved Successfully."
End If


If flagsave = 2 And ADDENDUMFALG = False Then
    flagsave = 0
    
   UPDREPORT_QR(0) = Trim(Right(cmb_modality.Text, 2))
   'UPD_CYTOGENETmod_QR
   'modalitycd =?,fsitudir =?,fsitupha =?,noca =?,probdesc =?,intrpre =?,comments =?,ca =?,ka =?,cb =?,ccnoca =?,ccintrpre =?,cccomments =?,ccpha =?,ccdir =?,ccphas =?,labno =?,enterby =?,reportby =?,commitby =?,result =?,OPPROCEDUR=?,fsituphab =? where reqno =?"
   
'   If fish_method1.Value = 1 Then
'    UPDREPORT_QR(1) = Trim(fish_method1.Caption)
'    UPD_CYTOGENETmod_QR(0) = Trim(fish_method1.Caption)
'   Else
'    UPDREPORT_QR(1) = ""
'    UPD_CYTOGENETmod_QR(0) = ""
'   End If
   
   '** Commented and added Shraddha 02/12/09 **
   If Trim(cmb_methodFI.Text) <> "" Then
    UPDREPORT_QR(1) = Trim(cmb_methodFI.Text)
    UPD_CYTOGENETmod_QR(0) = Trim(cmb_methodFI.Text)
   Else
    UPDREPORT_QR(1) = ""
    UPD_CYTOGENETmod_QR(0) = ""
   End If
   '*******************************************
   
   If Fish_method2.Value = 1 Then
      UPDREPORT_QR(2) = Trim(Fish_method2.Caption)
      UPD_CYTOGENETmod_QR(1) = Trim(Fish_method2.Caption)
   Else
      UPDREPORT_QR(2) = ""
      UPD_CYTOGENETmod_QR(1) = ""
   End If
   
   If Trim(txt_no_of_cell.Text) <> "" Then
    UPDREPORT_QR(3) = Trim(txt_no_of_cell.Text)
    UPD_CYTOGENETmod_QR(2) = Trim(txt_no_of_cell.Text)
   Else
    UPDREPORT_QR(3) = 0
    UPD_CYTOGENETmod_QR(2) = 0
   End If
   
   '***** Shraddha 07/12/09 *****
    '4 - Probdesc
    '21-OPPROCEDUR
    
    
    If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
        UPDREPORT_QR(4) = Trim(cmb_CCProbes.Text) 'Trim(txt_test.Text)
        UPD_CYTOGENETmod_QR(3) = Trim(cmb_CCProbes.Text) 'Trim(txt_test.Text)
    ElseIf Trim(Right(cmb_modality.Text, 2)) = "FI" Then
        UPDREPORT_QR(4) = Trim(cmb_probe.Text) 'Trim(txt_test.Text)
        UPD_CYTOGENETmod_QR(3) = Trim(cmb_probe.Text) 'Trim(txt_test.Text)
    End If
    '****************************
   UPDREPORT_QR(5) = Trim(FIsh_txt_inter.Text)
   UPD_CYTOGENETmod_QR(4) = Trim(FIsh_txt_inter.Text)
   
   UPDREPORT_QR(6) = Trim(fish_txt_comments.Text)
   UPD_CYTOGENETmod_QR(5) = Trim(fish_txt_comments.Text)
   
   UPDREPORT_QR(7) = opt_chromo.Value
   UPD_CYTOGENETmod_QR(6) = opt_chromo.Value
   
   UPDREPORT_QR(8) = opt_karyotype.Value
   UPD_CYTOGENETmod_QR(7) = opt_karyotype.Value
   
   UPDREPORT_QR(9) = opt_chromosome.Value
   UPD_CYTOGENETmod_QR(8) = opt_chromosome.Value
   
    If Trim(txt_no_of_mata.Text) = "" Then
        UPDREPORT_QR(10) = 0
        UPD_CYTOGENETmod_QR(9) = 0
    Else
        UPDREPORT_QR(10) = Trim(txt_no_of_mata.Text)
        UPD_CYTOGENETmod_QR(9) = Trim(txt_no_of_mata.Text)
    End If
   UPDREPORT_QR(11) = Trim(txt_Interpre.Text)
   UPD_CYTOGENETmod_QR(10) = Trim(txt_Interpre.Text)
   
   UPDREPORT_QR(12) = Trim(txt_comments.Text)
   UPD_CYTOGENETmod_QR(11) = Trim(txt_comments.Text)
   
'   If Chk_method1.Value = 1 Then
'     UPDREPORT_QR(13) = Trim(Chk_method1.Caption)
'     UPD_CYTOGENETmod_QR(12) = Trim(Chk_method1.Caption)
'   Else
'     UPDREPORT_QR(13) = ""
'     UPD_CYTOGENETmod_QR(12) = ""
'   End If
   
   '***** Shraddha 02/12/09 *****
   If cmb_methodCC.Text <> "" Then
     UPDREPORT_QR(13) = Trim(cmb_methodCC.Text)
     UPD_CYTOGENETmod_QR(12) = Trim(cmb_methodCC.Text)
   Else
     UPDREPORT_QR(13) = ""
     UPD_CYTOGENETmod_QR(12) = ""
   End If
   '*****************************
   
   If chk_method2.Value = 1 Then
     UPDREPORT_QR(14) = Trim(chk_method2.Caption)
     UPD_CYTOGENETmod_QR(13) = Trim(chk_method2.Caption)
   Else
     UPDREPORT_QR(14) = ""
     UPD_CYTOGENETmod_QR(13) = ""
   End If
   
   If chk_method3.Value = 1 Then
     UPDREPORT_QR(15) = Trim(chk_method3.Caption)
     UPD_CYTOGENETmod_QR(14) = Trim(chk_method3.Caption)
   Else
     UPDREPORT_QR(15) = ""
     UPD_CYTOGENETmod_QR(14) = ""
   End If
   
   UPDREPORT_QR(16) = Trim(txt_cytology.Text)
   UPD_CYTOGENETmod_QR(15) = Trim(txt_cytology.Text)
   
   UPDREPORT_QR(17) = Trim(cmb_enteryby.Text)
   UPD_CYTOGENETmod_QR(16) = Trim(cmb_enteryby.Text)
   
   UPDREPORT_QR(18) = Trim(cmb_finalby.Text)
   UPD_CYTOGENETmod_QR(17) = Trim(cmb_finalby.Text)
   
   UPDREPORT_QR(19) = Trim(cmb_finalby.Text)
   UPD_CYTOGENETmod_QR(18) = Trim(cmb_finalby.Text)
   
    If Trim(txt_fresult.Text) <> "" Then
        UPDREPORT_QR(20) = Trim(txt_fresult.Text)
        UPD_CYTOGENETmod_QR(19) = Trim(txt_fresult.Text)
    Else
        UPDREPORT_QR(20) = Trim(txt_result.Text)
        UPD_CYTOGENETmod_QR(19) = Trim(txt_result.Text)
    End If
    
    '***** Shraddha 07/12/09 *****
    If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
        UPDREPORT_QR(21) = Trim(cmb_CCTest.Text)
        UPD_CYTOGENETmod_QR(20) = Trim(cmb_CCTest.Text)
    ElseIf Trim(Right(cmb_modality.Text, 2)) = "FI" Then
        UPDREPORT_QR(21) = Trim(txt_test.Text)
        UPD_CYTOGENETmod_QR(20) = Trim(txt_test.Text)
    End If
   '4 - Probdesc
    '21-OPPROCEDUR
   '*******************************
   
   If fish_method3.Value = 1 Then
      UPDREPORT_QR(22) = Trim(fish_method3.Caption)
      UPD_CYTOGENETmod_QR(21) = Trim(fish_method3.Caption)
   Else
      UPDREPORT_QR(22) = ""
      UPD_CYTOGENETmod_QR(21) = ""
   End If
      
   UPDREPORT_QR(23) = req_chglen(Trim(cmb_reqno.Text))
   UPD_CYTOGENETmod_QR(22) = req_chglen(Trim(cmb_reqno.Text))
   
   UPDREPORT_QR.Execute
    UPD_CYTOGENETmod_QR.Execute
    
    
'    upddrequisitn_qr(0) = ""
'    upddrequisitn_qr(1) = cmb_cellblock.Text
'    If Len(Trim(txt_nmeterial.Text)) > 0 Then
'        upddrequisitn_qr(2) = Null
'        upddrequisitn_qr(2) = txt_nmeterial.Text
'    Else
'        upddrequisitn_qr(2) = Null
'        upddrequisitn_qr(2) = " "
'    End If
'    upddrequisitn_qr(3) = req_chglen(cmb_reqno.Text)
'    upddrequisitn_qr.Execute


    '**** updating INTO FREPORT

'    '**** UPDATE THE DISIMP
'    If Len(Trim(txt_impdiag.Text)) > 0 Then
'        UPDDISIMP_QR(0) = Null
'        UPDDISIMP_QR(0) = Trim(txt_impdiag.Text)
'        UPDDISIMP_QR(1) = req_chglen(cmb_reqno.Text)
'        UPDDISIMP_QR.Execute
'    End If
'
'    '*** UPDATE THE DISIMP



'
'    If ADDENDUMFALG = True Then
'            ''********SECURITY***************************
'        CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
'        CURUSRDET_RS.MoreResults
'        CURUSRDET_RS.Requery
'        If CURUSRDET_RS.RowCount > 0 Then
'            If LOGINTYP > CURUSRDET_RS("usrtype") Then
'              ''   no action as he can commit
'            ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
'                MsgBox "Sorry! U'r not an Authorized person to write an Addendum !!", vbInformation, "Pathlogy"
'                cmd_save.SetFocus
'                Exit Sub
'            End If
'            If CURUSRDET_RS("ADDENDUM") <> "YES" Then
'                MsgBox "Sorry! You are not authorised to Write Addendum of this Modality"
'                cmd_save.SetFocus
'                Exit Sub
'            End If
'        Else
'            MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
'            cmd_commit.SetFocus
'            Exit Sub
'        End If
'''********SECURITY***************************
'
'        SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'        SELDADDENDUM_RS.MoreResults
'        SELDADDENDUM_RS.Requery
'        If SELDADDENDUM_RS.RowCount > 0 Then
'            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'            INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
'            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'            INSDADDENDUM_QR(3) = time1
'            INSDADDENDUM_QR(4) = cmb_addendumby.Text
'            INSDADDENDUM_QR(5) = Null
'            INSDADDENDUM_QR(5) = txt_addendum.Text
'            INSDADDENDUM_QR.Execute
'        Else
'            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'            INSDADDENDUM_QR(1) = "1"
'            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'            INSDADDENDUM_QR(3) = time1
'            INSDADDENDUM_QR(4) = cmb_addendumby.Text
'            INSDADDENDUM_QR(5) = Null
'            INSDADDENDUM_QR(5) = txt_addendum.Text
'            INSDADDENDUM_QR.Execute
'        End If
'    End If

'**** Updating into CYTOG_ETO,CYTOG_chic,CYTOG_sp by v.s.k. on 03.04.2012

    str = "delete from dislib.CYTOG_ETO where reqno=?"
    Set delCytog = rdb.CreateQuery("", str)
    delCytog(0) = req_chglen(cmb_reqno.Text)
    delCytog.Execute
'  set 30-03-2012
    str = "delete from dislib.CYTOG_chic where reqno=?"
    Set delCytog = rdb.CreateQuery("", str)
    delCytog(0) = req_chglen(cmb_reqno.Text)
    delCytog.Execute
    
    str = "delete from dislib.CYTOG_sp where reqno=?"
    Set delCytog = rdb.CreateQuery("", str)
    delCytog(0) = req_chglen(cmb_reqno.Text)
    delCytog.Execute
    
    '   **** inserting into CYTOG_ETO added by v.s.k. on 29.03.2012
    For I = 0 To 9
        If Trim(msfAn1.TextMatrix(I, 0)) <> "" Or Trim(msfAn1.TextMatrix(I, 1)) <> "" Or Trim(msfAn1.TextMatrix(I, 2)) <> "" Or Trim(msfAn1.TextMatrix(I, 3)) <> "" Or Trim(msfAn1.TextMatrix(I, 4)) <> "" Or Trim(msfAn1.TextMatrix(I, 5)) <> "" Or Trim(msfAn1.TextMatrix(I, 6)) <> "" Then Exit For
    Next
    
    If I < 10 Then
        srno = 1
        For I = 0 To 9
            If msfAn1.TextMatrix(I, 0) <> "" Or msfAn1.TextMatrix(I, 1) <> "" Or msfAn1.TextMatrix(I, 2) <> "" Or msfAn1.TextMatrix(I, 3) <> "" Or msfAn1.TextMatrix(I, 4) <> "" Or Trim(msfAn1.TextMatrix(I, 5)) <> "" Or Trim(msfAn1.TextMatrix(I, 6)) <> "" Then
                InsCytog_ETO_qr(0) = req_chglen(cmb_reqno.Text)
                InsCytog_ETO_qr(1) = srno
                InsCytog_ETO_qr(2) = Trim(txtAn1Hd1.Text)
                InsCytog_ETO_qr(3) = Trim(txtAn1Hd2.Text)
                InsCytog_ETO_qr(4) = Trim(msfAn1.TextMatrix(I, 0))
                InsCytog_ETO_qr(5) = Trim(msfAn1.TextMatrix(I, 1))
                InsCytog_ETO_qr(6) = Trim(msfAn1.TextMatrix(I, 2))
                InsCytog_ETO_qr(7) = Trim(msfAn1.TextMatrix(I, 3))
                InsCytog_ETO_qr(8) = Trim(msfAn1.TextMatrix(I, 4))
                InsCytog_ETO_qr(9) = Trim(msfAn1.TextMatrix(I, 5))
                InsCytog_ETO_qr(10) = Trim(msfAn1.TextMatrix(I, 6))
                InsCytog_ETO_qr.Execute
                srno = srno + 1
            End If
        Next
    End If
     
 '   **** inserting into CYTOG_chic
    For I = 0 To 9
        If Trim(msfAn2.TextMatrix(I, 0)) <> "" Or Trim(msfAn2.TextMatrix(I, 1)) <> "" Or Trim(msfAn2.TextMatrix(I, 2)) <> "" Or Trim(msfAn2.TextMatrix(I, 3)) <> "" Or Trim(msfAn2.TextMatrix(I, 4)) <> "" Or Trim(msfAn2.TextMatrix(I, 5)) <> "" Or Trim(msfAn2.TextMatrix(I, 6)) <> "" Then Exit For
    Next
    
    If I < 10 Then
        srno = 1
        For I = 0 To 9
            If msfAn2.TextMatrix(I, 0) <> "" Or msfAn2.TextMatrix(I, 1) <> "" Or msfAn2.TextMatrix(I, 2) <> "" Or msfAn2.TextMatrix(I, 3) <> "" Or msfAn2.TextMatrix(I, 4) <> "" Or Trim(msfAn2.TextMatrix(I, 5)) <> "" Or Trim(msfAn2.TextMatrix(I, 6)) <> "" Then
                InsCytog_CHIC_qr(0) = req_chglen(cmb_reqno.Text)
                InsCytog_CHIC_qr(1) = srno
                InsCytog_CHIC_qr(2) = Trim(txtAn2Hd1.Text)
                InsCytog_CHIC_qr(3) = Trim(txtAn2Hd2.Text)
                InsCytog_CHIC_qr(4) = Trim(msfAn2.TextMatrix(I, 0))
                InsCytog_CHIC_qr(5) = Trim(msfAn2.TextMatrix(I, 1))
                InsCytog_CHIC_qr(6) = Trim(msfAn2.TextMatrix(I, 2))
                InsCytog_CHIC_qr(7) = Trim(msfAn2.TextMatrix(I, 3))
                InsCytog_CHIC_qr(8) = Trim(msfAn2.TextMatrix(I, 4))
                InsCytog_CHIC_qr(9) = Trim(msfAn2.TextMatrix(I, 5))
                InsCytog_CHIC_qr(10) = Trim(msfAn2.TextMatrix(I, 6))
                
                InsCytog_CHIC_qr.Execute
                srno = srno + 1
            End If
        Next
    End If
    
    '   **** inserting into CYTOG_sp
    For I = 0 To 9
        If Trim(msfAn3.TextMatrix(I, 0)) <> "" Or Trim(msfAn3.TextMatrix(I, 1)) <> "" Or Trim(msfAn3.TextMatrix(I, 2)) <> "" Or Trim(msfAn3.TextMatrix(I, 3)) <> "" Or Trim(msfAn3.TextMatrix(I, 4)) <> "" Or Trim(msfAn3.TextMatrix(I, 5)) <> "" Or Trim(msfAn3.TextMatrix(I, 6)) <> "" Then Exit For
    Next
    
    If I < 10 Then
        srno = 1
        For I = 0 To 9
            If msfAn3.TextMatrix(I, 0) <> "" Or msfAn3.TextMatrix(I, 1) <> "" Or msfAn3.TextMatrix(I, 2) <> "" Or msfAn3.TextMatrix(I, 3) <> "" Or msfAn3.TextMatrix(I, 4) <> "" Or Trim(msfAn3.TextMatrix(I, 5)) <> "" Or Trim(msfAn3.TextMatrix(I, 6)) <> "" Then
                InsCytog_SP_qr(0) = req_chglen(cmb_reqno.Text)
                InsCytog_SP_qr(1) = srno
                InsCytog_SP_qr(2) = Trim(txtAn3Hd1.Text)
                InsCytog_SP_qr(3) = Trim(txtAn3Hd2.Text)
                InsCytog_SP_qr(4) = Trim(msfAn3.TextMatrix(I, 0))
                InsCytog_SP_qr(5) = Trim(msfAn3.TextMatrix(I, 1))
                InsCytog_SP_qr(6) = Trim(msfAn3.TextMatrix(I, 2))
                InsCytog_SP_qr(7) = Trim(msfAn3.TextMatrix(I, 3))
                InsCytog_SP_qr(8) = Trim(msfAn3.TextMatrix(I, 4))
                InsCytog_SP_qr(9) = Trim(msfAn3.TextMatrix(I, 5))
                InsCytog_SP_qr(10) = Trim(msfAn3.TextMatrix(I, 6))
                InsCytog_SP_qr.Execute
                srno = srno + 1
            End If
        Next
    End If
   
    '   *** end insert
       MsgBox "Record is saved Successfully."
End If

'****************** Shraddha Pandit 25/02/2009 ******************
If ADDENDUMFALG = True Then
        
    If Trim(cmb_addendumby.Text) = "" Then
        
        MsgBox "The Addendum By field is EMPTY."
        chk_Adfld = True
        cmb_addendumby.SetFocus
        Exit Sub
    
    End If
    
    
    ''********SECURITY***************************
    'CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
    'above statement commented and next added by Shraddha Pandit on 30/03/2009
    CURUSRDET_QR(0) = Trim(cmb_addendumby.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! You're not an Authorized person to write an Addendum !!", vbInformation, "Pathlogy"
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
    SELDADDENDUM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    SELDADDENDUM_RS.MoreResults
    SELDADDENDUM_RS.Requery
    If SELDADDENDUM_RS.RowCount > 0 Then
       SELDADDENDUM_RS.MoveLast
        INSDADDENDUM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
        INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
        
        UPDTADNDMTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDTADNDMTIME_QR(1) = time1
        UPDTADNDMTIME_QR(2) = req_chglen(Trim(cmb_reqno.Text))
        UPDTADNDMTIME_QR.Execute
    Else
        INSDADDENDUM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
        INSDADDENDUM_QR(1) = "1"
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
        
        UPDTADNDMTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDTADNDMTIME_QR(1) = time1
        UPDTADNDMTIME_QR(2) = req_chglen(Trim(cmb_reqno.Text))
        UPDTADNDMTIME_QR.Execute
        
    End If
    MsgBox "Addendum added Successfully!"
End If
'****************************************************************
    rdb.CommitTrans
    'MsgBox "Record is saved Successfully."
Exit Sub
'MSG:
    rdb.RollbackTrans
    MsgBox Err.Description

End Sub

Public Sub field_clear()
Dim obj As Object
'lblicu.Visible = False
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
cmb_reqno.ListIndex = -1
cmb_reqno.clear
'cmb_reportdoc.ListIndex = -1
'cmb_addendumby.ListIndex = -1
'cmb_addendumby.Text = ""
cmb_enteryby.ListIndex = -1
cmb_cyspreinv.ListIndex = -1
cmb_cyspreinv.Text = ""
cmb_enteryby.Text = ""
'cmb_provby.ListIndex = -1
cmb_probe.Text = ""
cmb_finalby.ListIndex = -1
cmb_finalby.Text = ""
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
''MaskEdBox1.Mask = ">??/#####" ''commented by Monika on 02/12/2022
Chk_method1.Value = 0
cmb_methodCC.Text = "" 'Shraddha 02/12/09
chk_method2.Value = 0
chk_method3.Value = 0
fish_method1.Value = 0

cmb_methodFI.Text = "" 'Shraddha 02/12/09

cmb_CCProbes.Text = "" 'Shraddha 07/12/09
cmb_CCTest.Text = "" ''Shraddha 07/12/09

Fish_method2.Value = 0
fish_method3.Value = 0
Chk_method1.Value = 0
chk_method2.Value = 0
chk_method3.Value = 0

'*** Added by v.s.k. on 29.03.2012 for ETO/CHIC analysis
msfAn1.Rows = 0
msfAn1.Rows = 10
msfAn1.Cols = 7

msfAn1.ColWidth(0) = 2500
msfAn1.ColWidth(1) = 2500
msfAn1.ColWidth(2) = 2500
msfAn1.ColWidth(3) = 2500
msfAn1.ColWidth(4) = 4500
msfAn1.ColWidth(5) = 2500
msfAn1.ColWidth(6) = 2500


msfAn2.Rows = 0
msfAn2.Rows = 10
msfAn2.Cols = 7

msfAn2.ColWidth(0) = 2500
msfAn2.ColWidth(1) = 2500
msfAn2.ColWidth(2) = 2500
msfAn2.ColWidth(3) = 2500
msfAn2.ColWidth(4) = 4500
msfAn2.ColWidth(5) = 2500
msfAn2.ColWidth(6) = 2500


msfAn3.Rows = 0
msfAn3.Rows = 10
msfAn3.Cols = 7

msfAn3.ColWidth(0) = 2500
msfAn3.ColWidth(1) = 2500
msfAn3.ColWidth(2) = 2500
msfAn3.ColWidth(3) = 2500
msfAn3.ColWidth(4) = 4500
msfAn3.ColWidth(5) = 2500
msfAn3.ColWidth(6) = 2500


txtData.Text = ""
txtData2.Text = ""
txtData3.Text = ""

'TabStrip1
frmAn1.Visible = True
frmAn2.Visible = False
frmAn3.Visible = False
End Sub


Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In cytogenetic_report.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    'txt_addendum.Enabled = True

    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    'cmb_provby.Enabled = True
    cmb_addendumby.Enabled = True 'Shraddha Pandit 25/02/09
    txt_addendum.Enabled = True
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    'txt_report.Enabled = True
    'txt_impdiag.Enabled = True
    cmb_dept.Enabled = True
    cmb_reqno.Enabled = True
    txt_provdiag.Enabled = False
    'txt_sitefnac.Enabled = False
    'txt_noofslides.Enabled = False
    'txt_nmeterial.Enabled = True
    'txt_noofslidesdry.Enabled = False
    'txt_noslideswet.Enabled = False
    'cmb_cellblock.Enabled = True
    cmb_methodFI.Enabled = True
    txt_no_of_cell.Enabled = True
    txt_fresult.Enabled = True
    FIsh_txt_inter.Enabled = True
    fish_txt_comments.Enabled = True
    txt_test.Enabled = True
    cmb_probe.Enabled = True
    txt_test.Locked = False
    cmb_probe.Locked = False
End If
If str1 = "DIS" Then
    'txt_addendum.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    'cmb_provby.Enabled = False
    'cmb_addendumby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_reqno.Enabled = False
    'txt_report.Enabled = False
    'txt_impdiag.Enabled = False
    'txt_provdiag.Enabled = False
    'txt_sitefnac.Enabled = False
    'txt_noofslides.Enabled = False
    'txt_nmeterial.Enabled = False
    'txt_noofslidesdry.Enabled = False
    'txt_noslideswet.Enabled = False
    'cmb_cellblock.Enabled = False
End If
If str1 = "VIEW" Then
    
    'txt_addendum.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    'cmb_provby.Enabled = False
    'cmb_addendumby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    'txt_report.Enabled = False
    'txt_impdiag.Enabled = False
    'txt_provdiag.Enabled = False
    'txt_sitefnac.Enabled = False
    'txt_noofslides.Enabled = False
    'txt_nmeterial.Enabled = False
    'txt_noofslidesdry.Enabled = False
    'txt_noslideswet.Enabled = False
    'cmb_cellblock.Enabled = False
End If
If str1 = "MODIFY" Then
    txt_no_of_mata.Enabled = True
    txt_result.Enabled = True
    txt_Interpre.Enabled = True
    txt_comments.Enabled = True
    'txt_addendum.Enabled = False

    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    'cmb_provby.Enabled = True
    'cmb_addendumby.Enabled = False
    'cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    cmb_CCProbes.Enabled = True
    'txt_report.Enabled = True
    'txt_impdiag.Enabled = True
    'txt_provdiag.Enabled = False
    'txt_sitefnac.Enabled = False
    'txt_noofslides.Enabled = False
    'txt_nmeterial.Enabled = True
    'txt_noofslidesdry.Enabled = False
    'txt_noslideswet.Enabled = False
    'cmb_cellblock.Enabled = True
    cmb_probe.Enabled = True
    txt_test.Enabled = True
        cmb_probe.Locked = False
    txt_test.Locked = False

End If
If str1 = "ADDENDUM" Then
    '********* Shraddha Pandit 26/02/09 **********
    txt_addendum.Enabled = True
    cmb_addendumby.Enabled = True
    '*********************************************
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    'cmb_provby.Enabled = False
    
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    'txt_report.Enabled = False
    'txt_impdiag.Enabled = False
    'txt_provdiag.Enabled = False
    'txt_sitefnac.Enabled = False
    'txt_noofslides.Enabled = False
    'txt_nmeterial.Enabled = False
    'txt_noofslidesdry.Enabled = False
    'txt_noslideswet.Enabled = False
    'cmb_cellblock.Enabled = False
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
        txt_refdoc.Text = REFDR_RS("REFDOC")
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
        
        txt_cytology.Text = selpathdrequisitn_rs("PATHNO") 'Shraddha 27/12/10
        
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
     'prpathno = ""
     While Not seldrequisitn_rs.EOF
     'prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
     seldrequisitn_rs.MoveNext
     Wend
'     txt_prepathreqno.Text = prpathno
    End If
    'txt_addendum.Enabled = False
    'cmb_addendumby.Enabled = False
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
    '***** Shraddha 04/10/10 *****
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
       '***** Shraddha 07/10/10 *****
        
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
        txt_refdoc.Text = REFDR_RS("REFDOC")
    End If
        
       '*****************************
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
        
        
        sel_rreq_qr(0) = Trim(req_chglen(cmb_reqno.Text))
        Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
        
        'If viewflag = True Then
        If sel_rreq_rs.RowCount > 0 Then
            'txt_refdoc.Text = REFDR_RS("REFBYDOC")
            txt_refdoc.Text = REFDR_RS("REFDOC")
            If sel_rreq_rs("APNTDT") <> "" Then
                txt_rec_on.Text = sel_rreq_rs("APNTDT")
            Else
                txt_rec_on.Text = ""
            End If
            If Trim(sel_rreq_rs("ENTERBY")) <> "" Then
                cmb_enteryby.Text = Trim(sel_rreq_rs("ENTERBY"))
            Else
                cmb_enteryby.Text = ""
            End If
            
            If Trim(sel_rreq_rs("PREPORTBY")) <> "" Then
                cmb_finalby.Text = Trim(sel_rreq_rs("PREPORTBY"))
            Else
                cmb_finalby.Text = ""
            End If
            
            
            If sel_rreq_rs("COMMITFLAG") = "Y" Then
                'MsgBox "Final Report is Generated, You can't Modify the Final Report,You can Add ADDENDUMS Only"
                'cmd_modify.Enabled = True
                MsgBox "Final Report is Generated, You can't Modify the Final Report ,You can Add ADDENDUMS Only"
                cmd_modify.Enabled = True
                'cmd_modify.Enabled = False
                cmd_commit.Enabled = False
                
               If IsNull(Trim(sel_rreq_rs("FREPORTBY"))) Then
                   cmb_finalby.Text = ""
                Else
                   cmb_finalby.Text = Trim(sel_rreq_rs("FREPORTBY"))
               End If
                '************** Shraddha Pandit 25/02/09 ***************
                ADDENDUMFALG = True
                frame_addendum.Visible = True
                lbl_addendum.Visible = True
                txt_addendum.Visible = True
                txt_addendum.Enabled = False
                lbl_addendumby.Visible = True
                cmb_addendumby.Visible = True
                cmb_addendumby.Enabled = False
                
                txt_no_of_cell.Enabled = False
                txt_fresult.Enabled = False
                FIsh_txt_inter.Enabled = False
                fish_txt_comments.Enabled = False
                txt_test.Enabled = False
                cmb_probe.Enabled = False
                
                txt_no_of_mata.Enabled = False
                txt_result.Enabled = False
                txt_Interpre.Enabled = False
                txt_comments.Enabled = False
                '*******************************************************
            End If
          ' ALLOWING CHANGE OF PROBE & TEST FOR T108, T229 & T230
            
         If sel_rreq_rs("SER_CODE") = "T108" Or sel_rreq_rs("SER_CODE") = "T229" Or sel_rreq_rs("SER_CODE") = "T230" Then
          txt_test.Locked = False
          cmb_probe.Locked = False
          
         Else
           txt_test.Locked = True
           cmb_probe.Locked = True
          
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
        
        SELREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
        If SELREPORT_FLAG = False Then
           Set SELREPORT_RS = SELREPORT_QR.OpenResultset(2, 3)
           SELREPORT_FLAG = True
        Else
          SELREPORT_RS.MoreResults
          SELREPORT_RS.Requery
        End If
        If SELREPORT_RS.RowCount > 0 Then
                If Trim(SELREPORT_RS!modalitycd) = "CC" Then
                   frm_Conventional.Visible = True
                   Fish_modality.Visible = False
                End If
                If Trim(SELREPORT_RS!modalitycd) = "FI" Then
                    frm_Conventional.Visible = False
                    Fish_modality.Visible = True
                End If
                
                cmb_modality.Text = Trim(SELREPORT_RS!modalitycd)
                If Trim(SELREPORT_RS!FSITUDIR) = "" Then
                  fish_method1.Visible = True
                  fish_method1.Value = 0
                Else
                  fish_method1.Visible = True
                  fish_method1.Value = 1
                End If
                
                '***** Shraddha 02/12/09 *****
                If Trim(SELREPORT_RS!FSITUDIR) = "" Then
                  fish_method1.Visible = True
                  fish_method1.Value = 0
                  cmb_methodFI.Text = ""
                Else
                  fish_method1.Visible = True
                  fish_method1.Value = 1
                  cmb_methodFI.Text = Trim(SELREPORT_RS!FSITUDIR)
                End If
                'cmb_methodFI
                'In either modality methods should be available to select
                'So lets not disturb those combos (cmb_methodCC,cmb_methodFI)
                'Under view
                '*****************************
                
                '***** Shraddha 07/12/09 *****
                If Trim(SELREPORT_RS!OPPROCEDUR) <> "" Then
                    cmb_CCTest.Text = Trim(SELREPORT_RS!OPPROCEDUR)
                End If
                If Trim(SELREPORT_RS!PROBDESC) <> "" Then
                    cmb_CCProbes.Text = Trim(SELREPORT_RS!PROBDESC)
                End If
                '*****************************
                             
                If Trim(SELREPORT_RS!fsitupha) = "" Then
                 Fish_method2.Visible = True
                 Fish_method2.Value = 0
                Else
                 Fish_method2.Visible = True
                 Fish_method2.Value = 1
                End If
                
                If Trim(SELREPORT_RS!FSITUPHAB) = "" Then
                 fish_method3.Visible = True
                 fish_method3.Value = 0
                Else
                 fish_method3.Visible = True
                 fish_method3.Value = 1
                End If
                
                txt_no_of_cell.Text = Trim(SELREPORT_RS!NOCA)
                txt_test.Text = Trim(SELREPORT_RS!OPPROCEDUR)
                cmb_probe.Text = Trim(SELREPORT_RS!PROBDESC)
                FIsh_txt_inter.Text = Trim(SELREPORT_RS!INTRPRE)
                fish_txt_comments.Text = Trim(SELREPORT_RS!Comments)
                txt_no_of_mata.Text = Trim(SELREPORT_RS!CCNOCA)
                txt_Interpre.Text = Trim(SELREPORT_RS!CCINTRPRE)
                txt_comments.Text = Trim(SELREPORT_RS!CCCOMMENTS)
                If Trim(Right(cmb_modality.Text, 2)) = "CC" Then
                 txt_result.Text = Trim(SELREPORT_RS!Result)
                Else
                 txt_fresult.Text = Trim(SELREPORT_RS!Result)
                End If
                
                If Trim(SELREPORT_RS!CCPHA) = "" Then
                  Chk_method1.Visible = True
                  Chk_method1.Value = 0
                Else
                  Chk_method1.Value = 1
                  Chk_method1.Visible = True
                End If
                
                '***** Shraddha 02/12/09 *****
                
                If Trim(SELREPORT_RS!CCPHA) = "" Then
                  Chk_method1.Visible = True
                  Chk_method1.Value = 0
                  cmb_methodCC.Text = ""
                Else
                  Chk_method1.Value = 1
                  Chk_method1.Visible = True
                  cmb_methodCC.Text = Trim(SELREPORT_RS!CCPHA)
                End If
                
                'In either modality methods should be available to select
                'So lets not disturb those combos (cmb_methodCC,cmb_methodFI)
                'Under view
                '*****************************
                
                
                
                If Trim(SELREPORT_RS!CCDIR) = "" Then
                  chk_method2.Visible = True
                  chk_method2.Value = 0
                Else
                  chk_method2.Value = 1
                  chk_method2.Visible = True
                End If
                
                If Trim(SELREPORT_RS!CCPHAS) = "" Then
                  chk_method3.Visible = True
                  chk_method3.Value = 0
                Else
                  chk_method3.Visible = True
                  chk_method3.Value = 1
                End If
                txt_cytology.Text = Trim(SELREPORT_RS!LABNO)
                cmb_enteryby.Text = Trim(SELREPORT_RS!ENTERBY)
                cmb_finalby.Text = Trim(SELREPORT_RS!REPORTBY)
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
'            If Trim(SELREPORT_RS!MODALITYCD) = "FI" Then
            If Trim(selpathdrequisitn_rs("NMETERIALS")) <> "" Then
               txt_specimen.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
            Else
                txt_specimen.Text = ""
            End If
'            Else
'               txt_Interpre.Text = Trim(selpathdrequisitn_rs("nmeterials"))
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

        If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
        Dim prpathno
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
'        SELDADDENDUM_QR(0) = reqno
'            SELDADDENDUM_RS.MoreResults
'            SELDADDENDUM_RS.Requery
'            If SELDADDENDUM_RS.RowCount > 0 Then
'                While Not SELDADDENDUM_RS.EOF
'                    ADDENDUM = ADDENDUM + Chr(13) & "ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
'                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
'                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
'                SELDADDENDUM_RS.MoveNext
'                Wend
'                txt_addendum.Text = Trim(ADDENDUM)
'                ADDENDUMFALG = True
'            End If
'********************** Shraddha Pandit 25/02/09 *************
            Dim ADDENDUM As String

            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            SELDADDENDUM_RS.MoreResults
            SELDADDENDUM_RS.Requery
            If SELDADDENDUM_RS.RowCount > 0 Then
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
'**************************************************************
        'End If
        cmb_reqno.Enabled = False
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
    End If
    '*****************************
    cmb_reqno.Enabled = False
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    
Exit Sub
MSG:
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
Dim str1 As String
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
Dim str2 As String
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
        'STR1 = ip
        str2 = DATE11
        'STR3 = time1
        'STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
        'STR1 = Space(3) & STR1
        'CrystalReport1.Formulas(3) = "us_id = '" & STR1 & "'"
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
'        STR1 = ip
'        STR2 = DATE11
'        STR3 = time1
'        STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
'        STR1 = Space(3) & STR1
        CrystalReport1.Formulas(3) = "us_id = '"  'STR1 & "'"
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
                'STR1 = ""
                'STR1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(2) = "ENT = '" ' & STR1 & "'"
             Else

                CrystalReport1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(2) = "ENT = '" '& STR1 & "'"
            End If
            CrystalReport1.Formulas(3) = "USRID = '" '& STR & "'"
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
                'STR1 = ""
                'STR1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" '& STR1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" '& STR1 & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" '& STR & "'"
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
    'prpathno = ""
    'CMB_PATHNO.clear
    While Not seldrequisitn_rs11.EOF
        'prpathno = prpathno & " " & seldrequisitn_rs11("pathno")
        'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
        seldrequisitn_rs11.MoveNext
    Wend
    'txt_prepathreqno.Text = prpathno
End If
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
             CrystalReport1.Formulas(1) = "PREPATH = 'Prev. Path No." '& Trim(prpathno) & "'"
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
    cmd_commit.Enabled = True
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
    '30/03/09
    If chk_Adfld = True Then
    cmd_save.Enabled = True
    Else
    cmd_save.Enabled = False
    End If
    
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

'Private Sub cmd_back_Click()
'Unload Me
'End Sub


Private Sub MaskEdBox1_LostFocus()
 Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
        ''MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 02/12/2022
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
         If NEWFLAG = True Then
        
          If REQFLAG = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(Right(cmb_modality.Text, 2))
                Set req_rs = req_qr.OpenResultset(1, 3)
                REQFLAG = True
          Else
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(Right(cmb_modality.Text, 2))
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
                    SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
                    SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                    SELRREQUISITN_QR(3) = "F"
                    SELRREQUISITN_QR(4) = "F"
                    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                    SELRREQUISITNFLAG = True
                Else
                    SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                    SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
                    SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                    SELRREQUISITN_QR(3) = "F"
                    SELRREQUISITN_QR(4) = "F"
                    SELRREQUISITN_RS.MoreResults
                    SELRREQUISITN_RS.Requery
                End If
            Else
            If SELRREQUISITNFLAG = False Then
                
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
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
    
       'If REQFLAG = False Then
             req_qr(0) = MaskEdBox1.FormattedText
             req_qr(1) = Trim(Right(cmb_modality.Text, 2))
             Set req_rs = req_qr.OpenResultset(1, 3)
            ' REQFLAG = True
'          Else
'             req_qr(0) = MaskEdBox1.FormattedText
'             req_qr(1) = Trim(Right(Cmb_modality.Text, 2))
'             req_rs.MoreResults
'             req_rs.Requery
       '   End If
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
                SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
        Else
           If SELRREQUISITNFLAG = False Then
                
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(Right(cmb_modality.Text, 2))
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
'previous invetigation
  selreqnno_qr(0) = MaskEdBox1.FormattedText
  If selreqnno_flg = False Then
   Set selreqnno_rs = selreqnno_qr.OpenResultset(2, 3)
   selreqnno_flg = True
  Else
   selreqnno_rs.MoreResults
   selreqnno_rs.Requery
  End If
  
  If selreqnno_rs.RowCount > 0 Then
     
     selreqnno_rs.MoveFirst
     While Not selreqnno_rs.EOF
       selreqmoda_qr(0) = selreqnno_rs(0)
       Set selreqmoda_RS = selreqmoda_qr.OpenResultset(2, 3)
       cmb_cyspreinv.AddItem selreqnno_rs(0) & "   " & selreqmoda_RS(0)
     selreqnno_rs.MoveNext
     Wend
  End If
  
cmb_reqno.SetFocus
 Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub opt_chromo_Click()
If opt_chromo.Value = True Then
    Chk_method1.Visible = True
    chk_method2.Visible = False
    chk_method3.Visible = False
    'txt_result = " "
    
End If
End Sub

Private Sub opt_chromosome_Click()
'If opt_chromosome.Value = True Then
  'txt_result.Text = "SPONTANEOUS BREAKAGE FREQUENCY IN PATIENT IS      BREAKS/CELL" & Chr(13) & _
                    "SPONTANEOUS BREAKAGE FREQUENCY IN CONTROL IS      BREAKS/CELL" & Chr(13) & _
                    "MITOMYCIN C-(MMC) INDUCED CHROMOSOMAL BREAKAGE FREQUENCY IN PATIENT IS      BREAKS/CELL" & Chr(13) & _
                    "MITOMYCIN C-(MMC) INDUCED CHROMOSOMAL BREAKAGE FREQUENCY IN CONTROL IS      BREAKS/CELL"
'Else
'  txt_result.Text = ""
'End If
    Chk_method1.Visible = False
    chk_method2.Visible = False
    chk_method3.Visible = True
End Sub

Private Sub opt_karyotype_Click()
    Chk_method1.Visible = True
    chk_method2.Visible = True
    chk_method3.Visible = False
    'txt_result = ""
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
    frmAn1.Visible = True
    frmAn2.Visible = False
    frmAn3.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
    frmAn1.Visible = False
    frmAn2.Visible = True
    frmAn3.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 3 Then
    frmAn1.Visible = False
    frmAn2.Visible = False
    frmAn3.Visible = True
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


Private Sub TXT_COMMENTS_GotFocus()
On Error GoTo m14:
w.Visible = False
Set D = w.ActiveDocument
D.Content = ""
D.Content = txt_comments
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
D.Content = txt_comments.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set D = w.ActiveDocument
D.Content = txt_comments.Text
w.Visible = True
w.Activate
  
End Sub

Private Sub txt_comments_lostfocus()
 On Error GoTo m15:
Set D = w.ActiveDocument
txt_comments.Text = D.Content
'flagimp = False
D.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
D.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub txt_fresult_GotFocus()
On Error GoTo m14:
w.Visible = False
Set D = w.ActiveDocument
D.Content = ""
D.Content = txt_fresult
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
D.Content = txt_fresult.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set D = w.ActiveDocument
D.Content = txt_fresult.Text
w.Visible = True
w.Activate
End Sub

Private Sub txt_fresult_LostFocus()
 On Error GoTo m15:
Set D = w.ActiveDocument
txt_fresult.Text = D.Content
'flagimp = False
D.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
D.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub txt_Interpre_GotFocus()
On Error GoTo m14:
w.Visible = False
Set D = w.ActiveDocument
D.Content = ""
D.Content = txt_Interpre
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
D.Content = txt_Interpre.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set D = w.ActiveDocument
D.Content = txt_Interpre.Text
w.Visible = True
w.Activate
End Sub

Private Sub txt_Interpre_LostFocus()
On Error GoTo m15:
Set D = w.ActiveDocument
txt_Interpre.Text = D.Content
'flagimp = False
D.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
D.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub txt_result_gotfocus()
 On Error GoTo m14:
w.Visible = False
Set D = w.ActiveDocument
D.Content = ""
D.Content = txt_result.Text
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
D.Content = txt_Interpre.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set D = w.ActiveDocument
D.Content = txt_result.Text
w.Visible = True
w.Activate
End Sub

Private Sub txt_result_lostfocus()
On Error GoTo m15:
Set D = w.ActiveDocument
txt_result.Text = D.Content
'flagimp = False
D.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
D.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub txtAn1Hd1_Click()
msfAn1.Rows = 0
msfAn1.Rows = 12
selHead(0) = txtAn1Hd1.Text
Set rsSelCytog = selHead.OpenResultset(2, 3)
txtAn1Hd2.Text = rsSelCytog("head2")
I = 0
While Not rsSelCytog.EOF
    If msfAn1.Rows = I Then
        msfAn1.Rows = msfAn1.Rows + 2
    End If
    msfAn1.TextMatrix(I, 0) = rsSelCytog("col1")
    msfAn1.TextMatrix(I, 1) = rsSelCytog("col2")
    msfAn1.TextMatrix(I, 2) = rsSelCytog("col3")
    msfAn1.TextMatrix(I, 3) = rsSelCytog("col4")
    msfAn1.TextMatrix(I, 4) = rsSelCytog("col5")
    msfAn1.TextMatrix(I, 5) = rsSelCytog("col6")
    msfAn1.TextMatrix(I, 6) = rsSelCytog("col7")
    I = I + 2
    rsSelCytog.MoveNext
Wend
rsSelCytog.Close

End Sub

Private Sub txtAn2Hd1_Click()
msfAn2.Rows = 0
msfAn2.Rows = 12
selHead(0) = txtAn2Hd1.Text
Set rsSelCytog = selHead.OpenResultset(2, 3)
txtAn2Hd2.Text = rsSelCytog("head2")
I = 0
While Not rsSelCytog.EOF
    If msfAn2.Rows = I Then
        msfAn2.Rows = msfAn2.Rows + 2
    End If
    msfAn2.TextMatrix(I, 0) = rsSelCytog("col1")
    msfAn2.TextMatrix(I, 1) = rsSelCytog("col2")
    msfAn2.TextMatrix(I, 2) = rsSelCytog("col3")
    msfAn2.TextMatrix(I, 3) = rsSelCytog("col4")
    msfAn2.TextMatrix(I, 4) = rsSelCytog("col5")
    msfAn2.TextMatrix(I, 5) = rsSelCytog("col6")
    msfAn2.TextMatrix(I, 6) = rsSelCytog("col7")
    I = I + 2
    rsSelCytog.MoveNext
Wend
rsSelCytog.Close
End Sub

Private Sub txtAn3Hd1_Click()
msfAn3.Rows = 0
msfAn3.Rows = 12
selHead(0) = txtAn3Hd1.Text
Set rsSelCytog = selHead.OpenResultset(2, 3)
txtAn3Hd2.Text = rsSelCytog("head2")
I = 0
While Not rsSelCytog.EOF
    If msfAn3.Rows = I Then
        msfAn3.Rows = msfAn1.Rows + 2
    End If
    msfAn3.TextMatrix(I, 0) = rsSelCytog("col1")
    msfAn3.TextMatrix(I, 1) = rsSelCytog("col2")
    msfAn3.TextMatrix(I, 2) = rsSelCytog("col3")
    msfAn3.TextMatrix(I, 3) = rsSelCytog("col4")
    msfAn3.TextMatrix(I, 4) = rsSelCytog("col5")
    msfAn3.TextMatrix(I, 5) = rsSelCytog("col6")
    msfAn3.TextMatrix(I, 6) = rsSelCytog("col7")
    I = I + 2
    rsSelCytog.MoveNext
Wend
rsSelCytog.Close
End Sub

Private Sub txtData_Change()
msfAn1.TextMatrix(msfAn1.row, msfAn1.col) = txtData.Text
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If msfAn1.col = 6 Then
        msfAn1.col = 0
        If msfAn1.row = 9 Then
            msfAn1.row = 0
        Else
            msfAn1.row = msfAn1.row + 1
        End If
    Else
        msfAn1.col = msfAn1.col + 1
    End If
End If
End Sub

Private Sub txtData_LostFocus()
'msfAn1.TextMatrix(msfAn1.row, msfAn1.col) = Trim(txtData.Text)
End Sub

Private Sub txtData2_Change()
msfAn2.TextMatrix(msfAn2.row, msfAn2.col) = txtData2.Text
End Sub

Private Sub txtData2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If msfAn2.col = 6 Then
        msfAn2.col = 0
        If msfAn2.row = 9 Then
            msfAn2.row = 0
        Else
            msfAn2.row = msfAn2.row + 1
        End If
    Else
        msfAn2.col = msfAn2.col + 1
    End If
End If
End Sub

Private Sub txtData3_Change()
msfAn3.TextMatrix(msfAn3.row, msfAn3.col) = txtData3.Text
End Sub

Private Sub txtData3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If msfAn3.col = 6 Then
        msfAn3.col = 0
        If msfAn3.row = 9 Then
            msfAn3.row = 0
        Else
            msfAn3.row = msfAn3.row + 1
        End If
    Else
        msfAn3.col = msfAn3.col + 1
    End If
End If
End Sub
