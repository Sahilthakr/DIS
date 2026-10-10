VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form BIOCHEMISTRY1 
   Caption         =   "Biochemistry"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frame_EDL 
      Caption         =   "Estimation of Drug Level"
      ForeColor       =   &H000000C0&
      Height          =   2295
      Left            =   240
      TabIndex        =   172
      Top             =   2880
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CheckBox chk_cyclos 
         Caption         =   "Cyclosporin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   183
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chk_amik 
         Caption         =   "Amikacin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   182
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chk_vanc 
         Caption         =   "Vancomycin"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   181
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chk_metho 
         Caption         =   "Methotrexate"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6840
         TabIndex        =   180
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txt_Dosage 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   179
         Top             =   0
         Visible         =   0   'False
         Width           =   7815
      End
      Begin VB.TextBox txt_amt 
         Height          =   285
         Left            =   1920
         TabIndex        =   178
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmb_unit 
         Height          =   315
         ItemData        =   "BIOCHEMISTRY1.frx":0000
         Left            =   10440
         List            =   "BIOCHEMISTRY1.frx":000D
         TabIndex        =   177
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TXT_HH 
         Height          =   285
         Left            =   5040
         TabIndex        =   176
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXT_MM 
         Height          =   285
         Left            =   5640
         TabIndex        =   175
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CMB_FR 
         Height          =   315
         ItemData        =   "BIOCHEMISTRY1.frx":001D
         Left            =   6240
         List            =   "BIOCHEMISTRY1.frx":0027
         TabIndex        =   174
         Top             =   1680
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txt_hrs 
         Height          =   285
         Left            =   4440
         TabIndex        =   173
         Top             =   1200
         Width           =   855
      End
      Begin MSMask.MaskEdBox msk_fromdt 
         Height          =   255
         Left            =   2880
         TabIndex        =   184
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Drug Dosage"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   189
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dosage of Drug Given at "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   188
         Top             =   1680
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label8 
         Caption         =   "HH   :   MM"
         Height          =   255
         Left            =   5040
         TabIndex        =   187
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "mg and after"
         Height          =   255
         Left            =   3240
         TabIndex        =   186
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "hours"
         Height          =   255
         Left            =   5400
         TabIndex        =   185
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame urinerandom 
      Height          =   6255
      Left            =   1080
      TabIndex        =   127
      Top             =   1680
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame Frame16 
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
         Height          =   1335
         Left            =   8040
         TabIndex        =   141
         Top             =   2880
         Width           =   2895
         Begin VB.CheckBox chk_urobur 
            Caption         =   "Urobilinogen"
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
            Height          =   375
            Left            =   480
            TabIndex        =   143
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chk_Porphyrinur 
            Caption         =   "Porphyrins"
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
            Height          =   255
            Left            =   480
            TabIndex        =   142
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   1335
         Left            =   720
         TabIndex        =   137
         Top             =   720
         Width           =   4455
         Begin VB.CheckBox chk_ureaur 
            Caption         =   "Urea"
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
            Left            =   360
            TabIndex        =   140
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chk_creatur 
            Caption         =   "Creatinine"
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
            Left            =   360
            TabIndex        =   139
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox chk_uricacidur 
            Caption         =   "Uric Acid"
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
            Left            =   2280
            TabIndex        =   138
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   1335
         Left            =   5520
         TabIndex        =   133
         Top             =   720
         Width           =   5415
         Begin VB.CheckBox chk_sodiumur 
            Caption         =   "Sodium"
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
            Left            =   600
            TabIndex        =   136
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chk_potaur 
            Caption         =   "Potassium"
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
            Left            =   600
            TabIndex        =   135
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chk_chlorur 
            Caption         =   "Chlorides"
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
            TabIndex        =   134
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   600
         TabIndex        =   130
         Top             =   2880
         Width           =   4455
         Begin VB.CheckBox chk_umrandom 
            Caption         =   "URINARY MAGNESIUM(RANDOM)*"
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
            Left            =   360
            TabIndex        =   170
            Top             =   960
            Width           =   3495
         End
         Begin VB.CheckBox chk_calcur 
            Caption         =   "Calcium"
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
            Left            =   360
            TabIndex        =   132
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chk_phospur 
            Caption         =   "Phosphorus"
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
            Left            =   2160
            TabIndex        =   131
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1335
         Left            =   5520
         TabIndex        =   128
         Top             =   2880
         Width           =   2295
         Begin VB.CheckBox chk_protur 
            Caption         =   "Proteins"
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
            Left            =   600
            TabIndex        =   129
            Top             =   480
            Width           =   1575
         End
      End
   End
   Begin VB.Frame blood 
      Height          =   7335
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CheckBox serumB2_Microglobin 
         Caption         =   "Serum B2 - Microglobulin"
         BeginProperty Font 
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
         TabIndex        =   190
         Top             =   6600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame20 
         Height          =   855
         Left            =   5400
         TabIndex        =   166
         Top             =   5280
         Width           =   1575
         Begin VB.CheckBox chk_lactate 
            Caption         =   "Lactate"
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
            TabIndex        =   167
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame19 
         Height          =   1095
         Left            =   3480
         TabIndex        =   163
         Top             =   5160
         Width           =   1815
         Begin VB.CheckBox chk_se_cpk_mb 
            Caption         =   "CPK-MB"
            BeginProperty Font 
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
            TabIndex        =   165
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chk_se_cpk 
            Caption         =   "CPK"
            BeginProperty Font 
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
            TabIndex        =   164
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame18 
         Height          =   855
         Left            =   1920
         TabIndex        =   158
         Top             =   5280
         Width           =   1455
         Begin VB.CheckBox chk_crp 
            Caption         =   "CRP"
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
            TabIndex        =   159
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CheckBox chk_ggt 
         Caption         =   "GGT"
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
         Left            =   360
         TabIndex        =   156
         Top             =   5520
         Width           =   855
      End
      Begin VB.Frame SP 
         Caption         =   "Sugar Profile"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   6375
         Begin VB.CheckBox chk_trusugfas 
            Caption         =   "True Sugar Fasting"
            BeginProperty Font 
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
            TabIndex        =   18
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chk_ppsug 
            Caption         =   "PP Sugar"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk_ransug 
            Caption         =   "Random Sugar"
            BeginProperty Font 
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
            TabIndex        =   16
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chk_glucosug 
            Caption         =   "Glucocheck Sugar"
            BeginProperty Font 
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
            TabIndex        =   15
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox chk_glycahbb 
            Caption         =   "Glycateo H.B.B"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            MaskColor       =   &H00808080&
            TabIndex        =   14
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chk_glyhb 
            Caption         =   "Glycosylated Hb"
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
            TabIndex        =   13
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chk_glucotolerancetest 
            Caption         =   "Glucose Tolerance Test"
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
            Left            =   2520
            TabIndex        =   12
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame LP 
         Caption         =   "LIPID PROFILE"
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   6600
         TabIndex        =   64
         Top             =   360
         Width           =   4935
         Begin VB.CheckBox chk_cholest 
            Caption         =   "Cholesterol"
            BeginProperty Font 
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
            TabIndex        =   69
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chk_hdlcholest 
            Caption         =   "HDL Cholesterol"
            BeginProperty Font 
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
            Left            =   2040
            TabIndex        =   68
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chk_trigly 
            Caption         =   "Triglycerides"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   67
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chk_lippro 
            Caption         =   "Lipid  Profile"
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
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   1815
         End
         Begin VB.CheckBox chk_ldlcholest 
            Caption         =   "LDL Cholesterol"
            BeginProperty Font 
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
            Left            =   2040
            TabIndex        =   65
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame RFT 
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   120
         TabIndex        =   60
         Top             =   1560
         Width           =   5055
         Begin VB.CheckBox chk_creat 
            Caption         =   "Creatinine"
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
            Left            =   120
            TabIndex        =   168
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chk_renfuntest 
            Caption         =   "Renal Function Tests"
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
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   2655
         End
         Begin VB.CheckBox chk_urea 
            Caption         =   "Urea"
            BeginProperty Font 
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
            TabIndex        =   62
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chk_uricacid 
            Caption         =   "Uric Acid"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   61
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame ELECTROLYTES 
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   5160
         TabIndex        =   54
         Top             =   1560
         Width           =   2895
         Begin VB.CheckBox chk_elect 
            Caption         =   "Electrolytes"
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
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   1575
         End
         Begin VB.CheckBox chk_sodium 
            Caption         =   "Sodium"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   58
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chk_pota 
            Caption         =   "Potassium"
            BeginProperty Font 
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
            TabIndex        =   57
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chk_chlor 
            Caption         =   "Chlorides"
            BeginProperty Font 
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
            Left            =   1560
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox CHK_CO2 
            Caption         =   "Carbon dioxide"
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
            Left            =   1560
            TabIndex        =   55
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame FRM_CALCIUM 
         Height          =   1095
         Left            =   8160
         TabIndex        =   49
         Top             =   1560
         Width           =   3255
         Begin VB.CheckBox chk_calc 
            Caption         =   "Calcium"
            BeginProperty Font 
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
            TabIndex        =   53
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chk_ioncal 
            Caption         =   "Ionic Calcium"
            BeginProperty Font 
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
            TabIndex        =   52
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chk_phosp 
            Caption         =   "Phosphorus"
            BeginProperty Font 
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
            Left            =   1800
            TabIndex        =   51
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox chk_magn 
            Caption         =   "Magnesium"
            BeginProperty Font 
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
            Left            =   1800
            TabIndex        =   50
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame CE 
         Height          =   1095
         Left            =   7800
         TabIndex        =   43
         Top             =   4080
         Width           =   3615
         Begin VB.CheckBox chk_carenzy 
            Caption         =   "Cardiac - Enzymes"
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
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox chk_ldh 
            Caption         =   "LDH"
            BeginProperty Font 
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
            TabIndex        =   47
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chk_ldh1 
            Caption         =   "S.G.O.T"
            BeginProperty Font 
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
            Left            =   1920
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox CHK_cpk 
            Caption         =   "CPK"
            BeginProperty Font 
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
            TabIndex        =   45
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chk_cpkmb 
            Caption         =   "CPK-MB"
            BeginProperty Font 
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
            Left            =   1920
            TabIndex        =   44
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame PE 
         Height          =   1095
         Left            =   0
         TabIndex        =   39
         Top             =   4080
         Width           =   2895
         Begin VB.CheckBox chk_panenzy 
            Caption         =   "Pancreatic - Enzymes"
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
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   2655
         End
         Begin VB.CheckBox chk_amylase 
            Caption         =   "Amylase"
            BeginProperty Font 
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
            TabIndex        =   41
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chk_lipase 
            Caption         =   "Lipase"
            BeginProperty Font 
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
            TabIndex        =   40
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame FRM_IRON 
         Height          =   1095
         Left            =   3120
         TabIndex        =   36
         Top             =   4080
         Width           =   1815
         Begin VB.CheckBox chk_iron 
            Caption         =   "Iron"
            BeginProperty Font 
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
            TabIndex        =   38
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chk_tibc 
            Caption         =   "T.I.B.C"
            BeginProperty Font 
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
            TabIndex        =   37
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame FRM_ACID 
         Height          =   1095
         Left            =   5040
         TabIndex        =   33
         Top             =   4080
         Width           =   2655
         Begin VB.CheckBox chk_acidphosp 
            Caption         =   "Acid Phosph"
            BeginProperty Font 
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
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chk_prostacid 
            Caption         =   "Prost.Acid Phosph"
            BeginProperty Font 
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
            TabIndex        =   34
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame SA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8520
         TabIndex        =   31
         Top             =   5160
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CheckBox Check63 
            Caption         =   "Stone Analysis"
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
            Height          =   240
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.CheckBox chk_routinebio 
         BackColor       =   &H00404040&
         Caption         =   "Routine Biochemistry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   6480
         TabIndex        =   30
         Top             =   0
         Width           =   2655
      End
      Begin VB.Frame LFT 
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   0
         TabIndex        =   19
         Top             =   2760
         Width           =   11415
         Begin VB.CheckBox chk_livfuntest 
            Caption         =   "Liver Function Tests"
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
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   2535
         End
         Begin VB.CheckBox chk_prot 
            Caption         =   "Proteins"
            BeginProperty Font 
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
            TabIndex        =   28
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chk_albumin 
            Caption         =   "Albumin"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk_glob 
            Caption         =   "Globulins"
            BeginProperty Font 
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
            Left            =   5520
            TabIndex        =   26
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk_alkphos 
            Caption         =   "Alk.Phosphatase"
            BeginProperty Font 
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
            Left            =   7200
            TabIndex        =   25
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chk_totbil 
            Caption         =   "Total Bilirubin"
            BeginProperty Font 
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
            Left            =   9480
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chk_dirbil 
            Caption         =   "Direct Bilirubin"
            BeginProperty Font 
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
            TabIndex        =   23
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox chk_sgot 
            Caption         =   "S.G.O.T"
            BeginProperty Font 
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
            Left            =   5520
            TabIndex        =   22
            Top             =   900
            Width           =   1215
         End
         Begin VB.CheckBox chk_sgpt 
            Caption         =   "S.G.P.T"
            BeginProperty Font 
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
            Left            =   7200
            TabIndex        =   21
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox chk_indirbil 
            Caption         =   "Indirect Bilirubin"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   20
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame17 
         Height          =   855
         Left            =   120
         TabIndex        =   157
         Top             =   5280
         Width           =   1455
      End
   End
   Begin VB.Frame bodyfluid 
      Height          =   7695
      Left            =   120
      TabIndex        =   70
      Top             =   1560
      Visible         =   0   'False
      Width           =   11655
      Begin VB.ComboBox CMB_BODYF 
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
         Left            =   1920
         TabIndex        =   104
         Top             =   360
         Width           =   2295
      End
      Begin VB.Frame Frame11 
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
         Height          =   1095
         Left            =   5760
         TabIndex        =   100
         Top             =   4920
         Width           =   3615
         Begin VB.CheckBox CHK_LDHBF 
            Caption         =   "LDH"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   103
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox CHK_CALCI 
            Caption         =   "Calcium"
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
            Left            =   1680
            TabIndex        =   102
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox CHK_PHOS 
            Caption         =   "Phosphorus"
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
            Left            =   1680
            TabIndex        =   101
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame10 
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
         Height          =   1095
         Left            =   480
         TabIndex        =   97
         Top             =   4920
         Width           =   3975
         Begin VB.CheckBox CHK_AMYBF 
            Caption         =   "Amylase"
            BeginProperty Font 
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
            TabIndex        =   99
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox CHK_LIPBF 
            Caption         =   "Lipase"
            BeginProperty Font 
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
            TabIndex        =   98
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   480
         TabIndex        =   86
         Top             =   3600
         Width           =   11055
         Begin VB.CheckBox Check34 
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
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox CHK_PROBF 
            Caption         =   "Proteins"
            BeginProperty Font 
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
            TabIndex        =   95
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox CHK_ALBBF 
            Caption         =   "Albumin"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   94
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox CHK_GLOBF 
            Caption         =   "Globulins"
            BeginProperty Font 
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
            Left            =   5040
            TabIndex        =   93
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox CHK_ALKBF 
            Caption         =   "Alk.Phosphatase"
            BeginProperty Font 
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
            Left            =   6960
            TabIndex        =   92
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox CHK_TOTALBILBF 
            Caption         =   "Total Bilirubin"
            BeginProperty Font 
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
            Left            =   9120
            TabIndex        =   91
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox CHK_DIRECTBILBF 
            Caption         =   "Direct Bilirubin"
            BeginProperty Font 
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
            TabIndex        =   90
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox CHK_SGOTBF 
            Caption         =   "S.G.O.T"
            BeginProperty Font 
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
            TabIndex        =   89
            Top             =   900
            Width           =   1215
         End
         Begin VB.CheckBox CHK_SGPTBF 
            Caption         =   "S.G.P.T"
            BeginProperty Font 
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
            Left            =   6960
            TabIndex        =   88
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox CHK_IBBF 
            Caption         =   "Indirect Bilirubin"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   87
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
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
         Height          =   1095
         Left            =   5760
         TabIndex        =   82
         Top             =   2400
         Width           =   4095
         Begin VB.CheckBox CHK_NABF 
            Caption         =   "Sodium"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   85
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox CHK_KBF 
            Caption         =   "Potassium"
            BeginProperty Font 
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
            TabIndex        =   84
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox CHK_CLBF 
            Caption         =   "Chlorides"
            BeginProperty Font 
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
            Left            =   2040
            TabIndex        =   83
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
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
         Height          =   1095
         Left            =   480
         TabIndex        =   78
         Top             =   2400
         Width           =   3975
         Begin VB.CheckBox CHK_UREABF 
            Caption         =   "Urea"
            BeginProperty Font 
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
            TabIndex        =   81
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox CHK_URICACIDBF 
            Caption         =   "Uric Acid"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   80
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox CHK_CREATININEBF 
            Caption         =   "Creatinine"
            BeginProperty Font 
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
            TabIndex        =   79
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   1455
         Left            =   5760
         TabIndex        =   73
         Top             =   840
         Width           =   4095
         Begin VB.CheckBox CHK_CHOLBF 
            Caption         =   "Cholesterol"
            BeginProperty Font 
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
            TabIndex        =   77
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox CHK_HDLBF 
            Caption         =   "HDL Cholesterol"
            BeginProperty Font 
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
            Left            =   2040
            TabIndex        =   76
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox CHK_TRIBF 
            Caption         =   "Triglycerides"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   75
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox CHK_LDLBF 
            Caption         =   "LDL Cholesterol"
            BeginProperty Font 
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
            Left            =   2040
            TabIndex        =   74
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   1455
         Left            =   480
         TabIndex        =   71
         Top             =   840
         Width           =   3975
         Begin VB.CheckBox CHK_SUGARBF 
            Caption         =   "Sugar"
            BeginProperty Font 
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
            TabIndex        =   72
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Body Fluid"
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
         Left            =   600
         TabIndex        =   105
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame STONEANALYSIS 
      Height          =   3855
      Left            =   480
      TabIndex        =   106
      Top             =   1920
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CheckBox chk_stoneanalysis 
         Caption         =   "Stone Analysis"
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
         TabIndex        =   108
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txt_stoneanalysis 
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
         Height          =   1575
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   107
         Top             =   2400
         Visible         =   0   'False
         Width           =   9135
      End
      Begin VB.Label Label4 
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
         Height          =   495
         Left            =   240
         TabIndex        =   109
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame FRMGTT 
      Height          =   615
      Left            =   4800
      TabIndex        =   154
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox chk_gtt 
         Caption         =   "GTT"
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
         Left            =   360
         TabIndex        =   155
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FRAMECCT 
      Height          =   1815
      Left            =   4080
      TabIndex        =   149
      Top             =   4440
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CheckBox GROUPCCT 
         Caption         =   "CCT"
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
         Left            =   0
         TabIndex        =   153
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox CHK_CCC 
         Caption         =   "Corrected Creat. Clear"
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
         Left            =   240
         TabIndex        =   152
         Top             =   240
         Width           =   3615
      End
      Begin VB.CheckBox CHK_SC 
         Caption         =   "Serum Creatinine"
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
         Left            =   240
         TabIndex        =   151
         Top             =   720
         Width           =   3495
      End
      Begin VB.CheckBox CHK_UC 
         Caption         =   "Urine Creatinine "
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
         Left            =   240
         TabIndex        =   150
         Top             =   1200
         Width           =   3495
      End
   End
   Begin VB.Frame BFI 
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
      Height          =   1215
      Left            =   4800
      TabIndex        =   144
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chk_csfldh 
         Caption         =   " LDH"
         BeginProperty Font 
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
         Left            =   2520
         TabIndex        =   148
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox chk_csfchlor 
         Caption         =   "Chlorides"
         BeginProperty Font 
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
         TabIndex        =   147
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chk_csfprot 
         Caption         =   "Proteins"
         BeginProperty Font 
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
         Left            =   2520
         TabIndex        =   146
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chk_csfsug 
         Caption         =   "Sugar"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   145
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<<&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   8
      Top             =   9360
      Width           =   1215
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5760
      TabIndex        =   1
      Top             =   720
      Width           =   5655
   End
   Begin VB.ComboBox CMB_NMETERIAL 
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
      ItemData        =   "BIOCHEMISTRY1.frx":0033
      Left            =   1680
      List            =   "BIOCHEMISTRY1.frx":0035
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      ForeColor       =   16711680
      Enabled         =   0   'False
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
   Begin VB.Frame urine24hour 
      Height          =   6615
      Left            =   120
      TabIndex        =   110
      Top             =   1440
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame Frame12 
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
         Height          =   1095
         Left            =   240
         TabIndex        =   122
         Top             =   600
         Width           =   5055
         Begin VB.CheckBox chk_corrcreatclr24 
            Caption         =   "Corrected Creat. Clear"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   126
            Top             =   720
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox chk_creat24 
            Caption         =   "Creatinine"
            BeginProperty Font 
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
            Left            =   600
            TabIndex        =   125
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chk_uricacid24 
            Caption         =   "Uric Acid"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   124
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chk_urea24 
            Caption         =   "Urea"
            BeginProperty Font 
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
            Left            =   600
            TabIndex        =   123
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame13 
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
         Height          =   1095
         Left            =   5760
         TabIndex        =   118
         Top             =   600
         Width           =   3975
         Begin VB.CheckBox chk_chlor24 
            Caption         =   "Chlorides"
            BeginProperty Font 
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
            Left            =   2400
            TabIndex        =   121
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk_pota24 
            Caption         =   "Potassium"
            BeginProperty Font 
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
            TabIndex        =   120
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chk_sodium24 
            Caption         =   "Sodium"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   119
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1215
         Left            =   240
         TabIndex        =   115
         Top             =   1920
         Width           =   5055
         Begin VB.CheckBox chk_um24hrs 
            Caption         =   "URINARY MAGNESIUM (24 HOURS)"
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
            Left            =   600
            TabIndex        =   169
            Top             =   720
            Width           =   3735
         End
         Begin VB.CheckBox chk_phosp24 
            Caption         =   "Phosphorus"
            BeginProperty Font 
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
            Left            =   2520
            TabIndex        =   117
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox chk_calc24 
            Caption         =   "Calcium"
            BeginProperty Font 
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
            Left            =   600
            TabIndex        =   116
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame15 
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
         Height          =   1215
         Left            =   5880
         TabIndex        =   111
         Top             =   1920
         Width           =   3975
         Begin VB.CheckBox chk_urin5hiaa 
            Caption         =   "Urinary 5HIAA"
            BeginProperty Font 
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
            Left            =   1680
            TabIndex        =   114
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox chk_urinvma 
            Caption         =   "Urinary VMA"
            BeginProperty Font 
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
            Left            =   1680
            TabIndex        =   113
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chk_prot24 
            Caption         =   "Proteins"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   112
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin VB.Frame frame_Urine 
      Caption         =   "Urine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   2055
      Left            =   4440
      TabIndex        =   160
      Top             =   6360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox chk_ppurinsug 
         Caption         =   "P.P. Urine Sugar"
         BeginProperty Font 
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
         Left            =   720
         TabIndex        =   162
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chk_furinsug 
         Caption         =   "F Urine Sugar"
         BeginProperty Font 
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
         Left            =   720
         TabIndex        =   161
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label5 
      Caption         =   "05/11/21"
      Height          =   255
      Left            =   240
      TabIndex        =   171
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbl_msg 
      Caption         =   "Select  Eleven Procedures  only...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   9600
      Width           =   7575
   End
   Begin VB.Label lbl_specimen 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label lbl_caseno 
      Caption         =   "Case No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "BIOCHEMISTRY LABORATORY SERVICES"
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
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "BIOCHEMISTRY1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ccount As Integer
Dim chk As Boolean
Private Sub chk_bodyfludinv_Click()
If chk_bodyfludinv.Value = 1 Then
chk_csfsug.Value = 1
'    chk_csfprot.Value = 1
'    chk_amylase1.Value = 1
'    chk_csfchlor.Value = 1
'    chk_csfldh.Value = 1
'    chk_lipase1.Value = 1
End If
End Sub



Private Sub chk_albumin_Click()
If chk_albumin.Value = 0 Then
    'chk_prot.Value = 0
    chk_livfuntest.Value = 0
End If
If chk_albumin.Value = 1 Then
    chk_prot.Value = 1
End If
If chk_prot.Value = 1 And chk_albumin.Value = 1 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 1
End If
If chk_prot.Value = 0 Or chk_albumin.Value = 0 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 0
End If
End Sub

Private Sub chk_alkphos_Click()
If chk_alkphos.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub





Private Sub chk_carenzy_Click()
'If chk_carenzy.Value = 1 Then
'    chk_LDH.Value = 1
'    chk_ldh1.Value = 1
'    chk_CPK.Value = 1
'    chk_CPKMB.Value = 1
'Else
'    chk_LDH.Value = 0
'    chk_ldh1.Value = 0
'    chk_CPK.Value = 0
'    chk_CPKMB.Value = 0
'End If
End Sub

Private Sub chk_corrcreatclr_Click()
If chk_corrcreatclr.Value = 1 Then
    chk_creat.Value = 1
Else
    chk_creat.Value = 0
End If
End Sub

Private Sub chk_creat_Click()
If chk_creat.Value = 0 Then
    chk_renfuntest.Value = 0
End If
End Sub

Private Sub chk_creatur_Click()
If chk_creatur.Value = 0 Then
    'chk_renfuntestur.Value = 0
End If
End Sub

Private Sub chk_dirbil_Click()
 If chk_dirbil.Value = 1 And chk_indirbil.Value = 1 Then
    chk_totbil.Value = 1
 Else
    chk_totbil.Value = 0
 End If
End Sub

Private Sub chk_elect_Click()
If chk_elect.Value = 1 Then
    chk_sodium.Value = 1
    chk_chlor.Value = 1
    chk_pota.Value = 1
    CHK_CO2.Value = 1
Else
    chk_sodium.Value = 0
    chk_chlor.Value = 0
    chk_pota.Value = 0
    CHK_CO2.Value = 0
End If
End Sub



Private Sub chk_ggt_Click()
    'chk_ggt.Value = 1
End Sub

Private Sub chk_glob_Click()
If chk_glob.Value = 0 Then
    chk_livfuntest.Value = 0
End If
If chk_prot.Value = 1 And chk_albumin.Value = 1 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 1
End If
If chk_prot.Value = 0 Or chk_albumin.Value = 0 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 0
End If
End Sub




Private Sub chk_indirbil_Click()
   If chk_dirbil.Value = 1 And chk_indirbil.Value = 1 Then
    chk_totbil.Value = 1
   Else
    chk_totbil.Value = 0
   End If
End Sub



Private Sub chk_lippro_Click()
If chk_lippro.Value = 1 Then

    chk_cholest.Value = 1
    chk_hdlcholest.Value = 1
    chk_trigly.Value = 1
    chk_ldlcholest.Value = 1
Else
    chk_cholest.Value = 0
    chk_hdlcholest.Value = 0
    chk_trigly.Value = 0
    chk_ldlcholest.Value = 0
End If
End Sub
Private Sub chk_livfuntest_Click()
If chk_livfuntest.Value = 1 Then
    chk_prot.Value = 1
    chk_albumin.Value = 1
    chk_glob.Value = 1
    chk_alkphos.Value = 1
    chk_totbil.Value = 1
    chk_sgot.Value = 1
    chk_sgpt.Value = 1
Else
    chk_prot.Value = 0
    chk_albumin.Value = 0
    chk_glob.Value = 0
    chk_alkphos.Value = 0
    chk_totbil.Value = 0
    chk_sgot.Value = 0
    chk_sgpt.Value = 0

'    If chk_prot.Value = 0 Then
'        chk_prot.Value = 0
'     End If
'     If chk_albumin.Value = 0 Then
'        chk_albumin.Value = 0
'    End If
'    If chk_glob.Value = 0 Then
'        chk_glob.Value = 0
'    End If
'    If chk_alkphos.Value = 0 Then
'    chk_alkphos.Value = 0
'    End If
'    If chk_totbil.Value = 0 Then
'    chk_totbil.Value = 0
'    End If
'    If chk_SGOT.Value = 0 Then
'    chk_SGOT.Value = 0
'    End If
'    If chk_sgpt.Value = 0 Then
'    chk_sgpt.Value = 0
'    End If
End If
End Sub

Private Sub chk_panenzy_Click()
If chk_panenzy.Value = 1 Then
    chk_amylase.Value = 1
    chk_lipase.Value = 1
Else
    chk_amylase.Value = 0
    chk_lipase.Value = 0
End If
End Sub

Private Sub chk_ppsug_Click()
'Added by Shraddha on 09/09/11 by Shraddha
If chk_ppsug.Value = 1 Then
 
    For Each obj In BIOCHEMISTRY1.Controls
       If TypeOf obj Is CheckBox And obj.name <> "chk_ppsug" Then
           If obj.Value = 1 Then
               obj.Value = False
           End If
          obj.Enabled = False
       End If
    Next
    
    chk_ppsug.Enabled = True
 
Else
    For Each obj In BIOCHEMISTRY1.Controls
       If TypeOf obj Is CheckBox Then
          obj.Enabled = True
       End If
    Next
End If
 
 'chk_ppsug.Value = 1
End Sub

Private Sub chk_prot_Click()
If chk_prot.Value = 0 Then
    chk_albumin.Value = 0
    chk_livfuntest.Value = 0
End If
If chk_prot.Value = 1 And chk_albumin.Value = 1 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 1
End If
If chk_prot.Value = 0 Or chk_albumin.Value = 0 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 0
End If

End Sub

Private Sub chk_ransug_Click()
'Added by Shraddha on 09/09/11 by Shraddha
If chk_ransug.Value = 1 Then
 
    For Each obj In BIOCHEMISTRY1.Controls
       If TypeOf obj Is CheckBox And obj.name <> "chk_ransug" Then
           If obj.Value = 1 Then
               obj.Value = False
           End If
          obj.Enabled = False
       End If
    Next
    
    chk_ransug.Enabled = True
 
Else
    For Each obj In BIOCHEMISTRY1.Controls
       If TypeOf obj Is CheckBox Then
          obj.Enabled = True
       End If
    Next
End If
End Sub

Private Sub chk_renfuntest_Click()
If chk_renfuntest.Value = 1 Then
    chk_urea.Value = 1
    chk_uricacid.Value = 1
    chk_creat.Value = 1
    
Else
    chk_urea.Value = 0
    chk_uricacid.Value = 0
    chk_creat.Value = 0
    
'    If chk_urea.Value = 0 Then
'        chk_urea.Value = 0
'
'    End If
'    If chk_uricacid.Value = 0 Then
'        chk_uricacid.Value = 0
'
'    End If
'    If chk_creat.Value = 0 Then
'        chk_creat.Value = 0
'
'    End If
End If
End Sub



Private Sub chk_urininv_Click()
If chk_urininv.Value = 1 Then
'chk_urinvma.Value = 1
'chk_urob.Value = 1
'chk_urin5hiaa.Value = 1
'chk_Porphyrin.Value = 1
End If
End Sub




Private Sub chk_routinebio_Click()
If chk_routinebio.Value = 1 Then
    chk_trusugfas.Value = 1
    chk_renfuntest.Value = 1
    chk_livfuntest.Value = 1
Else
    chk_trusugfas.Value = 0
    chk_renfuntest.Value = 0
    chk_livfuntest.Value = 0
End If
End Sub

Private Sub chk_SGOT_Click()
If chk_sgot.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub

Private Sub chk_sgpt_Click()
If chk_sgpt.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub

Private Sub chk_stoneanalysis_Click()
    If chk_stoneanalysis.Value = 1 Then
        txt_stoneanalysis.Text = ""
        txt_stoneanalysis.Visible = True
    End If
    If chk_stoneanalysis.Value = 0 Then
        txt_stoneanalysis.Visible = False
    End If
End Sub

Private Sub chk_totbil_Click()
If chk_totbil.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub



Private Sub chk_urea_Click()
If chk_urea.Value = 0 Then
chk_renfuntest.Value = 0
End If
End Sub

Private Sub chk_ureaur_Click()
If chk_ureaur.Value = 0 Then
    chk_renfuntestur.Value = 0
End If
End Sub

Private Sub chk_uricacid_Click()
If chk_uricacid.Value = 0 Then
    chk_renfuntest.Value = 0
End If
End Sub

Private Sub chk_uricacidur_Click()
If chk_uricacidur.Value = 0 Then
    chk_renfuntestur.Value = 0
End If
End Sub

Private Sub CMB_BODYF_GotFocus()
CMB_BODYF.clear
CMB_BODYF.AddItem "Ascitic"
CMB_BODYF.AddItem "Pleural"
CMB_BODYF.AddItem "Drain"
CMB_BODYF.AddItem "Pericardial"
End Sub

Private Sub CMB_BODYF_LostFocus()
bodyf = CMB_BODYF.Text
End Sub

Private Sub CMB_NMETERIAL_Click()
If Trim(CMB_NMETERIAL.Text) = "CSF" Then
    blood.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    BFI.Visible = True
    FRAMECCT.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
ElseIf Trim(CMB_NMETERIAL.Text) = "Blood" Then
    blood.Visible = True
    BFI.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    FRAMECCT.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
'****** Added by Shraddha Pandit on 09/03/09 ******
ElseIf Trim(CMB_NMETERIAL.Text) = "Urine" Then
    frame_Urine.Visible = True
    blood.Visible = False
    BFI.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    FRAMECCT.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
'**************************************************
ElseIf Trim(CMB_NMETERIAL.Text) = "Stone Analysis" Then
    blood.Visible = False
    BFI.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = True
    FRAMECCT.Visible = False
    FRMGTT.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
    'txt_stoneanalysis.SetFocus
ElseIf Trim(CMB_NMETERIAL.Text) = "Body Fluid" Then
    blood.Visible = False
    BFI.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = True
    STONEANALYSIS.Visible = False
    FRAMECCT.Visible = False
    FRMGTT.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
    CMB_BODYF.SetFocus
ElseIf Trim(CMB_NMETERIAL.Text) = "Random Urine" Then
    blood.Visible = False
    BFI.Visible = False
    urinerandom.Visible = True
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    FRAMECCT.Visible = False
    FRMGTT.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
ElseIf Trim(CMB_NMETERIAL.Text) = "24 hours Urine" Then
    blood.Visible = False
    BFI.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = True
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    FRAMECCT.Visible = False
    FRMGTT.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
ElseIf Trim(CMB_NMETERIAL.Text) = "24 Hours CCT(blood+urine)" Then
    blood.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    BFI.Visible = False
    FRMGTT.Visible = False
    FRAMECCT.Visible = True
    frame_Urine.Visible = False
    frame_EDL.Visible = False 'Shraddha 31/10/2011
ElseIf Trim(CMB_NMETERIAL.Text) = "GTT" Then
    blood.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    BFI.Visible = False
    FRAMECCT.Visible = False
    FRMGTT.Visible = True
    frame_EDL.Visible = False 'Shraddha 31/10/2011
ElseIf Trim(CMB_NMETERIAL.Text) = "Estimated drug dosage" Then
    blood.Visible = False
    urinerandom.Visible = False
    urine24hour.Visible = False
    bodyfluid.Visible = False
    STONEANALYSIS.Visible = False
    BFI.Visible = False
    FRAMECCT.Visible = False
    FRMGTT.Visible = False
    frame_EDL.Visible = True
End If
End Sub

Private Sub cmb_nmeterial_GotFocus()
CMB_NMETERIAL.clear
CMB_NMETERIAL.AddItem "CSF"
CMB_NMETERIAL.AddItem "Body Fluid"
CMB_NMETERIAL.AddItem "Blood"
CMB_NMETERIAL.AddItem "Urine" 'Shraddha 09/03/09
CMB_NMETERIAL.AddItem "Random Urine"
CMB_NMETERIAL.AddItem "24 hours Urine"
''''ADDED BY SHWETA 20102010
CMB_NMETERIAL.AddItem "24 Hours CCT(blood+urine)"
'*******************************
'Added on 31/10/2011 by Shraddha
CMB_NMETERIAL.AddItem "Estimated drug dosage"

'***** Commented by Shraddha on 24/06/10 *****
'CMB_NMETERIAL.AddItem "24 Hours CCT(blood+urine)"
'CMB_NMETERIAL.AddItem "Stone Analysis"
'CMB_NMETERIAL.AddItem "GTT"
'*********************************************
End Sub

Private Sub CMB_NMETERIAL_LostFocus()
meterial = CMB_NMETERIAL.Text
End Sub

Private Sub cmd_back_Click()
meterial = CMB_NMETERIAL.Text
pathology.CMB_TEXT.clear
ccount = 0
Dim drug_dose As String
drug_dose = ""
DDG = ""

'***** Shraddha 1/12/2011 *****
chk = False
For Each c In Controls
If TypeOf c Is CheckBox Then
    If c.Value = 1 Then
        chk = True
    End If
End If
Next
If pathology.cmb_reqno.Visible = False Then
If chk = False Then
    MsgBox "Please select atleast 1 procedure"
    Exit Sub
End If
End If
'******************************

If BCFLAG = False Then
'VENKAT
'mono
'sundeep
    If Trim(CMB_NMETERIAL.Text) = "CSF" Then
        CSF
    End If
    If Trim(CMB_NMETERIAL.Text) = "Body Fluid" Then
        If Trim(CMB_BODYF.Text) = "" Then
            MsgBox "Please Enter Body Fluid !!", vbInformation, "Biochemistry"
            CMB_BODYF.SetFocus
            Exit Sub
        End If
        BODYFLUIDTEST
    End If
    'ADD  FOR CCT
    If Trim(CMB_NMETERIAL.Text) = "24 Hours CCT(blood+urine)" Then
        CCT
    End If
    ' END
    If Trim(CMB_NMETERIAL.Text) = "Blood" Then
        BLOODTEST
    '************* Shraddha Pandit 09/03/09 ********
    ElseIf Trim(CMB_NMETERIAL.Text) = "Urine" Then
        If chk_furinsug.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F846"
        End If
        If chk_ppurinsug.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F847"
        End If
    '***********************************************
    ElseIf Trim(CMB_NMETERIAL.Text) = "Random Urine" Then
         '**** Renal Function Tests
        RANDOMURINETEST
    ElseIf Trim(CMB_NMETERIAL.Text) = "Stone Analysis" Then
        ''STONE ANALYSIS
        If chk_stoneanalysis.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F947"
            ccount = ccount + 1
        End If
        ''STONE ANALYSIS
    ElseIf Trim(CMB_NMETERIAL.Text) = "24 hours Urine" Then
      urine24hourS
    ElseIf Trim(CMB_NMETERIAL.Text) = "GTT" Then
        ''GTT
        If chk_gtt.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F844"
            ccount = ccount + 1
        End If
        ''GTT
        If pathology.CMB_TEXT.ListCount = 0 Then
        MsgBox ("Please atleast 1 procedure code"), vbInformation
        Exit Sub
    End If
    '***** Shraddha 31/10/2011 *****
    ElseIf Trim(CMB_NMETERIAL.Text) = "Estimated drug dosage" Then
        
        If chk_metho.Value = 1 Then
      
            If Trim(txt_amt.Text) <> "" And Trim(txt_hrs.Text) <> "" Then
      
                drug_dose = Trim(txt_amt.Text) & " XXX " & Trim(txt_hrs.Text)
    
                If Trim(CMB_FR.Text) = "PM" Then
                    If Trim(TXT_HH.Text) <> "12" Then
                         TXT_HH.Text = 12 + CInt(TXT_HH.Text)
                    End If
                End If
    
                If Trim(CMB_FR.Text) = "AM" Then
                    If Trim(TXT_HH.Text) = "12" And Trim(TXT_MM.Text) = "00" Then
                        TXT_HH.Text = 12 + CInt(TXT_HH.Text)
                    End If
                    
                    If Trim(TXT_HH.Text) = "12" And Trim(TXT_MM.Text) <> "00" Then
                        TXT_HH.Text = "00"
                    End If
                End If
    
                Dim Yr, mn, DY As String
                Yr = Right(msk_fromdt.FormattedText, 4)
                mn = Mid(msk_fromdt.FormattedText, 4, 2)
                DY = Left(msk_fromdt.FormattedText, 2)
    
                datereq = Yr & "/" & mn & "/" & DY
                
                timereq = Trim(TXT_HH.Text) & ":" & Trim(TXT_MM.Text) & ":" & "00"
                
                If chk_metho.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox " You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F836"
                    ccount = ccount + 1
                    
                    'if drug_dose
                    DDG = drug_dose
                    AM = txt_rem
                    'Unload Me
                    'pathology.show
                End If
                
            Else
                drug_dose = ""
                DDG = drug_dose
            End If
            
        Else
            MsgBox "Please check the option"
        End If
       
       End If
       
    End If
    
    
    
    
    


    
   
    

    BCFLAG = False
'End If
'ADDED BY AMIT
'If chk_csfsug.Value = 1 Or chk_csfchlor.Value = 1 Or chk_csfprot.Value = 1 Or chk_csfldh.Value = 1 Then
 '   CSF_FLAG = True
'End If
Unload Me
pathology.show
'pathology.Biochemistry1.Visible = True
End Sub
Private Sub Form_Load()    '1

ccount = 0
meterial = ""
bodyf = ""
CMB_NMETERIAL.clear
CMB_NMETERIAL.AddItem "CSF"
CMB_NMETERIAL.AddItem "Body Fluid"
CMB_NMETERIAL.AddItem "Blood"
CMB_NMETERIAL.AddItem "Urine" 'Shraddha 09/03/09
CMB_NMETERIAL.AddItem "Random Urine"
CMB_NMETERIAL.AddItem "24 hours Urine"
'***** Commented by Shraddha on 24/06/10 *****
'CMB_NMETERIAL.AddItem "24 Hours CCT(blood+urine)"
'CMB_NMETERIAL.AddItem "Stone Analysis"
'CMB_NMETERIAL.AddItem "GTT"
'*********************************************
CMB_NMETERIAL.AddItem "Estimated drug dosage"
End Sub
Private Sub txt_nmeterial_LostFocus()
meterial = Trim(txt_nmeterial.Text)
End Sub
'ADD CSF
Public Sub CSF()
If chk_csfsug.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F884"
End If
If chk_csfprot.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F885"
End If
If chk_csfchlor.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F886"
End If
If chk_csfldh.Value = 1 Then
    pathology.CMB_TEXT.AddItem "F887"
End If
'Unload Me
'pathology.show
End Sub
' ADD  FOR CCT
Sub CCT()
            If CHK_CCC.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                ccount = ccount + 1
                pathology.CMB_TEXT.AddItem "F924"
            Else
            If CHK_SC.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                ccount = ccount + 1
                pathology.CMB_TEXT.AddItem "F974"
            End If
            If CHK_UC.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                ccount = ccount + 1
                pathology.CMB_TEXT.AddItem "F920"
            End If
            End If
End Sub
Sub BLOODTEST()
If chk_routinebio.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F802"
        Else
            If chk_trusugfas.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                ccount = ccount + 1
                pathology.CMB_TEXT.AddItem "F842"
            End If
            '**** Renal Function Tests
            If chk_renfuntest.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F854"
                ccount = ccount + 1
                If chk_corrcreatclr24.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F859"
                    ccount = ccount + 1
                End If
            Else
                If chk_urea.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F855"
                    ccount = ccount + 1
                End If
                If chk_uricacid.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F856"
                    ccount = ccount + 1
                End If
                If chk_creat.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F857"
                    ccount = ccount + 1
                End If
'                If chk_corrcreatclr.Value = 1 Then
'                    If ccount >= 11 Then
'                        MsgBox " You Can Select Max. of 11 Procedures only."
'                        Unload Me
'                        pathology.show
'                        Exit Sub
'                    End If
'                    pathology.CMB_TEXT.AddItem "F859"
'                    ccount = ccount + 1
'                End If
          End If
'            '**** Renal Function Tests
                '**** Liver function Tests
            If chk_livfuntest.Value = 1 Then
                If ccount > 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F865"
                ccount = ccount + 1
                If chk_dirbil.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F871"
                    ccount = ccount + 1
                End If
                If chk_indirbil.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F872"
                    ccount = ccount + 1
                End If
            Else
                If chk_prot.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F866"
                    ccount = ccount + 1
                End If
                If chk_albumin.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F867"
                    ccount = ccount + 1
                End If
                If chk_glob.Value = 1 Then
                    If ccount > 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F868"
                    ccount = ccount + 1
                End If
                If chk_alkphos.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F869"
                    ccount = ccount + 1
                End If
                If chk_totbil.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F870"
                    ccount = ccount + 1
                End If
                If chk_dirbil.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F871"
                    ccount = ccount + 1
                End If
                If chk_indirbil.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F872"
                    ccount = ccount + 1
                End If
                If chk_sgot.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F873"
                    ccount = ccount + 1
                End If
                If chk_sgpt.Value = 1 Then
                    If ccount >= 11 Then
                        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F874"
                    ccount = ccount + 1
                End If
            End If
            '**** Liver function Tests
        End If
        'venkat
        If chk_ppsug.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F843"
            ccount = ccount + 1
        End If
        If chk_ransug.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F841"
            ccount = ccount + 1
        End If
        If chk_glyhb.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F845"
            ccount = ccount + 1
        End If
'        If chk_furinsug.Value = 1 Then
'            If ccount >= 11 Then
'                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            ccount = ccount + 1
'            pathology.CMB_TEXT.AddItem "F846"
'        End If
'        If chk_ppurinsug.Value = 1 Then
'            If ccount >= 11 Then
'                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            ccount = ccount + 1
'            pathology.CMB_TEXT.AddItem "F847"
'        End If
        If chk_glucotolerancetest.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            '''ADDED AND COMMENTED BY CHARULATA ON 14/09/2010 FOR NEW SER_CODE
            'pathology.CMB_TEXT.AddItem "F810"
            pathology.CMB_TEXT.AddItem "F844"
        End If
        '***** LIQUID PROFILE
        If chk_lippro.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F849"
            ccount = ccount + 1
        Else
            If chk_cholest.Value = 1 Then
                If ccount > 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F850"
                ccount = ccount + 1
            End If
            If chk_hdlcholest.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                ccount = ccount + 1
                pathology.CMB_TEXT.AddItem "F851"
            End If
            If chk_trigly.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F853"
                ccount = ccount + 1
            End If
            If chk_ldlcholest.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F852"
                ccount = ccount + 1
            End If
        End If
        '''**** Renal Function Tests
        '**** Electrolytes
        If chk_elect.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F860"
            ccount = ccount + 1
        Else
            If chk_sodium.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F861"
                ccount = ccount + 1
            End If
            If chk_pota.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F862"
                ccount = ccount + 1
            End If
            If chk_chlor.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F863"
                ccount = ccount + 1
            End If
            If CHK_CO2.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F864"
                ccount = ccount + 1
            End If
        End If
        '**** Electrolytes
        If chk_calc.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F888"
            ccount = ccount + 1
        End If
        If chk_ioncal.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F889"
            ccount = ccount + 1
        End If
        If chk_phosp.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F890"
            ccount = ccount + 1
        End If
        If chk_magn.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F891"
            ccount = ccount + 1
        End If
        If chk_panenzy.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F880"
            ccount = ccount + 1
        Else
            If chk_amylase.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F881"
                ccount = ccount + 1
            End If
            If chk_lipase.Value = 1 Then
                If ccount >= 11 Then
                    MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                    Unload Me
                    pathology.show
                    Exit Sub
                End If
                pathology.CMB_TEXT.AddItem "F882"
                ccount = ccount + 1
            End If
        End If
        '**** pancreactic enzymes
        If chk_iron.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F893"
            ccount = ccount + 1
        End If
        If chk_tibc.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F894"
            ccount = ccount + 1
        End If
        If chk_acidphosp.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F895"
            ccount = ccount + 1
        End If
        If chk_prostacid.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F896"
            ccount = ccount + 1
        End If
        If chk_csfsug.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F884"
            ccount = ccount + 1
        End If
        If chk_csfprot.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F885"
            ccount = ccount + 1
        End If
        If chk_csfchlor.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F886"
            ccount = ccount + 1
        End If
        If chk_csfldh.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F887"
            ccount = ccount + 1
        End If
'        If chk_urob.Value = 1 Then
'            If ccount >= 11 Then
'                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            pathology.CMB_TEXT.AddItem "F899"
'            ccount = ccount + 1
'        End If
'        If chk_Porphyrin.Value = 1 Then
'            If ccount >= 11 Then
'                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'                Unload Me
'                pathology.show
'                Exit Sub
'            End If
'            pathology.CMB_TEXT.AddItem "F900"
'            ccount = ccount + 1
'        End If
         If chk_ldh.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F876"
            ccount = ccount + 1
        End If
        '******** Shraddha Pandit 17/12/08 ************
        If chk_crp.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F829"
            ccount = ccount + 1
        End If
        '**********************************************
        'GGT
        If chk_ggt.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F999"
            ccount = ccount + 1
        End If
        'GGT
        
        'newly added by shweta 16-june-2010
        '**********************************************
        'serum cpk
        If chk_se_cpk.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F055"
            ccount = ccount + 1
        End If
        'serum cpk
        
        '**********************************************
        'serum cpk-mb
        If chk_se_cpk_mb.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F056"
            ccount = ccount + 1
        End If
        'serum cpk-mb
        
        
        '**********************************************
        'serum lactate
        If chk_lactate.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F057"
            ccount = ccount + 1
        End If
        'serum cpk
      'shweta 26-07-2012


If serumB2_Microglobin.Value = 1 Then
If ccount >= 11 Then
        MsgBox " You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
pathology.CMB_TEXT.AddItem "F821"
ccount = ccount + 1
End If
        
End Sub
Sub STONEANALYSIS1()
        
End Sub
Sub RANDOMURINETEST()
If chk_ureaur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F925"
            ccount = ccount + 1
        End If
        If chk_uricacidur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F926"
            ccount = ccount + 1
        End If
        If chk_creatur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F927"
            ccount = ccount + 1
        End If
    '**** Renal Function Tests
        '**** Electrolytes
        If chk_sodiumur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F928"
            ccount = ccount + 1
        End If
        If chk_potaur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F929"
            ccount = ccount + 1
        End If
        
        '***** Shraddha 17/08/2011 *****
        If chk_umrandom.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F063"
            ccount = ccount + 1
        End If
        '*******************************
        
        If chk_chlorur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F930"
            ccount = ccount + 1
        End If
        '**** Electrolytes
        ''***calcium
        If chk_calcur.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F931"
            ccount = ccount + 1
        End If
        ''***calcium
        ''***phosphorus
        If chk_phospur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F932"
            ccount = ccount + 1
        End If
        ''***phosphorus
        ''***protein
        If chk_protur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F933"
            ccount = ccount + 1
        End If
        ''***protein
        ''*****urinary investigation
        If chk_urobur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F899"
            ccount = ccount + 1
        End If
        If chk_Porphyrinur.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F900"
            ccount = ccount + 1
        End If
        ''*****urinary investigation
    End Sub
Sub BODYFLUIDTEST()

''********SUGAR PROFILE***********************
If CHK_SUGARBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F962"
    ccount = ccount + 1
End If
''********SUGAR PROFILE***********************
''********LIPID PROFILE***********************
If CHK_CHOLBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F943"
    ccount = ccount + 1
End If
' ADD  FOR CALCIUM AND PHOSPHORUS

        If CHK_PHOS.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F969"
            ccount = ccount + 1
        End If
        If CHK_CALCI.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F970"
            ccount = ccount + 1
        End If

'END BY

If CHK_HDLBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F945"
    ccount = ccount + 1
End If
If CHK_TRIBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F944"
    ccount = ccount + 1
End If
If CHK_LDLBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F946"
    ccount = ccount + 1
End If
''********LIPID PROFILE***********************
''********RFT***********************
If CHK_UREABF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F934"
    ccount = ccount + 1
End If
If CHK_URICACIDBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F935"
    ccount = ccount + 1
End If
If CHK_CREATININEBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F936"
    ccount = ccount + 1
End If
''********RFT***********************
''********ELECTRILYTES***********************
If CHK_NABF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F937"
    ccount = ccount + 1
End If
If CHK_CLBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F939"
    ccount = ccount + 1
End If
If CHK_KBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F938"
    ccount = ccount + 1
End If
''********ELECTRILYTES***********************
''********LFT***********************
If CHK_PROBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F963"
    ccount = ccount + 1
End If
If CHK_ALBBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F964"
    ccount = ccount + 1
End If
If CHK_GLOBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F965"
    ccount = ccount + 1
End If
If CHK_ALKBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F966"
    ccount = ccount + 1
End If
If CHK_TOTALBILBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F940"
    ccount = ccount + 1
End If
If CHK_DIRECTBILBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F941"
    ccount = ccount + 1
End If
If CHK_IBBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F942"
    ccount = ccount + 1
End If
If CHK_SGOTBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F967"
    ccount = ccount + 1
End If
If CHK_SGPTBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F968"
    ccount = ccount + 1
End If
''********LFT***********************
''********PANCREATIC***********************
If CHK_AMYBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F971"
    ccount = ccount + 1
End If
If CHK_LIPBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F972"
    ccount = ccount + 1
End If
''********PANCREATIC***********************
If CHK_LDHBF.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F973"
    ccount = ccount + 1
End If

End Sub
Sub urine24hourS()

         '**** Renal Function Tests
        If chk_urea24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F918"
            ccount = ccount + 1
        End If
        If chk_uricacid24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F919"
            ccount = ccount + 1
        End If
    ' ADD FOR VMA AND 5-HIAA TEST
    
        If chk_urinvma.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F897"
            ccount = ccount + 1
        End If
        If chk_urin5hiaa.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F898"
            ccount = ccount + 1
        End If
    'END
        If chk_creat24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F920"
            ccount = ccount + 1
        End If
        If chk_corrcreatclr24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F924"
            ccount = ccount + 1
        End If
    
    '**** Renal Function Tests
        '**** Electrolytes
        If chk_sodium24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F915"
            ccount = ccount + 1
        End If
        If chk_pota24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F916"
            ccount = ccount + 1
        End If
        If chk_chlor24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F917"
            ccount = ccount + 1
        End If
        
        '***** Shraddha 17/08/2011 *****
        If chk_um24hrs.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F063"
            ccount = ccount + 1
        End If
        '*******************************
        '**** Electrolytes
        ''***calcium
        If chk_calc24.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F921"
            ccount = ccount + 1
        End If
        ''***calcium
        ''***phosphorus
        If chk_phosp24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F922"
            ccount = ccount + 1
        End If
        ''***phosphorus
        ''***protein
        If chk_prot24.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F923"
            ccount = ccount + 1
        End If
        ''***protein

End Sub


Private Sub GROUPCCT_Click()
 If GROUPCCT.Value = 1 Then
    CHK_CCC.Value = 1
    CHK_SC.Value = 1
    CHK_UC.Value = 1
 Else
    CHK_CCC.Value = 0
    CHK_SC.Value = 0
    CHK_UC.Value = 0
 
 End If
 
           
End Sub


