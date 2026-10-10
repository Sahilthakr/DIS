VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form biochemistry 
   Caption         =   "Biochemistry "
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15555
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
   ScaleHeight     =   9630
   ScaleWidth      =   15555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   240
      TabIndex        =   87
      Top             =   8040
      Width           =   5295
      Begin VB.CheckBox serumB2_Microglobin 
         Caption         =   "Serum B2 - Microglobulin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
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
      Left            =   6720
      TabIndex        =   85
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ComboBox CMB_NMETERIAL 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "biochemistry.frx":0000
      Left            =   1920
      List            =   "biochemistry.frx":0002
      TabIndex        =   0
      Text            =   "BLOOD"
      Top             =   1200
      Width           =   3735
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
      Left            =   8880
      TabIndex        =   79
      Top             =   6840
      Width           =   2775
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
         Left            =   120
         TabIndex        =   80
         Top             =   0
         Width           =   1935
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
      Height          =   375
      Left            =   10440
      TabIndex        =   77
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6000
      TabIndex        =   75
      Top             =   840
      Width           =   5655
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   1920
      TabIndex        =   73
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.Frame UI 
      Caption         =   "Urinary -  Investigations"
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
      Left            =   5520
      TabIndex        =   64
      Top             =   6840
      Width           =   3375
      Begin VB.CheckBox chk_Porphyrin 
         Caption         =   "Porphyrins"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   68
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chk_urob 
         Caption         =   "Urobilinogen"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   66
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chk_urin5hiaa 
         Caption         =   "Urinary 5HIAA"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chk_urinvma 
         Caption         =   "Urinary VMA"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame BFI 
      Caption         =   "Body Fluid - Investigations"
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
      Left            =   240
      TabIndex        =   55
      Top             =   6840
      Width           =   5295
      Begin VB.ComboBox CMB_BODYF 
         Height          =   315
         Left            =   2880
         TabIndex        =   56
         Top             =   0
         Width           =   2295
      End
      Begin VB.CheckBox chk_lipase1 
         Caption         =   "Lipase"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3720
         TabIndex        =   62
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chk_amylase1 
         Caption         =   "Amylase"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3720
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chk_csfldh 
         Caption         =   " LDH"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   61
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox chk_csfchlor 
         Caption         =   "Chlorides"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chk_csfprot 
         Caption         =   "Proteins"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   58
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chk_csfsug 
         Caption         =   "Sugar"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FRM_ACID 
      Height          =   1095
      Left            =   5160
      TabIndex        =   49
      Top             =   5640
      Width           =   2655
      Begin VB.CheckBox chk_prostacid 
         Caption         =   "Prost.Acid Phosph"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chk_acidphosp 
         Caption         =   "Acid Phosph"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame FRM_IRON 
      Height          =   1095
      Left            =   3240
      TabIndex        =   46
      Top             =   5640
      Width           =   1815
      Begin VB.CheckBox chk_tibc 
         Caption         =   "T.I.B.C"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chk_iron 
         Caption         =   "Iron"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame PE 
      Height          =   1095
      Left            =   240
      TabIndex        =   42
      Top             =   5640
      Width           =   2895
      Begin VB.CheckBox chk_lipase 
         Caption         =   "Lipase"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chk_amylase 
         Caption         =   "Amylase"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   1575
      End
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
         TabIndex        =   43
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Frame CE 
      Height          =   1095
      Left            =   8040
      TabIndex        =   52
      Top             =   5640
      Width           =   3615
      Begin VB.CheckBox chk_cpkmb 
         Caption         =   "CPK-MB"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   72
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chk_cpk 
         Caption         =   "CPK"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chk_ldh1 
         Caption         =   "S.G.O.T"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   69
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chk_ldh 
         Caption         =   "LDH"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   855
      End
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
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.Frame LFT 
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   240
      TabIndex        =   31
      Top             =   4320
      Width           =   11415
      Begin VB.CheckBox chk_indirbil 
         Caption         =   "Indirect Bilirubin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chk_sgpt 
         Caption         =   "S.G.P.T"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   41
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chk_sgot 
         Caption         =   "S.G.O.T"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5520
         TabIndex        =   40
         Top             =   900
         Width           =   1215
      End
      Begin VB.CheckBox chk_dirbil 
         Caption         =   "Direct Bilirubin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chk_totbil 
         Caption         =   "Total Bilirubin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9480
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chk_alkphos 
         Caption         =   "Alk.Phosphatase"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   36
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chk_glob 
         Caption         =   "Globulins"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk_albumin 
         Caption         =   "Albumin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk_prot 
         Caption         =   "Proteins"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
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
         TabIndex        =   32
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Frame FRM_CALCIUM 
      Height          =   1095
      Left            =   8400
      TabIndex        =   26
      Top             =   3120
      Width           =   3255
      Begin VB.CheckBox chk_magn 
         Caption         =   "Magnesium"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chk_phosp 
         Caption         =   "Phosphorus"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chk_ioncal 
         Caption         =   "Ionic Calcium"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chk_calc 
         Caption         =   "Calcium"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame ELECTROLYTES 
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   5400
      TabIndex        =   20
      Top             =   3120
      Width           =   2895
      Begin VB.CheckBox CHK_CO2 
         Caption         =   "Carbon dioxide"
         Height          =   360
         Left            =   1560
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chk_chlor 
         Caption         =   "Chlorides"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk_pota 
         Caption         =   "Potassium"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chk_sodium 
         Caption         =   "Sodium"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
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
         TabIndex        =   21
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame RFT 
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   5055
      Begin VB.CheckBox chk_corrcreatclr 
         Caption         =   "Corrected Creat. Clear"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chk_creat 
         Caption         =   "Creatinine"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chk_uricacid 
         Caption         =   "Uric Acid"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chk_urea 
         Caption         =   "Urea"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
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
         TabIndex        =   15
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Frame LP 
      Caption         =   "LIPID PROFILE"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   6720
      TabIndex        =   8
      Top             =   1560
      Width           =   4935
      Begin VB.CheckBox chk_ldlcholest 
         Caption         =   "LDL Cholesterol"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   720
         Width           =   1815
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
         TabIndex        =   9
         Top             =   0
         Width           =   1815
      End
      Begin VB.CheckBox chk_trigly 
         Caption         =   "Triglycerides"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chk_hdlcholest 
         Caption         =   "HDL Cholesterol"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chk_cholest 
         Caption         =   "Cholesterol"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
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
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   6375
      Begin VB.CheckBox chk_glucotolerancetest 
         Caption         =   "Glucose Tolerance Test"
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chk_glyhb 
         Caption         =   "Glycosylated Hb"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chk_glycahbb 
         Caption         =   "Glycateo H.B.B"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         MaskColor       =   &H00808080&
         TabIndex        =   78
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chk_ppurinsug 
         Caption         =   "P.P. Urine Sugar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chk_furinsug 
         Caption         =   "F Urine Sugar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chk_glucosug 
         Caption         =   "Glucocheck Sugar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   53
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chk_ransug 
         Caption         =   "Random Sugar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chk_ppsug 
         Caption         =   "PP Sugar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chk_trusugfas 
         Caption         =   "True Sugar Fasting"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label Label3 
      Caption         =   "This form is not in use"
      Height          =   255
      Left            =   240
      TabIndex        =   86
      Top             =   360
      Width           =   2055
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
      Left            =   360
      TabIndex        =   83
      Top             =   9840
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE,MULLANPUR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2520
      TabIndex        =   82
      Top             =   120
      Width           =   11655
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
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
      Left            =   2760
      TabIndex        =   81
      Top             =   480
      Width           =   6015
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
      Left            =   5280
      TabIndex        =   76
      Top             =   840
      Width           =   630
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
      Left            =   960
      TabIndex        =   74
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lbl_specimen 
      AutoSize        =   -1  'True
      Caption         =   "Nature of Metarial"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   44
      Top             =   1200
      Width           =   1545
   End
End
Attribute VB_Name = "biochemistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim ccount As Integer
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

Private Sub chk_creat_Click()
If chk_creat.Value = 0 Then
    chk_renfuntest.Value = 0
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

Private Sub chk_uricacid_Click()
If chk_uricacid.Value = 0 Then
    chk_renfuntest.Value = 0
End If
End Sub

Private Sub CMB_BODYF_GotFocus()

CMB_BODYF.clear
CMB_BODYF.AddItem "CSF"
CMB_BODYF.Text = "CSF"
'CMB_BODYF.AddItem "URINE "
'CMB_BODYF.AddItem "PLEURAL FLUID"
'CMB_BODYF.AddItem "ASCITIC FLUID"
'CMB_BODYF.AddItem "DRAIN FLUID"
End Sub

Private Sub CMB_BODYF_LostFocus()
bodyf = CMB_BODYF.Text
End Sub

Private Sub cmb_nmeterial_GotFocus()
CMB_NMETERIAL.clear
CMB_NMETERIAL.AddItem "Blood"
'COMMENTED FOR TEMPORARY
'CMB_NMETERIAL.AddItem "Body Fluid"
'CMB_NMETERIAL.AddItem "Random Urine"
'CMB_NMETERIAL.AddItem "24 hours Urine"
'CMB_NMETERIAL.AddItem "Stone Analysis"
End Sub

Private Sub CMB_NMETERIAL_LostFocus()
meterial = CMB_NMETERIAL.Text
End Sub

Private Sub cmd_back_Click()

meterial = CMB_NMETERIAL.Text
pathology.CMB_TEXT.clear
ccount = 0
If BCFLAG = False Then


'VENKAT



If pathology.CMB_TEXT.ListCount = 0 Then
    MsgBox ("Please atleast 1 procedure code"), vbInformation
    Exit Sub
End If


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
''''''''


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
                    If chk_corrcreatclr.Value = 1 Then
        
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
        
                If chk_corrcreatclr.Value = 1 Then
        
                    If ccount >= 11 Then
                        MsgBox " You Can Select Max. of 11 Procedures only."
                        Unload Me
                        pathology.show
                        Exit Sub
                    End If
                    pathology.CMB_TEXT.AddItem "F859"
                    ccount = ccount + 1
                End If
        End If
        '**** Renal Function Tests
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



'comented for test
''If chk_trusugfas.Value = 1 Then
''
''    If ccount >= 11 Then
''        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''        Unload Me
''        pathology.Show
''        Exit Sub
''    End If
''    ccount = ccount + 1
''    pathology.CMB_TEXT.AddItem "F842"
''End If

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

If chk_glucotolerancetest.Value = 1 Then
    
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    ccount = ccount + 1
    pathology.CMB_TEXT.AddItem "F810"
End If

'If chk_glycahbb.Value = 1 Then
'
'    If ccount >= 11 Then
'        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'        Unload Me
'        pathology.Show
'        Exit Sub
'    End If
'    pathology.CMB_TEXT.AddItem "F845"
'    ccount = ccount + 1
'End If
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



'commented for testing
'**** LIQUID PROFILE
'''**** Renal Function Tests
''If chk_renfuntest.Value = 1 Then
''            If ccount >= 11 Then
''                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''                Unload Me
''                pathology.Show
''                Exit Sub
''            End If
''            pathology.CMB_TEXT.AddItem "F854"
''            ccount = ccount + 1
''            If chk_corrcreatclr.Value = 1 Then
''
''            If ccount >= 11 Then
''                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''                Unload Me
''                pathology.Show
''                Exit Sub
''            End If
''            pathology.CMB_TEXT.AddItem "F859"
''            ccount = ccount + 1
''        End If
''Else
''        If chk_urea.Value = 1 Then
''
''            If ccount >= 11 Then
''                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''                Unload Me
''                pathology.Show
''                Exit Sub
''            End If
''            pathology.CMB_TEXT.AddItem "F855"
''            ccount = ccount + 1
''        End If
''
''        If chk_uricacid.Value = 1 Then
''
''            If ccount >= 11 Then
''                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''                Unload Me
''                pathology.Show
''                Exit Sub
''            End If
''            pathology.CMB_TEXT.AddItem "F856"
''            ccount = ccount + 1
''        End If
''
''        If chk_creat.Value = 1 Then
''
''            If ccount >= 11 Then
''                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''                Unload Me
''                pathology.Show
''                Exit Sub
''            End If
''            pathology.CMB_TEXT.AddItem "F857"
''            ccount = ccount + 1
''        End If
''
''        If chk_corrcreatclr.Value = 1 Then
''
''            If ccount >= 11 Then
''                MsgBox " You Can Select Max. of 11 Procedures only."
''                Unload Me
''                pathology.Show
''                Exit Sub
''            End If
''            pathology.CMB_TEXT.AddItem "F859"
''            ccount = ccount + 1
''        End If
''End If
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
' commented fro testing
'''**** Liver function Tests
''If chk_livfuntest.Value = 1 Then
''   If ccount > 11 Then
''        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''        Unload Me
''        pathology.Show
''        Exit Sub
''    End If
''    pathology.CMB_TEXT.AddItem "F865"
''    ccount = ccount + 1
''
''    If chk_dirbil.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F871"
''        ccount = ccount + 1
''    End If
''    If chk_indirbil.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F872"
''        ccount = ccount + 1
''    End If
''
''
''
''Else
''    If chk_prot.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F866"
''        ccount = ccount + 1
''    End If
''
''    If chk_albumin.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F867"
''        ccount = ccount + 1
''    End If
''
''    If chk_glob.Value = 1 Then
''
''        If ccount > 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F868"
''        ccount = ccount + 1
''    End If
''
''    If chk_alkphos.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F869"
''        ccount = ccount + 1
''    End If
''
''    If chk_totbil.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F870"
''        ccount = ccount + 1
''    End If
''    If chk_dirbil.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F871"
''        ccount = ccount + 1
''    End If
''    If chk_indirbil.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F872"
''        ccount = ccount + 1
''    End If
''
''    If chk_sgot.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F873"
''        ccount = ccount + 1
''    End If
''
''    If chk_sgpt.Value = 1 Then
''
''        If ccount >= 11 Then
''            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
''            Unload Me
''            pathology.Show
''            Exit Sub
''        End If
''        pathology.CMB_TEXT.AddItem "F874"
''        ccount = ccount + 1
''    End If
''
''End If
'''**** Liver function Tests
'**** Cardiac enzymes
'If chk_carenzy.Value = 1 Then
'    If ccount > 11 Then
'        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'        Unload Me
'        pathology.Show
'        Exit Sub
'    End If
'    pathology.CMB_TEXT.AddItem "F873"
'    ccount = ccount + 1
'Else
'    If chk_LDH.Value = 1 Then
'
'    If ccount >= 11 Then
'        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'        Unload Me
'        pathology.Show
'        Exit Sub
'    End If
'    pathology.CMB_TEXT.AddItem "F874"
'    ccount = ccount + 1
'End If
'
'If chk_ldh1.Value = 1 Then
'
'    If ccount >= 11 Then
'        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'        Unload Me
'        pathology.Show
'        Exit Sub
'    End If
'    pathology.CMB_TEXT.AddItem "F871"
'    ccount = ccount + 1
'End If
'
'If chk_CPK.Value = 1 Then
'
'    If ccount >= 11 Then
'        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'        Unload Me
'        pathology.Show
'        Exit Sub
'    End If
'    pathology.CMB_TEXT.AddItem "F876"
'    ccount = ccount + 1
'End If
'
'If chk_CPKMB.Value = 1 Then
'
'    If ccount >= 11 Then
'        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'        Unload Me
'        pathology.Show
'        Exit Sub
'    End If
'    pathology.CMB_TEXT.AddItem "F877"
'    ccount = ccount + 1
'End If
'
'End If
'**** Cardiac enzymes
'**** pancreactic enzymes
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

If chk_amylase1.Value = 1 Then

    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F881"
    ccount = ccount + 1
End If

If chk_lipase1.Value = 1 Then

    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F882"
    ccount = ccount + 1
End If
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

If chk_urob.Value = 1 Then

    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F899"
    ccount = ccount + 1
End If
If chk_Porphyrin.Value = 1 Then
    If ccount >= 11 Then
        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
        Unload Me
        pathology.show
        Exit Sub
    End If
    pathology.CMB_TEXT.AddItem "F900"
    ccount = ccount + 1
End If
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
End If
BCFLAG = False
Unload Me
pathology.show
'pathology.Biochemistry1.Visible = True
End Sub
Private Sub Form_Load()
ccount = 0
meterial = ""
bodyf = ""
End Sub
Private Sub txt_nmeterial_LostFocus()
meterial = Trim(txt_nmeterial.Text)
End Sub

