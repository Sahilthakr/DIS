VERSION 5.00
Begin VB.Form cytogeneticreq 
   Caption         =   "Cancer Cytogenetics Laboratory"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_save 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4665
      TabIndex        =   52
      Top             =   8415
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "&Back"
      Height          =   345
      Left            =   10560
      TabIndex        =   51
      Top             =   8235
      Width           =   1290
   End
   Begin VB.Frame cytogenetics 
      ForeColor       =   &H000000C0&
      Height          =   3375
      Left            =   75
      TabIndex        =   0
      Top             =   4995
      Width           =   11775
      Begin VB.Frame chk_amlm6 
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   0
         TabIndex        =   70
         Top             =   1320
         Width           =   11745
         Begin VB.CheckBox chk_amlm0 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M0"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   210
            TabIndex        =   88
            Top             =   75
            Width           =   1065
         End
         Begin VB.CheckBox chk_amlm1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M1"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   210
            TabIndex        =   87
            Top             =   360
            Width           =   1020
         End
         Begin VB.CheckBox chk_amlm2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M2"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   210
            TabIndex        =   86
            Top             =   630
            Width           =   1065
         End
         Begin VB.CheckBox chk_amlm3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M3"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1830
            TabIndex        =   85
            Top             =   90
            Width           =   1140
         End
         Begin VB.CheckBox chk_amlm4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M4"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1830
            TabIndex        =   84
            Top             =   360
            Width           =   1365
         End
         Begin VB.CheckBox chk_amlm4e0 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M4/E0"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1815
            TabIndex        =   83
            Top             =   630
            Width           =   1470
         End
         Begin VB.CheckBox chk_m5a 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M5a"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3585
            TabIndex        =   82
            Top             =   90
            Width           =   1155
         End
         Begin VB.CheckBox chk_m5b 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M5b"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3585
            TabIndex        =   81
            Top             =   405
            Width           =   1215
         End
         Begin VB.CheckBox chk_aml6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M6"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3585
            TabIndex        =   80
            Top             =   705
            Width           =   1080
         End
         Begin VB.CheckBox chk_amlm7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML-M7"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5340
            TabIndex        =   79
            Top             =   60
            Width           =   1035
         End
         Begin VB.CheckBox chk_bioleu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Biphenotypic Leukemia"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   5340
            TabIndex        =   78
            Top             =   285
            Width           =   1830
         End
         Begin VB.CheckBox chk_preball 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Pre B-ALL"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5340
            TabIndex        =   77
            Top             =   660
            Width           =   1365
         End
         Begin VB.CheckBox chk_tall 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "T-ALL"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7365
            TabIndex        =   76
            Top             =   75
            Width           =   900
         End
         Begin VB.CheckBox chk_cmlcp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "CML-CP"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7350
            TabIndex        =   75
            Top             =   345
            Width           =   1080
         End
         Begin VB.CheckBox chk_cmlbc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "CML-BC"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7350
            TabIndex        =   74
            Top             =   630
            Width           =   1095
         End
         Begin VB.CheckBox chk_cmlap 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "CML-AP"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8805
            TabIndex        =   73
            Top             =   60
            Width           =   1140
         End
         Begin VB.CheckBox chk_burkitt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Burkitt"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8820
            TabIndex        =   72
            Top             =   360
            Width           =   1020
         End
         Begin VB.CheckBox chk_cll 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "CLL"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8820
            TabIndex        =   71
            Top             =   660
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Clinico-hematological Diagnosis"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Width           =   3255
         End
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   0
         TabIndex        =   57
         Top             =   840
         Width           =   11760
         Begin VB.CheckBox chk_mds 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "MDS"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10830
            TabIndex        =   66
            Top             =   267
            Width           =   765
         End
         Begin VB.CheckBox chk_aplane 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Aplastic Anemia"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9366
            TabIndex        =   65
            Top             =   120
            Width           =   1035
         End
         Begin VB.CheckBox chk_lymphoma 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "LYMPHOMA"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5910
            TabIndex        =   64
            Top             =   225
            Width           =   1425
         End
         Begin VB.CheckBox chk_nhl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "NHL"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4725
            TabIndex        =   63
            Top             =   239
            Width           =   750
         End
         Begin VB.CheckBox chk_cml 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "CML"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3504
            TabIndex        =   62
            Top             =   246
            Width           =   795
         End
         Begin VB.CheckBox chk_all 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "ALL"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2415
            TabIndex        =   61
            Top             =   255
            Width           =   720
         End
         Begin VB.CheckBox chk_fanane 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Fanconi's anemia"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7770
            TabIndex        =   60
            Top             =   135
            Width           =   1155
         End
         Begin VB.CheckBox chk_aml 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AML"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1365
            TabIndex        =   59
            Top             =   240
            Width           =   750
         End
         Begin VB.CheckBox chk_al 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "AL"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   58
            Top             =   255
            Width           =   630
         End
      End
      Begin VB.TextBox txt_phone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   56
         Top             =   3075
         Width           =   1545
      End
      Begin VB.TextBox txt_fax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   54
         Top             =   3060
         Width           =   1620
      End
      Begin VB.ComboBox cmb_cyspreinv 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7410
         TabIndex        =   25
         Top             =   2925
         Width           =   4215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Clinical Details"
         Height          =   855
         Left            =   0
         TabIndex        =   14
         Top             =   75
         Width           =   11775
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   9840
            TabIndex        =   67
            Top             =   120
            Width           =   1815
            Begin VB.OptionButton Option2 
               Caption         =   "NO"
               Height          =   315
               Left            =   960
               TabIndex        =   69
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Option1 
               Caption         =   "YES"
               Height          =   315
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.CheckBox chk_at_dig 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "At Diagnosis"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   165
            TabIndex        =   22
            Top             =   225
            Width           =   1455
         End
         Begin VB.CheckBox chk_pr_treat 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Partially Treated"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2700
            TabIndex        =   21
            Top             =   195
            Width           =   1815
         End
         Begin VB.CheckBox chk_ps_indu 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Post Induction"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5220
            TabIndex        =   20
            Top             =   180
            Width           =   1695
         End
         Begin VB.CheckBox chk_remission 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Remission"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   165
            TabIndex        =   19
            Top             =   525
            Width           =   1695
         End
         Begin VB.CheckBox chk_relapse 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Relapse"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5220
            TabIndex        =   18
            Top             =   555
            Width           =   1095
         End
         Begin VB.TextBox txt_followup 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7200
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   510
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CheckBox chk_postbmt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Post BMT"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2700
            TabIndex        =   16
            Top             =   540
            Width           =   1335
         End
         Begin VB.ComboBox cmb_sex 
            Height          =   315
            ItemData        =   "cytogeneticreq.frx":0000
            Left            =   8775
            List            =   "cytogeneticreq.frx":000A
            TabIndex        =   15
            Top             =   135
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Follow -Up"
            Height          =   210
            Left            =   8880
            TabIndex        =   24
            Top             =   510
            Width           =   945
         End
         Begin VB.Label Label12 
            Caption         =   "BMT Donor Sex"
            Height          =   255
            Left            =   7260
            TabIndex        =   23
            Top             =   165
            Width           =   1575
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Specimen Type"
         Height          =   570
         Left            =   0
         TabIndex        =   2
         Top             =   2475
         Width           =   7335
         Begin VB.OptionButton opt_BMA 
            Caption         =   "Bone Marrow Aspiration"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton opt_option 
            Caption         =   "Others"
            Height          =   240
            Left            =   6225
            TabIndex        =   13
            Top             =   285
            Width           =   975
         End
         Begin VB.CheckBox chk_oa 
            Caption         =   "Only Aspirate"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chk_oi 
            Caption         =   "Only Imprint"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox CHK_ASPIMP 
            Caption         =   "Aspirate+Imprint"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chk_obas 
            Caption         =   "Outside BM asp.Slides"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox chk_oaspimp 
            Caption         =   "Outside BM asp.+imp."
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox CHK_OUTSIDEIMPRINT 
            Caption         =   "Outside Imprint"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Out side BM Biopsy"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton opt_PBlood 
            Caption         =   "Peripheral Blood"
            Height          =   255
            Left            =   2625
            TabIndex        =   4
            Top             =   270
            Width           =   1935
         End
         Begin VB.OptionButton opt_lymbnode 
            Caption         =   "Lymph Node"
            Height          =   255
            Left            =   4575
            TabIndex        =   3
            Top             =   255
            Width           =   1440
         End
         Begin VB.Label lbl_ihreqno 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Req No."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1035
            TabIndex        =   27
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.TextBox txt_others 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   7410
         TabIndex        =   1
         Top             =   2475
         Width           =   4215
      End
      Begin VB.Label Label28 
         Caption         =   "Phone"
         Height          =   300
         Left            =   2415
         TabIndex        =   55
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label Label27 
         Caption         =   "Fax"
         Height          =   240
         Left            =   60
         TabIndex        =   53
         Top             =   3120
         Width           =   405
      End
      Begin VB.Label lbl_hispreins 
         AutoSize        =   -1  'True
         Caption         =   "Previously Investigated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5280
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   1980
      End
   End
   Begin VB.Label Label26 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   570
      TabIndex        =   50
      Top             =   4335
      Width           =   270
   End
   Begin VB.Label Label25 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   585
      TabIndex        =   49
      Top             =   4005
      Width           =   300
   End
   Begin VB.Label Label24 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   570
      TabIndex        =   48
      Top             =   3465
      Width           =   255
   End
   Begin VB.Label Label23 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   585
      TabIndex        =   47
      Top             =   2325
      Width           =   315
   End
   Begin VB.Label Label8 
      Caption         =   $"cytogeneticreq.frx":0014
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   900
      TabIndex        =   46
      Top             =   4305
      Width           =   10860
   End
   Begin VB.Label Label7 
      Caption         =   "� Label the container/tube/bulb with patients name, case number, control's name and time of collection."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   45
      Top             =   3975
      Width           =   9795
   End
   Begin VB.Label Label22 
      Caption         =   "Transportation:"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   270
      TabIndex        =   44
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   $"cytogeneticreq.frx":012B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   915
      TabIndex        =   43
      Top             =   3465
      Width           =   10860
   End
   Begin VB.Label Label20 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   585
      TabIndex        =   42
      Top             =   2820
      Width           =   240
   End
   Begin VB.Label Label19 
      Caption         =   $"cytogeneticreq.frx":01EF
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   930
      TabIndex        =   41
      Top             =   2790
      Width           =   10725
   End
   Begin VB.Label Label18 
      Caption         =   "Chromosome breakage studies in Fanconi's and Aplastic anemia: "
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   330
      TabIndex        =   40
      Top             =   2565
      Width           =   5535
   End
   Begin VB.Label Label17 
      Caption         =   "� After collection of BMA and PB shake the container/tube/bulb to avoid clotting.  Do not freeze the specimen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   825
      TabIndex        =   39
      Top             =   2325
      Width           =   9615
   End
   Begin VB.Label Label16 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   600
      TabIndex        =   38
      Top             =   2040
      Width           =   225
   End
   Begin VB.Label Label15 
      Caption         =   "� Peripheral Blood (PB) Collection:  Collect 3-5 ml of sodium heparinized PB into a sterile heparinized vacutainer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   37
      Top             =   2025
      Width           =   10290
   End
   Begin VB.Label Label14 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   615
      TabIndex        =   36
      Top             =   1740
      Width           =   240
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   630
      TabIndex        =   35
      Top             =   1230
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "=>"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   630
      TabIndex        =   34
      Top             =   900
      Width           =   225
   End
   Begin VB.Label Label9 
      Caption         =   "� Outside Centres can collect BM Aspirate into a sterile heparinized vacutainer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   825
      TabIndex        =   33
      Top             =   1740
      Width           =   7065
   End
   Begin VB.Label Label6 
      Caption         =   $"cytogeneticreq.frx":02C9
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   945
      TabIndex        =   32
      Top             =   1215
      Width           =   10590
   End
   Begin VB.Label Label5 
      Caption         =   "� Peripheral blood is accepted only if blast percentage in peripheral blood is more than 30%."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   31
      Top             =   900
      Width           =   8880
   End
   Begin VB.Label Label4 
      Caption         =   "Leukemia specimens:"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   435
      TabIndex        =   30
      Top             =   645
      Width           =   1920
   End
   Begin VB.Label Label3 
      Caption         =   "Instruction for collection and transportation of specimens for Cytogenetic studies."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   2355
      TabIndex        =   29
      Top             =   375
      Width           =   8235
   End
   Begin VB.Label Label2 
      Caption         =   "CYTOGENETIC SERVICES INSTRUCTIONS"
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
      Height          =   315
      Left            =   3195
      TabIndex        =   28
      Top             =   90
      Width           =   5475
   End
End
Attribute VB_Name = "cytogeneticreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELRINVESTCY_QR As rdoQuery
Dim SELRINVESTCY_RS As rdoResultset
Dim SELRINVESTCY_FLAG As Boolean

Dim INSCYTOGCLDTA_QR As rdoQuery
Dim UPDCYTOGCLDTA_QR As rdoQuery

Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITN_FLAG As Boolean

Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
Dim SELMA101_FLAG As Boolean

Dim SELDAUTOREQNO_QR As rdoQuery
Dim SELDAUTOREQNO_RS As rdoResultset
Dim SELDAUTOREQNOFLAG As Boolean

Dim INSDAUTOREQNO_QR As rdoQuery
Dim UPDDAUTOREQNO_QR As rdoQuery

Dim INSDREQUISITN_QR As rdoQuery
Dim upddrequisitn_qr As rdoQuery

Dim SELDISICU_QR As rdoQuery
Dim SELDISICU_RS As rdoResultset
Dim SELDISICU_FLAG As Boolean

Dim SELMA101CY_QR As rdoQuery
Dim SELMA101CY_RS As rdoResultset
Dim SELMA101CY_FLAG As Boolean

Dim selreqnno_qr As rdoQuery
Dim selreqnno_rs As rdoResultset

Dim selreqmoda_qr As rdoQuery
Dim slereqmoda_rs As rdoResultset

Dim INSRREQUISITN_QR As rdoQuery
Dim updrrequisitn_qr As rdoQuery

Private Sub chk_ALL1_Click()
If chk_ALL1.Value = 1 Then
    chk_tall.Value = 1
    chk_biapheno.Value = 1
    chk_prebl.Value = 1
Else
    chk_tall.Value = 0
    chk_biapheno.Value = 0
    chk_prebl.Value = 0
End If
End Sub

Private Sub chk_AML1_Click()
If chk_AML1.Value = 1 Then
    chk_aml.Value = 1
    chk_amlm0.Value = 1
    chk_amlm1.Value = 1
    chk_amlm2.Value = 1
    chk_amlm4.Value = 1
    chk_AMLEO.Value = 1
    chk_m5a.Value = 1
    chk_m5b.Value = 1
    chk_aml6.Value = 1
    chk_amlm7.Value = 1
Else
    chk_aml.Value = 0
    chk_amlm0.Value = 0
    chk_amlm1.Value = 0
    chk_amlm2.Value = 0
    chk_amlm4.Value = 0
    chk_AMLEO.Value = 0
    chk_m5a.Value = 0
    chk_m5b.Value = 0
    chk_aml6.Value = 0
    chk_amlm7.Value = 0
End If
End Sub

Private Sub chk_al_Click()
If chk_al.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "AL"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "AL"
   End If
End If
End Sub

Private Sub chk_all_Click()
 If chk_all.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "ALL"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "ALL"
   End If
End If
End Sub

Private Sub chk_aml_Click()
If chk_al.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "AML"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "AML"
   End If
End If
End Sub

Private Sub chk_aml6_Click()
 If chk_aml6.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M6"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M6"
   End If

End If
End Sub

Private Sub chk_amlm0_Click()
'R1 = MSFlexGrid1.Row
'C1 = MSFlexGrid1.Col
End Sub

Private Sub chk_amlm1_Click()
 If chk_amlm1.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M1"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M1"
   End If

End If
End Sub

Private Sub chk_amlm2_Click()
 If chk_amlm2.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M2"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M2"
   End If
End If
End Sub

Private Sub chk_amlm3_Click()
 If chk_amlm3.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M3"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M3"
   End If

End If
End Sub

Private Sub chk_amlm4_Click()
 If chk_amlm4.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M4"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M4"
   End If

End If
End Sub

Private Sub chk_amlm4e0_Click()
 If chk_amlm4e0.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M4/E0"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M4/E0"
   End If
End If
End Sub

Private Sub chk_amlm7_Click()
 If chk_amlm7.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M7"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M7"
   End If

End If
End Sub

Private Sub chk_aplane_Click()
 If chk_aplane.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "APLASTIC ANEMIA"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "APLASTIC ANEMIA"
   End If

End If
End Sub

Private Sub chk_bioleu_Click()
 If chk_bioleu.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "BIPHENOTYPIC LEUKEMIA"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "BIPHENOTYPIC LEUKEMIA"
   End If

End If
End Sub

Private Sub chk_burkitt_Click()
 If chk_burkitt.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "Burkitt"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "Burkitt"
   End If

End If
End Sub

Private Sub chk_cll_Click()
 If chk_cll.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "CLL"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "CLL"
   End If

End If
End Sub

Private Sub chk_CML_Click()
If chk_cml.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "CML"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "CML"
   End If

End If
End Sub

Private Sub chk_cmlap_Click()
 If chk_cmlap.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "CML-AP"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "CML-AP"
   End If

End If
End Sub

Private Sub chk_cmlbc_Click()
 If chk_cmlbc.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "CML-BC"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "CML-BC"
   End If

End If
End Sub

Private Sub chk_cmlcp_Click()
 If chk_cmlcp.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "CML-CP"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "CML-CP"
   End If

End If
End Sub

Private Sub chk_fanane_Click()
 If chk_fanane.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "FANCONI'S ANEMIA"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "FANCONI'S ANEMIA"
   End If
End If
End Sub

Private Sub chk_lymphoma_Click()
 If chk_lymphoma.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "LYMPHOMA"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "LYMPHOMA"
   End If
End If
End Sub

Private Sub chk_m5a_Click()
 If chk_m5a.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M5a"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M5a"
   End If

End If
End Sub

Private Sub chk_m5b_Click()
 If chk_amlm0.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "AML-M5b"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M5b"
   End If
End If
End Sub

Private Sub chk_mds_Click()
 If chk_mds.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "MDS"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "MDS"
   End If

End If
End Sub

Private Sub chk_nhl_Click()
 If chk_nhl.Value = 1 Then
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag1 = "NHL"
   Else
      StrCgdiag1 = StrCgdiag1 & ", " & "NHL"
   End If
End If
End Sub

Private Sub chk_preball_Click()
 If chk_preball.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "Pre B-ALL"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "Pre B-ALL"
   End If

End If
End Sub

Private Sub chk_tall_Click()
 If chk_tall.Value = 1 Then
   
   If Trim(StrCgdiag1) = "" Then
      StrCgdiag2 = "T-ALL"
   Else
      StrCgdiag2 = StrCgdiag2 & ", " & "T-ALL"
   End If

End If
End Sub

Private Sub cmb_cyprocedure_Click()
'cmb_cyprocode.ListIndex = cmb_cyprocedure.ListIndex
End Sub

Private Sub cmb_cyprocode_Click()
'cmb_cyprocedure.ListIndex = cmb_cyprocode.ListIndex
End Sub

Private Sub cmd_back_Click()

If chk_at_dig.Value = 1 Then
  detail = detail & " " & "At Diagnosis"
End If
If chk_pr_treat.Value = 1 Then
   detail = detail & " " & "Partially Treated"
End If
If chk_ps_indu.Value = 1 Then
   detail = detail & " " & "Post Induction"
End If
If chk_relapse.Value = 1 Then
  detail = detail & " " & "Relapse"
End If
If chk_remission.Value = 1 Then
   detail = detail + " " & "Remission"
End If
If Option1.Value = True Then
  txt_followup = "FOLLOW-UP"
End If
If Option2.Value = True Then
  txt_followup = ""
End If
If txt_followup <> "" Then
   detail = detail & " " & Trim(txt_followup)
End If
If chk_postbmt.Value = 1 Then
   detail = detail & " " & "Post BMT"
End If
StrCgdiag2 = ""
StrCgdiag1 = ""
'''
If chk_amlm0.Value = 1 Then
   StrCgdiag2 = StrCgdiag2 & ", " & "AML-M0"
End If

If chk_amlm1.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M1"
End If
If chk_amlm2.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M2"
End If
If chk_amlm3.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M3"
End If
If chk_amlm4.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M4"
End If
'if chk_aml
If chk_m5a.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M5a"
End If
If chk_amlm4e0.Value = 1 Then
  StrCgdiag2 = StrCgdiag2 & ", " & "AML-M4E0"
End If
If chk_m5b.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M5b"
End If
If chk_aml6.Value = 1 Then
     StrCgdiag2 = StrCgdiag2 & ", " & "AML-M6"
End If
If chk_amlm7.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "AML-M7"
End If
If chk_bioleu.Value = 1 Then
      StrCgdiag2 = StrCgdiag2 & ", " & "BIPHENOTYPIC LEUKEMIA"
End If
If chk_preball.Value = 1 Then
 StrCgdiag2 = StrCgdiag2 & ", " & "Pre B-ALL"
End If
 If chk_tall.Value = 1 Then
   StrCgdiag2 = StrCgdiag2 & ", " & "T-ALL"
End If

If chk_cmlcp.Value = 1 Then
  StrCgdiag2 = StrCgdiag2 & ", " & "CML-CP"
End If
 If chk_cmlbc.Value = 1 Then
   StrCgdiag2 = StrCgdiag2 & ", " & "CML-BC"
 End If
If chk_cmlap.Value = 1 Then
  
   StrCgdiag2 = StrCgdiag2 & ", " & "CML-AP"
 End If

 If chk_burkitt.Value = 1 Then
   
   StrCgdiag2 = StrCgdiag2 & ", " & "Burkitt"

End If

 If chk_cll.Value = 1 Then
        StrCgdiag2 = StrCgdiag2 & ", " & "CLL"
 End If
''' 6666
If chk_al.Value = 1 Then
       StrCgdiag1 = StrCgdiag1 & ", " & "AL"
   
End If

 If chk_aml = 1 Then
 
      StrCgdiag1 = StrCgdiag1 & ", " & "AML"
  
End If
 If chk_all.Value = 1 Then
        StrCgdiag1 = StrCgdiag1 & ", " & "ALL"
  
End If
If chk_cml.Value = 1 Then
   
  
      StrCgdiag1 = StrCgdiag1 & ", " & "CML"
  
End If
 If chk_nhl.Value = 1 Then
   
      StrCgdiag1 = StrCgdiag1 & ", " & "NHL"
End If
If chk_lymphoma.Value = 1 Then
 
      StrCgdiag1 = StrCgdiag1 & ", " & "LYMPHOMA"
End If
If chk_fanane.Value = 1 Then
   
      StrCgdiag1 = StrCgdiag1 & ", " & "FANCONI'S ANEMIA"
   
End If
If chk_aplane.Value = 1 Then
   
  
     StrCgdiag1 = StrCgdiag1 & ", " & "APLASTIC ANEMIA"
End If
If chk_mds.Value = 1 Then
   
     StrCgdiag1 = StrCgdiag1 & ", " & "MDS"
End If
 If Left(StrCgdiag1, 1) = "," Then
    StrCgdiag1 = Mid(StrCgdiag1, 2, Len(Trim(StrCgdiag1)))
 End If
 If Left(StrCgdiag2, 1) = "," Then
    StrCgdiag2 = Mid(StrCgdiag2, 2, Len(Trim(StrCgdiag2)))
 End If
 
    
'''
pathology.show
Me.Hide
'pathology.bceltm.Visible = True
pathology.bceltm.Visible = True
End Sub

Private Sub cmd_save_Click()
Dim durill As String
Dim i As Integer
On Error GoTo MSG:
REQNOAUTO
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_cyreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(pathology.txt_time.Text)
INSRREQUISITN_QR(3) = Trim(pathology.msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(pathology.txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(pathology.txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(pathology.txt_status.Text)
INSRREQUISITN_QR(7) = Trim(pathology.txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(pathology.cmb_cyprocode.Text)
INSRREQUISITN_QR(9) = UCase(Trim(pathology.cmb_refdoc.Text))

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(pathology.msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(pathology.msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
    
    If pathology.Opt_patient.Value = True Then
        NO_CHARGE
      If No_Chargeflag = True Then
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
            No_Chargeflag = False
      Else
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
        
      End If
    End If
    
    If pathology.opt_staff.Value = True Then
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
    End If

    INSRREQUISITN_QR(14) = 0
    INSRREQUISITN_QR.Execute
    
    
    INSCYTOGCLDTA_QR(0) = req_chglen(Trim(txt_cyreqno.Text))
    INSCYTOGCLDTA_QR(1) = Trim(pathology.msk_caseno.FormattedText)
    INSCYTOGCLDTA_QR(2) = chk_at_dig.Value
    INSCYTOGCLDTA_QR(3) = chk_pr_treat.Value
    INSCYTOGCLDTA_QR(4) = chk_ps_indu.Value
    INSCYTOGCLDTA_QR(5) = chk_relapse.Value
    INSCYTOGCLDTA_QR(6) = chk_remission.Value
    INSCYTOGCLDTA_QR(7) = Trim(txt_followup.Text)
    INSCYTOGCLDTA_QR(8) = chk_postbmt.Value
    INSCYTOGCLDTA_QR(9) = Trim(cmb_sex.Text)
    INSCYTOGCLDTA_QR(10) = Trim(TXT_OTHERS.Text)
    INSCYTOGCLDTA_QR.Execute

        'inserting into drequisitn
    
        INSDREQUISITN_QR(0) = req_chglen(Trim(txt_cyreqno.Text))
        INSDREQUISITN_QR(1) = " "
        'INSDREQUISITN_QR(1) = Trim(txt_sppathno.Text)
        INSDREQUISITN_QR(2) = Trim(pathology.msk_caseno.FormattedText)
        INSDREQUISITN_QR(3) = Null
        INSDREQUISITN_QR(4) = Null
        INSDREQUISITN_QR(5) = Null
        INSDREQUISITN_QR(6) = Null
        INSDREQUISITN_QR(7) = Null
        INSDREQUISITN_QR(8) = Null
        INSDREQUISITN_QR(9) = Null
        INSDREQUISITN_QR(10) = Null
        INSDREQUISITN_QR(11) = Null
        INSDREQUISITN_QR(12) = Null
        INSDREQUISITN_QR(13) = Null
        INSDREQUISITN_QR(14) = Trim(pathology.cmb_refdoc.Text)
        INSDREQUISITN_QR(15) = Trim(pathology.cmb_reftodoc.Text)
        INSDREQUISITN_QR(16) = Trim(cmb_cyprocedure.Text)
        
        If Trim(pathology.txt_status.Text) = "IP" Then
            INSDREQUISITN_QR(17) = Trim(pathology.txt_ward.Text)
            INSDREQUISITN_QR(18) = Trim(pathology.txt_bedno.Text)
        Else
            INSDREQUISITN_QR(17) = Null
            INSDREQUISITN_QR(18) = Null
        End If
        INSDREQUISITN_QR(19) = " "
        INSDREQUISITN_QR(20) = Null
        INSDREQUISITN_QR(21) = Trim(ICU)
        INSDREQUISITN_QR.Execute
        
        '**** INSERTING INTO DREQUISITION

   MsgBox "Record saved successfully", vbInformation
Exit Sub
MSG:
MsgBox "Record not saved", vbInformation

End Sub

Private Sub Form_Load()
Dim str As String
Dim ICU As String
clear

'    STR = "INSERT INTO CYTOLIB.RREQUISITN (REQNNO, REQNDT,REQNTM,CASENO,DEPT_CODE,MODALITYCD,PATNSTAT,PAT_CATG,SER_CODE,REFBYDOC,PREVEXAM,RECPTNO,RECPTDT,REQNSTAT,RECORDNO) VALUES " _
'          & "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?)"
'    Set INSRREQUISITN_QR = rdb.CreateQuery("INSRREQUISITN_QRCY", STR)
'
'    STR = "UPDATE CYTOLIB.RREQUISITN REQNDT=?,REQNTM=?,CASENO=?,DEPT_CODE=?,MODALITYCD=?,PATNSTAT=?,PAT_CATG=?,SER_CODE=?,REFBYDOC=?,PREVEXAM=?,RECPTNO=?,RECPTDT=?,REQNSTAT=?,RECORDNO=? WHERE REQNNO=?"
'    Set updrrequisitn_qr = rdb.CreateQuery("UPDRREQUISITN_QRCY", STR)
'
'    STR = "INSERT INTO CYTOLIB.DREQUISITN(REQNO,PATHNO,CASENO,SITE,LATERALITS,DURILLNESS,METASTASIS,LATERALITM,OUTSSPECIM,OUTBLOCK,OUTSLIDES,FROZENSEC,FROZENDATE,IMMHISCHE,REFDOC,REFTODOC,OPPROCEDUR,WARD_NAME,BEDNO,PREPATHNO,NMETERIALS,ICDS2)VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
'    Set INSDREQUISITN_QR = rdb.CreateQuery("INSDREQUISITN_QRCY", STR)
'
'    STR = "UPDATE CYTOLIB.DREQUISITN SET REFDOC=?,REFTODOC=?,DURILLNESS=?,SITE=?,LATERALITs=?,OPPROCEDUR=?,NMETERIALS=?,METASTASIS=?,LATERALITM=?,OUTSLIDES=?,OUTSSPECIM=?,OUTBLOCK=?,FROZENSEC=?,FROZENDATE=?,LABNO=? WHERE REQNO=?"
'    Set upddrequisitn_qr = rdb.CreateQuery("UPDDREQUISITN1_QRCY", STR)
'
'    STR = "INSERT INTO CYTOLIB.CYTOGCLDTA (REQNO,CASEN0,AT_DIAG,PR_TREAT,PS_INDU,RELAPSE,REMISSION,FOLLUP,POSTBMT,B_SEX,SP_TYPE)" _
'         & "VALUES (?,?,?,?,?,?,?,?,?,?,?)"
'    Set INSCYTOGCLDTA_QR = rdb.CreateQuery("INSERTCYTGCL_QR", STR)
'
'    STR = "UPDATE CYTOLIB.CYTOGCLDTA CASEN0=?,AT_DIAG=?,PR_TREAT=?,PS_INDU=?,RELAPSE=?,REMISSIO=?,FOLLUP=?,POSTBMT=?,B_SEX=?,SP_TYPE=? WHERE REQNO=?"
'    Set UPDCYTOGCLDTA_QR = rdb.CreateQuery("UPDERTCYTGCL_QR", STR)
'
'    STR = "SELECT * FROM CYTOLIB.RREQUISITN WHERE CASENO=? ORDER BY MODALITYCD,REQNNO DESC"
'    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QRcy", STR)
'
'    STR = "SELECT * FROM CYTOLIB.MA101 WHERE SER_CODE=?"
'    Set SELMA101_QR = rdb.CreateQuery("SELMA101_QRrCY", STR)
'
'
'    '***** AUTOGENERATION OF THE REQUISITION NO.
'    STR = "SELECT * FROM CYTOLIB.DAUTOREQNO WHERE DEPT_CODE=? AND MODALITYCD=? AND REQNYEAR=?"
'    Set SELDAUTOREQNO_QR = rdb.CreateQuery("SELDAUTOREQNO_QRCY", STR)
'
'    STR = "INSERT INTO CYTOLIB.DAUTOREQNO (DEPT_CODE,MODALITYCD,REQNYEAR,REQNSRNO) VALUES(?,?,?,?)"
'    Set INSDAUTOREQNO_QR = rdb.CreateQuery("INSDAUTOREQNO_QRCY", STR)
'
'    STR = "UPDATE CYTOLIB.DAUTOREQNO SET REQNSRNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND REQNYEAR=?"
'    Set UPDDAUTOREQNO_QR = rdb.CreateQuery("UPDDAUTOREQNO_QRCY", STR)
'    ' ***** AUTOGENERATION OF THE REQUISTION NO. QURIES
'
'    STR = "SELECT * FROM DISLIB.DISICU WHERE IPADD=?"
'    Set SELDISICU_QR = rdb.CreateQuery("SELDISICU_STRcy", STR)
'
'    'cytology
'    STR = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=? AND SER_CODE<>'F106' ORDER BY SER_DET"
'    Set SELMA101CY_QR = rdb.CreateQuery("SELMA101CY_QRrr11cyc111", STR)
'
'    STR = "SELECT * FROM RISLIB.RINVESTIGN where modalitycd=?"
'    Set SELRINVESTCY_QR = rdb.CreateQuery("slerinevetcy11c11", STR)
'
'    '**** CYTOLOGY
'
'        If Trim(pathology.txt_modcode.Text) = "CC" Or Trim(pathology.txt_modcode.Text) = "FI" Then
'               If SELRINVESTCY_FLAG = False Then
'                SELRINVESTCY_QR(0) = Trim(pathology.txt_modcode.Text)
'                Set SELRINVESTCY_RS = SELRINVESTCY_QR.OpenResultset(1, 3)
'                SELRINVESTCY_FLAG = True
'            Else
'                SELRINVESTCY_QR(0) = Trim(pathology.txt_modcode.Text)
'                SELRINVESTCY_RS.MoreResults
'                SELRINVESTCY_RS.Requery
'            End If
'            If SELRINVESTCY_RS.RowCount > 0 Then
'                cmb_cyprocedure.clear
'                cmb_cyprocode.clear
'                While Not SELRINVESTCY_RS.EOF
'                    cmb_cyprocode.AddItem Trim(SELRINVESTCY_RS("ser_code"))
'                    'cmb_procedurecode.Text = Trim(SELRINVEST1_RS("ser_code"))
'                    If SELMA101CG_FLAG = False Then
'                        SELMA101CY_QR(0) = Trim(SELRINVESTCY_RS("SER_CODE"))
'                        Set SELMA101CY_RS = SELMA101CY_QR.OpenResultset(1, 3)
'                        SELMA101CY_FLAG = True
'                     Else
'                        SELMA101CY_QR(0) = SELRINVESTCY_RS("SER_CODE")
'                        SELMA101CY_RS.MoreResults
'                        SELMA101CY_RS.Requery
'                    End If
'                    If SELMA101CY_RS.RowCount > 0 Then
'                        While Not SELMA101CY_RS.EOF
'                        cmb_cyprocedure.AddItem SELMA101CY_RS("SER_DET")
'                        'cmb_procedure.Text SELMA101_RS1("SER_DET")
'                        SELMA101CY_RS.MoveNext
'                        Wend
'                    End If
'                    SELRINVESTCY_RS.MoveNext
'                Wend
'            End If
'        End If
'
'  STR = "select b.modlitydes from rislib.rrequisitn a,rislib.rmodality b  where a.modalitycd=b.modalitycd and a.reqnno=?"
'  Set selreqmoda_qr = rdb.CreateQuery("selmodareqc11", STR)
'
'  STR = "Select reqnno from rislib.rrequisitn where (dept_code='UZZ' or dept_code='FZZ' or dept_code='GZZ' or dept_code='VZZ') and caseno=?"
'  Set selreqnno_qr = rdb.CreateQuery("select_qrc11", STR)
'
'  selreqnno_qr(0) = pathology.Msk_caseno.FormattedText
'  Set selreqnno_rs = selreqnno_qr.OpenResultset(2, 3)
'
'  If selreqnno_rs.RowCount > 0 Then
'
'     selreqnno_rs.MoveFirst
'     While Not selreqnno_rs.EOF
'       selreqmoda_qr(0) = selreqnno_rs(0)
'       Set selreqmoda_rs = selreqmoda_qr.OpenResultset(2, 3)
'       cmb_cyspreinv.AddItem selreqnno_rs(0) & "   " & selreqmoda_rs(0)
'     selreqnno_rs.MoveNext
'     Wend
'  End If
        
        
'**** CYTOLOGY
    
    
'    SELDISICU_QR(0) = Trim(ipadd)
'    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
'
'    If Not SELDISICU_RS.EOF Then
'        lbl_icureq.Visible = True
'        ICU = "YES"
'    Else
'        lbl_icureq.Visible = False
'        ICU = "NO"
'    End If
End Sub

Public Sub NO_CHARGE()
'*****
On Error GoTo MSG:
No_Chargeflag = False
If SELMA101_FLAG = False Then
    SELMA101_QR(0) = Trim(pathology.cmb_cyprocode.Text)
    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
    SELMA101_FLAG = True
 Else
    SELMA101_QR(0) = Trim(pathology.cmb_cyprocode.Text)
    SELMA101_RS.MoreResults
    SELMA101_RS.Requery
End If

'SELMA101_QR(0) = Trim(cmb_procedurecode.Text)
'SELMA101_RS.MoreResults
'SELMA101_RS.Requery

If SELMA101_RS.RowCount > 0 Then
    If Trim(pathology.txt_catgory.Text) = "C" Then
        If SELMA101_RS("C_CH") = 0 Then
            No_Chargeflag = True
        End If
    ElseIf Trim(pathology.txt_catgory.Text) = "NC" Then
        If SELMA101_RS("NC_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
    If Trim(pathology.txt_catgory.Text) = "B" Then
        If SELMA101_RS("B_HOSP_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
    
Else
    MsgBox "There is NO  RECORD with this SERVICE CODE."
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Public Sub REQNOAUTO()
On Error GoTo MSG:
If SELDAUTOREQNOFLAG = False Then
    SELDAUTOREQNO_QR(0) = Trim(pathology.txt_dept.Text)
    SELDAUTOREQNO_QR(1) = Trim(pathology.txt_modcode.Text)
    SELDAUTOREQNO_QR(2) = Right(DATE11, 2)
    Set SELDAUTOREQNO_RS = SELDAUTOREQNO_QR.OpenResultset(1, 3)
    SELDAUTOREQNOFLAG = True
 Else
    SELDAUTOREQNO_QR(0) = Trim(pathology.txt_dept.Text)
    SELDAUTOREQNO_QR(1) = Trim(pathology.txt_modcode.Text)
    SELDAUTOREQNO_QR(2) = Right(DATE11, 2)
    SELDAUTOREQNO_RS.MoreResults
    SELDAUTOREQNO_RS.Requery
End If
If SELDAUTOREQNO_RS.RowCount > 0 Then
    AUTOREQNO = SELDAUTOREQNO_RS("REQNSRNO") + 1
    UPDDAUTOREQNO_QR(0) = AUTOREQNO
    UPDDAUTOREQNO_QR(1) = Trim(pathology.txt_dept.Text)
    UPDDAUTOREQNO_QR(2) = Trim(pathology.txt_modcode.Text)
    UPDDAUTOREQNO_QR(3) = Trim(Right(DATE11, 2))
    UPDDAUTOREQNO_QR.Execute
Else
    AUTOREQNO = 1
    INSDAUTOREQNO_QR(0) = Trim(pathology.txt_dept.Text)
    INSDAUTOREQNO_QR(1) = Trim(pathology.txt_modcode.Text)
    INSDAUTOREQNO_QR(2) = Trim(Right(DATE11, 2))
    INSDAUTOREQNO_QR(3) = 1
    INSDAUTOREQNO_QR.Execute
End If
'***********padding the reqslno with leading zeros"****************
Dim reqslno1
reqslno1 = AUTOREQNO
While Len(CStr(reqslno1)) < 6
reqslno1 = "0" & reqslno1
Wend
AUTOREQNO = reqslno1
'******************************************************************
DAUTOREQNO = Trim(pathology.txt_dept.Text) & "/" & Trim(pathology.txt_modcode.Text) & "/" & Trim(Right(DATE11, 2)) & "/" & AUTOREQNO
txt_cyreqno = DAUTOREQNO

'''If txt_modcode.Text = "SP" Then
'''    txt_cyreqno = DAUTOREQNO
'''    If DPATHYEARFLAG = False Then
'''        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
'''        Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
'''        DPATHYEARFLAG = True
'''    Else
'''        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
'''        DPATHYEAR_RS.MoreResults
'''        DPATHYEAR_RS.Requery
'''    End If
'''    If DPATHYEAR_RS.RowCount > 0 Then
'''        PATH = DPATHYEAR_RS("LASTTWO")
'''
'''    End If
'''    'MONO
'''    'txt_sppathno.Text = Right(DAUTOREQNO, 6) & "/" & Trim(PATH)
'''End If

'If txt_modcode.Text = "IH" Then  commented by candu 3/01/2002
'    txt_ihreqno.Text = DAUTOREQNO
'    If DPATHYEARFLAG = False Then
'        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
'        Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
'        DPATHYEARFLAG = True
'    Else
'        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
'        DPATHYEAR_RS.MoreResults
'        DPATHYEAR_RS.Requery
'    End If
'    If DPATHYEAR_RS.RowCount > 0 Then
'        PATH = DPATHYEAR_RS("LASTTWO")
'    End If
'    txt_ihpathno.Text = Right(DAUTOREQNO, 6) & "/" & Trim(PATH)
    
'End If

''If Trim(txt_modcode.Text) = "HT" Or Trim(txt_modcode.Text) = "SS" Or Trim(txt_modcode.Text) = "BP" Then
''    txt_bcreqno.Text = DAUTOREQNO
''End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Unload(cancel As Integer)
' Dim QR As rdoQuery
' Dim RS As rdoResultset
' For Each RS In rdb.rdoResultsets
'     RS.Close
' Next
'
' For Each QR In rdb.rdoQueries
'    QR.Close
' Next
' SELRINVESTCY_FLAG = False
' SELMA101CG_FLAG = False
End Sub

Private Sub opt_BMA_Click()
 TXT_OTHERS.Text = opt_BMA.Caption
End Sub

Private Sub opt_lymbnode_Click()
 TXT_OTHERS.Text = opt_lymbnode.Caption
End Sub

Private Sub opt_option_Click()
 TXT_OTHERS.Text = opt_option.Caption
End Sub

Private Sub opt_PBlood_Click()
 TXT_OTHERS.Text = opt_PBlood.Caption
End Sub

Public Sub clear()
 'cmb_cyprocedure.Text = ""
 'cmb_cyprocode.Text = ""
 cmb_sex.Text = ""
 cmb_cyspreinv.Text = ""
 txt_fax = ""
 txt_followup.Text = ""
 Option1.Value = True
 Option2.Value = False
 
 TXT_OTHERS.Text = ""
 txt_phone.Text = ""
 Dim obj As Object
 For Each obj In cytogeneticreq.Controls
    If TypeOf obj Is CheckBox Then
       obj.Value = False
    End If
 Next
 For Each obj In cytogeneticreq.Controls
    If TypeOf obj Is OptionButton Then
       obj.Value = False
    End If
 Next
 
End Sub
