VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form HEMBFREPORTOLD 
   Caption         =   "Body Fluid Report"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12945
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
   ScaleHeight     =   10905
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "DNA"
      Height          =   855
      Left            =   0
      TabIndex        =   193
      Top             =   8040
      Width           =   11655
      Begin VB.TextBox TXT_DNACLOIDY 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   3360
         TabIndex        =   195
         Top             =   240
         Width           =   8055
      End
      Begin VB.TextBox TXT_DNAINDEX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   194
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "CLOIDY"
         Height          =   255
         Left            =   2640
         TabIndex        =   197
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "INDEX"
         Height          =   255
         Left            =   120
         TabIndex        =   196
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   64
      Top             =   3720
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Body Fluid Details"
      TabPicture(0)   =   "HEMBFREPORT.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Surface Markers"
      TabPicture(1)   =   "HEMBFREPORT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   3735
         Left            =   240
         TabIndex        =   65
         Top             =   240
         Width           =   11655
         Begin VB.TextBox txt_nmaterial 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   68
            Top             =   240
            Width           =   5895
         End
         Begin VB.TextBox txt_cytomorphology 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   2520
            ScrollBars      =   3  'Both
            TabIndex        =   5
            Top             =   600
            Width           =   9015
         End
         Begin VB.TextBox txt_cytochemistry 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   6
            Top             =   1080
            Width           =   9015
         End
         Begin VB.TextBox txt_impression 
            Appearance      =   0  'Flat
            Height          =   645
            Left            =   2520
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   7
            Top             =   1440
            Width           =   9015
         End
         Begin VB.TextBox txt_cellcount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9720
            TabIndex        =   69
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txt_addendum 
            Appearance      =   0  'Flat
            Height          =   765
            Left            =   2520
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   67
            Top             =   2160
            Width           =   5655
         End
         Begin VB.ComboBox cmb_addendumby 
            Height          =   315
            Left            =   9720
            TabIndex        =   66
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txt_surfacemarkers 
            Appearance      =   0  'Flat
            Height          =   645
            Left            =   2520
            MaxLength       =   2000
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   3000
            Width           =   9015
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1560
            TabIndex        =   184
            Top             =   3000
            Width           =   870
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nature of Material"
            Height          =   195
            Left            =   840
            TabIndex        =   77
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cytomorphology"
            Height          =   195
            Left            =   1080
            TabIndex        =   76
            Top             =   600
            Width           =   1350
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cytochemistry"
            Height          =   195
            Left            =   1200
            TabIndex        =   75
            Top             =   1080
            Width           =   1305
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Impression"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1440
            TabIndex        =   74
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cell Count"
            Height          =   195
            Left            =   8520
            TabIndex        =   73
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Surface Markers"
            Height          =   195
            Left            =   720
            TabIndex        =   72
            Top             =   1920
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Addendum"
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   1440
            TabIndex        =   71
            Top             =   2280
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "Addendum by"
            Height          =   255
            Left            =   8400
            TabIndex        =   70
            Top             =   2160
            Width           =   1215
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   78
         Top             =   480
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   6588
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Screen I"
         TabPicture(0)   =   "HEMBFREPORT.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame9"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Screen II"
         TabPicture(1)   =   "HEMBFREPORT.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame10"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Screen III"
         TabPicture(2)   =   "HEMBFREPORT.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame11"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame9 
            ForeColor       =   &H00800000&
            Height          =   2295
            Left            =   120
            TabIndex        =   121
            Top             =   360
            Width           =   11055
            Begin VB.TextBox txt_CD15 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   142
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txt_cd14 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   141
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txt_cd13 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   140
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   139
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   138
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   137
               Top             =   600
               Width           =   1440
            End
            Begin VB.TextBox TXT_CD4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   136
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   135
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD8 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   134
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD10 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   8880
               TabIndex        =   133
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD11C 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   132
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txt_cd36 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   131
               Top             =   1320
               Width           =   1440
            End
            Begin VB.TextBox TXT_CD19 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3585
               TabIndex        =   130
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD22 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   129
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txt_cytocd221 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   1080
               TabIndex        =   128
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD23 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   127
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD25 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   126
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox txt_CD33 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   125
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TXT_CYTOCD3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   124
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_cytocd22 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   8880
               TabIndex        =   123
               Top             =   1320
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD20 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   122
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label lbl_cd15 
               AutoSize        =   -1  'True
               Caption         =   "CD 15"
               Height          =   195
               Left            =   8280
               TabIndex        =   163
               Top             =   960
               Width           =   540
            End
            Begin VB.Label lbl_cd14 
               AutoSize        =   -1  'True
               Caption         =   "CD 14"
               Height          =   195
               Left            =   5640
               TabIndex        =   162
               Top             =   960
               Width           =   540
            End
            Begin VB.Label LBL_CD13 
               AutoSize        =   -1  'True
               Caption         =   "CD 13"
               Height          =   195
               Left            =   3000
               TabIndex        =   161
               Top             =   960
               Width           =   540
            End
            Begin VB.Label LBL_CD2 
               AutoSize        =   -1  'True
               Caption         =   "CD 2"
               Height          =   195
               Left            =   480
               TabIndex        =   160
               Top             =   240
               Width           =   555
            End
            Begin VB.Label LBL_CD3 
               AutoSize        =   -1  'True
               Caption         =   "CD 3"
               Height          =   195
               Left            =   3000
               TabIndex        =   159
               Top             =   240
               Width           =   435
            End
            Begin VB.Label LBL_CD5 
               AutoSize        =   -1  'True
               Caption         =   "CD 5"
               Height          =   195
               Left            =   480
               TabIndex        =   158
               Top             =   600
               Width           =   555
            End
            Begin VB.Label LBL_CD4 
               AutoSize        =   -1  'True
               Caption         =   "CD 4"
               Height          =   195
               Left            =   8280
               TabIndex        =   157
               Top             =   240
               Width           =   435
            End
            Begin VB.Label LBL_CD7 
               AutoSize        =   -1  'True
               Caption         =   "CD 7"
               Height          =   195
               Left            =   3000
               TabIndex        =   156
               Top             =   600
               Width           =   435
            End
            Begin VB.Label LBL_CD8 
               AutoSize        =   -1  'True
               Caption         =   "CD 8"
               Height          =   195
               Left            =   5640
               TabIndex        =   155
               Top             =   600
               Width           =   435
            End
            Begin VB.Label LBL_CD10 
               AutoSize        =   -1  'True
               Caption         =   "CD 10"
               Height          =   195
               Left            =   8280
               TabIndex        =   154
               Top             =   600
               Width           =   540
            End
            Begin VB.Label LBL_CD11C 
               AutoSize        =   -1  'True
               Caption         =   "CD11c"
               Height          =   195
               Left            =   480
               TabIndex        =   153
               Top             =   960
               Width           =   705
            End
            Begin VB.Label LBL_CD36 
               Caption         =   "CD16"
               Height          =   255
               Left            =   480
               TabIndex        =   152
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label LBL_CD19 
               AutoSize        =   -1  'True
               Caption         =   "CD 19"
               Height          =   195
               Left            =   3000
               TabIndex        =   151
               Top             =   1320
               Width           =   540
            End
            Begin VB.Label LBL_CD22 
               AutoSize        =   -1  'True
               Caption         =   "CD 22"
               Height          =   195
               Left            =   8280
               TabIndex        =   150
               Top             =   1320
               Width           =   540
            End
            Begin VB.Label lbl_cyto221 
               AutoSize        =   -1  'True
               Caption         =   "Cyto CD22"
               Height          =   195
               Left            =   120
               TabIndex        =   149
               Top             =   1680
               Width           =   1035
            End
            Begin VB.Label LBL_CD23 
               AutoSize        =   -1  'True
               Caption         =   "CD 23"
               Height          =   195
               Left            =   3000
               TabIndex        =   148
               Top             =   1680
               Width           =   540
            End
            Begin VB.Label LBL_CD25 
               AutoSize        =   -1  'True
               Caption         =   "CD 25"
               Height          =   195
               Left            =   5640
               TabIndex        =   147
               Top             =   1680
               Width           =   540
            End
            Begin VB.Label LBL_CD33 
               AutoSize        =   -1  'True
               Caption         =   "CD 33"
               Height          =   195
               Left            =   8280
               TabIndex        =   146
               Top             =   1680
               Width           =   540
            End
            Begin VB.Label LBL_CYTOCD3 
               AutoSize        =   -1  'True
               Caption         =   "Cyto CD3"
               Height          =   195
               Left            =   5280
               TabIndex        =   145
               Top             =   240
               Width           =   810
            End
            Begin VB.Label LBL_CYTOCD22 
               AutoSize        =   -1  'True
               Caption         =   "CD22"
               Height          =   195
               Left            =   8280
               TabIndex        =   144
               Top             =   1320
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label Label16 
               Caption         =   "CD 20"
               Height          =   255
               Left            =   5640
               TabIndex        =   143
               Top             =   1320
               Width           =   735
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00800000&
            Height          =   2175
            Left            =   -74865
            TabIndex        =   82
            Top             =   360
            Width           =   11055
            Begin VB.TextBox txt_lanbda 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   168
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TXT_CYTOMU 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   96
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox txt_cytocd79a 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   90
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox txt_cd103 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   92
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txt_CD79B 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   91
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TXT_FMC7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   98
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox TXT_DBA44 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   97
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox txt_cd41 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   85
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD42 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   86
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD56 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   87
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD68 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   89
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD38 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   84
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox TXT_CD117 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   93
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD138 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   94
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox TXT_ANTIMPO 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   95
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox txt_Glycophorina 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   99
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox TXT_HLADR 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   101
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox txt_kappa 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   103
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD34 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   83
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD61 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   88
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label lbl_lanbda 
               AutoSize        =   -1  'True
               Caption         =   "Lambda"
               Height          =   195
               Left            =   8535
               TabIndex        =   179
               Top             =   1680
               Width           =   675
            End
            Begin VB.Label LBL_CYTOMU 
               AutoSize        =   -1  'True
               Caption         =   "Cyto mu"
               Height          =   195
               Left            =   3090
               TabIndex        =   120
               Top             =   1320
               Width           =   690
            End
            Begin VB.Label lbl_cytocd79a 
               AutoSize        =   -1  'True
               Caption         =   "CytoCD79a"
               Height          =   195
               Left            =   8445
               TabIndex        =   119
               Top             =   600
               Width           =   960
            End
            Begin VB.Label lbl_cd103 
               AutoSize        =   -1  'True
               Caption         =   "CD103"
               Height          =   195
               Left            =   3210
               TabIndex        =   118
               Top             =   960
               Width           =   585
            End
            Begin VB.Label lbl_DBA44 
               AutoSize        =   -1  'True
               Caption         =   "DBA 44"
               Height          =   195
               Left            =   6000
               TabIndex        =   117
               Top             =   1320
               Width           =   660
            End
            Begin VB.Label LBL_CD79B 
               AutoSize        =   -1  'True
               Caption         =   "CD79 b"
               Height          =   195
               Left            =   600
               TabIndex        =   116
               Top             =   960
               Width           =   645
            End
            Begin VB.Label LBL_FMC7 
               AutoSize        =   -1  'True
               Caption         =   "FMC7"
               Height          =   195
               Left            =   8880
               TabIndex        =   115
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label LBL_CD41 
               AutoSize        =   -1  'True
               Caption         =   "CD41"
               Height          =   195
               Left            =   6120
               TabIndex        =   114
               Top             =   240
               Width           =   480
            End
            Begin VB.Label Label15 
               Caption         =   "CD42"
               Height          =   255
               Left            =   8880
               TabIndex        =   113
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label14 
               Caption         =   "CD56"
               Height          =   255
               Left            =   720
               TabIndex        =   112
               Top             =   600
               Width           =   495
            End
            Begin VB.Label LBL_CD68 
               AutoSize        =   -1  'True
               Caption         =   "Cyto CD68"
               Height          =   195
               Left            =   5760
               TabIndex        =   111
               Top             =   600
               Width           =   915
            End
            Begin VB.Label LBL_CD38 
               AutoSize        =   -1  'True
               Caption         =   "CD 38"
               Height          =   195
               Left            =   3240
               TabIndex        =   110
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label13 
               Caption         =   "CD117"
               Height          =   255
               Left            =   6000
               TabIndex        =   109
               Top             =   960
               Width           =   735
            End
            Begin VB.Label LBL_CD138 
               AutoSize        =   -1  'True
               Caption         =   " CD138"
               Height          =   195
               Left            =   8760
               TabIndex        =   108
               Top             =   960
               Width           =   645
            End
            Begin VB.Label LBL_ANTIMPO 
               AutoSize        =   -1  'True
               Caption         =   "AntiMPO"
               Height          =   195
               Left            =   480
               TabIndex        =   107
               Top             =   1320
               Width           =   750
            End
            Begin VB.Label LBL_GLYCOPHORIN 
               AutoSize        =   -1  'True
               Caption         =   "Glycophorin A"
               Height          =   195
               Left            =   120
               TabIndex        =   106
               Top             =   1680
               Width           =   1200
            End
            Begin VB.Label LBL_HALDR 
               AutoSize        =   -1  'True
               Caption         =   "HLA DR"
               Height          =   195
               Left            =   3120
               TabIndex        =   105
               Top             =   1680
               Width           =   705
            End
            Begin VB.Label lbl_kappa 
               AutoSize        =   -1  'True
               Caption         =   "Kappa"
               Height          =   195
               Left            =   6000
               TabIndex        =   104
               Top             =   1680
               Width           =   555
            End
            Begin VB.Label lbl_cd34 
               AutoSize        =   -1  'True
               Caption         =   "CD 34"
               Height          =   195
               Left            =   720
               TabIndex        =   102
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label12 
               Caption         =   "CD 61"
               Height          =   255
               Left            =   3240
               TabIndex        =   100
               Top             =   600
               Width           =   615
            End
         End
         Begin VB.Frame Frame11 
            Height          =   3255
            Left            =   -74640
            TabIndex        =   79
            Top             =   360
            Width           =   10815
            Begin VB.TextBox txt_clnd1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8160
               TabIndex        =   172
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton opt_lsp 
               Caption         =   "Light scatter plots."
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   175
               Top             =   1680
               Width           =   2415
            End
            Begin VB.OptionButton opt_ssp 
               Caption         =   "Side scatter/CD 45 plots."
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   174
               Top             =   1320
               Width           =   3015
            End
            Begin VB.TextBox TXT_SMPCTG 
               Height          =   285
               Left            =   4680
               TabIndex        =   173
               Top             =   1320
               Width           =   615
            End
            Begin VB.Frame Frame14 
               Height          =   615
               Left            =   2640
               TabIndex        =   188
               Top             =   1920
               Width           =   7215
               Begin VB.OptionButton opt_drak 
                  Caption         =   "Dr. Ashok Kumar"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   176
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.OptionButton opt_mryb 
                  Caption         =   "Mr. Y. Badrinath"
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   177
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.OptionButton opt_sitaram 
                  Caption         =   "Mr. Sitaram"
                  Height          =   255
                  Left            =   4800
                  TabIndex        =   178
                  Top             =   240
                  Width           =   1455
               End
            End
            Begin VB.TextBox txt_tcrag 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5280
               TabIndex        =   171
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txt_tcrab 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               TabIndex        =   170
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox Txt_IgM 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2775
               TabIndex        =   165
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TXT_TDT 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5280
               TabIndex        =   167
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_smig 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8160
               TabIndex        =   169
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label26 
               Caption         =   "Cyclin D1"
               Height          =   255
               Left            =   7200
               TabIndex        =   192
               Top             =   720
               Width           =   855
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H000000C0&
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   2
               Height          =   1815
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   1200
               Width           =   10575
            End
            Begin VB.Label Label25 
               Caption         =   "The above positive percentage were determined on           % of the gated events on"
               Height          =   255
               Left            =   240
               TabIndex        =   191
               Top             =   1320
               Width           =   7215
            End
            Begin VB.Label Label24 
               Caption         =   "Flow Cytometry is Done By "
               Height          =   255
               Left            =   240
               TabIndex        =   190
               Top             =   2160
               Width           =   2295
            End
            Begin VB.Label Label23 
               Caption         =   "Instrument: BD-FACS Calibur E 2237"
               Height          =   255
               Left            =   240
               TabIndex        =   189
               Top             =   2640
               Width           =   3615
            End
            Begin VB.Label Label21 
               Caption         =   "ad"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   9
                  Charset         =   2
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4920
               TabIndex        =   183
               Top             =   720
               Width           =   255
            End
            Begin VB.Label Label20 
               Caption         =   "TCR"
               Height          =   255
               Left            =   4440
               TabIndex        =   182
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label3 
               Caption         =   "ab"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   9
                  Charset         =   2
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2280
               TabIndex        =   181
               Top             =   720
               Width           =   255
            End
            Begin VB.Label lbl_tmr 
               Caption         =   "TCR"
               Height          =   255
               Left            =   1800
               TabIndex        =   180
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label17 
               Caption         =   "IgM"
               Height          =   255
               Left            =   2310
               TabIndex        =   166
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lbl_tdt 
               AutoSize        =   -1  'True
               Caption         =   "Tdt"
               Height          =   195
               Left            =   4800
               TabIndex        =   81
               Top             =   240
               Width           =   300
            End
            Begin VB.Label LBL_SMIG 
               AutoSize        =   -1  'True
               Caption         =   "SmIg"
               Height          =   195
               Left            =   7560
               TabIndex        =   80
               Top             =   240
               Width           =   435
            End
         End
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   0
      TabIndex        =   52
      Top             =   9840
      Width           =   11775
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         MouseIcon       =   "HEMBFREPORT.frx":008C
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MouseIcon       =   "HEMBFREPORT.frx":0396
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "HEMBFREPORT.frx":06A0
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "HEMBFREPORT.frx":09AA
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5880
         MouseIcon       =   "HEMBFREPORT.frx":0CB4
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8760
         MouseIcon       =   "HEMBFREPORT.frx":0FBE
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10440
         MouseIcon       =   "HEMBFREPORT.frx":12C8
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7200
         MouseIcon       =   "HEMBFREPORT.frx":15D2
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10680
      TabIndex        =   48
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10680
      TabIndex        =   47
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame common 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   11775
      Begin VB.CommandButton cmd_emr 
         BackColor       =   &H00FFFF80&
         Caption         =   "&EMR"
         Height          =   375
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   2160
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   27
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   0
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   25
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2415
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
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   46
         Top             =   2520
         Width           =   1830
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   44
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   43
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   42
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   39
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   38
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   37
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   36
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   35
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Hem No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   34
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   33
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   32
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   1830
      End
   End
   Begin VB.Frame Frame12 
      Height          =   735
      Left            =   0
      TabIndex        =   60
      Top             =   9000
      Width           =   11775
      Begin VB.ComboBox CMB_PROVISIONAL 
         Height          =   315
         Left            =   4800
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         Left            =   9360
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmb_enterby 
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registrar"
         Height          =   195
         Left            =   3480
         TabIndex        =   63
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
         Height          =   195
         Left            =   7680
         TabIndex        =   62
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl_enter 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Label Label19 
      Caption         =   "Flow Cytometry is Done By "
      Height          =   255
      Left            =   840
      TabIndex        =   187
      Top             =   11040
      Width           =   2295
   End
   Begin VB.Label Label22 
      Caption         =   "Instrument: BD-FACS Calibur E 2237"
      Height          =   255
      Left            =   840
      TabIndex        =   186
      Top             =   11520
      Width           =   3615
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   164
      Top             =   0
      Width           =   9855
   End
   Begin VB.Label LBL_CP 
      AutoSize        =   -1  'True
      Caption         =   "BODY FLUID REPORT"
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
      Left            =   4080
      TabIndex        =   51
      Top             =   480
      Width           =   3060
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10080
      TabIndex        =   50
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10080
      TabIndex        =   49
      Top             =   480
      Width           =   420
   End
End
Attribute VB_Name = "HEMBFREPORTOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim MODCD As String
Dim testorder As String
Dim sel_deptqr As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SELDEPTRSFLAG1 As Boolean
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim SEL_MODALITYFLAG1 As Boolean
Dim sel_casno As rdoQuery
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim REQFLAG1 As Boolean
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim selta101flag1 As Boolean
Dim repdr_qr As rdoQuery
Dim repdr_rs As rdoResultset
Dim REPDRFLAG As Boolean
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim REFDRFLAG1 As Boolean
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset
Dim SELMA104FLAG1 As Boolean
Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset
Dim SELPATHDREQUISITNFALG1 As Boolean
Dim selma302_qr As rdoQuery
Dim selma302_rs As rdoResultset
Dim SELMA302FLAG1 As Boolean
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELTA1012FLAG1 As Boolean
'Dim SELEMPLIB_QR As rdoQuery
'Dim SELEMPLIB_RS As rdoResultset
Dim i As Integer
'Dim insdfsrep_qr As rdoQuery
'Dim insdgrossdesc_qr As rdoQuery
'Dim insdfs_qr As rdoQuery
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean
'Dim selempbasinfo1_qr As rdoQuery
'Dim selempbasinfo1_rs As rdoResultset
Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim INSDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUMFLAG As Boolean
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG1 As Boolean

'***** Shraddha 21/07/10 *****
Dim SELDREQDRNAME_QR As rdoQuery
Dim SELDREQDRNAME_RS As rdoResultset

Dim UPDDREQDRNAME_QR As rdoQuery
'*****************************


Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer
Dim ADDENDUMFALG As Boolean

Dim seldissecentr1_qr As rdoQuery
Dim seldissecentr1_rs As rdoResultset
Dim seldissecentrflag1 As Boolean

Dim seldissecprov1_qr As rdoQuery
Dim seldissecprov1_rs As rdoResultset
Dim seldissecprovflag1 As Boolean

Dim seldissecfin1_qr As rdoQuery
Dim seldissecfin1_rs As rdoResultset
Dim seldissecfinflag1 As Boolean

Dim SELDREQSTAFFDQR As rdoQuery
Dim SELDREQSTAFFDRS As rdoResultset
Dim SELDREQSTAFFD1FLAG As Boolean
Dim upddreqnoqr As rdoQuery

Dim selMA101QR As rdoQuery
Dim SELMA101RS As rdoResultset
Dim SELMA101FLAG As Boolean



'Dim updrrequisitn1_qr As rdoQuery
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean


'Dim w_pwave As Application, w_rhythm As Application, w_stsegment As Application, w_twave As Application, w_impression As Application
'Dim d_pwave As Document, d_rhythm As Document, d_stsegment As Document, d_twave As Document, d_impression As Document

Dim INSDHBFREPORT_QR As rdoQuery
Dim UPDDHBFREPORT_QR As rdoQuery
Dim SELDHBFREPORT_QR As rdoQuery
Dim SELDHBFREPORT_RS As rdoResultset
Dim SELDHBFREPORTFLAG As Boolean

Dim RNAGE As String
Dim SS As String

''******SECURITY***********

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String

Dim REPWTRNRS As rdoResultset
Dim REPWTRN_QR As rdoQuery

Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset
''******added by shweta 10-02-2014
Dim SELDCBCREPA_QR As rdoQuery
Dim SELDCBCREPA_RS As rdoResultset
Dim INSDCBCREPA_QR As rdoQuery
Dim UPDDCBCREPA_QR As rdoQuery
'''till here
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
    If SEL_MODALITYFLAG1 = False Then
        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
        Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
        SEL_MODALITYFLAG1 = True
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
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "UZZ" Then
            cmb_dept.ListIndex = ind_cbodept
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
MODCD = Right(cmb_modality.Text, (l - moda))
Exit Sub
MSG:
MsgBox Error(Err)
End Sub


Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "BODY FLUID REPORT$BF" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Public Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
Dim entr As String
On Error GoTo MSG:
If Opt_patient.Value = True Then
        If NEWFLAG = True Then
            NEWFLAG = False
            ENABL ("ENA")
            If selta101flag1 = False Then
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101flag1 = True
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
            
            If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
            End If
'            If SELMA104FLAG1 = False Then
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'                SELMA104FLAG1 = True
'            Else
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                selma104_rs.MoreResults
'                selma104_rs.Requery
'            End If
           If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
          
            If REFDR_RS.RowCount > 0 Then
                If REFDR_RS("PATNSTAT") = "IP" Then
                    txt_ward.Visible = True
                    txt_bedno.Visible = True
                    lbl_ward.Visible = True
                    lbl_bed.Visible = True
                    txt_ward.Text = selpathdrequisitn_rs("WARD_NAME")
                    txt_bedno.Text = selpathdrequisitn_rs("bedno")
                    txt_status.Text = "IP"
                Else
                    txt_status.Text = "OP"
                    txt_ward.Visible = False
                    txt_bedno.Visible = False
                    lbl_ward.Visible = False
                    lbl_bed.Visible = False
                End If
           End If
            
            If selpathdrequisitn_rs.RowCount > 0 Then
                TXT_PATH.Text = "" & Trim(selpathdrequisitn_rs("PATHNO"))
                txt_nmaterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
                If selpathdrequisitn_rs.EOF Then
                    txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
                'txt_refdoc.Text = SELPATHDREQUISITN_RS("REFDOC")
       
            End If
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
           
            
            
            

'        If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'
'            entr = cmb_enterby.Text
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'            seldissecentr1_rs.MovePrevious
'            cmb_enterby.Text = seldissecentr1_rs("CCNAME")
'
'        End If

'        If seldissecfinflag = False Then
'            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'            seldissecfinflag1 = True
'        Else
'            seldissecfin1_rs.MoreResults
'            seldissecfin1_rs.Requery
'        End If
'        If seldissecfin1_rs.RowCount > 0 Then
'            entr = ""
'            entr = cmb_finalby.Text
'            cmb_finalby.clear
'            While Not seldissecfin1_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                seldissecfin1_rs.MoveNext
'            Wend
'            cmb_finalby.Text = entr
'            entr = ""
'        End If
'        If seldissecprovflag1 = False Then
'            Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
'            seldissecprovflag1 = True
'        Else
'            seldissecprov1_rs.MoreResults
'            seldissecprov1_rs.Requery
'        End If
'        If seldissecprov1_rs.RowCount > 0 Then
'            entr = ""
'            entr = CMB_PROVISIONAL.Text
'            CMB_PROVISIONAL.clear
'            While Not seldissecprov1_rs.EOF
'                CMB_PROVISIONAL.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                seldissecprov1_rs.MoveNext
'            Wend
'            CMB_PROVISIONAL.Text = entr
'            entr = ""
'        End If
            
            
            
End If
If viewflag = True Then
     
            viewflag = False
            If selta101flag1 = False Then
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101flag1 = True
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
            
            If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Report is commited , You can not Modify"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                CMB_PROVISIONAL.Text = REFDR_RS("PREPORTBY")
                 If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
                End If
                
            End If
            
'            If SELMA104FLAG1 = False Then
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'                SELMA104FLAG1 = True
'            Else
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                selma104_rs.MoreResults
'                selma104_rs.Requery
'            End If
           If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
          
            If REFDR_RS.RowCount > 0 Then
                If REFDR_RS("PATNSTAT") = "IP" Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_ward.Text = selpathdrequisitn_rs("WARD_NAME")
                txt_bedno.Text = selpathdrequisitn_rs("bedno")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_status.Text = "OP"
            End If
            End If
         
           
            If selpathdrequisitn_rs.RowCount > 0 Then
                TXT_PATH.Text = "" & Trim(selpathdrequisitn_rs("PATHNO"))
                txt_nmaterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
                If selpathdrequisitn_rs.EOF = False Then
                    'txt_relclnhis.Visible = True
                    
                    'txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
            End If
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
            If SELDHBFREPORTFLAG = False Then
                SELDHBFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDHBFREPORT_RS = SELDHBFREPORT_QR.OpenResultset(1, 3)
                SELDHBFREPORTFLAG = True
            Else
                  SELDHBFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                  SELDHBFREPORT_RS.MoreResults
                  SELDHBFREPORT_RS.Requery
            End If
            ENABL ("DIS")
            If SELDHBFREPORT_RS.RowCount > 0 Then
                txt_cellcount.Text = "" & SELDHBFREPORT_RS("CELLCOUNT")
                txt_cytomorphology.Text = "" & SELDHBFREPORT_RS("CYTOMORPH")
                txt_cytochemistry.Text = "" & SELDHBFREPORT_RS("CYTOCHEMIS")
                txt_surfacemarkers.Text = "" & SELDHBFREPORT_RS("SURMARKER")
                txt_impression.Text = "" & SELDHBFREPORT_RS("IMPRESSION")
                ' *****************SURFACE MARKER
                'sundeep
                    txt_cd13.Text = " " & Trim(SELDHBFREPORT_RS("cd13"))
                    txt_cd14.Text = " " & Trim(SELDHBFREPORT_RS("CD14"))
                    txt_CD15.Text = " " & Trim(SELDHBFREPORT_RS("CD15"))
                    txt_CD33.Text = " " & Trim(SELDHBFREPORT_RS("CD33"))
                    txt_cd36.Text = " " & Trim(SELDHBFREPORT_RS("CD36"))
                    txt_cd41.Text = " " & Trim(SELDHBFREPORT_RS("CD4142"))
                    TXT_CD42.Text = "" & Trim(SELDHBFREPORT_RS("CD42"))
                    TXT_CD117.Text = "" & Trim(SELDHBFREPORT_RS("CD117"))
                    TXT_CD20.Text = "" & Trim(SELDHBFREPORT_RS("CD20"))
                    TXT_CD61.Text = "" & Trim(SELDHBFREPORT_RS("CD61"))
                    TXT_CD56.Text = "" & Trim(SELDHBFREPORT_RS("CD56"))
                    txt_Glycophorina.Text = " " & Trim(SELDHBFREPORT_RS("GLYCOPHOIN"))
                    TXT_CD68.Text = " " & Trim(SELDHBFREPORT_RS("CD68"))
                    TXT_CD34.Text = " " & Trim(SELDHBFREPORT_RS("CD34"))
                    TXT_HLADR.Text = " " & Trim(SELDHBFREPORT_RS("HLADR"))
                    TXT_CD38.Text = " " & Trim(SELDHBFREPORT_RS("CD38"))
                    TXT_CD138.Text = " " & Trim(SELDHBFREPORT_RS("CD138"))
                    TXT_TDT.Text = " " & Trim(SELDHBFREPORT_RS("TDT"))
                    TXT_ANTIMPO.Text = " " & Trim(SELDHBFREPORT_RS("ANTIMPO"))
                    TXT_CD10.Text = " " & Trim(SELDHBFREPORT_RS("CD10"))
                    TXT_CD19.Text = " " & Trim(SELDHBFREPORT_RS("CD19"))
                    TXT_FMC7.Text = " " & Trim(SELDHBFREPORT_RS("FMC7"))
                    TXT_CD22.Text = " " & Trim(SELDHBFREPORT_RS("CD22"))
                    TXT_CD23.Text = " " & Trim(SELDHBFREPORT_RS("CD23"))
                    txt_smig.Text = " " & Trim(SELDHBFREPORT_RS("SMIG"))
                    txt_CD79B.Text = " " & Trim(SELDHBFREPORT_RS("CD79B"))
                    TXT_CD25.Text = " " & Trim(SELDHBFREPORT_RS("CD25"))
                    TXT_CD11C.Text = " " & Trim(SELDHBFREPORT_RS("CD11C"))
                    txt_cd103.Text = " " & Trim(SELDHBFREPORT_RS("CD103"))
                    txt_kappa.Text = " " & Trim(SELDHBFREPORT_RS("KAPPA"))
                    txt_lanbda.Text = " " & Trim(SELDHBFREPORT_RS("LAMDA"))
                    TXT_DBA44.Text = " " & Trim(SELDHBFREPORT_RS("DBA44"))
                    txt_cytocd221(1).Text = " " & Trim(SELDHBFREPORT_RS("CYTOEP22"))
                    txt_cytocd79a.Text = " " & Trim(SELDHBFREPORT_RS("CYTOCD79A"))
                    TXT_CYTOMU.Text = " " & Trim(SELDHBFREPORT_RS("CYTOMU"))
                    TXT_CD2.Text = " " & Trim(SELDHBFREPORT_RS("CD2"))
                    TXT_CD3.Text = " " & Trim(SELDHBFREPORT_RS("CD3"))
                    TXT_CD4.Text = " " & Trim(SELDHBFREPORT_RS("CD4"))
                    TXT_CD5.Text = " " & Trim(SELDHBFREPORT_RS("CD5"))
                    TXT_CD7.Text = " " & Trim(SELDHBFREPORT_RS("CD7"))
                    TXT_CD8.Text = " " & Trim(SELDHBFREPORT_RS("CD8"))
                    TXT_CYTOCD3.Text = " " & Trim(SELDHBFREPORT_RS("CYTOCD3"))
                    txt_igm.Text = " " & Trim(SELDHBFREPORT_RS("IGM"))
                    
                    '***** ADDED BY Shraddha on 21/07/10 *****
                    txt_tcrab.Text = " " & Trim(SELDHBFREPORT_RS("TCRAB"))
                    txt_tcrag.Text = " " & Trim(SELDHBFREPORT_RS("TCRGD"))
                    txt_clnd1.Text = " " & Trim(SELDHBFREPORT_RS("CLND1"))
                    TXT_SMPCTG.Text = " " & Trim(SELDHBFREPORT_RS("SMPCTG"))
                    If SELDHBFREPORT_RS("OPTLSP") = "Y" Then
                        opt_lsp.Value = True
                    Else
                        opt_lsp.Value = False
                    End If
                    If SELDHBFREPORT_RS("OPTSCP") = "Y" Then
                        opt_ssp.Value = True
                    Else
                        opt_ssp.Value = False
                    End If
                    
                    
                    SELDREQDRNAME_QR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDREQDRNAME_RS = SELDREQDRNAME_QR.OpenResultset(1, 3)
                    If SELDREQDRNAME_RS.RowCount > 0 Then
                        If Trim(SELDREQDRNAME_RS("LATERALITS")) = "Dr. Ashok Kumar" Then
                            opt_drak.Value = True
                            opt_mryb.Value = False
                            opt_sitaram.Value = False
                        ElseIf Trim(SELDREQDRNAME_RS("METASTASIS")) = "Mr. Y. Badrinath" Then
                            opt_drak.Value = False
                            opt_mryb.Value = True
                            opt_sitaram.Value = False
                            '***** Added by Shradha on 05/05/10 *****
                        ElseIf Trim(SELDREQDRNAME_RS("METASTASIS")) = "Mr. Sitaram" Then
                            opt_drak.Value = False
                            opt_mryb.Value = False
                            opt_sitaram.Value = True
                        End If
                    End If
                    '*******************************************
                    
                    
               'sundeep
                ' *****************SURFACE MARKER
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
                ADDENDUM = ""
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = ADDENDUM & " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
      
    End If
End If
If opt_staff.Value = True Then
    STAFF
    

End If

VIEWDCBCREPA


Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub VIEWDCBCREPA()
           
    If SELDCBCFLAG = False Then
        SELDCBCREPA_QR(0) = req_chglen(cmb_reqno.Text)
        Set SELDCBCREPA_RS = SELDCBCREPA_QR.OpenResultset(1, 3)
        SELDCBCFLAG = True
    Else
        SELDCBCREPA_QR(0) = req_chglen(cmb_reqno.Text)
        SELDCBCREPA_RS.MoreResults
        SELDCBCREPA_RS.Requery
    End If
    
    If SELDCBCREPA_RS.RowCount > 0 Then
'        txt_crc.Text = SELDCBCREPA_RS("CRC")
'        txt_rpi.Text = SELDCBCREPA_RS("RPI")
'        txt_empc.Text = SELDCBCREPA_RS("EMPCNT")
'        txt_othertype = SELDCBCREPA_RS("fld2")
'        'txt_promonocyte.Text = SELDCBCREPA_RS("fld2") '19/11/08
'        txt_malarialparasite.Text = SELDCBCREPA_RS("maleria")
'        If IsNull(SELDCBCREPA_RS("fld3")) = False Then
'            txt_cbc_comment.Text = SELDCBCREPA_RS("fld3") 'Shraddha 01/03/11
'        End If
        
        If IsNull(SELDCBCREPA_RS("fld7")) = False Then
            TXT_DNAINDEX.Text = SELDCBCREPA_RS("fld7") 'SHWETA 10-02-2014
        End If
        
        If IsNull(SELDCBCREPA_RS("fld6")) = False Then
            TXT_DNACLOIDY.Text = SELDCBCREPA_RS("fld6") 'SHWETA 10-02-2014
        End If
        
    End If
End Sub


Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
btnenabl ("INI")
field_clear

viewflag = False
NEWFLAG = False


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
    
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = True
    cmd_view.Enabled = True
   
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
   
ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = True
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

Private Sub cmd_commit_Click()
Dim i As Integer
If Trim(txt_modcode.Text) = "PF" Then
    If Len(Trim(TXT_PFTIMPRESSION.Text)) = 0 Then
        MsgBox "Impression Field is Empty, You can not Commit the Report.", vbCritical
        TXT_PFTIMPRESSION.SetFocus
        Exit Sub
    End If
End If
i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If i = vbYes Then
On Error GoTo MSG:
 rdb.BeginTrans
    updrrequisitn1_qr(0) = UCase(cmb_finalby.Text)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    MsgBox "Record is Committed Successfully."
End If
rdb.CommitTrans
btnenabl ("INI")

cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub CMD_EMR_Click()
'Dim VAR1
''VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
'VAR1 = MaskEdBox1.FormattedText
'VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
'AppActivate VAR
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
Dim str As String
Dim D1 As String
Dim D2 As String
flagsave = 2
btnenabl ("NEW")
    FLDENABL ("MODIFY")
    ENABL ("ENA")
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

'      If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'        STR = ""
'            STR = cmb_enterby.Text
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'            cmb_enterby.Text = STR
'        End If

'        If seldissecfinflag = False Then
'            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'            seldissecfinflag1 = True
'        Else
'            seldissecfin1_rs.MoreResults
'            seldissecfin1_rs.Requery
'        End If
'        If seldissecfin1_rs.RowCount > 0 Then
'            STR = ""
'            STR = cmb_finalby.Text
'            cmb_finalby.clear
'            While Not seldissecfin1_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                seldissecfin1_rs.MoveNext
'            Wend
'            cmb_finalby.Text = STR
'        End If
'        If seldissecprovflag1 = False Then
'            Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
'            seldissecprovflag1 = True
'        Else
'            seldissecprov1_rs.MoreResults
'            seldissecprov1_rs.Requery
'        End If
'        If seldissecprov1_rs.RowCount > 0 Then
'            entr = ""
'            entr = CMB_PROVISIONAL.Text
'            CMB_PROVISIONAL.clear
'            While Not seldissecprov1_rs.EOF
'                CMB_PROVISIONAL.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                seldissecprov1_rs.MoveNext
'            Wend
'            CMB_PROVISIONAL.Text = entr
'            entr = ""
'        End If
        
        
cmb_enterby.Enabled = True
cmb_finalby.Enabled = True
ENABL ("ENA")
If ADDENDUMFALG = True Then
    cmb_addendumby.Enabled = True
    txt_addendum.Enabled = True
    FLDENABL ("ADDENDUM")
End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG:
btnenabl ("NEW")
field_clear
'TXT_CLEAR


If DATETIMEFLAG = False Then
   datetime
   DATETIMEFLAG = True
Else
    datetime
End If
txt_date.Text = DATE11
txt_time.Text = time1
cmb_enterby.Text = cmb_enterby.List(0)
cmb_enterby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enterby.List(0)) = cmb_finalby.List(L1) Then
       ' cmb_finalby.Text = cmb_finalby.List(L1)
        ' AS INFORMED BY DEDHIA ON 25-06-2004 FOLLOWING LINE HASBEEN HARDCODED BY AMIT
        cmb_finalby.Text = "C. N. NAIR"
    End If
Next L1
'cmb_finalby.Text = cmb_finalby.List(0)
'cmb_provisional.Text = cmb_provisional.List(0)
If SELDEPTRSFLAG1 = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SELDEPTRSFLAG1 = True
Else
    sel_deptrs.MoreResults
    sel_deptrs.Requery
End If
If sel_deptrs.RowCount > 0 Then
    cmb_dept.clear
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
        Wend
   End If
Opt_patient.Value = True

flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
MaskEdBox1.Enabled = True

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_cancel_Click()
btnenabl ("INI")
field_clear
'viewflag = False
'NEWFLAG = False

ADDENDUMFALG = False
FSFLAG = False
'FLDENABL ("DIS")
FLDENABL ("DIS")

End Sub

Public Sub field_clear()
Dim obj As Object
txt_addendum.Text = ""
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
txt_deptcode.Text = " "
cmb_reqno.ListIndex = -1
cmb_reqno.clear
If Opt_patient.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = ">??/#####"
End If
If opt_staff.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = "#######"   'VARINDER
End If
TXT_CD10.Text = ""
txt_cd103.Text = ""
TXT_CD117.Text = ""
TXT_CD11C.Text = ""
txt_cd13.Text = ""
TXT_CD138.Text = ""
txt_cd14.Text = ""
txt_CD15.Text = ""
TXT_CD19.Text = ""
TXT_CD2.Text = ""
TXT_CD20.Text = ""
TXT_CD22.Text = ""
TXT_CD23.Text = ""
TXT_CD25.Text = ""
TXT_CD3.Text = ""
txt_CD33.Text = ""
TXT_CD34.Text = ""
txt_cd36.Text = ""
TXT_CD38.Text = ""
TXT_CD4.Text = ""
txt_cd41.Text = ""
TXT_CD42.Text = ""
TXT_CD5.Text = ""
TXT_CD56.Text = ""
txt_cd13.Text = ""
        txt_cd14.Text = ""
        txt_CD15.Text = ""
        txt_CD33.Text = ""
        txt_cd36.Text = ""
        txt_cd41.Text = ""
        TXT_CD42.Text = ""
        TXT_CD117.Text = ""
        TXT_CD20.Text = ""
        TXT_CD61.Text = ""
        TXT_CD56.Text = ""
        txt_Glycophorina.Text = ""
        TXT_CD68.Text = ""
        TXT_CD34.Text = ""
        TXT_HLADR.Text = ""
        TXT_CD38.Text = ""
        TXT_CD138.Text = ""
        TXT_TDT.Text = ""
        TXT_ANTIMPO.Text = ""
        TXT_CD10.Text = ""
        TXT_CD19.Text = ""
        TXT_FMC7.Text = ""
        TXT_CD22.Text = ""
        TXT_CD23.Text = ""
        txt_smig.Text = ""
        txt_CD79B.Text = ""
        TXT_CD25.Text = ""
        TXT_CD11C.Text = ""
        txt_cd103.Text = ""
        txt_kappa.Text = ""
        txt_lanbda.Text = ""
        TXT_DBA44.Text = ""
        txt_cytocd221(1).Text = ""
        txt_cytocd79a.Text = ""
        TXT_CYTOMU.Text = ""
        TXT_CD2.Text = ""
        TXT_CD3.Text = ""
        TXT_CD4.Text = ""
        TXT_CD5.Text = ""
        TXT_CD7.Text = ""
        TXT_CD8.Text = ""
        TXT_CYTOCD3.Text = ""
txt_name.Text = " "
txt_sex.Text = " "
txt_age.Text = " "
txt_category.Text = " "
txt_status.Text = " "
txt_pridoc.Text = " "
txt_ward.Text = " "
txt_bedno.Text = " "
txt_refdoc.Text = " "
txt_procedure.Text = " "
txt_nmaterial.Text = ""
txt_cellcount.Text = ""
txt_cytochemistry.Text = ""
txt_cytomorphology.Text = ""
txt_surfacemarkers.Text = ""
txt_impression.Text = ""
cmb_enterby.Text = ""
CMB_PROVISIONAL.Text = ""
cmb_finalby.Text = ""

'***** Shraddha 21/07/10 *****
txt_igm.Text = ""
TXT_TDT.Text = ""
txt_smig.Text = ""
txt_tcrab.Text = ""
txt_tcrag.Text = ""
txt_clnd1.Text = ""
TXT_SMPCTG.Text = ""
opt_lsp.Value = False
opt_ssp.Value = False
opt_drak.Value = False
opt_mryb.Value = False
opt_sitaram.Value = False
'*****************************

End Sub




Public Sub FLDENABL(EN As String)
If EN = "ADDENDUM" Then
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
        cmb_reqno.Enabled = False
        TXT_PATH.Enabled = False
        txt_name.Enabled = False
        txt_sex.Enabled = False
        txt_age.Enabled = False
        txt_category.Enabled = False
        txt_status.Enabled = False
        txt_pridoc.Enabled = False
        txt_ward.Enabled = False
        txt_bedno.Enabled = False
        txt_modcode.Enabled = False
        txt_refdoc.Enabled = False
        txt_procedure.Enabled = False
        'txt_observed.Visible = False
        'TXT_COMMENTS.Enabled = False
        cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
        txt_nmaterial.Enabled = False
        txt_cellcount.Enabled = False
        txt_cytomorphology.Enabled = False
        txt_cytochemistry.Enabled = False
        txt_surfacemarkers.Enabled = False
        txt_impression.Enabled = False
        cmb_enterby.Enabled = False
        CMB_PROVISIONAL.Enabled = False
        cmb_finalby.Enabled = False
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = True
End If
If EN = "DIS" Then
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
        cmb_reqno.Enabled = False
        TXT_PATH.Enabled = False
        txt_name.Enabled = False
        txt_sex.Enabled = False
        txt_age.Enabled = False
        txt_category.Enabled = False
        txt_status.Enabled = False
        txt_pridoc.Enabled = False
        txt_ward.Enabled = False
        txt_bedno.Enabled = False
        txt_modcode.Enabled = False
        txt_refdoc.Enabled = False
        txt_procedure.Enabled = False
        'txt_observed.Visible = False
        'TXT_COMMENTS.Enabled = False
        cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
        CMB_PROVISIONAL.Enabled = False
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = False
End If
If EN = "ENA" Then
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        TXT_PATH.Enabled = False
        txt_name.Enabled = False
        txt_sex.Enabled = False
        txt_age.Enabled = False
        txt_category.Enabled = False
        txt_status.Enabled = False
        txt_pridoc.Enabled = False
        txt_ward.Enabled = False
        txt_bedno.Enabled = False
        txt_modcode.Enabled = False
        txt_refdoc.Enabled = False
        txt_procedure.Enabled = False
        'txt_observed.Visible = False
        'TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        CMB_PROVISIONAL.Enabled = True
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = False
End If
If EN = "MODIFY" Then
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        TXT_PATH.Enabled = True
        txt_name.Enabled = True
        txt_sex.Enabled = True
        txt_age.Enabled = True
        txt_category.Enabled = True
        txt_status.Enabled = True
        txt_pridoc.Enabled = True
        txt_ward.Enabled = True
        txt_bedno.Enabled = True
        txt_modcode.Enabled = True
        txt_refdoc.Enabled = True
        txt_procedure.Enabled = True
        'txt_observed.Visible = True
        'TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        CMB_PROVISIONAL.Enabled = True
        cmb_addendumby.Enabled = True
        txt_addendum.Enabled = True
        
        '***** Shraddha 21/07/10 *****
        txt_igm.Enabled = True
        TXT_TDT.Enabled = True
        txt_smig.Enabled = True
        txt_tcrab.Enabled = True
        txt_tcrag.Enabled = True
        txt_clnd1.Enabled = True
        TXT_SMPCTG.Enabled = True
        opt_lsp.Enabled = True
        opt_ssp.Enabled = True
        opt_drak.Enabled = True
        opt_mryb.Enabled = True
        opt_sitaram.Enabled = True
        '*****************************
        
        
End If
End Sub



Private Sub cmd_print_Click()
On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim i As Integer
 On Error GoTo prtmsg:

DoEvents

Dim VAR_REQNNO As String


        VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If


            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
'            If SELRREQUISITN_Rs.RowCount > 0 Then
'                prepath = ""
'                While Not SELRREQUISITN_Rs.EOF
'                    If Trim(SELRREQUISITN_Rs("reqnno")) <> req_chglen(Trim(cmb_reqno.Text)) Then
'                        prepath = prepath & req_chglen(SELRREQUISITN_Rs("REQNNO")) & " "
'                    End If
'
'                    SELRREQUISITN_Rs.MoveNext
'                Wend
'             End If
'
If Opt_patient.Value = True Then
             
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
      
     If Trim(txt_modcode.Text) = "BF" Then
         'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBF.RPT"
         'DISBFold.rpt
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBFold.rpt"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            If Len(Trim(cmb_enterby.Text)) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If (Trim(txt_cd13.Text) = "" And Trim(txt_cd14.Text) = "" And Trim(txt_CD15.Text) = "" And Trim(txt_CD33.Text) = "" And Trim(txt_cd36.Text) = "" And Trim(txt_cd41.Text) = "" And Trim(txt_Glycophorina.Text) = "" And Trim(TXT_CD68.Text) = "" And Trim(TXT_CD42.Text) = "" And Trim(TXT_CD117.Text) = "" And Trim(TXT_CD2.Text) = "" And Trim(TXT_CD3.Text) = "" And Trim(TXT_CD4.Text) = "" And Trim(TXT_CD5.Text) = "" And Trim(TXT_CD7.Text) = "" And Trim(TXT_CD8.Text) = "" And Trim(TXT_CYTOCD3.Text) = "" And Trim(TXT_CD56.Text) = "" And Trim(TXT_CD10.Text) = "" And Trim(TXT_CD19.Text) = "" And Trim(TXT_CD22.Text) = "" And Trim(TXT_CD23.Text) = "" And Trim(TXT_CD25.Text) = "" And Trim(txt_cd103.Text) = "" And Trim(TXT_CD61.Text) = "" And Trim(txt_lanbda.Text) = "" And Trim(txt_kappa.Text) = "" And Trim(TXT_DBA44.Text) = "" And Trim(TXT_CD11C.Text) = "" And Trim(TXT_FMC7.Text) = "" And Trim(txt_smig.Text) = "" And Trim(txt_CD79B.Text) = "" And Trim(txt_cytocd79a.Text) = "" And Trim(TXT_CYTOMU.Text) = "") Then
            CrystalReport1.Formulas(13) = "mm = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "mm = '" & "abc" & "'"
        End If
     End If
End If

If opt_staff.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
 
     If Trim(txt_modcode.Text) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBFS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
     End If

End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox Error(Err)
'txt_msg.Text = ""
'btnenabl ("print")
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmd_save_Click()
Dim pcount As Integer
    pcount = 0
Dim beforst As String, afterst As String
Dim K As Long
       While InStr(txt_impression, Chr(9)) > 0
            
            K = InStr(txt_impression, Chr(9))
            beforst = Mid(txt_impression, 1, K - 1) & " "
            afterst = Mid(txt_impression, K + 1)
            txt_impression = beforst & afterst
        Wend
        While InStr(txt_impression, Chr(146)) > 0

            K = InStr(txt_impression, Chr(146))
            beforst = Mid(txt_impression, 1, K - 1) & "`"
            afterst = Mid(txt_impression, K + 1)
            txt_impression = beforst & afterst
       Wend
     
       If Len(Trim(txt_impression.Text)) > 1000 Then
        MsgBox "Max.limit of Impression is 1000 characters."
        TXT_STIMPRESSION.SetFocus
        Exit Sub
       End If
       
       
       If Len(Trim(txt_impression.Text)) = 0 Then
        MsgBox "IMpression field is Empty.", vbCritical
        Exit Sub
       End If
       



If Len(Trim(cmb_enterby.Text)) = 0 Then
    MsgBox "Enter by field is Empty.", vbCritical
    Exit Sub
End If
If Len(Trim(CMB_PROVISIONAL.Text)) = 0 Then
    MsgBox "Registrar  field is Empty.", vbCritical
    Exit Sub
End If
If Len(Trim(cmb_finalby.Text)) = 0 Then
    MsgBox "To be finalised  field is Empty.", vbCritical
    Exit Sub
End If

    assign

btnenabl ("SAVE")
'cmd_print.Enabled = True
'cmd_cancel.Enabled = True



'cmd_commit.Enabled = True
End Sub

Public Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
FLDENABL ("DIS")
ENABL ("DIS")
field_clear
'TXT_CLEAR
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
 
'         If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'        End If
'
'        If seldissecfinflag = False Then
'            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'            seldissecfinflag1 = True
'        Else
'            seldissecfin1_rs.MoreResults
'            seldissecfin1_rs.Requery
'        End If
'        If seldissecfin1_rs.RowCount > 0 Then
'            cmb_finalby.clear
'            While Not seldissecfin1_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                seldissecfin1_rs.MoveNext
'            Wend
'        End If
'        If seldissecprovflag1 = False Then
'            Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
'            seldissecprovflag1 = True
'        Else
'            seldissecprov1_rs.MoreResults
'            seldissecprov1_rs.Requery
'        End If
'        If seldissecprov1_rs.RowCount > 0 Then
'            entr = ""
'            entr = CMB_PROVISIONAL.Text
'            CMB_PROVISIONAL.clear
'            While Not seldissecprov1_rs.EOF
'                CMB_PROVISIONAL.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                seldissecprov1_rs.MoveNext
'            Wend
'            CMB_PROVISIONAL.Text = entr
'            entr = ""
'        End If
        
cmb_modality.Enabled = True
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
Opt_patient.Value = True
cmb_dept.SetFocus
MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
'btnenabl ("INI")
NEWFLAG = False
viewflag = False
FLDENABL ("DIS")
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
If BFREPORTFLAG = False Then
    BFREPORTFLAG = True
'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1BBF", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.Text = (LOGINUSRNAME)
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = cmb_enterby.List(0)
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? and MODMO IN ('5','6','7') ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNSTRBF", str)
If Trim(USRDEPT) = "ALL" Then
    REPWTRN_QR(0) = "UZZ"
Else
    REPWTRN_QR(0) = Trim(USRDEPT)
End If
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)

If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    cmb_addendumby.clear
    cmb_finalby.clear
    CMB_PROVISIONAL.clear
    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 5 And (Trim(REPWTRNRS("MODMO")) = "5" Or Trim(REPWTRNRS("MODMO")) = "7") Then
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        ElseIf REPWTRNRS("USRTYPE") = 5 And (Trim(REPWTRNRS("MODMO")) = "6" Or Trim(REPWTRNRS("MODMO")) = "7") Then
            CMB_PROVISIONAL.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        
        
        REPWTRNRS.MoveNext
    Wend
   ' cmb_finalby.Text = cmb_finalby.List(0)
    
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BBF", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

''**********SECURITY****************************

btnenabl ("INI")

''for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr1bf", str)

'' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr1bf", str)

'    'for retriving reqnos
Set req_qr = rdb.CreateQuery("reqbf", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E'AND MODALITYCD=? ORDER BY REQNDT desc ")

'
Set SELTA101_QR = rdb.CreateQuery("selta101_qr11BF", "select * from PATOOLSLIB.ta101 where case_no =?")

'    ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1BF", str)

'    ' FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1BF", str)
''
'    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11BF", str)

'
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1BF", str)

    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1BF", str)

   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1BF", str)
'
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1BF", str)
'
   str = "INSERT INTO DISLIB.ADDENDUMBF(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QRBF", str)

   str = "SELECT * FROM DISLIB.ADDENDUMBF WHERE REQNO=?"
   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR1BF", str)
'   SELDADDENDUM_QR(0) = "FSS"
'   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
   str = "select * from dislib.drequisitn where caseno=?"
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1BF", str)

    'commented and new query added by amarjeet on 05-11-2002 to select manjure's name in enter by
    str = "select * from dislib.dissec where enterby='YES'AND MODALITY='BF' "
 
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1BF", str)
   
   str = "select * from dislib.dissec where PROVby='YES'AND MODALITY='BF'"
   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1BF", str)
'
   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='BF'"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1BF", str)
    
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1BF", str)
    
  str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
  Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRBF", str)
    
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
  Set selMA101QR = rdb.CreateQuery("SELMA101QRBF", str)
  

  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET=?"
  Set SELDMA101QR = rdb.CreateQuery("SELDMA101QRBF", str)
  
  str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1BF", str)
  
  
  str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrBF", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%EC%'"
  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRBF", str)
 
  ' commented by  sundeep kumar for testing
  'str = "INSERT INTO DISLIB.DHBFREPORT(REQNO,CELLCOUNT,CYTOMORPH,CYTOCHEMIS,surmarker,IMPRESSION,FLD1,FLD2,CD13,CD14,CD15,CD33,CD36,CD4142,GLYCOPHOIN,CD68,CD34,HLADR,CD38,CD138,TDT,ANTIMPO,CD10,CD19,FMC7,CD22,CD23,SMIG,CD79B,CD25,CD11C,CD103,KAPPA,LAMDA,DBA44,CYTOEP22,CYTOCD79A,CYTOMU,CD2,CD3,CD4,CD5,CD7,CD8,CYTOCD3,CD42,CD117,CD20,CD61,CD56,IGM)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  'Added by Shraddha A. Kesarkar on 21/07/10
  str = "INSERT INTO DISLIB.DHBFREPORT(REQNO,CELLCOUNT,CYTOMORPH,CYTOCHEMIS,surmarker,IMPRESSION,FLD1,FLD2,CD13,CD14,CD15,CD33,CD36,CD4142,GLYCOPHOIN,CD68,CD34,HLADR,CD38,CD138,TDT,ANTIMPO,CD10,CD19,FMC7,CD22,CD23,SMIG,CD79B,CD25,CD11C,CD103,KAPPA,LAMDA,DBA44,CYTOEP22,CYTOCD79A,CYTOMU,CD2,CD3,CD4,CD5,CD7,CD8,CYTOCD3,CD42,CD117,CD20,CD61,CD56,IGM,TCRAB,TCRGD,CLND1,SMPCTG,OPTLSP,OPTSCP)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  'str = "INSERT INTO DISLIB.DHBFREPORT(REQNO,CELLCOUNT,CYTOMORPH,CYTOCHEMIS,surmarker,IMPRESSION,FLD1,FLD2,CD13,CD14,CD15,CD33,CD36,CD4142,GLYCOPHOIN,CD68,CD34,HLADR,CD38,CD138,TDT,ANTIMPO,CD10,CD19,FMC7,CD22,CD23,SMIG,CD79B,CD25,CD11C,CD103,KAPPA,LAMDA,DBA44,CYTOEP22,CYTOCD79A,CYTOMU,CD2,CD3,CD4,CD5,CD7,CD8,CYTOCD3,CD42,CD117,CD20,CD61,CD56,IGM,TCRAB,TCRGD,CLND1,SMPCTG,OPTLSP,OPTSCP)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  'STR = "INSERT INTO DISLIB.DHBFREPORT(REQNO,CELLCOUNT,CYTOMORPH,CYTOCHEMIS,SURMARKER,IMPRESSION)VALUES(?,?,?,?,?,?)"
  Set INSDHBFREPORT_QR = rdb.CreateQuery("INSDHBFREPORT_QRBF", str)
  
  '***** Shraddha 21/07/10 *****
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
  Set SELDREQDRNAME_QR = rdb.CreateQuery("", str)
  
  str = "UPDATE DISLIB.DREQUISITN SET LATERALITS=?, METASTASIS=? WHERE REQNO=?"
  Set UPDDREQDRNAME_QR = rdb.CreateQuery("", str)
  '*****************************
  
  
  'Commented by sundeep kumar for testing
  'str = "UPDATE DISLIB.DHBFREPORT SET CELLCOUNT = ?,CYTOMORPH = ?, CYTOCHEMIS = ?, SURMARKER = ?, IMPRESSION = ?, FLD1 = ?, FLD2= ?, CD13=?,CD14=?,CD15=?,CD33=?,CD36=?,CD4142=?,GLYCOPHOIN=?,CD68=?,CD34=?,HLADR=?,CD38=?,CD138=?,TDT=?,ANTIMPO=?,CD10=?,CD19=?,FMC7=?,CD22=?,CD23=?,SMIG=?,CD79B=?,CD25=?,CD11C=?,CD103=?,KAPPA=?,LAMDA=?,DBA44=?,CYTOEP22=?,CYTOCD79A=?,CYTOMU=?,CD2=?,CD3=?,CD4=?,CD5=?,CD7=?,CD8=?,CYTOCD3=?,CD42=?,CD117=?,CD20=?,CD61=?,CD56=?,IGM=?  WHERE  REQNO=?"
  'str = "UPDATE DISLIB.DHBFREPORT SET CELLCOUNT = ?,CYTOMORPH = ?, CYTOCHEMIS = ?, SURMARKER = ?, IMPRESSION = ?, FLD1 = ?, FLD2= ?, CD13=?,CD14=?,CD15=?,CD33=?,CD36=?,CD4142=?,GLYCOPHOIN=?,CD68=?,CD34=?,HLADR=?,CD38=?,CD138=?,TDT=?,ANTIMPO=?,CD10=?,CD19=?,FMC7=?,CD22=?,CD23=?,SMIG=?,CD79B=?,CD25=?,CD11C=?,CD103=?,KAPPA=?,LAMDA=?,DBA44=?,CYTOEP22=?,CYTOCD79A=?,CYTOMU=?,CD2=?,CD3=?,CD4=?,CD5=?,CD7=?,CD8=?,CYTOCD3=?,CD42=?,CD117=?,CD20=?,CD61=?,CD56=?,IGM=?  WHERE  REQNO=?"
  'Added by Shraddha on 21/07/10
  str = "UPDATE DISLIB.DHBFREPORT SET CELLCOUNT = ?,CYTOMORPH = ?, CYTOCHEMIS = ?, SURMARKER = ?, IMPRESSION = ?, FLD1 = ?, FLD2= ?, CD13=?,CD14=?,CD15=?,CD33=?,CD36=?,CD4142=?,GLYCOPHOIN=?,CD68=?,CD34=?,HLADR=?,CD38=?,CD138=?,TDT=?,ANTIMPO=?,CD10=?,CD19=?,FMC7=?,CD22=?,CD23=?,SMIG=?,CD79B=?,CD25=?,CD11C=?,CD103=?,KAPPA=?,LAMDA=?,DBA44=?,CYTOEP22=?,CYTOCD79A=?,CYTOMU=?,CD2=?,CD3=?,CD4=?,CD5=?,CD7=?,CD8=?,CYTOCD3=?,CD42=?,CD117=?,CD20=?,CD61=?,CD56=?,IGM=?,TCRAB=?,TCRGD=?,CLND1=?,SMPCTG=?,OPTLSP=?,OPTSCP=?  WHERE  REQNO=?"
  'STR = "UPDATE DISLIB.DHBFREPORT SET CELLCOUNT=?,CYTOMORPH=?,CYTOCHEMIS=?,SURMARKER=?,IMPRESSION=? WHERE REQNO=?"
  Set UPDDHBFREPORT_QR = rdb.CreateQuery("UPDDHBFREPORT_QRBF", str)
  
  str = "SELECT * FROM DISLIB.DHBFREPORT WHERE REQNO=?"
  Set SELDHBFREPORT_QR = rdb.CreateQuery("SELDHBFREPORT_QRBF", str)
  
  End If
  
  '******added by shweta 10-02-2014
  
   str = "INSERT INTO DISLIB.DCBCREPA (REQNO,FLD7,FLD6) VALUES (?,?,?)"
  Set INSDCBCREPA_QR = rdb.CreateQuery("INSERTDCBCREPA", str)
  
   str = "UPDATE DISLIB.DCBCREPA SET FLD7=?, FLD6=? WHERE REQNO=?"
  Set UPDDCBCREPA_QR = rdb.CreateQuery("UPDATEDCBCREPA", str)
  
  
  str = "SELECT REQNO,FLD7,FLD6 FROM DISLIB.DCBCREPA WHERE REQNO=?"
  Set SELDCBCREPA_QR = rdb.CreateQuery("SELECTDCBCREPA", str)
  ''''till here
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
       ' MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 05/12/2022
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
            If REQFLAG1 = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
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
            
'          If SELDREQNOFLAG = False Then
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'            SELDREQNOFLAG = True
'          Else
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            SELDREQNORS.MoreResults
'            SELDREQNORS.Requery
'          End If
'          If SELDREQNORS.RowCount > 0 Then
'            CMB_PATHNO.clear
'            While Not SELDREQNORS.EOF
'                CMB_PATHNO.AddItem SELDREQNORS("PATHNO")
'                SELDREQNORS.MoveNext
'            Wend
'          End If
        End If
        
        If viewflag = True Then
           
            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
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
    If opt_staff.Value = True Then
            CASENO = Trim(MaskEdBox1.Text)
        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
         If NEWFLAG = True Then
            If REQFLAG1 = False Then
                req_qr(0) = Trim(MaskEdBox1.FormattedText)
                req_qr(1) = Trim(MODCD)
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = Trim(MaskEdBox1.FormattedText)
                req_qr(1) = Trim(MODCD)
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
            
'          If SELDREQNOFLAG = False Then
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'            SELDREQNOFLAG = True
'          Else
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            SELDREQNORS.MoreResults
'            SELDREQNORS.Requery
'          End If
'          If SELDREQNORS.RowCount > 0 Then
'            CMB_PATHNO.clear
'            While Not SELDREQNORS.EOF
'                CMB_PATHNO.AddItem SELDREQNORS("PATHNO")
'                SELDREQNORS.MoveNext
'            Wend
'          End If
        End If
        
        If viewflag = True Then
            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
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
 Exit Sub
MSG:
MsgBox Error(Err)
End Sub




Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
   ' MaskEdBox1.Mask = "??>/#####"
    MaskEdBox1.Enabled = True
    
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then

    lbl_caseno.Caption = "  CC No."
    MaskEdBox1.Mask = "########"
    MaskEdBox1.Enabled = True
    
End If
End Sub

Public Sub STAFF()
 If NEWFLAG = True Then
            NEWFLAG = False
            If SELDREQSTAFFD1FLAG = False Then
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
                Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
            Else
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
                SELDREQSTAFFDRS.MoreResults
                SELDREQSTAFFDRS.Requery
                
            End If
            If SELDREQSTAFFDRS.RowCount > 0 Then
                txt_sex.Text = SELDREQSTAFFDRS("sex")
                txt_name.Text = SELDREQSTAFFDRS("PNAME")
                txt_age = SELDREQSTAFFDRS("AGE")
                txt_pridoc.Text = ""
                txt_category.Text = SELDREQSTAFFDRS("RELATION")
            End If
            If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
            End If
          
'            If SELMA104FLAG1 = False Then
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'                SELMA104FLAG1 = True
'            Else
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                selma104_rs.MoreResults
'                selma104_rs.Requery
'            End If
          
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
          
            If REFDR_RS.RowCount > 0 Then
                If REFDR_RS("PATNSTAT") = "IP" Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                lbl_ward.Visible = True
                lbl_bed.Visible = True
                txt_ward.Text = selpathdrequisitn_rs("WARD_NAME")
                txt_bedno.Text = selpathdrequisitn_rs("bedno")
                txt_status.Text = "IP"
            Else
                txt_status.Text = "OP"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
            End If
           End If
            
            If selpathdrequisitn_rs.RowCount > 0 Then
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                '***** Added by Shraddha on 28/12/10
                TXT_PATH.Text = "" & Trim(selpathdrequisitn_rs("PATHNO"))
                
                If selpathdrequisitn_rs.EOF Then
                    txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
                'txt_refdoc.Text = SELPATHDREQUISITN_RS("REFDOC")
       
            End If
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
           
            
            
            

'        If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'
'            entr = cmb_enterby.Text
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'            seldissecentr1_rs.MovePrevious
'            cmb_enterby.Text = seldissecentr1_rs("CCNAME")
'
'        End If
'
'        If seldissecfinflag = False Then
'            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'            seldissecfinflag1 = True
'        Else
'            seldissecfin1_rs.MoreResults
'            seldissecfin1_rs.Requery
'        End If
'        If seldissecfin1_rs.RowCount > 0 Then
'            entr = ""
'            entr = cmb_finalby.Text
'            cmb_finalby.clear
'            While Not seldissecfin1_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                seldissecfin1_rs.MoveNext
'            Wend
'            cmb_finalby.Text = entr
'            entr = ""
'        End If
'
'        If seldissecprovflag1 = False Then
'            Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
'            seldissecprovflag1 = True
'        Else
'            seldissecprov1_rs.MoreResults
'            seldissecprov1_rs.Requery
'        End If
'        If seldissecprov1_rs.RowCount > 0 Then
'            entr = ""
'            entr = CMB_PROVISIONAL.Text
'            CMB_PROVISIONAL.clear
'            While Not seldissecprov1_rs.EOF
'                CMB_PROVISIONAL.AddItem Trim(seldissecprov1_rs("CCNAME"))
'                seldissecprov1_rs.MoveNext
'            Wend
'            CMB_PROVISIONAL.Text = entr
'            entr = ""
'        End If
         
          
        
        End If
        If viewflag = True Then
        
           
            
            viewflag = False
            If SELDREQSTAFFD1FLAG = False Then
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
                Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
            Else
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = MaskEdBox1.FormattedText
                SELDREQSTAFFDRS.MoreResults
                SELDREQSTAFFDRS.Requery
                
            End If
            If SELDREQSTAFFDRS.RowCount > 0 Then
                txt_sex.Text = SELDREQSTAFFDRS("sex")
                txt_name.Text = SELDREQSTAFFDRS("PNAME")
                txt_age = SELDREQSTAFFDRS("age")
                txt_pridoc.Text = ""
                txt_category.Text = SELDREQSTAFFDRS("RELATION")
            End If
        
            If REFDRFLAG1 = True Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
        
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Report is commited , You can not Modify"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                CMB_PROVISIONAL.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
                End If
                
            End If
            
'            If SELMA104FLAG1 = False Then
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'                SELMA104FLAG1 = True
'            Else
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                selma104_rs.MoreResults
'                selma104_rs.Requery
'            End If
           If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
          
            If REFDR_RS.RowCount > 0 Then
                If REFDR_RS("PATNSTAT") = "IP" Then
                    txt_ward.Visible = False
                    txt_bedno.Visible = False
                    lbl_ward.Visible = False
                    lbl_bed.Visible = False
                    txt_ward.Text = selpathdrequisitn_rs("WARD_NAME")
                    txt_bedno.Text = selpathdrequisitn_rs("bedno")
                    txt_status.Text = "IP"
                Else
                    txt_ward.Visible = False
                    txt_bedno.Visible = False
                    lbl_ward.Visible = False
                    lbl_bed.Visible = False
                    txt_status.Text = "OP"
                End If
            End If
         
           
            If selpathdrequisitn_rs.RowCount > 0 Then
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF = False Then
                    'txt_relclnhis.Visible = True
                    
                    'txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
            End If
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = MaskEdBox1.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
            If SELDHBFREPORTFLAG = False Then
                SELDHBFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDHBFREPORT_RS = SELDHBFREPORT_QR.OpenResultset(1, 3)
                SELDHBFREPORTFLAG = True
            Else
                  SELDHBFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                  SELDHBFREPORT_RS.MoreResults
                  SELDHBFREPORT_RS.Requery
            End If
            ENABL ("DIS")
            If SELDHBFREPORT_RS.RowCount > 0 Then
                txt_cellcount.Text = "" & SELDHBFREPORT_RS("CELLCOUNT")
                txt_cytomorphology.Text = "" & SELDHBFREPORT_RS("CYTOMORPH")
                txt_cytochemistry.Text = "" & SELDHBFREPORT_RS("CYTOCHEMIS")
                txt_surfacemarkers.Text = "" & SELDHBFREPORT_RS("SURMARKER")
                txt_impression.Text = "" & SELDHBFREPORT_RS("IMPRESSION")
                '''********************* SURFACE MARKER
                ' sundeep
                    txt_cd13.Text = " " & Trim(SELDHBFREPORT_RS("cd13"))
                    txt_cd14.Text = " " & Trim(SELDHBFREPORT_RS("CD14"))
                    txt_CD15.Text = " " & Trim(SELDHBFREPORT_RS("CD15"))
                    txt_CD33.Text = " " & Trim(SELDHBFREPORT_RS("CD33"))
                    txt_cd36.Text = " " & Trim(SELDHBFREPORT_RS("CD36"))
                    txt_cd41.Text = " " & Trim(SELDHBFREPORT_RS("CD4142"))
                    TXT_CD42.Text = "" & Trim(SELDHBFREPORT_RS("CD42"))
                    TXT_CD117.Text = "" & Trim(SELDHBFREPORT_RS("CD117"))
                    TXT_CD20.Text = "" & Trim(SELDHBFREPORT_RS("CD20"))
                    TXT_CD61.Text = "" & Trim(SELDHBFREPORT_RS("CD61"))
                    TXT_CD56.Text = "" & Trim(SELDHBFREPORT_RS("CD56"))
                    txt_Glycophorina.Text = " " & Trim(SELDHBFREPORT_RS("GLYCOPHOIN"))
                    TXT_CD68.Text = " " & Trim(SELDHBFREPORT_RS("CD68"))
                    TXT_CD34.Text = " " & Trim(SELDHBFREPORT_RS("CD34"))
                    TXT_HLADR.Text = " " & Trim(SELDHBFREPORT_RS("HLADR"))
                    TXT_CD38.Text = " " & Trim(SELDHBFREPORT_RS("CD38"))
                    TXT_CD138.Text = " " & Trim(SELDHBFREPORT_RS("CD138"))
                    TXT_TDT.Text = " " & Trim(SELDHBFREPORT_RS("TDT"))
                    TXT_ANTIMPO.Text = " " & Trim(SELDHBFREPORT_RS("ANTIMPO"))
                    TXT_CD10.Text = " " & Trim(SELDHBFREPORT_RS("CD10"))
                    TXT_CD19.Text = " " & Trim(SELDHBFREPORT_RS("CD19"))
                    TXT_FMC7.Text = " " & Trim(SELDHBFREPORT_RS("FMC7"))
                    TXT_CD22.Text = " " & Trim(SELDHBFREPORT_RS("CD22"))
                    TXT_CD23.Text = " " & Trim(SELDHBFREPORT_RS("CD23"))
                    txt_smig.Text = " " & Trim(SELDHBFREPORT_RS("SMIG"))
                    txt_CD79B.Text = " " & Trim(SELDHBFREPORT_RS("CD79B"))
                    TXT_CD25.Text = " " & Trim(SELDHBFREPORT_RS("CD25"))
                    TXT_CD11C.Text = " " & Trim(SELDHBFREPORT_RS("CD11C"))
                    txt_cd103.Text = " " & Trim(SELDHBFREPORT_RS("CD103"))
                    txt_kappa.Text = " " & Trim(SELDHBFREPORT_RS("KAPPA"))
                    txt_lanbda.Text = " " & Trim(SELDHBFREPORT_RS("LAMDA"))
                    TXT_DBA44.Text = " " & Trim(SELDHBFREPORT_RS("DBA44"))
                    txt_cytocd221(1).Text = " " & Trim(SELDHBFREPORT_RS("CYTOEP22"))
                    txt_cytocd79a.Text = " " & Trim(SELDHBFREPORT_RS("CYTOCD79A"))
                    TXT_CYTOMU.Text = " " & Trim(SELDHBFREPORT_RS("CYTOMU"))
                    TXT_CD2.Text = " " & Trim(SELDHBFREPORT_RS("CD2"))
                    TXT_CD3.Text = " " & Trim(SELDHBFREPORT_RS("CD3"))
                    TXT_CD4.Text = " " & Trim(SELDHBFREPORT_RS("CD4"))
                    TXT_CD5.Text = " " & Trim(SELDHBFREPORT_RS("CD5"))
                    TXT_CD7.Text = " " & Trim(SELDHBFREPORT_RS("CD7"))
                    TXT_CD8.Text = " " & Trim(SELDHBFREPORT_RS("CD8"))
                    TXT_CYTOCD3.Text = " " & Trim(SELDHBFREPORT_RS("CYTOCD3"))
                'sundeep
               
                '''********************* SURFACE MARKER
            End If

            
        End If
End Sub
'*************************************************
'*************************************************

Public Sub assign()
On Error GoTo MSG:
rdb.BeginTrans
    If flagsave = 1 Then
        '*** Inserting into  DHBFREPORT
        INSDHBFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
        If Len(Trim(txt_cellcount.Text)) > 0 Then
            INSDHBFREPORT_QR(1) = Null
            INSDHBFREPORT_QR(1) = Trim(txt_cellcount.Text)
        Else
            INSDHBFREPORT_QR(1) = Null
        End If
        If Len(Trim(txt_cytomorphology.Text)) > 0 Then
            INSDHBFREPORT_QR(2) = Null
            INSDHBFREPORT_QR(2) = Trim(txt_cytomorphology.Text)
        Else
            INSDHBFREPORT_QR(2) = Null
        End If
        If Len(Trim(txt_cytochemistry.Text)) > 0 Then
            INSDHBFREPORT_QR(3) = Null
            INSDHBFREPORT_QR(3) = Trim(txt_cytochemistry.Text)
        Else
            INSDHBFREPORT_QR(3) = Null
        End If
        If Len(Trim(txt_surfacemarkers.Text)) > 0 Then
            INSDHBFREPORT_QR(4) = Null
            INSDHBFREPORT_QR(4) = Trim(txt_surfacemarkers.Text)
        Else
            INSDHBFREPORT_QR(4) = Null
        End If
        If Len(Trim(txt_impression.Text)) > 0 Then
            INSDHBFREPORT_QR(5) = Null
            INSDHBFREPORT_QR(5) = Trim(txt_impression.Text)
        Else
            INSDHBFREPORT_QR(5) = Null
        End If
        '""""""""""
          '**** SURFACE MARKERS****
'sundeep
           INSDHBFREPORT_QR(6) = Null
           INSDHBFREPORT_QR(7) = Null

        If Len(Trim(txt_cd13.Text)) > 0 Then
            INSDHBFREPORT_QR(8) = Trim(txt_cd13.Text)
        Else
             INSDHBFREPORT_QR(8) = Null
        End If
        
        If Len(Trim(txt_cd14.Text)) > 0 Then
             INSDHBFREPORT_QR(9) = Trim(txt_cd14.Text)
        Else
            INSDHBFREPORT_QR(9) = Null
        End If
        
        If Len(Trim(txt_CD15.Text)) > 0 Then
            INSDHBFREPORT_QR(10) = Trim(txt_CD15.Text)
        Else
            INSDHBFREPORT_QR(10) = Null
        End If
        
        If Len(Trim(txt_CD33.Text)) > 0 Then
            INSDHBFREPORT_QR(11) = Trim(txt_CD33.Text)
        Else
            INSDHBFREPORT_QR(11) = Null
        End If
        
        If Len(Trim(txt_cd36.Text)) > 0 Then
            INSDHBFREPORT_QR(12) = Trim(txt_cd36.Text)
        Else
            INSDHBFREPORT_QR(12) = Null
        End If
        
        If Len(Trim(txt_cd41.Text)) > 0 Then
            INSDHBFREPORT_QR(13) = Trim(txt_cd41.Text)
        Else
            INSDHBFREPORT_QR(13) = Null
        End If
        
        If Len(Trim(txt_Glycophorina.Text)) > 0 Then
            INSDHBFREPORT_QR(14) = Trim(txt_Glycophorina.Text)
        Else
            INSDHBFREPORT_QR(14) = Null
        End If
        
        If Len(Trim(TXT_CD68.Text)) > 0 Then
            INSDHBFREPORT_QR(15) = Trim(TXT_CD68.Text)
        Else
            INSDHBFREPORT_QR(15) = Null
        End If
        
        If Len(Trim(TXT_CD34.Text)) > 0 Then
            INSDHBFREPORT_QR(16) = Trim(TXT_CD34.Text)
        Else
            INSDHBFREPORT_QR(16) = Null
        End If
        
        If Len(Trim(TXT_HLADR.Text)) > 0 Then
            INSDHBFREPORT_QR(17) = Trim(TXT_HLADR.Text)
        Else
            INSDHBFREPORT_QR(17) = Null
        End If
        
        If Len(Trim(TXT_CD38.Text)) > 0 Then
            INSDHBFREPORT_QR(18) = Trim(TXT_CD38.Text)
        Else
            INSDHBFREPORT_QR(18) = Null
        End If
        
        If Len(Trim(TXT_CD138.Text)) > 0 Then
            INSDHBFREPORT_QR(19) = Trim(TXT_CD138.Text)
        Else
           INSDHBFREPORT_QR(19) = Null
        End If
        
        If Len(Trim(TXT_TDT.Text)) > 0 Then
            INSDHBFREPORT_QR(20) = Trim(TXT_TDT.Text)
        Else
            INSDHBFREPORT_QR(20) = Null
        End If
        
        If Len(Trim(TXT_ANTIMPO.Text)) > 0 Then
            INSDHBFREPORT_QR(21) = Trim(TXT_ANTIMPO.Text)
        Else
            INSDHBFREPORT_QR(21) = Null
        End If
        
        If Len(Trim(TXT_CD10.Text)) > 0 Then
            INSDHBFREPORT_QR(22) = Trim(TXT_CD10.Text)
        Else
            INSDHBFREPORT_QR(22) = Null
        End If
        
        If Len(Trim(TXT_CD19.Text)) > 0 Then
            INSDHBFREPORT_QR(23) = Trim(TXT_CD19.Text)
        Else
            INSDHBFREPORT_QR(23) = Null
        End If
        
        If Len(Trim(TXT_FMC7.Text)) > 0 Then
            INSDHBFREPORT_QR(24) = Trim(TXT_FMC7.Text)
        Else
            INSDHBFREPORT_QR(24) = Null
        End If
        
        If Len(Trim(TXT_CD22.Text)) > 0 Then
            INSDHBFREPORT_QR(25) = Trim(TXT_CD22.Text)
        Else
           INSDHBFREPORT_QR(25) = Null
        End If
        
        If Len(Trim(TXT_CD23.Text)) > 0 Then
            INSDHBFREPORT_QR(26) = Trim(TXT_CD23.Text)
        Else
            INSDHBFREPORT_QR(26) = Null
        End If
        
        If Len(Trim(txt_smig.Text)) > 0 Then
            INSDHBFREPORT_QR(27) = Trim(txt_smig.Text)
        Else
            INSDHBFREPORT_QR(27) = Null
        End If
        
        If Len(Trim(txt_CD79B.Text)) > 0 Then
            INSDHBFREPORT_QR(28) = Trim(txt_CD79B.Text)
        Else
            INSDHBFREPORT_QR(28) = Null
        End If
        
        If Len(Trim(TXT_CD25.Text)) > 0 Then
            INSDHBFREPORT_QR(29) = Trim(TXT_CD25.Text)
        Else
            INSDHBFREPORT_QR(29) = Null
        End If
        
        If Len(Trim(TXT_CD11C.Text)) > 0 Then
            INSDHBFREPORT_QR(30) = Trim(TXT_CD11C.Text)
        Else
            INSDHBFREPORT_QR(30) = Null
        End If
        
        If Len(Trim(txt_cd103.Text)) > 0 Then
            INSDHBFREPORT_QR(31) = Trim(txt_cd103.Text)
        Else
            INSDHBFREPORT_QR(31) = Null
        End If
        
        If Len(Trim(txt_kappa.Text)) > 0 Then
            INSDHBFREPORT_QR(32) = Trim(txt_kappa.Text)
        Else
            INSDHBFREPORT_QR(32) = Null
        End If
        
        If Len(Trim(txt_lanbda.Text)) > 0 Then
            INSDHBFREPORT_QR(33) = Trim(txt_lanbda.Text)
        Else
            INSDHBFREPORT_QR(33) = Null
        End If
        
        If Len(Trim(TXT_DBA44.Text)) > 0 Then
            INSDHBFREPORT_QR(34) = Trim(TXT_DBA44.Text)
        Else
            INSDHBFREPORT_QR(34) = Null
        End If
        
        If Len(Trim(txt_cytocd221(1).Text)) > 0 Then
            INSDHBFREPORT_QR(35) = Trim(txt_cytocd221(1).Text)
        Else
            INSDHBFREPORT_QR(35) = Null
        End If
        
        If Len(Trim(txt_cytocd79a.Text)) > 0 Then
            INSDHBFREPORT_QR(36) = Trim(txt_cytocd79a.Text)
        Else
            INSDHBFREPORT_QR(36) = Null
        End If
        
        If Len(Trim(TXT_CYTOMU.Text)) > 0 Then
            INSDHBFREPORT_QR(37) = Trim(TXT_CYTOMU.Text)
        Else
           INSDHBFREPORT_QR(37) = Null
        End If
        
        If Len(Trim(TXT_CD2.Text)) > 0 Then
            INSDHBFREPORT_QR(38) = Trim(TXT_CD2.Text)
        Else
            INSDHBFREPORT_QR(38) = Null
        End If
        
        If Len(Trim(TXT_CD3.Text)) > 0 Then
            INSDHBFREPORT_QR(39) = Trim(TXT_CD3.Text)
        Else
           INSDHBFREPORT_QR(39) = Null
        End If
        
        If Len(Trim(TXT_CD4.Text)) > 0 Then
            INSDHBFREPORT_QR(40) = Trim(TXT_CD4.Text)
        Else
            INSDHBFREPORT_QR(40) = Null
        End If
        If Len(Trim(TXT_CD5.Text)) > 0 Then
            INSDHBFREPORT_QR(41) = Trim(TXT_CD5.Text)
        Else
            INSDHBFREPORT_QR(41) = Null
        End If
        If Len(Trim(TXT_CD7.Text)) > 0 Then
            INSDHBFREPORT_QR(42) = Trim(TXT_CD7.Text)
        Else
           INSDHBFREPORT_QR(42) = Null
        End If
        If Len(Trim(TXT_CD8.Text)) > 0 Then
            INSDHBFREPORT_QR(43) = Trim(TXT_CD8.Text)
        Else
            INSDHBFREPORT_QR(43) = Null
        End If
        If Len(Trim(TXT_CYTOCD3.Text)) > 0 Then
            INSDHBFREPORT_QR(44) = Trim(TXT_CYTOCD3.Text)
        Else
            INSDHBFREPORT_QR(44) = Null
        End If
        If Len(Trim(TXT_CD42.Text)) > 0 Then
            INSDHBFREPORT_QR(45) = Trim(TXT_CD42.Text)
        Else
           INSDHBFREPORT_QR(45) = Null
        End If
        If Len(Trim(TXT_CD117.Text)) > 0 Then
            INSDHBFREPORT_QR(46) = Trim(TXT_CD117.Text)
        Else
           INSDHBFREPORT_QR(46) = Null
        End If
        If Len(Trim(TXT_CD20.Text)) > 0 Then
            INSDHBFREPORT_QR(47) = Trim(TXT_CD20.Text)
        Else
           INSDHBFREPORT_QR(47) = Null
        End If
        If Len(Trim(TXT_CD61.Text)) > 0 Then
            INSDHBFREPORT_QR(48) = Trim(TXT_CD61.Text)
        Else
           INSDHBFREPORT_QR(48) = Null
        End If
        If Len(Trim(TXT_CD56.Text)) > 0 Then
            INSDHBFREPORT_QR(49) = Trim(TXT_CD56.Text)
        Else
           INSDHBFREPORT_QR(49) = Null
        End If
        If Len(Trim(txt_igm.Text)) > 0 Then
            INSDHBFREPORT_QR(50) = Trim(txt_igm.Text)
        Else
           INSDHBFREPORT_QR(50) = Null
        End If
  ' sundeep
        '"""""""""""
        
        '***** Shraddha 21/07/10 *****
        
        If Len(Trim(txt_tcrab.Text)) > 0 Then
            INSDHBFREPORT_QR(51) = Trim(txt_tcrab.Text)
        Else
            INSDHBFREPORT_QR(51) = Null
        End If
        If Len(Trim(txt_tcrag.Text)) > 0 Then
            INSDHBFREPORT_QR(52) = Trim(txt_tcrag.Text)
        Else
            INSDHBFREPORT_QR(52) = Null
        End If
        If Len(Trim(txt_clnd1.Text)) > 0 Then
            INSDHBFREPORT_QR(53) = Trim(txt_clnd1.Text)
        Else
            INSDHBFREPORT_QR(53) = Null
        End If
        If Len(Trim(TXT_SMPCTG.Text)) > 0 Then
            INSDHBFREPORT_QR(54) = Trim(TXT_SMPCTG.Text)
        Else
            INSDHBFREPORT_QR(54) = Null
        End If
        If opt_lsp.Value = True Then
            INSDHBFREPORT_QR(55) = "Y"
        Else
            INSDHBFREPORT_QR(55) = "N"
        End If
        If opt_ssp.Value = True Then
            INSDHBFREPORT_QR(56) = "Y"
        Else
            INSDHBFREPORT_QR(56) = "N"
        End If
        
        
        '*****************************
        
        
        INSDHBFREPORT_QR.Execute
        
        '***** Shraddha 21/07/10 *****
        If opt_drak.Value = True Then
            UPDDREQDRNAME_QR(0) = "Dr. Ashok Kumar"
        ElseIf opt_drak.Value = False Then
            UPDDREQDRNAME_QR(0) = ""
        End If

        If opt_mryb.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Y. Badrinath"
        ElseIf opt_mryb.Value = False And opt_sitaram.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Sitaram"
        Else
            UPDDREQDRNAME_QR(1) = ""
        End If
        
        UPDDREQDRNAME_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDREQDRNAME_QR.Execute
        
        '*****************************
            
       '*** Inserting into DHBFREPORT
                       
            
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = CMB_PROVISIONAL.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
        
            
             '***** SHWETA 10-02-2014 *****
             
        INSDCBCREPA_QR(0) = req_chglen(cmb_reqno.Text)
        If Len(Trim(TXT_DNAINDEX.Text)) > 0 Then
            INSDCBCREPA_QR(1) = Trim(TXT_DNAINDEX.Text)
        Else
            INSDCBCREPA_QR(1) = ""
        End If
        
            If Len(Trim(TXT_DNACLOIDY.Text)) > 0 Then
            INSDCBCREPA_QR(2) = Trim(TXT_DNACLOIDY.Text)
        Else
            INSDCBCREPA_QR(2) = ""
        End If
        INSDCBCREPA_QR.Execute
        '*****************************
    End If
    If flagsave = 2 Then
            
            '*** updating the DHBFREPORT
            
            
        If Len(Trim(txt_cellcount.Text)) > 0 Then
            UPDDHBFREPORT_QR(0) = Null
            UPDDHBFREPORT_QR(0) = Trim(txt_cellcount.Text)
        Else
            UPDDHBFREPORT_QR(0) = Null
        End If
        If Len(Trim(txt_cytomorphology.Text)) > 0 Then
            UPDDHBFREPORT_QR(1) = Null
            UPDDHBFREPORT_QR(1) = Trim(txt_cytomorphology.Text)
        Else
            UPDDHBFREPORT_QR(1) = Null
        End If
        If Len(Trim(txt_cytochemistry.Text)) > 0 Then
            UPDDHBFREPORT_QR(2) = Null
            UPDDHBFREPORT_QR(2) = Trim(txt_cytochemistry.Text)
        Else
            UPDDHBFREPORT_QR(2) = Null
        End If
        If Len(Trim(txt_surfacemarkers.Text)) > 0 Then
            UPDDHBFREPORT_QR(3) = Null
            UPDDHBFREPORT_QR(3) = Trim(txt_surfacemarkers.Text)
        Else
            UPDDHBFREPORT_QR(3) = Null
        End If
        If Len(Trim(txt_impression.Text)) > 0 Then
            UPDDHBFREPORT_QR(4) = Null
            UPDDHBFREPORT_QR(4) = Trim(txt_impression.Text)
        Else
            UPDDHBFREPORT_QR(4) = Null
        End If
        
        'UPDDHBFREPORT_QR(5) = req_chglen(cmb_reqno.Text)
        
        'SUNDEEP
        '*************** surface Marker

  'SUNDEEP

           UPDDHBFREPORT_QR(5) = Null  'fld1
           UPDDHBFREPORT_QR(6) = Null  'fld2
            
        If Len(Trim(txt_cd13.Text)) > 0 Then
            UPDDHBFREPORT_QR(7) = Trim(txt_cd13.Text)
        Else
             UPDDHBFREPORT_QR(7) = Null
        End If
        
        If Len(Trim(txt_cd14.Text)) > 0 Then
             UPDDHBFREPORT_QR(8) = Trim(txt_cd14.Text)
        Else
            UPDDHBFREPORT_QR(8) = Null
        End If
        
        If Len(Trim(txt_CD15.Text)) > 0 Then
            UPDDHBFREPORT_QR(9) = Trim(txt_CD15.Text)
        Else
            UPDDHBFREPORT_QR(9) = Null
        End If
        
        If Len(Trim(txt_CD33.Text)) > 0 Then
            UPDDHBFREPORT_QR(10) = Trim(txt_CD33.Text)
        Else
            UPDDHBFREPORT_QR(10) = Null
        End If
        
        If Len(Trim(txt_cd36.Text)) > 0 Then
            UPDDHBFREPORT_QR(11) = Trim(txt_cd36.Text)
        Else
            UPDDHBFREPORT_QR(11) = Null
        End If
        
        If Len(Trim(txt_cd41.Text)) > 0 Then
            UPDDHBFREPORT_QR(12) = Trim(txt_cd41.Text)
        Else
            UPDDHBFREPORT_QR(12) = Null
        End If
        
        If Len(Trim(txt_Glycophorina.Text)) > 0 Then
            UPDDHBFREPORT_QR(13) = Trim(txt_Glycophorina.Text)
        Else
            UPDDHBFREPORT_QR(13) = Null
        End If
        
        If Len(Trim(TXT_CD68.Text)) > 0 Then
            UPDDHBFREPORT_QR(14) = Trim(TXT_CD68.Text)
        Else
            UPDDHBFREPORT_QR(14) = Null
        End If
        
        If Len(Trim(TXT_CD34.Text)) > 0 Then
            UPDDHBFREPORT_QR(15) = Trim(TXT_CD34.Text)
        Else
            UPDDHBFREPORT_QR(15) = Null
        End If
        
        If Len(Trim(TXT_HLADR.Text)) > 0 Then
            UPDDHBFREPORT_QR(16) = Trim(TXT_HLADR.Text)
        Else
            UPDDHBFREPORT_QR(16) = Null
        End If
        
        If Len(Trim(TXT_CD38.Text)) > 0 Then
            UPDDHBFREPORT_QR(17) = Trim(TXT_CD38.Text)
        Else
            UPDDHBFREPORT_QR(17) = Null
        End If
        
        If Len(Trim(TXT_CD138.Text)) > 0 Then
            UPDDHBFREPORT_QR(18) = Trim(TXT_CD138.Text)
        Else
           UPDDHBFREPORT_QR(18) = Null
        End If
        
        If Len(Trim(TXT_TDT.Text)) > 0 Then
            UPDDHBFREPORT_QR(19) = Trim(TXT_TDT.Text)
        Else
            UPDDHBFREPORT_QR(19) = Null
        End If
        
        If Len(Trim(TXT_ANTIMPO.Text)) > 0 Then
            UPDDHBFREPORT_QR(20) = Trim(TXT_ANTIMPO.Text)
        Else
            UPDDHBFREPORT_QR(20) = Null
        End If
        
        If Len(Trim(TXT_CD10.Text)) > 0 Then
            UPDDHBFREPORT_QR(21) = Trim(TXT_CD10.Text)
        Else
            UPDDHBFREPORT_QR(21) = Null
        End If
        
        If Len(Trim(TXT_CD19.Text)) > 0 Then
            UPDDHBFREPORT_QR(22) = Trim(TXT_CD19.Text)
        Else
            UPDDHBFREPORT_QR(22) = Null
        End If
        
        If Len(Trim(TXT_FMC7.Text)) > 0 Then
            UPDDHBFREPORT_QR(23) = Trim(TXT_FMC7.Text)
        Else
            UPDDHBFREPORT_QR(23) = Null
        End If
        
        If Len(Trim(TXT_CD22.Text)) > 0 Then
            UPDDHBFREPORT_QR(24) = Trim(TXT_CD22.Text)
        Else
           UPDDHBFREPORT_QR(24) = Null
        End If
        
        If Len(Trim(TXT_CD23.Text)) > 0 Then
            UPDDHBFREPORT_QR(25) = Trim(TXT_CD23.Text)
        Else
            UPDDHBFREPORT_QR(25) = Null
        End If
        
        If Len(Trim(txt_smig.Text)) > 0 Then
            UPDDHBFREPORT_QR(26) = Trim(txt_smig.Text)
        Else
            UPDDHBFREPORT_QR(26) = Null
        End If
        
        If Len(Trim(txt_CD79B.Text)) > 0 Then
            UPDDHBFREPORT_QR(27) = Trim(txt_CD79B.Text)
        Else
            UPDDHBFREPORT_QR(27) = Null
        End If
        
        If Len(Trim(TXT_CD25.Text)) > 0 Then
            UPDDHBFREPORT_QR(28) = Trim(TXT_CD25.Text)
        Else
            UPDDHBFREPORT_QR(28) = Null
        End If
        
        If Len(Trim(TXT_CD11C.Text)) > 0 Then
            UPDDHBFREPORT_QR(29) = Trim(TXT_CD11C.Text)
        Else
            UPDDHBFREPORT_QR(29) = Null
        End If
        
        If Len(Trim(txt_cd103.Text)) > 0 Then
            UPDDHBFREPORT_QR(30) = Trim(txt_cd103.Text)
        Else
            UPDDHBFREPORT_QR(30) = Null
        End If
        
        If Len(Trim(txt_kappa.Text)) > 0 Then
            UPDDHBFREPORT_QR(31) = Trim(txt_kappa.Text)
        Else
            UPDDHBFREPORT_QR(31) = Null
        End If
        
        If Len(Trim(txt_lanbda.Text)) > 0 Then
            UPDDHBFREPORT_QR(32) = Trim(txt_lanbda.Text)
        Else
            UPDDHBFREPORT_QR(32) = Null
        End If
        
        If Len(Trim(TXT_DBA44.Text)) > 0 Then
            UPDDHBFREPORT_QR(33) = Trim(TXT_DBA44.Text)
        Else
            UPDDHBFREPORT_QR(33) = Null
        End If
        
        If Len(Trim(txt_cytocd221(1))) > 0 Then
            UPDDHBFREPORT_QR(34) = Trim(txt_cytocd221(1))
        Else
            UPDDHBFREPORT_QR(34) = Null
        End If
        
        If Len(Trim(txt_cytocd79a.Text)) > 0 Then
            UPDDHBFREPORT_QR(35) = Trim(txt_cytocd79a.Text)
        Else
            UPDDHBFREPORT_QR(35) = Null
        End If
        
        If Len(Trim(TXT_CYTOMU.Text)) > 0 Then
            UPDDHBFREPORT_QR(36) = Trim(TXT_CYTOMU.Text)
        Else
           UPDDHBFREPORT_QR(36) = Null
        End If
        
        If Len(Trim(TXT_CD2.Text)) > 0 Then
            UPDDHBFREPORT_QR(37) = Trim(TXT_CD2.Text)
        Else
            UPDDHBFREPORT_QR(37) = Null
        End If
        
        If Len(Trim(TXT_CD3.Text)) > 0 Then
            UPDDHBFREPORT_QR(38) = Trim(TXT_CD3.Text)
        Else
           UPDDHBFREPORT_QR(38) = Null
        End If
        
        If Len(Trim(TXT_CD4.Text)) > 0 Then
            UPDDHBFREPORT_QR(39) = Trim(TXT_CD4.Text)
        Else
            UPDDHBFREPORT_QR(39) = Null
        End If
        
        If Len(Trim(TXT_CD5.Text)) > 0 Then
            UPDDHBFREPORT_QR(40) = Trim(TXT_CD5.Text)
        Else
            UPDDHBFREPORT_QR(40) = Null
        End If
        
        If Len(Trim(TXT_CD7.Text)) > 0 Then
            UPDDHBFREPORT_QR(41) = Trim(TXT_CD7.Text)
        Else
           UPDDHBFREPORT_QR(41) = Null
        End If
        
        If Len(Trim(TXT_CD8.Text)) > 0 Then
            UPDDHBFREPORT_QR(42) = Trim(TXT_CD8.Text)
        Else
            UPDDHBFREPORT_QR(42) = Null
        End If
        
        If Len(Trim(TXT_CYTOCD3.Text)) > 0 Then
            UPDDHBFREPORT_QR(43) = Trim(TXT_CYTOCD3.Text)
        Else
            UPDDHBFREPORT_QR(43) = Null
        End If
        
        If Len(Trim(TXT_CD42.Text)) > 0 Then
            UPDDHBFREPORT_QR(44) = Trim(TXT_CD42.Text)
        Else
           UPDDHBFREPORT_QR(44) = Null
        End If
        
        If Len(Trim(TXT_CD117.Text)) > 0 Then
            UPDDHBFREPORT_QR(45) = Trim(TXT_CD117.Text)
        Else
           UPDDHBFREPORT_QR(45) = Null
        End If
        
        If Len(Trim(TXT_CD20.Text)) > 0 Then
            UPDDHBFREPORT_QR(46) = Trim(TXT_CD20.Text)
        Else
           UPDDHBFREPORT_QR(46) = Null
        End If
        
        If Len(Trim(TXT_CD61.Text)) > 0 Then
            UPDDHBFREPORT_QR(47) = Trim(TXT_CD61.Text)
        Else
           UPDDHBFREPORT_QR(47) = Null
        End If
        If Len(Trim(TXT_CD56.Text)) > 0 Then
            UPDDHBFREPORT_QR(48) = Trim(TXT_CD56.Text)
        Else
           UPDDHBFREPORT_QR(48) = Null
        End If
        
        If Len(Trim(txt_igm.Text)) > 0 Then
            UPDDHBFREPORT_QR(49) = Trim(txt_igm.Text)
        Else
           UPDDHBFREPORT_QR(49) = Null
        End If

        
        
        '***** Shraddha 21/07/10 *****
        If Len(Trim(txt_tcrab.Text)) > 0 Then
            UPDDHBFREPORT_QR(50) = Trim(txt_tcrab.Text)
        Else
            UPDDHBFREPORT_QR(50) = Null
        End If
        If Len(Trim(txt_tcrag.Text)) > 0 Then
            UPDDHBFREPORT_QR(51) = Trim(txt_tcrag.Text)
        Else
            UPDDHBFREPORT_QR(51) = Null
        End If
        If Len(Trim(txt_clnd1.Text)) > 0 Then
            UPDDHBFREPORT_QR(52) = Trim(txt_clnd1.Text)
        Else
            UPDDHBFREPORT_QR(52) = Null
        End If
        If Len(Trim(TXT_SMPCTG.Text)) > 0 Then
            UPDDHBFREPORT_QR(53) = Trim(TXT_SMPCTG.Text)
        Else
            UPDDHBFREPORT_QR(53) = Null
        End If
        If opt_lsp.Value = True Then
            UPDDHBFREPORT_QR(54) = "Y"
        Else
            UPDDHBFREPORT_QR(54) = "N"
        End If
        If opt_ssp.Value = True Then
            UPDDHBFREPORT_QR(55) = "Y"
        Else
            UPDDHBFREPORT_QR(55) = "N"
        End If
        '*****************************
        UPDDHBFREPORT_QR(56) = req_chglen(cmb_reqno.Text)
       
        
        UPDDHBFREPORT_QR.Execute
        
        '***** Shraddha 21/07/10 *****
        If opt_drak.Value = True Then
            UPDDREQDRNAME_QR(0) = "Dr. Ashok Kumar"
        ElseIf opt_drak.Value = False Then
            UPDDREQDRNAME_QR(0) = ""
        End If

        If opt_mryb.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Y. Badrinath"
        ElseIf opt_mryb.Value = False And opt_sitaram.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Sitaram"
        Else
            UPDDREQDRNAME_QR(1) = ""
        End If
        
        UPDDREQDRNAME_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDREQDRNAME_QR.Execute
        
        
        
         '***** SHWETA 10-02-2014 *****
             
'        INSDCBCREPA_QR(0) = req_chglen(cmb_reqno.Text)
'        If Len(Trim(TXT_DNAINDEX.Text)) > 0 Then
'            INSDCBCREPA_QR(1) = Trim(TXT_DNAINDEX.Text)
'        Else
'            INSDCBCREPA_QR(1) = ""
'        End If
'
'            If Len(Trim(TXT_DNACLOIDY.Text)) > 0 Then
'            INSDCBCREPA_QR(2) = Trim(TXT_DNACLOIDY.Text)
'        Else
'            INSDCBCREPA_QR(2) = ""
'        End If
'        INSDCBCREPA_QR.Execute
        '*****************************
        
        
        
        '*****************************
        '***** Added on 10-02-2014 by SHWETA
        UPDDCBCREPA_QR(0) = Trim(TXT_DNAINDEX.Text)
        UPDDCBCREPA_QR(1) = Trim(TXT_DNACLOIDY.Text)
        UPDDCBCREPA_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDCBCREPA_QR.Execute
        ''''till here
        
        
        
            '*** updating the DHBFREPORT
        If ADDENDUMFALG = False Then
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = CMB_PROVISIONAL.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
        End If
            '**** UPDATING INTO FREPORT
    
    If ADDENDUMFALG = True Then
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
MsgBox "Record is Saved"
btnenabl ("INI")

Exit Sub
MSG:
    rdb.RollbackTrans
MsgBox Err.Description
End Sub
Public Sub ENABL(a As String)
    If a = "DIS" Then
        txt_nmaterial.Enabled = False
        txt_cellcount.Enabled = False
        txt_cytomorphology.Enabled = False
        txt_cytochemistry.Enabled = False
        txt_surfacemarkers.Enabled = False
        txt_impression.Enabled = False
        cmb_enterby.Enabled = False
        CMB_PROVISIONAL.Enabled = False
        cmb_finalby.Enabled = False
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = False
        txt_cd13.Enabled = False
        txt_cd14.Enabled = False
        txt_CD15.Enabled = False
        txt_CD33.Enabled = False
        txt_cd36.Enabled = False
        txt_cd41.Enabled = False
        TXT_CD42.Enabled = False
        TXT_CD117.Enabled = False
        TXT_CD20.Enabled = False
        TXT_CD61.Enabled = False
        TXT_CD56.Enabled = False
        txt_Glycophorina.Enabled = False
        TXT_CD68.Enabled = False
        TXT_CD34.Enabled = False
        TXT_HLADR.Enabled = False
        TXT_CD38.Enabled = False
        TXT_CD138.Enabled = False
        TXT_TDT.Enabled = False
        TXT_ANTIMPO.Enabled = False
        TXT_CD10.Enabled = False
        TXT_CD19.Enabled = False
        TXT_FMC7.Enabled = False
        TXT_CD22.Enabled = False
        TXT_CD23.Enabled = False
        txt_smig.Enabled = False
        txt_CD79B.Enabled = False
        TXT_CD25.Enabled = False
        TXT_CD11C.Enabled = False
        txt_cd103.Enabled = False
        txt_kappa.Enabled = False
        txt_lanbda.Enabled = False
        TXT_DBA44.Enabled = False
        txt_cytocd221(1).Enabled = False
        txt_cytocd79a.Enabled = False
        TXT_CYTOMU.Enabled = False
        TXT_CD2.Enabled = False
        TXT_CD3.Enabled = False
        TXT_CD4.Enabled = False
        TXT_CD5.Enabled = False
        TXT_CD7.Enabled = False
        TXT_CD8.Enabled = False
        TXT_CYTOCD3.Enabled = False
        txt_impression.Enabled = False
        cmb_enterby.Enabled = False
        'cmb_finalby.Enabled = False
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
        
        '***** Shraddha 21/07/10 *****
        txt_igm.Enabled = False
        TXT_TDT.Enabled = False
        txt_smig.Enabled = False
        txt_tcrab.Enabled = False
        txt_tcrag.Enabled = False
        txt_clnd1.Enabled = False
        TXT_SMPCTG.Enabled = False
        opt_lsp.Enabled = False
        opt_ssp.Enabled = False
        opt_drak.Enabled = False
        opt_mryb.Enabled = False
        opt_sitaram.Enabled = False
        '*****************************
        
    
    End If
    If a = "ENA" Then
        txt_nmaterial.Enabled = True
        txt_cellcount.Enabled = True
        txt_cytomorphology.Enabled = True
        txt_cytochemistry.Enabled = True
        txt_surfacemarkers.Enabled = True
        txt_impression.Enabled = True
        cmb_enterby.Enabled = True
        CMB_PROVISIONAL.Enabled = True
        cmb_finalby.Enabled = True
        txt_cd13.Enabled = True
        txt_cd14.Enabled = True
        txt_CD15.Enabled = True
        txt_CD33.Enabled = True
        txt_cd36.Enabled = True
        txt_cd41.Enabled = True
        TXT_CD42.Enabled = True
        TXT_CD117.Enabled = True
        TXT_CD20.Enabled = True
        TXT_CD61.Enabled = True
        TXT_CD56.Enabled = True
        txt_Glycophorina.Enabled = True
        TXT_CD68.Enabled = True
        TXT_CD34.Enabled = True
        TXT_HLADR.Enabled = True
        TXT_CD38.Enabled = True
        TXT_CD138.Enabled = True
        TXT_TDT.Enabled = True
        TXT_ANTIMPO.Enabled = True
        TXT_CD10.Enabled = True
        TXT_CD19.Enabled = True
        TXT_FMC7.Enabled = True
        TXT_CD22.Enabled = True
        TXT_CD23.Enabled = True
        txt_smig.Enabled = True
        txt_CD79B.Enabled = True
        TXT_CD25.Enabled = True
        TXT_CD11C.Enabled = True
        txt_cd103.Enabled = True
        txt_kappa.Enabled = True
        txt_lanbda.Enabled = True
        TXT_DBA44.Enabled = True
        txt_cytocd221(1).Enabled = True
        txt_cytocd79a.Enabled = True
        TXT_CYTOMU.Enabled = True
        TXT_CD2.Enabled = True
        TXT_CD3.Enabled = True
        TXT_CD4.Enabled = True
        TXT_CD5.Enabled = True
        TXT_CD7.Enabled = True
        TXT_CD8.Enabled = True
        TXT_CYTOCD3.Enabled = True
        txt_impression.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        '***** Shraddha 21/07/10 *****
        txt_igm.Enabled = True
        TXT_TDT.Enabled = True
        txt_smig.Enabled = True
        txt_tcrab.Enabled = True
        txt_tcrag.Enabled = True
        txt_clnd1.Enabled = True
        TXT_SMPCTG.Enabled = True
        opt_lsp.Enabled = True
        opt_ssp.Enabled = True
        opt_drak.Enabled = True
        opt_mryb.Enabled = True
        opt_sitaram.Enabled = True
        '*****************************
        'cmb_addendumby.Enabled = False
        'txt_addendum.Enabled = False
    End If
End Sub





''Private Sub TXT_IMPRESSION_GotFocus()
''On Error GoTo m1:
''If newflag = True Then
''    TXT_ECIMPRESSION.Text = "With in Normal Limits"
''End If
''w_impression.Visible = False
''Set d_impression = w_impression.ActiveDocument
''d_impression.Content = ""
''d_impression.Content = TXT_ECIMPRESSION
''w_impression.Caption = "Impression"
'''w_frozen.WindowState = wdWindowStateMaximize
''w_impression.WindowState = wdWindowStateNormal
''d_impression.Content = TXT_ECIMPRESSION
''w_impression.Visible = True
''w_impression.Activate
''m1:
''On Error Resume Next
''Set w_impression = CreateObject("Word.Application")
''w_impression.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
''w_impression.Caption = "Impression"
'''w_frozen.WindowState = wdWindowStateMaximize
''w_impression.WindowState = wdWindowStateNormal
''Set d_impression = w_impression.ActiveDocument
''d_impression.Content = Trim(TXT_ECIMPRESSION)
''w_impression.Visible = True
''w_impression.Activate
''End Sub
''
''Private Sub TXT_ECIMPRESSION_LostFocus()
''On Error GoTo m2:
''Set d_impression = w_impression.ActiveDocument
''TXT_ECIMPRESSION.Text = d_impression.Content
'''flagimp = False
''d_impression.Content = ""
''d_impression.SaveAs "C:\fronzen.DOC"
''d_impression.Close
''w_impression.Quit
'''W1.Visible = False
''Exit Sub
''m2:
''MsgBox Err.Description
''End Sub









Private Sub Text6_Change()

End Sub



Private Sub txt_surfacemarkers_LostFocus()
'MsgBox (Len(txt_surfacemarkers.Text))
End Sub
