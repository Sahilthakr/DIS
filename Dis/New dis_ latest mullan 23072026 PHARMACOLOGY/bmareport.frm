VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form bmareport 
   Caption         =   "Bone Marrow Aspiration Report"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14625
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
   ScaleHeight     =   11800.88
   ScaleMode       =   0  'User
   ScaleWidth      =   1.18477e6
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Surface Markers"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      MaskColor       =   &H8000000A&
      TabIndex        =   249
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame Frame15 
      Caption         =   "DNA"
      Height          =   855
      Left            =   120
      TabIndex        =   245
      Top             =   7800
      Width           =   11655
      Begin VB.TextBox TXT_DNAINDEX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   247
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TXT_DNACLOIDY 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   3360
         TabIndex        =   246
         Top             =   240
         Width           =   8055
      End
      Begin VB.Label Label13 
         Caption         =   "INDEX"
         Height          =   255
         Left            =   0
         TabIndex        =   248
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.ComboBox cmb_enterby 
      Height          =   315
      Left            =   11880
      TabIndex        =   221
      Top             =   360
      Width           =   2295
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10680
      TabIndex        =   104
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9480
      TabIndex        =   103
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame12 
      Height          =   855
      Left            =   0
      TabIndex        =   101
      Top             =   8880
      Width           =   13335
      Begin VB.ComboBox regcbo 
         Height          =   315
         Left            =   960
         TabIndex        =   223
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   135
         Top             =   120
         Width           =   3495
      End
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   4560
         TabIndex        =   133
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txt_IMPRESSION 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   120
         Width           =   1935
      End
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         ItemData        =   "bmareport.frx":0000
         Left            =   9000
         List            =   "bmareport.frx":0002
         TabIndex        =   33
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TXT_COMMENTS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label15 
         Caption         =   "Registrar"
         Height          =   255
         Left            =   120
         TabIndex        =   222
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Addendum                                                                  Comments"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3360
         TabIndex        =   134
         Top             =   240
         Width           =   13215
      End
      Begin VB.Label Label4 
         Caption         =   "Addendum By"
         Height          =   255
         Left            =   3360
         TabIndex        =   132
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label LBL_FINALDIAGNOSIS 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
         Height          =   195
         Left            =   7080
         TabIndex        =   102
         Top             =   600
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   81
      Top             =   3480
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483646
      ForeColor       =   192
      TabCaption(0)   =   "MORPHOLOGY"
      TabPicture(0)   =   "bmareport.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CYTOCHEMISTRY"
      TabPicture(1)   =   "bmareport.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "SURFACE MARKERS"
      TabPicture(2)   =   "bmareport.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).ControlCount=   2
      Begin TabDlg.SSTab SSTab2 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   119
         Top             =   480
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   7011
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Screen I"
         TabPicture(0)   =   "bmareport.frx":0058
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame9"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Screen II"
         TabPicture(1)   =   "bmareport.frx":0074
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame10"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Screen III"
         TabPicture(2)   =   "bmareport.frx":0090
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame11"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame11 
            Height          =   3495
            Left            =   -74640
            TabIndex        =   208
            Top             =   360
            Width           =   10815
            Begin VB.Frame Frame14 
               Height          =   615
               Left            =   2640
               TabIndex        =   239
               Top             =   2040
               Width           =   7215
               Begin VB.OptionButton opt_sitaram 
                  Caption         =   "Mr. Sitaram"
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   244
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.OptionButton opt_mryb 
                  Caption         =   "Mr. Y. Badrinath"
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   241
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.OptionButton opt_drak 
                  Caption         =   "Dr. Ashok Kumar"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   240
                  Top             =   240
                  Width           =   2055
               End
            End
            Begin VB.TextBox TXT_SMPCTG 
               Height          =   285
               Left            =   4680
               TabIndex        =   235
               Top             =   1440
               Width           =   615
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
               TabIndex        =   234
               Top             =   1440
               Width           =   3015
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
               TabIndex        =   233
               Top             =   1800
               Width           =   2415
            End
            Begin VB.TextBox txt_tcrag 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5400
               TabIndex        =   228
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox txt_tcrab 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1680
               TabIndex        =   227
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox txt_cyclinD1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8520
               TabIndex        =   226
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txt_igm 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1680
               TabIndex        =   219
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_smig 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8520
               TabIndex        =   187
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TXT_TDT 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5400
               TabIndex        =   186
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label22 
               Caption         =   "Instrument: BD-FACS Calibur E 2237"
               Height          =   255
               Left            =   240
               TabIndex        =   238
               Top             =   2760
               Width           =   3615
            End
            Begin VB.Label Label1 
               Caption         =   "Flow Cytometry is Done By "
               Height          =   255
               Left            =   240
               TabIndex        =   237
               Top             =   2280
               Width           =   2295
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H000000C0&
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   2
               Height          =   1935
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   1320
               Width           =   10575
            End
            Begin VB.Label Label16 
               Caption         =   "The above positive percentage were determined on           % of the gated events on"
               Height          =   255
               Left            =   240
               TabIndex        =   236
               Top             =   1440
               Width           =   7215
            End
            Begin VB.Label Label21 
               Caption         =   "TCR"
               Height          =   255
               Left            =   4080
               TabIndex        =   232
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label20 
               Caption         =   "TCR"
               Height          =   255
               Left            =   720
               TabIndex        =   231
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label19 
               Caption         =   "gd"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4560
               TabIndex        =   230
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label18 
               Caption         =   "ab"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   229
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label17 
               Caption         =   "Cyclin D1"
               Height          =   255
               Left            =   7680
               TabIndex        =   225
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "IgM                                                  Tdt                                                       SmIg"
               Height          =   195
               Left            =   1080
               TabIndex        =   218
               Top             =   240
               Width           =   7335
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   1.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   2175
            Left            =   -74880
            TabIndex        =   124
            Top             =   360
            Width           =   11055
            Begin VB.TextBox TXT_CD61 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   164
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox TXT_CD34 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   159
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txt_kappa 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   183
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox TXT_HLADR 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   182
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox txt_Glycophorina 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   181
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox TXT_ANTIMPO 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   174
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD138 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   173
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD117 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   172
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD38 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   160
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox TXT_CD68 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   165
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD56 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   163
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TXT_CD42 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   162
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_cd41 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   161
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TXT_DBA44 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   178
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox TXT_FMC7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   179
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txt_CD79B 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   169
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txt_lanbda 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   185
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox txt_cd103 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   171
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txt_cytocd79a 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9480
               TabIndex        =   167
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TXT_CYTOMU 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   176
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "CD 61"
               Height          =   255
               Left            =   3240
               TabIndex        =   210
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lbl_cd34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CD 34"
               Height          =   195
               Left            =   720
               TabIndex        =   207
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl_kappa 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kappa"
               Height          =   195
               Left            =   6000
               TabIndex        =   205
               Top             =   1680
               Width           =   555
            End
            Begin VB.Label LBL_HALDR 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HLA DR"
               Height          =   195
               Left            =   3120
               TabIndex        =   204
               Top             =   1680
               Width           =   705
            End
            Begin VB.Label LBL_GLYCOPHORIN 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Glycophorin A"
               Height          =   195
               Left            =   120
               TabIndex        =   203
               Top             =   1680
               Width           =   1200
            End
            Begin VB.Label LBL_ANTIMPO 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AntiMPO"
               Height          =   195
               Left            =   480
               TabIndex        =   202
               Top             =   1320
               Width           =   750
            End
            Begin VB.Label LBL_CD138 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " CD138"
               Height          =   195
               Left            =   8760
               TabIndex        =   201
               Top             =   960
               Width           =   645
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "CD117"
               Height          =   255
               Left            =   6000
               TabIndex        =   200
               Top             =   960
               Width           =   735
            End
            Begin VB.Label LBL_CD38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CD 38"
               Height          =   195
               Left            =   3240
               TabIndex        =   199
               Top             =   240
               Width           =   540
            End
            Begin VB.Label LBL_CD68 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cyto CD68"
               Height          =   195
               Left            =   5760
               TabIndex        =   197
               Top             =   600
               Width           =   915
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "CD56"
               Height          =   255
               Left            =   720
               TabIndex        =   196
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "CD42"
               Height          =   255
               Left            =   8880
               TabIndex        =   195
               Top             =   240
               Width           =   615
            End
            Begin VB.Label LBL_CD41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CD41"
               Height          =   195
               Left            =   6120
               TabIndex        =   194
               Top             =   240
               Width           =   480
            End
            Begin VB.Label LBL_FMC7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FMC7"
               Height          =   195
               Left            =   8760
               TabIndex        =   131
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label LBL_CD79B 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CD79 b"
               Height          =   195
               Left            =   600
               TabIndex        =   130
               Top             =   960
               Width           =   645
            End
            Begin VB.Label lbl_DBA44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DBA 44"
               Enabled         =   0   'False
               Height          =   195
               Left            =   6000
               TabIndex        =   129
               Top             =   1320
               Width           =   660
            End
            Begin VB.Label lbl_lanbda 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lambda"
               Height          =   195
               Left            =   8760
               TabIndex        =   128
               Top             =   1680
               Width           =   675
            End
            Begin VB.Label lbl_cd103 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CD103"
               Height          =   195
               Left            =   3210
               TabIndex        =   127
               Top             =   960
               Width           =   585
            End
            Begin VB.Label lbl_cytocd79a 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CytoCD79a"
               Height          =   195
               Left            =   8445
               TabIndex        =   126
               Top             =   600
               Width           =   960
            End
            Begin VB.Label LBL_CYTOMU 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cyto mu"
               Height          =   195
               Left            =   3090
               TabIndex        =   125
               Top             =   1320
               Width           =   690
            End
         End
         Begin VB.Frame Frame9 
            ForeColor       =   &H00800000&
            Height          =   2295
            Left            =   120
            TabIndex        =   120
            Top             =   360
            Width           =   11055
            Begin VB.TextBox TXT_CD8 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   142
               Top             =   615
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD20 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   153
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txt_cytocd22 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   8880
               TabIndex        =   154
               Top             =   1320
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TXT_CYTOCD3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   138
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_CD33 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   158
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD25 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   157
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD23 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   156
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox txt_cytocd221 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   1080
               TabIndex        =   155
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD22 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   188
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD19 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3585
               TabIndex        =   152
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txt_cd36 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   151
               Top             =   1320
               Width           =   1440
            End
            Begin VB.TextBox TXT_CD11C 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   144
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD10 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   8880
               TabIndex        =   143
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   141
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   139
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   140
               Top             =   600
               Width           =   1440
            End
            Begin VB.TextBox TXT_CD3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   137
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TXT_CD2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   136
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_cd13 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   146
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txt_cd14 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6240
               TabIndex        =   148
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txt_CD15 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8880
               TabIndex        =   150
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label7 
               Caption         =   "CD 20"
               Height          =   255
               Left            =   5640
               TabIndex        =   209
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label LBL_CYTOCD22 
               AutoSize        =   -1  'True
               Caption         =   "CD22"
               Height          =   195
               Left            =   8280
               TabIndex        =   206
               Top             =   1320
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label LBL_CYTOCD3 
               AutoSize        =   -1  'True
               Caption         =   "Cyto CD3"
               Height          =   195
               Left            =   5280
               TabIndex        =   198
               Top             =   240
               Width           =   810
            End
            Begin VB.Label LBL_CD33 
               AutoSize        =   -1  'True
               Caption         =   "CD 33"
               Height          =   195
               Left            =   8280
               TabIndex        =   193
               Top             =   1680
               Width           =   540
            End
            Begin VB.Label LBL_CD25 
               AutoSize        =   -1  'True
               Caption         =   "CD 25"
               Height          =   195
               Left            =   5640
               TabIndex        =   192
               Top             =   1680
               Width           =   540
            End
            Begin VB.Label LBL_CD23 
               AutoSize        =   -1  'True
               Caption         =   "CD 23"
               Height          =   195
               Left            =   3000
               TabIndex        =   191
               Top             =   1680
               Width           =   540
            End
            Begin VB.Label lbl_cyto221 
               AutoSize        =   -1  'True
               Caption         =   "Cyto CD22"
               Height          =   195
               Left            =   120
               TabIndex        =   190
               Top             =   1680
               Width           =   1035
            End
            Begin VB.Label LBL_CD22 
               AutoSize        =   -1  'True
               Caption         =   "CD22"
               Height          =   195
               Left            =   8265
               TabIndex        =   189
               Top             =   1320
               Width           =   480
            End
            Begin VB.Label LBL_CD19 
               AutoSize        =   -1  'True
               Caption         =   "CD 19"
               Height          =   195
               Left            =   3000
               TabIndex        =   184
               Top             =   1320
               Width           =   540
            End
            Begin VB.Label LBL_CD36 
               Caption         =   "CD16"
               Height          =   255
               Left            =   480
               TabIndex        =   180
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label LBL_CD11C 
               AutoSize        =   -1  'True
               Caption         =   "CD11c"
               Height          =   195
               Left            =   480
               TabIndex        =   177
               Top             =   960
               Width           =   705
            End
            Begin VB.Label LBL_CD10 
               AutoSize        =   -1  'True
               Caption         =   "CD 10"
               Height          =   195
               Left            =   8280
               TabIndex        =   175
               Top             =   600
               Width           =   540
            End
            Begin VB.Label LBL_CD8 
               AutoSize        =   -1  'True
               Caption         =   "CD 8"
               Height          =   195
               Left            =   5640
               TabIndex        =   170
               Top             =   600
               Width           =   435
            End
            Begin VB.Label LBL_CD7 
               AutoSize        =   -1  'True
               Caption         =   "CD 7"
               Height          =   195
               Left            =   3000
               TabIndex        =   168
               Top             =   600
               Width           =   435
            End
            Begin VB.Label LBL_CD4 
               AutoSize        =   -1  'True
               Caption         =   "CD 4"
               Height          =   195
               Left            =   8280
               TabIndex        =   166
               Top             =   240
               Width           =   435
            End
            Begin VB.Label LBL_CD5 
               AutoSize        =   -1  'True
               Caption         =   "CD 5"
               Height          =   195
               Left            =   480
               TabIndex        =   149
               Top             =   600
               Width           =   555
            End
            Begin VB.Label LBL_CD3 
               AutoSize        =   -1  'True
               Caption         =   "CD 3"
               Height          =   195
               Left            =   3000
               TabIndex        =   147
               Top             =   240
               Width           =   435
            End
            Begin VB.Label LBL_CD2 
               AutoSize        =   -1  'True
               Caption         =   "CD 2"
               Height          =   195
               Left            =   480
               TabIndex        =   145
               Top             =   240
               Width           =   555
            End
            Begin VB.Label LBL_CD13 
               AutoSize        =   -1  'True
               Caption         =   "CD 13"
               Height          =   195
               Left            =   3000
               TabIndex        =   123
               Top             =   960
               Width           =   540
            End
            Begin VB.Label lbl_cd14 
               AutoSize        =   -1  'True
               Caption         =   "CD 14"
               Height          =   195
               Left            =   5625
               TabIndex        =   122
               Top             =   960
               Width           =   540
            End
            Begin VB.Label lbl_cd15 
               AutoSize        =   -1  'True
               Caption         =   "CD 15"
               Height          =   195
               Left            =   8280
               TabIndex        =   121
               Top             =   960
               Width           =   540
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   84
         Top             =   360
         Width           =   11535
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   12495
         Begin VB.TextBox txt_descmorpho 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   28
            Top             =   2400
            Width           =   8895
         End
         Begin VB.TextBox TXT_IMPRINT 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   30
            Top             =   2760
            Width           =   4095
         End
         Begin VB.TextBox txt_abcellblasts 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7560
            TabIndex        =   29
            Top             =   2760
            Width           =   3615
         End
         Begin VB.Frame Frame13 
            Caption         =   "Plasma Cells (%)"
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   3840
            TabIndex        =   108
            Top             =   1080
            Width           =   4815
            Begin VB.TextBox txt_mature 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_inmature 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2280
               TabIndex        =   21
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_plasmablasts 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4080
               TabIndex        =   23
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_atypbi 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               TabIndex        =   22
               Top             =   600
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lbl_mature 
               AutoSize        =   -1  'True
               Caption         =   "Mature           Immature"
               Height          =   195
               Left            =   150
               TabIndex        =   110
               Top             =   240
               Width           =   2025
            End
            Begin VB.Label lbl_plasmablasts 
               AutoSize        =   -1  'True
               Caption         =   "Plasmablasts"
               Height          =   195
               Left            =   2880
               TabIndex        =   111
               Top             =   240
               Width           =   1110
            End
            Begin VB.Label lbl_atypbin 
               AutoSize        =   -1  'True
               Caption         =   "Atyp/Binucleate"
               Height          =   195
               Left            =   120
               TabIndex        =   109
               Top             =   600
               Visible         =   0   'False
               Width           =   1380
            End
         End
         Begin VB.TextBox TXT_MERATION1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3720
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox TXT_MERATION 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3120
            TabIndex        =   7
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txt_cellularity 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame8 
            Caption         =   "Megakaryocytic Series"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   4440
            TabIndex        =   88
            Top             =   1680
            Width           =   6735
            Begin VB.TextBox txt_morphology 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   27
               Top             =   360
               Width           =   2535
            End
            Begin VB.TextBox txt_cellularity2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   26
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label LBL_MORPHOLOGYMS 
               AutoSize        =   -1  'True
               Caption         =   "Morphology"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   3120
               TabIndex        =   97
               Top             =   360
               Width           =   990
            End
            Begin VB.Label lbl_cellmegseries 
               AutoSize        =   -1  'True
               Caption         =   "Cellularity"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   96
               Top             =   360
               Width           =   840
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Erythroid Series"
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   8760
            TabIndex        =   87
            Top             =   600
            Width           =   3615
            Begin VB.TextBox TXT_ErythCells 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1440
               TabIndex        =   243
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txt_maturation 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               TabIndex        =   25
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txt_cellularity1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   960
               TabIndex        =   24
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "Erythroid Cells"
               Height          =   255
               Left            =   120
               TabIndex        =   242
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label lbl_cell1 
               AutoSize        =   -1  'True
               Caption         =   "Cellularity               Maturation"
               Height          =   195
               Left            =   120
               TabIndex        =   95
               Top             =   360
               Width           =   2640
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Lymphoid Series (%)"
            ForeColor       =   &H00800000&
            Height          =   975
            Left            =   4200
            TabIndex        =   86
            Top             =   0
            Width           =   8175
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3720
               TabIndex        =   217
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   216
               Top             =   600
               Width           =   495
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7440
               TabIndex        =   214
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txt_plasmacells 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5760
               TabIndex        =   19
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txt_atyply 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3960
               TabIndex        =   18
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox txt_matly 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   17
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Prolymphocytes            Hematogones"
               Height          =   195
               Left            =   120
               TabIndex        =   215
               Top             =   600
               Width           =   3210
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Hairy Cells"
               Height          =   195
               Left            =   6360
               TabIndex        =   213
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl_plasmacells 
               AutoSize        =   -1  'True
               Caption         =   "Plasma Cells"
               Height          =   195
               Left            =   4560
               TabIndex        =   94
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label lbl_maturely 
               AutoSize        =   -1  'True
               Caption         =   "Mature Lymphocytes         Atyp Lymphocytes"
               Height          =   195
               Left            =   120
               TabIndex        =   93
               Top             =   240
               Width           =   3825
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Myeloid Series (%)"
            ForeColor       =   &H00800000&
            Height          =   1695
            Left            =   0
            TabIndex        =   85
            Top             =   600
            Width           =   3735
            Begin VB.TextBox TXT_PROMONOCYTES 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3240
               TabIndex        =   16
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox txt_monocytes 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   15
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox txt_eosinophils 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   14
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txt_basophils 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   13
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txt_polymorphs 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   12
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txt_metamyelocytes 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   11
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txt_myelocytes 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   10
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_promyelocytes 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   9
               Top             =   240
               Width           =   375
            End
            Begin VB.Label LBL_PROMONOCYTES 
               AutoSize        =   -1  'True
               Caption         =   "Promonocytes"
               Height          =   195
               Left            =   2040
               TabIndex        =   117
               Top             =   1440
               Width           =   1200
            End
            Begin VB.Label lbl_monocytes 
               AutoSize        =   -1  'True
               Caption         =   "Monocytes"
               Height          =   195
               Left            =   120
               TabIndex        =   92
               Top             =   1440
               Width           =   930
            End
            Begin VB.Label lbl_basophils 
               AutoSize        =   -1  'True
               Caption         =   "Basophils                   Eosinophils"
               Height          =   195
               Left            =   120
               TabIndex        =   91
               Top             =   1080
               Width           =   2925
            End
            Begin VB.Label lbl_metamyelocytes 
               AutoSize        =   -1  'True
               Caption         =   "Metamyelocytes          Polymorphs"
               Height          =   195
               Left            =   120
               TabIndex        =   90
               Top             =   720
               Width           =   2925
            End
            Begin VB.Label lbl_promele 
               AutoSize        =   -1  'True
               Caption         =   "Promyelocytes            Myelocytes"
               Height          =   195
               Left            =   120
               TabIndex        =   89
               Top             =   360
               Width           =   2895
            End
         End
         Begin VB.Label Label9 
            Caption         =   "Descriptive Morphology"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   211
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label lbl_IMprint 
            AutoSize        =   -1  'True
            Caption         =   "IMPRINT"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   2760
            Width           =   795
         End
         Begin VB.Label lbl_abcellblasts 
            AutoSize        =   -1  'True
            Caption         =   "Abnormal Cells / Blasts"
            Height          =   195
            Left            =   5400
            TabIndex        =   113
            Top             =   2880
            Width           =   1980
         End
         Begin VB.Label LBL_COLON 
            AutoSize        =   -1  'True
            Caption         =   ":"
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
            Left            =   3600
            TabIndex        =   107
            Top             =   240
            Width           =   75
         End
         Begin VB.Label lbl_meration 
            AutoSize        =   -1  'True
            Caption         =   "M/E RATIO"
            Height          =   195
            Left            =   2040
            TabIndex        =   106
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lbl_cell 
            AutoSize        =   -1  'True
            Caption         =   "Cellularity"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   82
         Top             =   360
         Width           =   11535
         Begin VB.ComboBox CMB_GIDEROBLASTS 
            Height          =   315
            ItemData        =   "bmareport.frx":00AC
            Left            =   4560
            List            =   "bmareport.frx":00B6
            TabIndex        =   42
            Text            =   "Negative"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox TXT_GIDEROBLASTS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   43
            Top             =   2160
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox TXT_IRONSTAIN 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   41
            Top             =   1440
            Width           =   2775
         End
         Begin VB.ComboBox CMB_IRONSTAIN 
            Height          =   315
            ItemData        =   "bmareport.frx":00CE
            Left            =   6000
            List            =   "bmareport.frx":00D8
            TabIndex        =   40
            Text            =   "Negative"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox CMB_TOLUIDINEBLUE 
            Height          =   315
            ItemData        =   "bmareport.frx":00F0
            Left            =   7320
            List            =   "bmareport.frx":00FA
            TabIndex        =   46
            Text            =   "Negative"
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox CMB_SUDANBLACKB 
            Height          =   315
            ItemData        =   "bmareport.frx":0112
            Left            =   7320
            List            =   "bmareport.frx":011C
            TabIndex        =   44
            Text            =   "Negative"
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox CMB_TRAP 
            Height          =   315
            ItemData        =   "bmareport.frx":0134
            Left            =   6000
            List            =   "bmareport.frx":013E
            TabIndex        =   38
            Text            =   "Negative"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox CMB_NONSPEST 
            Height          =   315
            ItemData        =   "bmareport.frx":0156
            Left            =   1800
            List            =   "bmareport.frx":0163
            TabIndex        =   36
            Text            =   "Negative"
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox CMB_MYELOPEROXIDASE 
            Height          =   315
            ItemData        =   "bmareport.frx":018C
            Left            =   1800
            List            =   "bmareport.frx":019F
            TabIndex        =   34
            Text            =   "Negative"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txt_trapironstain 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   39
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txt_TOLUIDINEBLUE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8880
            TabIndex        =   47
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txt_nonspecificesterase 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   37
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txt_sudanblackb 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8880
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txt_myeloperoxidase 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LBL_GIDEROBLASTS 
            AutoSize        =   -1  'True
            Caption         =   "Sideroblasts"
            Height          =   195
            Left            =   480
            TabIndex        =   118
            Top             =   2160
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label LBL_IRONSTAIN 
            AutoSize        =   -1  'True
            Caption         =   "Iron Stain"
            Height          =   195
            Left            =   600
            TabIndex        =   116
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lbl_trapironstain 
            AutoSize        =   -1  'True
            Caption         =   "TRAP"
            Height          =   195
            Left            =   960
            TabIndex        =   112
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label lbl_nonspest 
            AutoSize        =   -1  'True
            Caption         =   "NSE"
            Height          =   195
            Left            =   960
            TabIndex        =   100
            Top             =   720
            Width           =   390
         End
         Begin VB.Label lbl_sudanBlackb 
            AutoSize        =   -1  'True
            Caption         =   "Sudan Black B"
            Height          =   195
            Left            =   5880
            TabIndex        =   99
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label lbl_myeloper 
            AutoSize        =   -1  'True
            Caption         =   "Myeloperoxidase"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   360
            Width           =   1425
         End
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   0
      TabIndex        =   74
      Top             =   9720
      Width           =   11775
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7080
         MouseIcon       =   "bmareport.frx":01FF
         MousePointer    =   99  'Custom
         TabIndex        =   80
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10440
         MouseIcon       =   "bmareport.frx":0509
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8775
         MouseIcon       =   "bmareport.frx":0813
         MousePointer    =   99  'Custom
         TabIndex        =   78
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5880
         MouseIcon       =   "bmareport.frx":0B1D
         MousePointer    =   99  'Custom
         TabIndex        =   77
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "bmareport.frx":0E27
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "bmareport.frx":1131
         MousePointer    =   99  'Custom
         TabIndex        =   76
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MouseIcon       =   "bmareport.frx":143B
         MousePointer    =   99  'Custom
         TabIndex        =   75
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         MouseIcon       =   "bmareport.frx":1745
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame common 
      Height          =   2895
      Left            =   0
      TabIndex        =   50
      Top             =   600
      Width           =   11775
      Begin VB.CommandButton cmd_emr 
         BackColor       =   &H00FFFF80&
         Caption         =   "&EMR"
         Height          =   255
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Full Report"
         Height          =   375
         Left            =   120
         TabIndex        =   220
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_nmaterial 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   5
         Top             =   2520
         Width           =   3615
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10320
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   64
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   63
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   62
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   61
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   60
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   59
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   58
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   57
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   56
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   55
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   0
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   600
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition to : HAEMATOPATHOLOGY"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1410
         TabIndex        =   73
         Top             =   360
         Width           =   3390
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No.                                           Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   72
         Top             =   720
         Width           =   4680
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "BMA No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9600
         TabIndex        =   71
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   70
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex                                      Age                             Category                                        Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   69
         Top             =   1440
         Width           =   8370
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   $"bmareport.frx":1A4F
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   1800
         Width           =   10035
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality : BONE MARROW ASPIRATION REPORT"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   67
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis                                                                    Nature of Material"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   65
         Top             =   2520
         Width           =   7440
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "TATA MEMORIAL HOSPITAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   450
      TabIndex        =   212
      Top             =   795
      Width           =   11655
   End
   Begin VB.Label LBL_CP 
      AutoSize        =   -1  'True
      Caption         =   "BONE MARROW ASPIRATION REPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3000
      TabIndex        =   49
      Top             =   120
      Width           =   4470
   End
End
Attribute VB_Name = "bmareport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELQR As rdoQuery

Dim blank As Boolean
Dim l As Integer

Dim str As String
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
Dim I As Integer
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
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG1 As Boolean


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

Dim INSDBMAREPORT_QR As rdoQuery
Dim UPDDBMAREPORT_QR As rdoQuery
Dim SELDBMAREPORT_QR As rdoQuery
Dim SELDBMAREPORT_RS As rdoResultset
Dim SELDBMAREPORTFLAG As Boolean

Dim RNAGE As String
Dim SS As String

'Dim d_grossdesc As Document
Dim d_grossdesc As Object
Dim W_grossdes As Application

'Dim d_descmorpho As Document
Dim d_descmorpho As Object
Dim W_descmorpho As Application

'Dim d_ADDENDUM As Document
Dim d_ADDENDUM As Object
Dim W_ADDENDUM As Application


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

''******SECURITY***********

''****ADDENDUM**************
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset

Dim seldissecentr11_QR As rdoQuery
Dim seldissecentr11_RS As rdoResultset
Dim SELDBMASUB_QR As rdoQuery
Dim INSDBMASUB_QR As rdoQuery
Dim UPDDBMASUB_QR As rdoQuery
Dim SELDBMASUB_RS As rdoResultset
Dim NOTHINGWRITTEN As Boolean

Dim UPD_DBMAREPORT_QR As rdoQuery

'***** Shraddha 12/04/10 *****
Dim selcbcsm_qr As rdoQuery
Dim selcbcsm_rs As rdoResultset

Dim SELDREQDRNAME_QR As rdoQuery
Dim SELDREQDRNAME_RS As rdoResultset

Dim UPDDREQDRNAME_QR As rdoQuery
Dim inscbcsm_qr As rdoQuery
Dim updcbcsm_qr As rdoQuery
''******added by shweta 10-02-2014
Dim SELDCBCREPA_QR As rdoQuery
Dim SELDCBCREPA_RS As rdoResultset
Dim INSDCBCREPA_QR As rdoQuery
Dim UPDDCBCREPA_QR As rdoQuery
'''till here



'SAVE THE DATE AND TIME INTO RTIMELOG
Dim UPDRTIME_QR As rdoQuery
Dim UPD_RS As rdoResultset

Dim UPDRTIMECADD_QR As rdoQuery

Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEC_RS As rdoResultset

Dim UPDRTIMEP_QR As rdoQuery
Dim UPDRTIMEM_QR As rdoQuery
'till here



Private Sub Check1_Click()
If Check1.Value = 1 Then
   If flagsave = 1 Or flagsave = 2 Then
   
                 txt_myeloperoxidase.Enabled = True
        txt_nonspecificesterase.Enabled = True
        txt_trapironstain.Enabled = True
        TXT_IRONSTAIN.Enabled = True
        TXT_GIDEROBLASTS.Enabled = True
        txt_sudanblackb.Enabled = True
        txt_TOLUIDINEBLUE.Enabled = True
        Text1.Enabled = True 'False
        Text2.Enabled = True 'False
        Text3.Enabled = True 'False
        txt_cellularity.Enabled = True
        TXT_MERATION.Enabled = True
        TXT_MERATION1.Enabled = True
        TXT_PROMONOCYTES.Enabled = True
        txt_promyelocytes.Enabled = True
        txt_myelocytes.Enabled = True
        txt_metamyelocytes.Enabled = True
        txt_polymorphs.Enabled = True
        txt_basophils.Enabled = True
        txt_eosinophils.Enabled = True
        txt_monocytes.Enabled = True
        TXT_PROMONOCYTES.Enabled = True
        txt_matly.Enabled = True
        txt_atyply.Enabled = True
        txt_plasmacells.Enabled = True
        txt_mature.Enabled = True
        txt_inmature.Enabled = True
        txt_atypbi.Enabled = True
        txt_plasmablasts.Enabled = True
        txt_cellularity1.Enabled = True
        txt_maturation.Enabled = True
        TXT_ErythCells.Enabled = True
        txt_cellularity2.Enabled = True
        txt_morphology.Enabled = True
        CMB_MYELOPEROXIDASE.Enabled = True
        'txt_myeloperoxidase.Enabled = False
        CMB_NONSPEST.Enabled = True
        'txt_nonspecificesterase.Enabled = False
        CMB_TRAP.Enabled = True
        'txt_trapironstain.Enabled = False
        CMB_IRONSTAIN.Enabled = True
        CMB_GIDEROBLASTS.Enabled = True
        CMB_SUDANBLACKB.Enabled = True
        CMB_TOLUIDINEBLUE.Enabled = True
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
        TXT_IMPRINT.Enabled = True
        txt_descmorpho.Enabled = True
        txt_IMPRESSION.Enabled = True
        txt_abcellblasts.Enabled = True
        TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        regcbo.Enabled = True
        
   End If
   
   
End If
If Check1.Value = 0 Then
    If flagsave = 1 Or flagsave = 2 Then
              txt_cellularity.Enabled = False
        TXT_MERATION.Enabled = False
        TXT_MERATION1.Enabled = False
        TXT_PROMONOCYTES.Enabled = False
        txt_promyelocytes.Enabled = False
        txt_myelocytes.Enabled = False
        txt_metamyelocytes.Enabled = False
        txt_polymorphs.Enabled = False
        txt_basophils.Enabled = False
        txt_eosinophils.Enabled = False
        txt_monocytes.Enabled = False
        TXT_PROMONOCYTES.Enabled = False
        txt_matly.Enabled = False
        txt_atyply.Enabled = False
        txt_plasmacells.Enabled = False
        txt_mature.Enabled = False
        txt_inmature.Enabled = False
        txt_atypbi.Enabled = False
        txt_plasmablasts.Enabled = False
        txt_cellularity1.Enabled = False
        txt_maturation.Enabled = False
        TXT_ErythCells.Enabled = False
        txt_cellularity2.Enabled = False
        txt_morphology.Enabled = False
        CMB_MYELOPEROXIDASE.Enabled = False
        'txt_myeloperoxidase.Enabled = False
        CMB_NONSPEST.Enabled = False
        'txt_nonspecificesterase.Enabled = False
        CMB_TRAP.Enabled = False
        'txt_trapironstain.Enabled = False
        CMB_IRONSTAIN.Enabled = False
        CMB_GIDEROBLASTS.Enabled = False
        CMB_SUDANBLACKB.Enabled = False
        CMB_TOLUIDINEBLUE.Enabled = False
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
        TXT_IMPRINT.Enabled = False
        txt_descmorpho.Enabled = False
        txt_IMPRESSION.Enabled = False
        txt_abcellblasts.Enabled = False
        'TXT_COMMENTS.Enabled = False
        
        txt_myeloperoxidase.Enabled = False
        txt_nonspecificesterase.Enabled = False
        txt_trapironstain.Enabled = False
        TXT_IRONSTAIN.Enabled = False
        TXT_GIDEROBLASTS.Enabled = False
        txt_sudanblackb.Enabled = False
        txt_TOLUIDINEBLUE.Enabled = False
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        cmb_enterby.Enabled = False
        'cmb_finalby.Enabled = False
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    '  txt_impression.Enabled = True
      'TXT_COMMENTS.Enabled = False
      cmb_enterby.Enabled = True
      
   End If
End If
End Sub

''****ADDENDUM**************


'Private Sub cmb_dept_Click()
'Dim deptno As String
'Dim dept As Integer
'Dim l As Integer
'On Error GoTo MSG:
'    If Trim(cmb_dept) = "" Then
'        Exit Sub
'    End If
'
'    dept = InStr(cmb_dept.Text, "$")
'    txt_deptcode.Text = Mid(cmb_dept.Text, dept + 1, 3)
'    If SEL_MODALITYFLAG1 = False Then
'        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
'        Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
'        SEL_MODALITYFLAG1 = True
'    Else
'        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
'        sel_modalityrs.MoreResults
'        sel_modalityrs.Requery
'    End If
'    If sel_modalityrs.RowCount > 0 Then
'        cmb_modality.clear
'        While Not sel_modalityrs.EOF
'            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
'            sel_modalityrs.MoveNext
'        Wend
'    End If
'Exit Sub
'MSG:
'MsgBox Error(Err)
'End Sub

'Private Sub cmb_dept_GotFocus()
'On Error GoTo MSG:
'Dim ind_cbodept As Integer
'If Trim(cmb_dept) = "" Then
'    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
'        If Trim(cmb_dept.List(ind_cbodept)) = "HAEMATO ONCOLOGY LABORATORY$UZZ" Then
'            cmb_dept.ListIndex = ind_cbodept
'        End If
'    Next ind_cbodept
'    End If
'Exit Sub
'MSG:
'MsgBox Error(Err)
'End Sub









Private Sub cmb_finalby_LostFocus()
'cmd_save.SetFocus
End Sub

Private Sub CMB_GIDEROBLASTS_Click()
'If CMB_GIDEROBLASTS.Text = "Positive" Then
 '   TXT_GIDEROBLASTS.Visible = True
'Else
 '   TXT_GIDEROBLASTS.Visible = False
'End If
End Sub



Private Sub CMB_IRONSTAIN_Click()
'If CMB_IRONSTAIN.Text = "Positive" Then
'    TXT_IRONSTAIN.Visible = True
'Else
'    TXT_IRONSTAIN.Visible = False
'End If
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

Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub


Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "BONE MARROW ASPIRATION REPORT$BM" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub CMB_MYELOPEROXIDASE_Click()
    If CMB_MYELOPEROXIDASE.Text = "Positive" Then
        txt_myeloperoxidase.Visible = True
    Else
        txt_myeloperoxidase.Visible = False
    End If
End Sub

Private Sub CMB_NONSPEST_Click()
If CMB_NONSPEST.Text = "Positive" Then
    txt_nonspecificesterase.Visible = True
Else
    txt_nonspecificesterase.Visible = False
End If
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


Public Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
Dim entr As String
'On Error GoTo MSG:
If Opt_patient.Value = True Then
        If NEWFLAG = True Then
            NEWFLAG = False
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
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
'                'lbl_ward.Visible = True
                'lbl_bed.Visible = True
                'txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_status.Text = "OP"
                'txt_ward.Visible = False
                'txt_bedno.Visible = False
                'lbl_ward.Visible = False
                'lbl_bed.Visible = False
            End If
          
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                TXT_PATH.Text = "" & Trim(selpathdrequisitn_rs("PATHNO"))
                'If selpathdrequisitn_rs.EOF Then
                    txt_nmaterial.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                'End If
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
           
            If Trim(txt_modcode.Text) = "BF" Then
                ENABL ("ENA")
                
                If seldissecprovflag1 = False Then
                    Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
                    seldissecprovflag1 = True
                Else
                    seldissecprov1_rs.MoreResults
                    seldissecprov1_rs.Requery
                End If
                If seldissecprov1_rs.RowCount > 0 Then
                    
                   CMB_PROVISIONAL.clear
                    While Not seldissecprov1_rs.EOF
                        CMB_PROVISIONAL.AddItem Trim(seldissecprov1_rs("CCNAME"))
                        
                        seldissecprov1_rs.MoveNext
                    Wend
                End If
                
                
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
'
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
'            If REFDR_RS.RowCount > 0 Then
'                txt_refdoc.Text = "" & REFDR_RS("REFBYDOC")
'                If REFDR_RS("COMMITFLAG") = "Y" Then
'                    MsgBox "Report is commited , You can not Modify"
'                    cmd_modify.Enabled = False
'                    FLDENABL ("ADDENDUM")
'                    ADDENDUMFALG = True
'                    cmd_save.Enabled = True
'                End If
'                cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
'                'CMB_PROVISIONAL.Text = REFDR_RS("PREPORTBY")
'                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
'                    cmb_finalby.Text = ""
'                Else
'                    cmb_finalby.Text = "" & Trim(REFDR_RS("FREPORTBY"))
'                End If
'
'            End If
            
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
                'lbl_ward.Visible = False
                'lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                'lbl_ward.Visible = False
                'lbl_bed.Visible = False
                txt_status.Text = "OP"
            End If
          
         
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                TXT_PATH.Text = "" & Trim(selpathdrequisitn_rs("PATHNO"))
                'If selpathdrequisitn_rs.EOF = False Then
                    'txt_relclnhis.Visible = True
                    
                    txt_nmaterial.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                'End If
            End If
            
            ''*******ADDENDUM*********************************
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
            ''*******ADDENDUM*********************************
            
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
                txt_procedure.Text = Trim(selta1012_rs("prvdiagdes"))
            End If
            If SELDBMAREPORTFLAG = False Then
                SELQR(0) = req_chglen(cmb_reqno.Text)
                Set SELDBMAREPORT_RS = SELQR.OpenResultset(1, 3)
                SELDBMAREPORTFLAG = True
            Else
                  SELQR(0) = req_chglen(cmb_reqno.Text)
                  SELDBMAREPORT_RS.MoreResults
                  SELDBMAREPORT_RS.Requery
            End If
            ENABL ("DIS")
            If SELDBMAREPORT_RS.RowCount > 0 Then
                    txt_cellularity.Text = " " & SELDBMAREPORT_RS("CELLULARIT")
                    If IsNull(SELDBMAREPORT_RS("MERATIO")) Then
                    Else
                        I = InStr(SELDBMAREPORT_RS("MERATIO"), ":")
                        If I <> 0 Then
                            TXT_MERATION.Text = " " & Mid(SELDBMAREPORT_RS("MERATIO"), 1, I - 1)
                            TXT_MERATION1.Text = " " & Mid(SELDBMAREPORT_RS("MERATIO"), I + 1, 2)
                        Else
                            TXT_MERATION.Text = Trim(SELDBMAREPORT_RS("MERATIO"))
                        End If
                    End If
                    TXT_COMMENTS.Text = "" & Trim(SELDBMAREPORT_RS("COMMENTS"))
                    txt_promyelocytes.Text = " " & Trim(SELDBMAREPORT_RS("PROMYELOCY"))
                    txt_myelocytes.Text = " " & Trim(SELDBMAREPORT_RS("MYELOCYTE"))
                    txt_metamyelocytes.Text = " " & Trim(SELDBMAREPORT_RS("MMYELOCYTE"))
                    txt_polymorphs.Text = " " & Trim(SELDBMAREPORT_RS("POLYMORPHS"))
                    txt_basophils.Text = " " & Trim(SELDBMAREPORT_RS("BASOPHILS"))
                    txt_eosinophils.Text = " " & Trim(SELDBMAREPORT_RS("EOSINOPHIL"))
                    txt_monocytes.Text = " " & Trim(SELDBMAREPORT_RS("MONOCYTES"))
                    TXT_PROMONOCYTES.Text = " " & Trim(SELDBMAREPORT_RS("PROMONOCYT"))
                    txt_matly.Text = " " & Trim(SELDBMAREPORT_RS("MLYMPHOCYT"))
                    txt_atyply.Text = " " & Trim(SELDBMAREPORT_RS("ALYMPHOCYT"))
                    txt_plasmacells.Text = " " & Trim(SELDBMAREPORT_RS("PLASMACELL"))
                    txt_mature.Text = " " & Trim(SELDBMAREPORT_RS("MATURE"))
                    txt_inmature.Text = " " & Trim(SELDBMAREPORT_RS("IMMATURE"))
                    txt_atypbi.Text = " " & Trim(SELDBMAREPORT_RS("ATYPBINU"))
                    txt_plasmablasts.Text = " " & Trim(SELDBMAREPORT_RS("PLASMABLAS"))
                    txt_cellularity1.Text = " " & Trim(SELDBMAREPORT_RS("ECELLULARI"))
                    txt_maturation.Text = " " & Trim(SELDBMAREPORT_RS("MATURATION"))
                    txt_cellularity2.Text = " " & Trim(SELDBMAREPORT_RS("MCELLULARI"))
                    txt_morphology.Text = " " & Trim(SELDBMAREPORT_RS("MORPHOLOGY"))
                    TXT_IMPRINT.Text = " " & Trim(SELDBMAREPORT_RS("IMPRINT"))
                    txt_descmorpho.Text = "" & Trim(SELDBMAREPORT_RS("DESCMORPHO"))
                    txt_abcellblasts.Text = "" & Trim(SELDBMAREPORT_RS("ABCELLBLAS"))
                    '***** Uncommented by Shradha on 09/03/10 *****
                    txt_tcrab.Text = "" & Trim(SELDBMAREPORT_RS("TCRAB"))
                    txt_tcrag.Text = "" & Trim(SELDBMAREPORT_RS("TCRAG"))
                    txt_cyclinD1.Text = "" & Trim(SELDBMAREPORT_RS("CLDN1"))
                    
                    '*********** Shraddha 12/04/10 ***********
                    selcbcsm_qr(0) = req_chglen(cmb_reqno.Text)
                    Set selcbcsm_rs = selcbcsm_qr.OpenResultset(1, 1)
               
                    If Not selcbcsm_rs.EOF Then
                        If Trim(selcbcsm_rs("optlsp")) = "Y" Then
                            opt_lsp.Value = True
                        ElseIf Trim(selcbcsm_rs("optlsp")) = "N" Then
                            opt_lsp.Value = False
                        End If
                    
                        If Trim(selcbcsm_rs("optscp")) = "Y" Then
                            opt_ssp.Value = True
                        ElseIf Trim(selcbcsm_rs("optscp")) = "N" Then
                            opt_ssp.Value = False
                        End If
                    End If
                    
                    
                    'ADDED BY SHWETA ON 19-07-2013 AS PER REQUEST UZZ DEPT.
                    Dim RS_ERYTHCELL As rdoResultset
                    Dim ERYTHSTR As String
                    ERYTHSTR = "SELECT * FROM DISLIB.DBMASUBNEW WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "'"
                    Set RS_ERYTHCELL = rdb.OpenResultset(ERYTHSTR, 1, 1)
                    If RS_ERYTHCELL.EOF = False Then
                       TXT_ErythCells.Text = RS_ERYTHCELL(1)
                    End If
                    
                    SELDREQDRNAME_QR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDREQDRNAME_RS = SELDREQDRNAME_QR.OpenResultset(1, 3)
                    If SELDREQDRNAME_RS.RowCount > 0 Then
                        If Trim(SELDREQDRNAME_RS("LATERALITS")) = "Dr. Ashok Kumar" Then
                            opt_drak.Value = True
                            opt_mryb.Value = False
                            opt_sitaram.Value = False
                        ElseIf Trim(SELDREQDRNAME_RS("METASTASIS")) = "Mr.Y. Badrinath" Then
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
                    '**********************************************
                    If Trim(SELDBMAREPORT_RS("myeloperox")) = "Positive" Then
                        CMB_MYELOPEROXIDASE.Text = " " & Trim(SELDBMAREPORT_RS("myeloperox"))
                        txt_myeloperoxidase.Visible = True
                        txt_myeloperoxidase.Text = " " & Trim(SELDBMAREPORT_RS("MYELOPERO"))
                    Else
                        CMB_MYELOPEROXIDASE.Text = " " & Trim(SELDBMAREPORT_RS("myeloperox"))
                        txt_myeloperoxidase.Visible = False
                    End If
                    If Trim(SELDBMAREPORT_RS("NONSPESTER")) = "Positive" Then
                        CMB_NONSPEST.Text = " " & Trim(SELDBMAREPORT_RS("NONSPESTER"))
                        txt_nonspecificesterase.Visible = True
                        txt_nonspecificesterase.Text = " " & Trim(SELDBMAREPORT_RS("NONSPESTE1"))
                    Else
                        CMB_NONSPEST.Text = " " & Trim(SELDBMAREPORT_RS("NONSPESTER"))
                        txt_nonspecificesterase.Visible = False
                    End If
                    'If Trim(SELDBMAREPORT_RS("TRAP")) = "Positive" Then
                        'CMB_TRAP.Text = " " & SELDBMAREPORT_RS("TRAP")
                        'txt_trapironstain.Visible = True
                        txt_trapironstain.Text = " " & Trim(SELDBMAREPORT_RS("TRAP1"))
                    'Else
                        'CMB_TRAP.Text = " " & SELDBMAREPORT_RS("TRAP")
                        'txt_trapironstain.Visible = False
                    'End If
                    'txt_trapironstain.Enabled = False
                    'If Trim(SELDBMAREPORT_RS("IRONSTAIN")) = "Positive" Then
                     '   CMB_IRONSTAIN.Text = " " & SELDBMAREPORT_RS("IRONSTAIN")
'                        TXT_IRONSTAIN.Visible = True
                        TXT_IRONSTAIN.Text = " " & Trim(SELDBMAREPORT_RS("IRONSTAIN1"))
                    'Else
                     '   CMB_IRONSTAIN.Text = " " & SELDBMAREPORT_RS("IRONSTAIN")
'                        TXT_IRONSTAIN.Visible = False
                    'End If
                    'If Trim(SELDBMAREPORT_RS("GIDEROBLAS")) = "Positive" Then
                     '   CMB_GIDEROBLASTS.Text = " " & SELDBMAREPORT_RS("GIDEROBLAS")
                      '  TXT_GIDEROBLASTS.Visible = True
                        TXT_GIDEROBLASTS.Text = " " & Trim(SELDBMAREPORT_RS("fld1"))
                    'Else
                     '   CMB_GIDEROBLASTS.Text = " " & SELDBMAREPORT_RS("GIDEROBLAS")
                      '  TXT_GIDEROBLASTS.Visible = False
                        
                    'End If
                    If Trim(SELDBMAREPORT_RS("sudanblack")) = "Positive" Then
                        CMB_SUDANBLACKB.Text = " " & Trim(SELDBMAREPORT_RS("sudanblack"))
                        txt_sudanblackb.Visible = True
                        txt_sudanblackb.Text = " " & Trim(SELDBMAREPORT_RS("sudanblac1"))
                    Else
                    CMB_SUDANBLACKB.Text = " " & Trim(SELDBMAREPORT_RS("sudanblack"))
                    txt_sudanblackb.Visible = False
                   
                    End If
                    If Trim(SELDBMAREPORT_RS("toluidineb")) = "Positive" Then
                        CMB_TOLUIDINEBLUE.Text = " " & Trim(SELDBMAREPORT_RS("toluidineb"))
                        txt_TOLUIDINEBLUE.Visible = True
                        txt_TOLUIDINEBLUE.Text = " " & Trim(SELDBMAREPORT_RS("toluidine1"))
                    Else
                        CMB_TOLUIDINEBLUE.Text = " " & Trim(SELDBMAREPORT_RS("toluidineb"))
                        txt_TOLUIDINEBLUE.Visible = False
                        
                    End If
                    txt_cd13.Text = " " & Trim(SELDBMAREPORT_RS("cd13"))
                    txt_cd14.Text = " " & Trim(SELDBMAREPORT_RS("CD14"))
                    txt_CD15.Text = " " & Trim(SELDBMAREPORT_RS("CD15"))
                    txt_CD33.Text = " " & Trim(SELDBMAREPORT_RS("CD33"))
                    txt_cd36.Text = " " & Trim(SELDBMAREPORT_RS("CD36"))
                    txt_cd41.Text = " " & Trim(SELDBMAREPORT_RS("CD4142"))
                    TXT_CD42.Text = "" & Trim(SELDBMAREPORT_RS("CD42"))
                    TXT_CD117.Text = "" & Trim(SELDBMAREPORT_RS("CD117"))
                    TXT_CD20.Text = "" & Trim(SELDBMAREPORT_RS("CD20"))
                    TXT_CD61.Text = "" & Trim(SELDBMAREPORT_RS("CD61"))
                    TXT_CD56.Text = "" & Trim(SELDBMAREPORT_RS("CD56"))
                    txt_Glycophorina.Text = " " & Trim(SELDBMAREPORT_RS("GLYCOPHOIN"))
                    TXT_CD68.Text = " " & Trim(SELDBMAREPORT_RS("CD68"))
                    TXT_CD34.Text = " " & Trim(SELDBMAREPORT_RS("CD34"))
                    TXT_HLADR.Text = " " & Trim(SELDBMAREPORT_RS("HLADR"))
                    TXT_CD38.Text = " " & Trim(SELDBMAREPORT_RS("CD38"))
                    TXT_CD138.Text = " " & Trim(SELDBMAREPORT_RS("CD138"))
                    TXT_TDT.Text = " " & Trim(SELDBMAREPORT_RS("TDT"))
                    txt_igm.Text = " " & Trim(SELDBMAREPORT_RS("fld2"))
                    TXT_ANTIMPO.Text = " " & Trim(SELDBMAREPORT_RS("ANTIMPO"))
                    TXT_CD10.Text = " " & Trim(SELDBMAREPORT_RS("CD10"))
                    TXT_CD19.Text = " " & Trim(SELDBMAREPORT_RS("CD19"))
                    TXT_FMC7.Text = " " & Trim(SELDBMAREPORT_RS("FMC7"))
                    TXT_CD22.Text = " " & Trim(SELDBMAREPORT_RS("CD22"))
                    TXT_CD23.Text = " " & Trim(SELDBMAREPORT_RS("CD23"))
                    txt_smig.Text = " " & Trim(SELDBMAREPORT_RS("SMIG"))
                    
                    '***** Shraddha 09/03/10 *****
                    If Trim(SELDBMAREPORT_RS("SMPCTG")) <> "" Then
                        TXT_SMPCTG.Text = "" & Trim(SELDBMAREPORT_RS("SMPCTG"))
                    End If
                    '*****************************
                    
                    txt_CD79B.Text = " " & Trim(SELDBMAREPORT_RS("CD79B"))
                    TXT_CD25.Text = " " & Trim(SELDBMAREPORT_RS("CD25"))
                    TXT_CD11C.Text = " " & Trim(SELDBMAREPORT_RS("CD11C"))
                    txt_cd103.Text = " " & Trim(SELDBMAREPORT_RS("CD103"))
                    txt_kappa.Text = " " & Trim(SELDBMAREPORT_RS("KAPPA"))
                    txt_lanbda.Text = " " & Trim(SELDBMAREPORT_RS("LAMDA"))
                    TXT_DBA44.Text = " " & Trim(SELDBMAREPORT_RS("DBA44"))
                    txt_cytocd221(1).Text = " " & Trim(SELDBMAREPORT_RS("CYTOEP22"))
                    txt_cytocd79a.Text = " " & Trim(SELDBMAREPORT_RS("CYTOCD79A"))
                    TXT_CYTOMU.Text = " " & Trim(SELDBMAREPORT_RS("CYTOMU"))
                    TXT_CD2.Text = " " & Trim(SELDBMAREPORT_RS("CD2"))
                    TXT_CD3.Text = " " & Trim(SELDBMAREPORT_RS("CD3"))
                    TXT_CD4.Text = " " & Trim(SELDBMAREPORT_RS("CD4"))
                    TXT_CD5.Text = " " & Trim(SELDBMAREPORT_RS("CD5"))
                    TXT_CD7.Text = " " & Trim(SELDBMAREPORT_RS("CD7"))
                    TXT_CD8.Text = " " & Trim(SELDBMAREPORT_RS("CD8"))
                    TXT_CYTOCD3.Text = " " & Trim(SELDBMAREPORT_RS("CYTOCD3"))
                    txt_IMPRESSION.Text = "" & Trim(SELDBMAREPORT_RS("impression"))
            End If
            
       ' FETCHING FROM DBMASUB
          SELDBMASUB_QR(0) = Trim(req_chglen(cmb_reqno.Text))
          
          
          Set SELDBMASUB_RS = SELDBMASUB_QR.OpenResultset(1, 1)
            If SELDBMASUB_RS.RowCount > 0 Then
                   Text2.Text = " " & Trim(SELDBMASUB_RS("PROLYMPHO"))
                   Text3.Text = " " & Trim(SELDBMASUB_RS("HEMATOGONE"))
                   Text1.Text = " " & Trim(SELDBMASUB_RS("HAIRYCELL"))
                   TXT_COMMENTS.Text = "" & Trim(SELDBMASUB_RS("COMMENTS"))
                   If Trim(SELDBMASUB_RS("fullrep")) = "Y" Then
                      Check1.Value = 1
                   ElseIf Trim(SELDBMASUB_RS("fullrep")) = "N" Then
                      Check1.Value = 0
                   End If
                   Check1.Enabled = False
                   
                      
            End If
            
           ' FETCHING FROM DBMASUB ENDS
                   
       
       
       
       ' FETCHING FROM DBMASUB
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
                txt_refdoc.Text = "" & REFDR_RS("REFBYDOC")
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Report is commited , You can not Modify"
                    cmd_modify.Enabled = False
                    FLDENABL ("ADDENDUM")
                    ADDENDUMFALG = True
                    cmd_save.Enabled = True
                End If
                cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
                'CMB_PROVISIONAL.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = "" & Trim(REFDR_RS("FREPORTBY"))
                End If
                If IsNull(Trim(REFDR_RS("PREPORTBY"))) Then
                    regcbo.Text = ""
                Else
                    regcbo.Text = "" & Trim(REFDR_RS("PREPORTBY"))
                End If
            End If
End If
End If
If opt_staff.Value = True Then
    STAFF
End If
VIEWDCBCREPA
If SSTab1.Caption = "MORPHOLOGY" Then
    'txt_cellularity.SetFocus
ElseIf SSTab1.Caption = "CYTOCHEMISTRY" Then
    'CMB_MYELOPEROXIDASE.SetFocus
ElseIf SSTab1.Caption = "SURFACE MARKERS" Then
    'txt_cd13.SetFocus
End If
Exit Sub
MSG:
MsgBox Err.Description
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
    'cmd_fullreport.Enabled = False
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
    'cmd_fullreport.Enabled = True
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
    'cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
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
    'cmd_fullreport.Enabled = False
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
    'cmd_fullreport.Enabled = False
ElseIf cap = "print" Then
     cmd_save.Enabled = False
     cmd_back.Enabled = True
     cmd_modify.Enabled = False
     cmd_print.Enabled = False
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False
End If
'If cap = "NEW" Then
'    cmd_save.Enabled = True
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'
'ElseIf cap = "VIEW" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'
'ElseIf cap = "SAVE" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = True
'    cmd_view.Enabled = True
'
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'
'ElseIf cap = "COMMIT" Then
'    cmd_save.Enabled = False
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'
'ElseIf cap = "print" Then
'     cmd_save.Enabled = False
'     cmd_back.Enabled = True
'     cmd_modify.Enabled = False
'     cmd_print.Enabled = False
'     cmd_cancel.Enabled = True
'     CMD_NEW.Enabled = False
'
'End If
End Sub



Private Sub CMB_SUDANBLACKB_Click()
If CMB_SUDANBLACKB.Text = "Positive" Then
    txt_sudanblackb.Visible = True
Else
    txt_sudanblackb.Visible = False
    
End If
End Sub




Private Sub CMB_TOLUIDINEBLUE_Click()
If CMB_TOLUIDINEBLUE.Text = "Positive" Then
    txt_TOLUIDINEBLUE.Visible = True
Else
    txt_TOLUIDINEBLUE.Visible = False
End If
End Sub

Private Sub CMB_TOLUIDINEBLUE_LostFocus()
txt_IMPRESSION.SetFocus
End Sub

Private Sub CMB_TRAP_Click()
'If CMB_TRAP.Text = "Positive" Then
 '   txt_trapironstain.Visible = True
'Else
 '   txt_trapironstain.Visible = False
'End If
End Sub

Private Sub cmd_commit_Click()
Dim I As Integer
If Trim(txt_modcode.Text) = "PF" Then
    If Len(Trim(TXT_PFTIMPRESSION.Text)) = 0 Then
        MsgBox "Impression Field is Empty, You can not Commit the Report.", vbCritical
        TXT_PFTIMPRESSION.SetFocus
        Exit Sub
    End If
End If
I = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If I = vbYes Then
On Error GoTo MSG:

''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
        If LOGINTYP >= CURUSRDET_RS("usrtype") Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            If LOGINTYP = 8 Or LOGINTYP = 9 Then
            
            Else
            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
            cmd_commit.SetFocus
            Exit Sub
            End If
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
            If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
                'If LOGINUSRID <> curusrdetrs("usr_id") Then
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
    updrrequisitn1_qr(0) = UCase(cmb_finalby.Text)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    
     'rtimelog
    UPDRTIMEC_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute
   'rtimelog
    
    
    
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
blank = True

flagsave = 2
btnenabl ("NEW")
    FLDENABL ("MODIFY")
If Trim(SELDBMASUB_RS("FULLREP")) = "Y" Then

    ENABL ("FENA")
End If
If Trim(SELDBMASUB_RS("FULLREP")) = "N" Then

    ENABL ("IENA")
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
cmb_enterby.Enabled = True
cmb_finalby.Enabled = True
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
'On Error GoTo MSG:
btnenabl ("NEW")
blank = False
field_clear
'TXT_CLEAR
ENABL ("ENA")
TCLEAR
If DATETIMEFLAG = False Then
   DATETIME
   DATETIMEFLAG = True
Else
    DATETIME
End If
txt_date.Text = DATE11
txt_time.Text = time1
cmb_enterby.Text = cmb_enterby.List(0)
cmb_enterby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_finalby.List(L1)) = "C N NAIR" Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
Next L1
'cmb_finalby.Locked = True
'cmb_finalby.Text = cmb_finalby.List(0)
If SELDEPTRSFLAG1 = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SELDEPTRSFLAG1 = True
Else
    sel_deptrs.MoreResults
    sel_deptrs.Requery
End If

'If sel_deptrs.RowCount > 0 Then
'    sel_deptrs.MoveFirst
'    cmb_dept.clear
'    While Not sel_deptrs.EOF
'        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
'        sel_deptrs.MoveNext
'    Wend
'End If

flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
'cmb_dept.SetFocus
MaskEdBox1.Enabled = True
Opt_patient.Value = True
Check1.Value = 1
Check1.Enabled = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_cancel_Click()
btnenabl ("INI")
field_clear
viewflag = False
NEWFLAG = False
TCLEAR

End Sub

Public Sub field_clear()
TXT_DNAINDEX.Text = ""
TXT_DNACLOIDY.Text = ""
Dim obj As Object
txt_abcellblasts.Text = ""
TXT_COMMENTS.Text = ""
txt_promyelocytes.Text = ""
TXT_IMPRINT.Text = ""
txt_descmorpho.Text = ""
regcbo.Text = ""
'cmb_dept.ListIndex = -1
'cmb_modality.ListIndex = -1
'txt_deptcode.Text = " "
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
' MaskEdBox1.Mask = "#######"  ' COMMENTED BY VARINDER
End If
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
txt_IMPRESSION.Text = ""
cmb_enterby.Text = ""
cmb_finalby.Text = ""
cmb_addendumby.Text = ""
txt_addendum.Text = ""
txt_myeloperoxidase.Text = ""
txt_nonspecificesterase.Text = ""
txt_trapironstain.Text = ""
TXT_IRONSTAIN.Text = ""
TXT_GIDEROBLASTS.Text = ""
txt_sudanblackb.Text = ""
txt_TOLUIDINEBLUE.Text = ""
CMB_SUDANBLACKB.Text = ""
CMB_TOLUIDINEBLUE.Text = ""
CMB_MYELOPEROXIDASE.Text = ""
CMB_NONSPEST.Text = ""
CMB_TRAP.Text = ""
CMB_IRONSTAIN.Text = ""
CMB_GIDEROBLASTS.Text = ""

End Sub




Public Sub FLDENABL(EN As String)
If EN = "DIS" Then
        'cmb_dept.Enabled = False
        'cmb_modality.Enabled = False
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
        'cmb_finalby.Enabled = False
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = False
End If
If EN = "ENA" Then
        'cmb_dept.Enabled = True
        'cmb_modality.Enabled = True
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
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = False
End If
If EN = "MODIFY" Then
        'cmb_dept.Enabled = True
        'cmb_modality.Enabled = True
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
        'cmb_addendumby.Enabled = True
        'txt_addendum.Enabled = True
End If
If EN = "ADDENDUM" Then
    'cmb_dept.Enabled = False
    'cmb_modality.Enabled = False
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
    cmb_addendumby.Enabled = False
    txt_addendum.Enabled = False
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
        cmb_addendumby.Enabled = True
    Else
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
End If
End Sub

Private Sub cmd_print_Click()
'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim fnby1 As String
Dim I As Integer
 'On Error GoTo prtmsg:

DoEvents

''' added for print emr report for sangrur by rahul 21/02/2018
Dim REQNO As String
REQNO = req_chglen(cmb_reqno.Text)
 strLink = "http://10.80.1.9/EMRREPORTS/allreports.aspx?reqno=" & REQNO & "  ~GUEST"
 
 
    'strLink = "http://10.80.1.9/EMRTMHACT/DEFAULT.ASPX?QPNO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink, vbMaximizedFocus
Exit Sub
'' till here by rahul 21/02/2018





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
                SELRREQUISITN_QR(1) = "BM" 'Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = "UZZ" 'Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = "BM" 'Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = "UZZ" 'Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
          SELDBMASUB_QR(0) = VAR_REQNNO
          Set SELDBMASUB_RS = SELDBMASUB_QR.OpenResultset(1, 1)
          If SELDBMASUB_RS.RowCount = 0 Then
            MsgBox "There is no indicator to show whether the Report is Full Report or otherwise"
            Exit Sub
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
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'ABOVE statement COMMENTED AND NEXT ADDED BY Shraddha ON 02/09/09
CrystalReport1.SelectionFormula = "{REPBMLF_1.REQNNO} = '" & VAR_REQNNO & "'"

'***** Commented by Shraddha on 09/03/10
'If Trim(txt_modcode.Text) = "BM" Then

    If Opt_patient.Value = True Then
      If SELDBMASUB_RS("fullrep") = "Y" Then
       'If ADDENDUMFALG = False Then
        'CrystalReport1.ReportFileName = "c:\DISREPORTSEC\DISBM3A.RPT"
        CrystalReport1.ReportFileName = "c:\DISREPORTSECLF\DISBM3A.RPT"
         'CrystalReport1.Formulas(27) = ""
        'Added by Milind on 01/01/06 for displaying designation of Dr. according to name
        'If Trim(cmb_finalby.Text) = "C N NAIR" Then
         '   CrystalReport1.Formulas(27) = "desg = 'Officer-Incharge'"
        'e 'nd If
        'If Trim(cmb_finalby.Text) = "MONA ANAND" Then
        '   CrystalReport1.Formulas(27) = "desg = 'Pathologist'"
        'End If
        'If Trim(cmb_finalby.Text) = "SUMEET GUJRAL" Then
        '   CrystalReport1.Formulas(27) = "desg = ' Acting Officer-Incharge'"
        'End If
        'upto here
       
       If ADDENDUMFALG = True Then
        CrystalReport2.ReportFileName = "c:\DISREPORTSEC\DISBM3AADD3.RPT"
        CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            
            'Added by Milind on 01/01/06 for displaying designation of Dr. according to name
'            If Trim(cmb_finalby.Text) = "C N NAIR" Then
'                CrystalReport2.Formulas(7) = "desg = 'Officer-Incharge'"
'            End If
'            If Trim(cmb_finalby.Text) = "MONA ANAND" Then
'                CrystalReport2.Formulas(27) = "desg = 'Pathologist'"
'            End If
'            If Trim(cmb_finalby.Text) = "SUMEET GUJRAL" Then
'                CrystalReport2.Formulas(27) = "desg = 'Acting Officer-Incharge'"
'            End If
            'upto here
            If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                CrystalReport2.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(cmb_enterby.Text)
                If Len(Trim(cmb_enterby.Text)) > 0 Then
                    CrystalReport2.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    CrystalReport2.Formulas(9) = "ENT = '" & "" & "'"
                End If
            Else
                CrystalReport2.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
                CrystalReport2.Formulas(9) = "ENT = '" & str1 & "'"
            End If
                CrystalReport2.Formulas(12) = "USRID = '" & str & "'"
                CrystalReport2.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                CrystalReport2.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
                CrystalReport2.WindowState = crptMaximized
                CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport2.Action = 1
       End If
      ElseIf SELDBMASUB_RS("fullrep") = "N" Then
                CrystalReport1.ReportFileName = "c:\DISREPORTSEC\DISBM3AI.RPT"
                'Added by Milind on 01/01/06 for displaying designation of Dr. according to name
'                If Trim(cmb_finalby.Text) = "C N NAIR" Then
'                    CrystalReport1.Formulas(7) = "desg = 'Officer-Incharge'"
'                End If
'                If Trim(cmb_finalby.Text) = "MONA ANAND" Then
'                    CrystalReport1.Formulas(27) = "desg = 'Pathologist'"
'                End If
'                If Trim(cmb_finalby.Text) = "SUMEET GUJRAL" Then
'                    CrystalReport1.Formulas(27) = "desg = 'Acting Officer-Incharge'"
'                End If
                'upto here
                
            End If
      
    Else
      If SELDBMASUB_RS("fullrep") = "Y" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBMstaff.RPT"
       ElseIf SELDBMASUB_RS("fullrep") = "N" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBMstaff.RPT"
      End If
      
    End If
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        'CrystalReport2.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'       If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        'CrystalReport2.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            'CrystalReport2.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            If Len(Trim(cmb_enterby.Text)) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                'CrystalReport2.Formulas(9) = "ENT = '" & STR1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
                ' CrystalReport2.Formulas(9) = "ENT = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            'CrystalReport2.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            'CrystalReport2.Formulas(9) = "ENT = '" & STR1 & "'"
            
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        'CrystalReport2.Formulas(12) = "USRID = '" & STR & "'"
        If Trim(TXT_CD2.Text) = "" And Trim(TXT_CD3.Text) = "" And Trim(TXT_CD4.Text) = "" And Trim(TXT_CD5.Text) = "" And Trim(TXT_CD7.Text) = "" And Trim(TXT_CD8.Text) = "" And Trim(TXT_CYTOCD3.Text) = "" Then
            CrystalReport1.Formulas(13) = "tmarkers = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "tmarkers = '" & "abc" & "'"
        End If
        If Trim(TXT_CD10.Text) = "" And Trim(TXT_CD19.Text) = "" And Trim(TXT_FMC7.Text) = "" And Trim(TXT_CD22.Text) = "" And Trim(TXT_CD23.Text) = "" And Trim(txt_smig.Text) = "" And Trim(txt_CD79B.Text) = "" And Trim(TXT_CD25.Text) = "" And Trim(TXT_CD11C.Text) = "" And Trim(txt_cd103.Text) = "" And Trim(txt_kappa.Text) = "" And Trim(txt_lanbda.Text) = "" And Trim(TXT_DBA44.Text) = "" And Trim(txt_cytocd22(0).Text) = "" And Trim(txt_cytocd79a.Text) = "" And Trim(TXT_CYTOMU.Text) = "" Then
            CrystalReport1.Formulas(14) = "bmarkers = '" & "" & "'"
        Else
            CrystalReport1.Formulas(14) = "bmarkers = '" & "abc" & "'"
        End If
        If Trim(TXT_CD34.Text) = "" And Trim(TXT_HLADR.Text) = "" And Trim(TXT_CD38.Text) = "" And Trim(TXT_TDT.Text) = "" And Trim(TXT_CD138.Text) = "" And Trim(TXT_ANTIMPO.Text) = "" Then
            CrystalReport1.Formulas(15) = "others = '" & "" & "'"
        Else
            CrystalReport1.Formulas(15) = "others = '" & "abc" & "'"
        End If
        If Trim(txt_cd13.Text) = "" And Trim(txt_cd14.Text) = "" And Trim(txt_CD15.Text) = "" And Trim(txt_CD33.Text) = "" And Trim(txt_cd36.Text) = "" And Trim(txt_cd41.Text) = "" And Trim(txt_Glycophorina.Text) = "" And Trim(TXT_CD68.Text) = "" And Trim(TXT_CD42.Text) = "" And Trim(TXT_CD117.Text) = "" And Trim(TXT_CD2.Text) = "" And Trim(TXT_CD3.Text) = "" And Trim(TXT_CD4.Text) = "" And Trim(TXT_CD5.Text) = "" And Trim(TXT_CD7.Text) = "" And Trim(TXT_CD8.Text) = "" And Trim(TXT_CYTOCD3.Text) = "" And Trim(TXT_CD56.Text) = "" And Trim(TXT_CD10.Text) = "" And Trim(TXT_CD19.Text) = "" And Trim(TXT_CD22.Text) = "" And Trim(TXT_CD23.Text) = "" And Trim(TXT_CD25.Text) = "" And Trim(txt_cd103.Text) = "" And Trim(TXT_CD61.Text) = "" And Trim(txt_lanbda.Text) = "" And Trim(txt_kappa.Text) = "" And Trim(TXT_DBA44.Text) = "" And Trim(TXT_CD11C.Text) = "" And Trim(TXT_FMC7.Text) = "" And Trim(txt_smig.Text) = "" And Trim(txt_CD79B.Text) = "" And Trim(txt_cytocd79a.Text) = "" And Trim(TXT_CYTOMU.Text) = "" Then
            CrystalReport1.Formulas(16) = "mm = '" & "" & "'"
        Else
            CrystalReport1.Formulas(16) = "mm = '" & "abc" & "'"
        End If
        If Trim(TXT_IRONSTAIN.Text) = "" And Trim(TXT_GIDEROBLASTS.Text) = "" And Trim(txt_sudanblackb.Text) = "" And Trim(txt_TOLUIDINEBLUE.Text) = "" Then
            CrystalReport1.Formulas(17) = "cci = '" & "" & "'"
        Else
            CrystalReport1.Formulas(17) = "cci = '" & "abc" & "'"
        End If
        If Trim(txt_myeloperoxidase.Text) = "" And Trim(CMB_MYELOPEROXIDASE.Text) = "" And Trim(txt_nonspecificesterase.Text) = "" And Trim(txt_trapironstain.Text) = "" Then
            CrystalReport1.Formulas(18) = "ccm = '" & "" & "'"
            'Commented and added by Shraddha on 18/02/11 Not needed
            'CrystalReport1.Formulas(18) = "ccm = 'abc'"
        Else
            CrystalReport1.Formulas(18) = "ccm = '" & "abc" & "'"
        End If
        If Trim(TXT_IMPRINT.Text) = "" And Trim(txt_descmorpho.Text) = "" Then
            CrystalReport1.Formulas(19) = "im = '" & "" & "'"
        Else
            CrystalReport1.Formulas(19) = "im = '" & "abc" & "'"
        End If
        If Trim(txt_cellularity2.Text) = "" And Trim(txt_morphology.Text) = "" Then
            CrystalReport1.Formulas(20) = "ms = '" & "" & "'"
        Else
            CrystalReport1.Formulas(20) = "ms = '" & "abc" & "'"
        End If
        If Trim(txt_cellularity1.Text) = "" And Trim(txt_maturation.Text) = "" Then
            CrystalReport1.Formulas(21) = "es = '" & "" & "'"
        Else
            CrystalReport1.Formulas(21) = "es = '" & "abc" & "'"
        End If
        If Trim(txt_mature.Text) = "" And Trim(txt_inmature.Text) = "" And Trim(txt_atypbi.Text) = "" And Trim(txt_plasmablasts.Text) = "" Then
            CrystalReport1.Formulas(22) = "pc = '" & "" & "'"
        Else
            CrystalReport1.Formulas(22) = "pc = '" & "abc" & "'"
        End If
        If Trim(txt_matly.Text) = "" And Trim(txt_atyply.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(Text1.Text) = "" And Trim(Text2.Text) = "" And Trim(Text3.Text) = "" Then
            CrystalReport1.Formulas(23) = "ls = '" & "" & "'"
        Else
            CrystalReport1.Formulas(23) = "ls = '" & "abc" & "'"
        End If
        If Trim(txt_promyelocytes.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_polymorphs.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(TXT_PROMONOCYTES.Text) = "" Then
            CrystalReport1.Formulas(24) = "mys = '" & "" & "'"
        Else
            CrystalReport1.Formulas(24) = "mys = '" & "abc" & "'"
        End If
        If Trim(txt_cellularity.Text) = "" And Trim(TXT_MERATION.Text) = "" Then
            CrystalReport1.Formulas(25) = "mor = '" & "" & "'"
        Else
            CrystalReport1.Formulas(25) = "mor = '" & "abc" & "'"
        End If
        If Trim(TXT_COMMENTS.Text) <> "" Then
            CrystalReport1.Formulas(26) = "nocomments = '" & "" & "'"
            CrystalReport1.Formulas(28) = "comments = '" & "COMMENTS" & "'"
        Else
            CrystalReport1.Formulas(26) = "nocomments = 'YES'"
            CrystalReport1.Formulas(28) = "comments = '" & "" & "'"
        End If
        
    'End If
    If opt_staff.Value = True Then
        CrystalReport1.Formulas(26) = "nam = '" & Trim(txt_name.Text) & "'"
        CrystalReport1.Formulas(27) = "stdob = '" & Trim(txt_name.Text) & "'"
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    End If
     If Mid(Trim(cmb_finalby.Text), 1, 2) = "Dr" Or Mid(Trim(cmb_finalby.Text), 1, 2) = "DR" Then
        fnby1 = Trim(cmb_finalby.Text)
    Else
        fnby1 = "Dr." + Trim(cmb_finalby.Text)
    End If
    'CRYSTALREPORT1.Formulas(28) = "PRFN = '" & Trim(fnby1) & "'"
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1
'If ADDENDUMFALG = True Then
'  CrystalReport2.WindowState = crptMaximized
'  CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'  CrystalReport2.Action = 1
'End If

btnenabl ("INI")

Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("PRINT")

Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_save_Click()
Dim pcount As Integer
    pcount = 0
Dim beforst As String, afterst As String
Dim K As Long
       While InStr(txt_IMPRESSION, Chr(9)) > 0
            
            K = InStr(txt_IMPRESSION, Chr(9))
            beforst = Mid(txt_IMPRESSION, 1, K - 1) & " "
            afterst = Mid(txt_IMPRESSION, K + 1)
            txt_IMPRESSION = beforst & afterst
        Wend
        While InStr(txt_IMPRESSION, Chr(146)) > 0

            K = InStr(txt_IMPRESSION, Chr(146))
            beforst = Mid(txt_IMPRESSION, 1, K - 1) & "`"
            afterst = Mid(txt_IMPRESSION, K + 1)
            txt_IMPRESSION = beforst & afterst
       Wend
     
       If Len(Trim(txt_IMPRESSION.Text)) > 1000 Then
        MsgBox "Max.limit of Impression is 1000 characters."
        TXT_STIMPRESSION.SetFocus
        Exit Sub
       End If
       
       
       If Len(Trim(txt_IMPRESSION.Text)) = 0 Then
        MsgBox "IMpression field is Empty.", vbCritical
        Exit Sub
       End If
       



If Len(Trim(cmb_enterby.Text)) = 0 Then
    MsgBox "Enter by field is Empty.", vbCritical
    Exit Sub
End If

If Len(Trim(cmb_finalby.Text)) = 0 Then
    MsgBox "To be finalised  field is Empty.", vbCritical
    Exit Sub
End If
If Len(Trim(regcbo.Text)) = 0 Then
    MsgBox "Registrar  field is Empty.", vbCritical
    Exit Sub
End If
If Len(Trim(txt_descmorpho.Text)) > 200 Then
  MsgBox "Max Length for Descriptive Moprhology is 200 characters"
  Exit Sub
End If
If Len(Trim(TXT_IMPRINT.Text)) > 200 Then
  MsgBox "Max Length for IMPRINT is 200 characters"
  Exit Sub
End If
If Len(Trim(TXT_COMMENTS.Text)) > 1499 Then
  MsgBox "Max Length for Comments is 1499 characters"
  Exit Sub
End If
If Len(Trim(txt_morphology.Text)) > 200 Then
  MsgBox "Max Length for Morphology is 200 characters"
  Exit Sub
End If
'If Trim(TXT_MERATION.Text) = "" Or Trim(TXT_MERATION1.Text) = "" Then
'    MsgBox "Pls Enter Data in M/E Ratio !!", vbInformation, "Bone Marrow "
'    TXT_MERATION.SetFocus
'    Exit Sub
'End If
    assign

'rtimelog
UPDRTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                
UPDRTIME_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIME_QR.Execute

' upto here
btnenabl ("SAVE")
'cmd_print.Enabled = True
'cmd_cancel.Enabled = True
ENABL ("DIS")


'cmd_commit.Enabled = True
End Sub

Public Sub cmd_view_Click()
'On Error GoTo MSG:
btnenabl ("VIEW")
FLDENABL ("DIS")
ENABL ("DIS")
field_clear
'TXT_CLEAR
TCLEAR
cmb_finalby.Locked = False
regcbo.Locked = False

If DATETIMEFLAG = False Then
   DATETIME
   DATETIMEFLAG = True
Else
    DATETIME
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
'If sel_deptrs.RowCount > 0 Then
'    sel_deptrs.MoveFirst
'    cmb_dept.clear
'    While Not sel_deptrs.EOF
'        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
'        sel_deptrs.MoveNext
'    Wend
'   End If
 
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
 
   
'cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
Opt_patient.Value = True
'cmb_dept.SetFocus
'cmb_modality.Enabled = True
MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
'AppActivate VAR
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = MaskEdBox1.FormattedText & "~" & Replace(Trim(cmb_reqno.Text), "/", "") & "~" & usr_name
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://intranet.tmc.gov.in/imunorpt/dis_all_1.ASPx?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If

End Sub

Private Sub Form_Load()
'On Error GoTo MSG:

NEWFLAG = False
viewflag = False
FLDENABL ("DIS")
ENABL ("DIS")

DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
If BMAREPORTFLAG = False Then
    BMAREPORTFLAG = True

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1BIBM", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = Trim(LOGINUSRNAME)
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? and MODMO IN ('5','6','7') ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNSTRBM", str)
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
    regcbo.clear
    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 5 And (Trim(REPWTRNRS("MODMO")) = "5" Or Trim(REPWTRNRS("MODMO")) = "7") Then
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        'ElseIf REPWTRNRS("usrtype") = 5 And (Trim(REPWTRNRS("MODMO")) = "6" Or Trim(REPWTRNRS("MODMO")) = "7") Then 'COMM and ADDED for sangrur security by rahul Dated on 21.02.2018
        ElseIf REPWTRNRS("usrtype") > 5 And (Trim(REPWTRNRS("MODMO")) = "6" Or Trim(REPWTRNRS("MODMO")) = "7") Then
            regcbo.AddItem Trim(REPWTRNRS("USERNAME"))
        
        End If
        REPWTRNRS.MoveNext
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BM", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

''**********SECURITY****************************

btnenabl ("INI")

''*******ADDENDUM***********
    str = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QRBM1", str)
'
'    SELDADDENDUM_QR(0) = "FSS"
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
'    STR = "select * from dislib.dissec where ADDENDUMBY='YES'AND MODALITY='BM' "
'    Set seldissecentr11_QR = rdb.CreateQuery("seldissecentr1SS_qr1BM", STR)
'    Set seldissecentr11_RS = seldissecentr11_QR.OpenResultset(1, 3)
'    If Not seldissecentr11_RS.EOF Then
'        Do While Not seldissecentr11_RS.EOF
'            cmb_addendumby.AddItem Trim(seldissecentr11_RS("CCNAME"))
'            seldissecentr11_RS.MoveNext
'        Loop
'    End If
   
''*******ADDENDUM***********

'*********SECURITY**********************************

'STR = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
'Set USRNAME_QR = rdb.CreateQuery("USRNAME1BMA", STR)
'USRNAME_QR(0) = UID
'Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
'If USRNAME_RS.RowCount > 0 Then
'    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
'    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
'    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
'    cmb_enteryby.AddItem LOGINUSRNAME
'Else
'    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
'    cmd_cancel_Click
'    Exit Sub
'End If
'
'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'Set REPWTRNRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USERNAME", 1, 1)
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
'STR = "select * from DISLIB.Dislnkusr where username = ?"
'Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BMA", STR)
'CURUSRDET_QR(0) = "ABC"
'Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'
'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1BMA", STR)
'SELFINALBY_QR(0) = 6
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If


'*********SECURITY**********************************
''for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr1BM", str)
'' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr1BM", str)

'    'for retriving reqnos
Set req_qr = rdb.CreateQuery("reqB2M", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT desc ")

'
Set SELTA101_QR = rdb.CreateQuery("selta101_qr11B2M", "select * from PATOOLSLIB.ta101 where case_no =?")

'    ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1BM", str)

'    ' FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1BM", str)
''
'    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11BM", str)

'
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1BM", str)

    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1BM", str)

   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1BM", str)
'
'   STR = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
'   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1BM", STR)
''
   str = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QRCCBM1", str)
'
'   STR = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
'   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11", STR)
'   SELDADDENDUM_QR(0) = "FSS"
'   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
   str = "select * from dislib.drequisitn where caseno=?"
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1BM", str)

    'commented and new query added by amarjeet on 05-11-2002 to select manjure's name in enter by
    str = "select * from dislib.dissec where enterby='YES'AND MODALITY='BM' "
 
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1BM", str)
   
   str = "select * from dislib.dissec where PROVby='YES'AND MODALITY='BM'"
   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1BM", str)
'
   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='BM'"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1BM", str)
    
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1BM", str)
    
  str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
  Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRBM", str)
    
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
  Set selMA101QR = rdb.CreateQuery("SELMA101QRBM", str)
  

  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET=?"
  Set SELDMA101QR = rdb.CreateQuery("SELDMA101QRBM", str)
  
  str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1BM", str)
  
  
  str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrBM", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%EC%'"
  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRBM", str)
                                             '?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
  'str = "INSERT INTO DISLIB.DBMAREPORT VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  'Commented and added by Shraddha on 21/05/10
  str = "INSERT INTO DISLIB.DBMAREPORT VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  
  Set INSDBMAREPORT_QR = rdb.CreateQuery("INSDBMAREPORT_QRTBM", str)
  
  
                                 ' SET CELLULARIT=?,MERATIO=?,PROMYELOCY=?,MYELOCYTE=?,MMYELOCYTE=?,POLYMORPHS=?,BASOPHILS=?,EOSINOPHIL=?,MONOCYTES=?,PROMONOCYT=?,MLYMPHOCYT=?,ALYMPHOCYT=?,PLASMACELL=?,MATURE=?,IMMATURE=?,ATYPBINU=?,PLASMABLAS=?,ECELLULARI=?,MATURATION=?,MCELLULARI=?,MORPHOLOGY=?,IMPRINT=?,ABCELLBLAS=?,MYELOPEROX=?,MYELOPERO=?,NONSPESTER=?,NONSPESTE1=?,TRAP=?,TRAP1=?,IRONSTAIN=?,IRONSTAIN1=?,GIDEROBLAS=?,SUDANBLACK=?,SUDANBLAC1=?,TOLUIDINEB=?,TOLUIDINE1=?,CD13=?,CD14=?,CD15=?,CD33=?,CD36=?,cd4142=?,GLYCOPHOIN=?,CD68=?,CD=34=?,HLADR=?,CD38=?,CD138=?,TDT=?,ANTIMPO=?,CD10=?,CD19=?,FMC7=?,CD22=?,CD23=?,SMIG=?,CD79B=?,CD25=?,CD11C=?,CD103=?,KAPPA=?,LAMDA=?,DBA44=?,CYTOEP22=?,CYTOCD79A=?,CYTOMU=?,CD2=?,CD3=?,CD4=?,CD5=?,CD7=?,CD8=?,CYTOCD3=?,IMPRESSION=?,COMMENTS=?,FLD1=?,FLD2=? WHERE REQNO=?"
  
  str = " UPDATE DISLIB.DBMAREPORT SET CELLULARIT=?,MERATIO=?,PROMYELOCY=?,MYELOCYTE=?,MMYELOCYTE=?,POLYMORPHS=?,BASOPHILS=?,EOSINOPHIL=?,MONOCYTES=?,PROMONOCYT=?,MLYMPHOCYT=?,ALYMPHOCYT=?,PLASMACELL=?,MATURE=?,IMMATURE=?,ATYPBINU=?,PLASMABLAS=?,ECELLULARI=?,MATURATION=?,MCELLULARI=?,MORPHOLOGY=?,IMPRINT=?,ABCELLBLAS=?,MYELOPEROX=?,MYELOPERO=?,NONSPESTER=?,NONSPESTE1=?,TRAP=?,TRAP1=?,IRONSTAIN=?,IRONSTAIN1=?,GIDEROBLAS=?,SUDANBLACK=?,SUDANBLAC1=?,TOLUIDINEB=?,TOLUIDINE1=?,CD13=?,CD14=?,CD15=?,CD33=?,CD36=?,cd4142=?,GLYCOPHOIN=?,CD68=?,CD34=?,HLADR=?,CD38=?,CD138=?,TDT=?,ANTIMPO=?,CD10=?,CD19=?,FMC7=?,CD22=?,CD23=?,SMIG=?,CD79B=?,CD25=?,CD11C=?,CD103=?,KAPPA=?,LAMDA=?,DBA44=?,CYTOEP22=?,CYTOCD79A=?,CYTOMU=?,CD2=?,CD3=?,CD4=?,CD5=?,CD7=?,CD8=?,CYTOCD3=?,IMPRESSION=?,COMMENTS=?,FLD1=?,FLD2=?,CD42=?,CD117=?,CD20=?,CD61=?,CD56=?,DESCMORPHO=?,TCRAB=?,TCRAG=? WHERE REQNO=?"
  Set UPDDBMAREPORT_QR = rdb.CreateQuery("UPDDBMAREPORT_QRBM", str)
  
  'Shraddha dislib/dbmareport SMPCTG
  '***** Shraddha 09/03/10 *****
  str = "UPDATE DISLIB.DBMAREPORT SET SMPCTG=?, TCRAB=?, TCRAG=?, CLDN1=? WHERE REQNO=?"
  Set UPD_DBMAREPORT_QR = rdb.CreateQuery("", str)
  
  
  '***** Shraddha 12/04/10 *****
  str = "select * from dislib.dcbcsmrep1 where reqno = ?"
  Set selcbcsm_qr = rdb.CreateQuery("", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
  Set SELDREQDRNAME_QR = rdb.CreateQuery("", str)
  
  str = "UPDATE DISLIB.DREQUISITN SET LATERALITS=?, METASTASIS=? WHERE REQNO=?"
  Set UPDDREQDRNAME_QR = rdb.CreateQuery("", str)
  
  str = "insert into dislib.dcbcsmrep1 (Reqno,OPTLSP,OPTSCP) values (?,?,?)"
  Set inscbcsm_qr = rdb.CreateQuery("", str)
  
  
  
  str = "UPDATE DISLIB.DCBCSMREP1 SET OPTLSP=?, OPTSCP=? WHERE REQNO = ?"
  Set updcbcsm_qr = rdb.CreateQuery("", str)
  '*****************************
  
  str = "select * from dislib.dbmareport where reqno=?"
  Set SELQR = rdb.CreateQuery("1", str)
  
  
  
  
  str = "INSERT INTO DISLIB.DBMASUBREP VALUES (?,?,?,?,?,?,?)"
  Set INSDBMASUB_QR = rdb.CreateQuery("INSDBMASUB", str)
  str = "UPDATE DISLIB.DBMASUBREP SET PROLYMPHO = ?, HEMATOGONE = ?, HAIRYCELL = ?, COMMENTS = ?,fullrep = ? WHERE REQNO = ?"
  Set UPDDBMASUB_QR = rdb.CreateQuery("UPDDBMASUB", str)
  
  str = "SELECT * FROM DISLIB.DBMASUBREP WHERE REQNO = ?"
  Set SELDBMASUB_QR = rdb.CreateQuery("SELDBMASUB", str)
  
  End If
 
 
 txt_sudanblackb.Visible = False
 'txt_trapironstain.Visible = False
 txt_nonspecificesterase.Visible = False
 txt_myeloperoxidase.Visible = False
 
  '******added by shweta 10-02-2014
  
   str = "INSERT INTO DISLIB.DCBCREPA (REQNO,FLD7,FLD6) VALUES (?,?,?)"
  Set INSDCBCREPA_QR = rdb.CreateQuery("", str)
  
   str = "UPDATE DISLIB.DCBCREPA SET FLD7=?, FLD6=? WHERE REQNO=?"
  Set UPDDCBCREPA_QR = rdb.CreateQuery("", str)
  
  
  str = "SELECT REQNO,FLD7,FLD6 FROM DISLIB.DCBCREPA WHERE REQNO=?"
  Set SELDCBCREPA_QR = rdb.CreateQuery("", str)
  ''''till here
 
 'THIS IS RTIMLOG
 str = "update rislib.rtimelog set ADPREPDATE=?,ADPREPTIME=? where reqnno=?"
 Set UPDRTIMECADD_QR = rdb.CreateQuery("updatetimecadd", str)
 
 str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=? WHERE REQNNO=?"
 Set UPDRTIME_QR = rdb.CreateQuery("QWEQWEQWE", str)

 str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEC_QR = rdb.CreateQuery("", str)
 
 str = "UPDATE RISLIB.RTIMELOG SET PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEP_QR = rdb.CreateQuery("QWEQWEQWEP", str)

 str = "UPDATE RISLIB.RTIMELOG SET MODREPDATE=?,MODREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEM_QR = rdb.CreateQuery("QWEQWEQWEM", str)

 
 
 
 
 
 
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

Private Sub lbl_time_Click()

End Sub



Private Sub lbl_inmature_Click()

End Sub

Private Sub lbl_myelocytes_Click()

End Sub

Public Sub MaskEdBox1_LostFocus()
Dim CASENO As String
'On Error GoTo MSG:
'***** Shraddha 09/03/10 *****
If Opt_patient.Value = False And opt_staff.Value = False Then
    MsgBox "Please select between PAtient/Staff"
    Exit Sub
End If
'*****************************
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
       '' MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 05/12/2022
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
                req_qr(1) = "BM" 'Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = "BM" 'Trim(txt_modcode.Text)
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
                SELRREQUISITN_QR(1) = "BM" 'Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = "UZZ" 'Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = "BM" 'Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = "UZZ" 'Trim(txt_deptcode.Text)
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
                req_qr(1) = "BM" 'Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = Trim(MaskEdBox1.FormattedText)
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
                SELRREQUISITN_QR(1) = "BM" 'Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = "UZZ" 'Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = "BM" 'Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = "UZZ" 'Trim(txt_deptcode.Text)
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
MsgBox ERROR(Err)
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
                SELDREQSTAFFDQR(0) = Trim(req_chglen(cmb_reqno))
                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
                Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
            Else
                SELDREQSTAFFDQR(0) = Trim(req_chglen(cmb_reqno))
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
          
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                'lbl_ward.Visible = True
'                lbl_bed.Visible = True
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_status.Text = "ST"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                'lbl_ward.Visible = False
'                lbl_bed.Visible = False
            End If
          
         
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
        
            End If
        
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
                       
           
        

        
'        If seldissecentrflag1 = False Then
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

' commeented as finalised by names are taken from security

'''''        If seldissecfinflag = False Then
'''''            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'''''            seldissecfinflag1 = True
'''''        Else
'''''            seldissecfin1_rs.MoreResults
'''''            seldissecfin1_rs.Requery
'''''        End If
'''''        If seldissecfin1_rs.RowCount > 0 Then
'''''            cmb_finalby.clear
'''''            While Not seldissecfin1_rs.EOF
'''''                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'''''                seldissecfin1_rs.MoveNext
'''''            Wend
'''''        End If
            
            '******
          
        
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
                    MsgBox "Final Report is Generated, You can't Modify the Final Report"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                'cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
                End If
                If IsNull(Trim(REFDR_RS("PREPORTBY"))) Then
                    regcbo.Text = ""
                Else
                    regcbo.Text = "" & Trim(REFDR_RS("PREPORTBY"))
                End If
            End If
            
            
            
          
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
'                lbl_ward.Visible = False
'                lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
'                lbl_ward.Visible = False
'                lbl_bed.Visible = False
                txt_status.Text = "ST"
            End If
          
         If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                'txt_ihmeterial.Text = SELPATHDREQUISITN_RS("NMETERIALS")
                txt_refdoc.Text = selpathdrequisitn_rs("REFDOC")
                TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
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
                        '*******
            '***********************************************
            ''*******ADDENDUM*********************************
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
            ''*******ADDENDUM*********************************
            
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
                txt_procedure.Text = Trim(selta1012_rs("prvdiagdes"))
            End If
            If SELDBMAREPORTFLAG = False Then
                SELQR(0) = req_chglen(cmb_reqno.Text)
                Set SELDBMAREPORT_RS = SELQR.OpenResultset(1, 3)
                SELDBMAREPORTFLAG = True
            Else
                  SELQR(0) = req_chglen(cmb_reqno.Text)
                  SELDBMAREPORT_RS.MoreResults
                  SELDBMAREPORT_RS.Requery
            End If
            ENABL ("DIS")
            If SELDBMAREPORT_RS.RowCount > 0 Then
                    txt_cellularity.Text = " " & SELDBMAREPORT_RS("CELLULARIT")
                    If IsNull(SELDBMAREPORT_RS("MERATIO")) Then
                    Else
                        I = InStr(SELDBMAREPORT_RS("MERATIO"), ":")
                        If I <> 0 Then
                            TXT_MERATION.Text = " " & Mid(SELDBMAREPORT_RS("MERATIO"), 1, I - 1)
                            TXT_MERATION1.Text = " " & Mid(SELDBMAREPORT_RS("MERATIO"), I + 1, 2)
                        Else
                            TXT_MERATION.Text = Trim(SELDBMAREPORT_RS("MERATIO"))
                        End If
                    End If
                    TXT_COMMENTS.Text = "" & Trim(SELDBMAREPORT_RS("COMMENTS"))
                    txt_promyelocytes.Text = " " & Trim(SELDBMAREPORT_RS("PROMYELOCY"))
                    txt_myelocytes.Text = " " & Trim(SELDBMAREPORT_RS("MYELOCYTE"))
                    txt_metamyelocytes.Text = " " & Trim(SELDBMAREPORT_RS("MMYELOCYTE"))
                    txt_polymorphs.Text = " " & Trim(SELDBMAREPORT_RS("POLYMORPHS"))
                    txt_basophils.Text = " " & Trim(SELDBMAREPORT_RS("BASOPHILS"))
                    txt_eosinophils.Text = " " & Trim(SELDBMAREPORT_RS("EOSINOPHIL"))
                    txt_monocytes.Text = " " & Trim(SELDBMAREPORT_RS("MONOCYTES"))
                    TXT_PROMONOCYTES.Text = " " & Trim(SELDBMAREPORT_RS("PROMONOCYT"))
                    txt_matly.Text = " " & Trim(SELDBMAREPORT_RS("MLYMPHOCYT"))
                    txt_atyply.Text = " " & Trim(SELDBMAREPORT_RS("ALYMPHOCYT"))
                    txt_plasmacells.Text = " " & Trim(SELDBMAREPORT_RS("PLASMACELL"))
                    txt_mature.Text = " " & Trim(SELDBMAREPORT_RS("MATURE"))
                    txt_inmature.Text = " " & Trim(SELDBMAREPORT_RS("IMMATURE"))
                    txt_atypbi.Text = " " & Trim(SELDBMAREPORT_RS("ATYPBINU"))
                    txt_plasmablasts.Text = " " & Trim(SELDBMAREPORT_RS("PLASMABLAS"))
                    txt_cellularity1.Text = " " & Trim(SELDBMAREPORT_RS("ECELLULARI"))
                    txt_maturation.Text = " " & Trim(SELDBMAREPORT_RS("MATURATION"))
                    
                    'ADDED BY SHWETA ON 19-07-2013 AS PER REQUEST UZZ DEPT.
                    Dim RS_ERYTHCELL As rdoResultset
                    Dim ERYTHSTR As String
                    ERYTHSTR = "SELECT * FROM DISLIB.DBMASUBNEW WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "'"
                    Set RS_ERYTHCELL = rdb.OpenResultset(ERYTHSTR, 1, 1)
                       TXT_ErythCells.Text = RS_ERYTHCELL(1)
                    
                    
                    
                    txt_cellularity2.Text = " " & Trim(SELDBMAREPORT_RS("MCELLULARI"))
                    txt_morphology.Text = " " & Trim(SELDBMAREPORT_RS("MORPHOLOGY"))
                    TXT_IMPRINT.Text = " " & Trim(SELDBMAREPORT_RS("IMPRINT"))
                    txt_descmorpho.Text = "" & Trim(SELDBMAREPORT_RS("DESCMORPHO"))
                    txt_abcellblasts.Text = "" & Trim(SELDBMAREPORT_RS("ABCELLBLAS"))
                    If Trim(SELDBMAREPORT_RS("myeloperox")) = "Positive" Then
                        CMB_MYELOPEROXIDASE.Text = " " & Trim(SELDBMAREPORT_RS("myeloperox"))
                        txt_myeloperoxidase.Visible = True
                        txt_myeloperoxidase.Text = " " & Trim(SELDBMAREPORT_RS("MYELOPERO"))
                    Else
                        CMB_MYELOPEROXIDASE.Text = " " & Trim(SELDBMAREPORT_RS("myeloperox"))
                        txt_myeloperoxidase.Visible = False
                    End If
                    If Trim(SELDBMAREPORT_RS("NONSPESTER")) = "Positive" Then
                        CMB_NONSPEST.Text = " " & Trim(SELDBMAREPORT_RS("NONSPESTER"))
                        txt_nonspecificesterase.Visible = True
                        txt_nonspecificesterase.Text = " " & Trim(SELDBMAREPORT_RS("NONSPESTE1"))
                    Else
                        CMB_NONSPEST.Text = " " & Trim(SELDBMAREPORT_RS("NONSPESTER"))
                        txt_nonspecificesterase.Visible = False
                    End If
                    'If Trim(SELDBMAREPORT_RS("TRAP")) = "Positive" Then
                        'CMB_TRAP.Text = " " & SELDBMAREPORT_RS("TRAP")
                        'txt_trapironstain.Visible = True
                        txt_trapironstain.Text = " " & Trim(SELDBMAREPORT_RS("TRAP1"))
                    'Else
                        'CMB_TRAP.Text = " " & SELDBMAREPORT_RS("TRAP")
                        'txt_trapironstain.Visible = False
                    'End If
                    'txt_trapironstain.Enabled = False
                    'If Trim(SELDBMAREPORT_RS("IRONSTAIN")) = "Positive" Then
                     '   CMB_IRONSTAIN.Text = " " & SELDBMAREPORT_RS("IRONSTAIN")
'                        TXT_IRONSTAIN.Visible = True
                        TXT_IRONSTAIN.Text = " " & Trim(SELDBMAREPORT_RS("IRONSTAIN1"))
                    'Else
                     '   CMB_IRONSTAIN.Text = " " & SELDBMAREPORT_RS("IRONSTAIN")
'                        TXT_IRONSTAIN.Visible = False
                    'End If
                    'If Trim(SELDBMAREPORT_RS("GIDEROBLAS")) = "Positive" Then
                     '   CMB_GIDEROBLASTS.Text = " " & SELDBMAREPORT_RS("GIDEROBLAS")
                      '  TXT_GIDEROBLASTS.Visible = True
                        TXT_GIDEROBLASTS.Text = " " & Trim(SELDBMAREPORT_RS("fld1"))
                    'Else
                     '   CMB_GIDEROBLASTS.Text = " " & SELDBMAREPORT_RS("GIDEROBLAS")
                      '  TXT_GIDEROBLASTS.Visible = False
                        
                    'End If
                    If Trim(SELDBMAREPORT_RS("sudanblack")) = "Positive" Then
                        CMB_SUDANBLACKB.Text = " " & Trim(SELDBMAREPORT_RS("sudanblack"))
                        txt_sudanblackb.Visible = True
                        txt_sudanblackb.Text = " " & Trim(SELDBMAREPORT_RS("sudanblac1"))
                    Else
                    CMB_SUDANBLACKB.Text = " " & Trim(SELDBMAREPORT_RS("sudanblack"))
                    txt_sudanblackb.Visible = False
                   
                    End If
                    If Trim(SELDBMAREPORT_RS("toluidineb")) = "Positive" Then
                        CMB_TOLUIDINEBLUE.Text = " " & Trim(SELDBMAREPORT_RS("toluidineb"))
                        txt_TOLUIDINEBLUE.Visible = True
                        txt_TOLUIDINEBLUE.Text = " " & Trim(SELDBMAREPORT_RS("toluidine1"))
                    Else
                        CMB_TOLUIDINEBLUE.Text = " " & Trim(SELDBMAREPORT_RS("toluidineb"))
                        txt_TOLUIDINEBLUE.Visible = False
                        
                    End If
                    txt_cd13.Text = " " & Trim(SELDBMAREPORT_RS("cd13"))
                    txt_cd14.Text = " " & Trim(SELDBMAREPORT_RS("CD14"))
                    txt_CD15.Text = " " & Trim(SELDBMAREPORT_RS("CD15"))
                    txt_CD33.Text = " " & Trim(SELDBMAREPORT_RS("CD33"))
                    txt_cd36.Text = " " & Trim(SELDBMAREPORT_RS("CD36"))
                    txt_cd41.Text = " " & Trim(SELDBMAREPORT_RS("CD4142"))
                    TXT_CD42.Text = "" & Trim(SELDBMAREPORT_RS("CD42"))
                    TXT_CD117.Text = "" & Trim(SELDBMAREPORT_RS("CD117"))
                    TXT_CD20.Text = "" & Trim(SELDBMAREPORT_RS("CD20"))
                    TXT_CD61.Text = "" & Trim(SELDBMAREPORT_RS("CD61"))
                    TXT_CD56.Text = "" & Trim(SELDBMAREPORT_RS("CD56"))
                    txt_Glycophorina.Text = " " & Trim(SELDBMAREPORT_RS("GLYCOPHOIN"))
                    TXT_CD68.Text = " " & Trim(SELDBMAREPORT_RS("CD68"))
                    TXT_CD34.Text = " " & Trim(SELDBMAREPORT_RS("CD34"))
                    TXT_HLADR.Text = " " & Trim(SELDBMAREPORT_RS("HLADR"))
                    TXT_CD38.Text = " " & Trim(SELDBMAREPORT_RS("CD38"))
                    TXT_CD138.Text = " " & Trim(SELDBMAREPORT_RS("CD138"))
                    TXT_TDT.Text = " " & Trim(SELDBMAREPORT_RS("TDT"))
                    TXT_ANTIMPO.Text = " " & Trim(SELDBMAREPORT_RS("ANTIMPO"))
                    TXT_CD10.Text = " " & Trim(SELDBMAREPORT_RS("CD10"))
                    TXT_CD19.Text = " " & Trim(SELDBMAREPORT_RS("CD19"))
                    TXT_FMC7.Text = " " & Trim(SELDBMAREPORT_RS("FMC7"))
                    TXT_CD22.Text = " " & Trim(SELDBMAREPORT_RS("CD22"))
                    TXT_CD23.Text = " " & Trim(SELDBMAREPORT_RS("CD23"))
                    txt_smig.Text = " " & Trim(SELDBMAREPORT_RS("SMIG"))
                    
                    '***** Shraddha 09/03/10 *****
                    If Trim(SELDBMAREPORT_RS("SMPCTG")) <> "" Then
                        TXT_SMPCTG.Text = "" & Trim(SELDBMAREPORT_RS("SMPCTG"))
                    End If
                    '*****************************
                    
                    txt_CD79B.Text = " " & Trim(SELDBMAREPORT_RS("CD79B"))
                    TXT_CD25.Text = " " & Trim(SELDBMAREPORT_RS("CD25"))
                    TXT_CD11C.Text = " " & Trim(SELDBMAREPORT_RS("CD11C"))
                    txt_cd103.Text = " " & Trim(SELDBMAREPORT_RS("CD103"))
                    txt_kappa.Text = " " & Trim(SELDBMAREPORT_RS("KAPPA"))
                    txt_lanbda.Text = " " & Trim(SELDBMAREPORT_RS("LAMDA"))
                    TXT_DBA44.Text = " " & Trim(SELDBMAREPORT_RS("DBA44"))
                    txt_cytocd221(1).Text = " " & Trim(SELDBMAREPORT_RS("CYTOEP22"))
                    txt_cytocd79a.Text = " " & Trim(SELDBMAREPORT_RS("CYTOCD79A"))
                    TXT_CYTOMU.Text = " " & Trim(SELDBMAREPORT_RS("CYTOMU"))
                    TXT_CD2.Text = " " & Trim(SELDBMAREPORT_RS("CD2"))
                    TXT_CD3.Text = " " & Trim(SELDBMAREPORT_RS("CD3"))
                    TXT_CD4.Text = " " & Trim(SELDBMAREPORT_RS("CD4"))
                    TXT_CD5.Text = " " & Trim(SELDBMAREPORT_RS("CD5"))
                    TXT_CD7.Text = " " & Trim(SELDBMAREPORT_RS("CD7"))
                    TXT_CD8.Text = " " & Trim(SELDBMAREPORT_RS("CD8"))
                    TXT_CYTOCD3.Text = " " & Trim(SELDBMAREPORT_RS("CYTOCD3"))
                    txt_IMPRESSION.Text = "" & Trim(SELDBMAREPORT_RS("impression"))
            End If
             ' FETCHING FROM DBMASUB
          SELDBMASUB_QR(0) = Trim(req_chglen(cmb_reqno.Text))
          
          
          Set SELDBMASUB_RS = SELDBMASUB_QR.OpenResultset(1, 1)
            If SELDBMASUB_RS.RowCount > 0 Then
                   Text2.Text = " " & Trim(SELDBMASUB_RS("PROLYMPHO"))
                   Text3.Text = " " & Trim(SELDBMASUB_RS("HEMATOGONE"))
                   Text1.Text = " " & Trim(SELDBMASUB_RS("HAIRYCELL"))
                   TXT_COMMENTS = " " & Trim(SELDBMASUB_RS("COMMENTS"))
                   If Trim(SELDBMASUB_RS("fullrep")) = "Y" Then
                      Check1.Value = 1
                   ElseIf Trim(SELDBMASUB_RS("fullrep")) = "N" Then
                      Check1.Value = 0
                   End If
                   Check1.Enabled = False
            End If
            
           ' FETCHING FROM DBMASUB ENDS
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
                txt_refdoc.Text = "" & REFDR_RS("REFBYDOC")
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Report is commited , You can not Modify"
                    cmd_modify.Enabled = False
                    FLDENABL ("ADDENDUM")
                    ADDENDUMFALG = True
                    cmd_save.Enabled = True
                End If
                cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
                'CMB_PROVISIONAL.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = "" & Trim(REFDR_RS("FREPORTBY"))
                End If
                
            End If
            
            '***********************************************
                   
            
        End If
End Sub

Public Sub assign()
'On Error GoTo MSG:
NOTHINGWRITTEN = False
 Dim erythcells As String
rdb.BeginTrans
    If flagsave = 1 Then
        '*** Inserting into  DBMAREPORT
        INSDBMAREPORT_QR(0) = req_chglen(cmb_reqno.Text)
        If Len(Trim(txt_cellularity)) > 0 Then
            INSDBMAREPORT_QR(1) = Null
            INSDBMAREPORT_QR(1) = txt_cellularity.Text
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(1) = "  "
        End If
        If Len(Trim(TXT_MERATION.Text)) > 0 Then
            INSDBMAREPORT_QR(2) = Trim(TXT_MERATION.Text) & ":" & Trim(TXT_MERATION1.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(2) = "  "
        End If
        If Len(Trim(txt_promyelocytes.Text)) > 0 Then
            INSDBMAREPORT_QR(3) = Trim(txt_promyelocytes.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(3) = "  "
        End If
        If Len(Trim(txt_myelocytes.Text)) > 0 Then
            INSDBMAREPORT_QR(4) = Trim(txt_myelocytes.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(4) = "  "
        End If
        If Len(Trim(txt_metamyelocytes.Text)) > 0 Then
            INSDBMAREPORT_QR(5) = Trim(txt_metamyelocytes.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(5) = "  "
        End If
        If Len(Trim(txt_polymorphs.Text)) > 0 Then
            INSDBMAREPORT_QR(6) = Trim(txt_polymorphs.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(6) = "  "
        End If
        If Len(Trim(txt_basophils.Text)) > 0 Then
            INSDBMAREPORT_QR(7) = Trim(txt_basophils.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(7) = "  "
        End If
        If Len(Trim(txt_eosinophils.Text)) > 0 Then
            INSDBMAREPORT_QR(8) = Trim(txt_eosinophils.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(8) = "  "
        End If
        If Len(Trim(txt_monocytes.Text)) > 0 Then
            INSDBMAREPORT_QR(9) = Trim(txt_monocytes.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(9) = "  "
        End If
        If Len(Trim(TXT_PROMONOCYTES.Text)) > 0 Then
            INSDBMAREPORT_QR(10) = Trim(TXT_PROMONOCYTES.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(10) = "  "
        End If
        If Len(Trim(txt_matly.Text)) > 0 Then
            INSDBMAREPORT_QR(11) = Trim(txt_matly.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(11) = "  "
        End If
        If Len(Trim(txt_atyply.Text)) > 0 Then
            INSDBMAREPORT_QR(12) = Trim(txt_atyply.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(12) = "  "
        End If
        If Len(Trim(txt_plasmacells.Text)) > 0 Then
            INSDBMAREPORT_QR(13) = Trim(txt_plasmacells.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(13) = "  "
        End If
        If Len(Trim(txt_mature.Text)) > 0 Then
            INSDBMAREPORT_QR(14) = Trim(txt_mature.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(14) = "  "
        End If
        If Len(Trim(txt_inmature.Text)) > 0 Then
            INSDBMAREPORT_QR(15) = Trim(txt_inmature.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(15) = "  "
        End If
        If Len(Trim(txt_atypbi.Text)) > 0 Then
            INSDBMAREPORT_QR(16) = Trim(txt_atypbi.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(16) = "  "
        End If
        If Len(Trim(txt_plasmablasts.Text)) > 0 Then
            INSDBMAREPORT_QR(17) = Trim(txt_plasmablasts.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(17) = "  "
        End If
        If Len(Trim(txt_cellularity1.Text)) > 0 Then
            INSDBMAREPORT_QR(18) = Null
            INSDBMAREPORT_QR(18) = Trim(txt_cellularity1.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(18) = "  "
        End If
        If Len(Trim(txt_maturation.Text)) > 0 Then
            INSDBMAREPORT_QR(19) = Null
            INSDBMAREPORT_QR(19) = Trim(txt_maturation.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(19) = "  "
        End If
        If Len(Trim(txt_cellularity2.Text)) > 0 Then
            INSDBMAREPORT_QR(20) = Null
            INSDBMAREPORT_QR(20) = Trim(txt_cellularity2.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(20) = "  "
        End If
        If Len(Trim(txt_morphology.Text)) > 0 Then
            INSDBMAREPORT_QR(21) = Null
            INSDBMAREPORT_QR(21) = Trim(txt_morphology.Text)
        Else
             NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(21) = "  "
        End If
        If Len(Trim(TXT_IMPRINT.Text)) > 0 Then
            INSDBMAREPORT_QR(22) = Null
            INSDBMAREPORT_QR(22) = Trim(TXT_IMPRINT.Text)
        Else
             NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(22) = "  "
        End If
        If Len(Trim(txt_abcellblasts.Text)) > 0 Then
            INSDBMAREPORT_QR(23) = Trim(txt_abcellblasts.Text)
        Else
             NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(23) = "  "
        End If
        If Len(Trim(CMB_MYELOPEROXIDASE.Text)) > 0 Then
            INSDBMAREPORT_QR(24) = Trim(CMB_MYELOPEROXIDASE.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(24) = "  "
        End If
        If Len(Trim(txt_myeloperoxidase.Text)) > 0 Then
            INSDBMAREPORT_QR(25) = Trim(txt_myeloperoxidase.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(25) = "  "
        End If
        If Len(Trim(CMB_NONSPEST.Text)) > 0 Then
            INSDBMAREPORT_QR(26) = Trim(CMB_NONSPEST.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(26) = "  "
        End If
        If Len(Trim(txt_nonspecificesterase.Text)) > 0 Then
            INSDBMAREPORT_QR(27) = Trim(txt_nonspecificesterase.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(27) = "  "
        End If
        If Len(Trim(CMB_TRAP.Text)) > 0 Then
            INSDBMAREPORT_QR(28) = Trim(CMB_TRAP.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(28) = "  "
        End If
        If Len(Trim(txt_trapironstain.Text)) > 0 Then
            INSDBMAREPORT_QR(29) = Trim(txt_trapironstain.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(29) = "  "
        End If
        If Len(Trim(CMB_IRONSTAIN.Text)) > 0 Then
            INSDBMAREPORT_QR(30) = Trim(CMB_IRONSTAIN.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(30) = "  "
        End If
        If Len(Trim(TXT_IRONSTAIN.Text)) > 0 Then
            INSDBMAREPORT_QR(31) = Trim(TXT_IRONSTAIN.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(31) = "  "
        End If
        If Len(Trim(CMB_GIDEROBLASTS.Text)) > 0 Then
            INSDBMAREPORT_QR(32) = Trim(CMB_GIDEROBLASTS.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(32) = "  "
        End If
'        If Len(Trim(TXT_GIDEROBLASTS.Text)) > 0 Then
'        End If
        If Len(Trim(CMB_SUDANBLACKB.Text)) > 0 Then
            INSDBMAREPORT_QR(33) = Trim(CMB_SUDANBLACKB.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(33) = "  "
        End If
        If Len(Trim(txt_sudanblackb.Text)) > 0 Then
            INSDBMAREPORT_QR(34) = Trim(txt_sudanblackb.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(34) = "  "
        End If
        If Len(Trim(CMB_TOLUIDINEBLUE.Text)) > 0 Then
            INSDBMAREPORT_QR(35) = Trim(CMB_TOLUIDINEBLUE.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(35) = "  "
        End If
        If Len(Trim(txt_TOLUIDINEBLUE.Text)) > 0 Then
            INSDBMAREPORT_QR(36) = Trim(txt_TOLUIDINEBLUE.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(36) = "  "
        End If
        '**** SURFACE MARKERS****
        If Len(Trim(txt_cd13.Text)) > 0 Then
            INSDBMAREPORT_QR(37) = Trim(txt_cd13.Text)
        Else
             NOTHINGWRITTEN = True
             INSDBMAREPORT_QR(37) = "  "
        End If
        If Len(Trim(txt_cd14.Text)) > 0 Then
             INSDBMAREPORT_QR(38) = Trim(txt_cd14.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(38) = "  "
        End If
        If Len(Trim(txt_CD15.Text)) > 0 Then
            INSDBMAREPORT_QR(39) = Trim(txt_CD15.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(39) = "  "
        End If
        If Len(Trim(txt_CD33.Text)) > 0 Then
            INSDBMAREPORT_QR(40) = Trim(txt_CD33.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(40) = "  "
        End If
        If Len(Trim(txt_cd36.Text)) > 0 Then
            INSDBMAREPORT_QR(41) = Trim(txt_cd36.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(41) = "  "
        End If
        If Len(Trim(txt_cd41.Text)) > 0 Then
            INSDBMAREPORT_QR(42) = Trim(txt_cd41.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(42) = "  "
        End If
        If Len(Trim(txt_Glycophorina.Text)) > 0 Then
            INSDBMAREPORT_QR(43) = Trim(txt_Glycophorina.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(43) = "  "
        End If
        If Len(Trim(TXT_CD68.Text)) > 0 Then
            INSDBMAREPORT_QR(44) = Trim(TXT_CD68.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(44) = "  "
        End If
        If Len(Trim(TXT_CD34.Text)) > 0 Then
            INSDBMAREPORT_QR(45) = Trim(TXT_CD34.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(45) = "  "
        End If
        If Len(Trim(TXT_HLADR.Text)) > 0 Then
            INSDBMAREPORT_QR(46) = Trim(TXT_HLADR.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(46) = "  "
        End If
        If Len(Trim(TXT_CD38.Text)) > 0 Then
            INSDBMAREPORT_QR(47) = Trim(TXT_CD38.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(47) = "  "
        End If
        If Len(Trim(TXT_CD138.Text)) > 0 Then
            INSDBMAREPORT_QR(48) = Trim(TXT_CD138.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(48) = "  "
        End If
        If Len(Trim(TXT_TDT.Text)) > 0 Then
            INSDBMAREPORT_QR(49) = Trim(TXT_TDT.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(49) = "  "
        End If
        If Len(Trim(TXT_ANTIMPO.Text)) > 0 Then
            INSDBMAREPORT_QR(50) = Trim(TXT_ANTIMPO.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(50) = "  "
        End If
        If Len(Trim(TXT_CD10.Text)) > 0 Then
            INSDBMAREPORT_QR(51) = Trim(TXT_CD10.Text)
        Else
             NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(51) = "  "
        End If
        If Len(Trim(TXT_CD19.Text)) > 0 Then
            INSDBMAREPORT_QR(52) = Trim(TXT_CD19.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(52) = "  "
        End If
        If Len(Trim(TXT_FMC7.Text)) > 0 Then
            INSDBMAREPORT_QR(53) = Trim(TXT_FMC7.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(53) = "  "
        End If
        If Len(Trim(TXT_CD22.Text)) > 0 Then
            INSDBMAREPORT_QR(54) = Trim(TXT_CD22.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(54) = "  "
        End If
        If Len(Trim(TXT_CD23.Text)) > 0 Then
            INSDBMAREPORT_QR(55) = Trim(TXT_CD23.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(55) = "  "
        End If
        If Len(Trim(txt_smig.Text)) > 0 Then
            INSDBMAREPORT_QR(56) = Trim(txt_smig.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(56) = "  "
        End If
        
        
        
        If Len(Trim(txt_CD79B.Text)) > 0 Then
            INSDBMAREPORT_QR(57) = Trim(txt_CD79B.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(57) = "  "
        End If
        If Len(Trim(TXT_CD25.Text)) > 0 Then
            INSDBMAREPORT_QR(58) = Trim(TXT_CD25.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(58) = "  "
        End If
        If Len(Trim(TXT_CD11C.Text)) > 0 Then
            INSDBMAREPORT_QR(59) = Trim(TXT_CD11C.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(59) = "  "
        End If
        If Len(Trim(txt_cd103.Text)) > 0 Then
            INSDBMAREPORT_QR(60) = Trim(txt_cd103.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(60) = "  "
        End If
        If Len(Trim(txt_kappa.Text)) > 0 Then
            INSDBMAREPORT_QR(61) = Trim(txt_kappa.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(61) = "  "
        End If
        If Len(Trim(txt_lanbda.Text)) > 0 Then
            INSDBMAREPORT_QR(62) = Trim(txt_lanbda.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(62) = "  "
        End If
        If Len(Trim(TXT_DBA44.Text)) > 0 Then
            INSDBMAREPORT_QR(63) = Trim(TXT_DBA44.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(63) = "  "
        End If
        If Len(Trim(txt_cytocd221(1))) > 0 Then
            INSDBMAREPORT_QR(64) = Trim(txt_cytocd221(1))
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(64) = "  "
        End If
        If Len(Trim(txt_cytocd79a.Text)) > 0 Then
            INSDBMAREPORT_QR(65) = Trim(txt_cytocd79a.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(65) = "  "
        End If
        If Len(Trim(TXT_CYTOMU.Text)) > 0 Then
            INSDBMAREPORT_QR(66) = Trim(TXT_CYTOMU.Text)
        Else
            NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(66) = "  "
        End If
        If Len(Trim(TXT_CD2.Text)) > 0 Then
            INSDBMAREPORT_QR(67) = Trim(TXT_CD2.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(67) = "  "
        End If
        If Len(Trim(TXT_CD3.Text)) > 0 Then
            INSDBMAREPORT_QR(68) = Trim(TXT_CD3.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(68) = "  "
        End If
        If Len(Trim(TXT_CD4.Text)) > 0 Then
            INSDBMAREPORT_QR(69) = Trim(TXT_CD4.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(69) = "  "
        End If
        If Len(Trim(TXT_CD5.Text)) > 0 Then
            INSDBMAREPORT_QR(70) = Trim(TXT_CD5.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(70) = "  "
        End If
        If Len(Trim(TXT_CD7.Text)) > 0 Then
            INSDBMAREPORT_QR(71) = Trim(TXT_CD7.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(71) = "  "
        End If
        If Len(Trim(TXT_CD8.Text)) > 0 Then
            INSDBMAREPORT_QR(72) = Trim(TXT_CD8.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(72) = "  "
        End If
        If Len(Trim(TXT_CYTOCD3.Text)) > 0 Then
            INSDBMAREPORT_QR(73) = Trim(TXT_CYTOCD3.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(73) = "  "
        End If
        If Len(Trim(txt_IMPRESSION.Text)) > 0 Then
            INSDBMAREPORT_QR(74) = Null
            INSDBMAREPORT_QR(74) = Trim(txt_IMPRESSION.Text)
        Else
            'NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(74) = "  "
        End If
        If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
            INSDBMAREPORT_QR(75) = Null
            INSDBMAREPORT_QR(75) = Trim(TXT_COMMENTS.Text)
        Else
            NOTHINGWRITTEN = True
            INSDBMAREPORT_QR(75) = "  "
        End If
        If Len(Trim(TXT_GIDEROBLASTS.Text)) > 0 Then
            INSDBMAREPORT_QR(76) = Trim(TXT_GIDEROBLASTS.Text)
        Else
            NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(76) = "  "
        End If
        If Len(Trim(txt_igm)) > 0 Then    ' fld2 is used for igm
         INSDBMAREPORT_QR(77) = Trim(txt_igm.Text)
        Else
          INSDBMAREPORT_QR(77) = "  "
        End If
        
        If Len(Trim(TXT_CD42.Text)) > 0 Then
            INSDBMAREPORT_QR(78) = Trim(TXT_CD42.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(78) = "  "
        End If
        If Len(Trim(TXT_CD117.Text)) > 0 Then
            INSDBMAREPORT_QR(79) = Trim(TXT_CD117.Text)
        Else
            NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(79) = "  "
        End If
        If Len(Trim(TXT_CD20.Text)) > 0 Then
            INSDBMAREPORT_QR(80) = Trim(TXT_CD20.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(80) = "  "
        End If
        If Len(Trim(TXT_CD61.Text)) > 0 Then
            INSDBMAREPORT_QR(81) = Trim(TXT_CD61.Text)
        Else
           NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(81) = "  "
        End If
        If Len(Trim(TXT_CD56.Text)) > 0 Then
            INSDBMAREPORT_QR(82) = Trim(TXT_CD56.Text)
        Else
            NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(82) = "  "
        End If
        
        If Len(Trim(txt_descmorpho.Text)) > 0 Then
            INSDBMAREPORT_QR(83) = Trim(txt_descmorpho.Text)
        Else
            NOTHINGWRITTEN = True
           INSDBMAREPORT_QR(83) = "  "
        End If
        
            INSDBMAREPORT_QR(84) = ""
            INSDBMAREPORT_QR(85) = ""
            
            
            '***** Shraddha 21/05/10 *****
            INSDBMAREPORT_QR(86) = ""
            INSDBMAREPORT_QR(87) = ""
            INSDBMAREPORT_QR(88) = ""
            INSDBMAREPORT_QR(89) = ""
            
            '*****************************
        INSDBMAREPORT_QR.Execute
        
        
           
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
       
        erythcells = "insert into DISLIB.DBMASUBNEW(REQNO,ERYTHROCT,SAVE_DT,SAVE_TM,FLD1) VALUES('" & Trim(req_chglen(cmb_reqno.Text)) & "','" & Trim(TXT_ErythCells.Text) & "',CURRENT DATE,CURRENT TIME,'I')"
        rdb.Execute erythcells
        
        
        '***** Shraddha 09/03/10 *****
        'UPDATE DISLIB.DBMAREPORT SET SMPCTG=?, TCRAB=?, TCRAG=? WHERE REQNO=?"
        UPD_DBMAREPORT_QR(0) = Trim(TXT_SMPCTG.Text) 'SMPCTG
        UPD_DBMAREPORT_QR(1) = Trim(txt_tcrab.Text) 'TCRAB
        UPD_DBMAREPORT_QR(2) = Trim(txt_tcrag.Text) 'TCRAG
        UPD_DBMAREPORT_QR(3) = Trim(txt_cyclinD1.Text) 'CLDN1
        UPD_DBMAREPORT_QR(4) = req_chglen(cmb_reqno.Text) 'REQNO
        UPD_DBMAREPORT_QR.Execute
        
        
        
        '***** Shraddha 12/04/10 *****
        inscbcsm_qr(0) = req_chglen(cmb_reqno.Text)
        If opt_lsp.Value = True Then
            inscbcsm_qr(1) = "Y"
        Else
            inscbcsm_qr(1) = "N"
        End If
            
        If opt_ssp.Value = True Then
            inscbcsm_qr(2) = "Y"
        Else
            inscbcsm_qr(2) = "N"
        End If
            
        inscbcsm_qr.Execute
        
        If opt_drak.Value = True Then
            UPDDREQDRNAME_QR(0) = "Dr. Ashok Kumar"
        ElseIf opt_drak.Value = False Then
            UPDDREQDRNAME_QR(0) = ""
        End If
'        If opt_mryb.Value = True Then
'            UPDDREQDRNAME_QR(1) = "Mr. Y Badrinath"
'        ElseIf opt_mryb.Value = False Then
'            UPDDREQDRNAME_QR(1) = ""
'        End If
        '***** Commented and added by Shraddha on 05/05/10 *****
        If opt_mryb.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Y. Badrinath"
        ElseIf opt_mryb.Value = False And opt_sitaram.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Sitaram"
        Else
            UPDDREQDRNAME_QR(1) = ""
        End If
        '*****************************************
        UPDDREQDRNAME_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDREQDRNAME_QR.Execute
        '*****************************
        
        ' INSERT ING DBMASUBREP
          
        ' inserting dbmasubrep
           INSDBMASUB_QR(0) = Trim(req_chglen(cmb_reqno.Text))
          '  INSDBMASUB_QR(0) = Trim(cmb_reqno.Text)
            If Len(Trim(Text2.Text)) > 0 Then
            
             INSDBMASUB_QR(1) = Text2.Text
            Else
               NOTHINGWRITTEN = True
              INSDBMASUB_QR(1) = "  "
            End If
            If Len(Trim(Text3.Text)) > 0 Then
            
             INSDBMASUB_QR(2) = Text3.Text
            Else
              NOTHINGWRITTEN = True
              INSDBMASUB_QR(2) = "  "
            End If
            If Len(Trim(Text1.Text)) > 0 Then
            
             INSDBMASUB_QR(3) = Text1.Text
            Else
               NOTHINGWRITTEN = True
              INSDBMASUB_QR(3) = "  "
            End If
            
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
               INSDBMASUB_QR(4) = Trim(TXT_COMMENTS.Text)
            Else
               NOTHINGWRITTEN = True
               INSDBMASUB_QR(4) = "  "
            End If
            If Check1.Value = 1 Then
                INSDBMASUB_QR(5) = "Y"
            ElseIf Check1.Value = 0 Then
                INSDBMASUB_QR(5) = "N"
            End If
            
            INSDBMASUB_QR.Execute
       '*** Inserting into DBMAREPORT
                       
                       
       '***** INSERTING TO DBMASUBNEW FOR ERYTHROCYTE CELLS
       
       
       '''''
       
       
                       
            
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = regcbo.Text   'cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
        
            
            
        
    
    End If
    If flagsave = 2 Then
            
            '*** updating the DHBFREPORT
            
        
        If Len(Trim(txt_cellularity.Text)) > 0 Then
            UPDDBMAREPORT_QR(0) = Null
            UPDDBMAREPORT_QR(0) = txt_cellularity.Text
        Else
            UPDDBMAREPORT_QR(0) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(TXT_MERATION.Text)) > 0 Then
            UPDDBMAREPORT_QR(1) = Trim(TXT_MERATION.Text) & ":" & Trim(TXT_MERATION1.Text)
        Else
            UPDDBMAREPORT_QR(1) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_promyelocytes.Text)) > 0 Then
            UPDDBMAREPORT_QR(2) = Trim(txt_promyelocytes.Text)
        Else
            UPDDBMAREPORT_QR(2) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_myelocytes.Text)) > 0 Then
            UPDDBMAREPORT_QR(3) = Trim(txt_myelocytes.Text)
        Else
            UPDDBMAREPORT_QR(3) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_metamyelocytes.Text)) > 0 Then
            UPDDBMAREPORT_QR(4) = Trim(txt_metamyelocytes.Text)
        Else
            UPDDBMAREPORT_QR(4) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_polymorphs.Text)) > 0 Then
            UPDDBMAREPORT_QR(5) = Trim(txt_polymorphs.Text)
        Else
            UPDDBMAREPORT_QR(5) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_basophils.Text)) > 0 Then
            UPDDBMAREPORT_QR(6) = Trim(txt_basophils.Text)
        Else
            UPDDBMAREPORT_QR(6) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_eosinophils.Text)) > 0 Then
            UPDDBMAREPORT_QR(7) = Trim(txt_eosinophils.Text)
        Else
            UPDDBMAREPORT_QR(7) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_monocytes.Text)) > 0 Then
            UPDDBMAREPORT_QR(8) = Trim(txt_monocytes.Text)
        Else
            UPDDBMAREPORT_QR(8) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(TXT_PROMONOCYTES.Text)) > 0 Then
            UPDDBMAREPORT_QR(9) = Trim(TXT_PROMONOCYTES.Text)
        Else
            UPDDBMAREPORT_QR(9) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_matly.Text)) > 0 Then
            UPDDBMAREPORT_QR(10) = Trim(txt_matly.Text)
        Else
            UPDDBMAREPORT_QR(10) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_atyply.Text)) > 0 Then
            UPDDBMAREPORT_QR(11) = Trim(txt_atyply.Text)
        Else
            UPDDBMAREPORT_QR(11) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_plasmacells.Text)) > 0 Then
            UPDDBMAREPORT_QR(12) = Trim(txt_plasmacells.Text)
        Else
            UPDDBMAREPORT_QR(12) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_mature.Text)) > 0 Then
            UPDDBMAREPORT_QR(13) = Trim(txt_mature.Text)
        Else
            UPDDBMAREPORT_QR(13) = "  "
            NOTHINGWRITTEN = True
        End If
        If Len(Trim(txt_inmature.Text)) > 0 Then
            UPDDBMAREPORT_QR(14) = Trim(txt_inmature.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(14) = "  "
        End If
        If Len(Trim(txt_atypbi.Text)) > 0 Then
            UPDDBMAREPORT_QR(15) = Trim(txt_atypbi.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(15) = "  "
        End If
        If Len(Trim(txt_plasmablasts.Text)) > 0 Then
            UPDDBMAREPORT_QR(16) = Trim(txt_plasmablasts.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(16) = "  "
        End If
        If Len(Trim(txt_cellularity1.Text)) > 0 Then
            UPDDBMAREPORT_QR(17) = Null
            UPDDBMAREPORT_QR(17) = Trim(txt_cellularity1.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(17) = "  "
        End If
        If Len(Trim(txt_maturation.Text)) > 0 Then
            UPDDBMAREPORT_QR(18) = Null
            UPDDBMAREPORT_QR(18) = Trim(txt_maturation.Text)
        Else
        NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(18) = "  "
        End If
        If Len(Trim(txt_cellularity2.Text)) > 0 Then
            UPDDBMAREPORT_QR(19) = Null
            UPDDBMAREPORT_QR(19) = Trim(txt_cellularity2.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(19) = "  "
        End If
        If Len(Trim(txt_morphology.Text)) > 0 Then
            UPDDBMAREPORT_QR(20) = Null
            UPDDBMAREPORT_QR(20) = Trim(txt_morphology.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(20) = "  "
        End If
        If Len(Trim(TXT_IMPRINT.Text)) > 0 Then
            UPDDBMAREPORT_QR(21) = Null
            UPDDBMAREPORT_QR(21) = Trim(TXT_IMPRINT.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(21) = "  "
        End If
        If Len(Trim(txt_abcellblasts.Text)) > 0 Then
            UPDDBMAREPORT_QR(22) = Trim(txt_abcellblasts.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(22) = "  "
        End If
        If Len(Trim(CMB_MYELOPEROXIDASE.Text)) > 0 Then
            UPDDBMAREPORT_QR(23) = Trim(CMB_MYELOPEROXIDASE.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(23) = "  "
        End If
        If Len(Trim(txt_myeloperoxidase.Text)) > 0 Then
            UPDDBMAREPORT_QR(24) = Trim(txt_myeloperoxidase.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(24) = "  "
        End If
        If Len(Trim(CMB_NONSPEST.Text)) > 0 Then
           UPDDBMAREPORT_QR(25) = Trim(CMB_NONSPEST.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(25) = "  "
        End If
        If Len(Trim(txt_nonspecificesterase.Text)) > 0 Then
            UPDDBMAREPORT_QR(26) = Trim(txt_nonspecificesterase.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(26) = "  "
        End If
        If Len(Trim(CMB_TRAP.Text)) > 0 Then
            UPDDBMAREPORT_QR(27) = Trim(CMB_TRAP.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(27) = "  "
        End If
        If Len(Trim(txt_trapironstain.Text)) > 0 Then
            UPDDBMAREPORT_QR(28) = Trim(txt_trapironstain.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(28) = "  "
        End If
        If Len(Trim(CMB_IRONSTAIN.Text)) > 0 Then
            UPDDBMAREPORT_QR(29) = Trim(CMB_IRONSTAIN.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(29) = "  "
        End If
        If Len(Trim(TXT_IRONSTAIN.Text)) > 0 Then
            UPDDBMAREPORT_QR(30) = Trim(TXT_IRONSTAIN.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(30) = "  "
        End If
        If Len(Trim(CMB_GIDEROBLASTS.Text)) > 0 Then
            UPDDBMAREPORT_QR(31) = Trim(CMB_GIDEROBLASTS.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(31) = "  "
        End If
'        If Len(Trim(TXT_GIDEROBLASTS.Text)) > 0 Then
'        End If
        If Len(Trim(CMB_SUDANBLACKB.Text)) > 0 Then
            UPDDBMAREPORT_QR(32) = Trim(CMB_SUDANBLACKB.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(32) = "  "
        End If
        If Len(Trim(txt_sudanblackb.Text)) > 0 Then
            UPDDBMAREPORT_QR(33) = Trim(txt_sudanblackb.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(33) = "  "
        End If
        If Len(Trim(CMB_TOLUIDINEBLUE.Text)) > 0 Then
            UPDDBMAREPORT_QR(34) = Trim(CMB_TOLUIDINEBLUE.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(34) = "  "
        End If
        If Len(Trim(txt_TOLUIDINEBLUE.Text)) > 0 Then
            UPDDBMAREPORT_QR(35) = Trim(txt_TOLUIDINEBLUE.Text)
        Else
             NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(35) = "  "
        End If
        '**** SURFACE MARKERS****
        If Len(Trim(txt_cd13.Text)) > 0 Then
            UPDDBMAREPORT_QR(36) = Trim(txt_cd13.Text)
        Else
             NOTHINGWRITTEN = True
             UPDDBMAREPORT_QR(36) = "  "
        End If
        If Len(Trim(txt_cd14.Text)) > 0 Then
             UPDDBMAREPORT_QR(37) = Trim(txt_cd14.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(37) = "  "
        End If
        If Len(Trim(txt_CD15.Text)) > 0 Then
            UPDDBMAREPORT_QR(38) = Trim(txt_CD15.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(38) = "  "
        End If
        If Len(Trim(txt_CD33.Text)) > 0 Then
            UPDDBMAREPORT_QR(39) = Trim(txt_CD33.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(39) = "  "
        End If
        If Len(Trim(txt_cd36.Text)) > 0 Then
            UPDDBMAREPORT_QR(40) = Trim(txt_cd36.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(40) = "  "
        End If
        If Len(Trim(txt_cd41.Text)) > 0 Then
            UPDDBMAREPORT_QR(41) = Trim(txt_cd41.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(41) = "  "
        End If
        If Len(Trim(txt_Glycophorina.Text)) > 0 Then
            UPDDBMAREPORT_QR(42) = Trim(txt_Glycophorina.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(42) = "  "
        End If
        If Len(Trim(TXT_CD68.Text)) > 0 Then
            UPDDBMAREPORT_QR(43) = Trim(TXT_CD68.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(43) = "  "
        End If
        If Len(Trim(TXT_CD34.Text)) > 0 Then
            UPDDBMAREPORT_QR(44) = Trim(TXT_CD34.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(44) = "  "
        End If
        If Len(Trim(TXT_HLADR.Text)) > 0 Then
            UPDDBMAREPORT_QR(45) = Trim(TXT_HLADR.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(45) = "  "
        End If
        If Len(Trim(TXT_CD38.Text)) > 0 Then
            UPDDBMAREPORT_QR(46) = Trim(TXT_CD38.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(46) = "  "
        End If
        If Len(Trim(TXT_CD138.Text)) > 0 Then
            UPDDBMAREPORT_QR(47) = Trim(TXT_CD138.Text)
        Else
            NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(47) = "  "
        End If
        If Len(Trim(TXT_TDT.Text)) > 0 Then
            UPDDBMAREPORT_QR(48) = Trim(TXT_TDT.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(48) = "  "
        End If
        If Len(Trim(TXT_ANTIMPO.Text)) > 0 Then
            UPDDBMAREPORT_QR(49) = Trim(TXT_ANTIMPO.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(49) = "  "
        End If
        If Len(Trim(TXT_CD10.Text)) > 0 Then
            UPDDBMAREPORT_QR(50) = Trim(TXT_CD10.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(50) = "  "
        End If
        If Len(Trim(TXT_CD19.Text)) > 0 Then
            UPDDBMAREPORT_QR(51) = Trim(TXT_CD19.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(51) = "  "
        End If
        If Len(Trim(TXT_FMC7.Text)) > 0 Then
            UPDDBMAREPORT_QR(52) = Trim(TXT_FMC7.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(52) = "  "
        End If
        If Len(Trim(TXT_CD22.Text)) > 0 Then
            UPDDBMAREPORT_QR(53) = Trim(TXT_CD22.Text)
        Else
            NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(53) = "  "
        End If
        If Len(Trim(TXT_CD23.Text)) > 0 Then
            UPDDBMAREPORT_QR(54) = Trim(TXT_CD23.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(54) = "  "
        End If
        If Len(Trim(txt_smig.Text)) > 0 Then
            UPDDBMAREPORT_QR(55) = Trim(txt_smig.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(55) = "  "
        End If
        
        '***** Shraddha 09/03/10 *****
        UPD_DBMAREPORT_QR(0) = Trim(TXT_SMPCTG.Text) 'SMPCTG
        UPD_DBMAREPORT_QR(1) = Trim(txt_tcrab.Text) 'TCRAB
        UPD_DBMAREPORT_QR(2) = Trim(txt_tcrag.Text) 'TCRAG
        UPD_DBMAREPORT_QR(3) = Trim(txt_cyclinD1.Text) 'CLDN1
        UPD_DBMAREPORT_QR(4) = req_chglen(cmb_reqno.Text) 'REQNO
        UPD_DBMAREPORT_QR.Execute
        
        
        'ADDED BY SHWETA ON 19-07-2013 AS PER REQUEST UZZ DEPT.
                    Dim RS_ERYTHCELL As rdoResultset
                    Dim ERYTHSTR As String
                    ERYTHSTR = "SELECT * FROM DISLIB.DBMASUBNEW WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "'"
                    Set RS_ERYTHCELL = rdb.OpenResultset(ERYTHSTR, 1, 1)
                    If RS_ERYTHCELL.EOF = False Then
                     
                    erythcells = "UPDATE DISLIB.DBMASUBNEW  SET  ERYTHROCT='" & Trim(TXT_ErythCells.Text) & "',FLD1='U',SAVE_DT=CURRENT DATE , SAVE_TM=CURRENT TIME WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "'"
                    rdb.Execute erythcells
                    
                    Else
                       
                    erythcells = "insert into DISLIB.DBMASUBNEW(REQNO,ERYTHROCT,SAVE_DT,SAVE_TM,FLD1) VALUES('" & Trim(req_chglen(cmb_reqno.Text)) & "','" & Trim(TXT_ErythCells.Text) & "',CURRENT DATE,CURRENT TIME,'IU')"
                    rdb.Execute erythcells
                       
                    End If
        
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
        '***** Added on 10-02-2014 by SHWETA
        UPDDCBCREPA_QR(0) = Trim(TXT_DNAINDEX.Text)
        UPDDCBCREPA_QR(1) = Trim(TXT_DNACLOIDY.Text)
        UPDDCBCREPA_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDCBCREPA_QR.Execute
        ''''till here
        
        
        
        '***** Shraddha 12/04/10 *****
        If opt_lsp.Value = True Then
            updcbcsm_qr(0) = "Y"
        Else
            updcbcsm_qr(0) = "N"
        End If
            
        If opt_ssp.Value = True Then
            updcbcsm_qr(1) = "Y"
        Else
            updcbcsm_qr(1) = "N"
        End If
                  
        updcbcsm_qr(2) = req_chglen(cmb_reqno.Text)
        
        updcbcsm_qr.Execute
        
        If opt_drak.Value = True Then
            UPDDREQDRNAME_QR(0) = "Dr. Ashok Kumar"
        ElseIf opt_drak.Value = False Then
            UPDDREQDRNAME_QR(0) = ""
        End If
'        If opt_mryb.Value = True Then
'            UPDDREQDRNAME_QR(1) = "Mr. Y Badrinath"
'        ElseIf opt_mryb.Value = False Then
'            UPDDREQDRNAME_QR(1) = ""
'        End If
        '***** Commented and added by Shraddha on 05/05/10 *****
        If opt_mryb.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Y. Badrinath"
        ElseIf opt_mryb.Value = False And opt_sitaram.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Sitaram"
        Else
            UPDDREQDRNAME_QR(1) = ""
        End If
        '*****************************************
        UPDDREQDRNAME_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDREQDRNAME_QR.Execute
        '*****************************
        
        If Len(Trim(txt_CD79B.Text)) > 0 Then
            UPDDBMAREPORT_QR(56) = Trim(txt_CD79B.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(56) = "  "
        End If
        If Len(Trim(TXT_CD25.Text)) > 0 Then
            UPDDBMAREPORT_QR(57) = Trim(TXT_CD25.Text)
        Else
             NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(57) = "  "
        End If
        If Len(Trim(TXT_CD11C.Text)) > 0 Then
            UPDDBMAREPORT_QR(58) = Trim(TXT_CD11C.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(58) = "  "
        End If
        If Len(Trim(txt_cd103.Text)) > 0 Then
            UPDDBMAREPORT_QR(59) = Trim(txt_cd103.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(59) = "  "
        End If
        If Len(Trim(txt_kappa.Text)) > 0 Then
            UPDDBMAREPORT_QR(60) = Trim(txt_kappa.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(60) = "  "
        End If
        If Len(Trim(txt_lanbda.Text)) > 0 Then
            UPDDBMAREPORT_QR(61) = Trim(txt_lanbda.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(61) = "  "
        End If
        If Len(Trim(TXT_DBA44.Text)) > 0 Then
            UPDDBMAREPORT_QR(62) = Trim(TXT_DBA44.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(62) = "  "
        End If
        If Len(Trim(txt_cytocd221(1))) > 0 Then
            UPDDBMAREPORT_QR(63) = Trim(txt_cytocd221(1))
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(63) = "  "
        End If
        If Len(Trim(txt_cytocd79a.Text)) > 0 Then
            UPDDBMAREPORT_QR(64) = Trim(txt_cytocd79a.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(64) = "  "
        End If
        If Len(Trim(TXT_CYTOMU.Text)) > 0 Then
            UPDDBMAREPORT_QR(65) = Trim(TXT_CYTOMU.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(65) = "  "
        End If
        If Len(Trim(TXT_CD2.Text)) > 0 Then
            UPDDBMAREPORT_QR(66) = Trim(TXT_CD2.Text)
        Else
             NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(66) = "  "
        End If
        If Len(Trim(TXT_CD3.Text)) > 0 Then
            UPDDBMAREPORT_QR(67) = Trim(TXT_CD3.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(67) = "  "
        End If
        If Len(Trim(TXT_CD4.Text)) > 0 Then
            UPDDBMAREPORT_QR(68) = Trim(TXT_CD4.Text)
        Else
             NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(68) = "  "
        End If
        If Len(Trim(TXT_CD5.Text)) > 0 Then
            UPDDBMAREPORT_QR(69) = Trim(TXT_CD5.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(69) = "  "
        End If
        If Len(Trim(TXT_CD7.Text)) > 0 Then
            UPDDBMAREPORT_QR(70) = Trim(TXT_CD7.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(70) = "  "
        End If
        If Len(Trim(TXT_CD8.Text)) > 0 Then
            UPDDBMAREPORT_QR(71) = Trim(TXT_CD8.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(71) = "  "
        End If
        If Len(Trim(TXT_CYTOCD3.Text)) > 0 Then
            UPDDBMAREPORT_QR(72) = Trim(TXT_CYTOCD3.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(72) = "  "
        End If
        If Len(Trim(txt_IMPRESSION.Text)) > 0 Then
            'UPDDBMAREPORT_QR(73) = Null
            UPDDBMAREPORT_QR(73) = Trim(txt_IMPRESSION.Text)
        Else
            'NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(73) = "  "
        End If
        If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
            UPDDBMAREPORT_QR(74) = Null
            UPDDBMAREPORT_QR(74) = Trim(TXT_COMMENTS.Text)
        Else
            NOTHINGWRITTEN = True
            UPDDBMAREPORT_QR(74) = "  "
        End If
        If Len(Trim(TXT_GIDEROBLASTS.Text)) > 0 Then
            UPDDBMAREPORT_QR(75) = Trim(TXT_GIDEROBLASTS.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(75) = "  "
        End If
        If Len(Trim(txt_igm.Text)) > 0 Then  ' fld2 used for igm
        
          UPDDBMAREPORT_QR(76) = Trim(txt_igm.Text)
        Else
           NOTHINGWRITTEN = True
        
          UPDDBMAREPORT_QR(76) = "  "
        End If
        
        If Len(Trim(TXT_CD42.Text)) > 0 Then
            UPDDBMAREPORT_QR(77) = Trim(TXT_CD42.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(77) = "  "
        End If
        If Len(Trim(TXT_CD117.Text)) > 0 Then
            UPDDBMAREPORT_QR(78) = Trim(TXT_CD117.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(78) = "  "
        End If
        If Len(Trim(TXT_CD20.Text)) > 0 Then
            UPDDBMAREPORT_QR(79) = Trim(TXT_CD20.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(79) = "  "
        End If
        If Len(Trim(TXT_CD61.Text)) > 0 Then
            UPDDBMAREPORT_QR(80) = Trim(TXT_CD61.Text)
        Else
            NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(80) = "  "
        End If
        If Len(Trim(TXT_CD56.Text)) > 0 Then
            UPDDBMAREPORT_QR(81) = Trim(TXT_CD56.Text)
        Else
           NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(81) = "  "
        End If
        If Len(Trim(txt_descmorpho.Text)) > 0 Then
            UPDDBMAREPORT_QR(82) = Trim(txt_descmorpho.Text)
        Else
            NOTHINGWRITTEN = True
           UPDDBMAREPORT_QR(82) = "  "
        End If
        
        '***** Shraddha 09/03/10 *****
        'TCRAB=?,TCRAG=? WHERE REQNO=?"
        UPDDBMAREPORT_QR(83) = txt_tcrab.Text 'TCRAB""
        UPDDBMAREPORT_QR(84) = txt_tcrag.Text 'TCRAG""
        '*****************************
        
        UPDDBMAREPORT_QR(85) = req_chglen(cmb_reqno.Text) 'REQNO
        
        UPDDBMAREPORT_QR.Execute
        '' updating dbmasubrep
            If Len(Trim(Text3.Text)) > 0 Then
            
              UPDDBMASUB_QR(1) = Trim(Text3.Text)
            Else
               NOTHINGWRITTEN = True
               UPDDBMASUB_QR(1) = "  "
            End If
            
            If Len(Trim(Text2.Text)) > 0 Then
            
              UPDDBMASUB_QR(0) = Trim(Text2.Text)
            Else
                NOTHINGWRITTEN = True
               UPDDBMASUB_QR(0) = "  "
            End If
            
            If Len(Trim(Text1.Text)) > 0 Then
            
              UPDDBMASUB_QR(2) = Trim(Text1.Text)
            Else
                NOTHINGWRITTEN = True
               UPDDBMASUB_QR(2) = "  "
            End If
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
               UPDDBMASUB_QR(3) = Trim(TXT_COMMENTS.Text)
            Else
               NOTHINGWRITTEN = True
               UPDDBMASUB_QR(3) = "  "
            End If
            
            If Check1.Value = 1 Then
               UPDDBMASUB_QR(4) = "Y"
            ElseIf Check1.Value = 0 Then
               UPDDBMASUB_QR(4) = "N"
            End If
            
            UPDDBMASUB_QR(5) = Trim(req_chglen(cmb_reqno.Text))
            UPDDBMASUB_QR.Execute
            
            
            
            
            '*** updating the DHBFREPORT
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = regcbo.Text  ' cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
           
    End If
                    
    ''******ADDENDUM*******************************
    If ADDENDUMFALG = True Then
        SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
        SELDADDENDUM_RS.MoreResults
        SELDADDENDUM_RS.Requery
        If SELDADDENDUM_RS.RowCount > 0 Then
            SELDADDENDUM_RS.MoveLast
            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INSDADDENDUM_QR(3) = time1
            INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
            INSDADDENDUM_QR.Execute
        Else
            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            INSDADDENDUM_QR(1) = "1"
            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INSDADDENDUM_QR(3) = time1
            INSDADDENDUM_QR(4) = cmb_addendumby.Text
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
            INSDADDENDUM_QR.Execute
        End If
    End If
    ''******ADDENDUM*******************************
            
'   If Check1.Value = 1 And NOTHINGWRITTEN = True Then
'     MsgBox "Nothing is Written in report, But FullReport is Selected"
'      rdb.RollbackTrans
'     Exit Sub
'   End If
   
    
rdb.CommitTrans
MsgBox "Record is Saved"
btnenabl ("INI")
ENABL ("DIS")
Exit Sub
'MSG:
    'rdb.RollbackTrans
'MsgBox Err.Description
End Sub




Public Sub ENABL(a As String)
    If a = "DIS" Then
        txt_cellularity.Enabled = False
        TXT_MERATION.Enabled = False
        TXT_MERATION1.Enabled = False
        TXT_PROMONOCYTES.Enabled = False
        txt_promyelocytes.Enabled = False
        txt_myelocytes.Enabled = False
        txt_metamyelocytes.Enabled = False
        txt_polymorphs.Enabled = False
        txt_basophils.Enabled = False
        txt_eosinophils.Enabled = False
        txt_monocytes.Enabled = False
        TXT_PROMONOCYTES.Enabled = False
        txt_matly.Enabled = False
        txt_atyply.Enabled = False
        txt_plasmacells.Enabled = False
        txt_mature.Enabled = False
        txt_inmature.Enabled = False
        txt_atypbi.Enabled = False
        txt_plasmablasts.Enabled = False
        txt_cellularity1.Enabled = False
        txt_maturation.Enabled = False
        TXT_ErythCells.Enabled = False
        txt_cellularity2.Enabled = False
        txt_morphology.Enabled = False
        CMB_MYELOPEROXIDASE.Enabled = False
        'txt_myeloperoxidase.Enabled = False
        CMB_NONSPEST.Enabled = False
        'txt_nonspecificesterase.Enabled = False
        CMB_TRAP.Enabled = False
        'txt_trapironstain.Enabled = False
        CMB_IRONSTAIN.Enabled = False
        CMB_GIDEROBLASTS.Enabled = False
        CMB_SUDANBLACKB.Enabled = False
        CMB_TOLUIDINEBLUE.Enabled = False
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
        TXT_IMPRINT.Enabled = False
        txt_descmorpho.Enabled = False
        txt_IMPRESSION.Enabled = False
        txt_abcellblasts.Enabled = False
        TXT_COMMENTS.Enabled = False
        
        txt_myeloperoxidase.Enabled = False
        txt_nonspecificesterase.Enabled = False
        txt_trapironstain.Enabled = False
        TXT_IRONSTAIN.Enabled = False
        TXT_GIDEROBLASTS.Enabled = False
        txt_sudanblackb.Enabled = False
        txt_TOLUIDINEBLUE.Enabled = False
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        cmb_enterby.Enabled = False
        'cmb_finalby.Enabled = False
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
    If a = "ENA" Then
        txt_myeloperoxidase.Enabled = True
        txt_nonspecificesterase.Enabled = True
        txt_trapironstain.Enabled = True
        TXT_IRONSTAIN.Enabled = True
        TXT_GIDEROBLASTS.Enabled = True
        txt_sudanblackb.Enabled = True
        txt_TOLUIDINEBLUE.Enabled = True
        Text1.Enabled = True 'False
        Text2.Enabled = True 'False
        Text3.Enabled = True 'False
        txt_cellularity.Enabled = True
        TXT_MERATION.Enabled = True
        TXT_MERATION1.Enabled = True
        TXT_PROMONOCYTES.Enabled = True
        txt_promyelocytes.Enabled = True
        txt_myelocytes.Enabled = True
        txt_metamyelocytes.Enabled = True
        txt_polymorphs.Enabled = True
        txt_basophils.Enabled = True
        txt_eosinophils.Enabled = True
        txt_monocytes.Enabled = True
        TXT_PROMONOCYTES.Enabled = True
        txt_matly.Enabled = True
        txt_atyply.Enabled = True
        txt_plasmacells.Enabled = True
        txt_mature.Enabled = True
        txt_inmature.Enabled = True
        txt_atypbi.Enabled = True
        txt_plasmablasts.Enabled = True
        txt_cellularity1.Enabled = True
        txt_maturation.Enabled = True
        TXT_ErythCells.Enabled = True
        txt_cellularity2.Enabled = True
        txt_morphology.Enabled = True
        CMB_MYELOPEROXIDASE.Enabled = True
        'txt_myeloperoxidase.Enabled = False
        CMB_NONSPEST.Enabled = True
        'txt_nonspecificesterase.Enabled = False
        CMB_TRAP.Enabled = True
        'txt_trapironstain.Enabled = False
        CMB_IRONSTAIN.Enabled = True
        CMB_GIDEROBLASTS.Enabled = True
        CMB_SUDANBLACKB.Enabled = True
        CMB_TOLUIDINEBLUE.Enabled = True
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
        TXT_IMPRINT.Enabled = True
        txt_descmorpho.Enabled = True
        txt_IMPRESSION.Enabled = True
        txt_abcellblasts.Enabled = True
        TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        regcbo.Enabled = True
        
    End If
If a = "FENA" Then
          txt_myeloperoxidase.Enabled = True
        txt_nonspecificesterase.Enabled = True
        txt_trapironstain.Enabled = True
        TXT_IRONSTAIN.Enabled = True
        TXT_GIDEROBLASTS.Enabled = True
        txt_sudanblackb.Enabled = True
        txt_TOLUIDINEBLUE.Enabled = True
        Text1.Enabled = True 'False
        Text2.Enabled = True 'False
        Text3.Enabled = True 'False
        txt_cellularity.Enabled = True
        TXT_MERATION.Enabled = True
        TXT_MERATION1.Enabled = True
        TXT_PROMONOCYTES.Enabled = True
        txt_promyelocytes.Enabled = True
        txt_myelocytes.Enabled = True
        txt_metamyelocytes.Enabled = True
        txt_polymorphs.Enabled = True
        txt_basophils.Enabled = True
        txt_eosinophils.Enabled = True
        txt_monocytes.Enabled = True
        TXT_PROMONOCYTES.Enabled = True
        txt_matly.Enabled = True
        txt_atyply.Enabled = True
        txt_plasmacells.Enabled = True
        txt_mature.Enabled = True
        txt_inmature.Enabled = True
        txt_atypbi.Enabled = True
        txt_plasmablasts.Enabled = True
        txt_cellularity1.Enabled = True
        txt_maturation.Enabled = True
        TXT_ErythCells.Enabled = True
        txt_cellularity2.Enabled = True
        txt_morphology.Enabled = True
        CMB_MYELOPEROXIDASE.Enabled = True
        'txt_myeloperoxidase.Enabled = False
        CMB_NONSPEST.Enabled = True
        'txt_nonspecificesterase.Enabled = False
        CMB_TRAP.Enabled = True
        'txt_trapironstain.Enabled = False
        CMB_IRONSTAIN.Enabled = True
        CMB_GIDEROBLASTS.Enabled = True
        CMB_SUDANBLACKB.Enabled = True
        CMB_TOLUIDINEBLUE.Enabled = True
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
        TXT_IMPRINT.Enabled = True
        txt_descmorpho.Enabled = True
        txt_IMPRESSION.Enabled = True
        txt_abcellblasts.Enabled = True
        TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        regcbo.Enabled = True
        
        Check1.Enabled = False
End If
If a = "IENA" Then
  txt_IMPRESSION.Enabled = True
  TXT_COMMENTS.Enabled = True
  Check1.Enabled = True
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



Private Sub cmd_back_Click()



btnenabl ("INI")
field_clear
TCLEAR


'btnenabl ("INI")
'BTNENABL1 ("INI")
'FLDENABL ("DIS")
'FLDENABL1 ("DIS")
'TXT_CLEAR
'field_clear

If WORKLIST = True Then
    'Unload Me
    Me.Hide
    disw.show
Else
    'Unload Me
    Me.Hide
    DIS_MAIN.show
End If


'Me.Hide
'DIS_MAIN.show
End Sub

Public Sub TCLEAR()
        txt_cellularity.Text = ""
        TXT_MERATION.Text = ""
        TXT_MERATION1.Text = ""
        TXT_PROMONOCYTES.Text = ""
        txt_myelocytes.Text = ""
        txt_metamyelocytes.Text = ""
        txt_polymorphs.Text = ""
        txt_basophils.Text = ""
        txt_eosinophils.Text = ""
        txt_monocytes.Text = ""
        TXT_PROMONOCYTES.Text = ""
        txt_matly.Text = ""
        txt_atyply.Text = ""
        txt_plasmacells.Text = ""
        txt_mature.Text = ""
        txt_inmature.Text = ""
        txt_atypbi.Text = ""
        txt_plasmablasts.Text = ""
        txt_cellularity1.Text = ""
        txt_maturation.Text = ""
        txt_cellularity2.Text = ""
        txt_morphology.Text = ""
        CMB_MYELOPEROXIDASE.Text = ""
        'txt_myeloperoxidase.Text =""
        CMB_NONSPEST.Text = ""
        'txt_nonspecificesterase.Text =""
        CMB_TRAP.Text = ""
        'txt_trapironstain.Enabled = False
        CMB_IRONSTAIN.Text = ""
        CMB_GIDEROBLASTS.Text = ""
        CMB_SUDANBLACKB.Text = ""
        CMB_TOLUIDINEBLUE.Text = ""
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
        TXT_SMPCTG.Text = "" 'Shraddha 09/03/10
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
        Text1.Text = " "
        Text2.Text = " "
        Text3.Text = " "
        txt_igm = " "
        TXT_PATH = " "
        txt_modcode = " "
        txt_nmaterial = " "
        
        TXT_CYTOCD3.Text = ""
End Sub






'***** Shraddha 09/03/10 *****
Private Sub Timer1_Timer()
'datetime
'txt_time.Text = time1
End Sub
'*****************************

Private Sub txt_addendum_GotFocus()
On Error GoTo m3:
W_ADDENDUM.Visible = False

Set d_ADDENDUM = W_ADDENDUM.ActiveDocument
d_ADDENDUM.Content = ""
d_ADDENDUM.Content = txt_addendum
W_ADDENDUM.Caption = "ADDENDUM"
'W_grossdes.WindowState = wdWindowStateMaximize
W_ADDENDUM.WindowState = wdWindowStateNormal
d_ADDENDUM.Content = txt_addendum
W_ADDENDUM.Visible = True
W_ADDENDUM.Activate
m3:
On Error Resume Next:
Set W_ADDENDUM = CreateObject("word.Application")
W_ADDENDUM.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_ADDENDUM.Caption = "ADDENDUM"
W_ADDENDUM.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_ADDENDUM = W_ADDENDUM.ActiveDocument
d_ADDENDUM.Content = txt_addendum.Text
W_ADDENDUM.Visible = True
W_ADDENDUM.Activate

End Sub

Private Sub txt_addendum_LostFocus()

'txt_addendum.Text = d_grossdesc.Content
On Error GoTo m4:
Set d_ADDENDUM = W_ADDENDUM.ActiveDocument
txt_addendum.Text = d_ADDENDUM.Content
'Set d_grossdesc = W_grossdes.ActiveDocument

If blank = True Then
    l = Len(d_ADDENDUM.Content)
    If l <= 4 Then
        txt_addendum.Text = ""
     Else
       txt_addendum.Text = d_ADDENDUM.Content
    End If
End If
d_ADDENDUM.Content = ""
d_ADDENDUM.Close (wdDoNotSaveChanges)
W_ADDENDUM.quit
Exit Sub
m4:
MsgBox Err.Description

'On Error GoTo m4:
'Set d_ADDENDUM = W_ADDENDUM.ActiveDocument
'txt_addendum.Text = d_ADDENDUM.Content
''flagimp = False
'd_ADDENDUM.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_ADDENDUM.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_ADDENDUM.Close (wdDoNotSaveChanges)
'W_ADDENDUM.quit
''W1.Visible = False
'Exit Sub
'm4:
'MsgBox Err.Description

End Sub

Private Sub TXT_ANTIMPO_LostFocus()
txt_IMPRESSION.SetFocus
End Sub

Private Sub txt_cellularity_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_cellularity
W_grossdes.Caption = "Cellularity"
'W_grossdes.WindowState = wdWindowStateMaximize
W_grossdes.WindowState = wdWindowStateNormal
d_grossdesc.Content = txt_cellularity
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Cellularity"
W_grossdes.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_cellularity.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub

Private Sub txt_cellularity_LostFocus()
'On Error GoTo m4:
'Set d_grossdesc = W_grossdes.ActiveDocument
'If blank = True Then
'    l = Len(d_grossdesc.Content)
'    If l <= 4 Then
'        txt_cellularity.Text = ""
'    Else
'       txt_cellularity.Text = d_grossdesc.Content
'      End If
'End If
'd_grossdesc.Content = ""
'd_grossdesc.Close (wdDoNotSaveChanges)
'W_grossdes.quit
'Exit Sub
'm4:
'MsgBox Err.Description

On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument
txt_cellularity.Text = d_grossdesc.Content
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

Private Sub txt_cellularity2_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_cellularity2
W_grossdes.Caption = "Cellularity"
'W_grossdes.WindowState = wdWindowStateMaximize
W_grossdes.WindowState = wdWindowStateNormal
d_grossdesc.Content = txt_cellularity2
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Cellularity"
W_grossdes.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_cellularity2.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub

Private Sub txt_cellularity2_LostFocus()
txt_cellularity2.Text = d_grossdesc.Content
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument

If blank = True Then
    l = Len(d_grossdesc.Content)
    If l <= 4 Then
        txt_cellularity2.Text = ""
    Else
       txt_cellularity2.Text = d_grossdesc.Content
     End If
End If
d_grossdesc.Content = ""
d_grossdesc.Close (wdDoNotSaveChanges)
W_grossdes.quit
Exit Sub
m4:
MsgBox Err.Description

'On Error GoTo m4:
'Set d_grossdesc = W_grossdes.ActiveDocument
'txt_cellularity2.Text = d_grossdesc.Content
''flagimp = False
'd_grossdesc.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_grossdesc.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_grossdesc.Close (wdDoNotSaveChanges)
'W_grossdes.quit
''W1.Visible = False
'Exit Sub
'm4:
'MsgBox Err.Description
End Sub


Private Sub TXT_COMMENTS_GotFocus()
On Error GoTo m31:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = TXT_COMMENTS.Text
W_grossdes.Caption = "Comments"
'W_grossdes.WindowState = wdWindowStateMaximize
W_grossdes.WindowState = wdWindowStateNormal
d_grossdesc.Content = TXT_COMMENTS
W_grossdes.Visible = True
W_grossdes.Activate
m31:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Comments"
W_grossdes.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = TXT_COMMENTS.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub

Private Sub txt_comments_lostfocus()

On Error GoTo m4:
'Added by Milind on 02/02/06 for removing Comments from Report when it is blank
TXT_COMMENTS.Text = d_grossdesc.Content
Set d_grossdesc = W_grossdes.ActiveDocument

If blank = True Then
    l = Len(d_grossdesc.Content)
    If l <= 4 Then
        TXT_COMMENTS.Text = ""
    Else
        TXT_COMMENTS.Text = d_grossdesc.Content
     End If
End If

'upto here

d_grossdesc.Content = ""
d_grossdesc.Close (wdDoNotSaveChanges)
W_grossdes.quit
Exit Sub
m4:
MsgBox Err.Description
End Sub
'On Error GoTo m41:
'Set d_grossdesc = W_grossdes.ActiveDocument
'TXT_COMMENTS.Text = d_grossdesc.Content
''flagimp = False
'd_grossdesc.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_grossdesc.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_grossdesc.Close (wdDoNotSaveChanges)
'W_grossdes.quit
''W1.Visible = False
'Exit Sub
'm41:
'MsgBox Err.Description
'End Sub

Private Sub txt_descmorpho_GotFocus()
On Error GoTo m3:
W_descmorpho.Visible = False
Set d_descmorpho = W_descmorpho.ActiveDocument
'd_descmorpho.Content = ""
d_descmorpho.Content = txt_descmorpho.Text
W_descmorpho.Caption = "Descriptive Morphology"
'W_grossdes.WindowState = wdWindowStateMaximize
W_descmorpho.WindowState = wdWindowStateNormal
d_descmorpho.Content = txt_descmorpho.Text
W_descmorpho.Visible = True
W_descmorpho.Activate
m3:
On Error Resume Next:
Set W_descmorpho = CreateObject("word.Application")
W_descmorpho.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_descmorpho.Caption = "Descriptive Morphology"
W_descmorpho.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_descmorpho = W_descmorpho.ActiveDocument
d_descmorpho.Content = txt_descmorpho.Text
W_descmorpho.Visible = True
W_descmorpho.Activate

'**********************************************




End Sub

Private Sub txt_descmorpho_LostFocus()
'ReDim X(10)
txt_descmorpho.Text = d_descmorpho.Content
On Error GoTo m4:
Set d_descmorhpo = W_descmorpho.ActiveDocument

'Added by Milind on 02/02/06 for removing Description Morphology from Report when it is blank
If blank = True Then
    l = Len(d_descmorpho.Content)
    If l <= 4 Then
'    For J = 0 To 9
'         X(J) = Mid(Trim(txt_descmorpho.Text), J + 1, 1)
'         MsgBox Asc(X(J))
'    Next
    
        txt_descmorpho.Text = ""
    Else
        txt_descmorpho.Text = d_descmorpho.Content
    End If
End If
'upto here

d_descmorpho.Content = ""
d_descmorpho.Close (wdDoNotSaveChanges)
W_descmorpho.quit
Call TXT_IMPRINT_GotFocus
Exit Sub
m4:
MsgBox Err.Description
'On Error GoTo m4:
'Set d_descmorpho = W_descmorpho.ActiveDocument
'txt_descmorpho.Text = d_descmorpho.Content
''flagimp = False
'd_descmorpho.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_descmorpho.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_descmorpho.Close (wdDoNotSaveChanges)
'W_descmorpho.quit
''W1.Visible = False
'Exit Sub
'm4:
'MsgBox Err.Description

End Sub

Private Sub txt_impression_gotfocus()
On Error GoTo m3:
'txt_impression.Visible = True
'txt_impression.Enabled = True
W_grossdes.Visible = False

'W_grossdes.Visible = True
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_IMPRESSION
W_grossdes.Caption = "IMPRESSION"
'W_grossdes.WindowState = wdWindowStateMaximize
W_grossdes.WindowState = wdWindowStateNormal
d_grossdesc.Content = txt_IMPRESSION
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "IMPRESSION"
W_grossdes.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_IMPRESSION.Text
W_grossdes.Visible = True
testW_grossdes.Activate
End Sub

Private Sub txt_impression_lostfocus()
txt_IMPRESSION.Text = d_grossdesc.Content
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument

If blank = True Then
    l = Len(d_grossdesc.Content)
    If l <= 4 Then
        txt_IMPRESSION.Text = ""
    Else
        txt_IMPRESSION.Text = d_grossdesc.Content
    End If
End If
d_grossdesc.Content = ""
d_grossdesc.Close (wdDoNotSaveChanges)
W_grossdes.quit
Exit Sub
m4:
MsgBox Err.Description

'On Error GoTo m4:
'Set d_grossdesc = W_grossdes.ActiveDocument
'txt_impression.Text = d_grossdesc.Content
''flagimp = False
'd_grossdesc.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_grossdesc.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_grossdesc.Close (wdDoNotSaveChanges)
'W_grossdes.quit
''W1.Visible = False
'Exit Sub
'm4:
'MsgBox Err.Description
End Sub

Private Sub TXT_IMPRINT_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = TXT_IMPRINT
W_grossdes.Caption = "IMPrint"
'W_grossdes.WindowState = wdWindowStateMaximize
W_grossdes.WindowState = wdWindowStateNormal
d_grossdesc.Content = TXT_IMPRINT
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "IMPrint"
W_grossdes.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = TXT_IMPRINT.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub
Private Sub TXT_IMPRINT_LostFocus()
TXT_IMPRINT.Text = d_grossdesc.Content
On Error GoTo m4:

Set d_grossdesc = W_grossdes.ActiveDocument

'Added by Milind on 02/02/06 for removing Description Morphology from Report when it is blank
If blank = True Then
    l = Len(d_grossdesc.Content)
    If l <= 4 Then
        TXT_IMPRINT.Text = ""
    Else
        TXT_IMPRINT.Text = d_grossdesc.Content
    End If
End If
'upto here
d_grossdesc.Content = ""
d_grossdesc.Close (wdDoNotSaveChanges)
W_grossdes.quit
Exit Sub

m4:
MsgBox Err.Description
'On Error GoTo m4:
'Set d_grossdesc = W_grossdes.ActiveDocument
'TXT_IMPRINT.Text = d_grossdesc.Content
''flagimp = False
'd_grossdesc.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_grossdesc.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_grossdesc.Close (wdDoNotSaveChanges)
'W_grossdes.quit
''W1.Visible = False
'Exit Sub
'm4:
'MsgBox Err.Description
End Sub

Private Sub txt_morphology_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_morphology
W_grossdes.Caption = "Morphology"
'W_grossdes.WindowState = wdWindowStateMaximize
W_grossdes.WindowState = wdWindowStateNormal
d_grossdesc.Content = txt_morphology
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Morphology"
W_grossdes.WindowState = wdWindowStateNormal
'W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_morphology.Text
W_grossdes.Visible = True
W_grossdes.Activate

'*****************************

End Sub

Private Sub txt_morphology_LostFocus()
'txt_descmorpho.Text = d_grossdesc.Content
'***** Commented andadded by Shraddha on 23/06/10 *****
txt_morphology.Text = d_grossdesc.Content
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument
'If blank = True Then
'    l = Len(d_grossdesc.Content)
'    If l <= 4 Then
'        txt_morphology.Text = ""
'    Else
'        txt_descmorpho.Text = d_grossdesc.Content
'    End If
'End If

d_grossdesc.Content = ""
d_grossdesc.Close (wdDoNotSaveChanges)
W_grossdes.quit
Call txt_descmorpho_GotFocus
Exit Sub
m4:
    MsgBox Err.Description
'On Error GoTo m4:
'Set d_grossdesc = W_grossdes.ActiveDocument
'TXT_MORPHOLOGY.Text = d_grossdesc.Content
''flagimp = False
'd_grossdesc.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_grossdesc.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_grossdesc.Close (wdDoNotSaveChanges)
'W_grossdes.quit
''W1.Visible = False
'Exit Sub
'm4:
'MsgBox Err.Description
End Sub

