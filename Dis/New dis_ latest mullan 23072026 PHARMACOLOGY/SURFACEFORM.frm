VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SURFACEFORM 
   Caption         =   "SURFACEMARKERFORM HEMATOLOGY"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   LinkTopic       =   "Form2"
   ScaleHeight     =   7485
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "BACK TO MAIN FORM"
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
      Left            =   8280
      TabIndex        =   83
      Top             =   5760
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5100
      Left            =   240
      TabIndex        =   0
      Top             =   525
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8996
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Screen I"
      TabPicture(0)   =   "SURFACEFORM.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LBL_CD2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LBL_CD3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LBL_CD4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LBL_CYTOCD3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LBL_CD5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LBL_CD7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LBL_CD8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LBL_CD10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_cd15"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_cd14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LBL_CD13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LBL_CD11C"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LBL_CD36"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "LBL_CD19"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "LBL_CD22"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl_cyto221"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "LBL_CD23"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "LBL_CD25"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LBL_CD33"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label20"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Shape3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label21"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Shape4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label18"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Shape5"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TXT_CD2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TXT_CD3"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TXT_CD4"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TXT_CYTOCD3"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TXT_CD5"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TXT_CD7"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TXT_CD8"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TXT_CD10"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt_CD15"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txt_cd14"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txt_cd13"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TXT_CD11C"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txt_cd36"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TXT_CD19"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TXT_CD22"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TXT_CD20"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt_cytocd221(1)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TXT_CD23"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "TXT_CD25"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt_CD33"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "opt_PrePBSC_CD3"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "opt_PBSCharvest_CD3"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "opt_BMharvest_CD3"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "opt_Tcell_CD4"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "opt_Tcell_CD8"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "Screen II"
      TabPicture(1)   =   "SURFACEFORM.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LBL_CD41"
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(2)=   "LBL_CD38"
      Tab(1).Control(3)=   "lbl_cd34"
      Tab(1).Control(4)=   "lbl_cytocd79a"
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(6)=   "LBL_CD68"
      Tab(1).Control(7)=   "Label12"
      Tab(1).Control(8)=   "lbl_cd103"
      Tab(1).Control(9)=   "LBL_CD79B"
      Tab(1).Control(10)=   "Label13"
      Tab(1).Control(11)=   "LBL_CD138"
      Tab(1).Control(12)=   "LBL_CYTOMU"
      Tab(1).Control(13)=   "lbl_DBA44"
      Tab(1).Control(14)=   "LBL_FMC7"
      Tab(1).Control(15)=   "LBL_ANTIMPO"
      Tab(1).Control(16)=   "LBL_GLYCOPHORIN"
      Tab(1).Control(17)=   "LBL_HALDR"
      Tab(1).Control(18)=   "lbl_kappa"
      Tab(1).Control(19)=   "lbl_lanbda"
      Tab(1).Control(20)=   "Shape2"
      Tab(1).Control(21)=   "Label19"
      Tab(1).Control(22)=   "txt_cd41"
      Tab(1).Control(23)=   "TXT_CD42"
      Tab(1).Control(24)=   "TXT_CD38"
      Tab(1).Control(25)=   "TXT_CD34"
      Tab(1).Control(26)=   "txt_cytocd79a"
      Tab(1).Control(27)=   "TXT_CD56"
      Tab(1).Control(28)=   "TXT_CD68"
      Tab(1).Control(29)=   "TXT_CD61"
      Tab(1).Control(30)=   "txt_cd103"
      Tab(1).Control(31)=   "txt_CD79B"
      Tab(1).Control(32)=   "TXT_CD117"
      Tab(1).Control(33)=   "TXT_CD138"
      Tab(1).Control(34)=   "TXT_CYTOMU"
      Tab(1).Control(35)=   "TXT_FMC7"
      Tab(1).Control(36)=   "TXT_DBA44"
      Tab(1).Control(37)=   "TXT_ANTIMPO"
      Tab(1).Control(38)=   "txt_Glycophorina"
      Tab(1).Control(39)=   "TXT_HLADR"
      Tab(1).Control(40)=   "txt_kappa"
      Tab(1).Control(41)=   "txt_lanbda"
      Tab(1).Control(42)=   "opt_PrePBSC_34"
      Tab(1).Control(43)=   "opt_PBSCharvest_34"
      Tab(1).Control(44)=   "opt_BMharvest_34"
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "Screen III"
      TabPicture(2)   =   "SURFACEFORM.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl_tdt"
      Tab(2).Control(1)=   "LBL_SMIG"
      Tab(2).Control(2)=   "Label17"
      Tab(2).Control(3)=   "Label1"
      Tab(2).Control(4)=   "Label2"
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(6)=   "Label4"
      Tab(2).Control(7)=   "Label5"
      Tab(2).Control(8)=   "Label6"
      Tab(2).Control(9)=   "Label7"
      Tab(2).Control(10)=   "Label8"
      Tab(2).Control(11)=   "Label9"
      Tab(2).Control(12)=   "Label10"
      Tab(2).Control(13)=   "Shape1"
      Tab(2).Control(14)=   "Label11"
      Tab(2).Control(15)=   "TXT_TDT"
      Tab(2).Control(16)=   "txt_smig"
      Tab(2).Control(17)=   "Txt_IgM"
      Tab(2).Control(18)=   "txt_clnd1"
      Tab(2).Control(19)=   "txt_tcrab"
      Tab(2).Control(20)=   "txt_tcrgd"
      Tab(2).Control(21)=   "txt_smpctg"
      Tab(2).Control(22)=   "Frame1"
      Tab(2).Control(23)=   "Frame2"
      Tab(2).Control(24)=   "Frame3"
      Tab(2).ControlCount=   25
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   -70920
         TabIndex        =   119
         Top             =   3120
         Width           =   2895
         Begin VB.OptionButton opt_sitaram 
            Caption         =   "Mr. Sitaram"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   122
            Top             =   840
            Width           =   2175
         End
         Begin VB.OptionButton opt_mryb 
            Caption         =   "Mr. Y. Badrinath"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   121
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton opt_drak 
            Caption         =   "Dr. Ashok Kumar"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.OptionButton opt_Tcell_CD8 
         Caption         =   "T cell enumeration"
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
         Left            =   7680
         TabIndex        =   118
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton opt_Tcell_CD4 
         Caption         =   "T cell enumeration"
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
         Left            =   4080
         TabIndex        =   116
         Top             =   3000
         Width           =   2295
      End
      Begin VB.OptionButton opt_BMharvest_CD3 
         Caption         =   "BM Harvest"
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
         Left            =   1080
         TabIndex        =   114
         Top             =   3720
         Width           =   1455
      End
      Begin VB.OptionButton opt_PBSCharvest_CD3 
         Caption         =   "PBSC Harvest"
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
         Left            =   1080
         TabIndex        =   113
         Top             =   3360
         Width           =   1575
      End
      Begin VB.OptionButton opt_PrePBSC_CD3 
         Caption         =   "Pre PBSC"
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
         Left            =   1080
         TabIndex        =   112
         Top             =   3000
         Width           =   1455
      End
      Begin VB.OptionButton opt_BMharvest_34 
         Caption         =   "BM Harvest"
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
         Left            =   -73560
         TabIndex        =   110
         Top             =   3600
         Width           =   1455
      End
      Begin VB.OptionButton opt_PBSCharvest_34 
         Caption         =   "PBSC Harvest"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -73560
         TabIndex        =   108
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton opt_PrePBSC_34 
         Caption         =   "Pre PBSC"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -73560
         TabIndex        =   107
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   -70800
         TabIndex        =   106
         Top             =   3000
         Width           =   4095
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   -70920
         TabIndex        =   103
         Top             =   2160
         Width           =   4095
         Begin VB.OptionButton opt_lsp 
            Caption         =   "Light scatter plots."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   480
            Width           =   3375
         End
         Begin VB.OptionButton opt_ssp 
            Caption         =   "Side scatter/CD 45 plots."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   120
            Width           =   3375
         End
      End
      Begin VB.TextBox txt_smpctg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   98
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txt_tcrgd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67680
         TabIndex        =   94
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txt_tcrab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70200
         TabIndex        =   91
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txt_clnd1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72960
         TabIndex        =   89
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txt_lanbda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65535
         TabIndex        =   86
         Top             =   2055
         Width           =   1455
      End
      Begin VB.TextBox Txt_IgM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72960
         TabIndex        =   84
         Top             =   510
         Width           =   1455
      End
      Begin VB.TextBox txt_smig 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67680
         TabIndex        =   80
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TXT_TDT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70200
         TabIndex        =   79
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt_kappa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68280
         TabIndex        =   75
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TXT_HLADR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71040
         TabIndex        =   74
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txt_Glycophorina 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         TabIndex        =   73
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TXT_ANTIMPO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         TabIndex        =   68
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TXT_DBA44 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68280
         TabIndex        =   67
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TXT_FMC7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65520
         TabIndex        =   66
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TXT_CYTOMU 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71040
         TabIndex        =   65
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TXT_CD138 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65520
         TabIndex        =   60
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD117 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68280
         TabIndex        =   59
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_CD79B 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         TabIndex        =   58
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_cd103 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71040
         TabIndex        =   57
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox TXT_CD61 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71040
         TabIndex        =   52
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TXT_CD68 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68280
         TabIndex        =   51
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TXT_CD56 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         TabIndex        =   50
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txt_cytocd79a 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65520
         TabIndex        =   49
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD34 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73560
         TabIndex        =   44
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TXT_CD38 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71040
         TabIndex        =   43
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TXT_CD42 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65520
         TabIndex        =   42
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txt_cd41 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68280
         TabIndex        =   41
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txt_CD33 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   36
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD25 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   35
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD23 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   34
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txt_cytocd221 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   33
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD20 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   28
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD22 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   27
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD19 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3585
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txt_cd36 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   1680
         Width           =   1440
      End
      Begin VB.TextBox TXT_CD11C 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txt_cd13 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txt_cd14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txt_CD15 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8880
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox TXT_CYTOCD3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TXT_CD2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         Height          =   495
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label18 
         Caption         =   "CD8"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   117
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape Shape4 
         Height          =   495
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label21 
         Caption         =   "CD4/CD8"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   115
         Top             =   3000
         Width           =   975
      End
      Begin VB.Shape Shape3 
         Height          =   1335
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   "CD3"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   111
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "CD 34"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   109
         Top             =   2880
         Width           =   615
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   -74640
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Flow Cytometry is Done By "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   102
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H000000C0&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   3255
         Left            =   -74160
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   8055
      End
      Begin VB.Label Label10 
         Caption         =   "Instrument: BD-FACS Calibur E 2237"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   101
         Top             =   4320
         Width           =   4095
      End
      Begin VB.Label Label9 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67320
         TabIndex        =   100
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "of the gated events on"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   99
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "The above positive percentage were determined on"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   97
         Top             =   1800
         Width           =   7815
      End
      Begin VB.Label Label6 
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
         Left            =   -67920
         TabIndex        =   96
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   -70560
         TabIndex        =   95
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "TCR"
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
         Left            =   -68400
         TabIndex        =   93
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   92
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "TCR"
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
         Left            =   -71040
         TabIndex        =   90
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Cyclin D1"
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
         Left            =   -73920
         TabIndex        =   88
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbl_lanbda 
         AutoSize        =   -1  'True
         Caption         =   "Lambda"
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
         Left            =   -66495
         TabIndex        =   87
         Top             =   2055
         Width           =   675
      End
      Begin VB.Label Label17 
         Caption         =   "IgM"
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
         Left            =   -73440
         TabIndex        =   85
         Top             =   510
         Width           =   375
      End
      Begin VB.Label LBL_SMIG 
         AutoSize        =   -1  'True
         Caption         =   "SmIg"
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
         Left            =   -68400
         TabIndex        =   82
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lbl_tdt 
         AutoSize        =   -1  'True
         Caption         =   "Tdt"
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
         Left            =   -71040
         TabIndex        =   81
         Top             =   480
         Width           =   300
      End
      Begin VB.Label lbl_kappa 
         AutoSize        =   -1  'True
         Caption         =   "Kappa"
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
         Left            =   -69000
         TabIndex        =   78
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label LBL_HALDR 
         AutoSize        =   -1  'True
         Caption         =   "HLA DR"
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
         Left            =   -71880
         TabIndex        =   77
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label LBL_GLYCOPHORIN 
         AutoSize        =   -1  'True
         Caption         =   "CD235a(Gly A)"
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
         Left            =   -74880
         TabIndex        =   76
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label LBL_ANTIMPO 
         AutoSize        =   -1  'True
         Caption         =   "AntiMPO"
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
         TabIndex        =   72
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label LBL_FMC7 
         AutoSize        =   -1  'True
         Caption         =   "FMC7"
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
         Left            =   -66120
         TabIndex        =   71
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lbl_DBA44 
         AutoSize        =   -1  'True
         Caption         =   "DBA 44"
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
         Height          =   195
         Left            =   -69000
         TabIndex        =   70
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label LBL_CYTOMU 
         AutoSize        =   -1  'True
         Caption         =   "Cyto mu"
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
         Left            =   -71910
         TabIndex        =   69
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label LBL_CD138 
         AutoSize        =   -1  'True
         Caption         =   " CD138"
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
         Left            =   -66240
         TabIndex        =   64
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label13 
         Caption         =   "CD117"
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
         Left            =   -69000
         TabIndex        =   63
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label LBL_CD79B 
         AutoSize        =   -1  'True
         Caption         =   "CD79 b"
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
         Left            =   -74400
         TabIndex        =   62
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lbl_cd103 
         AutoSize        =   -1  'True
         Caption         =   "CD103"
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
         Left            =   -71790
         TabIndex        =   61
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label12 
         Caption         =   "CD 61"
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
         Left            =   -71760
         TabIndex        =   56
         Top             =   960
         Width           =   615
      End
      Begin VB.Label LBL_CD68 
         AutoSize        =   -1  'True
         Caption         =   "IgD"
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
         Left            =   -68760
         TabIndex        =   55
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label14 
         Caption         =   "CD56"
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
         Left            =   -74280
         TabIndex        =   54
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbl_cytocd79a 
         AutoSize        =   -1  'True
         Caption         =   "CytoCD79a"
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
         Left            =   -66555
         TabIndex        =   53
         Top             =   960
         Width           =   960
      End
      Begin VB.Label lbl_cd34 
         AutoSize        =   -1  'True
         Caption         =   "CD 34"
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
         Left            =   -74280
         TabIndex        =   48
         Top             =   600
         Width           =   540
      End
      Begin VB.Label LBL_CD38 
         AutoSize        =   -1  'True
         Caption         =   "CD 38"
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
         Left            =   -71760
         TabIndex        =   47
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label15 
         Caption         =   "CD 55"
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
         Left            =   -66120
         TabIndex        =   46
         Top             =   600
         Width           =   615
      End
      Begin VB.Label LBL_CD41 
         AutoSize        =   -1  'True
         Caption         =   "CD41"
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
         Left            =   -68880
         TabIndex        =   45
         Top             =   600
         Width           =   480
      End
      Begin VB.Label LBL_CD33 
         AutoSize        =   -1  'True
         Caption         =   "CD 33"
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
         Left            =   8280
         TabIndex        =   40
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label LBL_CD25 
         AutoSize        =   -1  'True
         Caption         =   "CD 25"
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
         Left            =   5640
         TabIndex        =   39
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label LBL_CD23 
         AutoSize        =   -1  'True
         Caption         =   "CD 23"
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
         Left            =   3000
         TabIndex        =   38
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label lbl_cyto221 
         AutoSize        =   -1  'True
         Caption         =   "Cyto CD22"
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
         TabIndex        =   37
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "CD 20"
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
         Left            =   5640
         TabIndex        =   32
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label LBL_CD22 
         AutoSize        =   -1  'True
         Caption         =   "CD 22"
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
         Left            =   8280
         TabIndex        =   31
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label LBL_CD19 
         AutoSize        =   -1  'True
         Caption         =   "CD 19"
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
         Left            =   3000
         TabIndex        =   30
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label LBL_CD36 
         Caption         =   "CD16"
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
         Left            =   480
         TabIndex        =   29
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label LBL_CD11C 
         AutoSize        =   -1  'True
         Caption         =   "CD11c"
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
         Left            =   480
         TabIndex        =   24
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label LBL_CD13 
         AutoSize        =   -1  'True
         Caption         =   "CD 13"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lbl_cd14 
         AutoSize        =   -1  'True
         Caption         =   "CD 14"
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
         Left            =   5640
         TabIndex        =   22
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lbl_cd15 
         AutoSize        =   -1  'True
         Caption         =   "CD 15"
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
         Left            =   8280
         TabIndex        =   21
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label LBL_CD10 
         AutoSize        =   -1  'True
         Caption         =   "CD 10"
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
         Left            =   8280
         TabIndex        =   16
         Top             =   960
         Width           =   540
      End
      Begin VB.Label LBL_CD8 
         AutoSize        =   -1  'True
         Caption         =   "CD 8"
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
         Left            =   5640
         TabIndex        =   15
         Top             =   960
         Width           =   435
      End
      Begin VB.Label LBL_CD7 
         AutoSize        =   -1  'True
         Caption         =   "CD 7"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   960
         Width           =   435
      End
      Begin VB.Label LBL_CD5 
         AutoSize        =   -1  'True
         Caption         =   "CD 5"
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
         Left            =   480
         TabIndex        =   13
         Top             =   960
         Width           =   435
      End
      Begin VB.Label LBL_CYTOCD3 
         AutoSize        =   -1  'True
         Caption         =   "Cyto CD3"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   600
         Width           =   810
      End
      Begin VB.Label LBL_CD4 
         AutoSize        =   -1  'True
         Caption         =   "CD 4"
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
         Left            =   8280
         TabIndex        =   7
         Top             =   600
         Width           =   435
      End
      Begin VB.Label LBL_CD3 
         AutoSize        =   -1  'True
         Caption         =   "CD 3"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   600
         Width           =   435
      End
      Begin VB.Label LBL_CD2 
         AutoSize        =   -1  'True
         Caption         =   "CD 2"
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
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   435
      End
   End
End
Attribute VB_Name = "SURFACEFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SEL_HCD3_QR As rdoQuery
Dim SEL_HCD3_RS As rdoResultset

Dim SEL_HCD34_QR As rdoQuery
Dim SEL_HCD34_RS As rdoResultset

Dim SEL_CD4CD8_QR As rdoQuery
Dim SEL_CD4CD8_RS As rdoResultset

Dim str As String

Private Sub Command1_Click()
 Me.Hide
 hemcbcreport.show
 hemcbcreport.SSTab1.Tab = PREVTAB
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
opt_PrePBSC_34.Value = False
opt_PBSCharvest_34.Value = False
opt_BMharvest_34.Value = False

opt_Tcell_CD4.Value = False

opt_Tcell_CD8.Value = False
End Sub

Private Sub opt_BMharvest_34_Click()
Me.Hide
frm_CD34.show
End Sub



Private Sub opt_PBSCharvest_34_Click()
Me.Hide
frm_CD34.show
End Sub

Private Sub opt_PrePBSC_34_Click()
Me.Hide
frm_CD34.show
End Sub

Private Sub opt_PrePBSC_CD3_Click()
Me.Hide
frm_CD3.show
End Sub
Private Sub opt_PBSCharvest_CD3_Click()
Me.Hide
frm_CD3.show
End Sub

Private Sub opt_BMharvest_CD3_Click()
Me.Hide
frm_CD3.show
End Sub

Private Sub opt_Tcell_CD4_Click()
Me.Hide
frm_TcellEnumeration.show
End Sub

Private Sub opt_Tcell_CD8_Click()
Me.Hide
frm_TcellEnumeration.show
End Sub

Public Sub ShowProperOptions()
If hemcbcreport.txt_VIEW.Text = True Then
    '***** CD3 *****
    str = "SELECT * FROM DISLIB.HAEMATCD3 WHERE REQNO=?"
    Set SEL_HCD3_QR = rdb.CreateQuery("", str)
    
    SEL_HCD3_QR(0) = req_chglen(hemcbcreport.cmb_reqno)
    Set SEL_HCD3_RS = SEL_HCD3_QR.OpenResultset(1, 3)
    
    If Not SEL_HCD3_RS.EOF Then
        If Trim(SEL_HCD3_RS("FRMCAPTION")) = "Pre PBSC CD3 Enumeration" Then
            opt_PrePBSC_CD3.Enabled = True
            opt_PBSCharvest_CD3.Enabled = False
            opt_BMharvest_CD3.Enabled = False
        ElseIf Trim(SEL_HCD3_RS("FRMCAPTION")) = "PBSC Harvest CD3 Enumeration" Then
            opt_PrePBSC_CD3.Enabled = False
            opt_PBSCharvest_CD3.Enabled = True
            opt_BMharvest_CD3.Enabled = False
        ElseIf Trim(SEL_HCD3_RS("FRMCAPTION")) = "BM Harvest CD3 Enumeration" Then
            opt_PrePBSC_CD3.Enabled = False
            opt_PBSCharvest_CD3.Enabled = False
            opt_BMharvest_CD3.Enabled = True
        End If
    Else
        opt_PrePBSC_CD3.Enabled = False
        opt_PBSCharvest_CD3.Enabled = False
        opt_BMharvest_CD3.Enabled = False
    End If
    
    '***** CD4/CD8 *****
    str = "SELECT * FROM DISLIB.HMTCD4CD8 WHERE REQNO=?"
    Set SEL_CD4CD8_QR = rdb.CreateQuery("", str)
    
    SEL_CD4CD8_QR(0) = req_chglen(hemcbcreport.cmb_reqno)
    Set SEL_CD4CD8_RS = SEL_CD4CD8_QR.OpenResultset(1, 3)
    
    If Not SEL_CD4CD8_RS.EOF Then
        If Trim(SEL_CD4CD8_RS("FRMCAPTION")) = "CD4/CD8 T Cell Enumeration" Then
            opt_Tcell_CD4.Enabled = True
            'opt_Tcell_CD8.Enabled = False
'        ElseIf Trim(SEL_CD4CD8_RS("FRMCAPTION")) = "CD8 T Cell Enumeration" Then
'            opt_Tcell_CD4.Enabled = False
'            opt_Tcell_CD8.Enabled = True
        End If
    Else
        opt_Tcell_CD4.Enabled = False
 '       opt_Tcell_CD8.Enabled = False
    End If
    '*******************
    
    '***** CD34 *****
    str = "SELECT * FROM DISLIB.HAEMATCD34 WHERE REQNO=?"
    Set SEL_HCD34_QR = rdb.CreateQuery("", str)
    
    SEL_HCD34_QR(0) = req_chglen(hemcbcreport.cmb_reqno)
    Set SEL_HCD34_RS = SEL_HCD34_QR.OpenResultset(1, 3)
    
    If Not SEL_HCD34_RS.EOF Then
        If Trim(SEL_HCD34_RS("FRMCAPTION")) = "Pre PBSC CD34 Enumeration" Then
            opt_PrePBSC_34.Enabled = True
            opt_PBSCharvest_34.Enabled = False
            opt_BMharvest_34.Enabled = False
        ElseIf Trim(SEL_HCD34_RS("FRMCAPTION")) = "PBSC Harvest CD34 Enumeration" Then
            opt_PrePBSC_34.Enabled = False
            opt_PBSCharvest_34.Enabled = True
            opt_BMharvest_34.Enabled = False
        ElseIf Trim(SEL_HCD34_RS("FRMCAPTION")) = "BM Harvest CD34 Enumeration" Then
            opt_PrePBSC_34.Enabled = False
            opt_PBSCharvest_34.Enabled = False
            opt_BMharvest_34.Enabled = True
        End If
    Else
        opt_PrePBSC_34.Enabled = False
        opt_PBSCharvest_34.Enabled = False
        opt_BMharvest_34.Enabled = False
    End If
    '****************
End If
End Sub

Private Sub Option2_Click()

End Sub
