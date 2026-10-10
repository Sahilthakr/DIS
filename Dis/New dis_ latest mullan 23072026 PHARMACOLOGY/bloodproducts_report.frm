VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form bloodproducts_report 
   Caption         =   "Blood Products Report"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
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
   ScaleHeight     =   9765
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   360
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox txtusername 
      Height          =   315
      Left            =   1320
      TabIndex        =   63
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   73
      Top             =   4560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   192
      TabCaption(0)   =   "Blood Products Requisition"
      TabPicture(0)   =   "bloodproducts_report.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TXT_PBNAME"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TXT_NOOFUREP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Compatibility Report"
      TabPicture(1)   =   "bloodproducts_report.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox TXT_PBNAME 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73680
         TabIndex        =   104
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox TXT_NOOFUREP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67920
         TabIndex        =   103
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Frame Frame5 
         Caption         =   "Products Required"
         Height          =   975
         Left            =   -74880
         TabIndex        =   87
         Top             =   1320
         Width           =   11535
         Begin VB.TextBox TXT_FFP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5400
            TabIndex        =   108
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox CHK_WB 
            Caption         =   "WB"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            ToolTipText     =   "Whole Blood"
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox CHK_PC 
            Caption         =   "PC"
            Height          =   255
            Left            =   1560
            TabIndex        =   101
            ToolTipText     =   "Packed Cells"
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox CHK_RDP 
            Caption         =   "RDP"
            Height          =   255
            Left            =   2760
            TabIndex        =   100
            ToolTipText     =   "Platelet Conc."
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox CHK_FFP 
            Caption         =   "FFP"
            Height          =   255
            Left            =   5400
            TabIndex        =   99
            ToolTipText     =   "Fresh frozen Plasma"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox CHK_FVIIID 
            Caption         =   "FVIIID"
            Height          =   195
            Left            =   6960
            TabIndex        =   98
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox CHK_CRYO 
            Caption         =   "CRYO"
            Height          =   255
            Left            =   8520
            TabIndex        =   97
            ToolTipText     =   "Cryoprecipitate"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox CHK_ALB 
            Caption         =   "ALB"
            Height          =   255
            Left            =   10200
            TabIndex        =   96
            ToolTipText     =   "Albumin"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TXT_WB 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   95
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox TXT_PC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   94
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox TXT_RDP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            TabIndex        =   93
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox TXT_FVIIID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6960
            TabIndex        =   92
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox TXT_CRYO 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            TabIndex        =   91
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox TXT_ALB 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10200
            TabIndex        =   90
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox CHK_SDP 
            Caption         =   "SDP"
            Height          =   255
            Left            =   3960
            TabIndex        =   89
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox TXT_SDP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3960
            TabIndex        =   88
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   78
         Top             =   240
         Width           =   11535
         Begin VB.Frame Frame7 
            Caption         =   "Special  Requirements"
            Height          =   495
            Left            =   6120
            TabIndex        =   118
            Top             =   480
            Width           =   5175
            Begin VB.CheckBox CHK_IRRADIATION 
               Caption         =   "Irradiation"
               Height          =   255
               Left            =   2280
               TabIndex        =   120
               Top             =   120
               Width           =   255
            End
            Begin VB.CheckBox CHK_LEUCODEPLETION 
               Caption         =   "Leucodepletion"
               Height          =   255
               Left            =   3480
               TabIndex        =   119
               Top             =   120
               Width           =   255
            End
            Begin VB.Label Label16 
               Caption         =   "Leucodepletion"
               Height          =   255
               Left            =   3720
               TabIndex        =   123
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label15 
               Caption         =   "Irradiation"
               Height          =   255
               Left            =   2520
               TabIndex        =   122
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.TextBox TXT_PLTC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9720
            MaxLength       =   10
            TabIndex        =   116
            Top             =   120
            Width           =   1575
         End
         Begin VB.ComboBox CMB_PR 
            Height          =   315
            Left            =   2280
            TabIndex        =   82
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox CMB_BG 
            Height          =   315
            Left            =   4800
            TabIndex        =   81
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txt_hb 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            TabIndex        =   80
            Top             =   120
            Width           =   1215
         End
         Begin VB.ComboBox CMB_RFT 
            Height          =   315
            Left            =   2280
            TabIndex        =   79
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label Label14 
            Caption         =   "Plt. Count"
            Height          =   255
            Left            =   8640
            TabIndex        =   117
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Previous Recipent"
            Height          =   255
            Left            =   600
            TabIndex        =   86
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Blood Group"
            Height          =   195
            Left            =   3600
            TabIndex        =   85
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Reason for Transfusion"
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hb g/dl"
            Height          =   195
            Left            =   6240
            TabIndex        =   83
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   0
         TabIndex        =   74
         Top             =   300
         Width           =   11775
         Begin MSMask.MaskEdBox TXT_CROSSTIME 
            Height          =   255
            Left            =   10680
            TabIndex        =   19
            Top             =   600
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox txt_rem 
            Height          =   315
            ItemData        =   "bloodproducts_report.frx":0038
            Left            =   2040
            List            =   "bloodproducts_report.frx":003A
            TabIndex        =   9
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox TXT_BAGSEGNO 
            Height          =   285
            Left            =   2640
            TabIndex        =   10
            Top             =   600
            Width           =   1455
         End
         Begin VB.ComboBox TXT_IAT 
            Height          =   315
            Left            =   6960
            TabIndex        =   15
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox txt_enz 
            Height          =   315
            Left            =   6360
            TabIndex        =   14
            Top             =   600
            Width           =   615
         End
         Begin VB.ComboBox TXT_SAL 
            Height          =   315
            Left            =   5880
            TabIndex        =   13
            Top             =   600
            Width           =   495
         End
         Begin VB.ComboBox txt_compcode 
            Height          =   315
            Left            =   1320
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin VB.ComboBox TXT_GRP 
            Height          =   315
            Left            =   4080
            TabIndex        =   11
            Top             =   600
            Width           =   735
         End
         Begin MSMask.MaskEdBox txt_on 
            Height          =   255
            Left            =   7440
            TabIndex        =   16
            ToolTipText     =   "CROSSMATCH DATE"
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
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
         Begin VB.ComboBox txt_by 
            Height          =   315
            Left            =   8520
            TabIndex        =   17
            ToolTipText     =   "ENTERED BY"
            Top             =   600
            Width           =   975
         End
         Begin MSMask.MaskEdBox TXT_EXPDATE 
            Height          =   255
            Left            =   4800
            TabIndex        =   12
            ToolTipText     =   "EXP. DATE"
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox TXT_CROSSDATE 
            Height          =   255
            Left            =   9480
            TabIndex        =   18
            ToolTipText     =   "ISSUE DATE"
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
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
         Begin VB.TextBox txt_obvalue 
            Height          =   285
            Left            =   480
            TabIndex        =   7
            ToolTipText     =   "COMPATIBLE BAG NO."
            Top             =   600
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2175
            Left            =   0
            TabIndex        =   75
            Top             =   360
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   20
            Cols            =   14
            FormatString    =   $"bloodproducts_report.frx":003C
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "           ISSUED"
            Height          =   255
            Left            =   8520
            TabIndex        =   77
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            Caption         =   "            CROSSMATCH"
            Height          =   255
            Left            =   5520
            TabIndex        =   76
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Phlebotomist"
         Height          =   255
         Left            =   -74880
         TabIndex        =   106
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "No.of Units replaced"
         Height          =   195
         Left            =   -69840
         TabIndex        =   105
         Top             =   2400
         Width           =   1770
      End
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10680
      TabIndex        =   68
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10680
      TabIndex        =   67
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   62
      Top             =   7440
      Width           =   11775
      Begin VB.ComboBox TXT_COMMENTS 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   10215
      End
      Begin VB.TextBox txt_comments1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   64
         Top             =   240
         Width           =   10215
      End
      Begin VB.ComboBox cmb_enterby 
         Height          =   315
         Left            =   1080
         TabIndex        =   21
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label LBL_COMMENTS 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   66
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Entered by"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   720
         Width           =   930
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   855
      Left            =   0
      TabIndex        =   55
      Top             =   8520
      Width           =   11775
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         MouseIcon       =   "bloodproducts_report.frx":00D4
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MouseIcon       =   "bloodproducts_report.frx":03DE
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "bloodproducts_report.frx":06E8
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "bloodproducts_report.frx":09F2
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5880
         MouseIcon       =   "bloodproducts_report.frx":0CFC
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8760
         MouseIcon       =   "bloodproducts_report.frx":1006
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10440
         MouseIcon       =   "bloodproducts_report.frx":1310
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7200
         MouseIcon       =   "bloodproducts_report.frx":161A
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame common 
      Height          =   3015
      Left            =   0
      TabIndex        =   23
      Top             =   1440
      Width           =   11775
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   6360
         TabIndex        =   111
         Top             =   960
         Width           =   5175
         Begin VB.TextBox txt_redt 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   115
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txt_reqdt 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   114
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label13 
            Caption         =   "Required Date and Time"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Requisition Date and Time"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   1215
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10200
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_pathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10200
         TabIndex        =   35
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   34
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   32
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5280
         TabIndex        =   31
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   30
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   29
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5280
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5280
         TabIndex        =   27
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   26
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   25
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txt_relclnhis 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   24
         Top             =   2520
         Width           =   3615
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
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
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   53
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   52
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   51
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   50
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "DTM No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   49
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   48
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   47
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   46
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4320
         TabIndex        =   45
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   44
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   42
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4440
         TabIndex        =   41
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   40
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   39
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   2130
      End
      Begin VB.Label lbl_relclnhis 
         AutoSize        =   -1  'True
         Caption         =   "Clinical Diagnosis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6360
         TabIndex        =   37
         Top             =   2520
         Width           =   1515
      End
   End
   Begin VB.Label Label17 
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
      Height          =   495
      Left            =   0
      TabIndex        =   124
      Top             =   0
      Width           =   11775
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
      Left            =   8040
      TabIndex        =   121
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   195
      Left            =   240
      TabIndex        =   110
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label8 
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   109
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "COMPATIBILITY REPORT"
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
      Height          =   375
      Left            =   3840
      TabIndex        =   107
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10080
      TabIndex        =   70
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9840
      TabIndex        =   69
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl_transmeddtm 
      AutoSize        =   -1  'True
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
      Left            =   3600
      TabIndex        =   54
      Top             =   480
      Width           =   75
   End
End
Attribute VB_Name = "bloodproducts_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim i As Integer
Public j As Integer
Dim sql As String
Dim rqsqlinterface As rdoQuery
Dim Rssqlinterface As rdoResultset

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

'Dim seldissecprov1_qr As rdoQuery
'Dim seldissecprov1_rs As rdoResultset
'Dim seldissecprovflag1 As Boolean

Dim seldissecfin1_qr As rdoQuery
Dim seldissecfin1_rs As rdoResultset
Dim seldissecfinflag1 As Boolean

Dim SELDREQSTAFFDQR As rdoQuery
Dim SELDREQSTAFFDRS As rdoResultset
Dim SELDREQSTAFFD1FLAG As Boolean
Dim upddreqnoqr As rdoQuery
Dim SELDMULPROCQR As rdoQuery
Dim seldmulprocrs As rdoResultset
Dim seldmulprocflag As Boolean
Dim selMA101QR As rdoQuery
Dim SELMA101RS As rdoResultset
Dim SELMA101FLAG As Boolean
Dim SELDTEMPLETQR As rdoQuery
Dim SELDTEMPLETRS As rdoResultset
Dim SELDTEMPLETFLAG As Boolean
Dim INSDMPROCEDURQR As rdoQuery
Dim UPDDMPROCEDURQR As rdoQuery
Dim SELDMA101QR As rdoQuery
Dim SELDMA101RS As rdoResultset
Dim SELDMA101FLAG As Boolean
Dim SELDPROCEDURQR As rdoQuery
Dim SELDPROCEDURRS As rdoResultset
Dim SELDPROCEDURFLAG As Boolean
Dim SELDBCPROSUBQR As rdoQuery
Dim SELDBCPROSUBRS As rdoResultset
Dim SELDBCPROSUBFLAG As Boolean

Dim INSDMBCCOMMQR As rdoQuery
Dim UPDDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMRS As rdoResultset
Dim SELDMBCCOMMFLAG As Boolean
'Dim updrrequisitn1_qr As rdoQuery
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean

Dim SELDBCALKPHSQR As rdoQuery
Dim SELDBCALKPHSRS As rdoResultset
Dim SELDBCALKPHSFLAG As Boolean
Dim RNAGE As String
Dim SS As String
Dim ALKLV As Integer
Dim ALKHV As Integer
Dim LV As Integer
Dim HV As Integer
Dim PROTEIN As Variant
Dim GLOBULIN As Variant

Dim SELBLOODPRODUQR As rdoQuery
Dim SELBLOODPRODURS As rdoResultset
Dim SELBLOODPRODUFLAG As Boolean

''DCOMRP
Dim INSCOMRP_QR As rdoQuery
Dim INSCOMRP_RS As rdoResultset
Dim INSCOMRP_FLAG As Boolean

Dim NROWS
Dim rowno

Public alreadyRows As Integer
Public preMsfrow As Integer

Dim UPDTA1012_QR As rdoQuery
Dim UPDTA1012_RS As rdoResultset
Dim UPDTA1012_FLAG As Boolean


Dim INSDMBCCOMM_QR As rdoQuery

Dim UPDDMBCCOMM_QR As rdoQuery

Dim SELCOMEXT_QR As rdoQuery
Dim SELCOMEXT_RS As rdoResultset
Dim SELCOMEXT_FLAG As Boolean

Dim INSCOMEXT_QR As rdoQuery

Dim UPDCOMEXT_QR As rdoQuery

Dim SELDCOMRP_QR As rdoQuery
Dim SELDCOMRP_RS As rdoResultset
Dim SELDCOMRP_FLAG As Boolean

Dim SELDMBCCOMM_QR As rdoQuery
Dim SELDMBCCOMM_RS As rdoResultset
Dim SELDMBCCOMM_FLAG As Boolean

Dim UPDDCOMRP_QR As rdoQuery

Dim CHKDCOMRP_QR As rdoQuery
Dim CHKDCOMRP_RS As rdoResultset
Dim CHKDCOMRP_FLAG As Boolean

Dim MODFLAG As Boolean

Dim SELDTMNO1_QR As rdoQuery
Dim SELDTMNO1_RS As rdoResultset
Dim SELDTMNO1_FLAG As Boolean


''select data from dtmuser table

Dim DTMUSER_QR As rdoQuery
Dim DTMUSER_RS As rdoResultset

'SELECT DTMUSER
Dim SELDTMUSER_FLAG As Boolean
Dim SELDTMUSER_RS As rdoResultset
Dim SELDTMUSER_QR As rdoQuery


Dim usr As String

'''requisition date & time
Dim SELREQDT_QR As rdoQuery
Dim SELREQDT_RS As rdoResultset
Dim SELREQDT_FLAG As Boolean

'''requiRED date & time FOR BP
Dim SELREQDT1_QR As rdoQuery
Dim SELREQDT1_RS As rdoResultset
Dim SELREQDT1_FLAG As Boolean

'Dim SELdtmuser_QR As rdoQuery
'Dim SELdtmuser_RS As rdoResultset
'Dim SELdtmuser_FLAG As Boolean

Dim selsite_qr As rdoQuery
Dim selsite_rs As rdoResultset
Dim SELSITE_FLAG As Boolean

Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_FLAG As Boolean

''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean
Dim VAR_REQNNO As String
''*******ICU**********
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
Dim SELPATCATG_QR As rdoQuery
Dim SELPATCATG_RS As rdoResultset
Dim SELPATCATG_FLAG As Boolean


Private Sub CHK_ALB_Click()
If CHK_ALB.Value = 1 Then
    TXT_ALB.Visible = True
    'TXT_ALB.SetFocus
Else
    TXT_ALB.Visible = False
End If
End Sub

Private Sub CHK_CRYO_Click()
If CHK_CRYO.Value = 1 Then
    TXT_CRYO.Visible = True
    'TXT_CRYO.SetFocus
Else
    TXT_CRYO.Visible = False
End If
End Sub

Private Sub CHK_FFP_Click()
If CHK_FFP.Value = 1 Then
    TXT_FFP.Visible = True
    'TXT_FFP.SetFocus
Else
    TXT_FFP.Visible = False
End If
End Sub

Private Sub CHK_FVIIID_Click()
If CHK_FVIIID.Value = 1 Then
    TXT_FVIIID.Visible = True
    'TXT_FVIIID.SetFocus
Else
    TXT_FVIIID.Visible = False
End If
End Sub

Private Sub CHK_PC_Click()
If CHK_PC.Value = 1 Then
    TXT_PC.Visible = True
    'TXT_PC.SetFocus
Else
    TXT_PC.Visible = False
End If
End Sub

Private Sub CHK_RDP_Click()
If CHK_RDP.Value = 1 Then
    TXT_RDP.Visible = True
    'TXT_RDP.SetFocus
Else
    TXT_RDP.Visible = False
End If
End Sub

Private Sub CHK_SDP_Click()
If CHK_SDP.Value = 1 Then
    TXT_SDP.Visible = True
    'TXT_SDP.SetFocus
Else
    TXT_SDP.Visible = False
End If
End Sub

Private Sub CHK_WB_Click()
If CHK_WB.Value = 1 Then
    TXT_WB.Visible = True
    'TXT_WB.SetFocus
Else
    TXT_WB.Visible = False
End If
End Sub

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
        If Trim(cmb_dept.List(ind_cbodept)) = "TRANSFUSION MEDICINE$NZZ" Then
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
'If Trim(txt_modcode.Text) = "BP" Then
'    TXT_COMMENTS.clear
'End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:

Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "BLOOD PRODUCTS$BP" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
Dim INTFAC As Variant
Dim entr As String
On Error GoTo MSG:
'If Trim(usr) <> "" Then
'Else
'    MsgBox " Pls Enter User Name & Password !!!"
'    Exit Sub
'End If
If SELREQDT_FLAG = False Then
    SELREQDT_QR(0) = req_chglen(cmb_reqno.Text)
    Set SELREQDT_RS = SELREQDT_QR.OpenResultset(1, 3)
    SELREQDT_FLAG = False
Else
    SELREQDT_QR(0) = req_chglen(cmb_reqno.Text)
    SELREQDT_RS.MoreResults
    SELREQDT_RS.Requery
End If
If Not SELREQDT_RS.EOF Then
    'txt_redt.Enabled = False
    txt_redt.Text = SELREQDT_RS("REQNDT") & " " & SELREQDT_RS("REQNTM")
End If
If SELREQDT1_FLAG = False Then
    SELREQDT1_QR(0) = req_chglen(cmb_reqno.Text)
    Set SELREQDT1_RS = SELREQDT1_QR.OpenResultset(1, 3)
    SELREQDT1_FLAG = False
Else
    SELREQDT1_QR(0) = req_chglen(cmb_reqno.Text)
    SELREQDT1_RS.MoreResults
    SELREQDT1_RS.Requery
End If
If Not SELREQDT1_RS.EOF Then
'    txt_reqdt.Enabled = False
    txt_reqdt.Text = SELREQDT1_RS("REQUION") & " " & Trim(SELREQDT1_RS("TIME")) & " " & Trim(SELREQDT1_RS("fld1"))
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

'            If SELCOMEXT_FLAG = False Then
'                SELCOMEXT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
'                Set SELCOMEXT_RS = SELCOMEXT_QR.OpenResultset(1, 3)
'                SELCOMEXT_FLAG = False
'            Else
'                SELCOMEXT_RS.MoreResults
'                SELCOMEXT_RS.Requery
'            End If
'            If Not SELCOMEXT_RS.EOF Then
'                TXT_PLTC.Text = "" & SELCOMEXT_RS("PLTCT")
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
                If Trim(REFDR_RS("patNstat")) = "IP" Then
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
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF Then
                    txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
                'txt_refdoc.Text = SELPATHDREQUISITN_RS("REFDOC")
       
            End If
            If SELDTMNO1_FLAG = False Then
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                Set SELDTMNO1_RS = SELDTMNO1_QR.OpenResultset(1, 3)
                SELDTMNO1_FALG = True
            Else
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                SELDTMNO1_RS.MoreResults
                SELDTMNO1_RS.Requery
            End If
            If SELDTMNO1_RS.RowCount > 0 Then
                txt_pathno.Text = SELDTMNO1_RS("DTMNO")
                CMB_BG.Text = " " & SELDTMNO1_RS("fld1")
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
                txt_relclnhis.Text = "" & selta1012_rs("prvdiagdes")
            End If

           '****
           If SELBLOODPRODUFLAG = False Then
                SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                Set SELBLOODPRODURS = SELBLOODPRODUQR.OpenResultset(1, 3)
                SELBLOODPRODUFLAG = True
            Else
                SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                SELBLOODPRODURS.MoreResults
                SELBLOODPRODURS.Requery
            End If
                    If SELBLOODPRODURS.RowCount > 0 Then
                        'CMB_BG.Text = " " & SELBLOODPRODURS("BG")
                        CMB_RFT.Text = "" & SELBLOODPRODURS("RFT")
                        TXT_HB.Text = "" & SELBLOODPRODURS("HB")
                        TXT_NOOFUREP.Text = "" & SELBLOODPRODURS("NOOFUNREP")
'                        TXT_DT.Text = " " & Format(Trim(SELBLOODPRODURS("REQUION")), "DD/MM/YYYY")
'                        TXT_TM.Text = " " & SELBLOODPRODURS("TIME")
                        'cmb_tr.Text = "" & SELBLOODPRODURS("TYPEOFREQ")
                        CMB_PR.Text = "" & SELBLOODPRODURS("fld2")
                        If Len(SELBLOODPRODURS("WB")) > 0 Then
                            CHK_WB.Value = 1
                            TXT_WB.Text = " " & SELBLOODPRODURS("WB")
                        Else
                            CHK_WB.Value = 0
                            TXT_WB.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("PC")) > 0 Then
                            CHK_PC.Value = 1
                            TXT_PC.Text = " " & SELBLOODPRODURS("PC")
                        Else
                            CHK_PC.Value = 0
                            TXT_PC.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("RDP")) > 0 Then
                            CHK_RDP.Value = 1
                            TXT_RDP.Text = " " & SELBLOODPRODURS("RDP")
                        Else
                            CHK_RDP.Value = 0
                            TXT_RDP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("SDP")) > 0 Then
                            CHK_SDP.Value = 1
                            TXT_SDP.Text = " " & SELBLOODPRODURS("SDP")
                        Else
                            CHK_SDP.Value = 0
                            TXT_SDP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("FFP")) > 0 Then
                            CHK_FFP.Value = 1
                            TXT_FFP.Text = " " & SELBLOODPRODURS("FFP")
                        Else
                            CHK_FFP.Value = 0
                           TXT_FFP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("FVIIID")) > 0 Then
                            CHK_FVIIID.Value = 1
                            TXT_FVIIID.Text = " " & SELBLOODPRODURS("FVIIID")
                        Else
                            CHK_FVIIID.Value = 0
                            TXT_FVIIID.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("CRYO")) > 0 Then
                            CHK_CRYO.Value = 1
                            TXT_CRYO.Text = " " & SELBLOODPRODURS("CRYO")
                        Else
                            CHK_CRYO.Value = 0
                            TXT_CRYO.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("ALB")) > 0 Then
                            CHK_ALB.Value = 1
                            TXT_ALB.Text = " " & SELBLOODPRODURS("ALB")
                        Else
                            CHK_ALB.Value = 0
                            TXT_ALB.Text = " "
                        End If
                        
                        'TXT_OTHERS.Text = " " & SELBLOODPRODURS("OTHERS")
                        TXT_PBNAME.Text = "" & SELBLOODPRODURS("PHOBOTMIST")
                        TXT_NOOFUREP.Text = " " & SELBLOODPRODURS("NOOFUNREP")
                        If Len(SELBLOODPRODURS("SR2")) > 0 Then
                            CHK_LEUCODEPLETION.Value = 1
                        Else
                            CHK_LEUCODEPLETION.Value = 0
                        End If
                        If Len(SELBLOODPRODURS("SR1")) > 0 Then
                            CHK_IRRADIATION.Value = 1
                        Else
                            CHK_IRRADIATION.Value = 0
                        End If
                        TXT_PLTC.Text = "" & SELBLOODPRODURS("FLD3")
                End If
            
           
           '****
           
'
'                MSFlexGrid1.Rows = 10
'                MSFlexGrid1.Row = 1
'                MSFlexGrid1.Col = 1
'                txt_obvalue = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'                txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                txt_obvalue.Visible = True
'
            
        '
            

        
'        If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'        entr = ""
'            entr = cmb_enterby.Text
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'            cmb_enterby.Text = entr
'            entr = ""
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
                    cmd_modify.Enabled = False
                    'ADDENDUMFALG = True
                End If
                cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
                'cmb_finalby.Text = "" & REFDR_RS("FREPORTBY")
'                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
'                    cmb_finalby.Text = ""
'                Else
'                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
'                End If
            End If
            
'            If SELCOMEXT_FLAG = False Then
'                SELCOMEXT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
'                Set SELCOMEXT_RS = SELCOMEXT_QR.OpenResultset(1, 3)
'                SELCOMEXT_FLAG = False
'            Else
'                SELCOMEXT_RS.MoreResults
'                SELCOMEXT_RS.Requery
'            End If
'            If Not SELCOMEXT_RS.EOF Then
'                TXT_PLTC.Text = "" & SELCOMEXT_RS("PLTCT")
'            End If
'
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
                If Trim(REFDR_RS("patNstat")) = "IP" Then
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
                'txt_pathno.Text = "" & selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF = False Then
                    txt_relclnhis.Visible = True
                    lbl_relclnhis.Visible = True
                    txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
                End If
            End If
            If SELDTMNO1_FLAG = False Then
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                Set SELDTMNO1_RS = SELDTMNO1_QR.OpenResultset(1, 3)
                SELDTMNO1_FALG = True
            Else
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                SELDTMNO1_RS.MoreResults
                SELDTMNO1_RS.Requery
            End If
            If SELDTMNO1_RS.RowCount > 0 Then
                txt_pathno.Text = Trim(SELDTMNO1_RS("DTMNO"))
            End If
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = MaskEdBox1.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                txt_relclnhis.Text = "" & Trim(selta1012_rs("prvdiagdes"))
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
            '''''****** CHANGE
            If SELBLOODPRODUFLAG = False Then
                SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                Set SELBLOODPRODURS = SELBLOODPRODUQR.OpenResultset(1, 3)
                SELBLOODPRODUFLAG = True
            Else
                SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                SELBLOODPRODURS.MoreResults
                SELBLOODPRODURS.Requery
            End If
                    If SELBLOODPRODURS.RowCount > 0 Then
                        CMB_BG.Text = " " & Trim(SELBLOODPRODURS("BG"))
                        CMB_RFT.Text = "" & Trim(SELBLOODPRODURS("RFT"))
                        TXT_HB.Text = "" & Trim(SELBLOODPRODURS("HB"))
                        CMB_PR.Text = "" & Trim(SELBLOODPRODURS("FLD2"))
                        TXT_NOOFUREP.Text = "" & Trim(SELBLOODPRODURS("NOOFUNREP"))
                        'TXT_PLTC.Text = "" & Trim(SELBLOODPRODURS("FLD3"))
'                        TXT_DT.Text = " " & Format(Trim(SELBLOODPRODURS("REQUION")), "DD/MM/YYYY")
'                        TXT_TM.Text = " " & SELBLOODPRODURS("TIME")
                        'cmb_tr.Text = "" & SELBLOODPRODURS("TYPEOFREQ")
                        If Len(SELBLOODPRODURS("WB")) > 0 Then
                            CHK_WB.Value = 1
                            TXT_WB.Text = " " & SELBLOODPRODURS("WB")
                        Else
                            CHK_WB.Value = 0
                            TXT_WB.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("PC")) > 0 Then
                            CHK_PC.Value = 1
                            TXT_PC.Text = " " & SELBLOODPRODURS("PC")
                        Else
                            CHK_PC.Value = 0
                            TXT_PC.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("RDP")) > 0 Then
                            CHK_RDP.Value = 1
                            TXT_RDP.Text = " " & SELBLOODPRODURS("RDP")
                        Else
                            CHK_RDP.Value = 0
                            TXT_RDP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("SDP")) > 0 Then
                            CHK_SDP.Value = 1
                            TXT_SDP.Text = " " & SELBLOODPRODURS("SDP")
                        Else
                            CHK_SDP.Value = 0
                            TXT_SDP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("FFP")) > 0 Then
                            CHK_FFP.Value = 1
                            TXT_FFP.Text = " " & SELBLOODPRODURS("FFP")
                        Else
                            CHK_FFP.Value = 0
                           TXT_FFP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("FVIIID")) > 0 Then
                            CHK_FVIIID.Value = 1
                            TXT_FVIIID.Text = " " & SELBLOODPRODURS("FVIIID")
                        Else
                            CHK_FVIIID.Value = 0
                            TXT_FVIIID.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("CRYO")) > 0 Then
                            CHK_CRYO.Value = 1
                            TXT_CRYO.Text = " " & SELBLOODPRODURS("CRYO")
                        Else
                            CHK_CRYO.Value = 0
                            TXT_CRYO.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("ALB")) > 0 Then
                            CHK_ALB.Value = 1
                            TXT_ALB.Text = " " & SELBLOODPRODURS("ALB")
                        Else
                            CHK_ALB.Value = 0
                            TXT_ALB.Text = " "
                        End If
                        
                        'TXT_OTHERS.Text = " " & SELBLOODPRODURS("OTHERS")
                        TXT_PBNAME.Text = "" & Trim(SELBLOODPRODURS("PHOBOTMIST"))
                        TXT_NOUREPLACE = " " & Trim(SELBLOODPRODURS("NOOFUNREP"))
                        If Len(SELBLOODPRODURS("SR2")) > 0 Then
                            CHK_LEUCODEPLETION.Value = 1
                        Else
                            CHK_LEUCODEPLETION.Value = 0
                        End If
                        If Len(SELBLOODPRODURS("SR1")) > 0 Then
                            CHK_IRRADIATION.Value = 1
                        Else
                            CHK_IRRADIATION.Value = 0
                        End If
                        TXT_PLTC.Text = "" & SELBLOODPRODURS("FLD3")
                End If
'            txt_relclnhis.Visible = True
           
           '****
           
'
'                MSFlexGrid1.Rows = 10
'                MSFlexGrid1.Row = 1
'                MSFlexGrid1.Col = 1
'                txt_obvalue = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'                txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                txt_obvalue.Visible = True
''
            
        '
         ''*** REMARKS

        
'        If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'        entr = ""
'            entr = cmb_enterby.Text
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'            cmb_enterby.Text = entr
'            entr = ""
'        End If

'        If seldissecfinflag = False Then
'            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'            seldissecfinflag1 = True
'        Else
'            seldissecfin1_rs.MoreResults
'            seldissecfin1_rs.Requery
'        End If
        If SELDCOMRP_FLAG = False Then
            SELDCOMRP_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDCOMRP_RS = SELDCOMRP_QR.OpenResultset(1, 3)
            SELDCOMRP_FLAG = True
        Else
            SELDCOMRP_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDCOMRP_RS.MoreResults
            SELDCOMRP_RS.Requery
        End If
        
        If SELDMBCCOMM_FLAG = False Then
            SELDMBCCOMM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDMBCCOMM_RS = SELDMBCCOMM_QR.OpenResultset(1, 3)
            SELDMBCCOMM_FLAG = True
        Else
            SELDMBCCOMM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDMBCCOMM_RS.MoreResults
            SELDMBCCOMM_RS.Requery
        End If
        If Not SELDMBCCOMM_RS.EOF Then
            txt_comments.Text = "" & Trim(SELDMBCCOMM_RS("COMMENTS"))
        End If
        
        Dim i
        i = 1
        If Not SELDCOMRP_RS.EOF Then
            Do While Not SELDCOMRP_RS.EOF
                MSFlexGrid1.row = i
                MSFlexGrid1.TextMatrix(i, 0) = i
                MSFlexGrid1.TextMatrix(i, 1) = "" & Trim(SELDCOMRP_RS("CBBN"))
                MSFlexGrid1.TextMatrix(i, 2) = "" & Trim(SELDCOMRP_RS("compcode"))
                MSFlexGrid1.TextMatrix(i, 3) = "" & Trim(SELDCOMRP_RS("FLD1"))
                MSFlexGrid1.TextMatrix(i, 4) = "" & Trim(SELDCOMRP_RS("BSN"))
                MSFlexGrid1.TextMatrix(i, 5) = "" & Trim(SELDCOMRP_RS("GRP"))
                MSFlexGrid1.TextMatrix(i, 6) = "" & Trim(SELDCOMRP_RS("EXPDT"))
                MSFlexGrid1.TextMatrix(i, 7) = "" & Trim(SELDCOMRP_RS("SAL"))
                MSFlexGrid1.TextMatrix(i, 8) = "" & Trim(SELDCOMRP_RS("ENZ"))
                MSFlexGrid1.TextMatrix(i, 9) = "" & Trim(SELDCOMRP_RS("IAT"))
                MSFlexGrid1.TextMatrix(i, 10) = "" & Trim(SELDCOMRP_RS("ON"))
                MSFlexGrid1.TextMatrix(i, 11) = "" & Trim(SELDCOMRP_RS("BY"))
                MSFlexGrid1.TextMatrix(i, 12) = "" & Trim(SELDCOMRP_RS("IDT"))
                MSFlexGrid1.TextMatrix(i, 13) = "" & Trim(SELDCOMRP_RS("CTIME"))
                SELDCOMRP_RS.MoveNext
                i = i + 1
            Loop
        End If
  End If
End If
j = i
If opt_staff.Value = True Then
    STAFF
    

End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_reqno_LostFocus()
alreadyRows = 1
End Sub

Private Sub cmd_back_Click()
bloodproducts_report.Caption = "Dtm Lab"
Me.Hide
DIS_MAIN.show
btnenabl ("INI")
field_clear
viewflag = False
NEWFLAG = False
BCCLEAR
MSFlexGrid1.clear
GRIDSTRING
MSFlexGrid1.row = 1
MSFlexGrid1_Click
usr = ""
End Sub

Public Sub btnenabl(cap As String)
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
Private Sub cmd_commit_Click()
Dim i As Integer
'If Len(Trim(cmb_finalby)) = 0 Then
'    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
'    cmb_finalby.SetFocus
'    Exit Sub
'End If
    If MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation) = vbYes Then
        ''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_enterby.Text)
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
''********SECURITY*************************
'        If SELSITE_FLAG = False Then
'            SELSITE_QR(0) = req_chglen(Trim(cmb_reqno.Text))
'            Set SELSITE_RS = SELSITE_QR.OpenResultset(1, 3)
'            SELSITE_FLAG = True
'        Else
'            SELSITE_QR(0) = req_chglen(Trim(cmb_reqno.Text))
'            SELSITE_RS.MoreResults
'            SELSITE_RS.Requery
'        End If
'        If IsNull(SELSITE_RS("MEMONO")) Then
'            MsgBox "Memo is not made for this Requisition !!!", vbInformation, "DTM Report"
'            btnenabl ("INI")
'            Exit Sub
'        End If
'        If IsNull(SELSITE_RS("RECPTNO")) Then
'            MsgBox "Paymant is not made for this Requisition !!!", vbInformation, "DTM Report"
'            btnenabl ("INI")
'            Exit Sub
'        End If
        On Error GoTo MSG:
         rdb.BeginTrans
            updrrequisitn1_qr(0) = UCase(cmb_enterby.Text)
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

Private Sub cmd_modify_Click()
Dim str As String
flagsave = 2
MODFLAG = True
btnenabl ("NEW")
    FLDENABL ("MOD")
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
MSFlexGrid1.row = alreadyRows
'If seldissecentrflag1 = False Then
'    Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'    seldissecentrflag1 = True
'Else
'    seldissecentr1_rs.MoreResults
'    seldissecentr1_rs.Requery
'End If
'If seldissecentr1_rs.RowCount > 0 Then
'    STR = ""
'    STR = cmb_enterby.Text
'    cmb_enterby.clear
'    While Not seldissecentr1_rs.EOF
'        cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'        seldissecentr1_rs.MoveNext
'    Wend
'    cmb_enterby.Text = STR
'End If

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
'cmb_finalby.Enabled = True
'MSFlexGrid1.Row = 1
'MSFlexGrid1.Col = 2
'changes made by amarjeet on 08-11-2002 for  commit logic to disable txt_observed
'If ADDENDUMFALG Then 'means commited
'    txt_observed.Visible = False
'    TXT_COMMENTS.Enabled = True
'Else
'    Load frm_security
'    frm_security.show 1
'    If "RIDER" = UCase(password) Then
''        txt_obvalue = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
''        txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''        txt_obvalue.Visible = True
'        MSFlexGrid1_Click
'        password = ""
'    Else
'        MsgBox "Not Authorised to modify", vbCritical
'         Exit Sub
'
'    End If
'commented by amarjeet prev logic
'txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'txt_observed.Visible = True
txt_comments.Enabled = True

End Sub

Private Sub cmd_new_Click()
Dim z1 As Integer
Dim L1 As Integer
On Error GoTo MSG:
bloodproducts_report.Caption = "Dtm Lab"
MODFLAG = False
btnenabl ("NEW")
field_clear
TXT_CLEAR
BCCLEAR
txt_obvalue.Visible = True
TXT_BAGSEGNO.Visible = True
txt_rem.Visible = True
TXT_GRP.Visible = True
TXT_EXPDATE.Visible = True
TXT_SAL.Visible = True
txt_enz.Visible = True
TXT_IAT.Visible = True
txt_by.Visible = True
txt_on.Visible = True
TXT_CROSSDATE.Visible = True
TXT_CROSSTIME.Visible = True

If DATETIMEFLAG = False Then
   datetime
   DATETIMEFLAG = True
Else
    datetime
End If
txt_date.Text = DATE11
'TXT_CROSSDATE.text = DATE11
txt_time.Text = time1
cmb_enterby.Text = cmb_enterby.List(0)
cmb_enterby.Enabled = False

If SELDEPTRSFLAG1 = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SELDEPTRSFLAG1 = True
Else
    sel_deptrs.MoreResults
    sel_deptrs.Requery
End If
'If Trim(usr) <> "" Then
    If sel_deptrs.RowCount > 0 Then
        sel_deptrs.MoveFirst
        While Not sel_deptrs.EOF
            cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
            sel_deptrs.MoveNext
        Wend
    End If
'Else
 '   MsgBox " Pls Enter User Name & Password !!!"
  '  Exit Sub
'End If
Opt_patient.Value = True
'txt_observed.Visible = False
flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
MaskEdBox1.Enabled = True
MSFlexGrid1.clear
'MSFlexGrid1.Rows = 3
GRIDSTRING
'MSFlexGrid1.Row = 1
MSFlexGrid1_Click
'MsfSet
'txtusername.SetFocus
TXT_SAL.Text = "N"
TXT_SAL_Click
txt_enz.Text = "N"
txt_enz_Click
TXT_IAT.Text = "N"
TXT_IAT_Click
For z1 = 0 To txt_rem.ListCount - 1
    If Trim(txt_rem.List(z1)) = "-" Then
        txt_rem.Text = txt_rem.List(z1)
    End If
Next z1
txt_rem_Click
txt_on = DATE11
TXT_CROSSDATE = DATE11
'TXT_CROSSTIME = Format(time1, "####")
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_cancel_Click()
bloodproducts_report.Caption = "Dtm Lab"
btnenabl ("INI")
field_clear
viewflag = False
NEWFLAG = False
BCCLEAR
MSFlexGrid1.clear
GRIDSTRING
MSFlexGrid1.row = 1
MSFlexGrid1_Click
usr = ""
End Sub

Public Sub field_clear()
Dim obj As Object
lblicu.Visible = False
CHK_IRRADIATION.Value = 0
CHK_LEUCODEPLETION.Value = 0
TXT_PLTC.Text = ""
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
' MaskEdBox1.Mask = "#######"          ' COMMENTED BY VARINDER
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
txt_relclnhis.Text = " "
CHK_WB.Value = 0
CHK_PC.Value = 0
CHK_RDP.Value = 0
CHK_SDP.Value = 0
CHK_FFP.Value = 0
CHK_FVIIID.Value = 0
CHK_CRYO.Value = 0
CHK_ALB.Value = 0
txt_obvalue.Text = ""
'txt_obvalue.Visible = False
CMB_PR.Text = ""
CMB_RFT.Text = ""
CMB_BG.Text = ""
cmb_enterby.Text = ""
'cmb_finalby.Text = ""
txt_comments.Text = ""
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In bloodproducts_report.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub


Public Sub FLDENABL(EN As String)
If EN = "DIS" Then
        CHK_IRRADIATION.Enabled = False
        CHK_LEUCODEPLETION.Enabled = False
        TXT_PLTC.Enabled = False
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
        cmb_reqno.Enabled = False
        txt_pathno.Enabled = False
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
        txt_relclnhis.Visible = True
        'txt_observed.Visible = False
        txt_comments.Enabled = False
        'cmb_enterby.Enabled = False
        'cmb_finalby.Enabled = False
'        txt_relclnhis.Enabled = False
        CMB_PR.Enabled = False
        CMB_RFT.Enabled = False
        CMB_BG.Enabled = False
        TXT_HB.Enabled = False
        TXT_WB.Enabled = False
        TXT_PC.Enabled = False
        TXT_RDP.Enabled = False
        TXT_SDP.Enabled = False
        TXT_FFP.Enabled = False
        TXT_FVIIID.Enabled = False
        TXT_CRYO.Enabled = False
        TXT_ALB.Enabled = False
        CHK_WB.Enabled = False
        CHK_PC.Enabled = False
        CHK_RDP.Enabled = False
        CHK_SDP.Enabled = False
        CHK_FFP.Enabled = False
        CHK_FVIIID.Enabled = False
        CHK_CRYO.Enabled = False
        CHK_ALB.Enabled = False
        TXT_PBNAME.Enabled = False
        TXT_NOOFUREP.Enabled = False
        txt_redt.Enabled = False
        txt_reqdt.Enabled = False
        
End If
If EN = "ENA" Then
        'CHK_IRRADIATION.Enabled = True
        'CHK_LEUCODEPLETION.Enabled = True
        'TXT_PLTC.Enabled = True
        txt_relclnhis.Enabled = True
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        txt_pathno.Enabled = False
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
        txt_relclnhis.Visible = True
        txt_comments.Enabled = True
        cmb_enterby.Enabled = True
        'cmb_finalby.Enabled = True
End If
If EN = "MOD" Then
        'CHK_IRRADIATION.Enabled = True
        'CHK_LEUCODEPLETION.Enabled = True
        'TXT_PLTC.Enabled = True
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        txt_pathno.Enabled = True
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
        txt_relclnhis.Visible = True
        txt_relclnhis.Enabled = True
        txt_comments.Enabled = True
        cmb_enterby.Enabled = True
        'cmb_finalby.Enabled = True
End If
End Sub


Public Sub BCCLEAR()
TXT_CLEAR
cmb_dept.Text = ""
cmb_modality.Text = ""

'MSFlexGrid1.Rows = 3

End Sub

Private Sub cmd_print_Click()
'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim wb1 As String
Dim i As Integer
' On Error GoTo prtmsg:

DoEvents



For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i
VAR_REQNNO = req_chglen(cmb_reqno.Text)
   If Opt_patient.Value = True Then
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
    ''''memo check*********
        If SELPATCATG_FLAG = False Then
            SELPATCATG_QR(0) = VAR_REQNNO
            Set SELPATCATG_RS = SELPATCATG_QR.OpenResultset(1, 3)
            SELPATCATG_FLAG = True
        Else
            SELPATCATG_QR(0) = VAR_REQNNO
            SELPATCATG_RS.MoreResults
            SELPATCATG_RS.Requery
        End If
        If Not SELPATCATG_RS.EOF Then
            If IsNull(SELPATCATG_RS("MEMONO")) Then
            'commented by shweta 17-12-2013 as told by kalpan madam
                'MsgBox " Please Prepare the Memo !!!", vbInformation, "DIS"
              '  Exit Sub
            End If
        End If
    ''''memo check*********
    End If
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
            If SELRREQUISITN_RS.RowCount > 0 Then
                prepath = ""
                While Not SELRREQUISITN_RS.EOF
                    If Trim(SELRREQUISITN_RS("reqnno")) <> req_chglen(Trim(cmb_reqno.Text)) Then
                        prepath = prepath & req_chglen(SELRREQUISITN_RS("REQNNO")) & " "
                    End If
                    
                    SELRREQUISITN_RS.MoveNext
                Wend
             End If
             Dim pcount As Integer
             pcount = 0
            For i = 1 To MSFlexGrid1.Rows - 1
                If Trim(MSFlexGrid1.TextMatrix(i, 4)) <> "" Then
                  pcount = pcount + 1
                End If
            Next i
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
''add on 19/05
'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNO} = '" & req_chglen(cmb_reqno) & "'"
'***** Above statement commented and next added by Shraddha on 25/08/09 *****
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
If SELDTMNO_FLAG = False Then
     SELDTMNO_QR(0) = Trim(MaskEdBox1.FormattedText)
     Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
     SELDTMNO_FLAG = True
Else
    SELDTMNO_QR(0) = Trim(MaskEdBox1.FormattedText)
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
End If
If Opt_patient.Value = True Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBPNEW.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno.Text & "'"
                If Not SELDTMNO_RS.EOF Then
                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                        CrystalReport1.Formulas(6) = "DTMNO = '" & " " & " '"
                    Else
                        CrystalReport1.Formulas(6) = "DTMNO = '" & Trim(SELDTMNO_RS("DTMNO")) & " '"
                    End If
                End If
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                
    
    
'    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'
'
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.RPT"
'        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
'        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'
'
'
'        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(cmb_enterby.Text)
'            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'
'        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'        If pcount <> 0 Then
'
'            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'        Else
'            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
'        End If
End If
'**** STAFF
If opt_staff.Value = True Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBPSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno.Text & "'"
                CrystalReport1.Formulas(6) = "NAME = '" & Trim(txt_name.Text) & "'"
                
'STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'
'
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBCS.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'
'
'
'        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(cmb_enterby.Text)
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        If pcount <> 0 Then
'
'            CRYSTALREPORT1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(13) = "PAC = '" & "" & "'"
'        End If
'
'
End If
'***** STAFF
'CRYSTALREPORT1.WindowState = crptMaximized
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1

CrystalReport2.SelectionFormula = "{BLOODPROLF.REQNO} = '" & req_chglen(Trim(cmb_reqno.Text)) & "'"
   
        CrystalReport2.ReportFileName = "C:\DISREPORTSEC\DISCOMRP.RPT"
        CrystalReport2.Formulas(0) = "NAME = '" & Trim(txt_name.Text) & "'"
        CrystalReport2.Formulas(1) = "CASENO = '" & Trim(MaskEdBox1.FormattedText) & "'"
        CrystalReport2.Formulas(2) = "GRP = '" & Trim(CMB_BG.Text) & "'"
        CrystalReport2.Formulas(3) = "ent = '" & Trim(cmb_enterby.Text) & "'"
        CrystalReport2.Formulas(4) = "Dtmno = '" & Trim(txt_pathno.Text) & "'"
        CrystalReport2.Formulas(5) = "Udname = '" & Trim(txt_refdoc.Text) & "'"
        CrystalReport2.Formulas(6) = "Category = '" & Trim(txt_category.Text) & "'"
        wb1 = ""
        wb1 = Trim(txt_ward.Text) & " / " & Trim(txt_bedno.Text)
        CrystalReport2.Formulas(7) = "WB = '" & Trim(wb1) & "'"
        CrystalReport2.Formulas(8) = "Clindiag = '" & txt_relclnhis.Text & "'"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.Action = 1
Exit Sub
prtmsg: MsgBox Error(Err)
'txt_msg.Text = ""
btnenabl ("print")
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmd_save_Click()
Dim pcount As Integer
If Trim(cmb_enterby.Text) = "" Then
    MsgBox " ENTEREDBY FIELD IS EMPTY !!", vbInformation, "BLOOD PRODUCT"
    cmb_enterby.SetFocus
    Exit Sub
End If
'If Trim(cmb_finalby.Text) = "" Then
'    MsgBox " FINALBY FIELD IS EMPTY !!", vbInformation, "BLOOD PRODUCT"
'    cmb_finalby.SetFocus
'    Exit Sub
'End If
    pcount = 0
    For i = 1 To MSFlexGrid1.Rows - 1
        If Trim(MSFlexGrid1.TextMatrix(i, 1)) <> "" Then
          pcount = pcount + 1
        End If
    Next i
'    If pcount <> 0 Then
'      J = MsgBox(pcount & " Panic Value, Do you want to Save? ", vbYesNo)
'        If J = vbNo Then
'         Exit Sub
'        End If
'    End If

assign
btnenabl ("SAVE")
'cmd_print.Enabled = True
'cmd_cancel.Enabled = True


'commented by venkat
'If pcount = 0 Then
'    cmd_commit.Enabled = True
'Else
'    cmd_commit.Enabled = False
'End If
'cmd_commit.Enabled = True
MODFLAG = False
flagsave = 0
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
bloodproducts_report.Caption = "Dtm Lab"
MODFLAG = False
btnenabl ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR
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
'If Trim(usr) <> "" Then
    If sel_deptrs.RowCount > 0 Then
        sel_deptrs.MoveFirst
        While Not sel_deptrs.EOF
            cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
            sel_deptrs.MoveNext
        Wend
    End If
'Else
 '   MsgBox " Pls Enter User Name & Password !!!"
  '  Exit Sub
'End If
'    If seldissecentrflag1 = False Then
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
'
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
cmb_dept.SetFocus
cmb_modality.Enabled = True
MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True
MSFlexGrid1.clear
GRIDSTRING
'MSFlexGrid1.Rows = 3
'MSFlexGrid1.FormatString = "Sr.No. |^Com.Bl BagNo.|^Bag Seg No.|^  GRP  |^ EXP.DATE   |^  SAL    |^ ENZ   |^  IAT  |^  BY  |^  ON  |^     DATE      |^  TIME   |^ SIGN "
'MsfSet
txt_obvalue.Visible = False
TXT_BAGSEGNO.Visible = False
txt_rem.Visible = False
TXT_GRP.Visible = False
TXT_EXPDATE.Visible = False
TXT_SAL.Visible = False
txt_enz.Visible = False
TXT_IAT.Visible = False
txt_by.Visible = False
txt_on.Visible = False
TXT_CROSSDATE.Visible = False
TXT_CROSSTIME.Visible = False
txt_compcode.Visible = False
'txtusername.SetFocus
Opt_patient.Value = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo MSG:


NEWFLAG = False
viewflag = False
FLDENABL ("DIS")
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
MSFlexGrid1_Click
GRIDSTRING
'MSFlexGrid1.clear
'Set ConInterface = renv.OpenConnection(dsname:="biochem", Prompt:=rdDriverNo)
If BPREPORTFLAG = False Then
    BPREPORTFLAG = True
'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1BLOOD", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = Trim(LOGINUSRNAME)
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Enabled = False
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='NZZ' ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STR551", str)
'REPWTRN_QR(0) = Trim(USRDEPT)
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
If REPWTRNRS.RowCount > 0 Then
     txt_by.clear
'    'cmb_enteryby.clear
'    'cmb_addendumby.clear
'    'cmb_enterby.clear
'    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 1 Then
              txt_by.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        REPWTRNRS.MoveNext
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
    
'STR = "select * from DISLIB.Dislnkusr where username = ?"
'Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BLOOD", STR)
'CURUSRDET_QR(0) = "ABC"
'Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'
'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1BLOOD", STR)
'SELFINALBY_QR(0) = 6
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If


'*********SECURITY**********************************

''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR111", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG IS NULL)"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR111", str)

str = "SELECT MEMONO FROM RISLIB.RREQUISITN WHERE PAT_CATG<>'NC' AND REQNNO=?"
Set SELPATCATG_QR = rdb.CreateQuery("SELPATCATG_STR", str)

''********ICU***************
 
 
 
 str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
 Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO101_STR", str)
 

''dtm user
'STR = "select * from dislib.dtmuser"
'Set SELdtmuser_QR = rdb.CreateQuery("SELdtmuser_str11", STR)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set selsite_qr = rdb.CreateQuery("SELRECORDNO_STR", str)
    
''requisition date & time
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQDT_QR = rdb.CreateQuery("SELREQDT_STR", str)

''required date & time for BP
str = "SELECT * FROM DISLIB.BLOODPRODU WHERE REQNO=?"
Set SELREQDT1_QR = rdb.CreateQuery("SELREQDT1_STR", str)

''FOR DTMNO
str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
Set SELDTMNO1_QR = rdb.CreateQuery("SELDTMNO1_STR", str)


''for dept selection

str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr1BP", str)

'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
''******insert data into dcomrp table
str = "insert into DISLIB.dcomrp values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INSCOMRP_QR = rdb.CreateQuery("INSCOMRP_STR", str)

'******SELECT DATA FROM DCOMRP TABLE

str = "SELECT * FROM DISLIB.DCOMRP WHERE REQNO=?"
Set SELDCOMRP_QR = rdb.CreateQuery("SELDCOMRP_STR", str)

str = "select COUNT(*) REQNO from dislib.dcomrp where reqno=? and cbbn=? and compcode=?"
Set CHKDCOMRP_QR = rdb.CreateQuery("CHKDCOMRP_STR", str)

'*******UPDATE DCOMRP
str = "UPDATE DISLIB.DCOMRP SET BSN=?,GRP=?,EXPDT=?,SAL=?,ENZ=?,IAT=?,BY=?,ON=?,IDT=?,CTIME=?,USERNAME=?,FLD1=? WHERE REQNO=? AND CBBN=? AND COMPCODE=?"
Set UPDDCOMRP_QR = rdb.CreateQuery("UPDDCOMRP_STR", str)


'' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr1BP", str)



'    'for retriving reqnos
'''''****CHANGE
Set req_qr = rdb.CreateQuery("reqBP", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD='BP' ORDER BY REQNDT desc ") 'REQNSTAT='E'

'
Set SELTA101_QR = rdb.CreateQuery("selta101_qr11BP", "select * from PATOOLSLIB.ta101 where case_no =?")

'    ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1BP", str)

'    ' FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1BP", str)
''
'    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11BP", str)
'    selma104_qr(0) = "BN/00001"
'    Set selma104_rs = selma104_qr.OpenResultset(2, 3)
'    'for provisional diagnosis
'
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1BP", str)
'    selta1012_qr(0) = "BN/00001"
'    Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)


'''*****CHANGE
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1BP", str)

   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1BP", str)
'
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1BP", str)
'
'   STR = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
'   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1", STR)
'
'   STR = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
'   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11", STR)
'   SELDADDENDUM_QR(0) = "FSS"
'   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
   str = "select * from dislib.drequisitn where caseno=?"
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1BP", str)

    'commented and new query added by amarjeet on 05-11-2002 to select manjure's name in enter by
    str = "select * from dislib.dissec where (enterby='YES'AND MODALITY='BC') or (modality='EL' and ccname='Dr. S.S. Manjure')"
   ' STR = "select * from dislib.dissec where enterby='YES'AND MODALITY='BC'"
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1BP", str)
   
'   STR = "select * from dislib.dissec where PROVby='YES'AND MODALITY='SP'"
'   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1BC", STR)
'
   str = "select * from dislib.dissec where (FINALby='YES'AND MODALITY='BC')or (modality='EL' and ccname='Dr. S.S. Manjure')"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1BP", str)
    
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1BP", str)
    
  str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
  Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRBP", str)
    
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
  Set selMA101QR = rdb.CreateQuery("SELMA101QRBP", str)
  
'  STR = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET WHERE PRO_CODE=?"
'  Set SELDTEMPLETQR = rdb.CreateQuery("SELDTEMPLETQRBP", STR)
  
'  STR = "INSERT INTO DISLIB.DMPROCEDUR(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE) VALUES (?,?,?,?,?,?,?)"
'  Set INSDMPROCEDURQR = rdb.CreateQuery("INSDMPROCEDURQRBP", STR)
'
'  STR = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=? WHERE REQNO=? AND PRO_CODE=? "
'  Set UPDDMPROCEDURQR = rdb.CreateQuery("UPDDMPROCEDURQRBP", STR)
  
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_DET=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET=?"
  Set SELDMA101QR = rdb.CreateQuery("SELDMA101QRBP", str)
  
  str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1BP", str)
   
  'STR = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
'  STR = "SELECT PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
'  Set SELDPROCEDURQR = rdb.CreateQuery("SELDMPROCEDURQRBP", STR)
  
'  STR = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
'  Set SELDBCPROSUBQR = rdb.CreateQuery("SELDBCPROSUBQRBC", STR)
'
  str = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES(?,?)"
  Set INSDMBCCOMM_QR = rdb.CreateQuery("INSDMBCCOMMQRBC1", str)
'
  str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
  Set UPDDMBCCOMM_QR = rdb.CreateQuery("UPDDMBCCOMMQRBC1", str)
  
 str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
 ' STR = "SELECT * FROM DISLIB.DCOMRP WHERE REQNO=?"
  Set SELDMBCCOMM_QR = rdb.CreateQuery("SELDMBCCOMM_STR", str)
  
  str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrBC", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%BP%'"
  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRBC", str)
  
'  STR = "SELECT * FROM DISLIB.DBCALKPHS WHERE PROCODE=? AND AGE2=? AND SEX=? "
'  Set SELDBCALKPHSQR = rdb.CreateQuery("SELDBCALKPHSQRBC", STR)
'
'  sqlinterface = "select * from serial where reqno = ? and pro_code=?"
'  Set rqsqlinterface = ConInterface.CreateQuery("xxxx", sqlinterface)

  str = "SELECT * FROM DISLIB.BLOODPRODU WHERE REQNO=?"
  Set SELBLOODPRODUQR = rdb.CreateQuery("SELBLOODPRODUQRBPS", str)

  str = "UPDATE PATOOLSLIB.TA1012 SET PRVDIAGDES=? WHERE CASE_NO=?"
  Set UPDTA1012_QR = rdb.CreateQuery("UPDTA1012_QR11", str)

  str = "SELECT * FROM DISLIB.COMEXT WHERE REQNO=?"
  Set SELCOMEXT_QR = rdb.CreateQuery("SELCOMEXT_STR", str)
    
  str = "INSERT INTO DISLIB.COMEXT(REQNO,PLTCT) VALUES(?,?)"
  Set INSCOMEXT_QR = rdb.CreateQuery("INSCOMEXT_STR", str)
  
  str = "UPDATE DISLIB.COMEXT SET PLTCT=? WHERE REQNO=?"
  Set UPDCOMEXT_QR = rdb.CreateQuery("UPDCOMEXT_STR", str)
  
  str = "SELECT * FROM DISLIB.DTMUSER WHERE USERNAME=?"
    Set SELDTMUSER_QR = rdb.CreateQuery("SELDTMUSER_STR211", str)
  str = "select * from dislib.dtmuser"
    Set DTMUSER_QR = rdb.CreateQuery("DTMUSER_STR111", str)
End If

TXT_WB.Visible = False
TXT_PC.Visible = False
TXT_RDP.Visible = False
TXT_SDP.Visible = False
TXT_FFP.Visible = False
TXT_FVIIID.Visible = False
TXT_CRYO.Visible = False
TXT_ALB.Visible = False
'txt_obvalue.Visible = False

TXT_GRP.AddItem "A+"
TXT_GRP.AddItem "A-"
TXT_GRP.AddItem "B+"
TXT_GRP.AddItem "B-"
TXT_GRP.AddItem "O+"
TXT_GRP.AddItem "O-"
TXT_GRP.AddItem "AB+"
TXT_GRP.AddItem "AB-"
TXT_GRP.AddItem "Bombay"
TXT_GRP.AddItem " "

TXT_SAL.AddItem "N"
TXT_SAL.AddItem "P"
TXT_SAL.AddItem " "

TXT_IAT.AddItem "N"
TXT_IAT.AddItem "P"
TXT_IAT.AddItem " "

txt_enz.AddItem "N"
txt_enz.AddItem "P"
txt_enz.AddItem " "

txt_compcode.AddItem "WB"
txt_compcode.AddItem "PCS"
txt_compcode.AddItem "PC"
txt_compcode.AddItem "RDP"
txt_compcode.AddItem "FFP"
txt_compcode.AddItem "FVIIID"
txt_compcode.AddItem "CRYO"
txt_compcode.AddItem "ALB"
txt_compcode.AddItem "SDP"
txt_compcode.AddItem "IRR"
txt_compcode.AddItem "LEC"

txt_rem.AddItem "IR"
txt_rem.AddItem "LD"
txt_rem.AddItem "IR&LD"
txt_rem.AddItem "-"

'Set DTMUSER_RS = DTMUSER_QR.OpenResultset(1, 3)
'    If Not DTMUSER_RS.EOF Then
'        Do While Not DTMUSER_RS.EOF
'            txtusername.AddItem DTMUSER_RS("USERNAME")
'            txt_by.AddItem DTMUSER_RS("USERNAME")
'            cmb_enterby.AddItem DTMUSER_RS("username")
'            DTMUSER_RS.MoveNext
'        Loop
'    End If

btnenabl ("INI")
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
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
'        MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 05/12/2022
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
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = MaskEdBox1.FormattedText
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
            CASENO = MaskEdBox1.Text
        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
         If NEWFLAG = True Then
            If REQFLAG1 = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = MaskEdBox1.FormattedText
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
 Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo Errhandler
'If newflag = False Then
'    If MODFLAG = False Then
'        txt_obvalue.Visible = False
'        TXT_BAGSEGNO.Visible = False
'        TXT_GRP.Visible = False
'        TXT_EXPDATE.Visible = False
'        TXT_SAL.Visible = False
'        TXT_ENZ.Visible = False
'        TXT_IAT.Visible = False
'        txt_by.Visible = False
'        txt_on.Visible = False
'        TXT_CROSSDATE.Visible = False
'        TXT_CROSSTIME.Visible = False
'        txt_compcode.Visible = False
'        Exit Sub
'    End If
''End If
If (MSFlexGrid1.row < alreadyRows) And (flagsave = 1 Or flagsave = 2) Then
    txt_obvalue.Visible = True
    TXT_BAGSEGNO.Visible = True
    txt_rem.Visible = True
    TXT_GRP.Visible = True
    TXT_EXPDATE.Visible = True
    TXT_SAL.Visible = True
    txt_enz.Visible = True
    TXT_IAT.Visible = True
    txt_by.Visible = True
    txt_on.Visible = True
    TXT_CROSSDATE.Visible = True
    TXT_CROSSTIME.Visible = True
    txt_compcode.Visible = True
ElseIf (flagsave = 1 Or flagsave = 2) Then
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) <> "" Then
        txt_obvalue.Visible = True
        txt_obvalue.Enabled = False
        txt_compcode.Visible = True
        txt_compcode.Enabled = False
    Else
        txt_compcode.Visible = True
        txt_compcode.Enabled = True
        txt_obvalue.Visible = True
        txt_obvalue.Enabled = True
    End If
    TXT_BAGSEGNO.Visible = True
    txt_rem.Visible = True
    TXT_GRP.Visible = True
    TXT_EXPDATE.Visible = True
    TXT_SAL.Visible = True
    txt_enz.Visible = True
    TXT_IAT.Visible = True
    txt_by.Visible = True
    txt_on.Visible = True
    TXT_CROSSDATE.Visible = True
    TXT_CROSSTIME.Visible = True
    txt_compcode.Visible = True
    
End If
MSFlexGrid1.col = 1
txt_obvalue.Width = MSFlexGrid1.CellWidth
txt_obvalue.Height = MSFlexGrid1.CellHeight
txt_obvalue.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) <> "" Then
    txt_obvalue.Enabled = False
    If (flagsave = 2) Then
        txt_obvalue.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        txt_obvalue.Enabled = True
    Else
        txt_obvalue.Text = ""
    End If
Else
    txt_obvalue.Enabled = True
    If (flagsave = 2) Then
        txt_obvalue.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    Else
        txt_obvalue.Text = ""
    End If
End If

MSFlexGrid1.col = 2
txt_compcode.Width = MSFlexGrid1.CellWidth
'TXT_BAGSEGNO.Height = MSFlexGrid1.CellHeight
txt_compcode.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_compcode.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If (flagsave = 2) Then
    txt_compcode.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
Else
    txt_compcode.Text = ""
End If

MSFlexGrid1.col = 3
txt_rem.Width = MSFlexGrid1.CellWidth
'txt_rem.Height = MSFlexGrid1.CellHeight
txt_rem.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_rem.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If (flagsave = 2) Then
    txt_rem.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)
Else
    txt_rem.Text = ""
End If
'txt_rem.Text = ""


MSFlexGrid1.col = 4
TXT_BAGSEGNO.Width = MSFlexGrid1.CellWidth
'TXT_BAGSEGNO.Height = MSFlexGrid1.CellHeight
TXT_BAGSEGNO.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_BAGSEGNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

If (flagsave = 2) Then
    TXT_BAGSEGNO.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4)
Else
    TXT_BAGSEGNO.Text = ""
End If
'TXT_BAGSEGNO.Text = ""


MSFlexGrid1.col = 5
TXT_GRP.Width = MSFlexGrid1.CellWidth
'TXT_GRP.Height = MSFlexGrid1.CellHeight
TXT_GRP.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_GRP.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If (flagsave = 2) Then
    TXT_GRP.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5)
Else
    TXT_GRP.Text = ""
End If
'TXT_GRP.Text = ""

MSFlexGrid1.col = 6
TXT_EXPDATE.Width = MSFlexGrid1.CellWidth
'TXT_EXPDATE.Height = MSFlexGrid1.CellHeight
TXT_EXPDATE.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_EXPDATE.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = "" Then
    TXT_EXPDATE.Text = "__/__/____"
Else
    TXT_EXPDATE.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6)
End If

MSFlexGrid1.col = 7
TXT_SAL.Width = MSFlexGrid1.CellWidth
'txt_compcode.Height = MSFlexGrid1.CellHeight
TXT_SAL.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_SAL.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If (flagsave = 2) Then
    TXT_SAL.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
Else
    TXT_SAL.Text = ""
End If
'TXT_SAL.Text = ""

MSFlexGrid1.col = 8
txt_enz.Width = MSFlexGrid1.CellWidth
'TXT_SAL.Height = MSFlexGrid1.CellHeight
txt_enz.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_enz.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If (flagsave = 2) Then
    txt_enz.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8)
Else
    txt_enz.Text = ""
End If
'txt_enz.Text = ""

MSFlexGrid1.col = 9
TXT_IAT.Width = MSFlexGrid1.CellWidth
'txt_IAT.Height = MSFlexGrid1.CellHeight
TXT_IAT.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_IAT.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If (flagsave = 2) Then
    TXT_IAT.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 9)
Else
    TXT_IAT.Text = ""
End If
'TXT_IAT.Text = ""

MSFlexGrid1.col = 10
txt_on.Width = MSFlexGrid1.CellWidth
'txt_on.Height = MSFlexGrid1.CellHeight
txt_on.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_on.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'txt_on.Text = ""
If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) = "" Then
    txt_on.Text = "__/__/____"
Else
    txt_on.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10)
End If

MSFlexGrid1.col = 11
txt_by.Width = MSFlexGrid1.CellWidth
'txt_by.Height = MSFlexGrid1.CellHeight
txt_by.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_by.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If (flagsave = 2) Then
    txt_by.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 11)
Else
    txt_by.Text = ""
End If
'txt_by.Text = ""


MSFlexGrid1.col = 12
TXT_CROSSDATE.Width = MSFlexGrid1.CellWidth
'TXT_CROSSDATE.Height = MSFlexGrid1.CellHeight
TXT_CROSSDATE.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_CROSSDATE.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 12) = "" Then
    TXT_CROSSDATE.Text = "__/__/____"
Else
    TXT_CROSSDATE.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 12)
End If

MSFlexGrid1.col = 13
TXT_CROSSTIME.Width = MSFlexGrid1.CellWidth
'TXT_CROSSTIME.Height = MSFlexGrid1.CellHeight
TXT_CROSSTIME.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_CROSSTIME.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'TXT_CROSSTIME.Text = ""
If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13) = "" Then
    TXT_CROSSTIME.Text = "__:__"
Else
    TXT_CROSSTIME.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13)
    'TXT_CROSSTIME.Text = mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11), 1, 5)
End If

If ((flagsave = 1) Or (flagsave = 2)) And (MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 0) <> "") Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = CStr(Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 0)) + 1)
'    If txt_obvalue.Visible = True Then
'       txt_obvalue.SetFocus
'    End If
End If
    'ADD BY SUNDEEP
    'TXT_SAL.Text = "N"
    'TXT_SAL_Click
    'txt_enz.Text = "N"
    'txt_enz_Click
    'TXT_IAT.Text = "N"
    'TXT_IAT_Click
    
    'txt_on = DATE11
    TXT_CROSSDATE = DATE11
    ''ADD BY SUNDEEP
Exit Sub
Errhandler:
MsgBox Err.Description, vbInformation, "BLOOD PRODUCT"
Exit Sub
End Sub

Private Sub MSFlexGrid1_GotFocus()
'rowno = 1
'MSFlexGrid1.Row = i
End Sub

Private Sub MSFlexGrid1_LeaveCell()
preMsfrow = MSFlexGrid1.row
End Sub

Private Sub MSFlexGrid1_LostFocus()
'If Trim(txt_obvalue.Text) <> "" Then
'    i = i + 1
'End If
End Sub

'Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'If MSFlexGrid1.Col = 1 Then
'    MSFlexGrid1.ToolTipText = "COMPATIBLE BAG NO."
'ElseIf MSFlexGrid1.Col = 2 Then
'    MSFlexGrid1.ToolTipText = "COMPATIBLE SEGMENT NO."
'ElseIf MSFlexGrid1.Col > 2 Then
'    MSFlexGrid1.ToolTipText = ""
'End If
'End Sub

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

    lbl_caseno.Caption = "  CC No."
    MaskEdBox1.Mask = "########"
    MaskEdBox1.Enabled = True
    'MaskEdBox1.SetFocus
End If
End Sub

Private Sub txt_observed_LostFocus()
'If KeyAscii = 13 Then
'    If SELDMA101FLAG = False Then
'        SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
'        Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
'        SELDMA101FLAG = True
'    Else
'        SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
'        SELDMA101RS.MoreResults
'        SELDMA101RS.Requery
'    End If
'    If SELDMA101RS.RowCount > 0 Then
''        If SELDTEMPLETFLAG = False Then
''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''            SELDTEMPLETFLAG = True
''        Else
''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''            SELDTEMPLETRS.MoreResults
''            SELDTEMPLETRS.Requery
''        End If
'
'    End If
'End If
'        If SELDTEMPLETRS.RowCount > 0 Then
'            If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
'                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'                If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC"))) Then
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
'                Else
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'                End If
'                For i = 1 To 1000
'                Beep
'                Next i
'            Else
'                 If Val(Trim(txt_observed.Text)) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'                    MSFlexGrid1.CellBackColor = vbWhite
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'                    If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
'                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
'                    Else
'                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'                    End If
'                 Else
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
'                 End If
'                 If Val(Trim(txt_observed.Text)) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'                    MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'                        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
'                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
'                        Else
'                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'                        End If
'
'                 End If
'                 If Trim(txt_observed.Text) = "" Then
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'
'                 End If
'            End If
'        End If
'    End If
'
'    i = MSFlexGrid1.Rows - 1
'    If i = MSFlexGrid1.Row Then
'        cmd_save.SetFocus
'        Exit Sub
'    End If
'    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
'            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
'    End If
'    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'            txt_observed.Text = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'
'            txt_observed.Visible = True
'            txt_observed.SetFocus
'
'            txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'
'    End If
'End If
End Sub

Public Sub assign()
On Error GoTo MSG:
rdb.BeginTrans
    If flagsave = 1 Then
'         For i = 1 To MSFlexGrid1.Rows - 1
'
'
'
'         Next i
''
            updrrequisitn_qr(0) = Trim(cmb_enterby.Text)
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = Null 'Trim(cmb_finalby.Text)
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = Null 'Trim(cmb_finalby.Text)
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
            UPDTA1012_QR(0) = Trim(txt_relclnhis.Text)
            UPDTA1012_QR(1) = Trim(MaskEdBox1.FormattedText)
            UPDTA1012_QR.Execute
            
            INSDMBCCOMM_QR(0) = Trim(req_chglen(cmb_reqno.Text))
            If Len(Trim(txt_comments.Text)) > 0 Then
                INSDMBCCOMM_QR(1) = Null
                INSDMBCCOMM_QR(1) = txt_comments.Text
            Else
                INSDMBCCOMM_QR(1) = Null
            
            End If
            INSDMBCCOMM_QR.Execute
            
'            INSCOMEXT_QR(0) = Trim(req_chglen(cmb_reqno.Text))
'            If Len(Trim(TXT_PLTC.Text)) > 0 Then
'                INSCOMEXT_QR(1) = Null
'                INSCOMEXT_QR(1) = TXT_PLTC.Text
'            Else
'                INSCOMEXT_QR(1) = Null
'            End If
'            INSCOMEXT_QR.Execute
            
            
            i = 1
            While (MSFlexGrid1.TextMatrix(i, 1) <> "" And MSFlexGrid1.TextMatrix(i, 2) <> "")
                i = i + 1
            Wend
            NROWS = i - 1
            For i = 1 To NROWS
                INSCOMRP_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 1))) > 0 Then
                    INSCOMRP_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 1))
                Else
                    INSCOMRP_QR(1) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 4))) > 0 Then
                    INSCOMRP_QR(2) = Trim(MSFlexGrid1.TextMatrix(i, 4))
                Else
                    INSCOMRP_QR(2) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 5))) > 0 Then
                    INSCOMRP_QR(3) = Trim(MSFlexGrid1.TextMatrix(i, 5))
                Else
                    INSCOMRP_QR(3) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 0 Then
                    INSCOMRP_QR(4) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 6))))
                Else
                    INSCOMRP_QR(4) = Null
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 7))) > 0 Then
                    INSCOMRP_QR(5) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                Else
                    INSCOMRP_QR(5) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 8))) > 0 Then
                    INSCOMRP_QR(6) = Trim(MSFlexGrid1.TextMatrix(i, 8))
                Else
                    INSCOMRP_QR(6) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 9))) > 0 Then
                    INSCOMRP_QR(7) = Trim(MSFlexGrid1.TextMatrix(i, 9))
                Else
                    INSCOMRP_QR(7) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 11))) > 0 Then
                    INSCOMRP_QR(8) = Trim(MSFlexGrid1.TextMatrix(i, 11))
                Else
                    INSCOMRP_QR(8) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 10))) > 0 Then
                    INSCOMRP_QR(9) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 10))))
                Else
                    INSCOMRP_QR(9) = Null
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 12))) > 0 Then
                    INSCOMRP_QR(10) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 12))))
                Else
                    INSCOMRP_QR(10) = Null
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 13))) > 0 Then
                    INSCOMRP_QR(11) = Trim(MSFlexGrid1.TextMatrix(i, 13))
                Else
                    INSCOMRP_QR(11) = " "
                End If
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 2))) > 0 Then
                    INSCOMRP_QR(12) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                Else
                    INSCOMRP_QR(12) = " "
                End If
                INSCOMRP_QR(13) = Trim(usr)
                If Len(Trim(MSFlexGrid1.TextMatrix(i, 3))) > 0 Then
                    INSCOMRP_QR(14) = Trim(MSFlexGrid1.TextMatrix(i, 3))
                Else
                    INSCOMRP_QR(14) = " "
                End If
'INSERTING MAXNO 16-03-04
Dim TEMPSCHNO As Integer
Dim TEMP As rdoResultset
Dim str As String
str = "select MAX(fld1) from dislib.dcomrp where reqno='" & cmb_reqno & "'"
Set TEMP = rdb.OpenResultset(str)
If Not TEMP.AbsolutePosition = -1 Then
    If IsNull((TEMP(0))) = False Then
        TEMPSCHNO = Val(TEMP(0)) + 1
    Else
        TEMPSCHNO = 1

    End If
End If

                INSCOMRP_QR(13) = TEMPSCHNO
''''INSERTING MAX NO
                INSCOMRP_QR.Execute
            Next i
    End If
    If flagsave = 2 Then
         
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = Null 'cmb_finalby.Text
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = Null 'cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
            UPDTA1012_QR(0) = Trim(txt_relclnhis.Text)
            UPDTA1012_QR(1) = Trim(MaskEdBox1.FormattedText)
            UPDTA1012_QR.Execute
            
            If Len(Trim(txt_comments.Text)) > 0 Then
                UPDDMBCCOMM_QR(0) = Null
                UPDDMBCCOMM_QR(0) = Trim(txt_comments.Text)
            Else
                UPDDMBCCOMM_QR(0) = Null
            End If
            UPDDMBCCOMM_QR(1) = Trim(req_chglen(cmb_reqno.Text))
            UPDDMBCCOMM_QR.Execute
            
'            If Len(Trim(TXT_PLTC.Text)) > 0 Then
'                UPDCOMEXT_QR(0) = Null
'                UPDCOMEXT_QR(0) = TXT_PLTC.Text
'            Else
'                UPDCOMEXT_QR(0) = Null
'            End If
'            UPDCOMEXT_QR(1) = Trim(req_chglen(cmb_reqno.Text))
'            UPDCOMEXT_QR.Execute
            
            i = 1
            While MSFlexGrid1.TextMatrix(i, 1) <> "" And MSFlexGrid1.TextMatrix(i, 2) <> ""
                i = i + 1
            Wend
            NROWS = i - 1
            For i = 1 To NROWS
                If CHKDCOMRP_FLAG = False Then
                    CHKDCOMRP_QR(0) = Trim(req_chglen(cmb_reqno.Text))
                    CHKDCOMRP_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 1))
                    CHKDCOMRP_QR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                    Set CHKDCOMRP_RS = CHKDCOMRP_QR.OpenResultset(1, 3)
                    CHKDCOMRP_FLAG = True
                Else
                    CHKDCOMRP_QR(0) = Trim(req_chglen(cmb_reqno.Text))
                    CHKDCOMRP_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 1))
                    CHKDCOMRP_QR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                    CHKDCOMRP_RS.MoreResults
                    CHKDCOMRP_RS.Requery
                End If
                If CHKDCOMRP_RS("REQNO") > 0 Then
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 4))) > 0 Then
                        UPDDCOMRP_QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 4))
                    Else
                        UPDDCOMRP_QR(0) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 5))) > 0 Then
                        UPDDCOMRP_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 5))
                    Else
                        UPDDCOMRP_QR(1) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 0 Then
                        UPDDCOMRP_QR(2) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 6))))
                    Else
                        UPDDCOMRP_QR(2) = Null
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 7))) > 0 Then
                        UPDDCOMRP_QR(3) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    Else
                        UPDDCOMRP_QR(3) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 8))) > 0 Then
                        UPDDCOMRP_QR(4) = Trim(MSFlexGrid1.TextMatrix(i, 8))
                    Else
                        UPDDCOMRP_QR(4) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 9))) > 0 Then
                        UPDDCOMRP_QR(5) = Trim(MSFlexGrid1.TextMatrix(i, 9))
                    Else
                        UPDDCOMRP_QR(5) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 11))) > 0 Then
                        UPDDCOMRP_QR(6) = Trim(MSFlexGrid1.TextMatrix(i, 11))
                    Else
                        UPDDCOMRP_QR(6) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 10))) > 0 Then
                        UPDDCOMRP_QR(7) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 10))))
                    Else
                        UPDDCOMRP_QR(7) = Null
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 12))) > 0 Then
                        UPDDCOMRP_QR(8) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 12))))
                    Else
                        UPDDCOMRP_QR(8) = Null
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 13))) > 0 Then
                        UPDDCOMRP_QR(9) = Trim(MSFlexGrid1.TextMatrix(i, 13))
                    Else
                        UPDDCOMRP_QR(9) = " "
                    End If
                    j = NROWS 'here1
                 '   MsgBox J
                    If j = i Then
                            UPDDCOMRP_QR(10) = 1
                        Else
                            UPDDCOMRP_QR(10) = Trim(usr)
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 3))) > 0 Then
                        UPDDCOMRP_QR(11) = Trim(MSFlexGrid1.TextMatrix(i, 3))
                    Else
                        UPDDCOMRP_QR(11) = " "
                    End If
                    UPDDCOMRP_QR(12) = req_chglen(Trim(cmb_reqno.Text))
                    UPDDCOMRP_QR(13) = Trim(MSFlexGrid1.TextMatrix(i, 1))
                    UPDDCOMRP_QR(14) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                    UPDDCOMRP_QR.Execute
                   ' i = MSFlexGrid1.Rows 'H1
                   ' STR = "UPDATE DISLIB.DCOMRP SET USERNAME=1"
                    'Set UPDDCOMRP_QR1 = rdb.CreateQuery("UPDDCOMRP_STR", STR)

                Else
                    INSCOMRP_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 1))) > 0 Then
                        INSCOMRP_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 1))
                    Else
                        INSCOMRP_QR(1) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 4))) > 0 Then
                        INSCOMRP_QR(2) = Trim(MSFlexGrid1.TextMatrix(i, 4))
                    Else
                        INSCOMRP_QR(2) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 5))) > 0 Then
                        INSCOMRP_QR(3) = Trim(MSFlexGrid1.TextMatrix(i, 5))
                    Else
                        INSCOMRP_QR(3) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 0 Then
                        INSCOMRP_QR(4) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 6))))
                    Else
                        INSCOMRP_QR(4) = Null
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 7))) > 0 Then
                        INSCOMRP_QR(5) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    Else
                        INSCOMRP_QR(5) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 8))) > 0 Then
                        INSCOMRP_QR(6) = Trim(MSFlexGrid1.TextMatrix(i, 8))
                    Else
                        INSCOMRP_QR(6) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 9))) > 0 Then
                        INSCOMRP_QR(7) = Trim(MSFlexGrid1.TextMatrix(i, 9))
                    Else
                        INSCOMRP_QR(7) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 11))) > 0 Then
                        INSCOMRP_QR(8) = Trim(MSFlexGrid1.TextMatrix(i, 11))
                    Else
                        INSCOMRP_QR(8) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 10))) > 0 Then
                        INSCOMRP_QR(9) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 10))))
                    Else
                        INSCOMRP_QR(9) = Null
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 12))) > 0 Then
                        INSCOMRP_QR(10) = GET_REVDATE(CStr(Trim(MSFlexGrid1.TextMatrix(i, 12))))
                    Else
                        INSCOMRP_QR(10) = Null
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 13))) > 0 Then
                        INSCOMRP_QR(11) = Trim(MSFlexGrid1.TextMatrix(i, 13))
                    Else
                        INSCOMRP_QR(11) = " "
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 2))) > 0 Then
                        INSCOMRP_QR(12) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                    Else
                        INSCOMRP_QR(12) = " "
                    End If
                    INSCOMRP_QR(13) = Trim(usr)
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 3))) > 0 Then
                        INSCOMRP_QR(14) = Trim(MSFlexGrid1.TextMatrix(i, 3))
                    Else
                        INSCOMRP_QR(14) = " "
                    End If
                    
                  
''INSERTING MAXNO 16-03-04


str = "select MAX(FLD1) from dislib.dcomrp where reqno='" & cmb_reqno & "'"
Set TEMP = rdb.OpenResultset(str)
If Not TEMP.AbsolutePosition = -1 Then
    If IsNull((TEMP(0))) = False Then
        TEMPSCHNO = Val(TEMP(0)) + 1
    Else
        TEMPSCHNO = 1

    End If
End If



                INSCOMRP_QR(13) = Val(TEMPSCHNO)

                    
           ' MAX NO
                    
                    INSCOMRP_QR.Execute
                End If
            Next i
        End If
rdb.CommitTrans
MsgBox "Record is Saved"
btnenabl ("INI")
'txt_obvalue.Visible = False
Exit Sub
MSG:
    rdb.RollbackTrans
MsgBox Err.Description
End Sub

Public Sub STAFF()
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
                If Trim(REFDR_RS("patNstat")) = "IP" Then
                    txt_ward.Visible = True
                    txt_bedno.Visible = True
                    lbl_ward.Visible = True
                    lbl_bed.Visible = True
                    txt_ward.Text = selpathdrequisitn_rs("WARD_NAME")
                    txt_bedno.Text = selpathdrequisitn_rs("bedno")
                    txt_status.Text = "IP"
                Else
                    txt_status.Text = "ST"
                    txt_ward.Visible = False
                    txt_bedno.Visible = False
                    lbl_ward.Visible = False
                    lbl_bed.Visible = False
                End If
            End If
          
         
'            If SELPATHDREQUISITNFALG1 = False Then
'                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
'                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
'                SELPATHDREQUISITNFALG1 = True
'            Else
'                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
'                selpathdrequisitn_rs.MoreResults
'                selpathdrequisitn_rs.Requery
'            End If
'            If selpathdrequisitn_rs.RowCount > 0 Then
'                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
'            End If
            If SELDTMNO1_FLAG = False Then
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                Set SELDTMNO1_RS = SELDTMNO1_QR.OpenResultset(1, 3)
                SELDTMNO1_FALG = True
            Else
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                SELDTMNO1_RS.MoreResults
                SELDTMNO1_RS.Requery
            End If
            If SELDTMNO1_RS.RowCount > 0 Then
                txt_pathno.Text = SELDTMNO1_RS("DTMNO")
            End If


'            If SELTA1012FLAG1 = False Then
'                selta1012_qr(0) = MaskEdBox1.FormattedText
'                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'                SELTA1012FLAG1 = True
'            Else
'                selta1012_qr(0) = MaskEdBox1.FormattedText
'                selta1012_rs.MoreResults
'                selta1012_rs.Requery
'            End If
'            If selta1012_rs.RowCount > 0 Then
'                'txt_procedure.Text = selta1012_rs("prvdiagdes")
'            End If
           
           
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

        If seldissecfinflag = False Then
            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
            seldissecfinflag1 = True
        Else
            seldissecfin1_rs.MoreResults
            seldissecfin1_rs.Requery
        End If
        If seldissecfin1_rs.RowCount > 0 Then
            cmb_finalby.clear
            While Not seldissecfin1_rs.EOF
                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
                seldissecfin1_rs.MoveNext
            Wend
        End If
            
            '******
          
        
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
                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                'cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    'cmb_finalby.TEXT=""
                Else
                    'cmb_finalby.Text = REFDR_RS("FREPORTBY")
                End If
            End If
            
            
            
          
'            If SELMA104FLAG1 = False Then
'                selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
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
                If Trim(REFDR_RS("patNstat")) = "IP" Then
                    txt_ward.Visible = True
                    txt_bedno.Visible = True
                    lbl_ward.Visible = True
                    lbl_bed.Visible = True
                    txt_ward.Text = selpathdrequisitn_rs("WARD_NAME")
                    txt_bedno.Text = selpathdrequisitn_rs("bed_no")
                    txt_status.Text = "IP"
                Else
                    txt_status.Text = "OP"
                    txt_ward.Visible = False
                    txt_bedno.Visible = False
                    lbl_ward.Visible = False
                    lbl_bed.Visible = False
                    txt_status.Text = "ST"
                End If
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
                'txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
            End If
            If SELDTMNO1_FLAG = False Then
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                Set SELDTMNO1_RS = SELDTMNO1_QR.OpenResultset(1, 3)
                SELDTMNO1_FALG = True
            Else
                SELDTMNO1_QR(0) = Trim(MaskEdBox1.FormattedText)
                SELDTMNO1_RS.MoreResults
                SELDTMNO1_RS.Requery
            End If
            If SELDTMNO1_RS.RowCount > 0 Then
                txt_pathno.Text = SELDTMNO1_RS("DTMNO")
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
            ' ********************************************
            
             If SELBLOODPRODUFLAG = False Then
                SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                Set SELBLOODPRODURS = SELBLOODPRODUQR.OpenResultset(1, 3)
                SELBLOODPRODUFLAG = True
            Else
                SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                SELBLOODPRODURS.MoreResults
                SELBLOODPRODURS.Requery
            End If
                    If SELBLOODPRODURS.RowCount > 0 Then
                        CMB_BG.Text = " " & Trim(SELBLOODPRODURS("BG"))
                        CMB_RFT.Text = "" & Trim(SELBLOODPRODURS("RFT"))
                        TXT_HB.Text = "" & Trim(SELBLOODPRODURS("HB"))
                        CMB_PR.Text = "" & Trim(SELBLOODPRODURS("FLD2"))
                        TXT_NOOFUREP.Text = "" & Trim(SELBLOODPRODURS("NOOFUNREP"))
                        'TXT_PLTC.Text = "" & Trim(SELBLOODPRODURS("FLD3"))
'                        TXT_DT.Text = " " & Format(Trim(SELBLOODPRODURS("REQUION")), "DD/MM/YYYY")
'                        TXT_TM.Text = " " & SELBLOODPRODURS("TIME")
                        'cmb_tr.Text = "" & SELBLOODPRODURS("TYPEOFREQ")
                        If Len(SELBLOODPRODURS("WB")) > 0 Then
                            CHK_WB.Value = 1
                            TXT_WB.Text = " " & SELBLOODPRODURS("WB")
                        Else
                            CHK_WB.Value = 0
                            TXT_WB.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("PC")) > 0 Then
                            CHK_PC.Value = 1
                            TXT_PC.Text = " " & SELBLOODPRODURS("PC")
                        Else
                            CHK_PC.Value = 0
                            TXT_PC.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("RDP")) > 0 Then
                            CHK_RDP.Value = 1
                            TXT_RDP.Text = " " & SELBLOODPRODURS("RDP")
                        Else
                            CHK_RDP.Value = 0
                            TXT_RDP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("SDP")) > 0 Then
                            CHK_SDP.Value = 1
                            TXT_SDP.Text = " " & SELBLOODPRODURS("SDP")
                        Else
                            CHK_SDP.Value = 0
                            TXT_SDP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("FFP")) > 0 Then
                            CHK_FFP.Value = 1
                            TXT_FFP.Text = " " & SELBLOODPRODURS("FFP")
                        Else
                            CHK_FFP.Value = 0
                           TXT_FFP.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("FVIIID")) > 0 Then
                            CHK_FVIIID.Value = 1
                            TXT_FVIIID.Text = " " & SELBLOODPRODURS("FVIIID")
                        Else
                            CHK_FVIIID.Value = 0
                            TXT_FVIIID.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("CRYO")) > 0 Then
                            CHK_CRYO.Value = 1
                            TXT_CRYO.Text = " " & SELBLOODPRODURS("CRYO")
                        Else
                            CHK_CRYO.Value = 0
                            TXT_CRYO.Text = " "
                        End If
                        If Len(SELBLOODPRODURS("ALB")) > 0 Then
                            CHK_ALB.Value = 1
                            TXT_ALB.Text = " " & SELBLOODPRODURS("ALB")
                        Else
                            CHK_ALB.Value = 0
                            TXT_ALB.Text = " "
                        End If
                        
                        'TXT_OTHERS.Text = " " & SELBLOODPRODURS("OTHERS")
                        TXT_PBNAME.Text = "" & Trim(SELBLOODPRODURS("PHOBOTMIST"))
                        TXT_NOUREPLACE = " " & Trim(SELBLOODPRODURS("NOOFUNREP"))
                        If Len(SELBLOODPRODURS("SR2")) > 0 Then
                            CHK_LEUCODEPLETION.Value = 1
                        Else
                            CHK_LEUCODEPLETION.Value = 0
                        End If
                        If Len(SELBLOODPRODURS("SR1")) > 0 Then
                            CHK_IRRADIATION.Value = 1
                        Else
                            CHK_IRRADIATION.Value = 0
                        End If
                        TXT_PLTC.Text = "" & SELBLOODPRODURS("FLD3")

            
           
           '****
                If SELDCOMRP_FLAG = False Then
            SELDCOMRP_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDCOMRP_RS = SELDCOMRP_QR.OpenResultset(1, 3)
            SELDCOMRP_FLAG = True
        Else
            SELDCOMRP_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDCOMRP_RS.MoreResults
            SELDCOMRP_RS.Requery
        End If
        
        If SELDMBCCOMM_FLAG = False Then
            SELDMBCCOMM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDMBCCOMM_RS = SELDMBCCOMM_QR.OpenResultset(1, 3)
            SELDMBCCOMM_FLAG = True
        Else
            SELDMBCCOMM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDMBCCOMM_RS.MoreResults
            SELDMBCCOMM_RS.Requery
        End If
        If Not SELDMBCCOMM_RS.EOF Then
            txt_comments.Text = "" & Trim(SELDMBCCOMM_RS("COMMENTS"))
        End If
        
        Dim i
        i = 1
        If Not SELDCOMRP_RS.EOF Then
            Do While Not SELDCOMRP_RS.EOF
                MSFlexGrid1.row = i
                MSFlexGrid1.TextMatrix(i, 0) = i
                MSFlexGrid1.TextMatrix(i, 1) = "" & Trim(SELDCOMRP_RS("CBBN"))
                MSFlexGrid1.TextMatrix(i, 2) = "" & Trim(SELDCOMRP_RS("compcode"))
                MSFlexGrid1.TextMatrix(i, 3) = "" & Trim(SELDCOMRP_RS("FLD1"))
                MSFlexGrid1.TextMatrix(i, 4) = "" & Trim(SELDCOMRP_RS("BSN"))
                MSFlexGrid1.TextMatrix(i, 5) = "" & Trim(SELDCOMRP_RS("GRP"))
                MSFlexGrid1.TextMatrix(i, 6) = "" & Trim(SELDCOMRP_RS("EXPDT"))
                MSFlexGrid1.TextMatrix(i, 7) = "" & Trim(SELDCOMRP_RS("SAL"))
                MSFlexGrid1.TextMatrix(i, 8) = "" & Trim(SELDCOMRP_RS("ENZ"))
                MSFlexGrid1.TextMatrix(i, 9) = "" & Trim(SELDCOMRP_RS("IAT"))
                MSFlexGrid1.TextMatrix(i, 10) = "" & Trim(SELDCOMRP_RS("ON"))
                MSFlexGrid1.TextMatrix(i, 11) = "" & Trim(SELDCOMRP_RS("BY"))
                MSFlexGrid1.TextMatrix(i, 12) = "" & Trim(SELDCOMRP_RS("IDT"))
                MSFlexGrid1.TextMatrix(i, 13) = "" & Trim(SELDCOMRP_RS("CTIME"))
                SELDCOMRP_RS.MoveNext
                i = i + 1
            Loop
        End If
        End If
            
            '**********************************************
        End If
End Sub

Private Sub TXT_BAGSEGNO_Change()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = Trim(TXT_BAGSEGNO.Text)
End Sub

Private Sub TXT_BY_Change()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 11) = Trim(txt_by.Text)
End Sub

Private Sub txt_by_Click()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 11) = Trim(txt_by.Text)
End Sub

Private Sub txt_by_LostFocus()
    TXT_CROSSDATE = DATE11
End Sub

Private Sub txt_compcode_Change()
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = Trim(txt_compcode.Text)
End Sub

Private Sub txt_compcode_Click()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_compcode.Text)
End Sub

Private Sub txt_compcode_LostFocus()
    For z1 = 0 To txt_rem.ListCount - 1
        If Trim(txt_rem.List(z1)) = "-" Then
            txt_rem.Text = txt_rem.List(z1)
        End If
    Next z1
    txt_rem_Click
End Sub

Private Sub TXT_CROSSDATE_Change()
If TXT_CROSSDATE.Text <> "__/__/____" Then
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 12) = TXT_CROSSDATE.Text
Else
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 12) = ""
End If
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = Trim(TXT_CROSSDATE.Text)
End Sub


Private Sub TXT_CROSSDATE_GotFocus()
TXT_CROSSDATE.Text = DATE11
End Sub

Private Sub TXT_CROSSTIME_Change()
If TXT_CROSSTIME.Text <> "__:__" Then
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13) = TXT_CROSSTIME.Text
Else
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13) = ""
End If
'MSFlexGrid1_Click
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11) = Trim(TXT_CROSSTIME.Text)
End Sub


Private Sub TXT_CROSSTIME_GotFocus()
    TXT_CROSSTIME.Text = Format(time1, "hh:mm")
End Sub

Private Sub TXT_CROSSTIME_LostFocus()
Dim a
a = InStr(1, Trim(TXT_CROSSTIME.Text), ":")
b = Mid(Trim(TXT_CROSSTIME.Text), 1, a - 1)
If Len(TXT_CROSSTIME.Text) < 5 Then
    MsgBox " Please Enter a suitable format !!"
    Exit Sub
End If

'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''MsfSet
'TXT_CROSSDATE.Mask = ""

End Sub

Private Sub TXT_ENZ_Change()
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Trim(txt_enz.Text)
End Sub



Private Sub txt_enz_Click()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = Trim(txt_enz.Text)
End Sub

Private Sub txt_enz_LostFocus()
    TXT_IAT.Text = "N"
    TXT_IAT_Click
End Sub

Private Sub TXT_EXPDATE_Change()
If TXT_EXPDATE.Text <> "__/__/____" Then
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = TXT_EXPDATE.Text
Else
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = ""
End If
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Trim(TXT_EXPDATE.Text)
End Sub


Private Sub TXT_EXPDATE_LostFocus()
    TXT_SAL.Text = "N"
    TXT_SAL_Click
End Sub

Private Sub TXT_GRP_Change()
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Trim(TXT_GRP.Text)
End Sub



Private Sub TXT_GRP_Click()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = Trim(TXT_GRP.Text)
End Sub

Private Sub TXT_GRP_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TXT_IAT_Change()
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = Trim(TXT_IAT.Text)
End Sub





Private Sub TXT_IAT_Click()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 9) = Trim(TXT_IAT.Text)
End Sub

Private Sub TXT_IAT_LostFocus()
    txt_on = DATE11
End Sub

'Private Sub txt_obvalue_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 38 Then
'    If MSFlexGrid1.Row <> 1 Then
'    MSFlexGrid1.Row = MSFlexGrid1.Row - 1
'    txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'    txt_obvalue = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
'    'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
'
'    End If
'End If
'If KeyCode = 40 Then
'    If MSFlexGrid1.Row <> MSFlexGrid1.Rows - 1 Then
'        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'        txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'        txt_obvalue = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
'    End If
'End If
'If KeyCode = 13 Then
'    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = txt_obvalue.Text
'    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
'            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
'    End If
'    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'            txt_obvalue.Text = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'
'            txt_obvalue.Visible = True
'            txt_obvalue.SetFocus
'
'            txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'
'    End If
'End If
'End Sub
Private Sub txt_obvalue_Change()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) = Trim(txt_obvalue.Text)
End Sub

Private Sub txt_obvalue_GotFocus()
'If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) <> "" Then
'    txt_obvalue.Enabled = False
'Else
'    txt_obvalue.Enabled = True
'End If
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = CStr(Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 0)) + 1)
'MSFlexGrid1.Rows = MSFlexGrid1.Row + 1
'MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
End Sub

Private Sub txt_obvalue_KeyPress(KeyAscii As Integer)
KeyAscii = char_valid(KeyAscii)
End Sub

Private Sub txt_obvalue_LostFocus()
If (NEWFLAG = True Or viewflag = True) Then
    If Trim(txt_obvalue.Text) = "" Then
        MsgBox " COM. BAG NO. CAN'T BE EMPTY !!", vbInformation, "COMPATIBILITY REPORT"
        Exit Sub
    End If
End If
If ((NEWFLAG = True) Or (viewflag = True)) And (MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 0) <> "") Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = CStr(Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 0)) + 1)
End If
End Sub

Private Sub TXT_ON_Change()
If txt_on.Text <> "__/__/____" Then
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) = txt_on.Text
Else
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) = ""
End If
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = Trim(txt_on.Text)
End Sub

Private Sub TXT_ON_LostFocus()
If ((NEWFLAG = True) Or (viewflag = True)) And (MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 0) <> "") Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = CStr(Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 0)) + 1)
End If
End Sub

Private Sub txt_rem_Click()
If Trim(txt_rem.Text) <> "" Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = Trim(txt_rem.Text)
End If
End Sub

Private Sub TXT_SAL_Change()
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = Trim(TXT_SAL.Text)
End Sub

Private Sub MsfSet()

'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
MSFlexGrid1.col = 1
txt_obvalue.Width = MSFlexGrid1.CellWidth
txt_obvalue.Height = MSFlexGrid1.CellHeight
txt_obvalue.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_obvalue.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
txt_obvalue.Text = ""

MSFlexGrid1.col = 2
txt_compcode.Width = MSFlexGrid1.CellWidth
'TXT_BAGSEGNO.Height = MSFlexGrid1.CellHeight
txt_compcode.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_compcode.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
txt_compcode.Text = ""

MSFlexGrid1.col = 3
txt_rem.Width = MSFlexGrid1.CellWidth
'txt_rem.Height = MSFlexGrid1.CellHeight
txt_rem.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_rem.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
txt_rem.Text = ""



MSFlexGrid1.col = 4
TXT_BAGSEGNO.Width = MSFlexGrid1.CellWidth
TXT_BAGSEGNO.Height = MSFlexGrid1.CellHeight
TXT_BAGSEGNO.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_BAGSEGNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TXT_BAGSEGNO.Text = ""


MSFlexGrid1.col = 5
TXT_GRP.Width = MSFlexGrid1.CellWidth
'TXT_GRP.Height = MSFlexGrid1.CellHeight
TXT_GRP.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_GRP.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TXT_GRP.Text = ""

MSFlexGrid1.col = 6
TXT_EXPDATE.Width = MSFlexGrid1.CellWidth
'TXT_EXPDATE.Height = MSFlexGrid1.CellHeight
TXT_EXPDATE.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_EXPDATE.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TXT_EXPDATE.Mask = ""
TXT_EXPDATE.Mask = "##/##/####"


MSFlexGrid1.col = 7
TXT_SAL.Width = MSFlexGrid1.CellWidth
TXT_SAL.Height = MSFlexGrid1.CellHeight
TXT_SAL.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_SAL.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TXT_SAL.Text = ""

MSFlexGrid1.col = 6
txt_enz.Width = MSFlexGrid1.CellWidth
txt_enz.Height = MSFlexGrid1.CellHeight
txt_enz.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_enz.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
txt_enz.Text = ""

MSFlexGrid1.col = 7
TXT_IAT.Width = MSFlexGrid1.CellWidth
TXT_IAT.Height = MSFlexGrid1.CellHeight
TXT_IAT.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_IAT.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TXT_IAT.Text = ""

MSFlexGrid1.col = 8
txt_by.Width = MSFlexGrid1.CellWidth
txt_by.Height = MSFlexGrid1.CellHeight
txt_by.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_by.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
txt_by.Text = ""

MSFlexGrid1.col = 9
txt_on.Width = MSFlexGrid1.CellWidth
txt_on.Height = MSFlexGrid1.CellHeight
txt_on.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_on.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
txt_on.Text = ""

MSFlexGrid1.col = 10
TXT_CROSSDATE.Width = MSFlexGrid1.CellWidth
'TXT_CROSSDATE.Height = MSFlexGrid1.CellHeight
TXT_CROSSDATE.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_CROSSDATE.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TXT_CROSSDATE.Mask = ""
TXT_CROSSDATE.Text = DATE11
'TXT_CROSSDATE.Mask = "##/##/####"

MSFlexGrid1.col = 11
TXT_CROSSTIME.Width = MSFlexGrid1.CellWidth
'TXT_CROSSTIME.Height = MSFlexGrid1.CellHeight
TXT_CROSSTIME.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TXT_CROSSTIME.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TXT_CROSSTIME.Text = ""

End Sub

Private Sub TXT_SAL_Click()
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(TXT_SAL.Text)
End Sub

Private Sub TXT_SAL_LostFocus()
    txt_enz.Text = "N"
    txt_enz_Click
End Sub

Private Sub txt_time_LostFocus()
'MsfSet
End Sub
Private Sub GRIDSTRING()
MSFlexGrid1.FormatString = "SN|^ UnitNo.|^CompCo|^Sp.Pr |^     SEGNo.     |^ GRP  |^ Exp. Date   |^  S   |^   E  |^   I   |^     ON        |^     BY     |^    DATE     |^  TIME"
End Sub
Private Sub TXT_COMMENTS_GotFocus()
txt_comments.clear
'''TXT_COMMENTS.AddItem "Rechecked"
'''TXT_COMMENTS.AddItem "Report from E lab"
'''TXT_COMMENTS.AddItem "No Sample"
'''TXT_COMMENTS.AddItem "Sample ? Pl. repeat the sample."
'''TXT_COMMENTS.AddItem "Machine not working"
'''TXT_COMMENTS.AddItem "Serum haemolysed"
'''TXT_COMMENTS.AddItem "Serum higly haemolysed"
'''TXT_COMMENTS.AddItem "Serum lipaemic"
'''TXT_COMMENTS.AddItem "Serum highly lipaemic"
'''TXT_COMMENTS.AddItem "Reading out of range. Pl repeat the sample"
'''TXT_COMMENTS.AddItem "Electrolytes reported from E-Lab"
End Sub



Private Sub txtcompcode_Change()

End Sub

Private Sub txtPassword_LostFocus()
'If SELDTMUSER_FLAG = False Then
'    SELDTMUSER_QR(0) = Trim(txtusername.Text)
'    Set SELDTMUSER_RS = SELDTMUSER_QR.OpenResultset(1, 3)
'    SELDTMUSER_FLAG = True
'Else
'    SELDTMUSER_QR(0) = Trim(txtusername.Text)
'    SELDTMUSER_RS.MoreResults
'    SELDTMUSER_RS.Requery
'End If
'If Not SELDTMUSER_RS.EOF Then
'    If Trim(SELDTMUSER_RS("password")) <> Trim(txtpassword.Text) Then
'        MsgBox " U r not an authorized person to log in !!", vbInformation, "DTM LAB"
'        txtpassword.Text = ""
'        txtusername.Text = ""
'        btnenabl ("INI")
'        CMD_NEW.SetFocus
'        Exit Sub
'    Else
'        usr = txtusername.Text
'        bloodproducts_report.Caption = "Dtm Lab : User Name : " & usr
'        txtusername.Text = ""
'        txtpassword.Text = ""
'        If NEWFLAG = True Then
'            cmb_enterby.Text = usr
'        End If
'        Dim ind_cbodept As Integer
'        If Trim(cmb_dept) = "" Then
'            For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
'                If Trim(cmb_dept.List(ind_cbodept)) = "TRANSFUSION MEDICINE$NZZ" Then
'                    cmb_dept.ListIndex = ind_cbodept
'                End If
'            Next ind_cbodept
'        End If
'    End If
'End If
End Sub

Private Sub txtusername_LostFocus()
'If Trim(txtusername.Text) <> "" Then
'    txtpassword.SetFocus
'End If
End Sub
