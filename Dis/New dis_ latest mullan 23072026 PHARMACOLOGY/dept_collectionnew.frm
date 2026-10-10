VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form dept_collectionnew 
   Caption         =   "Department Sample Receiving  Centre"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   0
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   10200
      TabIndex        =   23
      Top             =   840
      Width           =   1215
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
      Left            =   10200
      TabIndex        =   22
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame buttons 
      Height          =   735
      Left            =   0
      TabIndex        =   20
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton CMD_REFRESH 
         Caption         =   "&Refresh"
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
         Left            =   10320
         TabIndex        =   35
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
         Left            =   8880
         MouseIcon       =   "dept_collectionnew.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         Left            =   7440
         MouseIcon       =   "dept_collectionnew.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   1215
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
         Left            =   6000
         MouseIcon       =   "dept_collectionnew.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   240
         Width           =   1215
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
         Left            =   4560
         MouseIcon       =   "dept_collectionnew.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1215
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
         MouseIcon       =   "dept_collectionnew.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   240
         Width           =   1215
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
         MouseIcon       =   "dept_collectionnew.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Default         =   -1  'True
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
         MouseIcon       =   "dept_collectionnew.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   0
      TabIndex        =   15
      Top             =   2880
      Width           =   11775
      Begin VB.Frame cy_frame 
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
         Height          =   1245
         Left            =   0
         TabIndex        =   41
         Top             =   3600
         Width           =   11745
         Begin VB.ComboBox cmb_sex 
            Height          =   315
            ItemData        =   "dept_collectionnew.frx":1546
            Left            =   3840
            List            =   "dept_collectionnew.frx":1550
            TabIndex        =   58
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_age 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1455
            TabIndex        =   57
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox txt_normal 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1455
            TabIndex        =   54
            Top             =   525
            Width           =   10185
         End
         Begin VB.ComboBox cmb_recename 
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
            ItemData        =   "dept_collectionnew.frx":1562
            Left            =   8820
            List            =   "dept_collectionnew.frx":1572
            TabIndex        =   52
            Top             =   870
            Width           =   2820
         End
         Begin MSMask.MaskEdBox MSK_CTIME 
            Height          =   270
            Left            =   11025
            TabIndex        =   51
            Top             =   180
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_cdate 
            Height          =   315
            Left            =   8805
            TabIndex        =   50
            Top             =   135
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.CheckBox chk_ok 
            Caption         =   "OK"
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
            Left            =   6510
            TabIndex        =   46
            Top             =   255
            Width           =   585
         End
         Begin VB.CheckBox chk_leaked 
            Caption         =   "Leaked"
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
            Left            =   4920
            TabIndex        =   45
            Top             =   180
            Width           =   1215
         End
         Begin VB.CheckBox chk_clotted 
            Caption         =   "Clotted"
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
            Left            =   3345
            TabIndex        =   44
            Top             =   255
            Width           =   1200
         End
         Begin VB.CheckBox chk_hyper 
            Caption         =   "Hypercellular"
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
            Left            =   1455
            TabIndex        =   43
            Top             =   210
            Width           =   1515
         End
         Begin VB.CheckBox chk_dilute 
            Caption         =   "Dilute"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   42
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label6 
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
            Height          =   270
            Left            =   3420
            TabIndex        =   56
            Top             =   900
            Width           =   375
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Left            =   1005
            TabIndex        =   55
            Top             =   885
            Width           =   390
         End
         Begin VB.Label Label4 
            Caption         =   "Normal Control"
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
            Left            =   135
            TabIndex        =   53
            Top             =   570
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Col.Date"
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
            Left            =   7935
            TabIndex        =   49
            Top             =   150
            Width           =   765
         End
         Begin VB.Label Label2 
            Caption         =   "Col.Time"
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
            Left            =   10200
            TabIndex        =   48
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Receivers Name"
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
            Left            =   7335
            TabIndex        =   47
            Top             =   960
            Width           =   1500
         End
      End
      Begin VB.CommandButton PRINTBARCODE 
         Caption         =   "Print_Barcode"
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
         Left            =   4560
         TabIndex        =   40
         Top             =   4350
         Width           =   1575
      End
      Begin VB.CommandButton cmd_VPP 
         Caption         =   "VPP"
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
         Left            =   10020
         TabIndex        =   37
         Top             =   4350
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CMB_PRINT1 
         Caption         =   "&Print"
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
         Height          =   375
         Left            =   165
         TabIndex        =   36
         Top             =   4290
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_pathno 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   10200
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chk_collected 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Collected"
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
         Left            =   6600
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chk_rcvd 
         BackColor       =   &H8000000E&
         Caption         =   "Received"
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
         Left            =   8265
         TabIndex        =   7
         Top             =   735
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3495
         Left            =   45
         TabIndex        =   6
         Top             =   135
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   3
         Cols            =   7
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   192
         AllowUserResizing=   3
         FormatString    =   $"dept_collectionnew.frx":15BC
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
      Begin VB.Label LBL_VPP 
         AutoSize        =   -1  'True
         Caption         =   "Letter for Vpp"
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
         Height          =   195
         Left            =   8685
         TabIndex        =   38
         Top             =   4440
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   11535
      Begin VB.TextBox TXT_NAME 
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
         Left            =   6360
         TabIndex        =   33
         Top             =   1200
         Width           =   4455
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   1080
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
      Begin VB.Frame frm_patientstaff 
         Height          =   735
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1335
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
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opt_patient 
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
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.OptionButton opt_caseno 
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
         Height          =   255
         Left            =   1680
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton opt_all 
         Caption         =   "All"
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
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_modality 
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
         Left            =   6360
         TabIndex        =   2
         Top             =   240
         Width           =   3135
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
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin MSMask.MaskEdBox msk_todt 
         Height          =   375
         Left            =   9480
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
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
      Begin MSMask.MaskEdBox msk_fromdt 
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   4210752
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
      Begin VB.Label LBL_NAME 
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
         Left            =   5640
         TabIndex        =   32
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbl_todate 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
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
         Left            =   8640
         TabIndex        =   19
         Top             =   855
         Width           =   705
      End
      Begin VB.Label lbl_mod 
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
         Left            =   5520
         TabIndex        =   18
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_fromdate 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
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
         TabIndex        =   17
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Department"
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
         TabIndex        =   16
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Label Label8 
      Caption         =   "This form is not in use (Its only for VZZ)"
      Height          =   375
      Left            =   600
      TabIndex        =   59
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITAL "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11655
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
      Left            =   9600
      TabIndex        =   25
      Top             =   840
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
      Left            =   9600
      TabIndex        =   24
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      Caption         =   "DEPARTMENT SAMPLE RECEIVING  CENTRE"
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
      Left            =   2640
      TabIndex        =   13
      Top             =   840
      Width           =   6405
   End
End
Attribute VB_Name = "dept_collectionnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELMA106_QR As rdoQuery
Dim SELMA106_RS As rdoResultset
Dim SELMA106FLAG As Boolean
Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset
Dim SELRMODALITYFLAG As Boolean
Dim deptcollectionflag As Boolean
Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim REC_DATE1, REC_TIME1
Dim SELRREQUISITN_QR1 As rdoQuery
Dim SELRREQUISITN_RS1 As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean
'INTERFACE
'ADDED BY AMIT FOR RTIMELOG
Dim INSTIME_QR As rdoQuery
Dim INSTIME_RS As rdoResultset

Dim UPDATE_CYTOG As rdoQuery
Dim PROVCY_FLG As Boolean

Public R1, C1, II
Dim BARCODE_NEW As Boolean
Dim BARCODE_VIEW As Boolean
Dim SELLABNO_QR As rdoQuery
Dim SELLABNO_RS As rdoResultset
Public REQNO As Variant
Public CASENO As Variant
Dim dept As String
Dim MODALITY As String
Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
Dim SELMA101FLAG As Boolean
Dim updrrequisitn_qr As rdoQuery
Dim MODIFYFLAG As Boolean
Dim SELRREQNO_QR As rdoQuery
Dim SELRREQNO_RS As rdoResultset
Dim SELRREQNOFLAG As Boolean

Dim SELRREQNO_QR1 As rdoQuery
Dim SELRREQNO_RS1 As rdoResultset
Dim SELRREQNOFLAG1 As Boolean


Dim SELTA1011_QR As rdoQuery
Dim SELTA1011_RS As rdoResultset
Dim SELTA1011FLAG As Boolean
Dim SELDREQSTAFFD1_QR As rdoQuery
Dim SELDREQSTAFFD1_RS As rdoResultset
Dim SELDREQSTAFFD1FLAG As Boolean
Dim upddrequisitn As rdoQuery
Dim SELDREQUISITNQR As rdoQuery
Dim SELDREQUISITNRS As rdoResultset
Dim seldequisitnflag As Boolean
Dim SELDREQQR As rdoQuery
Dim SELDREQRS As rdoResultset
Dim SELDREQFLAG As Boolean
Dim SELREQNOQR As rdoQuery
Dim SELREQNORS As rdoResultset
Dim SELREQNORSFLAG As Boolean
Dim DEPTMODALITY As String
Dim INSDBCINFACEQR As rdoQuery

Dim SELDMULPROCQR As rdoQuery
Dim seldmulprocrs As rdoResultset
Dim seldmulprocflag As Boolean

Dim SELELPROSUBQR As rdoQuery
Dim SELELPROSUBRS As rdoResultset
Dim SELELPROSUBFLAG As Boolean
Dim srow As Integer

''DTMNO

Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_FLAG As Boolean

''STAFF
Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFO_FLAG As Boolean


'''RECEIPT TIME
Dim UPDLATERALITS_QR As rdoQuery

Dim SELDAUTOPATH_QR As rdoQuery
Dim SELDAUTOPATH_RS As rdoResultset
Dim SELDAUTOPATHFLAG As Boolean

Dim DPATHYEAR_QR As rdoQuery
Dim DPATHYEAR_RS As rdoResultset
Dim DPATHYEARFLAG As Boolean

Dim UPDDAUTOPATH_QR As rdoQuery
Dim INSDAUTOPATH_QR As rdoQuery

Dim SELRREQUISITN_QR11 As rdoQuery
Dim SELRREQUISITN_RS11 As rdoResultset
Dim SELRREQUISITN_FLAG11 As Boolean

Dim SELRREQNO_QR11 As rdoQuery
Dim SELRREQNO_RS11 As rdoResultset
Dim SELRREQNO_FLAG11 As Boolean
'SAVE RECEIVING DATE AND TIME IN RTIMELOG BY AMIT
Dim UPDTIME_QR As rdoQuery
Dim UPDTIME_RS As rdoResultset
Dim SELTIME_QR As rdoQuery
Dim SELTIME_RS As rdoResultset

Dim pathno As String

''*********tracking System*************
Dim Ins_Spathology_qr As rdoQuery
''*********tracking System*************
'Public CASENO
Public PATNAME
Public COLL_CENTER
Public REC_DATE
Public REC_TIME

Dim SELDREQUISITN_PATH_QR As rdoQuery
Dim SELDREQUISITN_PATH_RS As rdoResultset
Dim SELDREQUISITN_PATH_flg As Boolean

''****INTERFACE CODE BY AMIT *****************
Dim INSDINTERFACE_QR As rdoQuery
Dim INSBINTERFACE_QR As rdoQuery
Dim RS_DMULPROCE  As rdoResultset

Dim seldmulproce_qr As rdoQuery
Dim seldmulproce_rs As rdoResultset
Dim SELDMULPROCE_FLAG As Boolean
Dim scrlck_flg As Boolean
Dim SELSAMPLECODE_QR As rdoQuery
Dim SELSAMPLECODE_RS As rdoResultset

'' HEMATO INTERFACE
Dim INSHINTERFACE_QR As rdoQuery
Dim INSCINTERFACE_QR As rdoQuery
Dim SELDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset
'Dim seldmulproce_qr As rdoQuery
'Dim seldmulproce_rs As rdoResultset
Dim seldhlprosub_qr As rdoQuery
Dim seldhlprosub_rs As rdoResultset

'' HEMATO INTERFACE

''****INTERFACE CODE BY AMIT AGRAWAL*****************
' tamankar begins
Dim interexpcd_qr As rdoQuery
Dim interexpxd_rs As rdoResultset
Dim INTERFACEDATA_QR0 As rdoQuery
Dim INTERFACEDATA_RS0 As rdoResultset
Dim INTERFACEDATA_QR1 As rdoQuery
Dim INTERFACEDATA_RS1 As rdoResultset
Dim SELDBCPROSUB_QR As rdoQuery
Dim SELDBCPROSUB_RS As rdoResultset



Private Sub chk_clotted_Click()
  If chk_clotted.Value = 1 Then
    PROVCY_FLG = True
  Else
    PROVCY_FLG = False
  End If
  
End Sub

Private Sub chk_collected_Click()
'MSFlexGrid1.Col = 4
'If chk_collected.Value = 1 Then
'    MSFlexGrid1.Text = "COLLECTED"
'Else
'    MSFlexGrid1.Text = ""
'End If

End Sub

Private Sub chk_collected_KeyPress(KeyAscii As Integer)
On Error GoTo MSG:
'If KeyAscii = 13 Then
'If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
'        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
'End If
'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'
'        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'        chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4))
'        chk_collected.SetFocus
'        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'
'End If
'End If
'Exit Sub
MSG:
'MsgBox Err.Description
End Sub

Private Sub chk_dilute_Click()
 If chk_dilute.Value = 1 Then
  PROVCY_FLG = True
 Else
  PROVCY_FLG = False
 End If
End Sub

Private Sub chk_hyper_Click()
 If chk_hyper.Value = 1 Then
   PROVCY_FLG = True
 Else
   PROVCY_FLG = False
 End If
End Sub

Private Sub chk_leaked_Click()
 If chk_leaked.Value = 1 Then
  PROVCY_FLG = True
 Else
  PROVCY_FLG = False
 End If
End Sub

Private Sub chk_ok_Click()
 If chk_ok.Value = 1 Then
   PROVCY_FLG = True
 Else
   PROVCY_FLG = False
 End If
End Sub

Private Sub chk_rcvd_Click()
Dim num As String
'If dept = "NZZ" Then
'    MSFlexGrid1.Col = 5
'    If chk_rcvd.Value = 1 Then
'        MSFlexGrid1.Text = "RECEIVED"
'
'    Else
'        MSFlexGrid1.Text = ""
'    End If
'    Exit Sub
'End If


''If MODALITY = "TM" And chk_rcvd.Value = 1 Then

''   'justin 23-08-2004
''    'i = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1), 11, 6)
''    'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "TM/" & i
''    num = Format(Now, "yymmdd")
''    SELDREQUISITN_PATH_QR(0) = num & "%"
''
''    If SELDREQUISITN_PATH_flg = False Then
''       Set SELDREQUISITN_PATH_RS = SELDREQUISITN_PATH_QR.OpenResultset(2, 3)
''       SELDREQUISITN_PATH_flg = True
''
''       If IsNull(SELDREQUISITN_PATH_RS(0)) = False Then
''          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = SELDREQUISITN_PATH_RS(0) + 1
''          pathno = SELDREQUISITN_PATH_RS(0) + 1
''       Else
''          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = num & "001"
''          pathno = num & "001"
''       End If
''    Else
''     pathno = "0" & Val(pathno) + 1
''     MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = pathno
''    End If
''
''  'JUSTIN
    

If Trim(MODALITY) = "TM" And chk_rcvd.Value = 0 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = " "
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = " "
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = ""
End If

MSFlexGrid1.col = 5
If chk_rcvd.Value = 1 Then
    MSFlexGrid1.Text = "RECEIVED"
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = "COLLECTED"
'chk_collected.Value = 1
Else
    MSFlexGrid1.Text = ""
End If

If Trim(dept) = "FZZ" And Trim(MODALITY) = "SP" And chk_rcvd.Value = 1 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = "RECEIVED"
    'SPAUTONO
ElseIf Trim(dept) = "FZZ" And Trim(MODALITY) = "SP" And chk_rcvd.Value = 0 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = ""
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
End If

If Trim(dept) = "UZZ" And chk_rcvd.Value = 1 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = "RECEIVED"
    'SPAUTONO
ElseIf Trim(dept) = "UZZ" Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = ""
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
End If

End Sub



Private Sub chk_rcvd_KeyPress(KeyAscii As Integer)
On Error GoTo MSG:
If KeyAscii = 13 Then
If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
End If
If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then

        MSFlexGrid1.row = MSFlexGrid1.row + 1
        chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
        chk_rcvd.SetFocus
        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

End If
End If


Exit Sub
MSG:
MsgBox Err.Description
End Sub

'Private Sub chk_rcvd_LostFocus()
'txt_pathno.SetFocus
'End Sub

Private Sub cmb_dept_Click()
Dim i As Integer
On Error GoTo MSG:
i = InStr(Trim(cmb_dept.Text), "$")
dept = Mid(cmb_dept.Text, i + 1, 4)
If SELRMODALITYFLAG = False Then
    
    SELRMODALITY_QR(0) = Trim(dept)
    'DEPT = Trim(Mid(Trim(cmb_dept.Text), i + 1, 4))
    Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(1, 3)
Else
    SELRMODALITY_QR(0) = Trim(dept)
    SELRMODALITY_RS.MoreResults
    SELRMODALITY_RS.Requery
End If
If SELRMODALITY_RS.RowCount > 0 Then
    cmb_modality.clear
    While Not SELRMODALITY_RS.EOF
        cmb_modality.AddItem SELRMODALITY_RS("modlitydes") & "$" & SELRMODALITY_RS("modalitycd")
        SELRMODALITY_RS.MoveNext
    Wend
End If
If dept = "NZZ" Then
    cmb_modality.AddItem "ALL"
    
End If
cmb_modality.Enabled = True
cmb_modality.SetFocus

If dept = "FEC" And viewflag = True Then
  cmd_modify.Enabled = True
End If

Exit Sub
MSG:
MsgBox Err.Description
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
MsgBox Error(Err)
End Sub

Private Sub cmb_modality_Click()
Dim i As Integer
i = InStr(cmb_modality, "$")
MODALITY = Mid(cmb_modality, i + 1, 4)
''If modality = "IH" Then
''dept = "FDZ"
''End If
''msk_fromdt.Text = Format(DATE11, "dd/mm/yyyy")
''msk_todt.Text = Format(DATE11, "dd/mm/yyyy")
''msk_fromdt.SetFocus
MSFlexGrid1.FormatString = " S.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    Path No. "
If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
    MSFlexGrid1.FormatString = " S.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    Lab No.   "
End If
If MODALITY = "SS" Or MODALITY = "BP" Or MODALITY = "HT" Or MODALITY = "ALL" Then
    MSFlexGrid1.FormatString = " S.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    DTM NO."
End If
If MODALITY = "BM" Then
    MSFlexGrid1.FormatString = " S.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    BMA No. "
End If
If MODALITY = "HL" Or MODALITY = "BF" Then
    MSFlexGrid1.FormatString = " S.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    HEM No. "
End If

If MODALITY = "CC" Or MODALITY = "FI" Then
   MSFlexGrid1.FormatString = " S.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    Lab No.   "
   cy_frame.Visible = True
   Call clearcy
   lbl_todate.Visible = False
   lbl_fromdate.Visible = False
   msk_fromdt.Visible = False
   msk_todt.Visible = False
   OPT_ALL.Visible = False
Else
   cy_frame.Visible = False
   lbl_todate.Visible = True
   lbl_fromdate.Visible = True
   msk_fromdt.Visible = True
   msk_todt.Visible = True
   OPT_ALL.Visible = True
End If
End Sub

Private Sub CMB_PRINT1_Click()
If Trim(msk_fromdt.Text) = "__/__/____" And Trim(msk_todt.Text) = "__/__/____" Then
    MsgBox "Enter From and To date", vbInformation
    Exit Sub
End If
        var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
        var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
    
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONN.RPT"
        'Cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
        CrystalReport1.SelectionFormula = "{RREQUISITN.EXAMDT} >= Date(" & var_date_f & ") AND {RREQUISITN.EXAMDT} <= Date(" & var_date_t & ") and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "'"
        CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
        CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
        CrystalReport1.Formulas(2) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
End Sub

Private Sub cmd_back_Click()
chk_rcvd.Value = 0
chk_rcvd.Visible = False
MSFlexGrid1.Rows = 1
BARCODE_NEW = False
BARCODE_VIEW = False
'chk_collected.Value = 0
'chk_rcvd.Value = 0
'chk_collected.Visible = False
'chk_rcvd.Visible = False
txt_pathno.Visible = False
msk_caseno.Mask = ""
msk_caseno.Text = ""
'msk_caseno.Mask = "??>/#####"
msk_caseno.Enabled = False
msk_fromdt.Mask = ""
msk_fromdt.Text = ""
msk_fromdt.Mask = "##/##/####"
msk_fromdt.Enabled = False
msk_todt.Mask = ""
msk_todt.Text = ""
msk_todt.Mask = "##/##/####"
msk_todt.Enabled = False
cmb_dept.ListIndex = -1
cmb_dept.Enabled = False
cmb_modality.ListIndex = -1
cmb_modality.Enabled = False
Opt_patient.Value = False
opt_staff.Value = False
OPT_ALL.Value = False
OPT_CASENO.Value = False
btnenabl ("INI")
Me.MousePointer = 11
DIS_MAIN.show
Me.Hide
Me.MousePointer = 0
End Sub

Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    ' codded for modify option
    If Trim(dept) = "GZZ" Or Trim(dept) = "FEC" Then
        cmd_modify.Enabled = True
    Else
        cmd_modify.Enabled = False
    End If
    ' codded for modify option
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_view.Enabled = False
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_view.Enabled = True
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = True
ElseIf cap = "MODIFY" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = False
End If


End Sub

Private Sub cmd_cancel_Click()
NEWFLAG = False
clearcy
MSFlexGrid1.Rows = 1
'chk_rcvd.Value = 0
BARCODE_NEW = False
BARCODE_VIEW = False
chk_rcvd.Visible = False
cmb_dept.Text = ""
cmb_modality.Text = ""
msk_fromdt.Mask = ""
msk_fromdt.Text = ""
txt_name.Text = ""
msk_fromdt.Mask = "##/##/####"
msk_todt.Mask = ""
msk_todt.Text = ""
msk_todt.Mask = "##/##/####"
OPT_ALL.Value = False
OPT_ALL.Enabled = False
OPT_CASENO.Value = False
OPT_CASENO.Enabled = False
opt_staff.Value = False
opt_staff.Enabled = False
Opt_patient.Value = False
Opt_patient.Enabled = False
msk_caseno.Mask = ""
msk_caseno.Text = ""
'msk_caseno.Mask = "??>/#####"
txt_pathno.Text = ""
msk_caseno.Enabled = False
txt_pathno.Visible = False
MODIFYFLAG = False
'chk_collected.Visible = False
btnenabl ("INI")
FLDENABL ("DIS")
End Sub

Private Sub cmd_modify_Click()
btnenabl ("NEW")
'chk_rcvd.Visible = True
txt_pathno.Visible = True
txt_pathno.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6)
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
MODIFYFLAG = True
' codded for modify option
If dept = "GZZ" Then
    If viewflag = True Then
        MSFlexGrid1_Click
        txt_pathno.Visible = True
    End If
End If
' codded for modify option
End Sub

Private Sub cmd_new_Click()
On Error GoTo MSG:
chk_rcvd.Value = 0
BARCODE_NEW = True
If SELMA106FLAG = False Then
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106FLAG = True
Else
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106_RS.MoreResults
    SELMA106_RS.Requery
    
End If

If SELMA106_RS.RowCount > 0 Then
    cmb_dept.clear
    While Not SELMA106_RS.EOF
        'If Trim(USRDEPT) <> "ALL" Then
         '   If Trim(USRDEPT) = Trim(SELMA106_RS("DEPT_CODE")) Then
          '      cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
           ' End If
        'Else
             cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
        'End If
        SELMA106_RS.MoveNext
    Wend
End If
cmb_dept.Text = cmb_dept.List(0)
    If Trim(cmb_dept.Text) = Trim("PATHOLOGY$FZZ") Then
        cmb_dept.AddItem Trim("PREVENTIVE ONCOLOGY$IZZ")
    End If
If Trim(cmb_dept.Text) <> "" Then
    cmb_dept_Click
End If
MODIFYFLAG = False
OPT_ALL.Enabled = True
OPT_CASENO.Enabled = True
Opt_patient.Enabled = True
opt_staff.Enabled = True
txt_name.Text = ""
chk_rcvd.Visible = False
btnenabl ("NEW")
FLDENABL ("ENA")
NEWFLAG = True
viewflag = False
MSFlexGrid1.Rows = 2
MSFlexGrid1.row = 1

'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
'opt_patient.Value = True
cmb_dept.SetFocus

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
Dim var_date_f
Dim var_date_t
Dim str1 As String
Dim str As String
str = "C"
R1 = MSFlexGrid1.row
C1 = MSFlexGrid1.col
On Error GoTo MSG:
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
     If OPT_ALL.Value = True Then
        var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
        var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
        If Right(Trim(cmb_modality.Text), 2) = "SS" Or Right(Trim(cmb_modality.Text), 2) = "BP" Or Right(Trim(cmb_modality.Text), 2) = "HT" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONDTM.RPT"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTION.RPT"
        End If
        'Cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
        CrystalReport1.SelectionFormula = "{RREQUISITN.EXAMDT} >= Date(" & var_date_f & ") AND {RREQUISITN.EXAMDT} <= Date(" & var_date_t & ") and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODALITYCD}= '" & MODALITY & "' "
        CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
        CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
        CrystalReport1.Formulas(2) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
     End If
     
     If OPT_CASENO.Value = True Then
        If Right(Trim(cmb_modality.Text), 2) = "SS" Or Right(Trim(cmb_modality.Text), 2) = "HT" Or Right(Trim(cmb_modality.Text), 2) = "BP" Then
            str1 = msk_caseno.FormattedText
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONCDTM.RPT"
            'CRYSTALREPORT1.SelectionFormula = " {RREQUISITN.CASENO} ='" & msk_caseno.FormattedText & "' and {RREQUISITN.DEPT_CODE}= '" & DEPT & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "' AND ({RREQUISITN.REQNSTAT}= 'E' OR {RREQUISITN.REQNSTAT}= 'C') "
            CrystalReport1.SelectionFormula = " {RREQUISITN.CASENO} ='" & msk_caseno.FormattedText & "' and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "' AND ({RREQUISITN.REQNSTAT}= 'E') "
            CrystalReport1.Formulas(0) = "FROMDT = '" & msk_caseno.FormattedText & "'"
            CrystalReport1.Formulas(1) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
        ElseIf Right(Trim(cmb_modality.Text), 2) = "CC" Or Right(Trim(cmb_modality.Text), 2) = "FI" Then
            
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONCY.RPT"
            
            CrystalReport1.SelectionFormula = " {RREQUISITN.REQNNO} ='" & req_chglen(MSFlexGrid1.TextMatrix(R1, 1)) & "' and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "' AND ({RREQUISITN.REQNSTAT}= 'E') "
      
            If Right(Trim(cmb_modality.Text), 2) = "CC" Then
            CrystalReport1.Formulas(1) = "MODALITY ='CONVENTIONAL CYTOGENETICS'"
            Else
            CrystalReport1.Formulas(1) = "MODALITY ='MOLECULAR CYTOGENETICS(FISH) REPORT'"
            End If
        Else
            str1 = msk_caseno.FormattedText
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONC.RPT"
            CrystalReport1.SelectionFormula = " {RREQUISITN.CASENO} ='" & msk_caseno.FormattedText & "' and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "' AND ({RREQUISITN.REQNSTAT}= 'E' OR {RREQUISITN.REQNSTAT}= 'C')"
            CrystalReport1.Formulas(0) = "FROMDT = '" & msk_caseno.FormattedText & "'"
            CrystalReport1.Formulas(1) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
        End If
     End If
     
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    
btnenabl ("PRINT")
Exit Sub
MSG:
 MsgBox Err.Description

End Sub

Private Sub cmd_refresh_Click()
msk_caseno.Mask = ""
msk_caseno.Text = ""
'msk_caseno.Mask = "??>/#####"
NEWFLAG = True
MSFlexGrid1.Rows = 1
txt_pathno.Text = ""
chk_rcvd.Visible = False
txt_pathno.Visible = False
btnenabl ("NEW")
MODIFYFLAG = False
'chk_rcvd.Visible = False
If msk_caseno.Enabled = True Then
msk_caseno.SetFocus
End If
End Sub

Private Sub cmd_save_Click()
Dim ST As String
datetime
Dim hlreqno As String
Dim path_nosp As String
Dim Case_nosp As String
Dim Locationsp As String
txt_time.Text = time1
'On Error GoTo MSG:
Dim REC As Double
''''ADDED BY CHARULATA ON 14/09/2010 FOR DATETIME
datetime
txt_date.Text = DATE11
txt_time.Text = time1

'''UPTO HERE


ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
If MODALITY = "SP" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Or MODALITY = "TM" Or MODALITY = "CC" Or MODALITY = "FI" Then
    For i = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(i, 5) <> "" Then
            REC = REC + 1
        End If
    Next i
    If MsgBox("U'r Going to Save " & Trim(REC) & " Records, is it OK ?", vbOKCancel) = vbOK Then
          SPAUTONO
    Else
        Exit Sub
    End If
End If
'For i = 1 To MSFlexGrid1.Rows - 1
'    If MSFlexGrid1.TextMatrix(i, 6) <> "" Then
'        If modality = "BC" Or modality = "EL" Then
'            If seldequisitnflag = False Then
'                seldrequisitnqr(0) = st & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                Set seldrequisitnrs = seldrequisitnqr.OpenResultset(1, 3)
'                seldequisitnflag = True
'            Else
'                seldrequisitnqr(0) = st & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                seldrequisitnrs.MoreResults
'                seldrequisitnrs.Requery
'            End If
'        Else
'            If seldequisitnflag = False Then
'                seldrequisitnqr(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                Set seldrequisitnrs = seldrequisitnqr.OpenResultset(1, 3)
'                seldequisitnflag = True
'            Else
'                seldrequisitnqr(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                seldrequisitnrs.MoreResults
'                seldrequisitnrs.Requery
'            End If
'        End If
'        If MODIFYFLAG = False Then
'            If seldrequisitnrs.RowCount > 0 Then
'                MsgBox "The Path No. is  Given already " & seldrequisitnrs("caseno")
'                txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
'                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
'                txt_pathno.Visible = True
'                'chk_rcvd.Visible = True
'                'chk_rcvd.SetFocus
'                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                Exit Sub
'
'            End If
'        Else
'            MODIFYFLAG = True
'        End If
'    End If
'Next i

rdb.BeginTrans
If NEWFLAG = True Then
    For i = 1 To MSFlexGrid1.Rows - 1
        'MSFlexGrid1.Col = 4
        srow = i
        If Trim(dept) = "NZZ" Then
            If Trim(MSFlexGrid1.TextMatrix(i, 4)) <> "" And Trim(MSFlexGrid1.TextMatrix(i, 5)) <> "" Then
                updrrequisitn_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                If SELREQNORSFLAG = False Then
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
                    SELREQNORSFLAG = True
                Else
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    SELREQNORS.MoreResults
                    SELREQNORS.Requery
                End If

                If Trim(SELREQNORS("reqnstat")) = "P" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("reqnstat")) = "C" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "E" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                    updrrequisitn_qr(1) = "T"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "F" Then
                    updrrequisitn_qr(1) = "F"
                End If
                updrrequisitn_qr(2) = "RECEIVED"
                If Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RG" Then
                    updrrequisitn_qr(3) = "COLLRCVD"
                    'Locationsp = "COLLRCVD"
                Else
                    If Len(Trim(SELREQNORS("LOCATION"))) = "" Then
                        updrrequisitn_qr(3) = "COLLECTED"
                       ' Locationsp = "COLLECTED"
                    Else
                       updrrequisitn_qr(3) = Trim(SELREQNORS("LOCATION"))
                      ' Locationsp = Trim(SELREQNORS("LOCATION"))
                    End If
                End If
                updrrequisitn_qr(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                updrrequisitn_qr(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                updrrequisitn_qr.Execute

                ''RECEIPT TIME

                UPDLATERALITS_QR(0) = time1
                UPDLATERALITS_QR(1) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                UPDLATERALITS_QR.Execute
                ''***********

                If dept <> "NZZ" Then
                    If MODALITY = "BC" Or MODALITY = "EL" Then
                        If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 3 Then
                        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                        End If
                        If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <> 0 Then
                         upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                        End If
            '            If modality = "BC" Then
            '                BCINF
            '            End If
            '            If modality = "EL" Then
            '                ELINF
            '            End If
                    Else
                        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    End If
                    upddrequisitn(1) = ""
                    upddrequisitn(2) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    upddrequisitn.Execute
                End If
            End If
        Else
            If MSFlexGrid1.TextMatrix(i, 4) <> "" And MSFlexGrid1.TextMatrix(i, 5) <> "" And MSFlexGrid1.TextMatrix(i, 6) <> "" Then
                updrrequisitn_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                If SELREQNORSFLAG = False Then
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    'FETCH REQNO  IN BARCODE AMIT
                    REQNO = SELREQNOQR(0)
                    Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
                    SELREQNORSFLAG = True
                Else
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    SELREQNORS.MoreResults
                    SELREQNORS.Requery
                End If
                If Trim(SELREQNORS("reqnstat")) = "P" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("reqnstat")) = "C" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "E" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                    updrrequisitn_qr(1) = "T"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "F" Then
                    updrrequisitn_qr(1) = "F"
                End If
                updrrequisitn_qr(2) = "RECEIVED"
                If Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RG" Then
                    updrrequisitn_qr(3) = "COLLRCVD"
                    Locationsp = "COLLRCVD"
                Else
                    If Len(Trim(SELREQNORS("LOCATION"))) = 0 Then  ' made 0 from ""
                        updrrequisitn_qr(3) = "COLLECTED"
                        Locationsp = "COLLECTED"
                    Else
                       updrrequisitn_qr(3) = Trim(SELREQNORS("LOCATION"))
                       If IsNull(Trim(SELREQNORS("LOCATION"))) = True Then
                            Locationsp = ""
                       Else
                            Locationsp = Trim(SELREQNORS("LOCATION"))
                       End If
                    End If
                End If
                updrrequisitn_qr(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                updrrequisitn_qr(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                updrrequisitn_qr.Execute
                'THIS IS FOR INSERT IN RTIMELOG BY AMIT
'                INSTIME_QR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
'                INSTIME_QR(1) = msk_caseno.FormattedText
'                INSTIME_QR(2) = SELREQNORS("REQNDT")
'                INSTIME_QR(3) = SELREQNORS("REQNTM")
'                INSTIME_QR.Execute
                'THIS IS FOR UPDATE RTIMELOG ADDED BY AMIT
                               
                UPDTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
                UPDTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                'UPDTIME_QR(2) = msk_caseno.FormattedText
                'UPDTIME_QR(3) = SELREQNORS("REQNDT")
                'UPDTIME_QR(4) = SELREQNORS("REQNTM")
                UPDTIME_QR(2) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                UPDTIME_QR.Execute
                If MODALITY = "BC" Or MODALITY = "EL" Then
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 3 Then
                        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <> 0 Then
                        upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    End If
        '            If modality = "BC" Then
        '                BCINF
        '            End If
        '            If modality = "EL" Then
        '                ELINF
        '            End If
                Else
                    l = UCase(MSFlexGrid1.TextMatrix(i, 6))
                    upddrequisitn(0) = l
                    upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    path_nosp = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                End If
                If MODALITY = "CC" Or MODALITY = "FI" Then
                    If PROVCY_FLG = False Then
                        MsgBox "Select Specimen Status", vbCritical
                        Exit Sub
                    End If
                    If chk_clotted.Value = 1 Then upddrequisitn(1) = chk_clotted.Caption
                    If chk_dilute.Value = 1 Then upddrequisitn(1) = chk_dilute.Caption
                    If chk_hyper.Value = 1 Then upddrequisitn(1) = chk_hyper.Caption
                    If chk_leaked.Value = 1 Then upddrequisitn(1) = chk_leaked.Caption
                    If chk_ok.Value = 1 Then upddrequisitn(1) = chk_ok.Caption
                Else
                    upddrequisitn(1) = ""
                End If
                              
                upddrequisitn(2) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                upddrequisitn.Execute
                If MODALITY = "CC" Or MODALITY = "FI" Then
                    UPDATE_CYTOG(0) = msk_cdate.FormattedText
                    UPDATE_CYTOG(1) = Format(MSK_CTIME, "hh:mm:ss") 'MSK_CTIME.FormattedText
                    UPDATE_CYTOG(2) = Trim(cmb_recename.Text)
                    UPDATE_CYTOG(3) = Trim(txt_normal.Text)
                    If Trim(txt_age.Text) <> "" Then
                        UPDATE_CYTOG(4) = Trim(txt_age.Text)
                    Else
                        UPDATE_CYTOG(4) = 0
                    End If
                    UPDATE_CYTOG(5) = Trim(cmb_sex.Text)
                    UPDATE_CYTOG(6) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    UPDATE_CYTOG.Execute
                End If
                
                ' Insert value into SPathology from here
                If MODALITY = "SP" Then
                'path_no,case_no,trans_dt,loct1
                        Ins_Spathology_qr(0) = path_nosp
                        Ins_Spathology_qr(1) = MSFlexGrid1.TextMatrix(i, 2)
                        Ins_Spathology_qr(2) = Format(DATE11, "YYYY/MM/DD")
                        Ins_Spathology_qr(3) = Locationsp
                        Ins_Spathology_qr.Execute
                    End If
                ' Insert value into SPathology from here
 ''  HEMATO INTERFACE CODE BEGINS
              If MODALITY = "HL" Then
                  hlreqno = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                   SELDHEMCYTO_QR(0) = hlreqno  ' req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                   Set SELDHEMCYTO_RS = SELDHEMCYTO_QR.OpenResultset(1, 1)
                 If SELDHEMCYTO_RS.RowCount > 0 Then
                      If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
                      '' INSERT FOR HEMOGRAM OR CBC INTERFACE
                    '  CASENO,REQNO,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
                     
                       INSHINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                       INSHINTERFACE_QR(1) = hlreqno
                       INSHINTERFACE_QR(2) = " "
                       INSHINTERFACE_QR(3) = "1"
                       INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "A"
                       INSHINTERFACE_QR.Execute
                        '' INSERT FOR HEMOGRAM OR CBC INTERFACE ENDS
                      End If
                       
                    
                     If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
                       seldmulproce_qr(0) = hlreqno
                       Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 1)
                        Do While seldmulproce_rs.EOF = False
                            seldhlprosub_qr(0) = seldmulproce_rs("pro_code")
                            Set seldhlprosub_rs = seldhlprosub_qr.OpenResultset(1, 1)
                            Do While seldhlprosub_rs.EOF = False
  ' STR = "INSERT INTO INTERFACE.CINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
                            If seldhlprosub_rs("samplecd") = "B" Then
                               INSCINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                               INSCINTERFACE_QR(1) = hlreqno
                               INSCINTERFACE_QR(2) = seldhlprosub_rs("procodesub")
                               INSCINTERFACE_QR(3) = " "
                               INSCINTERFACE_QR(4) = "1"
                               INSCINTERFACE_QR(5) = Trim(Mid(hlreqno, 4, 10)) + Trim(seldhlprosub_rs("samplecd"))
                               INSCINTERFACE_QR.Execute
                            End If
                               seldhlprosub_rs.MoveNext
                                               
                                                           
                            Loop
                           seldmulproce_rs.MoveNext
                        Loop
                        
                     End If
                 End If 'dhemctyo rowcount
                 
              End If  'HL
              
 ''  HEMATO INTERFACE ENDS
 
                If MODALITY = "BC" Then
                    If SELDMULPROCE_FLAG = False Then
                      seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                      Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
                      SELDMULPROCE_FLAG = True
                    Else
                      seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                      seldmulproce_rs.MoreResults
                      seldmulproce_rs.Requery
                    End If
                    REQNO = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    If Not seldmulproce_rs.EOF Then
                      Do While Not seldmulproce_rs.EOF
                        'checking for record in excption codes which are used as manual
                        ' and not in interface
                        'WRITTE BY AMIT STORING DATA IN INTERFACE/DINTERFACE
                    PCODE = seldmulproce_rs("PRO_CODE")
                    If PCODE = "F849" Or PCODE = "F854" Or PCODE = "F924" Or PCODE = "F860" Or PCODE = "F875" Or PCODE = "F880" Or PCODE = "F875" Or PCODE = "F865" Or PCODE = "F802" Then
                        SELDBCPROSUB_QR(0) = PCODE
                        Set SELDBCPROSUB_RS = SELDBCPROSUB_QR.OpenResultset(1, 3)
                        Do Until SELDBCPROSUB_RS.EOF
                            interexpcd_qr(0) = SELDBCPROSUB_RS!PROCODESUB
                            Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                            If interexpxd_rs.RowCount = 0 Then
                                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                Dim REQ1, P_CODE, SCODE
                                REQ1 = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                                P_CODE = SELDBCPROSUB_RS!PROCODESUB
                                SELSAMPLECODE_QR(0) = P_CODE
                                Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                                SCODE = SAMPLECODE_RS!SAMPLECD
                                REQ2 = Mid$(REQ1, 4)
                                SCODE = REQ2 & SCODE
                                INSDINTERFACE_QR(1) = REQ1
                                INSDINTERFACE_QR(2) = SELDBCPROSUB_RS!PROCODESUB
                                INSDINTERFACE_QR(3) = ""
                                INSDINTERFACE_QR(4) = "1"
                                INSDINTERFACE_QR(5) = SCODE
                                INSDINTERFACE_QR.Execute
                            End If
                            SELDBCPROSUB_RS.MoveNext
                        Loop
                     Else
                        interexpcd_qr(0) = seldmulproce_rs("PRO_CODE")
                        Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                        If interexpxd_rs.RowCount = 0 Then
                            INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                Dim REQ11, P_CODE1, SCODE1
                                REQ11 = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                                SELSAMPLECODE_QR(0) = PCODE
                                Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                                SCODE1 = SAMPLECODE_RS!SAMPLECD
                                REQ2 = Mid$(REQ11, 4)
                                SCODE1 = REQ2 & SCODE1
                                
                            INSDINTERFACE_QR(1) = REQ11
                            INSDINTERFACE_QR(2) = seldmulproce_rs("PRO_CODE")
                            INSDINTERFACE_QR(3) = ""
                            INSDINTERFACE_QR(4) = "1"
                            INSDINTERFACE_QR(5) = SCODE1
                            INSDINTERFACE_QR.Execute
                        End If
                    End If
                    seldmulproce_rs.MoveNext
                    Loop
                    End If
                        
                ' END BY AMIT
                  
                  End If
                  
                  
                  'TM INTERFACE 2005
                  
                  If MODALITY = "TM" Then
                     If SELDMULPROCE_FLAG = False Then
                        seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                        Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
                        SELDMULPROCE_FLAG = True
                     Else
                        seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                        seldmulproce_rs.MoreResults
                        seldmulproce_rs.Requery
                     End If
                     REQNO = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                     If Not seldmulproce_rs.EOF Then
                        Do While Not seldmulproce_rs.EOF
                            PCODE = seldmulproce_rs("PRO_CODE")
                            ''''''''''
                                                                                 
                        If PCODE = "F033" Or PCODE = "F043" Or PCODE = "F045" Or PCODE = "F049" Then
                        SELDBCPROSUB_QR(0) = PCODE
                        Set SELDBCPROSUB_RS = SELDBCPROSUB_QR.OpenResultset(1, 3)
                        
                            Do Until SELDBCPROSUB_RS.EOF
                                interexpcd_qr(0) = SELDBCPROSUB_RS("PROCODESUB")
                                Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                            If interexpxd_rs.RowCount = 0 Then
                            
                            ''
                              INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                Dim REQ1T, P_CODE1t, SCODE11
                                REQ1T = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                                P_CODE1t = SELDBCPROSUB_RS!PROCODESUB
                                SELSAMPLECODE_QR(0) = P_CODE1t
                                Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                                SCODE11 = SAMPLECODE_RS!SAMPLECD
                                REQ2 = Mid$(REQ1T, 4)
                                SCODE11 = REQ2 & SCODE11
                            ''
                                    
                                INSDINTERFACE_QR(1) = REQ1T
                                INSDINTERFACE_QR(2) = SELDBCPROSUB_RS("PROCODESUB")
                                INSDINTERFACE_QR(3) = ""
                                INSDINTERFACE_QR(4) = "1"
                                INSDINTERFACE_QR(5) = SCODE11
                                INSDINTERFACE_QR.Execute
                              End If
                            SELDBCPROSUB_RS.MoveNext
                            Loop
                        
                        ''''''''
                            
                    Else
                            interexpcd_qr(0) = seldmulproce_rs("PRO_CODE")
                            Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                            If interexpxd_rs.RowCount = 0 Then
                                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                Dim REQ11t, P_CODE1TT, SCODE1t
                                REQ11t = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                                SELSAMPLECODE_QR(0) = P_CODE1TT
                                Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                                If SAMPLECODE_RS.RowCount > 0 Then
                                   SCODE1t = SAMPLECODE_RS!SAMPLECD
                                Else
                                    SCODE1t = " "
                                End If
                                    
                                    REQ2 = Mid$(REQ11t, 4)
                                    SCODE1t = REQ2 & SCODE1t
                                    
                                INSDINTERFACE_QR(1) = REQ11t
                                INSDINTERFACE_QR(2) = seldmulproce_rs("PRO_CODE")
                                INSDINTERFACE_QR(3) = ""
                                INSDINTERFACE_QR(4) = "1"
                                INSDINTERFACE_QR(5) = SCODE1t
                                INSDINTERFACE_QR.Execute
                            End If
                         End If
                            seldmulproce_rs.MoveNext
                        Loop
                      End If
                      
                      ''''FOR SAVING IN BINTERFACE'''
                      
                      
                      Set RS_DMULPROCE = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE REQNO='" & REQNO & "' AND PRO_CODE='F816' ")
                      If RS_DMULPROCE.RowCount > 0 Then
                        Dim req As Variant
                        req = Mid$(REQNO, 4)
                        INSBINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                        INSBINTERFACE_QR(1) = REQNO
                        INSBINTERFACE_QR(2) = "F816"
                        INSBINTERFACE_QR(3) = ""
                        INSBINTERFACE_QR(4) = "1"
                        INSBINTERFACE_QR(5) = req & "A"
                        INSBINTERFACE_QR.Execute
                      End If
                      
                      
                      
                  End If
                 
                'TM INTERFACE 2005
              End If

            End If
            
          Next i
          
         
      MsgBox "Record is Successfully Saved."
    
End If
If MODIFYFLAG = True Then

For i = 1 To MSFlexGrid1.Rows - 1
     'MSFlexGrid1.Col = 4
    If MSFlexGrid1.TextMatrix(i, 4) <> "" And MSFlexGrid1.TextMatrix(i, 5) <> "" And MSFlexGrid1.TextMatrix(i, 6) <> "" Then
        updrrequisitn_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        If SELREQNORSFLAG = False Then
            SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
            Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
            SELREQNORSFLAG = True
        Else
            SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
            SELREQNORS.MoreResults
            SELREQNORS.Requery
        End If
        If Trim(SELREQNORS("reqnstat")) = "P" Then
            updrrequisitn_qr(1) = "E"
        End If
        If Trim(SELREQNORS("reqnstat")) = "C" Then
            updrrequisitn_qr(1) = "E"
        End If
        If Trim(SELREQNORS("REQNSTAT")) = "E" Then
            updrrequisitn_qr(1) = "E"
        End If
        If Trim(SELREQNORS("REQNSTAT")) = "T" Then
            updrrequisitn_qr(1) = "T"
        End If
        If Trim(SELREQNORS("REQNSTAT")) = "F" Then
            updrrequisitn_qr(1) = "F"
        End If
        updrrequisitn_qr(2) = "RECEIVED"
        If Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RG" Then
            updrrequisitn_qr(3) = "COLLRCVD"
        Else
            updrrequisitn_qr(3) = "COLLECTED"
        End If
        updrrequisitn_qr(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        updrrequisitn_qr(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
        updrrequisitn_qr.Execute
        If MODALITY = "BC" Or MODALITY = "EL" Then
            If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 3 Then
            upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
            End If
            If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <> 0 Then
             upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
            End If
            If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) = 0 Then
             upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
            End If
        Else
            upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
        End If
        upddrequisitn(1) = ""
        upddrequisitn(2) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
        upddrequisitn.Execute
    End If
Next i
MsgBox "Record is Successfully Saved."
End If
rdb.CommitTrans

btnenabl ("SAVE")
Exit Sub
'MSG:
rdb.RollbackTrans
'MsgBox Err.Description
btnenabl ("SAVE")
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
BARCODE_VIEW = True
chk_rcvd.Value = 0
viewflag = True
NEWFLAG = False
If SELMA106FLAG = False Then
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106FLAG = True
Else
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106_RS.MoreResults
    SELMA106_RS.Requery
    
End If
If SELMA106_RS.RowCount > 0 Then
    cmb_dept.clear
    While Not SELMA106_RS.EOF
        cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
        SELMA106_RS.MoveNext
    Wend
End If
txt_name.Text = ""
OPT_ALL.Enabled = True
OPT_ALL.Value = False
OPT_CASENO.Enabled = True
OPT_ALL.Value = False
Opt_patient.Enabled = True
Opt_patient.Value = False
opt_staff.Enabled = True
opt_staff.Value = False
chk_rcvd.Value = 0
chk_rcvd.Visible = False

btnenabl ("VIEW")
FLDENABL ("ENA")
cmb_dept.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_VPP_Click()
On Error GoTo MSG:
CrystalReport1.ReportFileName = "C:\DISREPORTSEC\vpp.RPT"
CrystalReport1.SelectionFormula = " {RREQUISITN.REQNNO} ='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) & "'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
clearcy
cy_frame.Visible = False
If deptcollectionflag = False Then
datetime
msk_cdate.Text = Format(DATE11, "dd/mm/yyyy")
txt_date.Text = Format(DATE11, "dd/mm/yyyy")
txt_time.Text = time1
MSK_CTIME.Text = Format(time1, "HH:MM")
'chk_rcvd.Visible = False
txt_pathno.Visible = False
deptcollectionflag = True
scrlck_flg = True
MODIFYFLAG = False


'INTERFACE
str = "SELECT DISTINCT LABNO FROM INTERFACE.DINTERFACE WHERE REQNO=?"
Set SELLABNO_QR = rdb.CreateQuery("LABNOD", str)

' FOR RTIMELOG
str = "INSERT INTO RISLIB.RTIMELOG(REQNNO,CASENO,REQNDATE,REQNTIME) VALUES(?,?,?,?)"
Set INSTIME_QR = rdb.CreateQuery("FHGSDF1D", str)

''***************INTERFACE**********************
str = "INSERT INTO INTERFACE.DINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
Set INSDINTERFACE_QR = rdb.CreateQuery("INSDINTERFACE_STRD", str)

str = "INSERT INTO INTERFACE.BINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
Set INSBINTERFACE_QR = rdb.CreateQuery("INSBINTERFACE_STRD", str)


'' HEMATO INTERFACE
str = "INSERT INTO INTERFACE.HINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?)"
Set INSHINTERFACE_QR = rdb.CreateQuery("INSHINTERFACE_STRD", str)
str = "INSERT INTO INTERFACE.CINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
Set INSCINTERFACE_QR = rdb.CreateQuery("INSCINTERFACE_STRD", str)
str = "SELECT * FROM DISLIB.DHEMCYTO1 WHERE REQNO = ?"
Set SELDHEMCYTO_QR = rdb.CreateQuery("SELDHEMCYTO1_QRD", str)

'STR = "select * from dislib.dmulproce where reqno = ?"
str = "select * from dislib.dhlprosub where hlprocode = ?"
Set seldhlprosub_qr = rdb.CreateQuery("SELDHLPROSUB_QR1D", str)

'' HEMATO INTERFACE

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set seldmulproce_qr = rdb.CreateQuery("SELDMULPROCE_STRDDD", str)

str = "SELECT * FROM DISLIB.INTEREXPCD WHERE excpcode =?"
Set interexpcd_qr = rdb.CreateQuery("interexpcd_qr1D", str)

str = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
Set SELDBCPROSUB_QR = rdb.CreateQuery("DBCPROSUB_SELD", str)

str = "SELECT SAMPLECD FROM DISLIB.SAMPLECODE WHERE SERCODE=?"
Set SELSAMPLECODE_QR = rdb.CreateQuery("SAMPLE1D", str)




''***************INTERFACE**********************
str = "SELECT SAMRECDATE,SAMRECTIME FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set SELTIME_QR = rdb.CreateQuery("FFHSHFD", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMRECDATE=?,SAMRECTIME=? WHERE REQNNO=? "
Set UPDTIME_QR = rdb.CreateQuery("DFHWHD", str)
''*********tracking System*************
str = "insert into dislib.spathology(path_no,case_no,trans_dt,loct1) values(?,?,?,?)"
Set Ins_Spathology_qr = rdb.CreateQuery("Ins_Spathology_qrD", str)
''*********tracking System*************

str = "INSERT INTO DISLIB.DAUTOPATH (DEPT_CODE,MODALITYCD,PATHYEAR,PATHYEAR1,PATHNO) VALUES (?,?,?,?,?)"
Set INSDAUTOPATH_QR = rdb.CreateQuery("INSDAUTOPATH_QRXS1D", str)

str = "UPDATE DISLIB.DAUTOPATH SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set UPDDAUTOPATH_QR = rdb.CreateQuery("UPDDAUTOPATH_Q11RD1", str)
        
' ***** AUTOGENERATION OF THE PATHNO.
str = "SELECT * FROM DISLIB.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set SELDAUTOPATH_QR = rdb.CreateQuery("SEXXLDAUTOPATH_QRD1", str)

str = "SELECT * FROM DISLIB.DPATHYEAR WHERE AYEAR=?"
Set DPATHYEAR_QR = rdb.CreateQuery("DPXXATHYEAR_QR1D", str)
'''RECEIPT TIME
str = "UPDATE DISLIB.DREQUISITN SET LATERALITS=? WHERE REQNO=?"
Set UPDLATERALITS_QR = rdb.CreateQuery("UPDLATERALITS_STRD", str)

''DTMNO
str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR11D", str)

str = "select * from patoolslib.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ'  OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' or dept_code='IZZ' or dept_code='VZZ')"
Set SELMA106_QR = rdb.CreateQuery("selma106_qr11D", str)

str = "select * from Rislib.rrequisitn where dept_code=? and modalitycd=? and (reqndt between ? and ?) AND (REQNSTAT=? or REQNSTAT=?) ORDER BY CASENO "
Set SELRREQUISITN_QR = rdb.CreateQuery("selrrequisitn_qr111D", str)

str = "select * from Rislib.rrequisitn where dept_code='NZZ' and (reqndt between ? and ?) AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=?) ORDER BY CASENO "
Set SELRREQUISITN_QR1 = rdb.CreateQuery("selrrequisitn_qr11D", str)

str = "select * from Rislib.rrequisitn where dept_code='NZZ' and (reqndt between ? and ?) AND MODALITYCD=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=?) ORDER BY CASENO "
Set SELRREQUISITN_QR11 = rdb.CreateQuery("selrrequisitn_qr10091D", str)


str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_STR111D", str)

'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE=? AND MODALITYCD=? AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=?)"
str = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE=? AND MODALITYCD=? AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? )"
Set SELRREQNO_QR = rdb.CreateQuery("SELRREQNO1_QR1D", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE='NZZ' AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? )"
Set SELRREQNO_QR1 = rdb.CreateQuery("SELRREQNO1_QR1001D", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE='NZZ' AND MODALITYCD=? AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=?)"
Set SELRREQNO_QR11 = rdb.CreateQuery("CCVSELRREQNO1_QR1001DD", str)

str = "select * from rislib.rmodality where dept_code=?"
Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr1D", str)

str = "select * from patoolslib.ma101 where ser_code=?"
Set SELMA101_QR = rdb.CreateQuery("selma1011_qr1D", str)

str = "UPDATE  RISLIB.RREQUISITN SET EXAMDT=?, REQNSTAT=?,EXAMINATN=?,LOCATION=?,APNTDT=? WHERE REQNNO=?"
Set updrrequisitn_qr = rdb.CreateQuery("UPDRREQUISITN_QR1D1", str)

str = "select * from patoolslib.ta101 where case_no=?"
Set SELTA1011_QR = rdb.CreateQuery("selta10111_qr1D", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE EMPCODE=?"
Set SELDREQSTAFFD1_QR = rdb.CreateQuery("SELDREQSTAFFD_QR11D", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE PATHNO=? AND  REQNO LIKE ?"
'Commented and added by Shraddha on 19/07/10
'str = "SELECT * FROM DISLIB.DREQLF_R WHERE PATHNO=? AND  REQNO LIKE ?"
Set SELDREQUISITNQR = rdb.CreateQuery("1SELDREQUISITNQR1D", str)

str = "UPDATE DISLIB.DREQUISITN SET PATHNO=?,NMETERIALS=? WHERE REQNO=?"
Set upddrequisitn = rdb.CreateQuery("UPDRREQUISITN11D", str)
            
str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDREQQR = rdb.CreateQuery("SELDREQQR1DEP", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNOQR = rdb.CreateQuery("SELREQNOQRDEP1", str)

'**** ADDED FOR BIOCHEMISTRY, EL, TM INTERFACEING
str = "INSERT INTO DISLIB.DBCINFACE(REQNO,LABNO,PROCODE) VALUES(?,?,?)"
Set INSDBCINFACEQR = rdb.CreateQuery("INSDBCINFACEQRDEPT", str)

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRBCdept", str)
 
str = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
Set SELDBCPROSUBQR = rdb.CreateQuery("SELDBCPROSUBQRBCdept", str)

str = "SELECT * FROM DISLIB.DELPROSUB WHERE PROCODE=?"
Set SELELPROSUBQR = rdb.CreateQuery("SELELPROSUBQRELdept", str)

str = "UPDATE DISLIB.CYTOGCLDTA SET CALDATE=?,CALTIME=?,REMARK1=?,NCONTROL=?,NCON_AGE=?,NCON_SEX=?  WHERE REQNO=?"
Set UPDATE_CYTOG = rdb.CreateQuery("CYTOGLDCYY", str)

'justin
 str = "SELECT max(pathno) FROM DISLIB.DREQUISITN WHERE REQNO LIKE '%TM%' AND PATHNO LIKE ?"
 Set SELDREQUISITN_PATH_QR = rdb.CreateQuery("1SELDREQUISITN_PATH_QR1", str)
  
'**** ADDED FOR BIOCHEMISTRY, EL, TM INTERFACEING
End If
MSFlexGrid1.row = 1
MSFlexGrid1.col = 5
chk_rcvd.Width = MSFlexGrid1.CellWidth
chk_rcvd.Height = MSFlexGrid1.CellHeight
chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

MSFlexGrid1.row = 1
MSFlexGrid1.col = 6
txt_pathno.Width = MSFlexGrid1.CellWidth
txt_pathno.Height = MSFlexGrid1.CellHeight
txt_pathno.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top


NEWFLAG = False
viewflag = False
btnenabl ("INI")
FLDENABL ("DIS")
OPT_ALL.Enabled = False
OPT_CASENO.Enabled = False
Opt_patient.Enabled = False
opt_staff.Enabled = False
OPT_ALL.Value = False
OPT_CASENO.Value = False
Opt_patient.Value = False
opt_staff.Value = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    cmb_dept.Enabled = True
    cmb_modality.Enabled = True
    msk_fromdt.Enabled = True
    msk_todt.Enabled = True

End If
If str1 = "DIS" Then
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
End If
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

Private Sub MSFlexGrid1_Click()
On Error GoTo MSG:
R1 = MSFlexGrid1.row
C1 = MSFlexGrid1.col
II = req_chglen(MSFlexGrid1.TextMatrix(R1, 1))
If MSFlexGrid1.col = 1 Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONCY.RPT"
            
            CrystalReport1.SelectionFormula = " {RREQUISITN.REQNNO} ='" & req_chglen(MSFlexGrid1.TextMatrix(R1, 1)) & "' and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "'"
            
            If Right(Trim(cmb_modality.Text), 2) = "CC" Then
            CrystalReport1.Formulas(1) = "MODALITY ='CONVENTIONAL CYTOGENETICS'"
            Else
            CrystalReport1.Formulas(1) = "MODALITY ='MOLECULAR CYTOGENETICS(FISH) REPORT'"
            End If
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        
            CrystalReport1.Action = 1
          Exit Sub


    On Error GoTo MSG:
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\vpp.RPT"
    CrystalReport1.SelectionFormula = " {RREQUISITN.REQNNO} ='" & req_chglen(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)) & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
    Exit Sub
End If
If MODIFYFLAG = False Then
If MODALITY = "SP" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Or Trim(MODALITY) = "CC" Or Trim(MODALITY) = "FI" Then
    If MSFlexGrid1.col = 5 Then
    'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
    
            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
            'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
    
             chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
             chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
            chk_rcvd.Value = 1
            chk_rcvd.SetFocus
            Else
            chk_rcvd.Value = 0
            End If
    
     End If
Else
    If dept <> "NZZ" Then
        If MSFlexGrid1.col = 6 Then
        'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
    
            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    
             'Chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            'If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
            'txt_pathno.SetFocus
'            chk_rcvd.Value = 1
'            chk_rcvd.SetFocus
'            Else
'            chk_rcvd.Value = 0
'            End If
        End If
    End If
    If MODALITY = "BP" Or MODALITY = "SS" Or MODALITY = "HT" Or MODALITY = "ALL" Or MODALITY = "TM" Then
        If MSFlexGrid1.col = 5 Then
        'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
    
            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
            'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
    
             chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
             chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
                chk_rcvd.Value = 1
                chk_rcvd.SetFocus
            Else
                chk_rcvd.Value = 0
            End If
        End If
     End If
     Exit Sub
 End If
Else
    If MODALITY = "SP" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Then
        If MSFlexGrid1.col = 5 Then
    'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
    
            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
            'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
    
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
            chk_rcvd.Value = 1
            Else
            chk_rcvd.Value = 0
            End If
     End If
   Else
  If MSFlexGrid1.col = 6 Then
    'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
    
            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    
            'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            'txt_pathno.SetFocus
'            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
'            chk_rcvd.Value = 1
'            Else
'            chk_rcvd.Value = 0
'            End If
     End If
 End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub MSFlexGrid1_GotFocus()
If MSFlexGrid1.row = 0 Then
Exit Sub
End If
'MSFlexGrid1.Row = 1
'MSFlexGrid1.Col = 5
'chk_rcvd.Width = MSFlexGrid1.CellWidth
'chk_rcvd.Height = MSFlexGrid1.CellHeight
'chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If dept <> "NZZ" Then
If scrlck_flg = True Then
    MSFlexGrid1.row = 1
End If
MSFlexGrid1.col = 6
 If txt_pathno.Visible <> True Then
    txt_pathno.Width = MSFlexGrid1.CellWidth
    txt_pathno.Height = MSFlexGrid1.CellHeight
    txt_pathno.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
 End If
End If
scrlck_flg = False
End Sub

Private Sub msk_caseno_LostFocus()
On Error GoTo MSG:
CASENO = msk_caseno.Text
If viewflag = True Then
    If Trim(dept) = "GZZ" Then
        cmd_modify.Enabled = True
    End If
End If
If Opt_patient.Value = True Then
    '''' Commented PADZERO FUNCTION INSTEAD USED NEWNOPAD FUNCTION
'cno_pad = padzero(5, "C", msk_caseno.Text, 3)
    

    
  If IsNumeric(Left(Trim(msk_caseno.Text), 1)) Then
          Dim Ks As Integer
        Dim str1s As String
        Dim strs As String
            strs = Trim(msk_caseno.Text)
            Ks = Len(strs)
            str1s = Mid(strs, 5, Ks)
              Dim sel_CASE_rs  As rdo.rdoResultset
              Set sel_CASE_rs = rdb.OpenResultset(" SELECT * FROM PATOOLSLIB.NEWCASETBL WHERE CASE_CODE = '" & Left(UCase(msk_caseno.FormattedText), 4) & "' ", 1, 1)
              If Not sel_CASE_rs.EOF Then
                    Dim CSNO As String
                    CSNO = Trim(sel_CASE_rs(1))
                    CSNO = CSNO & str1s
                     cno_pad = NEWNOpad(CSNO)
              Else
               cno_pad = NEWNOpad(UCase(msk_caseno))
              End If
    
  Else
    cno_pad = NEWNOpad(UCase(msk_caseno))
  End If
     msk_caseno.Text = cno_pad
''''''''''''''''''''till here (01/12/2022)'''''''''
    If SELTA1011FLAG = False Then
        SELTA1011_QR(0) = msk_caseno.FormattedText
        Set SELTA1011_RS = SELTA1011_QR.OpenResultset(1, 3)
        SELTA1011FLAG = True
    Else
        SELTA1011_QR(0) = msk_caseno.FormattedText
        SELTA1011_RS.MoreResults
        SELTA1011_RS.Requery
    End If
    If SELTA1011_RS.RowCount > 0 Then
        txt_name = Trim(SELTA1011_RS("tag")) & " " & Trim(SELTA1011_RS("first_name")) & " " & Trim(SELTA1011_RS("mid_name")) & " " & Trim(SELTA1011_RS("last_name"))
    End If
    If SELDTMNO_FLAG = False Then
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        SELDTMNO_FLAG = True
    Else
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        SELDTMNO_RS.MoreResults
        SELDTMNO_RS.Requery
    End If
    If dept = "NZZ" And Trim(MODALITY) = "ALL" Then
        DTM1
        Exit Sub
    ElseIf dept = "NZZ" And Trim(MODALITY) <> "ALL" Then
        DTM1MODWISE
        Exit Sub
    End If
    If NEWFLAG = True Then
        'NEWFLAG = False
        If SELRREQNOFLAG = False Then
            SELRREQNO_QR(0) = Trim(dept)
            SELRREQNO_QR(1) = Trim(MODALITY)
            SELRREQNO_QR(2) = msk_caseno.FormattedText
            SELRREQNO_QR(3) = "C"
            SELRREQNO_QR(4) = "P"
            SELRREQNO_QR(5) = "P"
            SELRREQNO_QR(6) = "P"
            Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
            SELRREQNOFLAG = True
        Else
            SELRREQNO_QR(0) = Trim(dept)
            SELRREQNO_QR(1) = Trim(MODALITY)
            SELRREQNO_QR(2) = msk_caseno.FormattedText
            SELRREQNO_QR(3) = "C"
            SELRREQNO_QR(4) = "P"
            SELRREQNO_QR(5) = "P"
            SELRREQNO_QR(6) = "P"
            Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
            SELRREQNO_RS.MoreResults
            SELRREQNO_RS.Requery
        End If
        If SELRREQNO_RS.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQNO_RS.EOF
                If SELRREQNO_RS("SER_CODE") <> "8888" Then
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If Trim(MODALITY) = "SP" Then
                        If SELDREQFLAG = False Then
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                            SELDREQFLAG = True
                        Else
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            SELDREQRS.MoreResults
                            SELDREQRS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                                If SELDREQRS.RowCount > 0 Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & SELDREQRS("PATHNO")
                                    'chk_collected.Value = 1
                                    'txt_pathno.Visible = False
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & ""
                                    'chk_collected.Value = 1
                                    'txt_pathno.Visible = False
                                End If
                            Else
                                If SELDREQRS.RowCount > 0 Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDREQRS("PATHNO"))
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & ""
                                End If
                            End If
                            i = i + 1
                        End If
                    Else
                        If SELMA101_RS.RowCount > 0 Then
                            
                            If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                                If dept = "NZZ" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                            GoTo m
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                    'chk_collected.Value = 1
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                End If
                            Else
                                If dept = "NZZ" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                            GoTo m
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                                End If
                            End If
                            i = i + 1
                        End If
                    End If
                Else
                    If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                        If dept = "NZZ" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                                    GoTo m
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo m
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                        'chk_collected.Value = 1
                        End If
                    Else
                        If dept = "NZZ" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                                    GoTo m
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo m
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE"
                        End If
                    End If
m:
                    i = i + 1
                End If
                SELRREQNO_RS.MoveNext
            Wend
            If Trim(MODALITY) = "SP" Or Trim(MODALITY) = "MP" Or Trim(MODALITY) = "CC" Or Trim(MODALITY) = "FI" Then
                chk_rcvd.Visible = True
                txt_pathno.Visible = False
            Else
                chk_rcvd.Visible = False
                txt_pathno.Visible = True
            End If
            If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or MODALITY = "TM" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Or MODALITY = "CC" Or MODALITY = "FI" Then
                MSFlexGrid1.row = 1
                MSFlexGrid1.col = 5
                chk_rcvd.Width = MSFlexGrid1.CellWidth
                chk_rcvd.Height = MSFlexGrid1.CellHeight
                chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                chk_rcvd.Value = 0
                chk_rcvd.Visible = True
                txt_pathno.Visible = False
            End If
            btnenabl ("NEW")
            'chk_collected.Visible = True
        Else
            MsgBox "No pending Requisition No."
            'chk_rcvd.Visible = False
            txt_pathno.Visible = False
            'btnenabl ("INI")
            FLDENABL ("DIS")
            'cmd_new.SetFocus
            Exit Sub
        End If
        If Trim(MODALITY) = "SP" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Value = 0
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
    End If
    If viewflag = True Then
    viewflag = False
    If SELRREQNOFLAG = False Then
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "E"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNOFLAG = True
                Else
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "E"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNO_RS.MoreResults
                    SELRREQNO_RS.Requery
                End If
                If SELRREQNO_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQNO_RS.EOF
                        If Trim(SELRREQNO_RS("SER_CODE")) <> "8888" Then
                                If SELMA101FLAG = False Then
                                    SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                    SELMA101FLAG = True
                                Else
                                    SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                                    SELMA101_RS.MoreResults
                                    SELMA101_RS.Requery
                                End If
                                If SELMA101_RS.RowCount > 0 Then
                                    If SELDREQFLAG = False Then
                                        SELDREQQR(0) = Trim(SELRREQNO_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                    Else
                                        SELDREQQR(0) = Trim(SELRREQNO_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                    End If
                                    
                                        If Trim(SELRREQNO_RS("REQNSTAT")) = "E" Then
                                            If SELDREQRS.RowCount > 0 Then
                                                 If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            i = i + 1
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    i = i + 1
                                                 Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                    i = i + 1
                                                 End If
                                            Else
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    i = i + 1
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                    i = i + 1
                                                End If
                                            End If
                                        Else
                                            If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            i = i + 1
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                    i = i + 1
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                    i = i + 1
                                                End If
                                            End If
                                        End If
                                    
                                   ' i = i + 1
                                    
                                End If
                        Else
                                    If SELDREQFLAG = False Then
                                        SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                    Else
                                        SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                    End If
                        
                                    If Trim(SELRREQNO_RS("REQNSTAT")) = "E" Then
                                            If SELDREQRS.RowCount > 0 Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            'i = i + 1
                                                            Exit Sub
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            
                                                            Exit Sub
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                End If
                                            Else
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                End If
                                            End If
                                        Else
                                            If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                End If
                                            End If
                                        End If
n:
                                     i = i + 1
                        End If
                        SELRREQNO_RS.MoveNext
                    Wend
                    If Trim(MODALITY) = "SP" Then
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    Else
                        chk_rcvd.Visible = False
                        txt_pathno.Visible = True
                    End If
                    If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "MP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Then
                        MSFlexGrid1.row = 1
                        MSFlexGrid1.col = 5
                        chk_rcvd.Width = MSFlexGrid1.CellWidth
                        chk_rcvd.Height = MSFlexGrid1.CellHeight
                        chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                        chk_rcvd.Value = 0
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    End If
                           'chk_rcvd.Visible = False
                   txt_pathno.Visible = False
                   'chk_collected.Visible = False
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

Private Sub msk_todt_LostFocus()
Dim i As Integer
On Error GoTo MSG:
If OPT_ALL.Value = True Then
    If dept = "NZZ" And Trim(MODALITY) = "ALL" Then
        DTM
        Exit Sub
    ElseIf dept = "NZZ" And Trim(MODALITY) <> "ALL" Then
        DTMMODWISE
        Exit Sub
    End If
    If NEWFLAG = True Then
                    'NEWFLAG = False
                    
                If SELRREQUISITNFLAG = False Then
                    SELRREQUISITN_QR(0) = Trim(dept)
                    SELRREQUISITN_QR(1) = Trim(MODALITY)
                    SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                    SELRREQUISITN_QR(3) = msk_todt.FormattedText
                    SELRREQUISITN_QR(4) = "C"
                    SELRREQUISITN_QR(5) = "P"
                    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                    SELRREQUISITNFLAG = True
                Else
                    SELRREQUISITN_QR(0) = Trim(dept)
                    SELRREQUISITN_QR(1) = Trim(MODALITY)
                    SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                    SELRREQUISITN_QR(3) = msk_todt.FormattedText
                    SELRREQUISITN_QR(4) = "C"
                    SELRREQUISITN_QR(5) = "P"
                    SELRREQUISITN_RS.MoreResults
                    SELRREQUISITN_RS.Requery
                End If
        
                
                If SELRREQUISITN_RS.RowCount = 0 Then
                    Exit Sub
                End If
                If SELRREQUISITN_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQUISITN_RS.EOF
                        If SELDTMNO_FLAG = False Then
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") ' msk_caseno.FormattedText
                            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                            SELDTMNO_FLAG = True
                        Else
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") 'msk_caseno.FormattedText
                            SELDTMNO_RS.MoreResults
                            SELDTMNO_RS.Requery
                        End If
                        If SELRREQUISITN_RS("SER_CODE") <> "8888" Then
                                If SELMA101FLAG = False Then
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                    SELMA101FLAG = True
                                Else
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    SELMA101_RS.MoreResults
                                    SELMA101_RS.Requery
                                End If
                                If Trim(MODALITY) = "SP" Then
                                        If SELDREQFLAG = False Then
                                                SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                                Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                                SELDREQFLAG = True
                                        Else
                                                SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                                SELDREQRS.MoreResults
                                                SELDREQRS.Requery
                                        End If
                                        If SELMA101_RS.RowCount > 0 Then
                                            If Trim(SELRREQUISITN_RS("reqnstat")) = "C" Then
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                                                'chk_collected.Value = 1
                                            Else
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                                            End If
                                            i = i + 1
                                        End If
                                    Else
                                        If SELMA101_RS.RowCount > 0 Then
                                            If Trim(SELRREQUISITN_RS("reqnstat")) = "C" Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                                            GoTo m
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                            GoTo m
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                                End If
                                            'chk_collected.Value = 1
                                            Else
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                            GoTo m
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                            GoTo m
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                                                End If
                                            End If
m:
                                            i = i + 1
                                        End If
                                    End If
                                Else
                                    If Trim(SELRREQUISITN_RS("reqnstat")) = "C" Then
                                        If dept = "NZZ" Then
                                            If Not SELDTMNO_RS.EOF Then
                                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                End If
                                            End If
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                        'chk_collected.Value = 1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                                        End If
                                    Else
                                        If dept = "NZZ" Then
                                            If Not SELDTMNO_RS.EOF Then
                                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                End If
                                            End If
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE"
                                        End If
                                    End If
n:
                                    i = i + 1
                        End If
                        SELRREQUISITN_RS.MoveNext
                    Wend
                    If Trim(MODALITY) = "SP" Or Trim(MODALITY) = "HL" Or Trim(MODALITY) = "BM" Or Trim(MODALITY) = "BF" Then
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                        chk_rcvd.SetFocus
                    Else
                        chk_rcvd.Visible = False
                        txt_pathno.Visible = True
                    End If
                    'chk_collected.Visible = True
                    If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or MODALITY = "TM" Or Trim(MODALITY) = "HL" Or Trim(MODALITY) = "BM" Or Trim(MODALITY) = "BF" Then
                        MSFlexGrid1.row = 1
                        MSFlexGrid1.col = 5
                        chk_rcvd.Width = MSFlexGrid1.CellWidth
                        chk_rcvd.Height = MSFlexGrid1.CellHeight
                        chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    End If
                    btnenabl ("NEW")
                Else
                    MsgBox "No pending Requisition No."
                    'chk_rcvd.Visible = False
                    txt_pathno.Visible = False
                    'btnenabl ("INI")
                    FLDENABL ("DIS")
                    'cmd_new.SetFocus
                    Exit Sub
                End If
        End If
        If viewflag = True Then
                viewflag = False
                If SELRREQUISITNFLAG = False Then
                    SELRREQUISITN_QR(0) = Trim(dept)
                    SELRREQUISITN_QR(1) = Trim(MODALITY)
                    SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                    SELRREQUISITN_QR(3) = msk_todt.FormattedText
                    SELRREQUISITN_QR(4) = "C"
                    SELRREQUISITN_QR(5) = "E"
                    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                    SELRREQUISITNFLAG = True
                Else
                    SELRREQUISITN_QR(0) = Trim(dept)
                    SELRREQUISITN_QR(1) = Trim(MODALITY)
                    SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                    SELRREQUISITN_QR(3) = msk_todt.FormattedText
                    SELRREQUISITN_QR(4) = "C"
                    SELRREQUISITN_QR(5) = "E"
                    SELRREQUISITN_RS.MoreResults
                    SELRREQUISITN_RS.Requery
                End If
                If SELRREQUISITN_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQUISITN_RS.EOF
                        If SELDTMNO_FLAG = False Then
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") ' msk_caseno.FormattedText
                            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                            SELDTMNO_FLAG = True
                        Else
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") 'msk_caseno.FormattedText
                            SELDTMNO_RS.MoreResults
                            SELDTMNO_RS.Requery
                        End If
                        If Trim(SELRREQUISITN_RS("SER_CODE")) <> "8888" Then
                                If SELMA101FLAG = False Then
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                    SELMA101FLAG = True
                                Else
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    SELMA101_RS.MoreResults
                                    SELMA101_RS.Requery
                                End If
                                If SELDREQFLAG = False Then
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                Else
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                End If
                                If SELMA101_RS.RowCount > 0 Then
                                    If Trim(SELRREQUISITN_RS("REQNSTAT")) = "E" Then
                                        If SELDREQRS.RowCount > 0 Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn"))
                                        End If
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
        
                                    End If
                                    i = i + 1
                                End If
                        Else
                                If SELDREQFLAG = False Then
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                Else
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                End If
                                    If Trim(SELRREQUISITN_RS("REQNSTAT")) = "E" Then
                                        If SELDREQRS.RowCount > 0 Then
                                            If dept = "NZZ" Then
                                                If Not SELDTMNO_RS.EOF Then
                                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    Else
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    End If
                                                End If
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            Else
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                            End If
                                        Else
                                            If dept = "NZZ" Then
                                                If Not SELDTMNO_RS.EOF Then
                                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    Else
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    End If
                                                End If
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            Else
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn"))
                                            End If
                                        End If
                                    Else
                                        If dept = "NZZ" Then
                                            If Not SELDTMNO_RS.EOF Then
                                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    GoTo G
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    GoTo G
                                                End If
                                            End If
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                                        End If
                                    End If
G:
                                    i = i + 1
                        End If
                        SELRREQUISITN_RS.MoveNext
                    Wend
                   'chk_rcvd.Visible = False
                   txt_pathno.Visible = False
                End If
        End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Opt_all_Click()
If OPT_ALL.Value = True Then
    msk_fromdt.Text = DATE11
    msk_todt.Text = DATE11
    OPT_CASENO.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    msk_caseno.Enabled = False
    lbl_name.Visible = False
    txt_name.Visible = False
    msk_fromdt.SetFocus
End If

End Sub

Private Sub OPT_CASENO_Click()
If OPT_CASENO.Value = True Then
    OPT_ALL.Value = False
    OPT_ALL.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    'Opt_patient.Value = True
    msk_caseno.Enabled = True
    lbl_name.Visible = True
    txt_name.Visible = True
    msk_caseno.Enabled = True
    msk_caseno.SetFocus
End If
End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    OPT_CASENO.Caption = "Case No."
   ' msk_caseno.Mask = "??>/#####"
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    OPT_CASENO.Caption = "CC No."
    msk_caseno.Mask = "######"
End If
End Sub

Public Sub STAFF()
        If SELEMPBASINFO_FLAG = False Then
             SELEMPBASINFO_QR(0) = msk_caseno.Text
            Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
             SELEMPBASINFO_FLAG = True
        Else
             SELEMPBASINFO_QR(0) = msk_caseno.Text
             SELEMPBASINFO_RS.MoreResults
             SELEMPBASINFO_RS.Requery
        End If
        If SELEMPBASINFO_RS.RowCount > 0 Then
            txt_name = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
        End If
        If dept = "NZZ" And Trim(MODALITY) = "ALL" Then
            DTM1
            Exit Sub
        ElseIf dept = "NZZ" And Trim(MODALITY) <> "ALL" Then
            DTM1MODWISE
            Exit Sub
        End If
        If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = msk_caseno.FormattedText
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
            SELDTMNO_QR(0) = msk_caseno.FormattedText
            SELDTMNO_RS.MoreResults
            SELDTMNO_RS.Requery
        End If
'        CASENO = padzero(5, "C", msk_caseno.Text, 3)
'        msk_caseno.Text = UCase(CASENO)
'                If SELTA1011FLAG = False Then
'                    SELTA1011_QR(0) = msk_caseno.FormattedText
'                    Set SELTA1011_RS = SELTA1011_QR.OpenResultset(1, 3)
'                    SELTA1011FLAG = True
'                Else
'                    SELTA1011_QR(0) = msk_caseno.FormattedText
'                    SELTA1011_RS.MoreResults
'                    SELTA1011_RS.Requery
'                End If
'                If SELTA1011_RS.RowCount > 0 Then
'                    txt_name = Trim(SELTA1011_RS("tag")) & " " & Trim(SELTA1011_RS("first_name")) & " " & Trim(SELTA1011_RS("mid_name")) & " " & Trim(SELTA1011_RS("last_name"))
'                End If


                If NEWFLAG = True Then
                'NEWFLAG = False
                If SELRREQNOFLAG = False Then
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = Trim(msk_caseno.FormattedText)
                    SELRREQNO_QR(3) = "C"
                    SELRREQNO_QR(4) = "P"
                    SELRREQNO_QR(5) = "P"
                    SELRREQNO_QR(6) = "P"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNOFLAG = True
                Else
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "C"
                    SELRREQNO_QR(4) = "P"
                    SELRREQNO_QR(5) = "P"
                    SELRREQNO_QR(6) = "P"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNO_RS.MoreResults
                    SELRREQNO_RS.Requery
                End If
                If SELDTMNO_FLAG = False Then
                    SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
                    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                    SELDTMNO_FLAG = True
                Else
                    SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
                    SELDTMNO_RS.MoreResults
                    SELDTMNO_RS.Requery
                End If
                If SELRREQNO_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQNO_RS.EOF
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                'chk_collected.Value = 1
                            Else
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                            End If
                            i = i + 1
                        End If
                        SELRREQNO_RS.MoveNext
                    Wend
                    'NEW LINE ADDED
                    If Trim(MODALITY) = "SP" Then
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    Else
                        chk_rcvd.Visible = False
                        txt_pathno.Visible = True
                    End If
                    'NEW LINE ADDED
                    
                    'If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Then
                    ' replaced from above code as  all mod were not there sundeep on 16-02-2004
                    
                    If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or MODALITY = "TM" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Then
                        MSFlexGrid1.row = 1
                        MSFlexGrid1.col = 5
                        chk_rcvd.Width = MSFlexGrid1.CellWidth
                        chk_rcvd.Height = MSFlexGrid1.CellHeight
                        chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
''                   ' Else    'COMMENTED
                        'chk_rcvd.Visible = True
''                   '     txt_pathno.Visible = True
                    End If
                    'chk_collected.Visible = True
                Else
                    MsgBox "No pending Requisition No."
                    'chk_rcvd.Visible = False
                    txt_pathno.Visible = False
                    'btnenabl ("INI")
                    FLDENABL ("DIS")
                    'cmd_new.SetFocus
                    Exit Sub
                End If
        End If
        If viewflag = True Then
                viewflag = False
                If SELRREQNOFLAG = False Then
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "C"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNOFLAG = True
                Else
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "C"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNO_RS.MoreResults
                    SELRREQNO_RS.Requery
                End If
                If SELRREQNO_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQNO_RS.EOF
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELDREQFLAG = False Then
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                            SELDREQFLAG = True
                        Else
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            SELDREQRS.MoreResults
                            SELDREQRS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If Trim(SELRREQNO_RS("REQNSTAT")) = "E" Then
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                            Else
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                            End If
                            i = i + 1
                        End If
                        SELRREQNO_RS.MoveNext
                    Wend
                   'chk_rcvd.Visible = False
                   txt_pathno.Visible = False
                   'chk_collected.Visible = False
                End If
        End If

End Sub
'ADDED BY AMIT FOR INTERFACES

Private Sub PRINTBARCODE_Click()
Dim QTY1 As String
Dim BAR_CD As String
Dim EXP_DT As String
Dim ITEM_DESC As String
    Dim TEMPDATA As String
    Dim rack As String
    Dim batchno1 As String
    Dim itemcode1 As String
    Dim i As Long, START
    Dim j As Integer
    Dim NOL As Integer
    Dim Nos As Integer
    'Dim QTY1 As Integer
    Nos = 0
    j = 1
     f1 = FreeFile()
     Open "c:\TATA4.TXT" For Output As #f1
     If BARCODE_NEW = True Then
     SELTIME_QR(0) = REQNO
     End If
     If BARCODE_VIEW = True Then
        SELTIME_QR(0) = II
    End If
     SELTIME_QR.Execute
     Set SELTIME_RS = SELTIME_QR.OpenResultset(1, 3)
     If SELTIME_RS.RowCount > 0 Then
        REC_DATE = Format(SELTIME_RS!SAMRECDATE, "DD/MM/YYYY")
        REC_TIME = Format(SELTIME_RS!SAMRECTIME, "hh:mm:ss")
     Else
        MsgBox "YOU CANNOT PRINT BARCODE LABEL FOR THIS REQUISITN B'COZE NO REQNO EXIST IN RTIMELOG FILE"
        Close #f1
        Exit Sub
    End If
            If BARCODE_NEW = True Then
                SELLABNO_QR(0) = REQNO ' req_chglen(cmb_reqno.Text)
                SELLABNO_QR.Execute
                QTY = 1
                Nos = 0
                Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                Do Until SELLABNO_RS.EOF
                    BAR_CD = Trim(SELLABNO_RS!LABNO)
                    'ITEM_DESC = CASENO
                    CASENO = msk_caseno.FormattedText
                     'REC_DATE = txt_date
                     'REC_TIME = txt_time
                     PATNAME = Trim(txt_name.Text)
                    COLL_CENTER = ""
                    QTY1 = "000" & QTY
                    EXP_DT = "EXP" '& Trim(SELLABNO_RS!TDATE) & Trim(SELLABNO_RS!TTIME)
                    Call Printlbl(BAR_CD)
                 
                    Nos = Nos + QTY
                    SELLABNO_RS.MoveNext
                Loop
                Close #f1
                Shell "c:\tata11.bat", 0
                MsgBox "Total " & Nos & " Labels Have Been Printed", vbInformation
            ElseIf BARCODE_VIEW = True Then
                    If II = "" Then
                        MsgBox "PLEASE CLICK ON REQUISITN NO."
                        Close #f1
                        Exit Sub
                    Else
                    SELLABNO_QR(0) = II ' req_chglen(cmb_reqno.Text)
                    SELLABNO_QR.Execute
                    QTY = 1
                    Nos = 0
                    Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                    Do Until SELLABNO_RS.EOF
                    BAR_CD = Trim(SELLABNO_RS!LABNO)
                    'ITEM_DESC = CASENO
                    CASENO = msk_caseno.FormattedText
                    'REC_DATE = txt_date
                    'REC_TIME = txt_time
                    PATNAME = Trim(txt_name.Text)
                    COLL_CENTER = ""
                    QTY1 = "000" & QTY
                    EXP_DT = "EXP" '& Trim(SELLABNO_RS!TDATE) & Trim(SELLABNO_RS!TTIME)
                    Call Printlbl(BAR_CD)
                    Nos = Nos + QTY
                    SELLABNO_RS.MoveNext
                    Loop
                    End If
                    Close #f1
                    Shell "c:\tata11.bat", 0
                    MsgBox "Total " & Nos & " Labels Have Been Printed", vbInformation
            End If
     Close #f1

End Sub
'ADDED BY AMIT FOR INTERFACES
Public Sub Printlbl(BAR_CD1 As String) ', caseno As String, patname As String, coll_centre As String, REC_DATE As String, rec_time)
''BAR_CD1 = BAR_CD
BAR_CD1 = Trim(BAR_CD1)
bar1 = Mid(BAR_CD1, 1, 2)
bar2 = Mid(BAR_CD1, 3, 8)
bar3 = Mid(BAR_CD1, 11, 1)
'bar4 = mid(bard
'''   NEW LABEL finalised on 30/07/2004
f1 = 1

'Print #f1, "{XC00;0080,0091,A,3,02,0,0160,+0000000000,000,0,00|}"

'ORIGINAL CODE  OLD CODES
'Print #f1, "{XB00;0080,0091,A,3,02,0,0060,+0000000000,000,0,00|}"
'Print #f1, "{PC000;0140,0031,05,05,I,00,B|}"
'Print #f1, "{PC001;0030,0066,05,05,H,00,B|}"
'Print #f1, "{PC002;0089,0187,1,1,H,00,B|}" ' her
'Print #f1, "{PC003;0019,0236,05,05,H,00,B|}"
'Print #f1, "{PC004;0231,0236,05,05,H,00,B|}"
'Print #f1, "{PC005;0041,0143,05,05,I,33,B|}"
'Print #f1, "{PC006;0336,0174,05,05,H,33,B|}"
'Print #f1, "{PC007;0371,0154,05,05,I,33,B|}"
''ENDED ORIGINAL
'Print #f1, "{RB00;>6" + bar1 + ">5" + bar2 + ">6" + bar3 + "|}"
'Print #f1, "{RC000;" + CASENO + "|}"
'Print #f1, "{RC001;" + PATNAME + "|}"
'Print #f1, "{RC002;" + BAR_CD1 + " |}"
''Print #f1, "{RC003;12335678901234567890|}"
'Print #f1, "{RC003;" + coll_centre + "|}"    ' Actual Collection centre to added later when Barcode printers are installed at collection centre
'Print #f1, "{RC004;Biochemistay Lab|}"
'Print #f1, "{RC005;TMH|}"
'Print #f1, "{RC006;" + REC_DATE + "|}"
'Print #f1, "{RC007;" + REC_TIME + "|}"
'Print #f1, "{XS;I,0001,0002C4201|}"
'   OLD CODES
' NEW CODES
'TEC BARCODE PRINTER JOB START
'<xpml><page></xpml>{D0281,0380,0251|}
'<xpml></page></xpml><xpml><page></xpml>{C|}
Print #f1, "{XB00;0030,0040,A,3,03,0,0170,+0000000000,000,0,00|}"
Print #f1, "{PC000;0281,0019,05,05,J,00,B|}"
Print #f1, "{PC001;0020,0036,05,05,H,00,B|}"
Print #f1, "{PC002;0131,0235,05,05,J,00,B|}"
Print #f1, "{PC003;0320,0240,05,05,J,00,B|}"
Print #f1, "{PC004;0030,0019,05,05,J,00,B|}"
Print #f1, "{PC005;0091,0016,05,05,H,00,B|}"
Print #f1, "{PC006;0190,0016,05,05,H,00,B|}"
Print #f1, "{PC007;0020,0240,05,05,J,00,B|}"
Print #f1, "{RB00;>6" + bar1 + ">5" + bar2 + ">6" + bar3 + "|}"
'Print #f1, "{RB00;>6BS>504048725>6D|}"
Print #f1, "{RC000;" + CASENO + "|}"
'Print #f1, "{RC000;TT/01500|}"
Print #f1, "{RC001;" + PATNAME + "|}"
'Print #f1, "{RC001;Mrs. MMMMMM OIOIPOI KXKLJKLJK|}"
Print #f1, "{RC002;" + BAR_CD1 + " |}"
'Print #f1, "{RC002;BS04048725D|}"
Print #f1, "{RC003;FEC|}"
Print #f1, "{RC004;TMH|}"
Print #f1, "{RC005;" + REC_DATE + "|}"
'Print #f1, "{RC005;30H07/2004|}"
Print #f1, "{RC006;" + REC_TIME + "|}"
'Print #f1, "{RC006;11:05|}"
Print #f1, "{RC007;" + "" + "|}"
Print #f1, "{XS;I,0001,0002C4201|}"
' TEC BARCODE PRINTER PAGE END
' <xpml></page></xpml>TEC BARCODE PRINTER JOB END
' <xpml><end/></xpml>


''' NEW CODES





'''' NEW LABEL END


End Sub

Private Sub Timer1_Timer()
'COMMENTED BY CHARULATA ON 14/09/2010 FOR LESS NO. OF CALLS TO SERVER
'datetime
'txt_time.Text = time1
If cy_frame.Visible = True Then
 MSK_CTIME.Text = Format(time1, "HH:MM")
End If
End Sub

Private Sub txt_cyname_Change()

End Sub

Private Sub txt_pathno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    If MSFlexGrid1.row <> 1 Then
    MSFlexGrid1.row = MSFlexGrid1.row - 1
    txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
    
    End If
End If
If KeyCode = 40 Then
    If MSFlexGrid1.row <> MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    End If
End If
End Sub

Private Sub TXT_PATHNO_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim S As String
Dim ST As String
On Error GoTo MSG:
If KeyAscii = 13 Then
ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
'*****
If NEWFLAG = True Then
    

    If Trim(txt_pathno.Text) <> "" Then
        If MODALITY = "BC" Or MODALITY = "EL" Then
            If seldequisitnflag = False Then
                If Len(Trim(txt_pathno.Text)) <= 3 And Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                If Len(Trim(txt_pathno.Text)) <= 3 Or Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        Else
            If seldequisitnflag = False Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        End If
        If MODIFYFLAG = False Then
            If SELDREQUISITNRS.RowCount > 0 Then
                MsgBox "The Path No. is  Given already " & SELDREQUISITNRS("caseno")
                'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                'txt_pathno.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                Exit Sub
                
            End If
        Else
            MODIFYFLAG = True
        End If
    End If

End If
'*****

If Trim(txt_pathno.Text) <> "" Then
  'trim(MSFlexGrid1.TextMatrix(i, 6) = ""
For i = 1 To MSFlexGrid1.Rows - 1
   If NEWFLAG = True Then
    If Trim(txt_pathno.Text) = Trim(MSFlexGrid1.TextMatrix(i, 6)) Then
         MsgBox "Lab No. Already exists"
         Exit Sub
    End If
   End If
   If MODIFYFLAG = True Then
    '*****
    If MODALITY = "BC" Or MODALITY = "EL" Then
            If seldequisitnflag = False Then
                If Len(Trim(txt_pathno.Text)) <= 3 And Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                If Len(Trim(txt_pathno.Text)) <= 3 And Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        Else
            If seldequisitnflag = False Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        End If
   
        If SELDREQUISITNRS.RowCount > 0 Then
                MsgBox "The Path No. is  Given already " & SELDREQUISITNRS("caseno")
                
                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                
                Exit Sub
                
            End If
   
   
   '****

         If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = MSFlexGrid1.TextMatrix(i, 6) Then
         MSFlexGrid1.TextMatrix(i, 6) = ""
         Else
         S = CStr(MSFlexGrid1.TextMatrix(i, 6))
         End If
        If Trim(txt_pathno.Text) = S Then
         MsgBox "Lab No. Already exists"
         Exit Sub
        Else

    End If
   End If
Next i

MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = "COLLECTED"
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = "RECEIVED"
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = txt_pathno.Text
Else
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = txt_pathno.Text


End If
'If Len(Trim(txt_positive)) = 0 Then
'     cmd_save.SetFocus
'     Exit Sub
'
'End If
'
    i = MSFlexGrid1.Rows - 1
    If i = MSFlexGrid1.row Then
        cmd_save.SetFocus
        Exit Sub
    End If
    If MODIFYFLAG = False Then
        If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
        End If
        If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.row = MSFlexGrid1.row + 1
                txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                txt_pathno.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
        End If
    Else
        If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
        End If
        If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.row = MSFlexGrid1.row + 1
                txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                txt_pathno.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        End If
    End If
    Exit Sub
MSG:
    MsgBox Err.Description
    End If
End Sub

Private Sub txt_pathno_Change()
If MODIFYFLAG = True Then
  MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = Trim(txt_pathno.Text)
End If
End Sub

Private Sub txt_pathno_LostFocus()

''If MODIFYFLAG = True Then
''If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)) <> "" Then
''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = txt_pathno
''End If
''End If
''Dim i As Integer
'On Error GoTo msg:
''
''MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = txt_pathno.Text
'''If Len(Trim(txt_positive)) = 0 Then
'''     cmd_save.SetFocus
'''     Exit Sub
'''
'''End If
'''
''i = MSFlexGrid1.Rows - 1
''If i = MSFlexGrid1.Row Then
''    cmd_save.SetFocus
''    Exit Sub
''End If
''If MODIFYFLAG = False Then
''    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
''    End If
''    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
''            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
''            txt_pathno.Visible = True
''            chk_rcvd.Visible = True
''            'chk_rcvd.SetFocus
''            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''    End If
''Else
''    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
''    End If
''    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
''            'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
''            txt_pathno.Visible = True
''            'chk_rcvd.Visible = True
''            'chk_rcvd.SetFocus
''            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''            'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''    End If
''End If
''Exit Sub
''msg:
''MsgBox Err.Description
End Sub

Public Sub BCINF()
''ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
''If seldmulprocflag = False Then
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
''     seldmulprocflag = True
''Else
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     seldmulprocrs.MoreResults
''     seldmulprocrs.Requery
''End If
''If seldmulprocrs.RowCount > 0 Then
''
''     While Not seldmulprocrs.EOF
''           If seldmulprocrs("pro_code") = "F849" Or seldmulprocrs("PRO_CODE") = "F854" Or seldmulprocrs("PRO_CODE") = "F860" Or seldmulprocrs("PRO_CODE") = "F865" Or seldmulprocrs("PRO_CODE") = "F875" Or seldmulprocrs("PRO_CODE") = "F880" Or seldmulprocrs("PRO_CODE") = "F802" Then
''
''                    If SELDBCPROSUBFLAG = False Then
''                        SELDBCPROSUBQR(0) = seldmulprocrs("pro_code")
''                        Set SELDBCPROSUBRS = SELDBCPROSUBQR.OpenResultset(1, 3)
''                        SELDBCPROSUBFLAG = True
''                    Else
''                        SELDBCPROSUBQR(0) = seldmulprocrs("pro_code")
''                        SELDBCPROSUBRS.MoreResults
''                        SELDBCPROSUBRS.Requery
''                    End If
''
''                    If SELDBCPROSUBRS.RowCount > 0 Then
''                        While Not SELDBCPROSUBRS.EOF
''                            INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(sRow, 1))
''                            If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) > 3 Then
''                                INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                            If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''                                INSDBCINFACEQR(1) =st & MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                            INSDBCINFACEQR(2) = SELDBCPROSUBRS("procodesub")
''                            INSDBCINFACEQR.Execute
''
''                           SELDBCPROSUBRS.MoveNext
''                        Wend
''                 End If
''           Else
''
''                          INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''                           If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) > 3 Then
''                                INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                            If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''                                INSDBCINFACEQR(1) = st & MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                         INSDBCINFACEQR(2) = seldmulprocrs("procodesub")
''                         INSDBCINFACEQR.Execute
''           End If
''        seldmulprocrs.MoveNext
''     Wend
''End If
End Sub

Public Sub ELINF()
''If seldmulprocflag = False Then
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
''     seldmulprocflag = True
''Else
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     seldmulprocrs.MoreResults
''     seldmulprocrs.Requery
''End If
''If seldmulprocrs.RowCount > 0 Then
''    While Not seldmulprocrs.EOF
''       If seldmulprocrs("PRO_CODE") = "F860" Or seldmulprocrs("PRO_CODE") = "F875" Or seldmulprocrs("PRO_CODE") = "F901" Or seldmulprocrs("PRO_CODE") = "F903" Or seldmulprocrs("PRO_CODE") = "F908" Then
''            If SELELPROSUBFLAG = False Then
''                SELELPROSUBQR(0) = Trim(seldmulprocrs("PRO_CODE"))
''                Set SELELPROSUBRS = SELELPROSUBQR.OpenResultset(1, 3)
''                SELELPROSUBFLAG = True
''            Else
''                SELELPROSUBQR(0) = Trim(seldmulprocrs("PRO_CODE"))
''                SELELPROSUBRS.MoreResults
''                SELELPROSUBRS.Requery
''            End If
''            If SELELPROSUBRS.RowCount > 0 Then
''                While Not SELELPROSUBRS.EOF
''                        INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''                        'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 3 Then
''                            INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''                        'End If
''                        'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''                         '   INSDBCINFACEQR(1) = ST & MSFlexGrid1.TextMatrix(srow, 6)
''                        'End If
''                        INSDBCINFACEQR(2) = SELELPROSUBRS("procodesub")
''                        INSDBCINFACEQR.Execute
''
''                    SELELPROSUBRS.MoveNext
''                Wend
''            End If
''       Else
''            INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''            'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 3 Then
''                INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''            'End If
''            'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''            '    INSDBCINFACEQR(1) = ST & MSFlexGrid1.TextMatrix(srow, 6)
''            'End If
''            INSDBCINFACEQR(2) = seldmulprocrs("PRO_CODE")
''            INSDBCINFACEQR.Execute
''
''       End If
''       seldmulprocrs.MoveNext
''   Wend
''End If
End Sub

Public Sub TMINF()

End Sub
Public Sub DTM()

If NEWFLAG = True Then
    If SELRREQUISITNFLAG1 = False Then
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "C"
        SELRREQUISITN_QR1(3) = "P"
        SELRREQUISITN_QR1(4) = "C"
        
        Set SELRREQUISITN_RS1 = SELRREQUISITN_QR1.OpenResultset(1, 3)
        SELRREQUISITNFLAG1 = True
    Else
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "C"
        SELRREQUISITN_QR1(3) = "P"
        SELRREQUISITN_QR1(4) = "C"
        SELRREQUISITN_RS1.MoreResults
        SELRREQUISITN_RS1.Requery
    End If
    If SELRREQUISITN_RS1.RowCount = 0 Then
        Exit Sub
    End If
    If SELRREQUISITN_RS1.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS1.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELRREQUISITN_RS1("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS1("reqnstat")) = "C" Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                            'chk_collected.Value = 1
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS1("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                            End If
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                            End If
                        End If
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQUISITN_RS1("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE"
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS1.MoveNext
        Wend
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
            chk_rcvd.SetFocus
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        'chk_collected.Visible = True
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
    MSFlexGrid1.row = 1
    MSFlexGrid1.col = 5
    chk_rcvd.Width = MSFlexGrid1.CellWidth
    chk_rcvd.Height = MSFlexGrid1.CellHeight
    chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    chk_rcvd.Visible = True
    txt_pathno.Visible = False
End If
If viewflag = True Then
    viewflag = False
    If SELRREQUISITNFLAG1 = False Then
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "E"
        SELRREQUISITN_QR1(3) = "T"
        SELRREQUISITN_QR1(4) = "F"
        Set SELRREQUISITN_RS1 = SELRREQUISITN_QR1.OpenResultset(1, 3)
        SELRREQUISITNFLAG1 = True
    Else
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "E"
        SELRREQUISITN_QR1(3) = "T"
        SELRREQUISITN_QR1(4) = "F"
        SELRREQUISITN_RS1.MoreResults
        SELRREQUISITN_RS1.Requery
    End If
    If SELRREQUISITN_RS1.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS1.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If Trim(SELRREQUISITN_RS1("SER_CODE")) <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    If Trim(SELRREQUISITN_RS1("REQNSTAT")) = "E" Then
                        If SELDREQRS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn"))
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                    End If
                    i = i + 1
                End If
            Else
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If Trim(SELRREQUISITN_RS1("REQNSTAT")) = "E" Then
                    If SELDREQRS.RowCount > 0 Then
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        End If
                    Else
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        End If
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("dtmno")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS1.MoveNext
        Wend
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
    End If
End If
End Sub
Public Sub DTMMODWISE()

If NEWFLAG = True Then
    If SELRREQUISITN_FLAG11 = False Then
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "C"
        SELRREQUISITN_QR11(4) = "P"
        SELRREQUISITN_QR11(5) = "C"
        Set SELRREQUISITN_RS11 = SELRREQUISITN_QR11.OpenResultset(1, 3)
        SELRREQUISITN_FLAG11 = True
    Else
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "C"
        SELRREQUISITN_QR11(4) = "P"
        SELRREQUISITN_QR11(5) = "C"
        SELRREQUISITN_RS11.MoreResults
        SELRREQUISITN_RS11.Requery
    End If
    If SELRREQUISITN_RS11.RowCount = 0 Then
        Exit Sub
    End If
    If SELRREQUISITN_RS11.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS11.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELRREQUISITN_RS11("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS11("reqnstat")) = "C" Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                            'chk_collected.Value = 1
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS11("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                            End If
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                            End If
                        End If
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQUISITN_RS11("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE"
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS11.MoveNext
        Wend
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
            chk_rcvd.SetFocus
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        'chk_collected.Visible = True
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
    MSFlexGrid1.row = 1
    MSFlexGrid1.col = 5
    chk_rcvd.Width = MSFlexGrid1.CellWidth
    chk_rcvd.Height = MSFlexGrid1.CellHeight
    chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    chk_rcvd.Visible = True
    txt_pathno.Visible = False
End If
If viewflag = True Then
    viewflag = False
    If SELRREQUISITN_FLAG11 = False Then
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "E"
        SELRREQUISITN_QR11(4) = "T"
        SELRREQUISITN_QR11(5) = "F"
        Set SELRREQUISITN_RS11 = SELRREQUISITN_QR11.OpenResultset(1, 3)
        SELRREQUISITN_FLAG11 = True
    Else
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "E"
        SELRREQUISITN_QR11(4) = "T"
        SELRREQUISITN_QR11(5) = "F"
        SELRREQUISITN_RS11.MoreResults
        SELRREQUISITN_RS11.Requery
    End If
    If SELRREQUISITN_RS11.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS11.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If Trim(SELRREQUISITN_RS11("SER_CODE")) <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    If Trim(SELRREQUISITN_RS11("REQNSTAT")) = "E" Then
                        If SELDREQRS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn"))
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn"))
                    End If
                    i = i + 1
                End If
            Else
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If Trim(SELRREQUISITN_RS11("REQNSTAT")) = "E" Then
                    If SELDREQRS.RowCount > 0 Then
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    Else
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("dtmno")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '" " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS11.MoveNext
        Wend
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
    End If
End If
End Sub

Public Sub DTM1()
If SELDTMNO_FLAG = False Then
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    SELDTMNO_FLAG = True
Else
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
End If
If NEWFLAG = True Then
    'NEWFLAG = False
    If SELRREQNOFLAG1 = False Then
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "C"
        SELRREQNO_QR1(2) = "C"
        SELRREQNO_QR1(3) = "P"   ' Made P from C on 27-01-2004
        SELRREQNO_QR1(4) = "C"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNOFLAG1 = True
    Else
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "C"
        SELRREQNO_QR1(2) = "C"
        SELRREQNO_QR1(3) = "P"  ' Made P from C on 27-01-2004
        SELRREQNO_QR1(4) = "C"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNO_RS1.MoreResults
        SELRREQNO_RS1.Requery
    End If
    If SELRREQNO_RS1.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQNO_RS1.EOF
            If SELRREQNO_RS1("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS1("reqnstat")) = "C" Then
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & SELDREQRS("PATHNO")
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & ""
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            End If
                        Else
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDREQRS("PATHNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & ""
                            End If
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS1("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                        End If
n:
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQNO_RS1("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                End If
m:
                i = i + 1
            End If
            SELRREQNO_RS1.MoveNext
        Wend
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Value = 0
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
        btnenabl ("NEW")
        'chk_collected.Visible = True
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
End If
If viewflag = True Then
'viewflag = False
If SELRREQNOFLAG1 = False Then
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "E"
        SELRREQNO_QR1(2) = "E"
        SELRREQNO_QR1(3) = "T"
        SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNOFLAG1 = True
    Else
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "E"
        SELRREQNO_QR1(2) = "E"
        SELRREQNO_QR1(3) = "T"
        SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNO_RS1.MoreResults
        SELRREQNO_RS1.Requery
    End If
        If SELRREQNO_RS1.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQNO_RS1.EOF
                If SELRREQNO_RS1("SER_CODE") <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If SELDREQFLAG = False Then
                                SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                                Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                SELDREQFLAG = True
                            Else
                                SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                                SELDREQRS.MoreResults
                                SELDREQRS.Requery
                            End If
                            If Trim(SELRREQNO_RS1("REQNSTAT")) = "E" Then
                                If SELDREQRS.RowCount > 0 Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                Else
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                End If
                            Else
                                If Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                End If
                            End If
m1:
                        i = i + 1
                        End If
                Else
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If Trim(SELRREQNO_RS1("REQNSTAT")) = "E" Then
                            If SELDREQRS.RowCount > 0 Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                            Else
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            If Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                            End If
                        End If
N1:
                     i = i + 1
                End If
                SELRREQNO_RS1.MoveNext
            Wend
           'chk_rcvd.Visible = False
           txt_pathno.Visible = False
           'chk_collected.Visible = False
        End If
    End If
'If opt_staff.Value = True Then
'    STAFF
'End If

End Sub
Public Sub DTM1MODWISE()
If SELDTMNO_FLAG = False Then
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    SELDTMNO_FLAG = True
Else
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
End If
If NEWFLAG = True Then
    'NEWFLAG = False
    If SELRREQNO_FLAG11 = False Then
        SELRREQNO_QR11(0) = MODALITY
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "C"
        SELRREQNO_QR11(3) = "P"   ' Made P from C on 27-01-2004
        SELRREQNO_QR11(4) = "C"
        'SELRREQNO_QR1(4) = "P"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNO_FLAG11 = True
    Else
        SELRREQNO_QR11(0) = Trim(MODALITY)
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "C"
        SELRREQNO_QR11(3) = "P"   ' Made P from C on 27-01-2004
        SELRREQNO_QR11(4) = "C"
        'SELRREQNO_QR1(4) = "P"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNO_RS11.MoreResults
        SELRREQNO_RS11.Requery
    End If
    If SELRREQNO_RS11.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQNO_RS11.EOF
            If SELRREQNO_RS11("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS11("reqnstat")) = "C" Then
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & SELDREQRS("PATHNO")
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & ""
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            End If
                        Else
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDREQRS("PATHNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & ""
                            End If
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS11("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                        End If
n:
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQNO_RS11("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                End If
m:
                i = i + 1
            End If
            SELRREQNO_RS11.MoveNext
        Wend
        
''''        If Trim(MODALITY) = "SP" Then
''''            chk_rcvd.Visible = True
''''            txt_pathno.Visible = False
''''        Else
''''            chk_rcvd.Visible = False
''''            txt_pathno.Visible = True
''''        End If
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
            chk_rcvd.SetFocus
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        'chk_collected.Visible = True
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If

        
''        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
''            MSFlexGrid1.Row = 1
''            MSFlexGrid1.Col = 5
''            chk_rcvd.Width = MSFlexGrid1.CellWidth
''            chk_rcvd.Height = MSFlexGrid1.CellHeight
''            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
''            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''            chk_rcvd.Value = 0
''            chk_rcvd.Visible = True
''            txt_pathno.Visible = False
''        End If
        btnenabl ("NEW")
        'chk_collected.Visible = True
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
End If
If viewflag = True Then
'viewflag = False
If SELRREQNO_FLAG11 = False Then
        SELRREQNO_QR11(0) = MODALITY
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "E"
        SELRREQNO_QR11(3) = "T"
        SELRREQNO_QR11(4) = "E"
        'SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNOFLAG1 = True
    Else
        SELRREQNO_QR11(0) = MODALITY
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "E"
        SELRREQNO_QR11(3) = "T"
        SELRREQNO_QR11(4) = "F"
        'SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNO_RS11.MoreResults
        SELRREQNO_RS11.Requery
    End If
        If SELRREQNO_RS11.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQNO_RS11.EOF
                If SELRREQNO_RS11("SER_CODE") <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If SELDREQFLAG = False Then
                                SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                                Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                SELDREQFLAG = True
                            Else
                                SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                                SELDREQRS.MoreResults
                                SELDREQRS.Requery
                            End If
                            If Trim(SELRREQNO_RS11("REQNSTAT")) = "E" Then
                                If SELDREQRS.RowCount > 0 Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                Else
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                End If
                            Else
                                If Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                End If
                            End If
m1:
                        i = i + 1
                        End If
                Else
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If Trim(SELRREQNO_RS11("REQNSTAT")) = "E" Then
                            If SELDREQRS.RowCount > 0 Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                            Else
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            If Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                            End If
                        End If
N1:
                     i = i + 1
                End If
                SELRREQNO_RS11.MoveNext
            Wend
           'chk_rcvd.Visible = False
           txt_pathno.Visible = False
           'chk_collected.Visible = False
        End If
    End If
'If opt_staff.Value = True Then
    'STAFF
'End If

End Sub

Public Sub SPAUTONO()
Dim PATH As String
On Error GoTo MSG:
i = 1
For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(i, 5)) <> "" Then
        If SELDAUTOPATHFLAG = False Then
            SELDAUTOPATH_QR(0) = Trim(dept)
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
            SELDAUTOPATHFLAG = True
        Else
            SELDAUTOPATH_QR(0) = Trim(dept)
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            SELDAUTOPATH_RS.MoreResults
            SELDAUTOPATH_RS.Requery
        End If
        If SELDAUTOPATH_RS.RowCount > 0 Then
            AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
            UPDDAUTOPATH_QR(0) = AUTOPATHNO
            UPDDAUTOPATH_QR(1) = Trim(dept)
            UPDDAUTOPATH_QR(2) = Trim(MODALITY)
            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDDAUTOPATH_QR.Execute
        Else
            AUTOPATHNO = 1
            INSDAUTOPATH_QR(0) = Trim(dept)
            INSDAUTOPATH_QR(1) = Trim(MODALITY)
            INSDAUTOPATH_QR(2) = Trim(Left(msk_caseno.FormattedText, 2))
            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            INSDAUTOPATH_QR(4) = 1
            INSDAUTOPATH_QR.Execute
        End If
        If Trim(dept) = "UZZ" Then
                PATH = Trim(SELDAUTOPATH_RS("PATHYEAR"))
        Else
            If DPATHYEARFLAG = False Then
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
                DPATHYEARFLAG = True
            Else
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                DPATHYEAR_RS.MoreResults
                DPATHYEAR_RS.Requery
            End If
            If DPATHYEAR_RS.RowCount > 0 Then
                PATH = DPATHYEAR_RS("LASTTWO")
            End If
        End If
    '***********padding the reqslno with leading zeros"****************
        Dim reqslno1
        reqslno1 = AUTOPATHNO
        If Trim(dept) = "UZZ" Then
            While Len(CStr(reqslno1)) < 6
            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
                reqslno1 = "0" & reqslno1
            Wend
        ElseIf Trim(dept) = "VZZ" Then
            While Len(CStr(reqslno1)) < 4
                reqslno1 = "0" & reqslno1
            Wend
            reqslno1 = "CG" & reqslno1
        Else
            While Len(CStr(reqslno1)) < 6
                reqslno1 = "0" & reqslno1
            Wend
        End If
        DAUTOPATHNO = reqslno1
'******************************************************************
        'If Trim(MODALITY) = "SP" Then
        If Trim(dept) = "UZZ" Then
            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
            MSFlexGrid1.TextMatrix(i, 6) = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
        Else
            MSFlexGrid1.TextMatrix(i, 6) = DAUTOPATHNO & "/" & PATH
        End If
        'End If
    End If
Next i
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub clearcy()
   chk_clotted.Value = 0
   chk_collected.Value = 0
   chk_dilute.Value = 0
   chk_hyper.Value = 0
   chk_leaked.Value = 0
   chk_ok.Value = 0
   chk_rcvd.Value = 0
End Sub
