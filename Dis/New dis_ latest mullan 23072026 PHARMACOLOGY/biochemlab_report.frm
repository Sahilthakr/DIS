VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form biochemlab_report 
   Caption         =   "Biochemistry Laboratory Report"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame REPORT 
      Height          =   4455
      Left            =   120
      TabIndex        =   48
      Top             =   4320
      Width           =   11775
      Begin VB.ComboBox cmb_authSign 
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
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   99
         Top             =   4080
         Width           =   2295
      End
      Begin VB.ComboBox cmb_Gtt_UrineValue 
         Height          =   315
         ItemData        =   "biochemlab_report.frx":0000
         Left            =   3720
         List            =   "biochemlab_report.frx":001C
         TabIndex        =   87
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txt_gttobsvdurn 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   3720
         TabIndex        =   80
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txt_gttobsvd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   1560
         TabIndex        =   78
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1815
         Left            =   360
         TabIndex        =   77
         Top             =   840
         Visible         =   0   'False
         Width           =   5320
         _ExtentX        =   9393
         _ExtentY        =   3201
         _Version        =   393216
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   4210816
         BackColorSel    =   4210752
         GridColorFixed  =   -2147483630
         FormatString    =   "Time Slots     |^    Blood Sugar          "
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
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   6240
         TabIndex        =   75
         Top             =   4080
         Width           =   2775
      End
      Begin VB.ComboBox cmb_machine 
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
         Height          =   315
         ItemData        =   "biochemlab_report.frx":006E
         Left            =   10440
         List            =   "biochemlab_report.frx":0084
         TabIndex        =   72
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmd_adenum 
         Caption         =   "&Addenum"
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
         Left            =   10680
         TabIndex        =   71
         Top             =   3240
         Width           =   975
      End
      Begin VB.ComboBox TXT_COMMENTS 
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
         Left            =   1680
         TabIndex        =   68
         Top             =   2880
         Width           =   8895
      End
      Begin VB.TextBox txt_addendum 
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
         Left            =   1680
         TabIndex        =   59
         Top             =   3360
         Visible         =   0   'False
         Width           =   8895
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
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   3720
         Width           =   2295
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
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txt_observed 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   3480
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ForeColorFixed  =   192
         AllowUserResizing=   3
         FormatString    =   "S.No.|^Procedure Name                   |^Observed Value|^Normal Range of the Procedure |^ Panic Value|^Source|^Preval|^MachineNo."
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
      Begin VB.Frame STONEFRM 
         Caption         =   "STONE ANALYSIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   82
         Top             =   120
         Visible         =   0   'False
         Width           =   11535
         Begin VB.ComboBox ser_code 
            Height          =   315
            Left            =   6720
            TabIndex        =   84
            Top             =   2040
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox STONETXT 
            Height          =   975
            Left            =   2280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   83
            Top             =   360
            Width           =   6615
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Authorised Signatory"
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
         TabIndex        =   98
         Top             =   4080
         Width           =   1770
      End
      Begin VB.Label Label4 
         Caption         =   "Addendum by"
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
         Left            =   4920
         TabIndex        =   74
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lbl_addendnum 
         Caption         =   "Addendum"
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
         TabIndex        =   73
         Top             =   3480
         Width           =   975
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
         Height          =   195
         Left            =   840
         TabIndex        =   60
         Top             =   2880
         Width           =   750
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Entered  by"
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
         Left            =   1080
         TabIndex        =   58
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label lbl_final 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized By"
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
         Left            =   4560
         TabIndex        =   57
         Top             =   3840
         Width           =   1590
      End
   End
   Begin VB.CheckBox chk_pno 
      Caption         =   "PathnoWise"
      Height          =   375
      Left            =   12480
      TabIndex        =   96
      Top             =   480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      TabIndex        =   92
      Top             =   3000
      Width           =   270
   End
   Begin VB.CommandButton CMD_ADDENDUM 
      Caption         =   "&Addendum"
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
      Left            =   10560
      MouseIcon       =   "biochemlab_report.frx":00AA
      MousePointer    =   99  'Custom
      TabIndex        =   91
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmd_refresh 
      Caption         =   "Repeat"
      Height          =   375
      Left            =   8640
      TabIndex        =   89
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1080
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   120
      TabIndex        =   43
      Top             =   8760
      Width           =   11775
      Begin VB.CommandButton CMD_NEW 
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
         Left            =   720
         MouseIcon       =   "biochemlab_report.frx":03B4
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
         Left            =   1920
         MouseIcon       =   "biochemlab_report.frx":06BE
         MousePointer    =   99  'Custom
         TabIndex        =   47
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
         MouseIcon       =   "biochemlab_report.frx":09C8
         MousePointer    =   99  'Custom
         TabIndex        =   46
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
         Left            =   4320
         MouseIcon       =   "biochemlab_report.frx":0CD2
         MousePointer    =   99  'Custom
         TabIndex        =   7
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
         Left            =   5520
         MouseIcon       =   "biochemlab_report.frx":0FDC
         MousePointer    =   99  'Custom
         TabIndex        =   8
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
         Left            =   8040
         MouseIcon       =   "biochemlab_report.frx":12E6
         MousePointer    =   99  'Custom
         TabIndex        =   9
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
         Left            =   9240
         MouseIcon       =   "biochemlab_report.frx":15F0
         MousePointer    =   99  'Custom
         TabIndex        =   45
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
         Left            =   6720
         MouseIcon       =   "biochemlab_report.frx":18FA
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
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
      Left            =   10440
      TabIndex        =   42
      Top             =   600
      Width           =   1095
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
      Left            =   10440
      TabIndex        =   41
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame common 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   11775
      Begin VB.CommandButton CMD_EMR 
         BackColor       =   &H00FFFF80&
         Caption         =   "&EMR"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   63
         Top             =   2520
         Width           =   270
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         TabIndex        =   62
         Top             =   2520
         Width           =   270
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   61
         Top             =   2520
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   1215
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
            Width           =   975
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
            Width           =   735
         End
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   23
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
         ItemData        =   "biochemlab_report.frx":1C04
         Left            =   6240
         List            =   "biochemlab_report.frx":1C06
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_pathno 
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
         TabIndex        =   22
         Top             =   600
         Width           =   1335
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_relclnhis 
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
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2520
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
      Begin VB.ComboBox CMB_PATHNO 
         Height          =   315
         Left            =   10200
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Addendum Value"
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
         Left            =   9480
         TabIndex        =   93
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Label LBL_OPPROC 
         Caption         =   "Label11"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6360
         TabIndex        =   90
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Normal Value"
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
         Left            =   6840
         TabIndex        =   66
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Low Value"
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
         Left            =   8760
         TabIndex        =   65
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "High Value"
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
         Left            =   10440
         TabIndex        =   64
         Top             =   2520
         Width           =   945
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_pathno 
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
         Left            =   9360
         TabIndex        =   36
         Top             =   600
         Width           =   690
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_specmen 
         AutoSize        =   -1  'True
         Caption         =   "Specimen"
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
         TabIndex        =   24
         Top             =   2520
         Width           =   840
      End
   End
   Begin VB.Label Label10 
      Caption         =   "SMS-critical alert"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Entered  by"
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
      TabIndex        =   97
      Top             =   0
      Width           =   990
   End
   Begin VB.Label Lbl_pno 
      Height          =   495
      Left            =   12240
      TabIndex        =   95
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "PCS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11640
      TabIndex        =   94
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "STAT"
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
      Left            =   1680
      TabIndex        =   88
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   1320
      TabIndex        =   86
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "02/11/21"
      Height          =   255
      Left            =   100
      TabIndex        =   85
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
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
      TabIndex        =   79
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "HOMI BHABHA CANCER  HOSPITAL AND RESEARCH CENTRE,MULLANPUR"
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
      Height          =   330
      Left            =   2175
      TabIndex        =   70
      Top             =   0
      Width           =   10545
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
      TabIndex        =   51
      Top             =   600
      Width           =   420
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
      TabIndex        =   50
      Top             =   960
      Width           =   420
   End
   Begin VB.Label lbl_biochemlab 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - BIOCHEMISTRY LABORATORY"
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
      Left            =   3000
      TabIndex        =   49
      Top             =   600
      Width           =   5835
   End
   Begin VB.Label lbl_rwd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RECOVERY WARD"
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
      Height          =   375
      Left            =   120
      TabIndex        =   76
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
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
      Left            =   120
      TabIndex        =   69
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "biochemlab_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PC As String
Dim R1, C1
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
Dim dm As Boolean
Dim DOBDET As String
'ADDED BY MILIND DT. 03/03/06
Dim userid_qr As rdoQuery
Dim userid_rs As rdoResultset
Dim userdesc_qr As rdoQuery
Dim userdesc_rs As rdoResultset
'upto here
Dim PREVTEST As String
Dim SELMSADINTF_QR As rdoQuery
Dim SELMSADINTF_RS As rdoResultset

Dim SELDTEMPLETR1 As rdoQuery
Dim SELDTEMPLETS1 As rdoResultset

'SUMOL NOV 2006 STARTS HERE
Dim INSSTONEQR As rdoQuery
Dim SELSTONEQR As rdoQuery
Dim SELSTONERS As rdoResultset
Dim SELSTONEFLAG As Boolean
Dim UPDSTONEQR As rdoQuery
Dim w_operationproc As Application
Dim d_operationproc As Document
'ENDS HERE

Dim sqlinterface As String
Dim rqsqlinterface As rdoQuery
Dim Rssqlinterface As rdoResultset

'ADDED BY MILIND DT. 06/09/06
Dim gtt_qr As rdoQuery
Dim gtt_rs As rdoResultset
'upto here

'ADDED BY MILIND DT. 23/01/06
Dim reqspn_qr As rdoQuery
Dim reqspn_rs As rdoResultset
'upto here

'GGT
Dim SELDBCALKPHSGGTQR As rdoQuery
Dim SELDBCALKPHSGGTRS As rdoResultset
Dim SELDBCALKPHSGGTFLAG As Boolean

'for GTT Test
Dim insgttdtlsqr As rdoQuery
Dim gttflag As Boolean
Dim UPDGTTDTLSQR As rdoQuery
Dim SELGTTDTLSQR As rdoQuery
Dim SELGTTDTLSRS As rdoResultset


'SAVE THE DATE AND TIME INTO RTIMELOG
Dim UPDRTIME_QR As rdoQuery
Dim UPD_RS As rdoResultset

Dim UPDRTIMECADD_QR As rdoQuery

Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEC_RS As rdoResultset

Dim UPDRTIMEP_QR As rdoQuery
Dim UPDRTIMEM_QR As rdoQuery
'till here


Dim REPDRFLAG As Boolean
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Public ICASENO, IREQNO
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
'Dim sqlinterface As String
'Dim rqsqlinterface As rdoQuery
'Dim Rssqlinterface As rdoResultset
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
'Dim INSDADDENDUM_QR As rdoQuery
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



Dim INSDMPROCEZ_QR As rdoQuery 'Shraddha 20/07/10
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
Dim RNAGE As Integer
Dim RNAGEAL2 As Double
Dim SS As String
Dim ALKLV As Integer
Dim ALKHV As Integer
Dim LV As Integer
Dim HV As Integer
Dim PROTEIN As Variant
Dim GLOBULIN As Variant

''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean

''*******ICU**********
''*****ADDED BY MONO
Dim SELDTEMPLETQR1 As rdoQuery
Dim SELDTEMPLETRS1 As rdoResultset
Dim SELDTEMPLETFLAG1 As Boolean

''******SECURITY***********

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String

Dim REPWTRNRS As rdoResultset
Dim REPWTRN_QR As rdoQuery

Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_FLAG As Boolean
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset
' INTERFACE
Dim SELINTER_QR As rdoQuery
Dim SELINTER_RS As rdoResultset

Dim SELINTER_manual_QR As rdoQuery
Dim SELINTER_manual_RS As rdoResultset

Dim SELINTER_FLAG As Boolean
Dim INTERFACEREAD_QR As rdoQuery
Dim INTERFACEUPD_QR As rdoQuery
Dim INTERFACEREAD_RS As rdoResultset
Dim UPDINTERFACE_QR As rdoQuery
Dim UPDINTERFACE_RS As rdoResultset
Dim SELLABNO_QR As rdoQuery
Dim SELLABNO_RS As rdoResultset
'INTERFACE ENDS


Dim INSDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUMFLAG As Boolean
'MACHINE
Dim SELDBCALKPHS_MQR As rdoQuery
Dim SELDBCALKPHS_MRS As rdoResultset
Dim SELDBCALKPHS_MFLAG As Boolean

Dim SELDTEMPLET_MQR As rdoQuery
Dim SELDTEMPLET_MRS As rdoResultset
Dim SELDTEMPLET_MFLAG As Boolean

Dim SELDTEMPLET_MQR1 As rdoQuery
Dim SELDTEMPLET_MRS1 As rdoResultset
Dim SELDTEMPLET_MFLAG1 As Boolean

Dim VAR_REQNNO As String
Dim adenum_flg As Boolean

'API DECLARATION BY FOR BARCODE PRINTING
Private Declare Function WaitForSingleObject Lib "Kernal32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "Kernal32" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "Kernal32" (ByVal hObject As Long) As Long
''******SECURITY***********


Private Sub cmb_addendnum_Change()

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
        cmb_modality.ListIndex = 0
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "FEC" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub


Private Sub cmb_Gtt_UrineValue_Click()
On Error Resume Next

Dim obs_val As String

obs_val = Trim(cmb_Gtt_UrineValue.Text)

'If KeyCode = 13 Then
    MSFlexGrid2.col = 2

    MSFlexGrid2.TextMatrix(MSFlexGrid2.row, 2) = obs_val
    cmb_Gtt_UrineValue.Visible = False
    'cmb_Gtt_UrineValue.Text = ""
    'cmb_Gtt_UrineValue.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left
    
'End If

'If KeyCode = 39 Then
'
'    If cmb_Gtt_UrineValue.Text <> "" Then
'        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, 2) = obs_val
'        cmb_Gtt_UrineValue.Text = ""
'    End If
'
'    MSFlexGrid2.col = MSFlexGrid2.col + 1
'    cmb_Gtt_UrineValue.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left
'
'End If
'
'If KeyCode = 40 Then
'
'    If MSFlexGrid2.row = MSFlexGrid2.Rows - 1 Then
'        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
'    End If
'
'    If cmb_Gtt_UrineValue.Text <> "" Then
'        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, 2) = obs_val
'        cmb_Gtt_UrineValue.Text = ""
'    End If
'
'    MSFlexGrid2.row = MSFlexGrid2.row + 1
'    cmb_Gtt_UrineValue.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top
'
'End If
'
'If KeyCode = 37 Then
'
'    If cmb_Gtt_UrineValue.Text <> "" Then
'        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, 2) = obs_val
'        cmb_Gtt_UrineValue.Text = ""
'    End If
'
'    MSFlexGrid2.col = MSFlexGrid2.col - 1
'
'    If MSFlexGrid2.col <> 0 Then
'        cmb_Gtt_UrineValue.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left
'    End If
'
'End If
'
'If KeyCode = 38 Then
'
'    If cmb_Gtt_UrineValue.Text <> "" Then
'        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, 2) = obs_val
'        'Text2.Text = ""
'    End If
'
'    MSFlexGrid2.row = MSFlexGrid2.row - 1
'    cmb_Gtt_UrineValue.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top
'
'End If
'
End Sub

Private Sub cmb_machine_Click()
'Added by Milind dt 19/01/06
Dim I As Integer
Dim j As Integer
I = MSFlexGrid1.Rows
I = I - 1

If MSFlexGrid1.row = 1 Then
    For j = 1 To I
        MSFlexGrid1.TextMatrix(j, 7) = cmb_machine.Text
    Next
End If
' upto here

Dim num As Integer
Dim sc As String

cmb_machine.Visible = False
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(cmb_machine.Text)
num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)

'MsgBox SCODE
If Opt_patient.Value = True Then
  If sc = "F869" Then 'ALKAGE
    
      If SELDBCALKPHS_MFLAG = False Then
            SELDBCALKPHS_MQR(0) = sc
            SELDBCALKPHS_MQR(1) = RNAGE
            SELDBCALKPHS_MQR(2) = SS
            SELDBCALKPHS_MQR(3) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
            Set SELDBCALKPHS_MRS = SELDBCALKPHS_MQR.OpenResultset(1, 3)
            SELDBCALKPHS_MFLAG = True
      Else
            SELDBCALKPHS_MQR(0) = sc
            SELDBCALKPHS_MQR(1) = RNAGE
            SELDBCALKPHS_MQR(2) = SS
            SELDBCALKPHS_MQR(3) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
            SELDBCALKPHS_MRS.MoreResults
            SELDBCALKPHS_MRS.Requery
       End If
       
       If SELDBCALKPHS_MRS.RowCount > 0 Then
         MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = Trim(SELDBCALKPHS_MRS!range)
       Else
         MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = ""
       End If
        
    
  Else
  
     If SELDTEMPLET_MFLAG = False Then
               SELDTEMPLET_MQR(0) = sc
               SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
               Set SELDTEMPLET_MRS = SELDTEMPLET_MQR.OpenResultset(1, 3)
               SELDTEMPLET_MFLAG = True
            Else
               SELDTEMPLET_MQR(0) = sc
               SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
               SELDTEMPLET_MRS.MoreResults
               SELDTEMPLET_MRS.Requery
     End If
  
      If sc = "F895" Or sc = "F896" Or sc = "F894" Or sc = "F846" Or sc = "F847" Or sc = "F897" Or sc = "F898" Or sc = "F947" Or sc = "F845" Then
            If SELDTEMPLET_MRS.RowCount > 0 Then
                RNAGE1 = Trim(SELDTEMPLET_MRS!LOWVALUE) & "-" & Trim(SELDTEMPLET_MRS!HIGHVALUE)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = RNAGE1
            Else
               MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = ""
            End If
      End If
     
     'THIS IS FOR MALE & FEMALE FROM DTEMPLET
       If sc = "F850" Or sc = "F851" Or sc = "F856" Or sc = "F878" Or sc = "F904" Or sc = "F907" Or sc = "F877" Or sc = "F961" Then
             If SELDTEMPLET_MFLAG1 = False Then
                SELDTEMPLET_MQR1(0) = sc
                SELDTEMPLET_MQR1(1) = Trim(txt_sex.Text)
                SELDTEMPLET_MQR1(2) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                Set SELDTEMPLET_MRS1 = SELDTEMPLET_MQR1.OpenResultset(1, 3)
                SELDTEMPLET_MFLAG1 = True
            Else
                SELDTEMPLET_MQR1(0) = sc
                SELDTEMPLET_MQR1(1) = Trim(txt_sex.Text)
                SELDTEMPLET_MQR1(2) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                SELDTEMPLET_MRS1.MoreResults
                SELDTEMPLET_MRS1.Requery
            End If
            If SELDTEMPLET_MRS1.RowCount > 0 Then
               RNAGE1 = Trim(SELDTEMPLET_MRS1!LOWVALUE) & "-" & Trim(SELDTEMPLET_MRS1!HIGHVALUE)
               MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = RNAGE1
            Else
               MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = ""
            End If
        Else
           If PC = "F869" Then
                Else
                  
                If SELDTEMPLET_MRS.RowCount > 0 Then
                    RNAGE1 = Trim(SELDTEMPLET_MRS!LOWVALUE) & "-" & Trim(SELDTEMPLET_MRS!HIGHVALUE)
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = RNAGE1
                Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = ""
                End If
           End If
      End If
  End If
End If

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
        If Trim(cmb_modality.List(ind_cbodept)) = "BIOCHEMISTRY$BC" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
End If

Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Sub FORNewInterface()
Dim sqlstr As String
Dim SELb_inter_QR As rdoQuery
Dim SELb_inter_RS As rdoResultset

 sqlstr = "SELECT * FROM interface.b_interfac WHERE REQNO=? AND (iflag='2' OR TESTCODE='F842' )"
 Set SELb_inter_QR = rdb.CreateQuery("", sqlstr)
 SELb_inter_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
Set SELb_inter_RS = SELb_inter_QR.OpenResultset(1, 3)

If SELb_inter_RS.EOF = False Then
While Not SELb_inter_RS.EOF
    
   UPDINTERFACE_QR(0) = 4
   UPDINTERFACE_QR(1) = Trim(SELb_inter_RS("machineno"))
   VAR_REQNNO = req_chglen(cmb_reqno.Text)
   UPDINTERFACE_QR(2) = Trim(VAR_REQNNO)
   UPDINTERFACE_QR(3) = Trim(SELb_inter_RS("TESTCODe"))
   UPDINTERFACE_QR.Execute
   SELb_inter_RS.MoveNext
Wend
End If
sqlstr = "SELECT * FROM interface.DINTERFACE WHERE REQNO=? AND TESTCODE NOT IN (SELECT TESTCODE FROM interface.b_interfac WHERE REQNO=?)"
 Set SELb_inter_QR = rdb.CreateQuery("", sqlstr)
 SELb_inter_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
 SELb_inter_QR(1) = Replace(Trim(cmb_reqno.Text), "/", "")
Set SELb_inter_RS = SELb_inter_QR.OpenResultset(1, 3)
If SELb_inter_RS.EOF = False Then
While Not SELb_inter_RS.EOF
    
   UPDINTERFACE_QR(0) = 2
   UPDINTERFACE_QR(1) = ""
   VAR_REQNNO = req_chglen(cmb_reqno.Text)
   UPDINTERFACE_QR(2) = Trim(VAR_REQNNO)
   UPDINTERFACE_QR(3) = Trim(SELb_inter_RS("TESTCODe"))
   UPDINTERFACE_QR.Execute
   SELb_inter_RS.MoveNext
Wend
End If

' sqlstr = "SELECT * FROM interface.DINTERFACE WHERE REQNO=? AND TESTCODE='F868' "
' Set SELb_inter_QR = rdb.CreateQuery("", sqlstr)
' SELb_inter_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
'Set SELb_inter_RS = SELb_inter_QR.OpenResultset(1, 3)
'If SELb_inter_RS.EOF = False Then
'While Not SELb_inter_RS.EOF
'
'   UPDINTERFACE_QR(0) = 2
'   UPDINTERFACE_QR(1) = Trim(SELb_inter_RS("machineno"))
'   VAR_REQNNO = req_chglen(cmb_reqno.Text)
'   UPDINTERFACE_QR(2) = Trim(VAR_REQNNO)
'   UPDINTERFACE_QR(3) = Trim(SELb_inter_RS("TESTCODe"))
'   UPDINTERFACE_QR.Execute
'   SELb_inter_RS.MoveNext
'Wend
'End If

End Sub
Public Sub cmb_reqno_Click()
On Error Resume Next
PREVTEST = ""
If ADDENDUMFALG = True Then

Dim sqlstr As String
Dim SELADDINTERFACE_QR As rdoQuery
Dim SELADDINTERFACE_RS As rdoResultset

 sqlstr = "SELECT * FROM DISLIB.DMPROCEADD WHERE REQNO=? AND OBSERVALUE IS NULL"
  Set SELADDINTERFACE_QR = rdb.CreateQuery("", sqlstr)
 SELADDINTERFACE_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
Set SELADDINTERFACE_RS = SELADDINTERFACE_QR.OpenResultset(1, 3)
If SELADDINTERFACE_RS.EOF = False Then
ADDENDUMFALG = True
viewflag = False
Else
ADDENDUMFALG = False
viewflag = True
End If

End If
Dim prpathno As String
Dim ADDENDUM As String
Dim INTFAC As Variant
Dim entr As String

'On Error GoTo MSG:
'ADDED BY SHWETA 26-04-2017 FOR NEW INTERFACING SYSTEM
FORNewInterface
'ADDED BY MILIND DT. 23/01/06 FOR RETREIVING SPECIMEN (NATURE OF MATERIALS)
Dim str, str1 As String

VAR_REQNNO = req_chglen(cmb_reqno.Text)

reqspn_qr(0) = MaskEdBox1.FormattedText
reqspn_qr(1) = VAR_REQNNO

Set reqspn_rs = reqspn_qr.OpenResultset(1, 3)
str = reqspn_rs("NMETERIALS")
If IsNull(reqspn_rs("METASTASIS")) = False Then
    str1 = reqspn_rs("METASTASIS")
Else
    str1 = ""
End If
txt_relclnhis.Text = str & "-" & str1
'UPTO here

'''''SUMOL START HERE
If reqspn_rs("NMETERIALS") = "Stone Analysis" And NEWFLAG = True Then '
    STONE ("NEW") '
    'ElseIf (reqspn_rs("NMETERIALS") = "Stone Analysis" And viewflag = True) Then ''
ElseIf (reqspn_rs("NMETERIALS") = "Stone Analysis") Then  ''
    STONE ("VIEW") '
    'viewflag = False
    'SUMOL STARTS HERE

    If SELSTONEFLAG = False Then
        SELSTONEQR(0) = VAR_REQNNO
        SELSTONEQR(1) = "F947"
        Set SELSTONERS = SELSTONEQR.OpenResultset(1, 3)
        SELSTONEFLAG = True
    Else
        SELSTONEQR(0) = VAR_REQNNO
        SELSTONEQR(1) = "F947"
        SELSTONERS.MoreResults
        SELSTONERS.Requery
        'STONETXT.Text = SELSTONERS("TXT")
    End If
    
    If SELSTONERS.RowCount = 0 Then
        Exit Sub
    End If
    STONETXT.Text = SELSTONERS("TXT")
            
Else
    STONE ("CANCEL") '
    'GoTo Label1:
End If '


'''''''''''''''''''''''''LAST HERE
'UPTO HERE
'added by milind on 14-09-2006 for GTT test
gtt_qr(0) = VAR_REQNNO
Set gtt_rs = gtt_qr.OpenResultset(1, 1)

If gtt_rs!PRO_CODE = "F844" Or gtt_rs!PRO_CODE = "F810a" Then
    gttflag = True
    gtt_rs.Close
    MSFlexGrid1.Visible = False
    txt_observed.Visible = False
    MSFlexGrid2.Visible = True
    cmb_Gtt_UrineValue.Visible = True
    MSFlexGrid2.clear
   ' MSFlexGrid2.FormatString = "Time Slots     |^    Blood Sugar       |^Urine Sugar           "
    MSFlexGrid2.FormatString = "Time Slots     |^    Blood Sugar       "
    'MSFlexGrid2.MergeCells = flexMergeRestrictRows
    'MSFlexGrid2.MergeRow(0) = True
'    cmb_finalby.Text = cmb_enterby.Text
    
    If Opt_patient.Value = True Then
        If NEWFLAG = True Then
            'NEWFLAG = False
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
                    lbl_rwd.Visible = False
                ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                    lblicu.Visible = False
                    lbl_rwd.Visible = True
                End If
              
            End If
                
            '******ICU****************
            
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
                
            '               If Trim(txt_age.Text) = 0 Then
            '                   Dim N As Integer
            '                   N = txt_age.Text / 12
            '                   txt_age.Text = N
            '               End If
               
            If Trim(txt_age.Text) < 12 Then
                Dim age1 As String
                age1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                age1 = "." & (age Mod 12)
                If age1 = ".0" Then
                    age1 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                    age1 = ".0" & (age1 Mod 30)
                End If
                txt_age.Text = Trim(age1)
            End If
            
            txt_pridoc.Text = SELTA101_RS("ud_name")
            txt_category.Text = SELTA101_RS("pat_catg")
            
            If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
                
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
            End If
            
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
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
                
                'STAT SAMPLE CHANGES
                If Trim(selpathdrequisitn_rs("SPIHREQNO")) = "U" Then
                Label9.Visible = True
                Else
                Label9.Visible = False
                End If
            End If
            
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If

            If selpathdrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
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
                ' txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If

            ICASENO = MaskEdBox1.FormattedText
            IREQNO = VAR_REQNNO
            Dim TEST_DET, MACHINENO
            MSFlexGrid2.Rows = 1
            I = 1
                
            'FETCH TEST_CODE FROM DMULPROCE
            If seldmulprocflag = False Then
                SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
                seldmulprocflag = True
            Else
                SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                seldmulprocrs.MoreResults
                seldmulprocrs.Requery
            End If
                
            If seldmulprocrs.RowCount > 0 Then
                Do Until seldmulprocrs.EOF
                    PC = seldmulprocrs!PRO_CODE
                    If SELMA101FLAG = False Then
                        selMA101QR(0) = PC
                        Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        selMA101QR(0) = PC
                        SELMA101RS.MoreResults
                        SELMA101RS.Requery
                    End If
                    
                    TEST_DET = Trim(SELMA101RS!SER_DET)
                                
                    'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE
    
                    If SELDTEMPLETFLAG = False Then
                        SELDTEMPLETQR(0) = PC
                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                        SELDTEMPLETFLAG = True
                    Else
                        SELDTEMPLETQR(0) = PC
                        SELDTEMPLETRS.MoreResults
                        SELDTEMPLETRS.Requery
                    End If
                    'RANGE = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                                
                    MSFlexGrid2.AddItem I & " . FASTING" '& vbTab & vbTab & vbTab & RANGE
                    I = I + 1
                   ' MSFlexGrid2.AddItem i & " . 30 min.   "
                   ' i = i + 1
                  '  MSFlexGrid2.AddItem i & " . 60 min.   "
                  '  i = i + 1
                  '  MSFlexGrid2.AddItem i & " . 90 min.   "
                 '   i = i + 1
                  '  MSFlexGrid2.AddItem i & " . 120 min. "
                  '  i = i + 1
                    MSFlexGrid2.AddItem I & " . 2 hrs. "
                    
                    seldmulprocrs.MoveNext
                Loop
            End If
        End If
         
        If viewflag = True Then
            viewflag = False
    
            ''******ICU****************
            VAR_REQNNO = req_chglen(Trim(cmb_reqno.Text))
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
                    lbl_rwd.Visible = False
                ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                    lblicu.Visible = False
                    lbl_rwd.Visible = True
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
                    ADDENDUMFALG = True
                    cmd_adenum.Visible = True
                End If
                cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
                cmb_finalby.Text = "" & REFDR_RS("FREPORTBY")
                
                '''' By mustufa on 21 05 2026
'                Dim RS_REQSTAT As rdoResultset
'
'                Set RS_REQSTAT = rdb.OpenResultset( _
'                    "SELECT REQ_STAT FROM RISLIB.RREQUISITN WHERE REQ_NO='" & VAR_REQNNO & "'", _
'                    1, 1)
'                If Not RS_REQSTAT.EOF Then
'                    If Trim(RS_REQSTAT("REQ_STAT")) = "T" Then
'                        cmd_commit.Enabled = True
'                    Else
'                        cmd_commit.Enabled = False
'                    End If
'
'                Else
'                    cmd_commit.Enabled = False
'                End If
'
'                RS_REQSTAT.Close
'                Set RS_REQSTAT = Nothing
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
                
                'STAT SAMPLE CHANGES
                If Trim(selpathdrequisitn_rs("SPIHREQNO")) = "U" Then
                Label9.Visible = True
                Else
                Label9.Visible = False
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
                txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF = False Then
                    txt_relclnhis.Visible = True
                    txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
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
        
        
            ''''''FILLING THE GRID
            MSFlexGrid2.Rows = 1
            I = 1
            MSFlexGrid2.AddItem I & " . FASTING" '& vbTab & vbTab & vbTab & RANGE
            I = I + 1
           ' MSFlexGrid2.AddItem i & " . 30 min.   "
           ' i = i + 1
           ' MSFlexGrid2.AddItem i & " . 60 min.   "
           ' i = i + 1
           ' MSFlexGrid2.AddItem i & " . 90 min.   "
           ' i = i + 1
          '  MSFlexGrid2.AddItem i & " . 120 min. "
           ' i = i + 1
            MSFlexGrid2.AddItem I & " . 2 hrs. "

            '*********** 23/04/2013 Aditi**************************************************
            '*********** view GTT details in grid******************************************

                For I = 1 To MSFlexGrid2.Rows - 1
                      SELGTTDTLSQR(0) = VAR_REQNNO
                      If MSFlexGrid2.TextMatrix(I, 0) = "1 . FASTING" Then
                            SELGTTDTLSQR(1) = 0
                      Else
                            SELGTTDTLSQR(1) = Mid(MSFlexGrid2.TextMatrix(I, 0), 5, 3)
                      End If
                      Set SELGTTDTLSRS = SELGTTDTLSQR.OpenResultset(1, 1)
                      
                      j = 2
                      If j <= 11 Then
                         If SELGTTDTLSRS.RowCount > 0 Then
                            If IsNull(SELGTTDTLSRS(2)) Then
                               MSFlexGrid2.TextMatrix(I, 1) = ""
                            Else
                               MSFlexGrid2.TextMatrix(I, 1) = CStr(SELGTTDTLSRS(2))
                            End If
                            MSFlexGrid2.TextMatrix(I, 2) = Trim(SELGTTDTLSRS(3))
                         End If
                    End If
                Next I
                
        If SELDMBCCOMMFLAG = False Then
            SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
            Set SELDMBCCOMMRS = SELDMBCCOMMQR.OpenResultset(1, 3)
            SELDMBCCOMMFLAG = True
        Else
            SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
            SELDMBCCOMMRS.MoreResults
            SELDMBCCOMMRS.Requery
        End If
        If SELDMBCCOMMRS.RowCount > 0 Then
            If SELDMBCCOMMRS.EOF = False Then
                TXT_COMMENTS.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
            End If
        End If

            '*****************************************************************************
        End If 'View flag end
    End If 'If Patient end
    
    '************** 25/04/2013 Aditi ********************************************
    '**************GTT report for staff******************************************
    If opt_staff.Value = True Then
           
           
           If NEWFLAG = True Then
              STAFF
              NEWFLAG = False
              MSFlexGrid2.Rows = 1
              I = 1
               MSFlexGrid2.AddItem I & " . FASTING" '& vbTab & vbTab & vbTab & RANGE
                    I = I + 1
                    MSFlexGrid2.AddItem I & " . 30 min.   "
                    I = I + 1
                   ' MSFlexGrid2.AddItem i & " . 60 min.   "
                   ' i = i + 1
                   ' MSFlexGrid2.AddItem i & " . 90 min.   "
                   ' i = i + 1
                  '  MSFlexGrid2.AddItem i & " . 120 min. "
                  '  i = i + 1
                    MSFlexGrid2.AddItem I & " . 2 hrs. "
           End If

           If viewflag = True Then
              STAFF
              viewflag = False

            MSFlexGrid2.Rows = 1
            I = 1
            MSFlexGrid2.AddItem I & " . FASTING" '& vbTab & vbTab & vbTab & RANGE
            I = I + 1
           ' MSFlexGrid2.AddItem i & " . 30 min.   "
           ' i = i + 1
           ' MSFlexGrid2.AddItem i & " . 60 min.   "
           ' i = i + 1
           ' MSFlexGrid2.AddItem i & " . 90 min.   "
           ' i = i + 1
           ' MSFlexGrid2.AddItem i & " . 120 min. "
           ' i = i + 1
            MSFlexGrid2.AddItem I & " . 2 hrs. "

                For I = 1 To MSFlexGrid2.Rows - 1
                      SELGTTDTLSQR(0) = VAR_REQNNO
                      If MSFlexGrid2.TextMatrix(I, 0) = "1 . FASTING" Then
                            SELGTTDTLSQR(1) = 0
                      Else
                            SELGTTDTLSQR(1) = Mid(MSFlexGrid2.TextMatrix(I, 0), 5, 3)
                      End If
                      Set SELGTTDTLSRS = SELGTTDTLSQR.OpenResultset(1, 1)
                      
                      j = 2
                      If j <= 11 Then
                         If SELGTTDTLSRS.RowCount > 0 Then
                            If IsNull(SELGTTDTLSRS(2)) Then
                               MSFlexGrid2.TextMatrix(I, 1) = ""
                            Else
                               MSFlexGrid2.TextMatrix(I, 1) = CStr(SELGTTDTLSRS(2))
                            End If
                            MSFlexGrid2.TextMatrix(I, 2) = Trim(SELGTTDTLSRS(3))
                         End If
                    End If
                Next I
              
        If SELDMBCCOMMFLAG = False Then
            SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
            Set SELDMBCCOMMRS = SELDMBCCOMMQR.OpenResultset(1, 3)
            SELDMBCCOMMFLAG = True
        Else
            SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
            SELDMBCCOMMRS.MoreResults
            SELDMBCCOMMRS.Requery
        End If
        If SELDMBCCOMMRS.RowCount > 0 Then
            If SELDMBCCOMMRS.EOF = False Then
                TXT_COMMENTS.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
            End If
        End If
           End If
    End If 'If Staff end
    '****************************************************************************
    
    
    
    Exit Sub

Else    'GTT ends here
    
    MSFlexGrid2.Visible = False
    
    If (reqspn_rs("NMETERIALS") = "Stone Analysis") Then
        MSFlexGrid1.Visible = False
    Else
        MSFlexGrid1.Visible = True
        'txt_observed.Visible = True
        MSFlexGrid1.clear
        MSFlexGrid1.FormatString = "S.No.|^Procedure Name                   |^Observed Value|^Normal Range of the Procedure |^ Panic Value|^Source|^Preval|^MachineNo."
    End If
    'MSFlexGrid1.FormatString = "SNo.|^    Procedure Name             |^  Observed Value  |^  Normal Range of the Procedure |^ Panic Value|^  Source|^Preval      |^Machine"
    'COMMENTED BY SHWETA 29-06-2017
   ' cmb_finalby.Text = cmb_enterby.Text
    If Opt_patient.Value = True Then
        If NEWFLAG = True Then
        'NEWFLAG = False
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
                lbl_rwd.Visible = False
            ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                lblicu.Visible = False
                lbl_rwd.Visible = True
            End If
        End If
        '******ICU****************
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
        txt_age = DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                
        '               If Trim(txt_age.Text) = 0 Then
        '                   Dim N As Integer
        '                   N = txt_age.Text / 12
        '                   txt_age.Text = N
        '               End If
                   
        If Trim(txt_age.Text) < 12 Then
            Dim age2 As String
            age2 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            age = "." & (age Mod 12)
            If age2 = ".0" Then
                age2 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                age2 = ".0" & (age2 Mod 30)
            End If
            txt_age.Text = Trim(age2)
        End If
                
        txt_pridoc.Text = SELTA101_RS("ud_name")
        txt_category.Text = SELTA101_RS("pat_catg")
                
        If REFDRFLAG1 = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            REFDRFLAG1 = True
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
        If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
        End If
        If SELPATHDREQUISITNFALG1 = False Then
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFALG1 = True
        Else
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
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
            If Trim(selpathdrequisitn_rs("spihreqno")) = "U" Then
            Label9.Visible = True
            Else
            Label9.Visible = False
            End If
        End If
                
        If SELPATHDREQUISITNFALG1 = False Then
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFALG1 = True
        Else
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
                
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
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
            ' txt_procedure.Text = selta1012_rs("prvdiagdes")
        End If
        
    
     ''/////////////////////////////////////////////////////////////////////////////////////////
     '' CODE FOR DISPLAY DATA IN FLEX GRID
     ''*****************************************************************************************
     ' WRITTEN BY AMIT FOR INTERFACING
       
        ICASENO = MaskEdBox1.FormattedText
        IREQNO = VAR_REQNNO
        Dim TEST_DET1, RANGE1, MACHINENO1
        MSFlexGrid1.Rows = 1
        I = 1
        'FETCH TEST_CODE FROM DMULPROCE
        If seldmulprocflag = False Then
            SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
            seldmulprocflag = True
        Else
            SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            seldmulprocrs.MoreResults
            seldmulprocrs.Requery
        End If
        
        '----- Commented by Shraddha on 29/01/10 -----
        '----- As told by Mrs.Kalpana m'am -----
'        If seldmulprocrs.RowCount > 0 Then
'            Do Until seldmulprocrs.EOF
'               PC = seldmulprocrs!PRO_CODE
'                ''sumol nov 2006 starts here
'                If PC = "F947" And seldmulprocrs.RowCount < 2 Then
'                    ser_code.Text = "F947"
'                Else
'                    'MsgBox "Requisitn Not Valid & Raised requisitn should be for stone analysis only"
'                End If
'                ''ends here
'                'Next check for F898 is added by Shraddha Pandit on 19/12/08
'                'If PC = "F895" Or PC = "F896" Or PC = "F894" Or PC = "F846" Or PC = "F847" Or PC = "F845" Or PC = "F897" Or PC = "F898" Then
'                'Above statement commented and next statement added by Shraddha on 02/06/09 '(Removed F845 and F894)
'                If PC = "F895" Or PC = "F896" Or PC = "F846" Or PC = "F847" Or PC = "F897" Or PC = "F898" Then
'                    If SELMA101FLAG = False Then
'                        selMA101QR(0) = PC
'                        Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
'                        SELMA101FLAG = True
'                    Else
'                        selMA101QR(0) = PC
'                        SELMA101RS.MoreResults
'                        SELMA101RS.Requery
'                    End If
'                    TEST_DET = Trim(SELMA101RS!SER_DET)
'                    'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE
'
'                    If SELDTEMPLETFLAG = False Then
'                       SELDTEMPLETQR(0) = PC
'
'                       Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                       SELDTEMPLETFLAG = True
'                    Else
'                       SELDTEMPLETQR(0) = PC
'                       SELDTEMPLETRS.MoreResults
'                       SELDTEMPLETRS.Requery
'                    End If
'                    RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
'                    ''****** COMMENTED BY SHWETA 20-JULY-09
'                    ''******IN GRID SHOWS 2 ROWS FOR SAME RESULT
'                    MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & vbTab & RNAGE1
'                    I = I + 1
'                    ''***** END COMMENTED BY SHWETA 20-JULY-09
'                End If
'                seldmulprocrs.MoveNext
'            Loop
'        End If
'----- Comment till here, Shraddha 29/01/10 ------
        ' FETCH TEST_CODE FROM INTERFACE
        
        'MsgBox "MERGEDATA "
        
        If SELINTER_FLAG = False Then
            SELINTER_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            Set SELINTER_RS = SELINTER_QR.OpenResultset(1, 3)
            SELINTER_FLAG = True
        Else
            SELINTER_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            SELINTER_RS.MoreResults
            SELINTER_RS.Requery
        End If
        'MsgBox "MERGEDATA DONE"
'         If SELINTER_RS.RowCount > 0 Then
'
'
'         End If
        
        If SELINTER_RS.RowCount > 0 Then
        
        
        
        
            Do Until SELINTER_RS.EOF
                PC = SELINTER_RS!testcode
                'THIS IS FOR GETTING TEST DESCRIPTION FOR TEST CODE
                If SELMA101FLAG = False Then
                    selMA101QR(0) = PC
                    Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    selMA101QR(0) = PC
                    SELMA101RS.MoreResults
                    SELMA101RS.Requery
                End If
                TEST_DET = Trim(SELMA101RS!SER_DET)
                'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE
                If SELDTEMPLETFLAG = False Then
                   SELDTEMPLETQR(0) = PC
                   Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                   SELDTEMPLETFLAG = True
                Else
                   SELDTEMPLETQR(0) = PC
                   SELDTEMPLETRS.MoreResults
                   SELDTEMPLETRS.Requery
                End If
                'THIS IF FOR F869 (A & C) IN DTEMPLET
                If PC = "F869" Then
                    ALKAGE
                    If SELDBCALKPHSFLAG = False Then
                        SELDBCALKPHSQR(0) = PC
                        SELDBCALKPHSQR(1) = RNAGE
                        SELDBCALKPHSQR(2) = SS
                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
                        SELDBCALKPHSFLAG = True
                    Else
                        SELDBCALKPHSQR(0) = PC
                        SELDBCALKPHSQR(1) = RNAGE
                        SELDBCALKPHSQR(2) = SS
                        SELDBCALKPHSRS.MoreResults
                        SELDBCALKPHSRS.Requery
                    End If
                    On Error Resume Next
                    RNAGE1 = SELDBCALKPHSRS!range
                    'justin 18-08-2004
                    ALKLV = Left(Trim(SELDBCALKPHSRS("RANGE")), 2)
                    ALKHV = Right(Trim(SELDBCALKPHSRS("RANGE")), 3)
                End If
                  
                'SHWETA 10-01-2013
''''23-02-2013
'NEW REFERANCE RANGE DATA
                Dim RS_VWREFRNG As rdoResultset
                 Set RS_VWREFRNG = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & PC & "' ", 1, 1)
        
                               If Not RS_VWREFRNG.EOF Then
                Dim agecal As Double
                agecal = DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                txt_age.Text = DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                If Trim(txt_age.Text) <= 1 Then
                Dim age3 As Double
                 age3 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                age3 = "." & age3 '(age Mod 12)
                If age3 = ".0" Then
                    age3 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                    age3 = ".0" & (age3 Mod 30)
                End If
                txt_age.Text = Trim(age3)
            End If
                   ' ALKAGEO 'FOR F876
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & PC & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range
                    
                    If PC = "F739" And UCase(Trim(seldmulprocrs("COLLECTION"))) = "YES" Then
                        While Not SELDBCALKPHSGGTRS.EOF
                        If Trim(SELDBCALKPHSGGTRS("AGE1")) = "P" Then
                        RNAGE1 = SELDBCALKPHSGGTRS!range
                        
                       
                        
                        End If
                        
                        SELDBCALKPHSGGTRS.MoveNext
                        Wend
                    Else
                        
                        
                    End If
                    
                    
                    If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Or InStr(Trim(SELDBCALKPHSGGTRS!range), "-") Then
                    Else
                    ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                    ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                    End If
                    End If
                End If
                
                
                ''''''''
                  
                If PC = "F999" Then
                    GTTAGE 'FOR GGT
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE='" & PC & "' AND (AGE1<" & RNAGE & " AND AGE2>=" & RNAGE & ") AND SEX='" & SS & "'", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range
                    
                    If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Then
                    Else
                    ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 2)
                    ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                    End If
                    End If
                End If
                'justin 18-08-2004
                'THIS IS FOR MALE & FEMALE FROM DTEMPLET
                If PC = "F850" Or PC = "F851" Or PC = "F856" Or PC = "F878" Or PC = "F904" Or PC = "F907" Or PC = "F877" Or PC = "F961" Then
                    If SELDTEMPLETFLAG1 = False Then
                        SELDTEMPLETQR1(0) = PC
                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
                        Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
                        SELDTEMPLETFLAG1 = True
                    Else
                        SELDTEMPLETQR1(0) = PC
                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
                        SELDTEMPLETRS1.MoreResults
                        SELDTEMPLETRS1.Requery
                    End If
                    RNAGE1 = Trim(SELDTEMPLETRS1!LOWVALUE) & "-" & Trim(SELDTEMPLETRS1!HIGHVALUE)
                Else
                    If PC = "F829" Then 'Shraddha 17/12/08
                        If Trim(SELDTEMPLETRS!LOWVALUE) = "" Then
                            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & Trim(SELDTEMPLETRS!HIGHVALUE)
                        Else
                            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                        End If
                    Else
                    'SHWETA BIO REG RANGE CHANGES 09-MAR-2013
                    Dim RS_VWREFRNG1 As rdoResultset
                 Set RS_VWREFRNG1 = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & PC & "' ", 1, 1)
        
                
              '  09-MAR-2013
                        If PC = "F869" Or PC = "F999" Or PC = "F876" Or PC = "F893" Or PC = "F894" Or RS_VWREFRNG1.EOF = False Then
                         
                        Else
                            SEX = Trim(txt_sex.Text)
                            If Not SELDTEMPLETRS.EOF Then
                            
                                If Trim(SELDTEMPLETRS("SEX")) = "" Then
                                    RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                                Else
                           If Trim(PC) = "F126" Then
                       If Trim(txt_age.Text) <= 15 Then
                          SEX = "C"
                       Else
                          SEX = "A"
                       End If
                       
                    End If
                    
                    
                    
Back:

                            If Trim(SELDTEMPLETRS("SEX")) = "" And PC = "F124" Then
                                    RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                            End If
If Trim(SEX) = Trim(SELDTEMPLETRS("SEX")) Then
                                    If Trim(SELDTEMPLETRS!LOWVALUE) = "" Then
                                        RNAGE1 = Trim(SELDTEMPLETRS!HIGHVALUE)
                                    Else
                                        RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                                    End If
                                    
                                Else
                                    SELDTEMPLETRS.MoveNext
                                    If Not SELDTEMPLETRS.EOF Then GoTo Back
                                End If
                                End If
                            End If
                        End If
                    End If
                End If
                'THIS IS FOR PICKUP TESTRESULT FROM INTERFACE/DINTERFACE
    
                INTERFACEREAD_QR(0) = IREQNO
                INTERFACEREAD_QR(1) = PC
                Set INTERFACEREAD_RS = INTERFACEREAD_QR.OpenResultset(1, 1)
    
                If INTERFACEREAD_RS.RowCount > 0 Then
                    INTFAC = INTERFACEREAD_RS("testresult")
                    If Trim(INTFAC) = "" Then
                        INTMODE = ""
                        MACHINENO = ""
                    Else
                        INTMODE = "I"
                        MACHINENO = INTERFACEREAD_RS("MACHINENO")
                    End If
                Else
                    INTFAC = ""
                    INTMODE = ""
                End If
                If INTERFACEREAD_RS.RowCount = 0 Then
                    '''FROM HERE
                   ' SELMSADINTF_QR(0) = IREQNO
                    'SELMSADINTF_QR(1) = PC
                    'Set SELMSADINTF_RS = SELMSADINTF_QR.OpenResultset(1, 1)
                    
'                    If SELMSADINTF_RS.RowCount > 0 Then
'                        INTFAC = SELMSADINTF_RS("testresult")
'                        If Trim(INTFAC) = "" Then
'                            INTMODE = ""
'                            MACHINENO = ""
'                        Else
'                            INTMODE = "I"
'                            MACHINENO = SELMSADINTF_RS("MACHINENO")
'                        End If
'                    Else
'                        INTFAC = ""
'                        INTMODE = ""
'                    End If
'                    'UPTO HERE
                    If PREVTEST = TEST_DET Then
                    Else
                                     
                    MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & vbTab & MACHINENO '& vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
                    End If
                    PREVTEST = TEST_DET
                    
                Else
                   'MSFlexGrid1.AddItem i & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL"))) & vbTab & MACHINENO
                     If PREVTEST = TEST_DET Then
                    Else
                                     
                    MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL"))) & vbTab & MACHINENO
                    End If
                    PREVTEST = TEST_DET
                    'THIS IS FOR DEFINE PANIC VALUE BASED ON OBSER VALUE
                    If Len(Trim(MSFlexGrid1.TextMatrix(I, 2))) = 0 Then
                    Else
                        'JUSTIN 18 - 8 - 2004
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                        NU = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), NU + 1)) Then
                            MSFlexGrid1.row = I
                            MSFlexGrid1.col = 2
                            MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = Val(Trim(MSFlexGrid1.TextMatrix(I, 2)))
                            Else
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                            End If
                        Else
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                MSFlexGrid1.row = I
                                MSFlexGrid1.CellBackColor = vbWhite
                                'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Val(Trim(MSFlexGrid1.TextMatrix(i, 2)))
                                If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) Then
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                                Else
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                                End If
                            End If
                            'If Len(Trim(MSFlexGrid1.TextMatrix(i, 2))) = 0 Then
                            'Else
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                MSFlexGrid1.row = I
                                MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
                                If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) Then
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                                Else
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                                End If
                            End If
                        End If
                    End If
                End If
                End If
                SELINTER_RS.MoveNext
                I = I + 1
            Loop
         
            
            
            
            
        End If
    End If
     
    ''' 'THIS IS FOR DEFINE PANIC VALUE BASED ON OBSER VALUE
    '''
    ''' If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
    '''    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
    '''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
    '''    If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC"))) Then
    '''        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
    '''    Else
    '''        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
    '''    End If
    '''    For i = 1 To 1000
    '''    Beep
    '''    Next i
    '''    Else
    '''    If Val(Trim(txt_observed.Text)) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
    '''        MSFlexGrid1.CellBackColor = vbWhite
    '''        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
    '''        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
    '''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
    '''        Else
    '''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
    '''        End If
    '''    Else
    '''        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
    '''    End If
    '''    If Val(Trim(txt_observed.Text)) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
    '''        MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
    '''        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
    '''        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
    '''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
    '''        Else
    '''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
    '''        End If
    '''
    '''    End If
     'ENDED INTERFACES BY AMITI
     '/////////////////////////////////////////////////////////////////////////////////////////
     'THIS IS FOR INTEREXPCD TABLE [F895,F896,F894,F846,F847,F897,F898,F947,F845]
    '            If seldmulprocflag = False Then
    '                SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
    '                Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
    '                seldmulprocflag = True
    '           Else
    '                SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
    '                seldmulprocrs.MoreResults
    '                seldmulprocrs.Requery
    '           End If
    '           If seldmulprocrs.RowCount > 0 Then
    '           i = 1
    '           Do Until seldmulprocrs.EOF
    '           PC = seldmulprocrs!PRO_CODE
    '
    '           If PC = "F895" Or PC = "F896" Or PC = "F894" Or PC = "F846" Or PC = "F847" Or PC = "F897" Or PC = "F898" Or PC = "F947" Or PC = "F845" Then
    '                'THIS IS FOR GETTING TEST DESCRIPTION FOR TEST CODE
    '
    '            If SELMA101FLAG = False Then
    '                selMA101QR(0) = PC
    '                Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
    '                SELMA101FLAG = True
    '            Else
    '                selMA101QR(0) = PC
    '                SELMA101RS.MoreResults
    '                SELMA101RS.Requery
    '            End If
    '            TEST_DET = Trim(SELMA101RS!SER_DET)
    '            'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE
    '
    '            If SELDTEMPLETFLAG = False Then
    '               SELDTEMPLETQR(0) = PC
    '               Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
    '               SELDTEMPLETFLAG = True
    '            Else
    '               SELDTEMPLETQR(0) = PC
    '               SELDTEMPLETRS.MoreResults
    '               SELDTEMPLETRS.Requery
    '            End If
    '            RANGE = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
    '            MSFlexGrid1.AddItem i & vbTab & TEST_DET & vbTab & vbTab & RANGE
    '            i = i + 1
    '            seldmulprocrs.MoveNext
    '
    '           End If
    '           Loop
    '           Else
    '
    '           End If
                '/////////////////////////////////////////////////////////////////////////////////////////
    '''        If seldmulprocrs.RowCount > 0 Then
    '''
    '''                MSFlexGrid1.Rows = 1
    '''                i = 1
    '''
    '''                While Not seldmulprocrs.EOF
    '''                   PC = seldmulprocrs("PRO_CODE")
    '''
    '''                   ''THIS CODE FOR GROUP PRO_CODE
    '''                   ''////////////////////////
    '''
    '''
    '''         ' INTERFACE
    '''                   INTERFACEREAD_QR(0) = req_chglen(cmb_reqno.Text)
    '''                   INTERFACEREAD_QR(1) = SELDBCPROSUBRS("PROCODESUB")
    '''                   Set INTERFACEREAD_RS = INTERFACEREAD_QR.OpenResultset(1, 1)
    '''
    '''                    If INTERFACEREAD_RS.RowCount > 0 Then
    '''                        INTFAC = INTERFACEREAD_RS("testresult")
    '''                        intmode = "I"
    '''                    Else
    '''                        INTFAC = ""
    '''                        intmode = ""
    '''                    End If
    ''''// ///////////////////CODE FOR PROCODE F869
    '''
    '''                If SELDBCPROSUBRS("PROCODESUB") = "F869" Then
    '''                    ALKAGE
    '''
    '''                    If SELDBCALKPHSFLAG = False Then
    '''                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
    '''                        SELDBCALKPHSQR(1) = RNAGE
    '''                        SELDBCALKPHSQR(2) = SS
    '''                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
    '''                        SELDBCALKPHSFLAG = True
    '''                    Else
    '''                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
    '''                        SELDBCALKPHSQR(1) = RNAGE
    '''                        SELDBCALKPHSQR(2) = SS
    '''                        'Set SELBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
    '''                        SELDBCALKPHSRS.MoreResults
    '''                        SELDBCALKPHSRS.Requery
    '''                    End If
    '''
    '''                    If SELDTEMPLETRS.RowCount > 0 Then
    '''                        If SELMA101RS.RowCount > 0 Then
    '''                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
    '''                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
    '''                            Else
    '''                                SELDTEMPLETRS.MoveNext
    '''                                  MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
    '''                            End If
    '''                            ALKLV = Left(Trim(SELDBCALKPHSRS("RANGE")), 2)
    '''                            ALKHV = Right(Trim(SELDBCALKPHSRS("RANGE")), 3)
    '''                            i = i + 1
    '''                        End If
    '''                    End If
    '''
    '''    '''ADD BY MONO ON 17-09-03 FOR M &F
    '''    ElseIf SELDBCPROSUBRS("PROCODESUB") = "F850" Or SELDBCPROSUBRS("PROCODESUB") = "F851" Or SELDBCPROSUBRS("PROCODESUB") = "F856" Or SELDBCPROSUBRS("PROCODESUB") = "F878" Or SELDBCPROSUBRS("PROCODESUB") = "F904" Or SELDBCPROSUBRS("PROCODESUB") = "F907" Then
    '''        If SELDTEMPLETFLAG1 = False Then
    '''            SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
    '''            SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
    '''            Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
    '''            SELDTEMPLETFLAG1 = True
    '''        Else
    '''            SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
    '''            SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
    '''            SELDTEMPLETRS1.MoreResults
    '''            SELDTEMPLETRS1.Requery
    '''        End If
    '''
    '''        If SELDTEMPLETRS1.RowCount > 0 Then
    '''            If SELMA101RS.RowCount > 0 Then
    '''                If SELDTEMPLETRS1("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS1("SEX")) = "" Then
    '''                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
    '''                Else
    '''                    SELDTEMPLETRS1.MoveNext
    '''                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))                                                 ' comments removed for interfaceing 0n 30-12-2003
    '''                End If
    '''                i = i + 1
    '''            End If
    '''        End If
    '''    ''AMIT CODDEE
    '''
    '''    ElseIf SELDBCPROSUBRS("PROCODESUB") = "F925" Or SELDBCPROSUBRS("PROCODESUB") = "F926" Or SELDBCPROSUBRS("PROCODESUB") = "F927" Or SELDBCPROSUBRS("PROCODESUB") = "F928" Or SELDBCPROSUBRS("PROCODESUB") = "F929" Or SELDBCPROSUBRS("PROCODESUB") = "F930" Then
    '''
    '''          If SELDTEMPLETFLAG1 = False Then
    '''              SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
    '''              Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
    '''              SELDTEMPLETFLAG1 = True
    '''          Else
    '''              SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
    '''              SELDTEMPLETRS1.MoreResults
    '''              SELDTEMPLETRS1.Requery
    '''          End If
    '''
    '''          If SELDTEMPLETRS1.RowCount > 0 Then
    '''              If SELMA101RS.RowCount > 0 Then
    '''                  If SELDTEMPLETRS1("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS1("SEX")) = "" Then
    '''                      MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & intmode & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
    '''                  Else
    '''                      SELDTEMPLETRS1.MoveNext
    '''                      MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & intmode & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
    '''
    '''                  End If
    '''                  i = i + 1
    '''              End If
    '''          End If
    '''    '''ADD BY MONO ON 17-09-03 FOR M &F
    '''    Else
    '''        If SELDTEMPLETRS.RowCount > 0 Then
    '''            If SELMA101RS.RowCount > 0 Then
    '''
    '''
    '''                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
    '''                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & intmode '& vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
    '''
    '''                Else
    '''                       MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & intmode & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
    '''
    '''                End If
    '''                i = i + 1
    '''            End If
    '''        End If
    '''    End If
    '''
    '''        SELDBCPROSUBRS.MoveNext
    '''
    '''Wend
    '''End If
    '''Else
    '''
    '''        ''' CODE FOR INDIVIDUAL PRO_CODE BY AMIT
    '''        ''//////////////////////////////////////////////////////////////////////////////
    '''
    '''                    If SELMA101FLAG = False Then
    '''                        selMA101QR(0) = seldmulprocrs("PRO_CODE")
    '''                        Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
    '''                        SELMA101FLAG = True
    '''                    Else
    '''                        selMA101QR(0) = seldmulprocrs("PRO_CODE")
    '''                        SELMA101RS.MoreResults
    '''                        SELMA101RS.Requery
    '''                    End If
    '''                    If SELDTEMPLETFLAG = False Then
    '''                        SELDTEMPLETQR(0) = seldmulprocrs("PRO_CODE")
    '''                        'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
    '''                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
    '''                        SELDTEMPLETFLAG = True
    '''                    Else
    '''                        SELDTEMPLETQR(0) = seldmulprocrs("PRO_CODE")
    '''                        'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
    '''                        SELDTEMPLETRS.MoreResults
    '''                        SELDTEMPLETRS.Requery
    '''                    End If
    '''                        '*******
    '''                        If seldmulprocrs("PRO_CODE") = "F869" Then
    '''                            ALKAGE
    '''
    '''                            If SELDBCALKPHSFLAG = False Then
    '''                                SELDBCALKPHSQR(0) = seldmulprocrs("PRO_CODE")
    '''                                SELDBCALKPHSQR(1) = RNAGE
    '''                                SELDBCALKPHSQR(2) = SS
    '''                                Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
    '''                                SELDBCALKPHSFLAG = True
    '''                            Else
    '''                                SELDBCALKPHSQR(0) = seldmulprocrs("PRO_CODE")
    '''                                SELDBCALKPHSQR(1) = RNAGE
    '''                                SELDBCALKPHSQR(2) = SS
    '''                                SELDBCALKPHSRS.MoreResults
    '''                                SELDBCALKPHSRS.Requery
    '''                            End If
    '''
    '''                            If SELDBCALKPHSRS.RowCount > 0 Then
    '''                                If SELMA101RS.RowCount > 0 Then
    '''                                    If SELDBCALKPHSRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDBCALKPHSRS("SEX")) = "" Then
    '''                                         MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDBCALKPHSRS("RANGE")) & vbTab & vbTab & intmode
    '''                                    Else
    '''                                        SELDBCALKPHSRS.MoveNext
    '''                                        MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDBCALKPHSRS("RANGE")) & vbTab & vbTab & intmode
    '''                                    End If
    '''                                    i = i + 1
    '''                                End If
    '''                            End If
    '''                        '''ADD BY MONO ON 17-09-03 FOR M &F
    '''                        '' add code by amit
    '''
    '''
    '''                        ElseIf PC = "F850" Or PC = "F851" Or PC = "F856" Or PC = "F878" Or PC = "F904" Or PC = "F907" Or PC = "F925" Or PC = "F926" Or PC = "F927" Or PC = "F928" Or PC = "F929" Or PC = "F930" Or PC = "F931" Or PC = "F932" Or PC = "F933" Or PC = "F975" Or PC = "F976" Or PC = "F918" Or PC = "F919" Or PC = "F920" Or PC = "F924" Or PC = "F915" Or PC = "F916" Or PC = "F917" Or PC = "F921" Or PC = "F922" Or PC = "F923" Then
    '''
    '''
    '''                            If SELDTEMPLETFLAG1 = False Then
    '''                                SELDTEMPLETQR1(0) = seldmulprocrs("PRO_CODE")
    '''                                'SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
    '''                                Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
    '''                                SELDTEMPLETFLAG1 = True
    '''                            Else
    '''                                SELDTEMPLETQR1(0) = seldmulprocrs("PRO_CODE")
    '''                                'SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
    '''                                SELDTEMPLETRS1.MoreResults
    '''                                SELDTEMPLETRS1.Requery
    '''                            End If
    '''
    '''                    If SELDTEMPLETRS1.RowCount > 0 Then
    '''                        If SELMA101RS.RowCount > 0 Then
    '''                            If SELDTEMPLETRS1("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS1("SEX")) = "" Then
    '''                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & intmode
    '''                        Else
    '''                            SELDTEMPLETRS1.MoveNext
    '''                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & intmode
    '''                        End If
    '''                        i = i + 1
    '''                    End If
    '''                    End If
    '''                Else
    '''                ' INTERFACE
    '''                   INTERFACEREAD_QR(0) = req_chglen(cmb_reqno.Text)
    '''                   INTERFACEREAD_QR(1) = seldmulprocrs("PRO_CODE")
    '''                   Set INTERFACEREAD_RS = INTERFACEREAD_QR.OpenResultset(1, 1)
    '''
    '''                    If INTERFACEREAD_RS.RowCount > 0 Then
    '''                        INTFAC = INTERFACEREAD_RS("testresult")
    '''                        intmode = "I"
    '''                    Else
    '''                        INTFAC = ""
    '''                        intmode = ""
    '''                    End If
    '''
    '''If SELDTEMPLETRS.RowCount > 0 Then
    '''If SELMA101RS.RowCount > 0 Then
    '''    If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
    '''        MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & intmode & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
    '''    Else
    '''        SELDTEMPLETRS.MoveNext
    '''        MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & intmode & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
    '''    End If
    '''    i = i + 1
    '''End If
    '''End If
    '''End If
    '''End If
    '''seldmulprocrs.MoveNext
    '''Wend
    '''    MSFlexGrid1.Row = 1
    '''    MSFlexGrid1.Col = 2
    '''    txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
    '''    txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    '''    txt_observed.Visible = True
    '''End If
    '''End If
    ''''End If
    
    
    '''' THIS CODE FOR VIEW BUTTON****************************************************************************************
    ''/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''*********************************************************************************************************************
    ''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
If ADDENDUMFALG = True Then
    ADDENDUMFALG = False
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
                lbl_rwd.Visible = False
            ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                lblicu.Visible = False
                lbl_rwd.Visible = True
            End If
        End If
        '******ICU****************
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
        txt_age = DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                
        '               If Trim(txt_age.Text) = 0 Then
        '                   Dim N As Integer
        '                   N = txt_age.Text / 12
        '                   txt_age.Text = N
        '               End If
                   
        If Trim(txt_age.Text) < 12 Then
            'Dim age2 As String
            age2 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            age = "." & (age Mod 12)
            If age2 = ".0" Then
                age2 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                age2 = ".0" & (age2 Mod 30)
            End If
            txt_age.Text = Trim(age2)
        End If
                
        txt_pridoc.Text = SELTA101_RS("ud_name")
        txt_category.Text = SELTA101_RS("pat_catg")
                
        If REFDRFLAG1 = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            REFDRFLAG1 = True
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
        If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
        End If
        If SELPATHDREQUISITNFALG1 = False Then
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFALG1 = True
        Else
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
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
            If Trim(selpathdrequisitn_rs("spihreqno")) = "U" Then
            Label9.Visible = True
            Else
            Label9.Visible = False
            End If
        End If
                
        If SELPATHDREQUISITNFALG1 = False Then
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFALG1 = True
        Else
            SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
                
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
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
            ' txt_procedure.Text = selta1012_rs("prvdiagdes")
        End If
        
    
     ''/////////////////////////////////////////////////////////////////////////////////////////
     '' CODE FOR DISPLAY DATA IN FLEX GRID
     ''*****************************************************************************************
     ' WRITTEN BY AMIT FOR INTERFACING
       
        ICASENO = MaskEdBox1.FormattedText
        IREQNO = VAR_REQNNO
'        Dim TEST_DET1, RANGE1, MACHINENO1
        MSFlexGrid1.Rows = 1
        I = 1
        'FETCH TEST_CODE FROM DMULPROCE
        If seldmulprocflag = False Then
            SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
            seldmulprocflag = True
        Else
            SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            seldmulprocrs.MoreResults
            seldmulprocrs.Requery
        End If
    
            Dim sqlstr1 As String
            Dim SELADDINTERFACE_QR1 As rdoQuery
            Dim SELADDINTERFACE_RS1 As rdoResultset

        sqlstr1 = "SELECT * FROM DISLIB.DMPROCEADD WHERE REQNO=? AND OBSERVALUE IS NULL"
            Set SELADDINTERFACE_QR1 = rdb.CreateQuery("", sqlstr1)
            SELADDINTERFACE_QR1(0) = Replace(Trim(cmb_reqno.Text), "/", "")
            Set SELADDINTERFACE_RS1 = SELADDINTERFACE_QR1.OpenResultset(1, 3)
           
                 If SELADDINTERFACE_RS1.RowCount > 0 Then
            Do Until SELADDINTERFACE_RS1.EOF
                PC = SELADDINTERFACE_RS1!PRO_CODE
                'THIS IS FOR GETTING TEST DESCRIPTION FOR TEST CODE
                If SELMA101FLAG = False Then
                    selMA101QR(0) = PC
                    Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    selMA101QR(0) = PC
                    SELMA101RS.MoreResults
                    SELMA101RS.Requery
                End If
                TEST_DET = Trim(SELMA101RS!SER_DET)
                'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE
                If SELDTEMPLETFLAG = False Then
                   SELDTEMPLETQR(0) = PC
                   Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                   SELDTEMPLETFLAG = True
                Else
                   SELDTEMPLETQR(0) = PC
                   SELDTEMPLETRS.MoreResults
                   SELDTEMPLETRS.Requery
                End If
                'THIS IF FOR F869 (A & C) IN DTEMPLET
                If PC = "F869" Then
                    ALKAGE
                    If SELDBCALKPHSFLAG = False Then
                        SELDBCALKPHSQR(0) = PC
                        SELDBCALKPHSQR(1) = RNAGE
                        SELDBCALKPHSQR(2) = SS
                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
                        SELDBCALKPHSFLAG = True
                    Else
                        SELDBCALKPHSQR(0) = PC
                        SELDBCALKPHSQR(1) = RNAGE
                        SELDBCALKPHSQR(2) = SS
                        SELDBCALKPHSRS.MoreResults
                        SELDBCALKPHSRS.Requery
                    End If
                    On Error Resume Next
                    RNAGE1 = SELDBCALKPHSRS!range
                    'justin 18-08-2004
                    ALKLV = Left(Trim(SELDBCALKPHSRS("RANGE")), 2)
                    ALKHV = Right(Trim(SELDBCALKPHSRS("RANGE")), 3)
                End If
                  
                'SHWETA 10-01-2013
''''23-02-2013
'NEW REFERANCE RANGE DATA
'                Dim RS_VWREFRNG As rdoResultset
                 Set RS_VWREFRNG = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & PC & "' ", 1, 1)
        
                
                If Not RS_VWREFRNG.EOF Then
'                Dim agecal As Double
                agecal = DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                txt_age.Text = DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                If Trim(txt_age.Text) <= 1 Then
'                Dim age3 As Double
                 age3 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                age3 = "." & age3 '(age Mod 12)
                If age3 = ".0" Then
                    age3 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                    age3 = ".0" & (age3 Mod 30)
                End If
                txt_age.Text = Trim(age3)
            End If
                   ' ALKAGEO 'FOR F876
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & PC & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range

                    If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Or InStr(Trim(SELDBCALKPHSGGTRS!range), "-") Then
                    Else
                    ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                    ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                    End If
                    End If
                End If
                
                
                ''''''''
                  
                If PC = "F999" Then
                    GTTAGE 'FOR GGT
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE='" & PC & "' AND (AGE1<" & RNAGE & " AND AGE2>=" & RNAGE & ") AND SEX='" & SS & "'", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range
                    
                    If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Then
                    Else
                    ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 2)
                    ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                    End If
                    End If
                End If
                'justin 18-08-2004
                'THIS IS FOR MALE & FEMALE FROM DTEMPLET
                If PC = "F850" Or PC = "F851" Or PC = "F856" Or PC = "F878" Or PC = "F904" Or PC = "F907" Or PC = "F877" Or PC = "F961" Then
                    If SELDTEMPLETFLAG1 = False Then
                        SELDTEMPLETQR1(0) = PC
                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
                        Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
                        SELDTEMPLETFLAG1 = True
                    Else
                        SELDTEMPLETQR1(0) = PC
                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
                        SELDTEMPLETRS1.MoreResults
                        SELDTEMPLETRS1.Requery
                    End If
                    RNAGE1 = Trim(SELDTEMPLETRS1!LOWVALUE) & "-" & Trim(SELDTEMPLETRS1!HIGHVALUE)
                Else
                    If PC = "F829" Then 'Shraddha 17/12/08
                        If Trim(SELDTEMPLETRS!LOWVALUE) = "" Then
                            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & Trim(SELDTEMPLETRS!HIGHVALUE)
                        Else
                            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                        End If
                    Else
                    'SHWETA BIO REG RANGE CHANGES 09-MAR-2013
'                    Dim RS_VWREFRNG1 As rdoResultset
                 Set RS_VWREFRNG1 = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & PC & "' ", 1, 1)
        
                
              '  09-MAR-2013
                        If PC = "F869" Or PC = "F999" Or PC = "F876" Or PC = "F893" Or PC = "F894" Or RS_VWREFRNG1.EOF = False Then
                         
                        Else
                            SEX = Trim(txt_sex.Text)
                            If Not SELDTEMPLETRS.EOF Then
                            
                                If Trim(SELDTEMPLETRS("SEX")) = "" Then
                                    RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                                Else
                           If Trim(PC) = "F126" Then
                       If Trim(txt_age.Text) <= 15 Then
                          SEX = "C"
                       Else
                          SEX = "A"
                       End If
                       
                    End If
                    
                    
                    
Back1:

                            If Trim(SELDTEMPLETRS("SEX")) = "" And PC = "F124" Then
                                    RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                            End If
If Trim(SEX) = Trim(SELDTEMPLETRS("SEX")) Then
                                    If Trim(SELDTEMPLETRS!LOWVALUE) = "" Then
                                        RNAGE1 = Trim(SELDTEMPLETRS!HIGHVALUE)
                                    Else
                                        RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                                    End If
                                    
                                Else
                                    SELDTEMPLETRS.MoveNext
                                    If Not SELDTEMPLETRS.EOF Then GoTo Back
                                End If
                                End If
                            End If
                        End If
                    End If
                End If
                'THIS IS FOR PICKUP TESTRESULT FROM INTERFACE/DINTERFACE
    
                INTERFACEREAD_QR(0) = IREQNO
                INTERFACEREAD_QR(1) = PC
                Set INTERFACEREAD_RS = INTERFACEREAD_QR.OpenResultset(1, 1)
    
                If INTERFACEREAD_RS.RowCount > 0 Then
                    INTFAC = INTERFACEREAD_RS("testresult")
                    If Trim(INTFAC) = "" Then
                        INTMODE = ""
                        MACHINENO = ""
                    Else
                        INTMODE = "I"
                        MACHINENO = INTERFACEREAD_RS("MACHINENO")
                    End If
                Else
                    INTFAC = ""
                    INTMODE = ""
                End If
                If INTERFACEREAD_RS.RowCount = 0 Then
                    '''FROM HERE
                   ' SELMSADINTF_QR(0) = IREQNO
                    'SELMSADINTF_QR(1) = PC
                    'Set SELMSADINTF_RS = SELMSADINTF_QR.OpenResultset(1, 1)
                    
'                    If SELMSADINTF_RS.RowCount > 0 Then
'                        INTFAC = SELMSADINTF_RS("testresult")
'                        If Trim(INTFAC) = "" Then
'                            INTMODE = ""
'                            MACHINENO = ""
'                        Else
'                            INTMODE = "I"
'                            MACHINENO = SELMSADINTF_RS("MACHINENO")
'                        End If
'                    Else
'                        INTFAC = ""
'                        INTMODE = ""
'                    End If
'                    'UPTO HERE

                    MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & vbTab & MACHINENO '& vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
                Else
                    MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL"))) & vbTab & MACHINENO
                    'THIS IS FOR DEFINE PANIC VALUE BASED ON OBSER VALUE
                    If Len(Trim(MSFlexGrid1.TextMatrix(I, 2))) = 0 Then
                    Else
                        'JUSTIN 18 - 8 - 2004
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                        NU = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), NU + 1)) Then
                            MSFlexGrid1.row = I
                            MSFlexGrid1.col = 2
                            MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = Val(Trim(MSFlexGrid1.TextMatrix(I, 2)))
                            Else
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                            End If
                        Else
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                MSFlexGrid1.row = I
                                MSFlexGrid1.CellBackColor = vbWhite
                                'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Val(Trim(MSFlexGrid1.TextMatrix(i, 2)))
                                If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) Then
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                                Else
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                                End If
                            End If
                            'If Len(Trim(MSFlexGrid1.TextMatrix(i, 2))) = 0 Then
                            'Else
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                MSFlexGrid1.row = I
                                MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
                                If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) Then
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                                Else
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                                End If
                            End If
                        End If
                    End If
                End If
                End If
                SELADDINTERFACE_RS1.MoveNext
                I = I + 1
            Loop
        End If
            
End If
    
    If viewflag = True Then
        viewflag = False
    
            ''******ICU****************
        VAR_REQNNO = req_chglen(Trim(cmb_reqno.Text))
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
              lbl_rwd.Visible = False
           ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
              lblicu.Visible = False
              lbl_rwd.Visible = True
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
              'cmd_adenum.Visible = True
          End If
          cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
          cmb_finalby.Text = "" & REFDR_RS("FREPORTBY")
          
          If Trim(REFDR_RS("REQNSTAT")) = "T" Then
             cmd_commit.Enabled = True
          Else
             cmd_commit.Enabled = False
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
          'STAT SAMPLE CHANGES
                If Trim(selpathdrequisitn_rs("SPIHREQNO")) = "U" Then
                Label9.Visible = True
                Else
                Label9.Visible = False
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
        txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
        If selpathdrequisitn_rs.EOF = False Then
            txt_relclnhis.Visible = True
    
            txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
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
    
    
    ''**** filling the flex grid
    If SELDPROCEDURFLAG = False Then
        SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
        Set SELDPROCEDURRS = SELDPROCEDURQR.OpenResultset(1, 3)
        SELDPROCEDURFLAG = True
    Else
        SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
        SELDPROCEDURRS.MoreResults
        SELDPROCEDURRS.Requery
    End If
        If SELDPROCEDURRS.RowCount > 0 Then
                 I = 1
                 srno = 1
             While Not SELDPROCEDURRS.EOF
                 If SELMA101FLAG = False Then
                     selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
                     Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
                     
                     SELMA101FLAG = True
                 Else
                     selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
                     SELMA101RS.MoreResults
                     SELMA101RS.Requery
                 End If
                If SELDTEMPLETFLAG = False Then
                    SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                    SELDTEMPLETFLAG = True
                Else
                    SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    SELDTEMPLETRS.MoreResults
                    SELDTEMPLETRS.Requery
                End If
                If Trim(SELDPROCEDURRS("PRO_CODE")) = "F869" Then
                    'ALKLV = Left(Trim(SELDPROCEDURRS("NRANGE")), 3)
                    'Commented and added by Shraddha on 09/12/10
                    ALKLV = Left(Trim(SELDPROCEDURRS("NRANGE")), 2)
                    ALKHV = Right(Trim(SELDPROCEDURRS("NRANGE")), 3)
                ElseIf Trim(SELDPROCEDURRS("PRO_CODE")) = "F999" Then
                    ALKLV1 = Left(Trim(SELDPROCEDURRS("NRANGE")), 3)
                    ALKHV1 = Right(Trim(SELDPROCEDURRS("NRANGE")), 3)
                End If
    
                If SELDTEMPLETRS.RowCount > 0 Then
        
                    If SELMA101RS.RowCount > 0 Then
                        
                        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                        MSFlexGrid1.row = srno
                        MSFlexGrid1.col = 0
                        
                            'added by shweta fro addendum
                            If IsNull(SELDPROCEDURRS("OBSERVALUE")) = False Or Trim(SELDPROCEDURRS("OBSERVALUE")) <> "" Then
                            
                            
                            
                        'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
                            MSFlexGrid1.TextMatrix(I, 0) = I
                            MSFlexGrid1.TextMatrix(I, 1) = Trim(SELMA101RS("SER_DET")) & "$" & Trim(SELDPROCEDURRS("PRO_CODE"))
                            If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
                                MSFlexGrid1.row = srno
                                MSFlexGrid1.col = 2
                                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                            Else
                                If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
    
                                    MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(100, 255, 150)
                                End If
                                If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
    
                                     MSFlexGrid1.col = 2
                                     MSFlexGrid1.CellBackColor = vbWhite
                                End If
                                If Trim(SELDPROCEDURRS("obser_type")) = "" Then
    
                                    MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = vbWhite
                                End If
                            End If
                            j = 0
                            j = InStr(SELDPROCEDURRS("observalue"), "*")
                            If j = 0 Then
                                MSFlexGrid1.TextMatrix(I, 2) = SELDPROCEDURRS("OBSERVALUE")
                            Else
                                MSFlexGrid1.TextMatrix(I, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, j - 1)
                            End If
                            MSFlexGrid1.TextMatrix(I, 3) = Trim(SELDPROCEDURRS("nrange")) 'Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                            MSFlexGrid1.TextMatrix(I, 4) = SELDPROCEDURRS("pvalue")
                        Else
                            SELDTEMPLETRS.MoveNext
                            MSFlexGrid1.TextMatrix(I, 0) = I
                            MSFlexGrid1.TextMatrix(I, 1) = Trim(SELMA101RS("SER_DET")) & "$" & Trim(SELDPROCEDURRS("PRO_CODE"))
                            If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
                                MSFlexGrid1.row = srno
                                 MSFlexGrid1.col = 2
                                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                            Else
                            If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
            
                              MSFlexGrid1.col = 2
                              MSFlexGrid1.CellBackColor = RGB(150, 255, 100)
                            End If
                            If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
            
                               MSFlexGrid1.col = 2
                               MSFlexGrid1.CellBackColor = vbWhite
                            End If
                            If Trim(SELDPROCEDURRS("obser_type")) = "" Then
            
                                MSFlexGrid1.col = 2
                              MSFlexGrid1.CellBackColor = vbWhite
                            End If
                        End If
            
                        j = 0
                        j = InStr(SELDPROCEDURRS("observalue"), "*")
                        If j = 0 Then
                            MSFlexGrid1.TextMatrix(I, 2) = Trim(SELDPROCEDURRS("OBSERVALUE"))
                        Else
                            MSFlexGrid1.TextMatrix(I, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, j - 1)
                        End If
                        MSFlexGrid1.TextMatrix(I, 3) = Trim(SELDPROCEDURRS("nrange")) 'Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                        MSFlexGrid1.TextMatrix(I, 4) = SELDPROCEDURRS("pvalue")
    
                    End If
                    
                    
                     Else
                            
                     
                     
                    'added by shweta for addendum
                    Dim QR_dadd As rdoQuery
                    Dim RS_dadd As rdoResultset
                    str = "SELECT * FROM DISLIB.dmproceadd WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "'"
                    Set QR_dadd = rdb.CreateQuery("", str)
                    Set RS_dadd = QR_dadd.OpenResultset(1, 3)
                    If RS_dadd.EOF = False Then
                                
                                
                              
                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
                            MSFlexGrid1.TextMatrix(I, 0) = I
                            MSFlexGrid1.TextMatrix(I, 1) = Trim(SELMA101RS("SER_DET")) & "$" & Trim(RS_dadd("PRO_CODE"))
                            If Trim(RS_dadd("obser_type")) = "H" Then
                                MSFlexGrid1.row = srno
                                MSFlexGrid1.col = 2
                                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                            Else
                                If Trim(RS_dadd("obser_type")) = "L" Then
    
                                    MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(100, 255, 150)
                                End If
                                If Trim(RS_dadd("obser_type")) = "N" Then
    
                                     MSFlexGrid1.col = 2
                                     MSFlexGrid1.CellBackColor = vbWhite
                                End If
                                If Trim(RS_dadd("obser_type")) = "" Then
    
                                    MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = vbWhite
                                End If
                            End If
                            j = 0
                            j = InStr(RS_dadd("observalue"), "*")
                            If j = 0 Then
                                MSFlexGrid1.TextMatrix(I, 2) = RS_dadd("OBSERVALUE")
                            Else
                                MSFlexGrid1.TextMatrix(I, 2) = Mid(RS_dadd("OBSERVALUE"), 1, j - 1)
                            End If
                            MSFlexGrid1.TextMatrix(I, 3) = Trim(RS_dadd("nrange")) 'Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                            MSFlexGrid1.TextMatrix(I, 4) = RS_dadd("pvalue")
                        Else
                            SELDTEMPLETRS.MoveNext
                            MSFlexGrid1.TextMatrix(I, 0) = I
                            MSFlexGrid1.TextMatrix(I, 1) = Trim(SELMA101RS("SER_DET")) & "$" & Trim(RS_dadd("PRO_CODE"))
                            If Trim(RS_dadd("obser_type")) = "H" Then
                                MSFlexGrid1.row = srno
                                 MSFlexGrid1.col = 2
                                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                            Else
                            If Trim(RS_dadd("obser_type")) = "L" Then
            
                              MSFlexGrid1.col = 2
                              MSFlexGrid1.CellBackColor = RGB(150, 255, 100)
                            End If
                            If Trim(RS_dadd("obser_type")) = "N" Then
            
                               MSFlexGrid1.col = 2
                               MSFlexGrid1.CellBackColor = vbWhite
                            End If
                            If Trim(RS_dadd("obser_type")) = "" Then
            
                                MSFlexGrid1.col = 2
                              MSFlexGrid1.CellBackColor = vbWhite
                            End If
                        End If
            
                        j = 0
                        j = InStr(RS_dadd("observalue"), "*")
                        If j = 0 Then
                            MSFlexGrid1.TextMatrix(I, 2) = Trim(RS_dadd("OBSERVALUE"))
                        Else
                            MSFlexGrid1.TextMatrix(I, 2) = Mid(RS_dadd("OBSERVALUE"), 1, j - 1)
                        End If
                        MSFlexGrid1.TextMatrix(I, 3) = Trim(RS_dadd("nrange")) 'Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                        MSFlexGrid1.TextMatrix(I, 4) = RS_dadd("pvalue")
    
                    End If
                    MSFlexGrid1.row = I
                           For kp = 0 To MSFlexGrid1.Cols - 1
                            MSFlexGrid1.col = kp
                            MSFlexGrid1.row = I
                            MSFlexGrid1.CellBackColor = vbCyan
        
                          Next
            End If
                    
                    'tiil here for addendum
                                
                                
                                
                                
                            End If
                End If
    
                    If IsNull(SELDPROCEDURRS("SOURCE")) = False Then
                        MSFlexGrid1.TextMatrix(I, 5) = SELDPROCEDURRS("SOURCE")
                    End If
    
                    If IsNull(SELDPROCEDURRS("REMARKS")) = False Then
                        MSFlexGrid1.TextMatrix(I, 7) = Trim(SELDPROCEDURRS("REMARKS"))
                    End If
                       srno = srno + 1
                        I = I + 1
    
                End If
                SELDPROCEDURRS.MoveNext
    
            Wend
            txt_observed.Visible = False
    
       End If
    
        If SELDMBCCOMMFLAG = False Then
            SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
            Set SELDMBCCOMMRS = SELDMBCCOMMQR.OpenResultset(1, 3)
            SELDMBCCOMMFLAG = True
        Else
            SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
            SELDMBCCOMMRS.MoreResults
            SELDMBCCOMMRS.Requery
        End If
        If SELDMBCCOMMRS.RowCount > 0 Then
            If SELDMBCCOMMRS.EOF = False Then
                TXT_COMMENTS.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
            End If
        End If
    End If
    End If
    End If
'End If

'***** Shraddha 25/11/2011 *****
'For showing previous values
Call prev_val
'*******************************

If opt_staff.Value = True Then
    STAFF
End If


'ADDED BY SHWETA DEC 2015
Call UV


Exit Sub
MSG:
MsgBox Err.Description

End Sub
Sub UV()
LBL_OPPROC.Caption = ""
Dim QR_UV As rdoQuery
Dim RS_UV As rdoResultset
str = "SELECT * FROM DISLIB.drequisitn WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "'"
Set QR_UV = rdb.CreateQuery("", str)

Set RS_UV = QR_UV.OpenResultset(1, 3)

If RS_UV.EOF = False Then
If IsNull(RS_UV("OPPROCEDUR")) = True Then
Else
LBL_OPPROC.Caption = RS_UV("OPPROCEDUR")
txt_relclnhis.Text = "24 hours Urine Volume" & " - " & LBL_OPPROC.Caption
End If

Else
 LBL_OPPROC.Caption = ""
End If
End Sub
Public Sub prev_val()

'*********** T F *********
SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
SELRREQUISITN_QR(3) = "T"
SELRREQUISITN_QR(4) = "F"
Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)

If Not SELRREQUISITN_RS.EOF Then


'*********************************************

Dim DMPROCEDUR_QR As rdoQuery
Dim DMPROCEDUR_RS As rdoResultset
str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO = ? "
Set DMPROCEDUR_QR = rdb.CreateQuery("", str)
DMPROCEDUR_QR(0) = Trim(SELRREQUISITN_RS("reqnno")) 'req_chglen(cmb_reqno.List(1))
Set DMPROCEDUR_RS = DMPROCEDUR_QR.OpenResultset(1, 3)
While Not DMPROCEDUR_RS.EOF
    I = 1
    While I < MSFlexGrid1.Rows
    
        If Trim(DMPROCEDUR_RS("PRO_CODE")) = Right(Trim(MSFlexGrid1.TextMatrix(I, 1)), 4) Then
            'MsgBox "match"
            MSFlexGrid1.TextMatrix(I, 6) = Trim(DMPROCEDUR_RS("OBSERVALUE"))
        End If
        I = I + 1
    Wend
    DMPROCEDUR_RS.MoveNext
Wend

End If
End Sub

Private Sub cmd_addendum_Click()

flagsave = 3
GoForAddendum  'added by shweta 30-09-2016 for all comments
adenum_flg = True
BTNENABL1 ("VIEW")
FLDENABL ("MOD")
'field_clear
'TXT_CLEAR

ADDENDUMFALG = True
cmd_adenum_Click
Exit Sub

On Error GoTo MSG:

ADDENDUMFALG = True
viewflag = False
flagsave = 3


BTNENABL1 ("VIEW")
FLDENABL ("MOD")
field_clear
TXT_CLEAR
MSFlexGrid1.clear
MSFlexGrid1.FormatString = "S.No.  |^                 Procedure Name                        |^    Observed Value   |^        Normal Range of the Procedure |^ Panic Value"
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
ADDENDUMFALG = True
viewflag = False
'cmb_dept.SetFocus
'cmb_modality.Enabled = True
MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True
Opt_patient.Value = True
MaskEdBox1.SetFocus
Exit Sub
MSG:
MsgBox Err.Description



'On Error GoTo MSG:
'
'
'flagsave = 3
'
'
'BTNENABL1 ("VIEW")
'FLDENABL ("VIEW")
'field_clear
'TXT_CLEAR
'MSFlexGrid1.clear
'MSFlexGrid1.FormatString = "S.No.  |^                 Procedure Name                        |^    Observed Value   |^        Normal Range of the Procedure |^ Panic Value"
'If DATETIMEFLAG = False Then
'   DATETIME
'   DATETIMEFLAG = True
'Else
'    DATETIME
'End If
'txt_date.Text = DATE11
'txt_time.Text = time1
'If SEL_DEPTRSFLAG = False Then
'    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
'    SEL_DEPTFLAG = True
'Else
'    sel_deptrs.MoreResults
'    sel_deptrs.Requery
'End If
'ADDENDUMFALG = True
'viewflag = False
''cmb_dept.SetFocus
''cmb_modality.Enabled = True
'MaskEdBox1.Enabled = True
'cmb_reqno.Enabled = True
'cmd_print.Enabled = True
'Opt_patient.Value = True
'MaskEdBox1.SetFocus
'Exit Sub
'MSG:
'MsgBox Err.Description
End Sub

Private Sub cmd_adenum_Click()
 flagsave = 3
 adenum_flg = True
 lbl_addendnum.Visible = True
 txt_addendum.Visible = True
 txt_addendum.Enabled = True
 Label4.Visible = True
 cmb_addendumby.Visible = True
 cmd_save.Enabled = True
End Sub

Private Sub cmd_back_Click()
'Me.Hide
'DIS_MAIN.show
''btnenabl ("INI")
'BTNENABL1 ("INI")
'field_clear
'viewflag = False
'NEWFLAG = False
'BCCLEAR
SELDEPTRSFLAG1 = False
cmd_cancel_Click
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
Private Sub cmd_commit_Click()
If MSFlexGrid1.Visible = True Then

 Set SELREPDN_RS = rdb.OpenResultset("SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO='" & Replace(Trim(cmb_reqno.Text), "/", "") & "'  ", 1, 1)
                   
If SELREPDN_RS.EOF = True Then

    updrrequisitn1_qr(0) = UCase(cmb_finalby.Text)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "E"
    updrrequisitn1_qr(3) = ""
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute

MsgBox "PLEASE REPORT IS NOT ENTERED PROPERLY PLEASE REENTER USING NEW!!!!!!!"
Exit Sub
End If
End If
                   


Dim I As Integer
'If Len(Trim(cmb_finalby)) = 0 Then
'    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
'    cmb_finalby.SetFocus
'    Exit Sub
'End If
I = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If I = vbYes Then
On Error GoTo MSG:
  
 
''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
    
        If (LOGINTYP > CURUSRDET_RS("usrtype") Or LOGINTYP >= 6) Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
'            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
'            cmd_commit.SetFocus
'            Exit Sub
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
            'If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
                '  If loginusrid <> curusrdetrs("usr_id") Then
             '   MsgBox "Sorry! This Report is not to be Finalized by You"
              '  cmd_commit.SetFocus
               ' Exit Sub
            'End If
        End If
        If CURUSRDET_RS("COMMIT") <> "YES" Then
'            MsgBox "Sorry! You are not authorised to Commit the Report of this Modality"
'            cmd_commit.SetFocus
'            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
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
    MsgBox "Record is Committed Successfully."
    
   'rtimelog
   DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
   
   
   UPDRTIMEC_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(txt_time.Text, "HH:MM:SS")
      
   
    
    UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute
   'rtimelog
    
End If


'******  CRITICAL MSG
Dim CR_PRCODE As String

 Dim CR_I As Integer
 Dim CR_RES_Q As rdoQuery
 Dim CR_RES_R As rdoResultset
 Dim CR_STR As String
 
 
 Dim SELSMSTBL_QR As rdoQuery
Dim SELSMSTBL_RS As rdoResultset

Dim SELMOBNO_QR As rdoQuery
Dim SELMOBNO_RS As rdoResultset


Dim SELSMSSEND_RS As rdoResultset
' Dim dbsql As New ADODB.Connection
' CR_STR = "SELECT * FROM DISLIB.CRTALERT WHERE PROCODE=? "
''CR_STR = "SELECT * FROM DISLIB.CRTALERT WHERE PROCODE=? AND AGE2 > ? AND AGE1 < ? "
''MsgBox "crstr"
' Set CR_RES_Q = rdb.CreateQuery("", CR_STR)
''CR_RES_Q(1) = CInt(txt_age.Text)
''CR_RES_Q(2) = CInt(txt_age.Text)
' 'MsgBox CR_STR
' For CR_I = 1 To MSFlexGrid1.Rows - 1
'  CR_PRCODE = Right(MSFlexGrid1.TextMatrix(CR_I, 1), 4)
'    CR_RES_Q(0) = Trim(CR_PRCODE)
'    Set CR_RES_R = CR_RES_Q.OpenResultset(1, 3)
'        If CR_RES_R.EOF = False Then
'        Dim STRSQL As String
'        Dim UPD_DMULFORCRT As rdoQuery
'        While Not CR_RES_R.EOF
'                CR_AGE2 = CR_RES_R("AGE2")
'                CR_AGE1 = CR_RES_R("AGE1")
'                'MsgBox "Age check"
'                If CR_AGE2 > CInt(txt_age.Text) And CR_AGE1 < CInt(txt_age.Text) Then
'                    RHVAL = Trim(CR_RES_R("H_CRVALUE"))
'                    RLVAL = Trim(CR_RES_R("L_CRVALUE"))
'                    PVAL = Replace(Trim(MSFlexGrid1.TextMatrix(CR_I, 2)), "#", "")
'                '    MsgBox "Critical Values so MSG Will get send to Patient"
'                    If Val(PVAL) < Val(RLVAL) Or Val(PVAL) > Val(RHVAL) Then
'                    'INSERT TO TABLE FOR REFERENCE 22-AUG-2014
'                   ' MsgBox "now it sill send"
'                    ''commented by shweta dec2015 bcoz table is locked
''                    STRSQL = "UPDATE DISLIB.DMULPROCE SET COLLECTION='C' WHERE REQNO=? AND PRO_CODE=?"
''                     Set UPD_DMULFORCRT = rdb.CreateQuery("", STRSQL)
''                    UPD_DMULFORCRT(0) = ""
''                    UPD_DMULFORCRT(1) = Trim(CR_PRCODE)
''                    UPD_DMULFORCRT.Execute
'
'                    'ANOTHER TABLE
'                     str = "insert into DISLIB.CRTSMSLOG(REQNO,CASENO,PRO_CODE,OB_VAL,RNG,UID,TDATE,TTIME,REM,FLAG1,AGE) values ('" & Replace(Trim(cmb_reqno.Text), "/", "") & "','" & Trim(MaskEdBox1.FormattedText) & "','" & Trim(CR_PRCODE) & "','" & Trim(PVAL) & " ','" & Trim(RLVAL) & "-" & Trim(RHVAL) & "','" & Trim(LOGINUSRID) & "',CURRENT DATE,CURRENT TIME,'','','" & Trim(txt_age.Text) & "')"
'
'                        rdb.Execute str
'
'                    SelStr = "SELECT * FROM PATOOLSLIB.TBLSMSMSGS WHERE MSGID='26'"
'                    Set SELSMSTBL_QR = rdb.CreateQuery("", SelStr)
'                     Set SELSMSTBL_RS = SELSMSTBL_QR.OpenResultset(1, 1)
'                       If SELSMSTBL_RS.EOF = False Then
'                            If Trim(SELSMSTBL_RS("MSGDESC")) <> "" And Len(Trim(SELSMSTBL_RS("MSGDESC"))) <> 0 Then
'                              MSG = Trim(SELSMSTBL_RS("MSGDESC"))
'                                  '  MSG = Trim(SELSMSTBL_RS("MSGDESC"))
'                            MSG = Replace(MSG, "@1", Trim(MaskEdBox1.FormattedText))
'                            MSG = Replace(MSG, "@2", Replace(Trim(cmb_reqno.Text), "/", ""))
'                            MSG = Replace(MSG, "@3", Trim("Dept. of Biochemistry"))
'
'            Set rs_smsdtls = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.VWSMSCASES WHERE CASENO='" & Trim(MaskEdBox1.FormattedText) & "' ", 1, 1)
'            If Not rs_smsdtls.EOF Then
'                If Trim(rs_smsdtls("pmob")) <> "" Then
'
'
'
'               Set SELSMSSEND_RS = rdb.OpenResultset("SELECT * FROM DISLIB.RPTSMSLOG WHERE REQNO='" & Replace(Trim(cmb_reqno.Text), "/", "") & "' AND MSGID='26' ", 1, 1)
'
'                   If SELSMSSEND_RS.EOF = True Then
'                     Set rs_server = rdb.OpenResultset("select server from patoolslib.serverip where APPLICATN ='SMS'", 1, 1)
'
'                        dbsql.Open ("Provider=SQLOLEDB.1;Password=tmhit123;Persist Security Info=True;User ID=SA;Initial Catalog=SMS_DB;Data Source=" & Trim(rs_server(0)))
'                       ' Set rq_ins = dbsql.CreateQuery("", "insert into SMS_IO_LOG_MASTER(SMS_MOBILE,SMS_TEXT,STAMP,PDU_SMS,PRIORITY_INDEX,SMS_METHOD,SMS_IO_IND) values ('" & Trim(rs_smsdtls("MOBILE1")) & "','" & Trim(MSG) & "',getdate(),0,'" & Trim(rs_smsdtls("priority")) & "','2','3')")
'                        str = "insert into SMS_IO_LOG_MASTER(SMS_MOBILE,SMS_TEXT,STAMP,PDU_SMS,PRIORITY_INDEX,SMS_METHOD,SMS_IO_IND,transaction_id,case_no,msgID) values ('" & Trim(rs_smsdtls("pmob")) & "','" & Trim(MSG) & "',getdate(),0,'" & Trim(SELSMSTBL_RS("priority")) & "','2','3','" & Replace(Trim(cmb_reqno.Text), "/", "") & "','" & Trim(MaskEdBox1.FormattedText) & "','26')"
'
'                        dbsql.Execute str
'                        dbsql.Close
'                        MsgBox "Crtitical Alert SMS Sent to Patient(Case No: " & Trim(MaskEdBox1.FormattedText) & ""
'
'                        str = "insert into DISLIB.RPTSMSLOG(REQNO,CASE_NO,REQSTAT,SMSDEL_DT,SMSDEL_TM,FLAG,IPADD,MSGID) values ('" & Replace(Trim(cmb_reqno.Text), "/", "") & "','" & Trim(MaskEdBox1.FormattedText) & "','F',CURRENT DATE,CURRENT TIME,'Y','" & Trim(IPAD) & "','26')"
'
'                        rdb.Execute str
'                        Else
'                        MsgBox "Crtitical Alert SMS Already Sent to Patient(Case No: " & Trim(MaskEdBox1.FormattedText) & ""
'
'                    End If
'
'                    Else
'                    MsgBox "Patient's mobile number is not Registered, So Critical Alert Message will not send to Patient."
'                 End If
'                 End If
'            End If
'
'                       End If
'                       End If
'                'MsgBox "TRUE"
'                Else
'
'                End If
'                CR_RES_R.MoveNext
'                Wend
'        End If
' Next
'
'
'
''******'''TILL HERE



rdb.CommitTrans


 
'insertIntoRAUTHOSIGN "U", 1, VAR_REQNNO, MaskEdBox1.FormattedText, cmb_enterby.Text, Winsock1.LocalIP, cmb_enterby.Text, Winsock1.LocalIP, cmb_authSign.Text, Winsock1.LocalIP, cmb_finalby.Text, Winsock1.LocalIP, "", ""

GoForAddendum

'btnenabl ("INI")
BTNENABL1 ("INI")
'TXT_CLEAR
'field_clear
cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub
Sub GoForAddendum()

'Dim sqlstr As String
'Dim SELADDINTERFACE_QR As rdoQuery
'Dim SELADDINTERFACE_RS As rdoResultset
'
' sqlstr = "SELECT * FROM interface.B_INTERFAC WHERE REQNO=? AND IFLAG='1' AND REQNO NOT IN (SELECT REQNO FROM DISLIB.DMPROCEADD WHERE REQNO='" & Trim(Replace(Trim(cmb_reqno.Text), "/", "")) & "')"
'  Set SELADDINTERFACE_QR = rdb.CreateQuery("", sqlstr)
' SELADDINTERFACE_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
'Set SELADDINTERFACE_RS = SELADDINTERFACE_QR.OpenResultset(1, 3)
'If SELADDINTERFACE_RS.EOF = False Then
'While SELADDINTERFACE_RS.EOF = False
'
'  str = "insert into DISLIB.DMPROCEADD(REQNO,PRO_CODE) VALUES('" & Trim(Replace(Trim(cmb_reqno.Text), "/", "")) & "','" & Trim(SELADDINTERFACE_RS("TESTCODE")) & "')"
'  rdb.Execute str
'  SELADDINTERFACE_RS.MoveNext
'Wend
'
'End If


Dim sqlstr As String
Dim SELADDINTERFACE_QR As rdoQuery
Dim SELADDINTERFACE_RS As rdoResultset

Dim SELADDDMPROCE_QR As rdoQuery
Dim SELADDDMPROCE_RS As rdoResultset


'sqlstr = "SELECT * FROM INTERFACE.DINTERFACE WHERE REQNO=? AND IFLAG='1' AND REQNO NOT IN (SELECT REQNO FROM DISLIB.DMPROCEADD WHERE REQNO='" & Trim(Replace(Trim(cmb_reqno.Text), "/", "")) & "')"
'FOLLOWING CODE COMMENTED AND ADDED BY SHWETA AS PER USER REQUEST
 'sqlstr = "SELECT * FROM dislib.DMBCCOMM WHERE REQNO=? AND (COMMENTS <>'' or COMMENTS is not null) "
 sqlstr = "SELECT * FROM dislib.DMBCCOMM WHERE REQNO=?"
  Set SELADDINTERFACE_QR = rdb.CreateQuery("", sqlstr)
 SELADDINTERFACE_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
Set SELADDINTERFACE_RS = SELADDINTERFACE_QR.OpenResultset(1, 3)
If SELADDINTERFACE_RS.EOF = False Then

         sqlstr = "SELECT * FROM dislib.DMPROCEDUR WHERE REQNO=? "
         Set SELADDDMPROCE_QR = rdb.CreateQuery("", sqlstr)
         SELADDDMPROCE_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
        Set SELADDDMPROCE_RS = SELADDDMPROCE_QR.OpenResultset(1, 3)
        If SELADDDMPROCE_RS.EOF = False Then
On Error Resume Next
            While SELADDDMPROCE_RS.EOF = False
              str = "insert into DISLIB.DMPROCEADD(REQNO,PRO_CODE) VALUES('" & Trim(Replace(Trim(cmb_reqno.Text), "/", "")) & "','" & Trim(SELADDDMPROCE_RS("PRO_CODE")) & "')"
              rdb.Execute str
              SELADDDMPROCE_RS.MoveNext
            Wend
        End If
End If



End Sub

Private Sub CMD_EMR_Click()
'If Len(Trim(MaskEdBox1.Text)) > 0 Then
'    Dim VAR1
'    'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
'    VAR1 = MaskEdBox1.FormattedText
'    VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
'    AppActivate VAR
'Else
'    MsgBox "Please Enter The Case No."
'End If
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://10.1.4.13/emrtmhact/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
    
End Sub

Private Sub cmd_modify_Click()
'On Error GoTo MSG:
Dim str As String

If adenum_flg = True Then
flagsave = 3
Else
flagsave = 2
End If

'If adenum_flg = True Then
'flagsave = 3
'Else
'flagsave = 2
'End If
'btnenabl ("NEW")
BTNENABL1 ("NEW")
FLDENABL ("MOD")

If MaskEdBox1.Text = "" Then
    MsgBox "PLEASE FETCH THE VALUES FROM VIEW THAT ARE TO BE MODIFIED"
    cmd_view.Enabled = True
    
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


'changes made by amarjeet on 08-11-2002 for  commit logic to disable txt_observed
If ADDENDUMFALG Then 'means commited
    txt_observed.Visible = False
    txt_gttobsvd.Visible = False
    'txt_gttobsvdurn.Visible = False
    TXT_COMMENTS.Enabled = True
Else
    'Load frm_security
    'frm_security.show 1
    'If "RIDER" = UCase(PASSWORD) Then
     '   txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
      '  txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
       ' txt_observed.Visible = True
        'PASSWORD = ""
    'Else
     '   MsgBox "Not Authorised to modify", vbCritical
      '   Exit Sub
        
    'End If
'commented by amarjeet prev logic
'txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'txt_observed.Visible = True
    TXT_COMMENTS.Enabled = True
End If

UPDRTIMEM_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIMEM_QR(1) = Format(txt_time.Text, "HH:MM:SS")
             
UPDRTIMEM_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIMEM_QR.Execute

If Trim(txt_relclnhis.Text) = "Stone Analysis-" Then
    Exit Sub
Else
    MSFlexGrid1.row = 0
    MSFlexGrid1.col = 2
End If
'MSG:
'    MsgBox "PLEASE ENTER THE CASE NUMBER", vbOKOnly, "MODIFY"
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
'On Error GoTo MSG:
'btnenabl ("NEW")
BTNENABL1 ("NEW")
field_clear
TXT_CLEAR
BCCLEAR
MSFlexGrid1.clear
MSFlexGrid2.clear
'MSFlexGrid1.FormatString = "S.No.  |^                 Procedure Name                        |^    Observed Value   |^        Normal Range of the Procedure |^ Panic Value"
'MSFlexGrid1.FormatString = "S.No.  |^    Procedure Name             |^  Observed Value  |^  Normal Range of the Procedure |^ Panic Value|^Source|^Preval      |^Machine "
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
    If Trim(cmb_enterby.List(0)) = Trim(cmb_finalby.List(L1)) Then
        cmb_finalby.Text = Trim(cmb_finalby.List(L1))
    End If
Next L1

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
'    While Not sel_deptrs.EOF
'        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
'        sel_deptrs.MoveNext
'        Wend
'   End If


Opt_patient.Value = True
txt_observed.Visible = False
txt_gttobsvd.Visible = False
'txt_gttobsvdurn.Visible = False
cmb_machine.Visible = False
'cmb_modality.ListIndex = 0
flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
nextpno:
If Lbl_pno.Caption = "" Then
txt_pathno.Text = Format(DATE11, "YYMMDD")
Else
txt_pathno.Text = Val(Lbl_pno.Caption) + 1
End If

Dim VWpathstat_QR As rdoQuery
Dim VWpathstat_RS As rdoResultset

If Len(Trim(txt_pathno.Text)) > 6 Then
str = "SELECT * FROM RISLIB.Rrequisitn where reqnno in (Select reqno from dislib.drequisitn where pathno='" & Trim(txt_pathno.Text) & "')"
Set VWpathstat_QR = rdb.CreateQuery("", str)
'VWDMULPROCE_QR(0) = Trim(txt_pathno.Text)
Set VWpathstat_RS = VWpathstat_QR.OpenResultset(1, 3)
If VWpathstat_RS.EOF Then
Exit Sub
End If
If NEWFLAG = True Then
        If Trim(VWpathstat_RS("reqnstat")) = "E" Then
        Else
        Lbl_pno.Caption = Val(Lbl_pno.Caption) + 1
           GoTo nextpno
        End If
ElseIf viewflag = True Then
        If VWpathstat_RS("reqnstat") = "E" Then
        GoTo nextpno
        Else
        End If

End If
End If
txt_pathno.SelStart = Len(txt_pathno.Text)
txt_pathno.SelLength = 1
txt_pathno.SetFocus
txt_pathno_LostFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_cancel_Click()
'btnenabl ("INI")
BTNENABL1 ("INI")
field_clear
viewflag = False
NEWFLAG = False
BCCLEAR
MSFlexGrid2.clear
'MSFlexGrid2.FormatString = "Time Slots     |^    Blood Sugar        |^Urine Sugar          "
MSFlexGrid2.FormatString = "Time Slots     |^    Blood Sugar        "
    'MSFlexGrid2.MergeCells = flexMergeRestrictRows

End Sub

Public Sub field_clear()
Dim obj As Object
lblicu.Visible = False
lbl_rwd.Visible = False
LBL_OPPROC.Caption = ""
'cmb_dept.ListIndex = -1
'cmb_modality.ListIndex = -1
'txt_deptcode.Text = " "
'cmb_reqno.ListIndex = -1
'following commented by shweta on 06-jan-2015
'cmb_finalby.Text = ""
'cmb_enterby.Text = ""
cmb_reqno.clear

If Opt_patient.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
' MaskEdBox1.Mask = ">??/#####"   ''commented by Monika
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
txt_relclnhis.Text = " "
TXT_COMMENTS.Text = " "

End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In biochemlab_report.Controls
If TypeOf obj Is TextBox Then
    If obj <> txt_deptcode.Text And obj <> txt_modcode.Text Then
        'If obj <> txt_modcode.Text Then
            obj = ""
        'End If
    End If
End If
Next

End Sub


Public Sub FLDENABL(EN As String)
If EN = "DIS" Then
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
        txt_relclnhis.Visible = False
        txt_observed.Visible = False
        txt_gttobsvd.Visible = False
        'txt_gttobsvdurn.Visible = False
        cmb_machine.Visible = False
        TXT_COMMENTS.Enabled = False
        cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
End If
If EN = "ENA" Then
        'cmb_dept.Enabled = True
        'cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        txt_pathno.Enabled = False
        txt_pathno.Enabled = True
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
        txt_observed.Visible = False
        txt_gttobsvd.Visible = False
        'txt_gttobsvdurn.Visible = False
        cmb_machine.Visible = False
        TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
End If
If EN = "MOD" Then
        'cmb_dept.Enabled = True
        'cmb_modality.Enabled = True
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
        
        'txt_gttobsvd.Visible = True
        'txt_gttobsvdurn.Visible = True
        
        If Trim(txt_relclnhis.Text) = "Stone Analysis-" Then
            txt_observed.Visible = False
            cmb_machine.Visible = False
        Else
            txt_observed.Visible = True
            cmb_machine.Visible = True
        End If
        
        TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
End If
'If EN = "VIEW" Then
'        MaskEdBox1.SetFocus
'End If
End Sub


Public Sub BCCLEAR()
TXT_CLEAR
'cmb_dept.Text = ""
'cmb_modality.Text = ""
txt_observed.Visible = False
txt_gttobsvd.Visible = False
'txt_gttobsvdurn.Visible = False
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
MSFlexGrid1.Rows = 1
MSFlexGrid2.Rows = 1
End Sub


Public Sub CLPTM()
Dim RTLOG_QR As rdoQuery
Dim RTLOG_RS As rdoResultset

Dim VWDMULPROCE_QR As rdoQuery
Dim VWDMULPROCE_RS As rdoResultset


str = "SELECT SAMCOLDATE,SAMCOLTIME,ADPREPDATE,ADPREPTIME FROM RISLIB.RTIMELOG WHERE REQNNO=? "
Set RTLOG_QR = rdb.CreateQuery("RTM", str)

str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set VWDMULPROCE_QR = rdb.CreateQuery("VE", str)

VWDMULPROCE_QR(0) = req_chglen(Trim(cmb_reqno.Text))
Set VWDMULPROCE_RS = VWDMULPROCE_QR.OpenResultset(1, 3)


If VWDMULPROCE_RS("PRO_CODE") = "F833" Or VWDMULPROCE_RS("PRO_CODE") = "F834" Or VWDMULPROCE_RS("PRO_CODE") = "F835" Or VWDMULPROCE_RS("PRO_CODE") = "F836" Then

    RTLOG_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    Set RTLOG_RS = RTLOG_QR.OpenResultset(1, 3)

    Dim t1, T2, D1, D2 As String
    Dim H1, H2, h, m1, m2, m As Integer

    D1 = Trim(RTLOG_RS("SAMCOLDATE"))
    t1 = Trim(RTLOG_RS("SAMCOLTIME"))
    If Trim(RTLOG_RS("ADPREPDATE")) <> "" Then
        D2 = Trim(RTLOG_RS("ADPREPDATE"))
        T2 = Trim(RTLOG_RS("ADPREPTIME"))
    End If

    If Len(t1) = 10 Then
        t1 = "0" & t1
    End If
    If Len(T2) = 10 Then
        T2 = "0" & T2
    End If

    H1 = Left(t1, 2)
    H2 = Left(T2, 2)
    m1 = Mid(t1, 4, 2)
    m2 = Mid(T2, 4, 2)
    

    If Right(t1, 2) = "PM" Then
        If H1 = 12 Then
        ' H1 = 0
        Else
         H1 = H1 + 12
        End If
    End If


    If Right(T2, 2) = "PM" Then
        If H2 = 12 Then
         '   H2 = 0
        Else
            H2 = H2 + 12
        End If
    End If

    'Commented by Shraddha on 07/01/09
    '********************************
'    If D2 <> "" And D1 <> "" Then
'        If DateDiff("D", D2, D1) = 0 Then
'            'H = H1 - H2
'            H = H2 - H1
'        ElseIf DateDiff("D", D2, D1) = 1 Then
'            H = 24 - H2 + H1
'        ElseIf DateDiff("D", D2, D1) > 1 Then
'            H = DateDiff("D", D2, D1) * 24
'            H = H + (24 - H2 + H1)
'        End If
'    End If

'***************************



'********* Added by Shraddha on 11/04/09 ****************
If D2 <> "" And D1 <> "" Then
    If DateDiff("D", D2, D1) = 0 Then
   
        If H2 = 0 Then
            If H2 = 0 Then
                h = H1 - H2
            End If
                 
            If H2 = 0 Then
                m = m1 - m2
            End If
        ElseIf m2 < 60 And m2 > 0 Then
            m2 = 60 - m2
            H2 = H2 + 1
           
            m = m2 + m1
            h = Val(H1) - Val(H2)
        Else
        
            If Val(H1) >= Val(H2) Then
                h = H1 - H2
            End If
            
            If m1 >= m2 Then
                m = m1 - m2
            End If
            
        End If
   End If
End If
'************************************************************

    If m > 60 Or m = 60 Then
        h = h + 1
        m = m - 60
    End If

    CLPTIME = h & ":" & m

    RTLOG_RS.Close
    RTLOG_QR.Close
    
    
Else
    CLPTIME = ""
End If

VWDMULPROCE_RS.Close
VWDMULPROCE_QR.Close

End Sub

Private Sub cmd_print_Click()
'On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = Trim(Replace(cmb_reqno.Text, "/", ""))    'MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
  '  strLink = "http://intranet.tmc.gov.in/EMRREPORTS/allreports.aspx?reqno=" & VAR1 & "~guest"
    'MRUNAL
    strLink = "http://intranet.tmc.gov.in/IMUNO_RPT_N/allreports.aspx?reqno=" & VAR1 & "~guest"
    'END
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If

Exit Sub
Dim n, d As String

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
On Error GoTo prtmsg:

DoEvents

Dim VAR_REQNNO As String

        VAR_REQNNO = req_chglen(cmb_reqno.Text)
        
        ''*******icu****************
'    If Opt_patient.Value = True Then
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
'        If Not SELMEMO_RS.EOF Then
'            If IsNull(SELMEMO_RS("MEMONO")) And Trim(SELMEMO_RS("PAT_CATG")) <> "NC" Then
'                MsgBox " Please Prepare the Memo !!!", vbInformation, "DIS"
'                Exit Sub
'            End If
'        End If
'    End If
'******************************* for displaying designation ************'
'********** added by milind on 03-03-06
userid_qr(0) = Trim(cmb_finalby.Text)
Set userid_rs = userid_qr.OpenResultset(1, 3)
n = userid_rs(0)
userdesc_qr(0) = n
Set userdesc_rs = userdesc_qr.OpenResultset(1, 3)
d = userdesc_rs(0)
'************upto here**********

       
        
        ''*******icu****************
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
             
             
'FOLLOWING LOOP ADDED BY SHWETA 07-11-2012
For VAR11 = 0 To 20 '..................................
CrystalReport1.Formulas(VAR11) = ""
Next VAR11
             '''''SUMOL START HERE
                    If reqspn_rs("NMETERIALS") = "Stone Analysis" Then  '
                    
                        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' "
                        'CrystalReport1.SelectionFormula = "{STONERPT.REQNNO} = '" & VAR_REQNNO & "' "
                        'and {STONERPT.CASE_NO} = '" & Trim(MaskEdBox1.FormattedText) & "'"
                        'CrystalReport1.SelectionFormula = "{DMULPROCE.REQNNO} = '" & VAR_REQNNO & "'"
                    '    CrystalReport1.Formulas(1) = "REQNNO = '" & "" & "'"
                            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
                    
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCSTONENEW.RPT"
                        GoTo display:
                    End If
            ''''ends here

             
             
             
             Dim pcount As Integer
             pcount = 0
            For I = 1 To MSFlexGrid1.Rows - 1
                If Trim(MSFlexGrid1.TextMatrix(I, 4)) <> "" Then
                  pcount = pcount + 1
                End If
            Next I
If Opt_patient.Value = True Then
             
         If SELDADDENDUMFLAG = False Then
            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
            SELDADDENDUMFLAG = True
          Else
            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            SELDADDENDUM_RS.MoreResults
            SELDADDENDUM_RS.Requery
          End If
   
        If SELDADDENDUM_RS.RowCount > 1 Then
            SELDADDENDUM_RS.MoveLast
        End If
        
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    
    'CrystalReport1.ReportFileName = "+"
    
    If SELDADDENDUM_RS.RowCount > 0 Then
    
    'CrystalReport1.SelectionFormula = "{DISBCUNIlf.REQNO} = '" & VAR_REQNNO & "' AND {DISBCUNIlf.ADDEND_NO}=" & SELDADDENDUM_RS("ADDEND_NO") & ""
    'Commented and added by Shraddha on 03/09/10
    'CrystalReport1.SelectionFormula = "{DISBCUNIlf.REQNO} = '" & VAR_REQNNO & "' AND {DISBCUNIlf.ADDEND_NO}=" & SELDADDENDUM_RS("ADDEND_NO") & " AND ({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F')"
    
    '({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F')
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCUNI.RPT"
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    Else
    'CrystalReport1.SelectionFormula = "{DISBCUNIlf.REQNO} = '" & VAR_REQNNO & "'"
    'Commented and added by Shraddha on 03/09/10
    CrystalReport1.SelectionFormula = "{DISBCUNIlf.REQNO} = '" & VAR_REQNNO & "' and ({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F')"
    CrystalReport1.ReportFileName = "C:\DISREPORTSEClf\DISBCUNI.RPT"
   
    '***** Shraddha 01/11/2011 *****
    SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
    Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
        If Not seldmulprocrs.EOF Then
            If Trim(seldmulprocrs!PRO_CODE) = "F836" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\BCMETHO.rpt"
                CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"

            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
           
            If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & Trim(cmb_enterby.Text)
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
    
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(17) = "DESG = '" & d & "'"
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            
            ''' ADDED BY KANCHAN
            CLPTM
            CrystalReport1.Formulas(20) = "COLLAPS TIME='" & CLPTIME & "'"
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
                Exit Sub
            
'********** Aditi 22/04/2013 ************************************************************
'********** Glucose Tolerance test report ***********************************************
            
            ElseIf Trim(seldmulprocrs!PRO_CODE) = "F844" Or Trim(seldmulprocrs!PRO_CODE) = "F810a" Then
                 CrystalReport1.SelectionFormula = "{DISBCGTLF.REQNO} = '" & VAR_REQNNO & "' "
                 CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISBCGTT.rpt"
                 CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"

                 If Trim(prepath) = "" Then
                     CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                 Else
                     CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
                 End If
                     
                     CrystalReport1.Formulas(1) = "NAME = '" & Trim(txt_name.Text) & "'"
                     CrystalReport1.Formulas(2) = "AGE = '" & Trim(txt_age.Text) & "'"
                     CrystalReport1.Formulas(3) = "GEN = '" & Trim(txt_sex.Text) & "'"
                     
                     
                 If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                    str1 = ""
                    str1 = "Entered by " & Trim(cmb_enterby.Text)
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                 Else
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                 End If
                    
                 CrystalReport1.Formulas(17) = "DESG = '" & d & "'"
                 CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            
                 ''' ADDED BY KANCHAN
                 'CLPTM
                 'CrystalReport1.Formulas(20) = "COLLAPS TIME='" & CLPTIME & "'"
                 CrystalReport1.WindowState = crptMaximized
                 CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                 CrystalReport1.Action = 1
                 Exit Sub
'****************************************************************************************
            End If
            
            
            
            
        End If
        
        
    End If

       ' CrystalReport1.ReportFileName = "C:\DISREPORTSEClf\DISBCUNI.RPT"
        
        ''*******icu****************
    If Opt_patient.Value = True Then
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
            If SELMEMO_FLAG = False Then
                SELMEMO_QR(0) = VAR_REQNNO
                Set SELMEMO_RS = SELMEMO_QR.OpenResultset(1, 3)
                SELMEMO_FLAG = True
            Else
                SELMEMO_QR(0) = VAR_REQNNO
                SELMEMO_RS.MoreResults
                SELMEMO_RS.Requery
            End If
        End If
        If Not SELMEMO_RS.EOF Then
            If IsNull(SELMEMO_RS("MEMONO")) And Trim(SELMEMO_RS("PAT_CATG")) <> "NC" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEClf\DISBCUNI.RPT"
'            Else
'                CrystalReport1.ReportFileName = "C:\DISREPORTSEClf\DISBCUNI.RPT"
            End If
        End If
    End If
        ''*******icu****************
        
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                
               
        
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
  
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If pcount <> 0 Then
        
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        End If
            CrystalReport1.Formulas(17) = "DESG = '" & d & "'"
            
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
             If Trim(SELDADDENDUM_RS("add_max_ln")) <> "" Then
              CrystalReport1.Formulas(14) = "Addendum ='Addendum'"
              CrystalReport1.Formulas(15) = "Addtime ='Adden time/dt :'"
             End If
          End If
End If
'**** STAFF
If opt_staff.Value = True Then
        
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

      
        CrystalReport1.ReportFileName = "C:\DISREPORTSEClf\DISBCSUNI.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                
               
        
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
  
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If pcount <> 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        End If

'********** Aditi 25/04/2013 ************************************************************
'********** Glucose Tolerance test report ***********************************************
For VAR11 = 0 To 20 '..................................
CrystalReport1.Formulas(VAR11) = ""
Next VAR11

    SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
    Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
            If Trim(seldmulprocrs!PRO_CODE) = "F844" Or Trim(seldmulprocrs!PRO_CODE) = "F810a" Then
                 CrystalReport1.SelectionFormula = "{DISBCGTLF.REQNO} = '" & VAR_REQNNO & "' "
                 CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISBCGTT.rpt"
                 CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"

                 If Trim(prepath) = "" Then
                     CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                 Else
                     CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
                 End If
                     CrystalReport1.Formulas(1) = "NAME = '" & Trim(txt_name.Text) & "'"
                     CrystalReport1.Formulas(2) = "AGE = '" & Trim(txt_age.Text) & "'"
                     CrystalReport1.Formulas(3) = "GEN = '" & Trim(txt_sex.Text) & "'"
                     
                 If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                    str1 = ""
                    str1 = "Entered by " & Trim(cmb_enterby.Text)
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                 Else
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                 End If
                    
                 CrystalReport1.Formulas(17) = "DESG = '" & d & "'"
                 CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            
                 ''' ADDED BY KANCHAN
                 'CLPTM
                 'CrystalReport1.Formulas(20) = "COLLAPS TIME='" & CLPTIME & "'"
                 CrystalReport1.WindowState = crptMaximized
                 CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                 CrystalReport1.Action = 1
                 Exit Sub
              End If
'****************************************************************************************
 


End If
'***** STAFF
display:
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1

'rtime log
UPDRTIMEP_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIMEP_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                
UPDRTIMEP_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIMEP_QR.Execute
'rtimelog

Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("print")
BTNENABL1 ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_refresh_Click()
If Trim(cmb_reqno.Text) <> "" Then
Dim UPDRREQ_DC_QR As rdoQuery
str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT='T',COMMITFLAG='' WHERE REQNNO=?"
Set UPDRREQ_DC_QR = rdb.CreateQuery("", str)
UPDRREQ_DC_QR(0) = Replace(cmb_reqno.Text, "/", "")
UPDRREQ_DC_QR.Execute
MsgBox "Record Updated Successfully"
'cmd_view_Click
viewflag = True
cmb_reqno_Click
End If
End Sub

Private Sub cmd_save_Click()
'If Trim(cmb_authSign.Text) = "" Then
'MsgBox ("Please select Authorised By")
'Exit Sub
'End If

Dim SELFLAG_QR As rdoQuery
Dim SELFLAG_RS As rdoResultset



'following check added by shweta 22-05-2018
For I = 1 To MSFlexGrid1.Rows - 1
  num = InStr(Trim(MSFlexGrid1.TextMatrix(I, 2)), "$")
  sc = Right(Trim(MSFlexGrid1.TextMatrix(I, 2)), 3)
  
   sqlstr = "SELECT * FROM DISLIB.FLAGMASTER WHERE UCASE(ACTIVE)='Y'"
         Set SELFLAG_QR = rdb.CreateQuery("", sqlstr)
         'SELFLAG_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
         Set SELFLAG_RS = SELFLAG_QR.OpenResultset(1, 3)
   If SELFLAG_RS.EOF = False Then
   While Not SELFLAG_RS.EOF
    If InStr(sc, SELFLAG_RS("FLAG")) = True Then
        MsgBox "Sample should be rerun so Report will not save", vbCritical
        Exit Sub
    End If
    SELFLAG_RS.MoveNext
   Wend
   End If
  
  
  
  
'  If Right(sc, 1) = "K" Or Right(sc, 1) = "J" Then
'
'  End If
Next I



Dim SELADDINTERFACE_QR As rdoQuery
Dim SELADDINTERFACE_RS As rdoResultset

 sqlstr = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO =?"
  Set SELADDINTERFACE_QR = rdb.CreateQuery("", sqlstr)
 SELADDINTERFACE_QR(0) = Replace(Trim(cmb_reqno.Text), "/", "")
Set SELADDINTERFACE_RS = SELADDINTERFACE_QR.OpenResultset(1, 3)
If SELADDINTERFACE_RS.EOF = False Then
    If SELADDINTERFACE_RS("reqnstat") = "T" Or SELADDINTERFACE_RS("reqnstat") = "F" Then
    MsgBox "Report is already Saved..."
    Exit Sub
    End If
End If




Dim pcount As Integer
Dim RCOUNT As Integer
Dim t1 As String
t1 = VAR_REQNNO
Dim STONE As String
STONE = PC


''''sumol nov 2006 starts here

If reqspn_rs("NMETERIALS") <> "Stone Analysis" Then
''ends here

        If Trim(cmb_finalby.Text) = "" Then
          MsgBox "Please Enter finalby", vbInformation, "DIS"
          cmb_finalby.SetFocus
          Exit Sub
        End If
            pcount = 0
            RCOUNT = 0
            For I = 1 To MSFlexGrid1.Rows - 1
                If Trim(MSFlexGrid1.TextMatrix(I, 4)) <> "" Then
                  pcount = pcount + 1
                End If
                
                If InStr(Trim(MSFlexGrid1.TextMatrix(I, 2)), "#") Then
                  RCOUNT = RCOUNT + 1
                End If
                
            Next I
            If pcount <> 0 Then
              j = MsgBox(pcount & " Panic Value, Do you want to Save? ", vbYesNo)
                If j = vbNo Then
                 Exit Sub
                End If
            End If
            
            'added by shweta 28-04-2015 as suggested by nitin sir
            
            If RCOUNT <> 0 Then
            If Len(Trim(TXT_COMMENTS.Text)) > 3 Then
            Else
               MsgBox ("Please Enter Remark")
               TXT_COMMENTS.SetFocus
               Exit Sub
                End If
            End If
            
            Dim str As String
            For I = 1 To MSFlexGrid1.Rows - 2
                If Trim(MSFlexGrid1.TextMatrix(I, 3)) = "" Then
                   str = "Range "
                End If
                
                If Trim(MSFlexGrid1.TextMatrix(I, 7)) = "" Then
                   str = str & " Machine "
                End If
            Next I
            
            If Len(Trim(str)) <> 0 Then
              If MsgBox(" Machine is Blank Do you want to save ", vbYesNo) = vbNo Then
               Exit Sub
              End If
            End If
'***** Aditi 18/04/2013 *****************************
'***** Grid validation - Insert all values **********


''For i = 1 To MSFlexGrid2.Rows - 1
''
''       If MSFlexGrid2.TextMatrix(i, 1) = "" Then
''          MsgBox "Please Enter All Values", vbInformation, "DIS"
''          Exit Sub
''
''       End If
''Next i

'****************************************************
            
              If adenum_flg = True Then
        If cmb_addendumby.Text = "" Then
        MsgBox "Please Select Addendum by Field", vbCritical
        cmb_addendumby.SetFocus
        Exit Sub
        End If
    End If
            
            
            
         assign
        'ASSIGNNEW

        
        
ElseIf reqspn_rs("NMETERIALS") = "Stone Analysis" Then
        assignstone
End If

'***** Comment later 20/07/10 *****
'***** Shraddha 20/07/10 *****
SELPATHDREQUISITN_QR(0) = t1
Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)

'INSDMPROCEZ_QR(0) = t1
'INSDMPROCEZ_QR(1) = Trim(selpathdrequisitn_rs("PATHNO"))
'INSDMPROCEZ_QR(2) = Trim(selpathdrequisitn_rs("REFDOC"))
'INSDMPROCEZ_QR.Execute
'**********************************



SELDTEMPLETFLAG = False
'rtimelog

DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1

UPDRTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
'

                
UPDRTIME_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIME_QR.Execute
BTNENABL1 ("SAVE")
' upto here

'TO UPDATE FLAG=3 IN DINTERFACE
'WRITTEN BY AMIT FOR INTERFACE TO UPDATE FLAG=3
For I = 1 To MSFlexGrid1.Rows - 1
  num = InStr(Trim(MSFlexGrid1.TextMatrix(I, 1)), "$")
  sc = Mid(Trim(MSFlexGrid1.TextMatrix(I, 1)), num + 1)
   UPDINTERFACE_QR(0) = 3
   VAR_REQNNO = req_chglen(cmb_reqno.Text)
   UPDINTERFACE_QR(1) = Trim(MSFlexGrid1.TextMatrix(I, 7))
   UPDINTERFACE_QR(2) = VAR_REQNNO
   UPDINTERFACE_QR(3) = sc
   ' commented as dinterface file is getting locked
  ' UPDINTERFACE_QR.Execute
Next I


' If flagsave = 1 Then
' insertIntoRAUTHOSIGN "I", 3, VAR_REQNNO, MaskEdBox1.FormattedText, cmb_enterby.Text, Winsock1.LocalIP, cmb_enterby.Text, Winsock1.LocalIP, cmb_authSign.Text, "", cmb_finalby.Text, "", "", ""
'
' ElseIf flagsave = 2 Then
' insertIntoRAUTHOSIGN "U", 2, VAR_REQNNO, MaskEdBox1.FormattedText, cmb_enterby.Text, Winsock1.LocalIP, cmb_enterby.Text, Winsock1.LocalIP, cmb_authSign.Text, Winsock1.LocalIP, cmb_finalby.Text, "", "", ""
' ElseIf flagsave = 3 Then
' insertIntoRAUTHOSIGN "U", 4, VAR_REQNNO, MaskEdBox1.FormattedText, cmb_enterby.Text, Winsock1.LocalIP, cmb_enterby.Text, Winsock1.LocalIP, cmb_authSign.Text, Winsock1.LocalIP, cmb_finalby.Text, "", cmb_addendumby.Text, Winsock1.LocalIP
' End If





End Sub

Public Sub cmd_view_Click()
On Error GoTo MSG:
'btnenabl ("VIEW")
BTNENABL1 ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR
MSFlexGrid1.clear
MSFlexGrid1.FormatString = "S.No.  |^                 Procedure Name                        |^    Observed Value   |^        Normal Range of the Procedure |^ Panic Value"
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
'    While Not sel_deptrs.EOF
'        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
'        sel_deptrs.MoveNext
'        Wend
'   End If
 
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
 
   
'cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
'cmb_dept.SetFocus
'cmb_modality.Enabled = True
MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True
Opt_patient.Value = True
MaskEdBox1.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
'On Error GoTo MSG:
'btnenabl ("INI")
dm = False
adenum_flg = False
lbl_addendnum.Visible = False
cmd_adenum.Visible = False

NEWFLAG = False
viewflag = False
cmb_dept.Text = "BIOCHEMISTRY$FEC"
cmb_modality.Text = "BIOCHEMISTRY$BC"
txt_deptcode = "FEC"
txt_modcode = "BC"
FLDENABL ("DIS")
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
 'prev reqns
 str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
  Set SELRREQUISITNpv_QR = rdb.CreateQuery("", str)
 
 
 
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

'SUMOL STARTS HERE
  str = "SELECT * FROM DISLIB.STONERPT WHERE REQNNO=? AND PRO_CODE=? "
  Set SELSTONEQR = rdb.CreateQuery("SELSTONEQR", str)
  
  str = "INSERT INTO DISLIB.STONERPT (REQNNO,PRO_CODE,LABNO,CASE_NO,MACID,TXT) VALUES (?,?,?,?,?,?)"
  Set INSSTONEQR = rdb.CreateQuery("INSSTONEQR", str)
  
  str = "UPDATE DISLIB.STONERPT SET REQNNO=?,PRO_CODE=?,TXT=?,MACID=? WHERE REQNNO=? AND PRO_CODE=? "
  Set UPDSTONEQR = rdb.CreateQuery("UPDSTONEQR", str)
  
'  Set ConInterface = renv.OpenConnection(dsname:="biochem", Prompt:=rdDriverNo)
'
'  sqlinterface = "select * from serial where reqno = ? and pro_code=? "
'  Set rqsqlinterface = ConInterface.CreateQuery("xxxx", sqlinterface)
'
'''FINISHES HERE

'********************displaying designation in crystal report*********'
'******************added by milind on 03-03-06
str = "select usr_id from dislib.dislnkusr where username=?"
Set userid_qr = rdb.CreateQuery("userid_qr1", str)

str = "select user_des from dislib.dsecurity where usr_id=?"
Set userdesc_qr = rdb.CreateQuery("userdesc_qr1", str)

'*******************************'
'************upto here*********************


'COMMENT
' comments removed for interfaceing 0n 30-12-2003
'Set ConInterface = renv.OpenConnection(dsname:="biochem", Prompt:=rdDriverNo)
' comments removed for interfaceing 0n 30-12-2003
If BCREPORTFLAG = False Then
    BCREPORTFLAG = True

'*********SECURITY**********************************
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? and DEPT_CODE=? "
'str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? and  USRTYPE >0 ORDER BY USERNAME"
Set USRNAME_QR = rdb.CreateQuery("", str)
USRNAME_QR(0) = UID
USRNAME_QR(1) = "FEC" 'removed comment by shweta 15-10-2018
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)

If USRNAME_RS.RowCount > 0 Then
'While Not USRNAME_RS.EOF
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = Trim(LOGINUSRNAME)
    cmb_enterby.Text = cmb_enterby.List(0)
    ' USRNAME_RS.MoveNext
 '   Wend
Else
    MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
    End
    cmd_cancel_Click
    Exit Sub
End If
'commented and added by shweta 19-02-2014 as per kalpana mam Instructions
'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='FEC' AND USRTYPE>=8 ORDER BY USERNAME" '
'AND USER_DES <>''
'modus condition added by shweta as per Chitra mam Mail.26-10-2018
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='FEC' AND USRTYPE>=8 AND USER_DES <>'' and modus='1' ORDER BY USERNAME" '
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNSTR", str)
'REPWTRN_QR(0) = Trim(USRDEPT)
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)

If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    'cmb_addendumby.clear
    cmb_finalby.clear
    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 0 Then
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        REPWTRNRS.MoveNext
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
cmb_finalby.Text = cmb_finalby.List(0)

Dim REPWTRN_QR_AS As rdoQuery
Dim REPWTRNRS_AS As rdoResultset
'commented and added by shweta 19-02-2014 as per kalpana mam Instructions
'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='FEC' AND USRTYPE>=8 ORDER BY USERNAME" '
'AND USER_DES <>''
'modus condition added by shweta as per Chitra mam Mail.26-10-2018
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='FEC' and modus='2' ORDER BY USERNAME" '
Set REPWTRN_QR_AS = rdb.CreateQuery("S", str)
'REPWTRN_QR(0) = Trim(USRDEPT)
Set REPWTRNRS_AS = REPWTRN_QR_AS.OpenResultset(1, 1)

If REPWTRNRS_AS.RowCount > 0 Then
    'cmb_enteryby.clear
    'cmb_addendumby.clear
    cmb_authSign.clear
    'cmb_provby.clear
    REPWTRNRS_AS.MoveFirst
    While Not REPWTRNRS_AS.EOF
        If REPWTRNRS_AS("USRTYPE") > 0 Then
           cmb_authSign.AddItem Trim(REPWTRNRS_AS("USERNAME"))
         
        End If
        REPWTRNRS_AS.MoveNext
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If



str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1BIO", STR)
'SELFINALBY_QR(0) = 6
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If

'*********SECURITY**********************************


''ADDED BY MONO SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC
'STR = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE=? AND SEX=?"


str = "SELECT * FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND SEX = ? " 'AM
Set SELDTEMPLETQR1 = rdb.CreateQuery("SELDTEMPLETQRBCFM1", str)

''********ICU***************
str = "SELECT ICDS2,spihreqno FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR11", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR11", str)

''********ICU***************

''for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr1BC", str)

'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)

'' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr1BC", str)

'    'for retriving reqnos
'Set req_qr = rdb.CreateQuery("reqBC", "select REQNNO from Rislib.RREQUISITN where CASENO = ? AND REQNSTAT = 'E' AND MODALITYCD = 'BC' ORDER BY REQNNO DESC ")
Set req_qr = rdb.CreateQuery("reqBC", "select REQNNO from Rislib.RREQUISITN where CASENO = ? AND REQNSTAT = 'E' AND MODALITYCD = 'BC' AND DEPT_CODE='FEC' ORDER BY REQNNO DESC ")
'

'ADDED BY MILIND DT. 23/01/06 FOR RETREIVING SPECIEMEN (NATURE OF MATERIALS)
str = "SELECT NMETERIALS,METASTASIS FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO=? "
Set reqspn_qr = rdb.CreateQuery("reqspn_BC", str)
'upto here

'for gtt by milind on 06-09-2006
str = "select pro_code from dislib.dmulproce where reqno=?"
Set gtt_qr = rdb.CreateQuery("gtt_qr", str)
'upto here

Set SELTA101_QR = rdb.CreateQuery("selta101_qr11BC", "select * from PATOOLSLIB.ta101 where case_no =?")

'   ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1BC", str)

'   'FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1BC", str)
''
'    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11BC", str)
'    selma104_qr(0) = "BN/00001"
'    Set selma104_rs = selma104_qr.OpenResultset(2, 3)
'    'for provisional diagnosis
'
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1BC", str)
'    selta1012_qr(0) = "BN/00001"
'    Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)

    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1111_QR1BC", str)

   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("", str)
'
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("", str)
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
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1BC", str)

    'commented and new query added by amarjeet on 05-11-2002 to select manjure's name in enter by
    str = "select * from dislib.dissec where (enterby='YES'AND MODALITY='BC') or (modality='EL' and ccname='Dr. S.S. Manjure')"
   ' STR = "select * from dislib.dissec where enterby='YES'AND MODALITY='BC'"
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1BC", str)
   
'   STR = "select * from dislib.dissec where PROVby='YES'AND MODALITY='SP'"
'   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1BC", STR)
'
   str = "select * from dislib.dissec where (FINALby='YES'AND MODALITY='BC')or (modality='EL' and ccname='Dr. S.S. Manjure')"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1BC", str)
    
  str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
  Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1BC", str)
    'INTERFACE BY AMIT
    
 ' str = "SELECT DISTINCT SNO,TESTCODE,TESTRESULT,MACHINENO,IFLAG FROM dislib.MERGEDATA1 WHERE REQNO=? AND (iflag IN ('2','1')) and  TESTCODE<>'F027'  ORDER BY IFLAG ,SNO"
  str = "SELECT DISTINCT SNO,TESTCODE,TESTRESULT,MACHINENO,IFLAG FROM INTERFACE.MERGEDATA1 WHERE REQNO=? AND (iflag IN ('2','1')) and  TESTCODE<>'F027'  ORDER BY IFLAG ,SNO"
  Set SELINTER_QR = rdb.CreateQuery("", str)
  
   str = "SELECT * FROM dislib.MERGEDATA1 WHERE REQNO=? AND iflag ='1' and  TESTCODE<>'F027' ORDER BY SNO"
  Set SELINTER_manual_QR = rdb.CreateQuery("", str)
  
  
  str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
  Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRBC", str)
    
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
  Set selMA101QR = rdb.CreateQuery("", str)
  
  str = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? "
  Set SELDTEMPLETQR = rdb.CreateQuery("SELDTEMPLETQRBC", str)
  
  str = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=?" 'AND MACHINNA=?"
  Set SELDTEMPLETR1 = rdb.CreateQuery("SELDTEMPLETMQR", str)
  
  str = "INSERT INTO DISLIB.DMPROCEDUR(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE,SOURCE,REMARKS) VALUES (?,?,?,?,?,?,?,?,?)"
  Set INSDMPROCEDURQR = rdb.CreateQuery("INSDMPROCEDURQRBC", str)
  
  '***** Shraddha 20/07/10 *****
  'str = "INSERT INTO DISLIB.DMPROCED_R(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE,SOURCE,REMARKS) VALUES (?,?,?,?,?,?,?,?,?)"
  str = "INSERT INTO DISLIB.DMPROCED_R(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE,SOURCE,REMARKS, ENTRYBY , pathno, CASENO, REQNSTAT, PATSTAT, PAT_CATG, REFDOC, ICSD2, PREPORTBY, SAMCOLDATE, SAMCOLTIME, SAMRECDATE, SAMRECTIME, COMREPDATE, COMREPTIME, Comments, MATERIAL, bodyfluid, wardname, BEDNAME, METHOD) VALUES (?,?,?,?,?,?,?,?,?)"

  'ENTRYBY, PATHNO, CASENO, REQNSTAT, PATSTAT, PAT_CATG, REFDOC, ICSD2, PREPORTBY, SAMCOLDATE, SAMCOLTIME, SAMRECDATE, SAMRECTIME, COMREPDATE, COMREPTIME, COMMENTS, MATERIAL, BODYFLUID, WARDNAME, BEDNAME, METHOD
  'REQNO, RO_CODE, OBSERVALUE, OBSER_TYPE, SOURCE, ENTRYBY, REMARKS, UOM, NRANGE, PVALUE, PATHNO, CASENO, REQNSTAT, PATSTAT, PAT_CATG, REFDOC, ICSD2, PREPORTBY, SAMCOLDATE, SAMCOLTIME, SAMRECDATE, SAMRECTIME, COMREPDATE, COMREPTIME, COMMENTS, MATERIAL, BODYFLUID, WARDNAME, BEDNAME, METHOD
  Set INSDMPROCEZ_QR = rdb.CreateQuery("INSDMPROCEZQR", str)
  '*****************************
  
  '********************* 23/04/2013 Aditi***************************************
  '********************* gtt details insertion/updation/view *******************
  str = "insert into dislib.GTDETAILS(REQNO,PRO_CODE,BLOOD,URINE,TIMESLOT,COMMENTS) values (?,?,?,?,?,?)"
  Set insgttdtlsqr = rdb.CreateQuery("", str)

  str = "UPDATE DISLIB.GTDETAILS SET PRO_CODE=?,BLOOD=?,URINE=? WHERE REQNO=? AND TIMESLOT=?"
  Set UPDGTTDTLSQR = rdb.CreateQuery("UPDGTTDTLSQR", str)
  
  str = "SELECT * FROM DISLIB.GTDETAILS WHERE REQNO=? AND TIMESLOT=?"
  Set SELGTTDTLSQR = rdb.CreateQuery("SELGTTDTLSQR", str)
  '*****************************************************************************
  
  str = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=? WHERE REQNO=? AND PRO_CODE=? "
  Set UPDDMPROCEDURQR = rdb.CreateQuery("UPDDMPROCEDURQRBC", str)
  
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_DET=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET=?"
  Set SELDMA101QR = rdb.CreateQuery("SELDMA101QRBC", str)
  
  str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
  Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1bc", str)
   
  'STR = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
  str = "SELECT PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE,SOURCE,REMARKS,SNO FROM DISLIB.DMPROCEDUR A LEFT OUTER JOIN INTERFACE.COAGPARM B ON A.PRO_CODE=B.PARAMCODE AND A.REMARKS=B.MACHINEID WHERE REQNO=? ORDER BY B.SNO"
  Set SELDPROCEDURQR = rdb.CreateQuery("SELDMPROCEDURQRBC", str)
  
  str = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
  Set SELDBCPROSUBQR = rdb.CreateQuery("SELDBCPROSUBQRBC", str)
  
  str = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES(?,?)"
  Set INSDMBCCOMMQR = rdb.CreateQuery("INSDMBCCOMMQRBC", str)
  
  str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
  Set UPDDMBCCOMMQR = rdb.CreateQuery("UPDDMBCCOMMQRBC", str)
  
  str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
  Set SELDMBCCOMMQR = rdb.CreateQuery("SELDMBCCOMMQRBC", str)
  
  str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrBC1", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%BC%'"
  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRBC1", str)
  
  str = "SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE=? AND AGE2=? AND SEX=? "
  Set SELDBCALKPHSQR = rdb.CreateQuery("SELDBCALKPHSQRBC", str)
  
 'STR = "SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE=? AND (AGE1<? AND AGE2>?) "
  str = "SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE=? AND (AGE2=?) AND SEX=? "
  'STR = "SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE=? AND AGE1<=?  AND SEX=? "
  Set SELDBCALKPHSGGTQR = rdb.CreateQuery("SELDBCALKPHSQRBCGGT", str)
  
  'MACHINE
  str = "SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE=? AND AGE2=? AND SEX=? AND MACHINNA=?"
  Set SELDBCALKPHS_MQR = rdb.CreateQuery("SELDBCALKPHS_MQRBC", str)

  str = "SELECT * FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND SEX = ? AND MACHINNA=?"  'AM
  Set SELDTEMPLET_MQR1 = rdb.CreateQuery("SELDTEMPLETMQRBCFM1", str)

  str = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=?" 'AND MACHINNA=?"
  Set SELDTEMPLET_MQR = rdb.CreateQuery("SELDTEMPLETMQRBC", str)

''  STR = "SELECT * FROM DISLIB.DBCALKPHS WHERE PROCODE=? AND AGE2=? AND SEX=? "
''  Set SELDBCALKPHS_MQR = rdb.CreateQuery("SELDBCALKPHS_MQRBC", STR)
''
''  STR = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE=? AND SEX = ? " 'AM
''  Set SELDTEMPLET_MQR1 = rdb.CreateQuery("SELDTEMPLETMQRBCFM1", STR)
''
''  STR = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET WHERE PRO_CODE=? "
''  Set SELDTEMPLET_MQR = rdb.CreateQuery("SELDTEMPLETMQRBC", STR)
  
' INTERFACE
'**************** Commented and added by Shraddha Pandit 29/06/09 *****************
'str = "SELECT * FROM interface.B_INTERFAC WHERE REQNO = ? AND TESTCODE = ? AND IFLAG = '2'"
'str = "SELECT * FROM interface.B_INTERFAC WHERE REQNO = ? AND TESTCODE = ? AND (IFLAG = '2' OR IFLAG = '1')"
str = "SELECT * FROM DISLIB.mergedata1 WHERE REQNO = ? AND TESTCODE = ? AND (IFLAG = '2' OR IFLAG = '1') ORDER BY TESTCODE"
Set INTERFACEREAD_QR = rdb.CreateQuery("INTERFACEDATA1", str)
'**********************************************************************************

'ADDED BY MILIND FOR OLYMPUS MACHINE TO FETCH DATA FROM MS-ACCESS DINTERFACE TABLE
'Set ConInterface = renv.OpenConnection(dsname:="biochem", Prompt:=rdDriverNo)
str = "SELECT * FROM DISLIB.BCGRIDFIL WHERE REQNO=? AND TESTCODE=? AND IFLAG='2' "
'Set SELMSADINTF_QR = ConInterface.CreateQuery("SELMSADINTF_QR", STR)
''UPTO HERE

str = "UPDATE INTERFACE.DINTERFACE SET IFLAG = ? WHERE REQNO = ?"
Set INTERFACEUPD_QR = rdb.CreateQuery("INTERFACEUPD1", str)
str = "UPDATE INTERFACE.DINTERFACE SET IFLAG=?,MACHINENO=? WHERE REQNO=? AND TESTCODE=?"
Set UPDINTERFACE_QR = rdb.CreateQuery("", str)
'STR = "SELECT * FROM interface.B_INTERFAC WHERE REQNO=?"
'Set SELLABNO_QR = rdb.CreateQuery("LABNO", STR)
'INTERFACE ENDS
  
 str = "INSERT INTO DISLIB.ADDENDUMCY(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
 Set INSDADDENDUM_QR = rdb.CreateQuery("", str)
   
 str = "SELECT * FROM DISLIB.ADDENDUMCY WHERE REQNO=?"
 Set SELDADDENDUM_QR = rdb.CreateQuery("", str)
  'COMMENT
  ' comments removed for interfaceing 0n 30-12-2003
  'sqlinterface = "select * from serial where reqno = ? and pro_code=?"
 ' Set rqsqlinterface = ConInterface.CreateQuery("xxxx", sqlinterface)
' comments removed for interfaceing 0n 30-12-2003
End If
MSFlexGrid1.col = 2
txt_observed.Width = MSFlexGrid1.CellWidth
txt_observed.Height = MSFlexGrid1.CellHeight
txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
BTNENABL1 ("INI")

Exit Sub
'MSG:
'MsgBox Err.Description
End Sub


Public Sub MaskEdBox1_LostFocus()
If MSFlexGrid2.Visible = True Then
Exit Sub
End If
Dim CASENO As String
'On Error GoTo MSG:
If Trim(MaskEdBox1.Text) = "" Then
 Exit Sub
End If
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
      ''  MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))   ''02/12/2022
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
            
            'added  by shweta 01-03-2018
            Dim SELADDINTERFACE_QR As rdoQuery
            Dim SELADDINTERFACE_RS As rdoResultset

                sqlstr = "SELECT REQNO FROM DISLIB.DREQUISITN WHERE PATHNO=?"
                Set SELADDINTERFACE_QR = rdb.CreateQuery("", sqlstr)
                SELADDINTERFACE_QR(0) = Trim(txt_pathno.Text)
                Set SELADDINTERFACE_RS = SELADDINTERFACE_QR.OpenResultset(1, 3)
                If SELADDINTERFACE_RS.EOF = False Then
                
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
        
        If viewflag = True Or ADDENDUMFALG = True Then
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
                'btnenabl ("INI")
                BTNENABL1 ("INI")
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
                'btnenabl ("INI")
                BTNENABL1 ("INI")
                Exit Sub
            End If
        
        End If
        On Error GoTo MSG
                cmb_reqno.SetFocus
End If

If NEWFLAG = True Then
If cmb_reqno.ListCount > 0 Then
cmb_reqno.Text = cmb_reqno.List(0)



cmb_reqno_Click
End If
End If
cmb_reqno.SetFocus


'cmb_reqno.SetFocus
 Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub MSFlexGrid1_Click()
'If MSFlexGrid1.Col = 2 Then
 '       txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
  '      txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'End If
R1 = MSFlexGrid1.row
C1 = MSFlexGrid1.col
If C1 = 2 Then
   
   ' If MSFlexGrid1.TextMatrix(R1, C1) = "" Then
     txt_observed.Text = MSFlexGrid1.TextMatrix(R1, C1)
    'End If
    txt_observed.Visible = True
      '  txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
       ' txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    txt_observed.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
    txt_observed.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
    txt_observed.Width = MSFlexGrid1.CellWidth
    txt_observed.Height = MSFlexGrid1.CellHeight
    txt_observed.SelStart = 0
    txt_observed.SelLength = Len(txt_observed.Text)
    txt_observed.SetFocus
 
End If

If C1 = 7 Then
   ' If MSFlexGrid1.TextMatrix(R1, C1) = "" Then
        cmb_machine.Text = MSFlexGrid1.TextMatrix(R1, C1)
    'End If
    cmb_machine.Visible = True
      '  cmb_machine = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
       ' cmb_machine.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    cmb_machine.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
    cmb_machine.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
    cmb_machine.Width = MSFlexGrid1.CellWidth
    'cmb_machine.Height = MSFlexGrid1.CellHeight
    cmb_machine.SelStart = 0
    cmb_machine.SelLength = Len(cmb_machine.Text)
'    cmb_machine.SetFocus
End If
End Sub

Private Sub MSFlexGrid1_GotFocus()
'MSFlexGrid1.Col = 2
'txt_observed.Width = MSFlexGrid1.CellWidth
'txt_observed.Height = MSFlexGrid1.CellHeight
'txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
End Sub

Private Sub MSFlexGrid2_Click()
R1 = MSFlexGrid2.row
C1 = MSFlexGrid2.col
If C1 = 1 Then
   
    txt_gttobsvd.Text = MSFlexGrid2.TextMatrix(R1, C1)
    txt_gttobsvd.Visible = True
    txt_gttobsvd.Left = MSFlexGrid2.Left + MSFlexGrid2.CellLeft
    txt_gttobsvd.Top = MSFlexGrid2.Top + MSFlexGrid2.CellTop
    txt_gttobsvd.Width = MSFlexGrid2.CellWidth
    txt_gttobsvd.Height = MSFlexGrid2.CellHeight
    'txt_gttobsvd.SelStart = 0
    'txt_gttobsvd.SelLength = Len(txt_gttobsvd.Text)
    txt_gttobsvd.SetFocus
 
End If

If C1 = 2 Then
   
    cmb_Gtt_UrineValue.Text = MSFlexGrid2.TextMatrix(R1, C1)
    cmb_Gtt_UrineValue.Visible = True
    cmb_Gtt_UrineValue.Left = MSFlexGrid2.Left + MSFlexGrid2.CellLeft
    cmb_Gtt_UrineValue.Top = MSFlexGrid2.Top + MSFlexGrid2.CellTop
    cmb_Gtt_UrineValue.Width = MSFlexGrid2.CellWidth
'    cmb_Gtt_UrineValue.Height = MSFlexGrid2.CellHeight
    'cmb_Gtt_UrineValue.SelStart = 0
   ' cmb_Gtt_UrineValue.SelLength = Len(txt_gttobsvd.Text)
    cmb_Gtt_UrineValue.SetFocus
 
End If
'If C1 = 7 Then
'   ' If MSFlexGrid1.TextMatrix(R1, C1) = "" Then
'        cmb_machine.Text = MSFlexGrid1.TextMatrix(R1, C1)
'    'End If
'    cmb_machine.Visible = True
'      '  cmb_machine = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'       ' cmb_machine.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'    cmb_machine.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
'    cmb_machine.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
'    cmb_machine.Width = MSFlexGrid1.CellWidth
'    'cmb_machine.Height = MSFlexGrid1.CellHeight
'    cmb_machine.SelStart = 0
'    cmb_machine.SelLength = Len(cmb_machine.Text)
''    cmb_machine.SetFocus
'End If

End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
   ' MaskEdBox1.Mask = "??>/#####"
    MaskEdBox1.Enabled = True
    'AMIT
    'MaskEdBox1.SetFocus
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then

    lbl_caseno.Caption = "  CC No."
    MaskEdBox1.Mask = "########"
    MaskEdBox1.Enabled = True
    'AMIT
    'MaskEdBox1.SetFocus
End If
End Sub
'ADDED by AMIT

Function WaitForProcess(taskID As Long, Optional msecs As Long = -1) As Boolean
'Dim procHandle As Long
'procHandle = OpenProcess(&H100000, True, taskID)
'WaitForProcess = WaitForSingleObject(procHandle, msecs) <> -1
'CloseHandle procHandle
End Function
Private Sub PRINTBARCODE_Click()
''Dim QTY1 As String
''Dim BAR_CD As String
''Dim EXP_DT As String
''Dim ITEM_DESC As String
''    Dim TEMPDATA As String
''    Dim rack As String
''    Dim batchno1 As String
''    Dim itemcode1 As String
''    Dim i As Long, START
''    Dim J As Integer
''    Dim NOL As Integer
''    Dim Nos As Integer
''    'Dim QTY1 As Integer
''    Nos = 0
''    J = 1
''     f1 = FreeFile()
''     Open "c:\TATA4.TXT" For Output As #f1   '
'''    For j = 1 To MSFlexGrid1.Row
''        'If Trim(MSFlexGrid1.TextMatrix(j, 1)) <> "" Then
''         '   For NOL = 1 To Trim(MSF_Barcd_qty.TextMatrix(j, 4))
''                SELLABNO_QR(0) = req_chglen(cmb_reqno.Text)
''                SELLABNO_QR.Execute
''                QTY = 1
''                Nos = 0
''                Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
''                Do Until SELLABNO_RS.EOF
''
''                    'itemcode1 = MSFlexGrid1.TextMatrix(j, 1)
''                    'ITEM_DESC = MSFlexGrid1.TextMatrix(j, 2)
''                    'batchno1 = MSFlexGrid1.TextMatrix(j, 3)
''                    'BAR_CD = itemcode1 & "$" & batchno1   ' MSF_Barcd_qty.TextMatrix(j, 1) & "$" & MSF_Barcd_qty.TextMatrix(j, 3)
''                    'EXP_DT = MSFlexGrid1.TextMatrix(j, 5)
''                    'QTY1 = Format(MSF_Barcd_qty.TextMatrix(j, 4), "0000")
''                    BAR_CD = Trim(SELLABNO_RS!LABNO)
''                    ITEM_DESC = Trim(SELLABNO_RS!CASENO)
''                    QTY1 = "Q00" & QTY
''                    EXP_DT = "EXP" & Trim(SELLABNO_RS!TDATE) & Trim(SELLABNO_RS!TTIME)
''                    Call NEWPRINTLBL(BAR_CD, EXP_DT, ITEM_DESC, QTY1)
''                    QTY = QTY + 1
''                    Nos = Nos + QTY
''                    SELLABNO_RS.MoveNext
''                Loop
''
''       ' End If
''
'' '   Next j
''    Close #f1
''  WaitForProcess Shell("c:\tata11.bat", 0)
''    MsgBox "Total " & Nos & " Labels Have Been Printed", vbInformation

End Sub
'ADDED BY AMIT
Public Sub NEWPRINTLBL(BAR_CD1 As String, EXP_DT1 As String, ITEM_DESC1 As String, QTY11 As String)
'Dim BARCD As String
'Dim EXPDT As String
'Dim BARPR As String
'Dim BARTXT As String
'Dim QTYPR11 As String
'QTYPR11 = "Q" + Trim(QTY11)
'
'BARCD = BAR_CD1
'EXPDT = EXP_DT1
'BARPR = "191100300050015" + Trim(EXPDT)
'BARTXT = "1E1102000200015 " + Trim(BARCD)
'f1 = 1
'Print #f1, Chr(2) & "L"
'Print #f1, "D11"
'Print #f1, "H10"
'Print #f1, "PC"
'Print #f1, "SC"
'Print #f1, "C0000"
'Print #f1, "R0000"
'Print #f1, BARPR  '   "191100300050015&EXPDT"
'Print #f1, BARTXT   ' "IE1102000200015  bar_cd1"
'Print #f1, QTYPR11
'Print #f1, "E"
End Sub

Private Sub Timer1_Timer()
    'txt_time.Text = Format(Time, "HH:MM:SS")
    '***** Commented and added on 10/06/09 by Shraddha *****
'    datetime
'    txt_time.Text = time1
    '*******************************************************
End Sub

Private Sub TXT_COMMENTS_GotFocus()
'If Trim(txt_comments.Text) = "" Then
    TXT_COMMENTS.clear
    
    TXT_COMMENTS.AddItem "# Rechecked"
    TXT_COMMENTS.AddItem "# Rechecked. Sample Query.Please Repeat the Sample"
    TXT_COMMENTS.AddItem "#No sample, ## Rechecked"
    
    TXT_COMMENTS.AddItem "# No Sample"
    TXT_COMMENTS.AddItem "# Serum haemolysed, Please repeat the sample"
    TXT_COMMENTS.AddItem "# Serum highly haemolysed, Please repeat the sample"
    TXT_COMMENTS.AddItem "# Serum lipaemic, Please repeat the sample"
    TXT_COMMENTS.AddItem "# Serum highly lipaemic, Please repeat the sample"
    TXT_COMMENTS.AddItem "24 Hrs Urine Collection from 8.00am/dt:          to 8.00am/dt:         ."
    TXT_COMMENTS.AddItem "Second sample Received at "
    TXT_COMMENTS.AddItem "# Sample ? Pl. repeat the sample."
    TXT_COMMENTS.AddItem "# Machine not working"
    TXT_COMMENTS.AddItem "# Out of Range (High/Low)sample ?,Please repeat the sample"
    TXT_COMMENTS.AddItem "# Sample ? Diluated with IV fluid ? Please repeat the sample."
    TXT_COMMENTS.AddItem "# Serum repeatedly clotting, Please repeat the sample"
    TXT_COMMENTS.AddItem "# Value too low cannot be detected"
    TXT_COMMENTS.AddItem "# Rechecked sample ? Pl.repeat the sample"
    TXT_COMMENTS.AddItem "# Rechecked sample ? Diluted with Iv fluid?"
    TXT_COMMENTS.AddItem "sample received on                          "

    TXT_COMMENTS.AddItem "# Clotted sample not received"
    TXT_COMMENTS.AddItem "# The test will be reported later"
    TXT_COMMENTS.AddItem "# The Final test Result will be reported later"
    TXT_COMMENTS.AddItem "# Sample collected in wrong container"
'End If
End Sub



Private Sub txt_gttobsvd_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo MSG


If KeyCode = 13 Then
    MSFlexGrid2.TextMatrix(MSFlexGrid2.row, MSFlexGrid2.col) = txt_gttobsvd.Text
    txt_gttobsvd.Text = ""
    txt_gttobsvd.Visible = False
    
End If

If KeyCode = 39 Then
    
    If txt_gttobsvd.Text <> "" Then
        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, MSFlexGrid2.col) = txt_gttobsvd.Text
        txt_gttobsvd.Text = ""
        
    End If
 
    MSFlexGrid2.col = MSFlexGrid2.col + 1
    txt_gttobsvd.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left

End If

If KeyCode = 40 Then
    
    If MSFlexGrid2.row = MSFlexGrid2.Rows - 1 Then
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
    End If
    
    If txt_gttobsvd.Text <> "" Then
        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, MSFlexGrid2.col) = txt_gttobsvd.Text
        txt_gttobsvd.Text = ""
    End If
 
    MSFlexGrid2.row = MSFlexGrid2.row + 1
    txt_gttobsvd.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top

End If

If KeyCode = 37 Then
 
    If txt_gttobsvd.Text <> "" Then
        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, MSFlexGrid2.col) = txt_gttobsvd.Text
        txt_gttobsvd.Text = ""
    End If
 
    MSFlexGrid2.col = MSFlexGrid2.col - 1
    
    If MSFlexGrid2.col <> 0 Then
        txt_gttobsvd.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left
    End If

End If

If KeyCode = 38 Then

    If txt_gttobsvd.Text <> "" Then
        MSFlexGrid2.TextMatrix(MSFlexGrid2.row, MSFlexGrid2.col) = txt_gttobsvd.Text
        'Text2.Text = ""
    End If
 
    MSFlexGrid2.row = MSFlexGrid2.row - 1
    txt_gttobsvd.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top

End If

MSG:
Exit Sub

End Sub

Private Sub txt_gttobsvd_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then
'
'    i = MSFlexGrid2.Rows
'    J = MSFlexGrid2.Cols
'    MSFlexGrid2.TextMatrix(i, J) = txt_gttobsvd.Text
'
'
'    If i = MSFlexGrid2.row Then
'        cmd_save.SetFocus
'        Exit Sub
'    End If
'
'    If MSFlexGrid2.row = MSFlexGrid2.Rows - 1 Then
'            MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
'            MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 0) = MSFlexGrid2.Rows - 1
'    End If
'
'    If MSFlexGrid2.row < MSFlexGrid2.Rows - 1 Then
'
'            MSFlexGrid2.row = MSFlexGrid2.row + 1
'            txt_gttobsvd.Text = CStr(MSFlexGrid2.TextMatrix(MSFlexGrid2.row, 2))
'
'            txt_gttobsvd.Visible = True
'            txt_gttobsvd.SetFocus
'
'            txt_gttobsvd.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top
'
'    End If
'End If

End Sub


Private Sub txt_observed_KeyPress(KeyAscii As Integer)

'MsgBox Trim(SELDMA101RS("ser_code"))
Dim NU As Integer
Dim num As Integer
Dim sc As String
Dim low, high As Variant

'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Trim(cmb_machine.Text)
num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)
If KeyAscii = 13 Then
    
    SELDTEMPLETR1(0) = sc
    Set SELDTEMPLETS1 = SELDTEMPLETR1.OpenResultset(1, 3)

    If Trim(sc) = "F866" Then
        PROTEIN = Val(Trim(txt_observed.Text))
    End If

    If Trim(sc) = "F867" Then
        GLOBULIN = Val(Trim(txt_observed.Text))
        If Trim(txt_observed.Text) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 2) <> "" Then
            Dim SS As Variant
            Dim v As Integer
            SS = Val(PROTEIN) - Val(GLOBULIN)
            v = InStr(SS, ".")
            If v = 0 Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.row + 1, 2) = SS & ".0"
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.row + 1, 2) = Val(PROTEIN) - Val(GLOBULIN)
            End If
        End If
   End If
   On Error Resume Next
   
   '------------- For F829 Shraddha Pandit 18/12/08 ---------------------------
    If Trim(sc) = "F829" Then
        If Val(Trim(txt_observed.Text)) <= 1 Then 'Normal
            MSFlexGrid1.CellBackColor = vbWhite
        Else 'High
            MSFlexGrid1.CellBackColor = vbRed
        End If
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
   Else
       '------------------------------------------------------------------
       
        NU = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), "-")
        low = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 1, NU - 1)
        high = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), NU + 1)
      
        'JUSTIN 05-10-2004
        If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)) <> "" Then
            'If ((low = 0 And high = 0) Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)))) = "ABSENT") Then
            'MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = vbWhite
            'End If
            
            
            
            NU = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGr0id1.row, 3)), "-")
            If Val(Trim(txt_observed.Text)) > Val(Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), NU + 1)) Then
                'If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                'Milind dt 19/01/06
                If ((low = 0 And high = 0) Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)))) = "ABSENT") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "PRESENT+") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "PRESENT++") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "PRESENT+++") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "PRESENT++++") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "+") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "++") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "+++") Or (UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))) = "++++") Then
                    dm = True
                    MSFlexGrid1.CellBackColor = vbWhite
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                End If
                'upto here
                If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETS1("HighPANIC"))) Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                    'Milind dt 19/01/06
                    If Val(Trim(SELDTEMPLETRS("HighPANIC"))) = 0 And Val(Trim(SELDTEMPLETS1("LowPANIC"))) = 0 Then
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                    End If
                    'upto here
                Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                End If
                For I = 1 To 2
                    Beep
                Next I
            Else
                If Val(Trim(txt_observed.Text)) >= Val(Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 1, NU - 1)) Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                    MSFlexGrid1.CellBackColor = vbWhite
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                    If Val(Trim(SELDTEMPLETS1("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                    Else
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                    End If
                Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = ""
                End If
                If Val(Trim(txt_observed.Text)) < Val(Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 1, NU - 1)) Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                    MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                    If Val(Trim(SELDTEMPLETS1("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                    Else
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                    End If
                End If
                If Trim(txt_observed.Text) = "" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                    MSFlexGrid1.CellBackColor = RGB(255, 255, 255)
                End If
            End If
            
            '**** Shraddha 12/11/2011 *****
            If Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 1) = "<" Then
                low = ""
                high = Right(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 3)
                If Val(Trim(txt_observed.Text)) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                    MSFlexGrid1.CellBackColor = vbGreen
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                ElseIf Val(Trim(txt_observed.Text)) > high Then
                    MSFlexGrid1.CellBackColor = vbRed
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                Else
                    MSFlexGrid1.CellBackColor = RGB(255, 255, 255)
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                End If
                
            End If
            '******************************
             If Val(Trim(txt_observed.Text)) < low And low <> "" Then
                    MSFlexGrid1.CellBackColor = vbGreen
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
              End If
              If Val(Trim(txt_observed.Text)) > Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), NU + 1) Then  'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                    MSFlexGrid1.CellBackColor = vbRed
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
              End If
        End If
    
        'up to here justine
        
        
        
        '''If KeyAscii = 13 Then
        '''    If SELDMA101FLAG = False Then
        '''        SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
        '''        Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
        '''        SELDMA101FLAG = True
        '''    Else
        '''        SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
        '''        SELDMA101RS.MoreResults
        '''        SELDMA101RS.Requery
        '''    End If
        '''    If SELDMA101RS.RowCount > 0 Then
        ''''        **** VALIDATION FOR PROTEIN-ALBUMIN
        '''
        '''        If Trim(SELDMA101RS("ser_code")) = "F866" Then
        '''          PROTEIN = Val(Trim(txt_observed.Text))
        '''        End If
        '''
        '''
        '''
        '''        If Trim(SELDMA101RS("ser_code")) = "F867" Then
        '''          GLOBULIN = Val(Trim(txt_observed.Text))
        '''          If Trim(txt_observed.Text) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 2) <> "" Then
        '''            Dim SS As Variant
        '''            Dim v As Integer
        '''            SS = Val(PROTEIN) - Val(GLOBULIN)
        '''            v = InStr(SS, ".")
        '''            If v = 0 Then
        '''                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + 1, 2) = SS & ".0"
        '''            Else
        '''                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + 1, 2) = Val(PROTEIN) - Val(GLOBULIN)
        '''            End If
        '''          End If
        '''        End If
        ''''        '**** VALIDATION FOR PROTEIN-ALBUMIN
        '''
        '''        If SELDTEMPLETFLAG = False Then
        '''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
        '''            'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
        '''            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
        '''            SELDTEMPLETFLAG = True
        '''        Else
        '''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
        '''            'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
        '''            SELDTEMPLETRS.MoreResults
        '''            SELDTEMPLETRS.Requery
        '''        End If
        '''
        '''
        '''        If SELDTEMPLETRS.RowCount > 0 Then
        '''                '******
        '''                If Trim(SELDMA101RS("SER_CODE")) = "F869" Then
        '''
        '''                        If Val(Trim(txt_observed.Text)) > Val(ALKHV) Then
        '''                            MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
        '''                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                            If Val(Trim(txt_observed.Text)) > 500 Then
        '''                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
        '''                            Else
        '''                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
        '''                            End If
        '''                            For i = 1 To 1000
        '''                            Beep
        '''                            Next i
        '''                        Else
        '''                            If Val(Trim(txt_observed.Text)) >= Val(ALKLV) Then
        '''                               MSFlexGrid1.CellBackColor = vbWhite
        '''                               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                               If Val(ALKLV) > Val(Trim(txt_observed.Text)) Then
        '''                                   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
        '''                               Else
        '''                                   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
        '''                               End If
        '''                            Else
        '''                               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
        '''                            End If
        '''                            If Val(Trim(txt_observed.Text)) < Val(ALKLV) Then
        '''                               MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
        '''                               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                                   If Val(ALKLV) > Val(Trim(txt_observed.Text)) Then
        '''                                       MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
        '''                                   Else
        '''                                       MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
        '''                                   End If
        '''
        '''                            End If
        '''                            If Trim(txt_observed.Text) = "" Then
        '''                               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                               MSFlexGrid1.CellBackColor = RGB(255, 255, 255)
        '''
        '''                            End If
        '''                End If
        '''
        '''
        '''
        '''
        '''
        '''                '******
        '''            Else
        '''                'JUSTIN 05-10-2004
        '''                NU = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), "-")
        '''                If Val(Trim(txt_observed.Text)) > Val(Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), NU + 1)) Then
        '''
        '''                'If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
        '''                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
        '''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                    If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC"))) Then
        '''                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
        '''                    Else
        '''                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
        '''                    End If
        '''                    For i = 1 To 1000
        '''                    Beep
        '''                    Next i
        '''                Else
        '''
        '''                        If Val(Trim(txt_observed.Text)) >= Val(Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), 1, NU - 1)) Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
        '''                           MSFlexGrid1.CellBackColor = vbWhite
        '''                           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                           If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
        '''                               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
        '''                           Else
        '''                               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
        '''                           End If
        '''                        Else
        '''                           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
        '''                        End If
        '''
        '''                        If Val(Trim(txt_observed.Text)) < Val(Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), 1, NU - 1)) Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
        '''                           MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
        '''                           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                               If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
        '''                                   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
        '''                               Else
        '''                                   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
        '''                               End If
        '''
        '''                        End If
        '''                        If Trim(txt_observed.Text) = "" Then
        '''                           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
        '''                           MSFlexGrid1.CellBackColor = RGB(255, 255, 255)
        '''
        '''                        End If
        '''                End If
        '''             End If
        '''        'End If
        '''    End If
            
        'justine make the comments
        
        
        I = MSFlexGrid1.Rows
    End If
    If I = MSFlexGrid1.row Then
        cmd_save.SetFocus
        Exit Sub
    End If
    If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
    End If
    If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        txt_observed.Text = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
        txt_observed.Visible = True
        txt_observed.SetFocus
        txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    End If
    
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

'On Error GoTo MSG123: 'shweta 07-NOV-2012

rdb.BeginTrans
 
 Dim NU As Integer
 Dim num As Integer
 Dim sc As String
 Dim gttcode As String
 
Dim count, low, high, W1
 
    If flagsave = 1 Then 'Insert
        'ADDED BY MILIND ON 12-09-2006 FOR GTT TEST
        If gttflag = True Then
            gttcode = ""
            gtt_qr(0) = VAR_REQNNO
            Set gtt_rs = gtt_qr.OpenResultset(1, 1)
            If gtt_rs!PRO_CODE = "F844" Or gtt_rs!PRO_CODE = "F810a" Then
                gttcode = Trim(gtt_rs!PRO_CODE)
                gtt_rs.Close
                gttflag = False
        '******* 23/04/2013 Aditi*********************************************
        '******* Insert GTT details ******************************************

                For I = 1 To MSFlexGrid2.Rows - 1
                    insgttdtlsqr(0) = VAR_REQNNO
                    insgttdtlsqr(1) = gttcode
                    j = 2
                    
                    If j <= 11 Then
                        If MSFlexGrid2.TextMatrix(I, 1) <> "" Then
                            insgttdtlsqr(j) = MSFlexGrid2.TextMatrix(I, 1)
                            j = j + 1
                        Else
                            insgttdtlsqr(j) = Null
                            j = j + 1
                        End If
                        If MSFlexGrid2.TextMatrix(I, 2) <> "" Then
                            insgttdtlsqr(j) = MSFlexGrid2.TextMatrix(I, 2)
                            j = j + 1
                        Else
                            insgttdtlsqr(j) = MSFlexGrid2.TextMatrix(I, 2)
                            j = j + 1
                        End If
                        
                        If MSFlexGrid2.TextMatrix(I, 0) = "1 . FASTING" Then
                            insgttdtlsqr(4) = 0
                            j = j + 1
                        ElseIf Trim(MSFlexGrid2.TextMatrix(I, 0)) = "2 . 2 hrs." Then
                            insgttdtlsqr(4) = "2"
                            j = j + 1
                        Else
                            insgttdtlsqr(4) = Mid(MSFlexGrid2.TextMatrix(I, 0), 5, 3)
                            j = j + 1
                        End If
                    End If

                    insgttdtlsqr(5) = ""
                    insgttdtlsqr.Execute
                    
                Next I
             End If
             
            Dim one As String
            Dim two As String
            Dim three As String
            Dim four As String
            Dim five As String
            one = Mid(VAR_REQNNO, 1, 3)
            two = Mid(VAR_REQNNO, 4, 2)
            three = Mid(VAR_REQNNO, 6, 2)
            four = Mid(VAR_REQNNO, 8, 6)
            five = one + two + three + four
            
            INSDMBCCOMMQR(0) = Trim(five)
            'upto here
            'INSDMBCCOMMQR(0) = Trim(req_chglen(VAR_REQNNO))
            
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                INSDMBCCOMMQR(1) = Null
                INSDMBCCOMMQR(1) = TXT_COMMENTS.Text
            Else
                INSDMBCCOMMQR(1) = Null
            
            End If
            INSDMBCCOMMQR.Execute
         '********************************************************************
             
             
             'ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(VAR_REQNNO)
            updrrequisitn_qr.Execute
            rdb.CommitTrans
            MsgBox "RECORD SAVED SUCCESSFLY"
            Exit Sub
MSG:
                rdb.RollbackTrans
                MsgBox Err.Description
        
        End If
        
        
         For I = 1 To MSFlexGrid1.Rows - 1
         
'JUSTINE 3-12-2004

         num = InStr(Trim(MSFlexGrid1.TextMatrix(I, 1)), "$")
         sc = Mid(Trim(MSFlexGrid1.TextMatrix(I, 1)), num + 1)
If sc = "F868" And Trim(MSFlexGrid1.TextMatrix(I, 2)) = "" Then
Else


            If SELDMA101FLAG = False Then
                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(I, 1))
                Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
                SELDMA101FLAG = True
            Else
                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(I, 1))
                SELDMA101RS.MoreResults
                SELDMA101RS.Requery
            End If

            'If SELDMA101RS.RowCount > 0 Then
            
           
             If Trim(sc) <> "" Then
                Dim t As String
                t = VAR_REQNNO
                INSDMPROCEDURQR(0) = t
                INSDMPROCEDURQR(1) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE")) 'trim(sc)
                'INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                
'                '***** Shraddha 20/07/10 *****
                'INSDMPROCEZ_QR(0) = t
                'INSDMPROCEZ_QR(1) = Trim(sc)
'                '*****************************
                
                If SELDTEMPLET_MFLAG = False Then
                    SELDTEMPLET_MQR(0) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
                    'SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    Set SELDTEMPLETRS = SELDTEMPLET_MQR.OpenResultset(1, 3)
                    'SELDTEMPLET_MFLAG = True
                Else
                    SELDTEMPLET_MQR(0) = Trim(sc) ' Trim(SELDMA101RS("SER_CODE"))
                    'SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    SELDTEMPLETRS.MoreResults
                    SELDTEMPLETRS.Requery
                End If
                
  ' 03-12-2004 JUSTINE
                '*****
                If Trim(sc) = "F869" Then
                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-") + 1, Len(Trim(MSFlexGrid1.TextMatrix(I, 3))) - InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-"))) Then
                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                        INSDMPROCEDURQR(3) = "H"
                        '***** Shraddha 20/07/10 *****
'                        INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                        INSDMPROCEZ_QR(3) = "H"
                        '*****************************
                    Else
                       ''' If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), 1, Len(Trim(MSFlexGrid1.TextMatrix(i, 3))) - InStr(Trim(MSFlexGrid1.TextMatrix(i, 3)), "-"))) Then
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1, InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-") - 1)) Then
                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            INSDMPROCEDURQR(3) = "N"
                            '***** Shraddha 20/07/10 *****
'                            INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                            INSDMPROCEZ_QR(3) = "N"
                            '*****************************
                        End If
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1, InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-") - 1)) Then
                           ' If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(ALKLV) Then
                               ' INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
                                INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                '***** Shraddha 20/07/10 *****
                                'INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                '*****************************
                                '***** 08/12/10 *****
                                INSDMPROCEDURQR(3) = "L"
                                'INSDMPROCEZ_QR(3) = "L"
                                '*******************
                            Else
                                INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                '***** Shraddha 20/07/10 *****
                                'INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                '*****************************
                           ' End If
                            'INSDMPROCEDURQR(3) = "L"
                        End If
                        If (Trim(MSFlexGrid1.TextMatrix(I, 2))) = "" Then
                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            '***** Shraddha 20/07/10 *****
'                            INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                            INSDMPROCEZ_QR(3) = ""
                            '*****************************
                            INSDMPROCEDURQR(3) = ""
                        End If
                    End If
                Else
                    '------For F829 ----- Shraddha Pandit 18/12/08 -------------------------
                    If Trim(sc) = "F829" Then
                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                        '***** Shraddha 20/07/10 *****
                        'INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                        '*****************************
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) <= 1 Then
                            INSDMPROCEDURQR(3) = "N"
                            '***** Shraddha 20/07/10 *****
                            'INSDMPROCEZ_QR(3) = "N"
                            '*****************************
                        Else
                            INSDMPROCEDURQR(3) = "H"
                            '***** Shraddha 20/07/10 *****
                            'INSDMPROCEZ_QR(3) = "H"
                            '*****************************
                        End If
                    Else
                '-----------------------------------------------------------------------
                        If SELDTEMPLETRS.RowCount > 0 Then
                            NU = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
                            
                            'Added by Shraddha on 12/09/2011
                            If Left(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1) = "<" Then
                                low = 0
                                high = Right(Trim(MSFlexGrid1.TextMatrix(I, 3)), Len(Trim(MSFlexGrid1.TextMatrix(I, 3))) - 1)
                            Else
                            low = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1, NU - 1)
                            high = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), NU + 1)
                            End If
                            W1 = "0"
                            'If (low <> "0" And high <> "0") And ((Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > high) Or ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE"))) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                            If high <> "0" And ((Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > high) Or ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE"))) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                                'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > high Then 'Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
                                    If ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE")) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '*****************************
                                    Else
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                        '*****************************
                                    End If
                                Else
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '***** Shraddha 20/07/10 *****
                                    'INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '*****************************
                                End If
                                INSDMPROCEDURQR(3) = "H"
                                '***** Shraddha 20/07/10 *****
                                'INSDMPROCEZ_QR(3) = "H"
                                '*****************************
                                W1 = "1"
                            Else
                                ' for normal value
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= low Or (UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "ABSENT" Or (low = "0" And high = "0")) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847"))
                                    'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= low Then ' Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    INSDMPROCEDURQR(3) = "N"
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                                    INSDMPROCEZ_QR(3) = "N"
                                    '*****************************
                                    W1 = "1"
                                End If
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < low And W1 <> "1" Then ' Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                        '***** Shraddha 20/07/10 *****
                                        'INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                        '*****************************
                                    Else
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '***** Shraddha 20/07/10 *****
                                        'INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '*****************************
                                    End If
                                    If low = "UPTO" Then
                                        INSDMPROCEDURQR(3) = "N"
                                    '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(3) = "N"
                                    '*****************************
                                    Else
                                        INSDMPROCEDURQR(3) = "L"
                                        '***** Shraddha 20/07/10 *****
                                        'INSDMPROCEZ_QR(3) = "L"
                                    '*****************************
                                    End If
                                End If
                                If (Trim(MSFlexGrid1.TextMatrix(I, 2))) = "" Then
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    INSDMPROCEDURQR(3) = ""
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                                    INSDMPROCEZ_QR(3) = ""
                                    '*****************************
                                End If
                            End If
                        End If
                    End If 'Of F829
                End If
                '*****
                   
                   
                   
                   
                   If SELDTEMPLETRS.EOF = False Then
                   INSDMPROCEDURQR(4) = Trim(SELDTEMPLETRS(3)) 'UOM
                   Else
                   
                   
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & PC & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    
                    'following code added by shweta 17-09-2014
                    If SELDBCALKPHSGGTRS.EOF Then
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & PC & "' AND (AGE_1)<=" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    End If
                    If Not SELDBCALKPHSGGTRS.EOF Then
                   
                   
                   INSDMPROCEDURQR(4) = Trim(SELDBCALKPHSGGTRS("REMARKS1")) 'UOM
                    End If
                   End If
                
                INSDMPROCEDURQR(5) = Trim(MSFlexGrid1.TextMatrix(I, 3))
                INSDMPROCEDURQR(6) = Trim(MSFlexGrid1.TextMatrix(I, 4))
                INSDMPROCEDURQR(7) = Trim(MSFlexGrid1.TextMatrix(I, 5))
                INSDMPROCEDURQR(8) = Trim(MSFlexGrid1.TextMatrix(I, 7))
                '***** Shraddha 22/07/10 *****
'                INSDMPROCEZ_QR(4) = Trim(SELDTEMPLETRS(3)) 'UOM
'                INSDMPROCEZ_QR(5) = Trim(MSFlexGrid1.TextMatrix(I, 3))
'                INSDMPROCEZ_QR(6) = Trim(MSFlexGrid1.TextMatrix(I, 4))
'                INSDMPROCEZ_QR(7) = Trim(MSFlexGrid1.TextMatrix(I, 5))
'                INSDMPROCEZ_QR(8) = Trim(MSFlexGrid1.TextMatrix(I, 7))
                '*****************************
                
                If dm = True Then
                    INSDMPROCEDURQR(3) = "N"
'                    INSDMPROCEZ_QR(3) = "N"
                End If
                
                
                'CHECK FOR GTT
                If PC = "F737" Or PC = "F738" Or PC = "F739" Or PC = "F740" Then
                    INSDMPROCEDURQR(3) = "N"
                End If
                
                
                
                INSDMPROCEDURQR.Execute
                'UPDATE DATE AND TIME IN RTIMELOG
                
            End If
            End If
    
         Next I
         dm = False
         
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(VAR_REQNNO)
            updrrequisitn_qr.Execute
            
            'Added by Milind 19/01/06

            one = Mid(VAR_REQNNO, 1, 3)
            two = Mid(VAR_REQNNO, 4, 2)
            three = Mid(VAR_REQNNO, 6, 2)
            four = Mid(VAR_REQNNO, 8, 6)
            five = one + two + three + four
            
            INSDMBCCOMMQR(0) = Trim(five)
            'upto here
            'INSDMBCCOMMQR(0) = Trim(req_chglen(VAR_REQNNO))
            
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                INSDMBCCOMMQR(1) = Null
                INSDMBCCOMMQR(1) = TXT_COMMENTS.Text
            Else
                INSDMBCCOMMQR(1) = Null
            
            End If
            INSDMBCCOMMQR.Execute
            
        
    
    End If
    
     '***** ADDED BY SHWETA  02/04/14 *****
        Dim RS_UPPATH As rdoQuery
        'STR11 = "UPDATE DISLIB.DREQUISITN SET PATHNO='" & Trim(txt_pathno.Text) & "' WHERE REQNO='" & Trim(Replace(cmb_reqno.Text), "/", "") & "' AND CASENO='" & Trim(MaskEdBox1.FormattedText) & "'"
        Set RS_UPPATH = rdb.CreateQuery("", "UPDATE DISLIB.DREQUISITN SET PATHNO='" & Trim(txt_pathno.Text) & "' WHERE REQNO='" & Trim(Replace(cmb_reqno.Text, "/", "")) & "' AND CASENO='" & Trim(MaskEdBox1.FormattedText) & "'")
  
        
        '*****************************
         
         
    RS_UPPATH.Execute
    
    
    If flagsave = 2 Then   'Update
            
        'ADDED BY MILIND ON 12-09-2006 FOR GTT TEST
        If gttflag = True Then
            
                gttflag = False
  '******* 23/04/2013 Aditi*********************************************
  '******* Insert GTT details ******************************************
                gttcode = ""
                gtt_qr(0) = VAR_REQNNO
                Set gtt_rs = gtt_qr.OpenResultset(1, 1)
                gttcode = Trim(gtt_rs!PRO_CODE)
                gtt_rs.Close
                
                For I = 1 To MSFlexGrid2.Rows - 1
                     
                    UPDGTTDTLSQR(0) = gttcode
                    j = 1
                    
                    If j <= 10 Then
                        If MSFlexGrid2.TextMatrix(I, 1) <> "" Then
                            UPDGTTDTLSQR(j) = MSFlexGrid2.TextMatrix(I, 1)
                            j = j + 1
                        Else
                            UPDGTTDTLSQR(j) = Null
                            j = j + 1
                            
                        End If
                        
                        If MSFlexGrid2.TextMatrix(I, 1) <> "" Then
                            UPDGTTDTLSQR(j) = "" ' MSFlexGrid2.TextMatrix(i, 2)
                            j = j + 1
                        Else
                            UPDGTTDTLSQR(j) = "" 'MSFlexGrid2.TextMatrix(i, 2)
                            j = j + 1

                        End If
                        UPDGTTDTLSQR(3) = Trim(VAR_REQNNO)
                        
                        If MSFlexGrid2.TextMatrix(I, 0) = "1 . FASTING" Then
                            UPDGTTDTLSQR(4) = 0
                            j = j + 1
                            ElseIf Trim(MSFlexGrid2.TextMatrix(I, 0)) = "2 . 2 hrs." Then
                            insgttdtlsqr(4) = "2"
                            j = j + 1
                        Else
                            UPDGTTDTLSQR(4) = Mid(MSFlexGrid2.TextMatrix(I, 0), 5, 3)
                            j = j + 1
                        End If
                    End If
                    
                    
                    UPDGTTDTLSQR.Execute
               Next I
               
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                UPDDMBCCOMMQR(0) = Null
                UPDDMBCCOMMQR(0) = Trim(TXT_COMMENTS.Text)
            Else
                UPDDMBCCOMMQR(0) = Null
                
            End If
            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
            UPDDMBCCOMMQR.Execute
   '*******************************************************************
                rdb.CommitTrans
                MsgBox "RECORD SAVED SUCCESSFLY..."
                Exit Sub

        End If
        'UPTO HERE GTT TEST
        
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
            For I = 1 To MSFlexGrid1.Rows - 1
             num = InStr(Trim(MSFlexGrid1.TextMatrix(I, 1)), "$")
             sc = Mid(Trim(MSFlexGrid1.TextMatrix(I, 1)), num + 1)
         
        
''
                'UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
             If Trim(sc) <> "" Then
                If SELDTEMPLET_MFLAG = False Then
                    SELDTEMPLET_MQR(0) = Trim(sc)
                   ' SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    Set SELDTEMPLETRS = SELDTEMPLET_MQR.OpenResultset(1, 3)
                    'SELDTEMPLET_MFLAG = True
                Else
                    SELDTEMPLET_MQR(0) = Trim(sc)
                    'SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    SELDTEMPLETRS.MoreResults
                    SELDTEMPLETRS.Requery
                End If
                    
                If SELDTEMPLETRS.RowCount > 0 Then
                    '****
                    
                    If Trim(sc) = "F869" Then
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(ALKHV) Then
                        
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= 500 Then
                                'UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
                                
                            'Else
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            'End If
                           
                            UPDDMPROCEDURQR(1) = "H"
                        Else
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(ALKLV) Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = "N"
                            End If
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                Else
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                End If
                                UPDDMPROCEDURQR(1) = "L"
                            End If
                            If Trim(MSFlexGrid1.TextMatrix(I, 2)) = "" Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = ""
                            End If
                        End If
                        UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                    Else
                      'JUSTIN 05-10-2004
                       '---------- Shraddha Pandit 18/12/08 -----------
                       'For CRP
                        If Trim(sc) = "F829" Then
                            
                            UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) <= 1 Then
                                UPDDMPROCEDURQR(1) = "N"
                            Else
                                UPDDMPROCEDURQR(1) = "H"
                            End If
                            
                        Else
                       '-----------------------------------------------
                      
                        If Left(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1) = "<" Then
                        
                        Else
                         NU = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
                         low = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1, NU - 1)
                         high = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), NU + 1)
                          W1 = "0"
                        End If
                        
                        If (low <> "0" And high <> "0") And Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > high Or ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE")) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > high Then 'Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
                                'If ((UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "TRACE") And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                                If ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE") And (Right(MSFlexGrid1.TextMatrix(I, 1), 4) = "F846" Or Right(MSFlexGrid1.TextMatrix(I, 1), 4) = "F847")) Then
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                Else
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                End If
                            Else
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            End If
                            UPDDMPROCEDURQR(1) = "H"
                         Else
                           'for normal value
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= low Or (UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "ABSENT" Or (low = "0" And high = "0")) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Or (low = "0" And high = "0") Then
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = "N"
                            End If
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < low Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
                                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) <> "0" Then
                                        UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                    Else
                                        UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    End If
                                Else
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                End If
                                If low = "UPTO" Then
                                    UPDDMPROCEDURQR(1) = "N"
                                Else
                                    UPDDMPROCEDURQR(1) = "L"
                                End If
                            End If
                            If Trim(MSFlexGrid1.TextMatrix(I, 2)) = "" Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = ""
                            End If
                        End If
                        End If 'OF F829
                    End If
                    
       
                    
                    
                    '****
                End If
                 UPDDMPROCEDURQR(2) = Trim(SELDTEMPLETRS("UOM"))
                 UPDDMPROCEDURQR(3) = Trim(MSFlexGrid1.TextMatrix(I, 3))
                 UPDDMPROCEDURQR(4) = Trim(MSFlexGrid1.TextMatrix(I, 4))
                 UPDDMPROCEDURQR(5) = req_chglen(cmb_reqno.Text)
                 UPDDMPROCEDURQR(6) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
                 UPDDMPROCEDURQR.Execute
            End If
            
    
         Next I
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
           
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                UPDDMBCCOMMQR(0) = Null
                UPDDMBCCOMMQR(0) = Trim(TXT_COMMENTS.Text)
            Else
                UPDDMBCCOMMQR(0) = Null
                
            End If
            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
            UPDDMBCCOMMQR.Execute
            
            
    End If
    
    
    
    
    
    
    
    
      If flagsave = 3 Then  'Update
            Dim UPDDMPROCEaddQR As rdoQuery
         'str = "UPDATE DISLIB.DMPROCEadd SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=? WHERE REQNO=? AND PRO_CODE=? "
         str = "UPDATE DISLIB.DMPROCEadd SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=?,ADDENDDT =current date, ADDENDTM=current time,IPADD=?,USERID=?,ADDENDBY=?,fld1=?,REMARKS=? WHERE REQNO=? AND PRO_CODE=? "
         Set UPDDMPROCEaddQR = rdb.CreateQuery("", str)
            
            
            
        'ADDED BY MILIND ON 12-09-2006 FOR GTT TEST
        If gttflag = True Then
            
                gttflag = False
  '******* 23/04/2013 Aditi*********************************************
  '******* Insert GTT details ******************************************
                gttcode = ""
                gtt_qr(0) = VAR_REQNNO
                Set gtt_rs = gtt_qr.OpenResultset(1, 1)
                gttcode = Trim(gtt_rs!PRO_CODE)
                gtt_rs.Close
                
                For I = 1 To MSFlexGrid2.Rows - 1
                     
                    UPDGTTDTLSQR(0) = gttcode
                    j = 1
                    
                    If j <= 10 Then
                        If MSFlexGrid2.TextMatrix(I, 1) <> "" Then
                            UPDGTTDTLSQR(j) = MSFlexGrid2.TextMatrix(I, 1)
                            j = j + 1
                        Else
                            UPDGTTDTLSQR(j) = Null
                            j = j + 1
                            
                        End If
                        
                        If MSFlexGrid2.TextMatrix(I, 2) <> "" Then
                            UPDGTTDTLSQR(j) = MSFlexGrid2.TextMatrix(I, 2)
                            j = j + 1
                        Else
                            UPDGTTDTLSQR(j) = MSFlexGrid2.TextMatrix(I, 2)
                            j = j + 1
                            
                        End If
                        UPDGTTDTLSQR(3) = Trim(VAR_REQNNO)
                        
                        If MSFlexGrid2.TextMatrix(I, 0) = "1 . FASTING" Then
                            UPDGTTDTLSQR(4) = 0
                            j = j + 1
                        Else
                            UPDGTTDTLSQR(4) = Mid(MSFlexGrid2.TextMatrix(I, 0), 5, 3)
                            j = j + 1
                        End If
                    End If
                    
                    
                    UPDGTTDTLSQR.Execute
               Next I
               
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                UPDDMBCCOMMQR(0) = Null
                UPDDMBCCOMMQR(0) = Trim(TXT_COMMENTS.Text)
            Else
                UPDDMBCCOMMQR(0) = Null
                
            End If
            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
            UPDDMBCCOMMQR.Execute
   '*******************************************************************
                rdb.CommitTrans
                MsgBox "RECORD SAVED SUCCESSFLY..."
                Exit Sub

        End If
        'UPTO HERE GTT TEST
        
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
            For I = 1 To MSFlexGrid1.Rows - 1
             num = InStr(Trim(MSFlexGrid1.TextMatrix(I, 1)), "$")
             sc = Mid(Trim(MSFlexGrid1.TextMatrix(I, 1)), num + 1)
         
        
''
                'UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
             If Trim(sc) <> "" Then
                If SELDTEMPLET_MFLAG = False Then
                    SELDTEMPLET_MQR(0) = Trim(sc)
                   ' SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    Set SELDTEMPLETRS = SELDTEMPLET_MQR.OpenResultset(1, 3)
                    'SELDTEMPLET_MFLAG = True
                Else
                    SELDTEMPLET_MQR(0) = Trim(sc)
                    'SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    SELDTEMPLETRS.MoreResults
                    SELDTEMPLETRS.Requery
                End If
                    
                If SELDTEMPLETRS.RowCount > 0 Then
                    '****
                    
                    If Trim(sc) = "F869" Then
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(ALKHV) Then
                        
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= 500 Then
                                'UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
                                
                            'Else
                                UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            'End If
                           
                            UPDDMPROCEaddQR(1) = "H"
                        Else
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(ALKLV) Then
                                UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEaddQR(1) = "N"
                            End If
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                    UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                Else
                                    UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                End If
                                UPDDMPROCEaddQR(1) = "L"
                            End If
                            If Trim(MSFlexGrid1.TextMatrix(I, 2)) = "" Then
                                UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEaddQR(1) = ""
                            End If
                        End If
                        UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                    Else
                      'JUSTIN 05-10-2004
                       '---------- Shraddha Pandit 18/12/08 -----------
                       'For CRP
                        If Trim(sc) = "F829" Then
                            
                            UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) <= 1 Then
                                UPDDMPROCEaddQR(1) = "N"
                            Else
                                UPDDMPROCEaddQR(1) = "H"
                            End If
                            
                        Else
                       '-----------------------------------------------
                      
                        If Left(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1) = "<" Then
                        
                        Else
                         NU = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
                         low = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1, NU - 1)
                         high = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), NU + 1)
                          W1 = "0"
                        End If
                        
                        If (low <> "0" And high <> "0") And Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > high Or ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE")) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > high Then 'Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
                                'If ((UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(i, 2))) = "TRACE") And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                                If ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE") And (Right(MSFlexGrid1.TextMatrix(I, 1), 4) = "F846" Or Right(MSFlexGrid1.TextMatrix(I, 1), 4) = "F847")) Then
                                    UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                Else
                                    UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                End If
                            Else
                                UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            End If
                            UPDDMPROCEaddQR(1) = "H"
                         Else
                           'for normal value
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= low Or (UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "ABSENT" Or (low = "0" And high = "0")) Then 'And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Or (low = "0" And high = "0") Then
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEaddQR(1) = "N"
                            End If
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < low Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
                                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) <> "0" Then
                                        UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                    Else
                                        UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    End If
                                Else
                                    UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                End If
                                If low = "UPTO" Then
                                    UPDDMPROCEaddQR(1) = "N"
                                Else
                                    UPDDMPROCEaddQR(1) = "L"
                                End If
                            End If
                            If Trim(MSFlexGrid1.TextMatrix(I, 2)) = "" Then
                                UPDDMPROCEaddQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEaddQR(1) = ""
                            End If
                        End If
                        End If 'OF F829
                    End If
                    
       
                    
                    
                    '****
                End If
'                 UPDDMPROCEaddQR(2) = Trim(SELDTEMPLETRS("UOM"))
'                 UPDDMPROCEaddQR(3) = Trim(MSFlexGrid1.TextMatrix(i, 3))
'                 UPDDMPROCEaddQR(4) = Trim(MSFlexGrid1.TextMatrix(i, 4))
'                 UPDDMPROCEaddQR(5) = req_chglen(cmb_reqno.Text)
'                 UPDDMPROCEaddQR(6) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
'                 UPDDMPROCEaddQR.Execute

UPDDMPROCEaddQR(2) = Trim(SELDTEMPLETRS("UOM"))
                 UPDDMPROCEaddQR(3) = Trim(MSFlexGrid1.TextMatrix(I, 3))
                 UPDDMPROCEaddQR(4) = Trim(MSFlexGrid1.TextMatrix(I, 4))
                 
                 UPDDMPROCEaddQR(5) = Winsock1.LocalIP
                 UPDDMPROCEaddQR(6) = Trim(logonfrm.usrid)
                 UPDDMPROCEaddQR(7) = Trim(cmb_addendumby.Text)
                 
                  UPDDMPROCEaddQR(8) = Trim(lbl_addtext) & " " & Trim(txt_addendum.Text)
                  UPDDMPROCEaddQR(9) = Trim(MSFlexGrid1.TextMatrix(I, 7))
                 UPDDMPROCEaddQR(10) = req_chglen(cmb_reqno.Text)
                 UPDDMPROCEaddQR(11) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
                 UPDDMPROCEaddQR.Execute



            End If
            
    
         Next I
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
           
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                UPDDMBCCOMMQR(0) = Null
                UPDDMBCCOMMQR(0) = Trim(TXT_COMMENTS.Text)
            Else
                UPDDMBCCOMMQR(0) = Null
                
            End If
            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
            UPDDMBCCOMMQR.Execute
            
            
    End If
    
    
    
    
        If adenum_flg = True Then
           UPDRTIMECADD_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
           UPDRTIMECADD_QR(1) = time1
           UPDRTIMECADD_QR(2) = req_chglen(cmb_reqno.Text)
           UPDRTIMECADD_QR.Execute
        
'          If SELDADDENDUMFLAG = False Then
'            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'            Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'            SELDADDENDUMFLAG = True
'          Else
'            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'            SELDADDENDUM_RS.MoreResults
'            SELDADDENDUM_RS.Requery
'          End If
          
            
'
'          If SELDADDENDUM_RS.RowCount > 0 Then
'            SELDADDENDUM_RS.MoveLast
'            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'            INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
'            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'            INSDADDENDUM_QR(3) = time1
'            INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
'            INSDADDENDUM_QR(5) = Null
'            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
'            INSDADDENDUM_QR.Execute
'         Else
'            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'            INSDADDENDUM_QR(1) = "1"
'            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'            INSDADDENDUM_QR(3) = time1
'            INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
'            INSDADDENDUM_QR(5) = Null
'            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
'            INSDADDENDUM_QR.Execute
'         End If
       End If
        
    
rdb.CommitTrans
MsgBox "Record is Saved"
'btnenabl ("INI")
BTNENABL1 ("INI")
txt_observed.Visible = False
txt_gttobsvd.Visible = False
'txt_gttobsvdurn.Visible = False
cmb_machine.Visible = False
Exit Sub
MSG123: 'shweta 07-NOV-2012
   rdb.RollbackTrans
MsgBox Err.Description
End Sub

Public Sub STAFF()
 If NEWFLAG = True Or viewflag = False Then
            NEWFLAG = False
            ''******ICU & RECOVERY WARD****************
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
                  lbl_rwd.Visible = False
               ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                  lblicu.Visible = False
                  lbl_rwd.Visible = True
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
                txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
        
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
 '12-01-2005 JUSTINE '******
 
            If (reqspn_rs("NMETERIALS") = "Stone Analysis") Then
                Exit Sub
            End If
 
 ICASENO = MaskEdBox1.FormattedText
 IREQNO = VAR_REQNNO
 Dim TEST_DET, MACHINENO
 MSFlexGrid1.Rows = 1
 I = 1
          'FETCH TEST_CODE FROM DMULPROCE
          If seldmulprocflag = False Then
                SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
                seldmulprocflag = True
           Else
                SELDMULPROCQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                seldmulprocrs.MoreResults
                seldmulprocrs.Requery
           End If
           If seldmulprocrs.RowCount > 0 Then
           Do Until seldmulprocrs.EOF
           PC = seldmulprocrs!PRO_CODE
           'If PC = "F895" Or PC = "F896" Or PC = "F894" Or PC = "F846" Or PC = "F847" Or PC = "F897" Or PC = "F898" Or PC = "F947" Or PC = "F845" Then
           'Next check for F898 is added by Shraddha Pandit on 30/01/09
           'If PC = "F895" Or PC = "F896" Or PC = "F894" Or PC = "F846" Or PC = "F847" Or PC = "F947" Or PC = "F845" Or PC = "F897" Then ' Or Or PC = "F898"
           'Above statement commented and next statement added by Shraddha on 09/06/09 '(Removed F845 and F894)
            If PC = "F895" Or PC = "F896" Or PC = "F846" Or PC = "F847" Or PC = "F897" Or PC = "F898" Or PC = "F845" Then
            If SELMA101FLAG = False Then
                selMA101QR(0) = PC
                Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
               
                SELMA101FLAG = True
            Else
                selMA101QR(0) = PC
                SELMA101RS.MoreResults
                SELMA101RS.Requery
            End If


            TEST_DET = Trim(SELMA101RS!SER_DET)
            'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE

            If SELDTEMPLETFLAG = False Then
               SELDTEMPLETQR(0) = PC
               Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
               SELDTEMPLETFLAG = True
            Else
               SELDTEMPLETQR(0) = PC
               SELDTEMPLETRS.MoreResults
               SELDTEMPLETRS.Requery
            End If
            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
            'MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & vbTab & RNAGE1
            
'            I = I + 1 'Commented by Shraddha on 08/01/10
Else
''''23-02-2013
'NEW REFERANCE RANGE DATA
                Dim RS_VWREFRNG As rdoResultset
                 Set RS_VWREFRNG = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & PC & "' ", 1, 1)
        
                
                If Not RS_VWREFRNG.EOF Then
                Dim agecal As Double
                agecal = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                txt_age.Text = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                If Trim(txt_age.Text) <= 1 Then
                Dim age3 As Double
                 age3 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                age3 = "." & age3 '(age Mod 12)
                If age3 = ".0" Then
                    age3 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                    age3 = ".0" & (age3 Mod 30)
                End If
                txt_age.Text = Trim(age3)
            End If
                   ' ALKAGEO 'FOR F876
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & PC & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range

                        If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Or Left(Trim(SELDBCALKPHSGGTRS!range), 1) = ">" Then
                        Else
                        ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        End If
                    End If
                End If
        End If
            seldmulprocrs.MoveNext
            Loop
            End If
 ' FETCH TEST_CODE FROM INTERFACE
            If SELINTER_FLAG = False Then
                SELINTER_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set SELINTER_RS = SELINTER_QR.OpenResultset(1, 3)
                SELINTER_FLAG = True
            Else
                SELINTER_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                SELINTER_RS.MoreResults
                SELINTER_RS.Requery
            End If
            Dim PREVCODE As String
 If SELINTER_RS.RowCount > 0 Then
    Do Until SELINTER_RS.EOF
    PC = SELINTER_RS!testcode
    If PREVCODE = PC Then
    Else
            'THIS IS FOR GETTING TEST DESCRIPTION FOR TEST CODE

            If SELMA101FLAG = False Then
                selMA101QR(0) = PC
                Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
                
                SELMA101FLAG = True
            Else
                selMA101QR(0) = PC
                SELMA101RS.MoreResults
                SELMA101RS.Requery
            End If
            TEST_DET = Trim(SELMA101RS!SER_DET)
            'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE

            If SELDTEMPLETFLAG = False Then
               SELDTEMPLETQR(0) = PC
               Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
               SELDTEMPLETFLAG = True
            Else
               SELDTEMPLETQR(0) = PC
               SELDTEMPLETRS.MoreResults
               SELDTEMPLETRS.Requery
            End If
            'THIS IF FOR F869 (A & C) IN DTEMPLET
            If PC = "F869" Then
                ALKAGE
                
                If SELDBCALKPHSFLAG = False Then
                        SELDBCALKPHSQR(0) = PC
                        SELDBCALKPHSQR(1) = RNAGE
                        SELDBCALKPHSQR(2) = SS
                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
                        SELDBCALKPHSFLAG = True
                Else
                        SELDBCALKPHSQR(0) = PC
                        SELDBCALKPHSQR(1) = RNAGE
                        SELDBCALKPHSQR(2) = SS
                        SELDBCALKPHSRS.MoreResults
                        SELDBCALKPHSRS.Requery
                End If

                RNAGE1 = SELDBCALKPHSRS!range
             'justin 18-08-2004
                ALKLV = Left(Trim(SELDBCALKPHSRS("RANGE")), 2)
                ALKHV = Right(Trim(SELDBCALKPHSRS("RANGE")), 3)
                

              End If
              
              If PC = "F999" Then
                    GTTAGE
                    If SELDBCALKPHSGGTFLAG = False Then
                            SELDBCALKPHSGGTQR(0) = PC
                            SELDBCALKPHSGGTQR(1) = RNAGE
                           
                            SELDBCALKPHSGGTQR(2) = SS
                            Set SELDBCALKPHSGGTRS = SELDBCALKPHSGGTQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDBCALKPHSGGTQR(0) = PC
                            SELDBCALKPHSGGTQR(1) = RNAGE
                            SELDBCALKPHSGGTQR(2) = SS
                            SELDBCALKPHSGGTRS.MoreResults
                            SELDBCALKPHSGGTRS.Requery
                    End If
                            
                            
                    RNAGE1 = SELDBCALKPHSGGTRS!range
                 
                    If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Then
                    
                    Else
                        ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 2)
                        ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        RNAGE1 = ALKLV & ALKHV
                    End If
                    
                End If
            
            'THIS IS FOR MALE & FEMALE FROM DTEMPLET
           ' If PC = "F850" Or PC = "F851" Or PC = "F856" Or PC = "F878" Or PC = "F904" Or PC = "F907" Or PC = "F877" Or PC = "F961" Or PC = "F055" Then
            If PC = "F856" Or PC = "F878" Or PC = "F904" Or PC = "F907" Or PC = "F877" Or PC = "F961" Or PC = "F055" Then
                If SELDTEMPLETFLAG1 = False Then
                    SELDTEMPLETQR1(0) = PC
                    SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
                    Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
                    SELDTEMPLETFLAG1 = True
                Else
                    SELDTEMPLETQR1(0) = PC
                    SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
                    SELDTEMPLETRS1.MoreResults
                    SELDTEMPLETRS1.Requery
                End If
                RNAGE1 = Trim(SELDTEMPLETRS1!LOWVALUE) & "-" & Trim(SELDTEMPLETRS1!HIGHVALUE)
            Else
                If PC = "F869" Or PC = "F999" Then
                Else
                'ADDED BY SHWETA FOR NEW REFRENCE RANGE 01-10-2013
                 Set RS_VWREFRNG = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & PC & "' ", 1, 1)
        
                
                If Not RS_VWREFRNG.EOF Then
                    ' Dim agecal As Double
                agecal = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                txt_age.Text = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                        If Trim(txt_age.Text) <= 1 Then
                       ' Dim age3 As Double
                         age3 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                        age3 = "." & age3 '(age Mod 12)
                        If age3 = ".0" Then
                            age3 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                            age3 = ".0" & (age3 Mod 30)
                        End If
                        txt_age.Text = Trim(age3)
                    End If
                           ' ALKAGEO 'FOR F876
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & PC & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    
                    'following code added by shweta 17-09-2014
                    If SELDBCALKPHSGGTRS.EOF Then
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & PC & "' AND (AGE_1)<=" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    End If
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range

                        If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Then
                        Else
                        ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        End If
                    End If
                Else
                 If PC = "F810a" Then
                 Else
                 
                   If Trim(SELDTEMPLETRS!LOWVALUE) = "" Then
                   RNAGE1 = Trim(SELDTEMPLETRS!HIGHVALUE)
                   Else
                   RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                   End If
                 End If
                 
                End If
                End If
            End If
            'THIS IS FOR PICKUP TESTRESULT FROM INTERFACE/DINTERFACE

            INTERFACEREAD_QR(0) = IREQNO
            INTERFACEREAD_QR(1) = PC
            Set INTERFACEREAD_RS = INTERFACEREAD_QR.OpenResultset(1, 1)

            If INTERFACEREAD_RS.RowCount > 0 Then
                INTFAC = INTERFACEREAD_RS("testresult")
                If Trim(INTFAC) = "" Then
                    INTMODE = ""
                    MACHINENO = ""
                    Else
                    INTMODE = "I"
                    MACHINENO = INTERFACEREAD_RS("MACHINENO")
                End If
            Else
                INTFAC = ""
                INTMODE = ""
            End If
            If INTERFACEREAD_RS.RowCount = 0 Then
                '''FROM HERE
'''                    SELMSADINTF_QR(0) = IREQNO
'''                    SELMSADINTF_QR(1) = PC
'''                    Set SELMSADINTF_RS = SELMSADINTF_QR.OpenResultset(1, 1)
'''
'''                    If SELMSADINTF_RS.RowCount > 0 Then
'''                        INTFAC = SELMSADINTF_RS("testresult")
'''                        If Trim(INTFAC) = "" Then
'''                            INTMODE = ""
'''                            MACHINENO = ""
'''                        Else
'''                            INTMODE = "I"
'''                            MACHINENO = SELMSADINTF_RS("MACHINENO")
'''                        End If
'''                    Else
'''                        INTFAC = ""
'''                        INTMODE = ""
'''                    End If
                    'UPTO HERE
                    
                    If PREVTEST = TEST_DET Then
                    Else
                                       
                    MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & vbTab & MACHINENO '& vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL")))
                    End If
                    PREVTEST = TEST_DET
                
            Else
                MSFlexGrid1.AddItem I & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL"))) & vbTab & MACHINENO
            ''THIS IS FOR DEFINE PANIC VALUE BASED ON OBSER VALUE

            If Len(Trim(MSFlexGrid1.TextMatrix(I, 2))) = 0 Then
            Else

            'JUSTIN 18-08-2004
             'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
        
        If Trim(MSFlexGrid1.TextMatrix(I, 3)) <> "" Then
            NU = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
            RN1 = Trim(MSFlexGrid1.TextMatrix(I, 3))
            If InStr(RN1, "-") Then
            RN1 = Val(Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), NU + 1))
            ElseIf InStr(RN1, "<") Then
            RN1 = Replace(Trim(MSFlexGrid1.TextMatrix(I, 3)), "<", "")
            End If
            
            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(RN1) Then
                MSFlexGrid1.row = I
                MSFlexGrid1.col = 2
                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = Val(Trim(MSFlexGrid1.TextMatrix(I, 2)))
                Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                End If
             Else
                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                    MSFlexGrid1.row = I
                    MSFlexGrid1.CellBackColor = vbWhite
                    'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Val(Trim(MSFlexGrid1.TextMatrix(i, 2)))
                    If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) Then
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    Else
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                    End If
                End If
                'If Len(Trim(MSFlexGrid1.TextMatrix(i, 2))) = 0 Then
                'Else
                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                    MSFlexGrid1.row = I
                    MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
                        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) Then
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                        Else
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                        End If
                End If
              End If
            End If
           End If
        End If
        End If
           I = I + 1
        PREVCODE = PC
            SELINTER_RS.MoveNext
         
    Loop
 End If
 
 
 
'''            If seldmulprocflag = False Then
'''                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
'''                Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
'''                seldmulprocflag = True
'''           Else
'''                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
'''                seldmulprocrs.MoreResults
'''                seldmulprocrs.Requery
'''           End If
'''           If seldmulprocrs.RowCount > 0 Then
'''
'''                MSFlexGrid1.Rows = 1
'''                i = 1
'''
'''                While Not seldmulprocrs.EOF
'''                   If seldmulprocrs("pro_code") = "F849" Or seldmulprocrs("PRO_CODE") = "F854" Or seldmulprocrs("PRO_CODE") = "F860" Or seldmulprocrs("PRO_CODE") = "F865" Or seldmulprocrs("PRO_CODE") = "F875" Or seldmulprocrs("PRO_CODE") = "F880" Or seldmulprocrs("PRO_CODE") = "F802" Then
'''                       If SELDBCPROSUBFLAG = False Then
'''                            SELDBCPROSUBQR(0) = seldmulprocrs("pro_code")
'''                            Set SELDBCPROSUBRS = SELDBCPROSUBQR.OpenResultset(1, 3)
'''                            SELDBCPROSUBFLAG = True
'''                        Else
'''                            SELDBCPROSUBQR(0) = seldmulprocrs("pro_code")
'''                            SELDBCPROSUBRS.MoreResults
'''                            SELDBCPROSUBRS.Requery
'''                        End If
'''                        If SELDBCPROSUBRS.RowCount > 0 Then
'''                            While Not SELDBCPROSUBRS.EOF
'''                                If SELMA101FLAG = False Then
'''                                    selMA101QR(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                    Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
'''                                    SELMA101FLAG = True
'''                                Else
'''                                    selMA101QR(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                    SELMA101RS.MoreResults
'''                                    SELMA101RS.Requery
'''                                End If
'''                                If SELDTEMPLETFLAG = False Then
'''                                    SELDTEMPLETQR(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
'''                                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'''                                    SELDTEMPLETFLAG = True
'''                                Else
'''                                    SELDTEMPLETQR(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
'''                                    SELDTEMPLETRS.MoreResults
'''                                    SELDTEMPLETRS.Requery
'''                                End If
'''
'''                                '******
'''                                If SELDBCPROSUBRS("PROCODESUB") = "F869" Then
'''                                        ALKAGE
'''                                    If SELDBCALKPHSFLAG = False Then
'''                                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                        SELDBCALKPHSQR(1) = RNAGE
'''                                        SELDBCALKPHSQR(2) = SS
'''                                        Set SELBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
'''                                        SELDBCALKPHSFLAG = True
'''                                    Else
'''                                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                        SELDBCALKPHSQR(1) = RNAGE
'''                                        SELDBCALKPHSQR(2) = SS
'''                                        Set SELBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
'''                                        SELBCALKPHSRS.MoreResults
'''                                        SELBCALKPHSRS.Requery
'''                                    End If
'''                                    If SELDTEMPLETRS.RowCount > 0 Then
'''                                    If SELMA101RS.RowCount > 0 Then
'''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELBCALKPHSRS("RANGE"))
'''                                        Else
'''                                            SELDTEMPLETRS.MoveNext
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELBCALKPHSRS("RANGE"))
'''                                        End If
'''                                        ALKLV = Left(Trim(SELBCALKPHSRS("RANGE")), 2)
'''                                        ALKHV = Right(Trim(SELBCALKPHSRS("RANGE")), 3)
'''                                        i = i + 1
'''                                    End If
'''                                    End If
'''                               '''ADD BY MONO ON 17-09-03 FOR M &F
'''                                ElseIf SELDBCPROSUBRS("PROCODESUB") = "F850" Or SELDBCPROSUBRS("PROCODESUB") = "F851" Or SELDBCPROSUBRS("PROCODESUB") = "F856" Or SELDBCPROSUBRS("PROCODESUB") = "F878" Or SELDBCPROSUBRS("PROCODESUB") = "F904" Or SELDBCPROSUBRS("PROCODESUB") = "F907" Then
'''                                    If SELDTEMPLETFLAG1 = False Then
'''                                        SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
'''                                        Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
'''                                        SELDTEMPLETFLAG1 = True
'''                                    Else
'''                                        SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
'''                                        SELDTEMPLETRS1.MoreResults
'''                                        SELDTEMPLETRS1.Requery
'''                                    End If
'''                                    If SELDTEMPLETRS1.RowCount > 0 Then
'''                                        If SELMA101RS.RowCount > 0 Then
'''                                            If SELDTEMPLETRS1("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS1("SEX")) = "" Then
'''                                                'MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS1("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS1("HIGHVALUE"))
'''                                                'COMMENT
'''                                                ' comments removed for interfaceing 0n 30-12-2003
'''                                                '' INTERFACEX
'''                                               MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                                 ' comments removed for interfaceing 0n 30-12-2003
'''                                            Else
'''                                                SELDTEMPLETRS1.MoveNext
'''                                                'MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS1("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS1("HIGHVALUE"))
'''                                                'COMMENT
'''                                                ' comments removed for interfaceing 0n 30-12-2003
'''                                                '' INTERFACEX
'''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                                ' comments removed for interfaceing 0n 30-12-2003
'''                                            End If
'''                                            i = i + 1
'''                                        End If
'''                                    End If
'''                                '''ADD BY MONO ON 17-09-03 FOR M &F
'''                               Else
'''                                    If SELDTEMPLETRS.RowCount > 0 Then
'''                                    If SELMA101RS.RowCount > 0 Then
'''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                        Else
'''                                            SELDTEMPLETRS.MoveNext
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                        End If
'''                                        i = i + 1
'''                                    End If
'''                                    End If
'''                                End If
''''                                '******
''''                                If SELDTEMPLETRS.RowCount > 0 Then
''''                                    If SELMA101RS.RowCount > 0 Then
''''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
''''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                                        Else
''''                                            SELDTEMPLETRS.MoveNext
''''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                                        End If
''''                                        i = i + 1
''''                                    End If
''''                                End If
'''
'''                                SELDBCPROSUBRS.MoveNext
'''                            Wend
'''                        End If
'''                   Else
'''                            If SELMA101FLAG = False Then
'''                                selMA101QR(0) = seldmulprocrs("PRO_CODE")
'''                                Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
'''                                SELMA101FLAG = True
'''                            Else
'''                                selMA101QR(0) = seldmulprocrs("PRO_CODE")
'''                                SELMA101RS.MoreResults
'''                                SELMA101RS.Requery
'''                            End If
'''                            If SELDTEMPLETFLAG = False Then
'''                                SELDTEMPLETQR(0) = seldmulprocrs("PRO_CODE")
'''                                'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
'''                                Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'''                                SELDTEMPLETFLAG = True
'''                            Else
'''                                SELDTEMPLETQR(0) = seldmulprocrs("PRO_CODE")
'''                                'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
'''                                SELDTEMPLETRS.MoreResults
'''                                SELDTEMPLETRS.Requery
'''                            End If
'''                            '*******
'''                                If seldmulprocrs("PRO_CODE") = "F869" Then
'''                                     ALKAGE
'''                                    If SELDBCALKPHSFLAG = False Then
'''                                        SELDBCALKPHSQR(0) = seldmulprocrs("PRO_CODE")
'''                                        SELDBCALKPHSQR(1) = RNAGE
'''                                        SELDBCALKPHSQR(2) = SS
'''                                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
'''                                        SELDBCALKPHSFLAG = True
'''                                    Else
'''                                        SELDBCALKPHSQR(0) = seldmulprocrs("PRO_CODE")
'''                                        SELDBCALKPHSQR(1) = RNAGE 'Trim(txt_age.Text)
'''                                        SELDBCALKPHSQR(2) = SS
'''                                        'Set SELBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
'''                                        'SELDBCALKPHSFLAG = True
'''                                        SELDBCALKPHSRS.MoreResults
'''                                        SELDBCALKPHSRS.Requery
'''                                    End If
'''                                    If SELDTEMPLETRS.RowCount > 0 Then
'''                                        If SELMA101RS.RowCount > 0 Then
'''                                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
'''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
'''                                            Else
'''                                                SELDTEMPLETRS.MoveNext
'''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
'''                                            End If
'''                                            ALKLV = Left(Trim(SELDBCALKPHSRS("RANGE")), 2)
'''                                            ALKHV = Right(Trim(SELDBCALKPHSRS("RANGE")), 3)
'''                                            i = i + 1
'''                                        End If
'''                                    End If
'''                                '''ADD BY MONO ON 17-09-03 FOR M &F
'''                                ElseIf SELDBCPROSUBRS("PROCODESUB") = "F850" Or SELDBCPROSUBRS("PROCODESUB") = "F851" Or SELDBCPROSUBRS("PROCODESUB") = "F856" Or SELDBCPROSUBRS("PROCODESUB") = "F878" Or SELDBCPROSUBRS("PROCODESUB") = "F904" Or SELDBCPROSUBRS("PROCODESUB") = "F907" Then
'''                                    If SELDTEMPLETFLAG1 = False Then
'''                                        SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
'''                                        Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
'''                                        SELDTEMPLETFLAG1 = True
'''                                    Else
'''                                        SELDTEMPLETQR1(0) = SELDBCPROSUBRS("PROCODESUB")
'''                                        SELDTEMPLETQR1(1) = Trim(txt_sex.Text)
'''                                        SELDTEMPLETRS1.MoreResults
'''                                        SELDTEMPLETRS1.Requery
'''                                    End If
'''                                    If SELDTEMPLETRS1.RowCount > 0 Then
'''                                        If SELMA101RS.RowCount > 0 Then
'''                                            If SELDTEMPLETRS1("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS1("SEX")) = "" Then
'''                                                'MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS1("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS1("HIGHVALUE"))
'''                                                'COMMENT
'''                                                '' INTERFACEX
'''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                            Else
'''                                                SELDTEMPLETRS1.MoveNext
'''                                                'MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS1("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS1("HIGHVALUE"))
'''                                                'COMMENT
'''                                                '' INTERFACEX
'''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & INTFAC & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                            End If
'''                                            i = i + 1
'''                                        End If
'''                                    End If
'''                                '''ADD BY MONO ON 17-09-03 FOR M &F
'''                                Else
'''                                        If SELDTEMPLETRS.RowCount > 0 Then
'''                                            If SELMA101RS.RowCount > 0 Then
'''                                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
'''                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                                Else
'''                                                    SELDTEMPLETRS.MoveNext
'''                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                                End If
'''                                                i = i + 1
'''                                            End If
'''                                        End If
'''                                End If
'''                            '*******
''''                        If SELDTEMPLETRS.RowCount > 0 Then
''''                            If SELMA101RS.RowCount > 0 Then
''''                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
''''                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                                Else
''''                                    SELDTEMPLETRS.MoveNext
''''                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                                End If
''''                                i = i + 1
''''                            End If
''''                        End If
'''                    End If
'''                    seldmulprocrs.MoveNext
'''
'''                Wend
'''                MSFlexGrid1.Row = 1
'''                MSFlexGrid1.Col = 2
'''                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'''                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'''                txt_observed.Visible = True
'''
'''          ' 12-01-2004 justine
'''
'''
'''          '12-01-2004 justine
'''           End If
            
        '    seldrequisitn_qr(0) = MaskEdBox1.FormattedText
        '    seldrequisitn_rs.MoreResults
        '    seldrequisitn_rs.Requery
        '    If seldrequisitn_rs.RowCount > 0 Then
        '     prpathno = ""
        '     While Not seldrequisitn_rs.EOF
        '     prpathno = prpathno & " " & seldrequisitn_rs("pathno")
        '     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
        '     seldrequisitn_rs.MoveNext
        '     Wend
        '     txt_prepathreqno.Text = prpathno
        '    End If
        '    txt_addendum.Enabled = False
        '    cmb_addendumby.Enabled = False
            

        
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
                  lbl_rwd.Visible = False
               ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                  lblicu.Visible = False
                  lbl_rwd.Visible = True
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
                    MsgBox "Final Report is Generated, You cann't Modify the Final Report ,You can Add ADDENDUMS Only"
                    cmd_modify.Enabled = False 'Added on 09/06/09 by Shraddha
                    ADDENDUMFALG = True
                    cmd_adenum.Visible = True
                  Else
                  cmd_modify.Enabled = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                'cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
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
                txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
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
            
            If (reqspn_rs("NMETERIALS") = "Stone Analysis") Then
                Exit Sub
            End If
             '**** filling the flex grid
            If SELDPROCEDURFLAG = False Then
                SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
                Set SELDPROCEDURRS = SELDPROCEDURQR.OpenResultset(1, 3)
                SELDPROCEDURFLAG = True
           Else
                SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
                SELDPROCEDURRS.MoreResults
                SELDPROCEDURRS.Requery
           End If
           If SELDPROCEDURRS.RowCount > 0 Then
                
''                MSFlexGrid1.Rows = 1
''                i = 1
''
''                While Not SELDPROCEDURRS.EOF
''                    If SELMA101FLAG = False Then
''                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
''                        Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
''                        SELMA101FLAG = True
''                    Else
''                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
''                        SELMA101RS.MoreResults
''                        SELMA101RS.Requery
''                    End If
''                    If SELDTEMPLETFLAG = False Then
''                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
''                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                        SELDTEMPLETFLAG = True
''                    Else
''                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
''                        SELDTEMPLETRS.MoreResults
''                        SELDTEMPLETRS.Requery
''                    End If
''                    If SELDTEMPLETRS.RowCount > 0 Then
''                        If SELMA101RS.RowCount > 0 Then
''                            If SELDTEMPLETRS("SEX") = "M" Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''                                 MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & SELDPROCEDURRS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                            Else
''                                SELDTEMPLETRS.MoveNext
''                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & SELDPROCEDURRS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                            End If
''                        End If
''
''                            i = i + 1
''
''                    End If
''                    SELDPROCEDURRS.MoveNext
''
''                Wend
'''                MSFlexGrid1.Row = 1
'''                MSFlexGrid1.Col = 2
'''                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'''                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''                txt_observed.Visible = False
                    'MSFlexGrid1.Rows = 1
                    'MSFlexGrid1.Row = 1
                    I = 1
                    srno = 1
                While Not SELDPROCEDURRS.EOF
                    If SELMA101FLAG = False Then
                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
                        Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
                        
                        SELMA101FLAG = True
                    Else
                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
                        SELMA101RS.MoreResults
                        SELMA101RS.Requery
                    End If
                    If SELDTEMPLETFLAG = False Then
                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
                        'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                        SELDTEMPLETFLAG = True
                    Else
                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
                        'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                        SELDTEMPLETRS.MoreResults
                        SELDTEMPLETRS.Requery
                    End If
                    
'                    If Trim(SELDPROCEDURRS("PRO_CODE")) = "F869" Then
'                        'ALKLV = Left(Trim(SELDPROCEDURRS("NRANGE")), 2)
'                        'ALKHV = Right(Trim(SELDPROCEDURRS("NRANGE")), 3)
'                    End If
                    
                    
                    If SELDTEMPLETRS.RowCount > 0 Then
                        If SELMA101RS.RowCount > 0 Then
                            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                            MSFlexGrid1.row = srno
                            MSFlexGrid1.col = 0
                            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
                                MSFlexGrid1.TextMatrix(I, 0) = I
                                MSFlexGrid1.TextMatrix(I, 1) = Trim(SELMA101RS("SER_DET")) & "$" & Trim(SELDPROCEDURRS("PRO_CODE"))
                                If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
                                    MSFlexGrid1.row = srno
                                    MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                                Else
                                    If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
                                        MSFlexGrid1.col = 2
                                        MSFlexGrid1.CellBackColor = RGB(100, 255, 150)
                                    End If
                                    If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
                                        MSFlexGrid1.col = 2
                                        MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                    If Trim(SELDPROCEDURRS("obser_type")) = "" Then
                                        MSFlexGrid1.col = 2
                                        MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                End If
                                j = 0
                                
                                
                                
                                j = InStr(SELDPROCEDURRS("observalue"), "*")
                                
                                RNANGE = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                ''''
                                If PC = "F869" Or PC = "F999" Then
                Else
                
                j = 0
            j = InStr(SELDPROCEDURRS("observalue"), "*")
            RNAGE1 = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                ''''
                
                'ADDED BY SHWETA FOR NEW REFRENCE RANGE 01-10-2013
                 Set RS_VWREFRNG = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & SELDPROCEDURRS("PRO_CODE") & "' ", 1, 1)
        
                
                If Not RS_VWREFRNG.EOF Then
                    ' Dim agecal As Double
                agecal = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                txt_age.Text = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        If Trim(txt_age.Text) <= 1 Then
'                       ' Dim age3 As Double
'                         age3 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        age3 = "." & age3 '(age Mod 12)
'                        If age3 = ".0" Then
'                            age3 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                            age3 = ".0" & (age3 Mod 30)
'                        End If
'                        txt_age.Text = Trim(age3)
'                    End If
'                           ' ALKAGEO 'FOR F876
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & SELDPROCEDURRS("PRO_CODE") & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range

                        If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Then
                        Else
                        ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        End If
                    End If
                Else
               ' RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                End If
               End If
           
                  '''TILL HERE
                                If j = 0 Then
                                    MSFlexGrid1.TextMatrix(I, 2) = SELDPROCEDURRS("OBSERVALUE")
                                Else
                                    MSFlexGrid1.TextMatrix(I, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, j - 1)
                                End If
                                MSFlexGrid1.TextMatrix(I, 3) = RNAGE1
                                MSFlexGrid1.TextMatrix(I, 4) = SELDPROCEDURRS("pvalue")
                            Else
                                SELDTEMPLETRS.MoveNext
                                MSFlexGrid1.TextMatrix(I, 0) = I
                                MSFlexGrid1.TextMatrix(I, 1) = Trim(SELMA101RS("SER_DET")) & "$" & Trim(SELDPROCEDURRS("PRO_CODE"))
                                If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
                                    MSFlexGrid1.row = srno
                                    MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                                Else
                                    If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
                                        MSFlexGrid1.col = 2
                                        MSFlexGrid1.CellBackColor = RGB(150, 255, 100)
                                    End If
                                    If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
                                        MSFlexGrid1.col = 2
                                        MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                    If Trim(SELDPROCEDURRS("obser_type")) = "" Then
                                        MSFlexGrid1.col = 2
                                        MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                End If
                                j = 0
                                j = InStr(SELDPROCEDURRS("observalue"), "*")
                                If j = 0 Then
                                    MSFlexGrid1.TextMatrix(I, 2) = Trim(SELDPROCEDURRS("OBSERVALUE"))
                                Else
                                    MSFlexGrid1.TextMatrix(I, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, j - 1)
                                End If
                                
                                'Next line added by Shraddha Pandit on 18/12/08
                                'SELDTEMPLETQR(0) = Trim(SELDPROCEDURRS("PRO_CODE"))
                                'Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                                'Commented and added by Shraddha on 12/09/11
                                
                                'If condition commented by Shraddha on 29/09/2011
                                'If Trim(SELDTEMPLETRS("LOWVALUE")) = 0 Then
                                    'MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("HIGHVALUE"))
                                'Else
                                
                              '*********  ''''''''
                                RNAGE1 = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                ''''
                
                'ADDED BY SHWETA FOR NEW REFRENCE RANGE 01-10-2013
                 Set RS_VWREFRNG = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & SELDPROCEDURRS("PRO_CODE") & "' ", 1, 1)
        
                
                If Not RS_VWREFRNG.EOF Then
                    ' Dim agecal As Double
                agecal = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                txt_age.Text = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        If Trim(txt_age.Text) <= 1 Then
'                       ' Dim age3 As Double
'                         age3 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        age3 = "." & age3 '(age Mod 12)
'                        If age3 = ".0" Then
'                            age3 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                            age3 = ".0" & (age3 Mod 30)
'                        End If
'                        txt_age.Text = Trim(age3)
'                    End If
'                           ' ALKAGEO 'FOR F876
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & SELDPROCEDURRS("PRO_CODE") & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range

                        If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Then
                        Else
                        ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        End If
                    End If
                Else
               ' RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                End If
               
           
                  '''TILL HERE
                                
                                
                              '**********  ''''''''
                                
                                    MSFlexGrid1.TextMatrix(I, 3) = RNAGE1   ' Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                'End If
                                MSFlexGrid1.TextMatrix(I, 4) = SELDPROCEDURRS("pvalue")
                            End If
                        End If
                           srno = srno + 1
                            I = I + 1

                    End If
                    SELDPROCEDURRS.MoveNext

                Wend
'                MSFlexGrid1.Row = 1
'                MSFlexGrid1.Col = 2
'                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                txt_observed.Visible = False

           End If
        
           If SELDMBCCOMMFLAG = False Then
                SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
                Set SELDMBCCOMMRS = SELDMBCCOMMQR.OpenResultset(1, 3)
                SELDMBCCOMMFLAG = True
            Else
                SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
                SELDMBCCOMMRS.MoreResults
                SELDMBCCOMMRS.Requery
            End If
            If SELDMBCCOMMRS.RowCount > 0 Then
                If SELDMBCCOMMRS.EOF = False Then
                    TXT_COMMENTS.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
                End If
            End If
            '**** filling the flex grid
            
            
                   
            '*******
           
        End If
End Sub
Public Sub ALKAGE()
If txt_age.Text > 17 Then
RNAGE = 100
End If
SS = Trim(txt_sex.Text)
'If txt_age.Text <= 0.6 Then
'RNAGE = 0.6
'SS = ""
'End If
'If txt_age.Text <= 6 And txt_age.Text >= 4 Then
'RNAGE = 6
''SS = ""
'End If
If txt_age.Text <= 1 And txt_age.Text > 0.7 Then
RNAGE = 1
'SS = ""
End If
'If txt_age.Text <= 3 And txt_age.Text > 1 Then
'RNAGE = 3
''SS = ""
'End If
If (txt_age.Text <= 3 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "M" Then
RNAGE = 3
SS = "M"
End If
If (txt_age.Text <= 3 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "F" Then
RNAGE = 3
SS = "F"
End If
If (txt_age.Text <= 9 And txt_age.Text >= 7) And Trim(txt_sex.Text) = "M" Then
RNAGE = 9
SS = "M"
End If
If (txt_age.Text <= 9 And txt_age.Text >= 7) And Trim(txt_sex.Text) = "F" Then
RNAGE = 9
SS = "F"
End If
If (txt_age.Text <= 6 And txt_age.Text >= 4) And Trim(txt_sex.Text) = "M" Then
RNAGE = 6
SS = "M"
End If
If (txt_age.Text <= 6 And txt_age.Text >= 4) And Trim(txt_sex.Text) = "F" Then
RNAGE = 6
SS = "F"
End If
If (txt_age.Text <= 12 And txt_age.Text >= 10) And Trim(txt_sex.Text) = "M" Then
RNAGE = 12
SS = "M"
End If
If (txt_age.Text <= 12 And txt_age.Text >= 10) And Trim(txt_sex.Text) = "F" Then
RNAGE = 12
SS = "F"
End If
If (txt_age.Text <= 15 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "M" Then
RNAGE = 15
SS = "M"
End If
If (txt_age.Text <= 15 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "F" Then
RNAGE = 15
SS = "F"
End If
If (txt_age.Text <= 17 And txt_age.Text >= 16) And Trim(txt_sex.Text) = "M" Then
RNAGE = 17
SS = "M"
End If
If (txt_age.Text <= 17 And txt_age.Text >= 16) And Trim(txt_sex.Text) = "F" Then
RNAGE = 17
SS = "F"
End If
End Sub
'If (txt_age.Text <= 12 And txt_age.Text >= 7) And Trim(txt_sex.Text) = "F" Then
'RNAGE = 12
'SS = "F"
'End If
'If (txt_age.Text <= 17 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "M" Then
'RNAGE = 17
'SS = "M"
'End If
'If (txt_age.Text <= 17 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "F" Then
'RNAGE = 17
'SS = "F"
'End If
'End Sub

'Public Sub ALKAGE()
'If txt_Age.Text > 17 Then
'RNAGE = 100
'SS = Trim(txt_sex.Text)
'End If
'
''If txt_age.Text <= 0.6 Then
''RNAGE = 0.6
''SS = ""
''End If
'If txt_Age.Text <= 0.3 And txt_Age.Text >= 0.01 Then
'RNAGE = 0.3
'SS = Trim(txt_sex.Text)
'End If
'
'If txt_Age.Text <= 1 And txt_Age.Text >= 0.3 Then
'RNAGE = 1
'SS = Trim(txt_sex.Text)
'End If
'
'If txt_Age.Text <= 3 And txt_Age.Text >= 1 Then
'RNAGE = 3
'SS = Trim(txt_sex.Text)
'End If
'
'If txt_Age.Text <= 6 And txt_Age.Text >= 4 Then
'RNAGE = 6
'SS = Trim(txt_sex.Text)
'End If
'
'If txt_Age.Text <= 9 And txt_Age.Text >= 7 Then
'RNAGE = 9
'SS = Trim(txt_sex.Text)
'End If
'
'If txt_Age.Text <= 12 And txt_Age.Text >= 10 Then
'RNAGE = 12
'SS = Trim(txt_sex.Text)
'End If
'
'If txt_Age.Text <= 15 And txt_Age.Text >= 13 Then
'RNAGE = 15
'SS = Trim(txt_sex.Text)
'End If
'
'If txt_Age.Text <= 17 And txt_Age.Text >= 16 Then
'RNAGE = 17
'SS = Trim(txt_sex.Text)
'End If
'
'''If txt_age.Text <= 6 And txt_age.Text >= 4 Then
'''RNAGE = 6
'''SS = ""
'''End If
'''If txt_age.Text <= 1 And txt_age.Text > 0.7 Then
'''RNAGE = 1
'''SS = ""
'''End If
'''If txt_age.Text <= 3 And txt_age.Text > 1 Then
'''RNAGE = 3
'''SS = ""
'''End If
'''If (txt_age.Text <= 12 And txt_age.Text >= 7) And Trim(txt_sex.Text) = "M" Then
'''RNAGE = 12
'''SS = "M"
'''End If
'''If (txt_age.Text <= 12 And txt_age.Text >= 7) And Trim(txt_sex.Text) = "F" Then
'''RNAGE = 12
'''SS = "F"
'''End If
'''If (txt_age.Text <= 17 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "M" Then
'''RNAGE = 17
'''SS = "M"
'''End If
'''If (txt_age.Text <= 17 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "F" Then
'''RNAGE = 17
'''SS = "F"
'''End If
'''
''''JUSTIN 18-08-2004
'''If (txt_age.Text <= 99 And txt_age.Text >= 17) And Trim(txt_sex.Text) = "F" Then
'''RNAGE = 17
'''SS = "F"
'''End If
'''
'''If (txt_age.Text <= 99 And txt_age.Text >= 17) And Trim(txt_sex.Text) = "M" Then
'''RNAGE = 17
'''SS = "M"
'''End If
'''
'
'End Sub


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
    cmd_adenum.Visible = False
    'cmd_fullreport.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    If USRNAME_RS("MODIFY") = "YES" Then
         cmd_modify.Enabled = True
    Else
        cmd_modify.Enabled = False
    End If
    If LOGINTYP >= 6 Then
             cmd_modify.Enabled = True
    End If
    
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    cmd_adenum.Visible = False
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
    cmd_adenum.Visible = False
    'cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = True
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
    cmd_adenum.Visible = False
    adenum_flg = False
    lbl_addendnum.Visible = False
    txt_addendum.Visible = False
    txt_addendum.Enabled = False
    Label4.Visible = False
    cmb_addendumby.Visible = False
    
    
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
    cmd_adenum.Visible = False
    'cmd_fullreport.Enabled = False
ElseIf cap = "print" Then
     cmd_save.Enabled = False
     cmd_back.Enabled = True
     cmd_modify.Enabled = False
     cmd_print.Enabled = False
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False
     cmd_adenum.Visible = False
End If
End Sub



' ADD BY AMIT FOR DISPLAY PREVAL IN FLEX GRID

Public Function getpreval(inuptstr As String) As String
Dim INSTR As String
Dim I As Integer
Dim j As Integer
If Len(Trim(inuptstr)) > 0 Then
   I = InStr(1, inuptstr, ",")
  If I = 0 Then
     getpreval = Trim(inuptstr)
  ' Exit Sub
  Else
  j = InStr(I + 1, inuptstr, ",")
    If j = 0 Then
       getpreval = Trim(inuptstr)
   '    Exit Sub
       
    Else
        getpreval = Mid(inuptstr, 1, j - 1)
    '    Exit Sub
        
    End If
  End If
Else

getpreval = ""

End If


End Function

Public Sub ALKPHS1AGE()

SS = Trim(txt_sex.Text)


If (txt_age.Text <= 0.6 And txt_age.Text >= 0.01) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 0.6
    SS = "M"
End If

If (txt_age.Text <= 0.6 And txt_age.Text >= 0.01) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 0.6
    SS = "F"
End If

If (txt_age.Text <= 1 And txt_age.Text >= 0.6) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 1
    SS = "M"
End If

If (txt_age.Text <= 1 And txt_age.Text >= 0.6) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 0.6
    SS = "F"
End If

'***** Shraddha 14/09/2011 *****
If (txt_age.Text <= 12 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 12
    SS = "M"
End If

If (txt_age.Text <= 12 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 12
    SS = "F"
End If

If (txt_age.Text <= 18 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 18
    SS = "M"
End If

If (txt_age.Text <= 18 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 18
    SS = "F"
End If
'********************************

If (txt_age.Text <= 100 And txt_age.Text >= 18) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 100
    SS = "M"
End If

If (txt_age.Text <= 100 And txt_age.Text >= 18) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 100
    SS = "F"
End If


End Sub
Sub ALKAGEO()
SS = Trim(txt_sex.Text)


If (txt_age.Text <= 0.04 And txt_age.Text >= 0.01) And Trim(txt_sex.Text) = "M" Then
    RNAGEAL2 = 0.004
    SS = "M"
End If

If (txt_age.Text <= 0.04 And txt_age.Text >= 0.01) And Trim(txt_sex.Text) = "F" Then
    RNAGEAL2 = 0.004
    SS = "F"
End If


If (txt_age.Text < 0.11 And txt_age.Text >= 0.05) And Trim(txt_sex.Text) = "M" Then
    RNAGEAL2 = 0.011
    SS = "M"
End If

If (txt_age.Text < 0.11 And txt_age.Text >= 0.05) And Trim(txt_sex.Text) = "F" Then
    RNAGEAL2 = 0.011
    SS = "F"
End If

If (txt_age.Text <= 2 And txt_age.Text >= 0.11) And Trim(txt_sex.Text) = "M" Then
    RNAGEAL2 = 2
    SS = "M"
End If


If (txt_age.Text <= 2 And txt_age.Text >= 0.11) And Trim(txt_sex.Text) = "F" Then
    RNAGEAL2 = 2
    SS = "F"
End If

If (txt_age.Text <= 12 And txt_age.Text >= 2) And Trim(txt_sex.Text) = "M" Then
    RNAGEAL2 = 12
    SS = "M"
End If

If (txt_age.Text <= 12 And txt_age.Text >= 2) And Trim(txt_sex.Text) = "F" Then
    RNAGEAL2 = 12
    SS = "F"
End If


If (txt_age.Text <= 60 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "M" Then
    RNAGEAL2 = 60
    SS = "M"
End If

If (txt_age.Text <= 60 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "F" Then
    RNAGEAL2 = 60
    SS = "F"
End If


If (txt_age.Text <= 150 And txt_age.Text >= 60) And Trim(txt_sex.Text) = "M" Then
    RNAGEAL2 = 150
    SS = "M"
End If

If (txt_age.Text <= 150 And txt_age.Text >= 60) And Trim(txt_sex.Text) = "F" Then
    RNAGEAL2 = 150
    SS = "F"
End If





End Sub

Public Sub GTTAGE()

SS = Trim(txt_sex.Text)


If (txt_age.Text <= 0.6 And txt_age.Text >= 0.01) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 0.6
    SS = "M"
End If

If (txt_age.Text <= 0.6 And txt_age.Text >= 0.01) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 0.6
    SS = "F"
End If

If (txt_age.Text <= 1 And txt_age.Text >= 0.6) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 1
    SS = "M"
End If

If (txt_age.Text <= 1 And txt_age.Text >= 0.6) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 0.6
    SS = "F"
End If

'***** Shraddha 14/09/2011 *****
If (txt_age.Text <= 12 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 12
    SS = "M"
End If

If (txt_age.Text <= 12 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 12
    SS = "F"
End If

If (txt_age.Text <= 18 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 18
    SS = "M"
End If

If (txt_age.Text <= 18 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 18
    SS = "F"
End If
'********************************

If (txt_age.Text <= 100 And txt_age.Text >= 18) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 100
    SS = "M"
End If

If (txt_age.Text <= 100 And txt_age.Text >= 18) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 100
    SS = "F"
End If


End Sub

Public Sub STONE(cap As String)
''''SUMOL
If cap = "NEW" Then

MSFlexGrid1.Visible = False
STONEFRM.Visible = True
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False

ElseIf cap = "VIEW" Then
MSFlexGrid1.Visible = False
STONEFRM.Visible = True
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False

ElseIf cap = "CANCEL" Then
MSFlexGrid1.Visible = True
STONEFRM.Visible = False
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True

End If
'SUMOL
End Sub

Public Sub assignstone()
On Error GoTo MSG:
rdb.BeginTrans
'''sumol nov 2006 starts here
''' Dim NU As Integer
''' Dim num As Integer
''' Dim sc As String

 
 
    If flagsave = 1 Then
'''''''''''sumol nov 2006 starts here
'''''''''''
'''''''''''
'''''''''''
'''''''''''
'''''''''''
'''''''''''
'''''''''''         For i = 1 To MSFlexGrid1.Rows - 1
'''''''''''
''''''''''''JUSTINE 3-12-2004
'''''''''''
'''''''''''         num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
'''''''''''         sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
'''''''''''
'''''''''''''            If SELDMA101FLAG = False Then
'''''''''''''                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 1))
'''''''''''''                Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
'''''''''''''                SELDMA101FLAG = True
'''''''''''''            Else
'''''''''''''                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 1))
'''''''''''''                SELDMA101RS.MoreResults
'''''''''''''                SELDMA101RS.Requery
'''''''''''''            End If
'''''''''''
'''''''''''''            If SELDMA101RS.RowCount > 0 Then
'''''''''''
'''''''''''             If Trim(sc) <> "" Then
'''''''''''                Dim t As String
'''''''''''                t = VAR_REQNNO
'''''''''''                INSDMPROCEDURQR(0) = t
'''''''''''                INSDMPROCEDURQR(1) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE")) 'trim(sc)
'''''''''''                'INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                If SELDTEMPLET_MFLAG = False Then
'''''''''''                    SELDTEMPLET_MQR(0) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
'''''''''''                    'SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
'''''''''''                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
'''''''''''                    Set SELDTEMPLETRS = SELDTEMPLET_MQR.OpenResultset(1, 3)
'''''''''''                    SELDTEMPLET_MFLAG = True
'''''''''''                Else
'''''''''''                    SELDTEMPLET_MQR(0) = Trim(sc) ' Trim(SELDMA101RS("SER_CODE"))
'''''''''''                    'SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
'''''''''''                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
'''''''''''                    SELDTEMPLETRS.MoreResults
'''''''''''                    SELDTEMPLETRS.Requery
'''''''''''                End If
'''''''''''
'''''''''''  ' 03-12-2004 JUSTINE
'''''''''''                '*****
'''''''''''                If Trim(sc) = "F869" Then
'''''''''''                'If Trim(SELDMA101RS("ser_code")) = "F869" Then
'''''''''''
'''''''''''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(ALKHV) Then
'''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= 500 Then
'''''''''''                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
'''''''''''                                    'INSDMPROCEDURQR(3) = "H"
'''''''''''
'''''''''''                                Else
'''''''''''                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                End If
'''''''''''                                INSDMPROCEDURQR(3) = "H"
'''''''''''
'''''''''''                            Else
'''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(ALKLV) Then
'''''''''''                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                    INSDMPROCEDURQR(3) = "N"
'''''''''''
'''''''''''                                End If
'''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(ALKLV) Then
'''''''''''                                    If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(ALKLV) Then
'''''''''''                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
'''''''''''                                    Else
'''''''''''                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                     End If
'''''''''''                                        INSDMPROCEDURQR(3) = "L"
'''''''''''
'''''''''''                                End If
'''''''''''                                If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
'''''''''''                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                    INSDMPROCEDURQR(3) = ""
'''''''''''
'''''''''''                                End If
'''''''''''                            End If
'''''''''''
'''''''''''                Else
'''''''''''
'''''''''''                        If SELDTEMPLETRS.RowCount > 0 Then
'''''''''''                           'JUSTIN 05-10-2004
'''''''''''                         NU = InStr(Trim(MSFlexGrid1.TextMatrix(i, 3)), "-")
'''''''''''                         If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), NU + 1)) Then
'''''''''''
'''''''''''                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
'''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= (Val(Trim(SELDTEMPLETRS("LOWPANIC") = 0)) And Val(Trim(SELDTEMPLETRS("HIGHPANIC") = 0))) Then
'''''''''''                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                Else
'''''''''''                                    If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
'''''''''''                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
'''''''''''                                    Else
'''''''''''                                         INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                    End If
'''''''''''                                End If
'''''''''''                                INSDMPROCEDURQR(3) = "H"
'''''''''''
'''''''''''                            Else
'''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), 1, NU + 1)) Then ' Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'''''''''''                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                    INSDMPROCEDURQR(3) = "N"
'''''''''''
'''''''''''                                End If
'''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), 1, NU + 1)) Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'''''''''''                                    If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
'''''''''''                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
'''''''''''                                    Else
'''''''''''                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                     End If
'''''''''''                                        INSDMPROCEDURQR(3) = "L"
'''''''''''
'''''''''''                                End If
'''''''''''                                If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
'''''''''''                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''''''''''                                    INSDMPROCEDURQR(3) = ""
'''''''''''
'''''''''''                                End If
'''''''''''                            End If
'''''''''''                        End If
'''''''''''                End If
'''''''''''                '*****
'''''''''''
'''''''''''
'''''''''''                INSDMPROCEDURQR(4) = Trim(SELDTEMPLETRS(3))
'''''''''''
'''''''''''                INSDMPROCEDURQR(5) = Trim(MSFlexGrid1.TextMatrix(i, 3))
'''''''''''                INSDMPROCEDURQR(6) = Trim(MSFlexGrid1.TextMatrix(i, 4))
'''''''''''                INSDMPROCEDURQR(7) = Trim(MSFlexGrid1.TextMatrix(i, 5))
'''''''''''                INSDMPROCEDURQR(8) = Trim(MSFlexGrid1.TextMatrix(i, 7))
'''''''''''                If dm = True Then
'''''''''''                    INSDMPROCEDURQR(3) = "N"
'''''''''''                End If
'''''''''''
'''''''''''                INSDMPROCEDURQR.Execute
'''''''''''                'UPDATE DATE AND TIME IN RTIMELOG
'''''''''''
'''''''''''            End If
'''''''''''
'''''''''''
'''''''''''         Next i
'''''''''''         dm = False

            'SUMOL NOV 2006 STARTS HERE FOR STONE REPORT(STONERPT)
            INSSTONEQR(0) = VAR_REQNNO
            INSSTONEQR(1) = "F947" 'STONE ANALYSIS CODE F947
            INSSTONEQR(2) = Trim(txt_pathno.Text)
            INSSTONEQR(3) = Trim(MaskEdBox1.FormattedText)
            INSSTONEQR(4) = "NA"
            INSSTONEQR(5) = Trim(STONETXT.Text)
            INSSTONEQR.Execute
            
            'ENDS HERE
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(VAR_REQNNO)
            updrrequisitn_qr.Execute
            
            'Added by Milind 19/01/06
            Dim one As String
            Dim two As String
            Dim three As String
            Dim four As String
            Dim five As String
            one = Mid(VAR_REQNNO, 1, 3)
            two = Mid(VAR_REQNNO, 4, 2)
            three = Mid(VAR_REQNNO, 6, 2)
            four = Mid(VAR_REQNNO, 8, 6)
            five = one + two + three + four
            
            INSDMBCCOMMQR(0) = Trim(five)
            'upto here
            'INSDMBCCOMMQR(0) = Trim(req_chglen(VAR_REQNNO))
            
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                INSDMBCCOMMQR(1) = Null
                INSDMBCCOMMQR(1) = TXT_COMMENTS.Text
            Else
                INSDMBCCOMMQR(1) = Null
            
            End If
            INSDMBCCOMMQR.Execute
            
        
    
    End If
    If flagsave = 2 Then
''''''sumol nov 2006 STONE starts here
''''''            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
''''''            For i = 1 To MSFlexGrid1.Rows - 1
''''''             num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
''''''             sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
''''''end here
         
        'JUSTINE 03-12-04
''            If SELDMA101FLAG = False Then
''                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 1))
''                Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
''                SELDMA101FLAG = True
''            Else
''                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 1))
''                SELDMA101RS.MoreResults
''                SELDMA101RS.Requery
''            End If
            
            
''            If SELDMA101RS.RowCount > 0 Then
''
                'UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                
                
'''''''''''''''''sumol nov 2006 STONE statrs here
                
''''''''''''''''             If Trim(sc) <> "" Then
''''''''''''''''                If SELDTEMPLET_MFLAG = False Then
''''''''''''''''                    SELDTEMPLET_MQR(0) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
''''''''''''''''                   ' SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
''''''''''''''''                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
''''''''''''''''                    Set SELDTEMPLETRS = SELDTEMPLET_MQR.OpenResultset(1, 3)
''''''''''''''''                    SELDTEMPLET_MFLAG = True
''''''''''''''''                Else
''''''''''''''''                    SELDTEMPLET_MQR(0) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
''''''''''''''''                    'SELDTEMPLET_MQR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
''''''''''''''''                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
''''''''''''''''                    SELDTEMPLETRS.MoreResults
''''''''''''''''                    SELDTEMPLETRS.Requery
''''''''''''''''                End If
''''''''''''''''                If SELDTEMPLETRS.RowCount > 0 Then
''''''''''''''''                    '****
''''''''''''''''                    'If Trim(SELDMA101RS("SER_CODE")) = "F869" Then
''''''''''''''''                    If Trim(sc) = "F869" Then
''''''''''''''''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(ALKHV) Then
''''''''''''''''
''''''''''''''''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= 500 Then
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''''''''''''''''
''''''''''''''''                            Else
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                            End If
''''''''''''''''
''''''''''''''''                            UPDDMPROCEDURQR(1) = "H"
''''''''''''''''                         Else
''''''''''''''''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(ALKLV) Then
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                                UPDDMPROCEDURQR(1) = "N"
''''''''''''''''                            End If
''''''''''''''''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(ALKLV) Then
''''''''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(ALKLV) Then
''''''''''''''''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''''''''''''''''                                Else
''''''''''''''''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                                End If
''''''''''''''''                                UPDDMPROCEDURQR(1) = "L"
''''''''''''''''                            End If
''''''''''''''''                            If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                                UPDDMPROCEDURQR(1) = ""
''''''''''''''''                            End If
''''''''''''''''                        End If
''''''''''''''''
''''''''''''''''                    Else
''''''''''''''''                      'JUSTIN 05-10-2004
''''''''''''''''                         NU = InStr(Trim(MSFlexGrid1.TextMatrix(i, 3)), "-")
''''''''''''''''                         If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), NU + 1)) Then
''''''''''''''''                        'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
''''''''''''''''
''''''''''''''''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''''''''''''''''
''''''''''''''''                            Else
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                            End If
''''''''''''''''
''''''''''''''''                            UPDDMPROCEDURQR(1) = "H"
''''''''''''''''                         Else
''''''''''''''''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), 1, NU + 1)) Then ' Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                                UPDDMPROCEDURQR(1) = "N"
''''''''''''''''                            End If
''''''''''''''''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), 1, NU + 1)) Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''''''''''''''''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
''''''''''''''''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''''''''''''''''                                Else
''''''''''''''''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                                End If
''''''''''''''''                                UPDDMPROCEDURQR(1) = "L"
''''''''''''''''                            End If
''''''''''''''''                            If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
''''''''''''''''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''''''''''''''''                                UPDDMPROCEDURQR(1) = ""
''''''''''''''''                            End If
''''''''''''''''                        End If
''''''''''''''''
''''''''''''''''                    End If
                    '****
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
''
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''
''                            Else
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                            End If
''
''                            UPDDMPROCEDURQR(1) = "H"
''                         Else
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                                UPDDMPROCEDURQR(1) = "N"
''                            End If
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''                                Else
''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                                End If
''                                UPDDMPROCEDURQR(1) = "L"
''                            End If
''                            If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                                UPDDMPROCEDURQR(1) = ""
''                            End If
''                        End If

'''SUMOL NOV 2006 STONE STARTS HERE
'''sumol update dmprodeure not required for word (NON interface) report

''                End If
                
'''sumol update dmprodeure not required for word (NON interface) report
''                 UPDDMPROCEDURQR(2) = Trim(SELDTEMPLETRS("UOM"))
''                 UPDDMPROCEDURQR(3) = Trim(MSFlexGrid1.TextMatrix(i, 3))
''                 UPDDMPROCEDURQR(4) = Trim(MSFlexGrid1.TextMatrix(i, 4))
''                 UPDDMPROCEDURQR(5) = req_chglen(cmb_reqno.Text)
''                 UPDDMPROCEDURQR(6) = Trim(sc) 'Trim(SELDMA101RS("SER_CODE"))
''                UPDDMPROCEDURQR.Execute
''            End If
''
''
''         Next i


            'SUMOL NOV 2006 STARTS HERE FOR STONE REPORT(STONERPT)
            
            
            UPDSTONEQR(0) = VAR_REQNNO
            UPDSTONEQR(1) = "F947"
            UPDSTONEQR(2) = STONETXT.Text
            UPDSTONEQR(3) = "NA"
            UPDSTONEQR(4) = VAR_REQNNO
            If (reqspn_rs("NMETERIALS") = "Stone Analysis") Then
            UPDSTONEQR(5) = "F947"
            Else
            UPDSTONEQR(5) = "NOMATCH"
            End If
            UPDSTONEQR.Execute
            
            'REQNO=?,PRO_CODE=?,TXT=?,MACNAME=? WHERE REQNO=? AND PRO_CODE=?
            'ENDS HERE
                
                
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
           
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                UPDDMBCCOMMQR(0) = Null
                UPDDMBCCOMMQR(0) = Trim(TXT_COMMENTS.Text)
            Else
                UPDDMBCCOMMQR(0) = Null
                
            End If
            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
            UPDDMBCCOMMQR.Execute
            
            
    End If
        If adenum_flg = True Then
           UPDRTIMECADD_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
           UPDRTIMECADD_QR(1) = time1
           UPDRTIMECADD_QR(2) = req_chglen(cmb_reqno.Text)
           UPDRTIMECADD_QR.Execute
        
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
            INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
            INSDADDENDUM_QR.Execute
         Else
            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            INSDADDENDUM_QR(1) = "1"
            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INSDADDENDUM_QR(3) = time1
            INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
            INSDADDENDUM_QR.Execute
         End If
       End If
        
    
rdb.CommitTrans
MsgBox " Record is Saved Sucessfully "
'btnenabl ("INI")

BTNENABL1 ("INI")
txt_observed.Visible = False
cmb_machine.Visible = False
Exit Sub
MSG:
rdb.RollbackTrans

MsgBox Err.Description

End Sub

Public Sub ASSIGNNEW()
On Error GoTo MSG:

Dim count, low, high, W1
rdb.BeginTrans
    If flagsave = 1 Then
         For I = 1 To MSFlexGrid1.Rows - 1
            If SELDMA101FLAG = False Then
                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(I, 1))
                Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
                SELDMA101FLAG = True
            Else
                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(I, 1))
                SELDMA101RS.MoreResults
                SELDMA101RS.Requery
            End If
            If SELDMA101RS.RowCount > 0 Then
                INSDMPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
                INSDMPROCEDURQR(1) = Trim(SELDMA101RS("SER_CODE"))
                'INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                '***** Shraddha 20/07/10 *****
'                INSDMPROCEZ_QR(0) = req_chglen(cmb_reqno.Text)
'                INSDMPROCEZ_QR(1) = Trim(SELDMA101RS("SER_CODE"))
                '*****************************
                If SELDTEMPLETFLAG = False Then
                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                    SELDTEMPLETFLAG = True
                Else
                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    SELDTEMPLETRS.MoreResults
                    SELDTEMPLETRS.Requery
                End If
                '*****
                If Trim(SELDMA101RS("ser_code")) = "F869" Then

                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(ALKHV) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= 1000 Then
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '*****************************
                                Else
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '*****************************
                                End If
                                INSDMPROCEDURQR(3) = "H"
                                '***** Shraddha 20/07/10 *****
'                                INSDMPROCEZ_QR(3) = "H"
                                '*****************************
                            Else
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(ALKLV) Then
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    INSDMPROCEDURQR(3) = "N"
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                                    INSDMPROCEZ_QR(3) = "N"
                                    '*****************************
                                End If
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '*****************************
                                    Else
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '*****************************
                                    End If
                                        INSDMPROCEDURQR(3) = "L"
                                        '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(3) = "L"
                                        '*****************************
                                End If
                                If (Trim(MSFlexGrid1.TextMatrix(I, 2))) = "" Then
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    INSDMPROCEDURQR(3) = ""
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                                    INSDMPROCEZ_QR(3) = ""
                                    '*****************************
                                End If
                            End If

                Else
                        If SELDTEMPLETRS.RowCount > 0 Then
                         
                         count = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
                         low = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1, count - 1)
                         high = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), count + 1)
                          W1 = "0"
                             
                            If (low <> "0" And high <> "0") And Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > high Or ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE") And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > high Then 'Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
                                    If ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE") And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                                     INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                     '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '*****************************
                                    Else
                                     INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                     '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                    '*****************************
                                    End If
                                Else
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    '*****************************
                                End If
                                INSDMPROCEDURQR(3) = "H"
                                '***** Shraddha 20/07/10 *****
'                                INSDMPROCEZ_QR(3) = "H"
                                '*****************************
                                W1 = "1"
                            Else
                               ' for normal value
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= low Or (UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "ABSENT" And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Or (low = "0" And high = "0") Then
                                'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= low Then ' Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    INSDMPROCEDURQR(3) = "N"
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                                    INSDMPROCEZ_QR(3) = "N"
                                '*****************************
                                    W1 = "1"
                                End If
                                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < low And W1 <> "1" Then ' Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                        '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                        '*****************************
                                    Else
                                        INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '***** Shraddha 20/07/10 *****
'                                        INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                        '*****************************
                                    End If
                                    If low = "UPTO" Then
                                        INSDMPROCEDURQR(3) = "N"
                                    Else
                                        INSDMPROCEDURQR(3) = "L"
                                    End If
                                End If
                                If (Trim(MSFlexGrid1.TextMatrix(I, 2))) = "" Then
                                    INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    INSDMPROCEDURQR(3) = ""
                                    '***** Shraddha 20/07/10 *****
'                                    INSDMPROCEZ_QR(2) = Trim(MSFlexGrid1.TextMatrix(I, 2))
'                                    INSDMPROCEZ_QR(3) = ""
                                    '*****************************
                                End If
                            End If
                        End If
                End If
                '*****
'''                '******
'''                If SELDTEMPLETRS.RowCount > 0 Then
'''                    If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
'''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
'''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
'''
'''                        Else
'''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''                        End If
'''                        INSDMPROCEDURQR(3) = "H"
'''
'''                    Else
'''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''                            INSDMPROCEDURQR(3) = "N"
'''
'''                        End If
'''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
'''                                INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
'''                            Else
'''                                INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''                             End If
'''                                INSDMPROCEDURQR(3) = "L"
'''
'''                        End If
'''                        If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
'''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
'''                            INSDMPROCEDURQR(3) = ""
'''
'''                        End If
'''                    End If
'''                End If
'''                '******
                INSDMPROCEDURQR(4) = Trim(SELDTEMPLETRS("uom"))
                INSDMPROCEDURQR(5) = Trim(MSFlexGrid1.TextMatrix(I, 3))
                INSDMPROCEDURQR(6) = Trim(MSFlexGrid1.TextMatrix(I, 4))
                '***** Shraddha 22/07/10 *****
'                INSDMPROCEZ_QR(4) = Trim(SELDTEMPLETRS("uom"))
'                INSDMPROCEZ_QR(5) = Trim(MSFlexGrid1.TextMatrix(I, 3))
'                INSDMPROCEZ_QR(6) = Trim(MSFlexGrid1.TextMatrix(I, 4))
                '*****************************
                
                
                                
                                'added by milind on 18-05-2006 for inserting machine name
                                'according to procedure code
                                rqsqlinterface(0) = Trim(txt_pathno.Text)
                                rqsqlinterface(1) = Trim(SELDMA101RS("SER_CODE"))
                                Set Rssqlinterface = rqsqlinterface.OpenResultset(1, 3)
                                If Rssqlinterface.RowCount > 0 Then
                                   INTFMCN = Rssqlinterface("INSTRU_NAME")
                                Else
                                   INTFMCN = ""
                                End If
                                INSDMPROCEDURQR(7) = INTFMCN
                                INSDMPROCEZ_QR(7) = INTFMCN 'Shraddha 22/07/10
                                'upto here
                
                INSDMPROCEDURQR.Execute
            End If
            
    
         Next I
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
        
            INSDMBCCOMMQR(0) = Trim(req_chglen(cmb_reqno.Text))
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                INSDMBCCOMMQR(1) = Null
                INSDMBCCOMMQR(1) = TXT_COMMENTS.Text
            Else
                INSDMBCCOMMQR(1) = Null
            
            End If
            INSDMBCCOMMQR.Execute
            
        
    
    End If
    If flagsave = 2 Then
            For I = 1 To MSFlexGrid1.Rows - 1
         
            If SELDMA101FLAG = False Then
                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(I, 1))
                Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
                SELDMA101FLAG = True
            Else
                SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(I, 1))
                SELDMA101RS.MoreResults
                SELDMA101RS.Requery
            End If
            If SELDMA101RS.RowCount > 0 Then
'
                'UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                
                If SELDTEMPLETFLAG = False Then
                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                    SELDTEMPLETFLAG = True
                Else
                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
                    'SELDTEMPLETQR(1) = Trim(txt_sex.Text)
                    SELDTEMPLETRS.MoreResults
                    SELDTEMPLETRS.Requery
                End If
                If SELDTEMPLETRS.RowCount > 0 Then
                    '****
                    If Trim(SELDMA101RS("SER_CODE")) = "F869" Then
                        If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > Val(ALKHV) Then
                        
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= 500 Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                
                            Else
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                            End If
                           
                            UPDDMPROCEDURQR(1) = "H"
                         Else
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(ALKLV) Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = "N"
                            End If
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(ALKLV) Then
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                Else
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                End If
                                UPDDMPROCEDURQR(1) = "L"
                            End If
                            If Trim(MSFlexGrid1.TextMatrix(I, 2)) = "" Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = ""
                            End If
                        End If
                        
                    Else
                        
                         
                         count = InStr(Trim(MSFlexGrid1.TextMatrix(I, 3)), "-")
                         low = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), 1, count - 1)
                         high = Mid(Trim(MSFlexGrid1.TextMatrix(I, 3)), count + 1)
                             
                            If (low <> "0" And high <> "0") And Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) > high Or ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE") And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > high Then 'Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
                                    If ((UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "PRESENT++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "+++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "++++" Or UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "TRACE") And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Then
                                     UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    Else
                                     UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                    End If
                                Else
                                   UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                End If
                            UPDDMPROCEDURQR(1) = "H"
                         Else
                           'for normal value
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) >= low Or (UCase(Trim(MSFlexGrid1.TextMatrix(I, 2))) = "ABSENT" And (Trim(SELDMA101RS("SER_CODE")) = "F846" Or Trim(SELDMA101RS("SER_CODE")) = "F847")) Or (low = "0" And high = "0") Then
                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = "N"
                            End If
                            If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < low Then 'Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
                                    If Val(Trim(MSFlexGrid1.TextMatrix(I, 2))) <> "0" Then
                                     UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2)) & "*"
                                    Else
                                     UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                    End If
                                Else
                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                End If
                                If low = "UPTO" Then
                                    UPDDMPROCEDURQR(1) = "N"
                                Else
                                    UPDDMPROCEDURQR(1) = "L"
                                End If
                            End If
                            If Trim(MSFlexGrid1.TextMatrix(I, 2)) = "" Then
                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(I, 2))
                                UPDDMPROCEDURQR(1) = ""
                            End If
                        End If
                    
                    End If
                    '****
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
''
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''
''                            Else
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                            End If
''
''                            UPDDMPROCEDURQR(1) = "H"
''                         Else
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                                UPDDMPROCEDURQR(1) = "N"
''                            End If
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''                                Else
''                                    UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                                End If
''                                UPDDMPROCEDURQR(1) = "L"
''                            End If
''                            If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                                UPDDMPROCEDURQR(1) = ""
''                            End If
''                        End If
                End If
                 UPDDMPROCEDURQR(2) = Trim(SELDTEMPLETRS("UOM"))
                 UPDDMPROCEDURQR(3) = Trim(MSFlexGrid1.TextMatrix(I, 3))
                 UPDDMPROCEDURQR(4) = Trim(MSFlexGrid1.TextMatrix(I, 4))
                 UPDDMPROCEDURQR(5) = req_chglen(cmb_reqno.Text)
                 UPDDMPROCEDURQR(6) = Trim(SELDMA101RS("SER_CODE"))
                UPDDMPROCEDURQR.Execute
            End If
            
    
         Next I
        
            updrrequisitn_qr(0) = cmb_enterby.Text
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = cmb_authSign.Text     ' cmb_finalby.Text  'CHANGES DONE BY SHWETA 26-10-2018
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = cmb_finalby.Text
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
           
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                UPDDMBCCOMMQR(0) = Null
                UPDDMBCCOMMQR(0) = TXT_COMMENTS.Text
            Else
                UPDDMBCCOMMQR(0) = Null
                
            End If
            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
            UPDDMBCCOMMQR.Execute
            
            
    End If
    
    If adenum_flg = True Then
           UPDRTIMECADD_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
           UPDRTIMECADD_QR(1) = time1
           UPDRTIMECADD_QR(2) = req_chglen(cmb_reqno.Text)
           UPDRTIMECADD_QR.Execute
        
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
            INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
            INSDADDENDUM_QR.Execute
         Else
            INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            INSDADDENDUM_QR(1) = "1"
            INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            INSDADDENDUM_QR(3) = time1
            INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
            INSDADDENDUM_QR(5) = Null
            INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
            INSDADDENDUM_QR.Execute
         End If
    End If
        
    
    
rdb.CommitTrans
MsgBox "Record is Saved"
'btnenabl ("INI")
BTNENABL1 ("INI")
txt_observed.Visible = False
Exit Sub
MSG:
    rdb.RollbackTrans
MsgBox Err.Description

End Sub



Private Sub txt_pathno_LostFocus()
If chk_pno.Value = False Then
Exit Sub
End If
If Len(txt_pathno.Text) > 6 Then
Else
Exit Sub
End If
On Error Resume Next
If Trim(txt_pathno.Text) = "" Then
Exit Sub
End If


Lbl_pno.Caption = Trim(txt_pathno.Text)
Dim sqlstr As String
Dim SELADDINTERFACE_QR As rdoQuery
Dim SELADDINTERFACE_RS As rdoResultset

 sqlstr = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN (SELECT REQNO FROM DISLIB.DREQUISITN WHERE PATHNO=? AND REQNO LIKE 'FECBC%') "
  Set SELADDINTERFACE_QR = rdb.CreateQuery("", sqlstr)
 SELADDINTERFACE_QR(0) = Trim(txt_pathno.Text)
Set SELADDINTERFACE_RS = SELADDINTERFACE_QR.OpenResultset(1, 3)
If SELADDINTERFACE_RS.EOF = False Then
If Len(Trim(SELADDINTERFACE_RS("CASENO"))) = 6 Then
opt_staff.Value = True
MaskEdBox1.Text = Trim(SELADDINTERFACE_RS("CASENO"))
Else
MaskEdBox1.Text = Trim(SELADDINTERFACE_RS("CASENO"))
End If
'MsgBox Trim(SELADDINTERFACE_RS("CASENO"))
cmd_cancel.Enabled = True
'MsgBox ("CASENO")
MaskEdBox1_LostFocus
'MsgBox "LOST"
    If Trim(SELADDINTERFACE_RS("reqnstat")) = Null Or Trim(SELADDINTERFACE_RS("reqnstat")) = "T" Then
    'cmd_save.Enabled = True
    cmd_modify.Enabled = True
    ElseIf Trim(SELADDINTERFACE_RS("reqnstat")) = "F" Then
 cmd_modify.Enabled = False
        MsgBox ("Report is committed you can not modify the report!!!Addendum you can add")
    End If
    Else
    cmd_cancel_Click
    'MsgBox Trim("CANCEL")
    cmd_new_Click
    'MsgBox Trim("NEW")
End If
End Sub
