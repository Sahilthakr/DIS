VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{9B8DFCFF-6726-11D6-99D4-000244193CE4}#1.0#0"; "PROJECT1.OCX"
Begin VB.Form DTMREP 
   Caption         =   "DTMREPORT"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   11775
      Begin VB.OptionButton OPT_ALL 
         Caption         =   "All"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton OPT_CASENO 
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox txt_name 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6720
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   4560
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_todt 
         Height          =   375
         Left            =   9360
         TabIndex        =   21
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msk_fromdt 
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.Label lbl_fromdt 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8520
         TabIndex        =   27
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   26
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6120
         TabIndex        =   25
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3720
         TabIndex        =   24
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   11775
      Begin VB.TextBox txt_rowcount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   4680
         Width           =   9375
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   20
         Cols            =   11
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   192
         AllowUserResizing=   3
         FormatString    =   $"DTMREP.frx":0000
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
   End
   Begin VB.Frame buttons 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   9360
         MouseIcon       =   "DTMREP.frx":0095
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   7920
         MouseIcon       =   "DTMREP.frx":039F
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6480
         MouseIcon       =   "DTMREP.frx":06A9
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   5040
         MouseIcon       =   "DTMREP.frx":09B3
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3600
         MouseIcon       =   "DTMREP.frx":0CBD
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "DTMREP.frx":0FC7
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Height          =   375
         Left            =   720
         MouseIcon       =   "DTMREP.frx":12D1
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   510
      Left            =   2760
      TabIndex        =   28
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   900
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      Caption         =   "DTMLAB  REPORT"
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
      Left            =   4800
      TabIndex        =   29
      Top             =   600
      Width           =   2610
   End
End
Attribute VB_Name = "DTMREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_BACK_Click()
DIS_MAIN.show
Me.Hide
End Sub



