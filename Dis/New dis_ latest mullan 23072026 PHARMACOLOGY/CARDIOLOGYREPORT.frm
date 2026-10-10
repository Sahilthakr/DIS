VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form EChocardiologyreport 
   Caption         =   "Echo Cardiography and Colour Doppler Report"
   ClientHeight    =   9780
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12780
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      TabIndex        =   86
      Top             =   6840
      Width           =   11775
      Begin VB.ComboBox cmb_reportby 
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
         Left            =   4680
         TabIndex        =   55
         Top             =   600
         Width           =   2535
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
         Left            =   9000
         TabIndex        =   56
         Top             =   600
         Width           =   2415
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
         Left            =   1080
         TabIndex        =   54
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txt_remarks 
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
         Left            =   1080
         TabIndex        =   52
         Top             =   240
         Width           =   10335
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reported By"
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
         Left            =   3600
         TabIndex        =   140
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
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
         Left            =   7320
         TabIndex        =   89
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lbl_enter 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
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
         Left            =   240
         TabIndex        =   88
         Top             =   600
         Width           =   720
      End
      Begin VB.Label LBL_COMM 
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
         Left            =   240
         TabIndex        =   87
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame COMMAND 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   84
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
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
         Left            =   480
         MouseIcon       =   "CARDIOLOGYREPORT.frx":0000
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
         MouseIcon       =   "CARDIOLOGYREPORT.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   1
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
         Left            =   3240
         MouseIcon       =   "CARDIOLOGYREPORT.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   2
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
         Left            =   4560
         MouseIcon       =   "CARDIOLOGYREPORT.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   3
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
         Left            =   5880
         MouseIcon       =   "CARDIOLOGYREPORT.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   4
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
         Left            =   8760
         MouseIcon       =   "CARDIOLOGYREPORT.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   6
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
         Left            =   10200
         MouseIcon       =   "CARDIOLOGYREPORT.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   7
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
         Left            =   7200
         MouseIcon       =   "CARDIOLOGYREPORT.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   5
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
      TabIndex        =   81
      Top             =   240
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
      TabIndex        =   80
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame common 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   57
      Top             =   960
      Width           =   11775
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
         Sorted          =   -1  'True
         TabIndex        =   8
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
         TabIndex        =   62
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
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_microno 
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_provdiag 
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
         TabIndex        =   23
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
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmb_reportdoc 
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
         Height          =   315
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   0
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         TabIndex        =   58
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
            TabIndex        =   60
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
            TabIndex        =   59
            Top             =   600
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         PromptChar      =   "_"
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Echo No."
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
         TabIndex        =   75
         Top             =   600
         Width           =   810
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
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
         Left            =   480
         TabIndex        =   63
         Top             =   2520
         Width           =   1830
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   90
      Top             =   3960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5318
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Echocardiographic Measurement"
      TabPicture(0)   =   "CARDIOLOGYREPORT.frx":1850
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TXT_EPSS"
      Tab(0).Control(1)=   "txt_lvids"
      Tab(0).Control(2)=   "txt_lvidd"
      Tab(0).Control(3)=   "txt_ivsd"
      Tab(0).Control(4)=   "txt_rvidd"
      Tab(0).Control(5)=   "txt_lvpwd"
      Tab(0).Control(6)=   "txt_la"
      Tab(0).Control(7)=   "txt_ao"
      Tab(0).Control(8)=   "txt_ef"
      Tab(0).Control(9)=   "txt_fs"
      Tab(0).Control(10)=   "LBL_EPSS"
      Tab(0).Control(11)=   "lbl_lvids"
      Tab(0).Control(12)=   "LBL_LVIDD"
      Tab(0).Control(13)=   "LBL_IVSD"
      Tab(0).Control(14)=   "LBL_RVIDD"
      Tab(0).Control(15)=   "LBL_LVPWD"
      Tab(0).Control(16)=   "LBL_LA"
      Tab(0).Control(17)=   "lbl_ap"
      Tab(0).Control(18)=   "lbl_Ef"
      Tab(0).Control(19)=   "lbl_fs"
      Tab(0).Control(20)=   "Label3"
      Tab(0).Control(21)=   "Label4"
      Tab(0).Control(22)=   "Label5"
      Tab(0).Control(23)=   "Label6"
      Tab(0).Control(24)=   "Label7"
      Tab(0).Control(25)=   "Label8"
      Tab(0).Control(26)=   "Label9"
      Tab(0).Control(27)=   "Label10"
      Tab(0).Control(28)=   "Label11"
      Tab(0).Control(29)=   "lbl_dhemotherapy"
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "CARDIOLOGYREPORT.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmb_procdone"
      Tab(1).Control(1)=   "msk_procdate"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "TXT_HB"
      Tab(1).Control(4)=   "TXT_ECG"
      Tab(1).Control(5)=   "TXT_XRAYCHEST"
      Tab(1).Control(6)=   "txt_summary"
      Tab(1).Control(7)=   "txt_comments"
      Tab(1).Control(8)=   "lbl_star"
      Tab(1).Control(9)=   "lbl_procdone"
      Tab(1).Control(10)=   "lbl_examdt"
      Tab(1).Control(11)=   "lbl_hb"
      Tab(1).Control(12)=   "lbl_gms"
      Tab(1).Control(13)=   "lbl_ecg"
      Tab(1).Control(14)=   "lbl_xray"
      Tab(1).Control(15)=   "lbl_summary"
      Tab(1).Control(16)=   "lbl_comments"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Doppler study"
      TabPicture(2)   =   "CARDIOLOGYREPORT.frx":1888
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "msflex_doppler"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txt_1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txt_2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txt_3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.ComboBox cmb_procdone 
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
         Left            =   -66720
         Sorted          =   -1  'True
         TabIndex        =   138
         Top             =   2280
         Width           =   3135
      End
      Begin MSMask.MaskEdBox msk_procdate 
         Height          =   375
         Left            =   -66720
         TabIndex        =   137
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame6 
         Caption         =   "Valve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8760
         TabIndex        =   132
         Top             =   240
         Width           =   2895
         Begin VB.TextBox Txt_tricuspid 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1080
            TabIndex        =   49
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txt_pulmonary 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1080
            TabIndex        =   47
            Top             =   120
            Width           =   1695
         End
         Begin VB.ComboBox cmb_arotic 
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
            Left            =   1110
            TabIndex        =   48
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Pulmonary"
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
            TabIndex        =   135
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Aortic"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   480
            TabIndex        =   134
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tricuspid"
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
            TabIndex        =   133
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6480
         TabIndex        =   127
         Top             =   360
         Width           =   2295
         Begin VB.TextBox txt_pht 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   930
            TabIndex        =   46
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txt_planimetry 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   930
            TabIndex        =   45
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "cm2"
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
            Left            =   1890
            TabIndex        =   131
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "cm2"
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
            Left            =   1890
            TabIndex        =   130
            Top             =   360
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Planimetry"
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
            TabIndex        =   129
            Top             =   240
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label18 
            Caption         =   "M. V. A.  PHT"
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
            Height          =   435
            Left            =   240
            TabIndex        =   128
            Top             =   600
            Width           =   780
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mitral Valve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         TabIndex        =   122
         Top             =   360
         Width           =   6495
         Begin VB.ComboBox cmb_mitral 
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
            Left            =   4440
            TabIndex        =   44
            Top             =   600
            Width           =   1935
         End
         Begin VB.ComboBox cmb_subvalvular 
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
            Left            =   4440
            TabIndex        =   42
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox cmb_pml 
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
            Left            =   480
            TabIndex        =   43
            Top             =   600
            Width           =   1935
         End
         Begin VB.ComboBox cmb_aml 
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
            Left            =   480
            TabIndex        =   41
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "M. V. A. Planimetry"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2760
            TabIndex        =   126
            Top             =   720
            Width           =   1635
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Subvalvular apparatus"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2400
            TabIndex        =   125
            Top             =   360
            Width           =   1920
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "PML"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   124
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "AML"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.TextBox TXT_EPSS 
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
         Left            =   -71850
         TabIndex        =   33
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txt_lvids 
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
         Left            =   -67050
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_lvidd 
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
         Left            =   -71850
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_ivsd 
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
         Left            =   -71850
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_rvidd 
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
         Left            =   -71850
         TabIndex        =   28
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_lvpwd 
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
         Left            =   -67050
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_la 
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
         Left            =   -67050
         TabIndex        =   29
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_ao 
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
         Left            =   -71850
         TabIndex        =   30
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txt_ef 
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
         Left            =   -67050
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txt_fs 
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
         Left            =   -71850
         TabIndex        =   32
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "CHEMOTHERAPY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -66780
         TabIndex        =   93
         Top             =   480
         Width           =   2415
         Begin VB.TextBox txt_dose 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   285
            Left            =   720
            TabIndex        =   37
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TXT_DRUG 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   285
            Left            =   720
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label LBL_DOSE 
            AutoSize        =   -1  'True
            Caption         =   "Dose"
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
            TabIndex        =   96
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "mg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   95
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Drug"
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
            Left            =   150
            TabIndex        =   94
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.TextBox TXT_HB 
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
         Left            =   -73380
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TXT_ECG 
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
         Left            =   -73380
         TabIndex        =   35
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TXT_XRAYCHEST 
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
         Left            =   -73380
         TabIndex        =   38
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox txt_summary 
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
         Left            =   -73380
         TabIndex        =   40
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txt_comments 
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
         Left            =   -73380
         TabIndex        =   39
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox txt_3 
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
         TabIndex        =   53
         Top             =   1800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txt_2 
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
         Left            =   4560
         TabIndex        =   51
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txt_1 
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
         Left            =   1320
         TabIndex        =   50
         Top             =   1800
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSFlexGridLib.MSFlexGrid msflex_doppler 
         Height          =   1335
         Left            =   120
         TabIndex        =   92
         Top             =   1560
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   -2147483642
         ForeColorFixed  =   192
         Appearance      =   0
         FormatString    =   $"CARDIOLOGYREPORT.frx":18A4
      End
      Begin VB.Label lbl_star 
         Caption         =   "*"
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
         Height          =   255
         Left            =   -74640
         TabIndex        =   141
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label lbl_procdone 
         AutoSize        =   -1  'True
         Caption         =   "Performed By"
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
         Left            =   -67920
         TabIndex        =   139
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Label lbl_examdt 
         AutoSize        =   -1  'True
         Caption         =   "Exam Dt"
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
         Left            =   -67560
         TabIndex        =   136
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label LBL_EPSS 
         AutoSize        =   -1  'True
         Caption         =   "EPSS"
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
         Left            =   -72480
         TabIndex        =   121
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lbl_lvids 
         AutoSize        =   -1  'True
         Caption         =   "Left vent int dimension (LVIDs)"
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
         Left            =   -69810
         TabIndex        =   120
         Top             =   600
         Width           =   2640
      End
      Begin VB.Label LBL_LVIDD 
         AutoSize        =   -1  'True
         Caption         =   "Left vent int dimension (LVIDd)"
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
         Left            =   -74610
         TabIndex        =   119
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label LBL_IVSD 
         AutoSize        =   -1  'True
         Caption         =   "Inter ventricular septum (IVSD)"
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
         Left            =   -74595
         TabIndex        =   118
         Top             =   1080
         Width           =   2640
      End
      Begin VB.Label LBL_RVIDD 
         AutoSize        =   -1  'True
         Caption         =   "Right vent int dimension (RVIDd)"
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
         Left            =   -74760
         TabIndex        =   117
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label LBL_LVPWD 
         AutoSize        =   -1  'True
         Caption         =   "Left vent posterior wall (LVPWD)"
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
         Left            =   -69975
         TabIndex        =   116
         Top             =   960
         Width           =   2805
      End
      Begin VB.Label LBL_LA 
         AutoSize        =   -1  'True
         Caption         =   "Left atrium (LA)"
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
         Left            =   -68490
         TabIndex        =   115
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label lbl_ap 
         AutoSize        =   -1  'True
         Caption         =   "Aortic annulus (AO)"
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
         Left            =   -73620
         TabIndex        =   114
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label lbl_Ef 
         AutoSize        =   -1  'True
         Caption         =   "Ejection fraction (EF)"
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
         Left            =   -68985
         TabIndex        =   113
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lbl_fs 
         AutoSize        =   -1  'True
         Caption         =   "Fractional shortening (FS)"
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
         Left            =   -74160
         TabIndex        =   112
         Top             =   2160
         Width           =   2205
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   -65970
         TabIndex        =   111
         Top             =   1800
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   -70770
         TabIndex        =   110
         Top             =   2160
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   -70770
         TabIndex        =   109
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   -65970
         TabIndex        =   108
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   -70770
         TabIndex        =   107
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   -65970
         TabIndex        =   106
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   -70770
         TabIndex        =   105
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   -65970
         TabIndex        =   104
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   -70770
         TabIndex        =   103
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label lbl_hb 
         AutoSize        =   -1  'True
         Caption         =   "Hb"
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
         Left            =   -73770
         TabIndex        =   102
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lbl_gms 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "gms/dl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -72060
         TabIndex        =   101
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lbl_ecg 
         AutoSize        =   -1  'True
         Caption         =   "ECG"
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
         Left            =   -73905
         TabIndex        =   100
         Top             =   960
         Width           =   390
      End
      Begin VB.Label lbl_xray 
         AutoSize        =   -1  'True
         Caption         =   "X-RAY CHEST"
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
         Left            =   -74760
         TabIndex        =   99
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lbl_summary 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74430
         TabIndex        =   98
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label lbl_comments 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74220
         TabIndex        =   97
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label lbl_dhemotherapy 
         AutoSize        =   -1  'True
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
         TabIndex        =   91
         Top             =   480
         Width           =   75
      End
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   143
      Top             =   0
      Width           =   9015
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
      TabIndex        =   142
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl_cap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   " ECHOCARDIOGRAPHY REPORT"
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   3375
      TabIndex        =   85
      Top             =   480
      Width           =   4575
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
      TabIndex        =   83
      Top             =   240
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
      TabIndex        =   82
      Top             =   600
      Width           =   420
   End
End
Attribute VB_Name = "EChocardiologyreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VAR_REQNNO As String
Dim callstatus          As Boolean
Dim objWord             As Word.Application
Dim objDoc              As Word.Document
Dim i                   As Byte
Dim srno                As Byte
Dim sel_deptqr          As rdo.rdoQuery
Dim sel_deptrs          As rdo.rdoResultset
Dim sel_moadlityqr      As rdo.rdoQuery
Dim sel_modalityrs      As rdo.rdoResultset
Dim sel_FINAL_qr        As rdo.rdoQuery ' to get status of commit flag
Dim sel_ENTER_qr        As rdo.rdoQuery ' to get ENTER BY USER LIST
Dim SEL_REPORT_QR       As rdo.rdoQuery 'TO GET PERFORMED BY NAME
Dim SEL_REPORT_RS       As rdo.rdoResultset
Dim SEL_dreq_staff_qr   As rdo.rdoQuery ' to get staff detail
Dim SEL_dreq_staff_rs   As rdo.rdoResultset ' to get staff detail
Dim sel_ENTER_rs        As rdo.rdoResultset ' to  get ENTER BY USER LIST
Dim sel_FINAL_rs        As rdo.rdoResultset ' to  get FINAL BY USER LIST
Dim sel_ris_qr          As rdo.rdoQuery ' to get status of commit flag
Dim sel_ris_rs          As rdo.rdoResultset ' to get status of commit flag
Dim req_qr              As rdo.rdoQuery
Dim req_rs              As rdo.rdoResultset
Dim SELRREQUISITN_QR    As rdo.rdoQuery
Dim SELRREQUISITN_RS    As rdo.rdoResultset
Dim SEL_CHEMO_QR        As rdo.rdoQuery     'USED TO GET THE CHEMOTHERAPY OUTPUT FROM INDICATION TABLE
Dim SEL_CHEMO_RS        As rdo.rdoResultset 'WHICH FILLED DURING REQUISITION
'Dim sel_micro_qr        As RDO.rdoQuery ' to  get ENDOSCOPY number
'Dim sel_micro_rs        As RDO.rdoResultset ' to  get ENDOSCOPY number
Dim SEL_DOPPLER_QR      As rdo.rdoQuery 'USED TO GET DATA OF DOPPLER STUDY
Dim SEL_DOPPLER_RS      As rdo.rdoResultset
Dim SEL_OTHER_QR        As rdo.rdoQuery 'USED TO GET DATA OF OTHER DETAIL TEST
Dim SEL_OTHER_RS        As rdo.rdoResultset
Dim SEL_MODALITY_QR     As rdo.rdoQuery
Dim SEL_MODALITY_RS     As rdo.rdoResultset
Dim SEL_detail_qr       As rdo.rdoQuery
Dim SEL_detail_rs       As rdo.rdoResultset
Dim sel_bed_qr          As rdo.rdoQuery
Dim sel_bed_rs          As rdo.rdoResultset
Dim sel_proccode_qr     As rdo.rdoQuery
Dim sel_proccode_rs     As rdo.rdoResultset 'for selecting procedure code
Dim REFDR_QR            As rdo.rdoQuery
Dim REFDR_RS            As rdo.rdoResultset
Dim sel_prdiag_qr       As rdo.rdoQuery 'to get provisional diagnosis
Dim sel_prdiag_rs       As rdo.rdoResultset
Dim sel_prevdia_qr      As rdo.rdoQuery 'used for previous diagnosis
Dim sel_prevdia_rs      As rdo.rdoResultset
Dim upd_1_qr            As rdo.rdoQuery 'FOR UPDATING DOPPLER TABLE
Dim upd_2_qr            As rdo.rdoQuery  'FOR UPDATING ECCHOREPORT TABLE
Dim upd_11_qr           As rdo.rdoQuery
Dim upd_8_qr            As rdo.rdoQuery
Dim upd_6_qr            As rdo.rdoQuery
Dim ins_1_qr            As rdo.rdoQuery 'FOR INSERTING IN DOPPLER TABLE
Dim ins_2_qr            As rdo.rdoQuery 'FOR INSERTING IN MAIN ECCHOREPORT TABLE
Dim mod_reqno           As String 'used to store reqno by calling req_chglen method
Dim new_flag            As Boolean
Dim mod_flag            As Boolean
Dim VIEW_FLAG           As Boolean
Dim no_req              As Boolean 'to check before printing report whether req is existing or not
Dim S1                  As String
Dim S2                  As String

''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean

''*******ICU**********
''******SECURITY***********

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String
Dim REPWTRNRS As rdoResultset
Dim REPWTRNRS_REPBY As rdoResultset
Dim REPWTRNRS_FNLBY As rdoResultset
Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset

Dim REPWTRN_QR_REPBY As rdoQuery
Dim REPWTRN_QR_FNLBY As rdoQuery

''******SECURITY***********

Private Sub cmb_modality_Click()
On Error GoTo MSG
    i = InStr(cmb_modality.Text, "$")
    txt_modcode.Text = Mid(cmb_modality.Text, i + 1, 2)
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_reqno_Click()
On Error GoTo MSG
    'clearlowerpanel
    If cmb_reqno.ListIndex > 0 And (cmb_reqno.Text) = "" Or Len(cmb_reqno.Text) < 1 Then
        MsgBox "Select Requisition First", vbOKOnly
        btnenabl "INI"
    Exit Sub
    End If
        mod_reqno = req_chglen(cmb_reqno) 'put change value to this var so that it used later on
        'GETREPORTERLIST
        'GETFINALLIST
        'getenterlist
        'getperformedlist (Trim(USRDEPT))
    If Opt_patient.Value Then
        If new_flag Then
            ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = mod_reqno ' VAR_REQNNO
               Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
               SELICU_FLAG = True
            Else
               SELICU_QR(0) = mod_reqno 'VAR_REQNNO
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
            getupperpanel
        ElseIf VIEW_FLAG Then
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
            RETRIVE
            msk_procdate.Text = Format(DATE11, "dd/mm/yyyy")
        End If
        no_req = True
    ElseIf opt_staff.Value Then 'FOR STAFF MEMBER
        'if staff call staff procedure
        STAFF
        no_req = True
    End If 'END OF TOP IF
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_modify_Click()
On Error GoTo MSG
    btnenabl "NEW"
    new_flag = False
    VIEW_FLAG = False
    mod_flag = True
    If Len(Trim(MaskEdBox1.Text)) = 0 Then
        MsgBox "Case Number should not be empty!", vbCritical
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(cmb_reqno.Text)) = 0 Then
        MsgBox "Requisition Number should not be empty!", vbCritical
        If cmb_reqno.Enabled = True Then
            cmb_reqno.SetFocus
           ' SendKeys "^{f4}"
        End If
        Exit Sub
    End If
    'check for commit flag status
    sel_ris_qr(0) = mod_reqno
    Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
    While Not sel_ris_rs.EOF
        If Trim(sel_ris_rs(0)) = "Y" Then
            MsgBox "Sorry Record is Already Commited", vbOKOnly
            Exit Sub
        Else
            set_flag "MOD"
        End If
        sel_ris_rs.MoveNext
    Wend
    enable_field
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
'On Error GoTo MSG
DoEvents
 If no_req Then 'check for entered selected req no.
'    If Opt_patient.Value = True Then
'        If SELICU_FLAG = False Then
'            SELICU_QR(0) = mod_reqno
'            Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
'            SELICU_FLAG = True
'        Else
'            SELICU_QR(0) = mod_reqno
'            SELICU_RS.MoreResults
'            SELICU_RS.Requery
'        End If
'        If Not SELICU_RS.EOF Then
'            If SELMEMO_FLAG = False Then
'                SELMEMO_QR(0) = mod_reqno
'                Set SELMEMO_RS = SELMEMO_QR.OpenResultset(1, 3)
'                SELMEMO_FLAG = True
'            Else
'                SELMEMO_QR(0) = mod_reqno
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
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
    CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
    CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Echo2d_report.rpt"
    CrystalReport1.Formulas(1) = "proFIN = 'ECHO CARDIOGRAPHY REPORT'"
    CrystalReport1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enterby.Text) & "'"
    CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
    CrystalReport1.Formulas(7) = "pathnumber='" & txt_microno.Text & "'"
    If Not Trim(txt_ward.Text) = "" Then
         If InStr(1, txt_ward.Text, "WARD") = 0 Then
             CrystalReport1.Formulas(8) = "ward='" & txt_ward.Text & "'"
         Else
             CrystalReport1.Formulas(8) = "ward='" & Trim(Mid(txt_ward.Text, 1, InStr(1, txt_ward.Text, "WARD") - 1)) & "'"
         End If
    Else
            CrystalReport1.Formulas(8) = "ward="
    End If
    CrystalReport1.Formulas(9) = "NAM='" & txt_name.Text & "'"
    CrystalReport1.Formulas(10) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
    If Opt_patient.Value Then
        CrystalReport1.Formulas(11) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
    Else
        CrystalReport1.Formulas(11) = "ST='Staff /" & txt_status.Text & "'"
    End If
        CrystalReport1.Formulas(12) = "provdiag='" & txt_provdiag.Text & "'"
        Dim checkstatus As Boolean
        Dim j As Byte
        i = 1
        For i = 1 To 4
            For j = 1 To 3
                If Len(Trim(msflex_doppler.TextMatrix(i, j))) Then
                    checkstatus = True
                    GoTo here
                End If
            Next j
        Next i
here:
        If checkstatus Then
            CrystalReport1.Formulas(13) = "doppler='print'"
        Else
            CrystalReport1.Formulas(13) = "doppler=''"
        End If
        If Len(Trim(txt_comments.Text)) > 2 Then
            CrystalReport1.Formulas(14) = "result='print'"
        Else
            CrystalReport1.Formulas(14) = "result=''"
        End If
 Else
   MsgBox "Please select valid Requisition No.", vbOKOnly
   Exit Sub
 End If
 CrystalReport1.WindowState = crptMaximized
 CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

   CrystalReport1.Action = 1
   Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_procdone_lostfocus()
   ' cmb_reportby.Text = cmb_procdone.Text
End Sub

Private Sub Opt_patient_Click()
    If Opt_patient.Value = True Then
       lbl_caseno.Caption = "Case No."
      '  MaskEdBox1.Mask = "??>/#####"
    End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    'MaskEdBox1.Mask = "#######"         ' COMMENTED BY VARINDER
End If
End Sub

Private Sub txt_1_Change()
On Error GoTo MSG
    msflex_doppler.TextMatrix(msflex_doppler.row, 1) = txt_1.Text
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub txt_1_lostfocus()
'On Error GoTo msg
'    msflex_doppler.TextMatrix(msflex_doppler.Row, 1) = txt_1.Text
'Exit Sub
'msg:
'    MsgBox Err.Description
End Sub

Private Sub txt_2_Change()
On Error GoTo MSG
    msflex_doppler.TextMatrix(msflex_doppler.row, 2) = txt_2.Text
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub txt_2_lostfocus()
'On Error GoTo msg
'    msflex_doppler.TextMatrix(msflex_doppler.Row, 2) = txt_2.Text
'Exit Sub
'msg:
'    MsgBox Err.Description
End Sub

Private Sub txt_3_Change()
On Error GoTo MSG
    msflex_doppler.TextMatrix(msflex_doppler.row, 3) = txt_3.Text
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub txt_3_lostfocus()
'On Error GoTo msg
'    msflex_doppler.TextMatrix(msflex_doppler.Row, 3) = txt_3.Text
'Exit Sub
'msg:
'    MsgBox Err.Description
End Sub


Private Sub msflex_doppler_Click()
On Error GoTo MSG
    If msflex_doppler.col = 1 Or msflex_doppler.col = 2 Or msflex_doppler.col = 3 Then
        txt_1 = CStr(msflex_doppler.TextMatrix(msflex_doppler.row, 1))
        txt_2 = CStr(msflex_doppler.TextMatrix(msflex_doppler.row, 2))
        txt_3 = CStr(msflex_doppler.TextMatrix(msflex_doppler.row, 3))
        msflex_doppler.col = 1
        txt_1.Top = msflex_doppler.CellTop + msflex_doppler.Top
        msflex_doppler.col = 2
        txt_2.Top = msflex_doppler.CellTop + msflex_doppler.Top
        msflex_doppler.col = 3
        txt_3.Top = msflex_doppler.CellTop + msflex_doppler.Top
    End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_back_Click()
On Error GoTo MSG
    clearlastpanel
    clearlowerpanel
    clearupperpanel
    set_flag "All"
    btnenabl "INI"
    Me.Hide
    DIS_MAIN.show
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_cancel_Click()
On Error GoTo MSG
    btnenabl ("INI")
    clearupperpanel
    clearlowerpanel
    clearlastpanel
    CMD_NEW.SetFocus
    cmb_dept.Enabled = True
    cmb_modality.Enabled = True
    cmb_dept.Enabled = True
    cmb_reqno.Enabled = True
    MaskEdBox1.Enabled = True
    txt_1.Visible = False
    txt_2.Visible = False
    txt_3.Visible = False
    set_flag "Ni"
    Exit Sub
MSG:
    MsgBox Err.Description

End Sub

Private Sub cmd_commit_Click()
    Dim i As Integer
    i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
    If i = vbYes Then
    On Error GoTo MSG:
      ''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
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
''********SECURITY***************************

    
    
     rdb.BeginTrans
        upd_6_qr(0) = Left(Trim(cmb_finalby.Text), 30)
        upd_6_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        upd_6_qr(2) = "F"
        upd_6_qr(3) = "Y"
        upd_6_qr(4) = mod_reqno
        upd_6_qr.Execute
        End If
        'SHWETA UPDATE RTIME LOG 02-04-2013
      If UPDRTIMELOG(Trim(mod_reqno), "COMMIT") = True Then
      Else
      GoTo MSG
      End If
      '''TILL HERE
        
    rdb.CommitTrans
    btnenabl "INI"
    MsgBox "Record is Committed Successfully."
    cmd_print.Enabled = True
    Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG
    btnenabl ("NEW")
    new_flag = True
    VIEW_FLAG = False
    mod_flag = False
    'cmb_finalby.Text = cmb_finalby.List(0)
    cmb_enterby.Text = cmb_enterby.List(0)
    For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enterby.List(0)) = cmb_finalby.List(L1) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
    Next L1
    cmb_reportby.Text = cmb_reportby.List(0)
    cmb_procdone.Text = cmb_procdone.List(0)
    cmb_finalby.Enabled = True
    cmb_reportby.Enabled = True
    cmb_procdone.Enabled = True
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
    Wend
    For i = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(i)) = "GEN. MEDICINE SER.$CAZ" Then
                cmb_dept.ListIndex = i
        End If
    Next i
     If Trim(cmb_dept) = "" Then
        Exit Sub
    End If
    i = InStr(cmb_dept.Text, "$")
    txt_deptcode.Text = Mid(cmb_dept.Text, i + 1, 3)
    sel_moadlityqr(0) = Trim(txt_deptcode.Text)
    Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
        For i = 0 To (cmb_modality.ListCount - 1)
            If Trim(cmb_modality.List(i)) = "2DECHO$2D" Then
                cmb_modality.ListIndex = i
            End If
        Next i
        sel_modalityrs.Close
    Else
        sel_modalityrs.Close
    End If
sel_deptrs.Close
set_flag "NEW"
Opt_patient.Value = True
msflex_doppler.FormatString = "     Valve     |^       Peak/mean velocity M/s          |^          Peak/mean gradient mmHg       |^          Regurgitation grade           "
msflex_doppler.TextMatrix(1, 0) = "Mitral"
msflex_doppler.TextMatrix(2, 0) = "Tricuspid"
msflex_doppler.TextMatrix(3, 0) = "Aortic"
msflex_doppler.TextMatrix(4, 0) = "Pulmonary"
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_save_Click()
'On Error GoTo MSG:
If new_flag Then    'if new record then use insert query
    If Len(Trim(cmb_enterby.Text)) = 0 Then
        MsgBox "The Enter by field is EMPTY.", vbCritical
        cmb_enterby.SetFocus
    Exit Sub
    End If
    If Len(Trim(cmb_procdone)) = 0 Then
        MsgBox "The Procedure Done by field is EMPTY.", vbCritical
        cmb_procdone.SetFocus
        Exit Sub
    End If
    If Len(Trim(cmb_finalby)) = 0 Then
        MsgBox "The Finalized by field is EMPTY.", vbCritical
        cmb_finalby.SetFocus
        Exit Sub
    End If
    If Not Len(Trim(txt_summary.Text)) > 2 Then
         MsgBox "The Impression field should not be EMPTY", vbCritical
         txt_summary.SetFocus
         Exit Sub
    End If
        'write to save insert query
    btnenabl "SAVE"
    new_flag = False
    insert1
    update_rislib
   ' cmd_commit.Enabled = True
ElseIf mod_flag Then 'else if modifying a record then use update query
     'write to update record update query
        mod_flag = False
        btnenabl "SAVE"
        MODIFY
        update_rislib
End If

Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
new_flag = False
VIEW_FLAG = True
mod_flag = False
Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    'get the result
    While Not sel_deptrs.EOF
            cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
            sel_deptrs.MoveNext
    Wend
sel_deptrs.Close
If Trim(cmb_dept) = "" Then
        For i = 0 To (cmb_dept.ListCount - 1)
            If Trim(cmb_dept.List(i)) = "GEN. MEDICINE SER.$CAZ" Then
                cmb_dept.ListIndex = i
            End If
        Next i
End If
i = InStr(cmb_dept.Text, "$")
    txt_deptcode.Text = Mid(cmb_dept.Text, i + 1, 3)
    sel_moadlityqr(0) = Trim(txt_deptcode.Text)
    Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
        For i = 0 To (cmb_modality.ListCount - 1)
            If Trim(cmb_modality.List(i)) = "2DECHO$2D" Then
                cmb_modality.ListIndex = i
            End If
        Next i
        sel_modalityrs.Close
    Else
        sel_modalityrs.Close
    End If
    
set_flag "VIEW"
Opt_patient.Value = True
cmd_print.Enabled = True
msflex_doppler.FormatString = "     Valve     |^       Peak/mean velocity M/s          |^          Peak/mean gradient mmHg       |^          Regurgitation grade           "
msflex_doppler.TextMatrix(1, 0) = "Mitral"
msflex_doppler.TextMatrix(2, 0) = "Tricuspid"
msflex_doppler.TextMatrix(3, 0) = "Aortic"
msflex_doppler.TextMatrix(4, 0) = "Pulmonary"
'cmb_dept.SetFocus
Exit Sub
MSG:
    MsgBox Err.Description

End Sub

Private Sub Form_Load()
On Error GoTo MSG
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
msk_procdate.Text = Format(DATE11, "DD/MM/YYYY")
initialize_combo
new_flag = False
mod_flag = False
VIEW_FLAG = False
SSTab1.Tab = 0
'query definition starts here
'*******************************************************************************************
'************       ALL SELECT QUERY            *****************************************
'*******************************************************************************************

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1ECHO", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = Trim(USRNAME_RS("USERNAME"))
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Enabled = False
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'SUN ON 28/11/2003
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='CAZ' AND USRTYPE>=6 ORDER BY USERNAME"
Set REPWTRN_QR_FNLBY = rdb.CreateQuery("REPWTRN_STRCAZ1", str)

'REPWTRN_QR_FNLBY(0) = Trim(USRDEPT)
Set REPWTRNRS_FNLBY = REPWTRN_QR_FNLBY.OpenResultset(1, 1)
If REPWTRNRS_FNLBY.RowCount > 0 Then
    cmb_finalby.clear
    REPWTRNRS_FNLBY.MoveFirst
    While Not REPWTRNRS_FNLBY.EOF
           cmb_finalby.AddItem Trim(REPWTRNRS_FNLBY("USERNAME"))
           REPWTRNRS_FNLBY.MoveNext
    Wend
    cmb_finalby.Text = cmb_finalby.List(0)
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'SUN ON 28/11/2003
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='CAZ' AND USRTYPE>3 ORDER BY USERNAME"
Set REPWTRN_QR_REPBY = rdb.CreateQuery("REPWTRN_STRCAZ", str)

'REPWTRN_QR_REPBY(0) = Trim(USRDEPT)
Set REPWTRNRS_REPBY = REPWTRN_QR_REPBY.OpenResultset(1, 1)
If REPWTRNRS_REPBY.RowCount > 0 Then
     cmb_reportby.clear
     REPWTRNRS_REPBY.MoveFirst
    While Not REPWTRNRS_REPBY.EOF
               cmb_reportby.AddItem Trim(REPWTRNRS_REPBY("USERNAME"))
               REPWTRNRS_REPBY.MoveNext
    Wend
    cmb_reportby.Text = cmb_reportby.List(0)
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11ECHO", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'
'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1ECHO", STR)
'SELFINALBY_QR(0) = 6
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If

getperformedlist (Trim(USRDEPT))
'*********SECURITY**********************************

''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR114", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR114", str)

''********ICU***************
'for dept selection

str = "select dept_name, dept_code from PATOOLSLIB.ma106 where dept_code='CAZ'"
Set sel_deptqr = rdb.CreateQuery("echo_select1", str)

'FOR MODALITY SELECTION
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("echo_SELECT2", str)

'to check status of commitflag
str = "select * from rislib.rrequisitn where reqnno=?"
Set sel_ris_qr = rdb.CreateQuery("echo_select3", str)

'to get entered by name
str = "select CCNAME from dislib.dissec where PROVBY='YES' AND MODALITY='EC'"
Set sel_ENTER_qr = rdb.CreateQuery("echo_select4", str)

'to get final by list name
str = "select CCNAME from dislib.dissec where FINALby='YES' AND MODALITY='EC'"
Set sel_FINAL_qr = rdb.CreateQuery("echo_select5", str)

'for staff query
str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
Set SEL_dreq_staff_qr = rdb.CreateQuery("echo_select6", str)

'select requisition no.
str = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? and DEPT_CODE=? AND REQNSTAT='E' ORDER BY REQNDT DESC"
Set req_qr = rdb.CreateQuery("echo_select7", str)

'for view mode
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND  DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
Set SELRREQUISITN_QR = rdb.CreateQuery("echo_SELECT8", str)

'STR = "SELECT * from  DISLIB.Drequisitn where reqno=?"
'Set sel_micro_qr = rdb.CreateQuery("echo_SELECT11", STR)

'TO GET MODALITY DESC.
str = "SELECT MODLITYDES,MODALITYCD FROM RISLIB.RMODALITY WHERE DEPT_CODE=? AND MODALITYCD=?"
Set SEL_MODALITY_QR = rdb.CreateQuery("echo_SELECT12", str)

'select other detail for patient
str = "select * from PATOOLSLIB.ta101 where case_no =?"
Set SEL_detail_qr = rdb.CreateQuery("echo_select13", str)

'for referred by dr.
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
Set REFDR_QR = rdb.CreateQuery("echo_select14", str)

'for bed and ward retriving
str = "select * from PATOOLSLIB.ma104 where case_no =?"
Set sel_bed_qr = rdb.CreateQuery("echo_select15", str)

'FOR SELECTING WHICH PROCEDURE HE HAS SELECTED
str = "SELECT SER_CODE,SER_DET FROM DISLIB.MA101 WHERE SER_CODE IN(SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO= ?)"
Set sel_proccode_qr = rdb.CreateQuery("echo_select16", str)

'to get provisional diagnosis
str = "select PRVDIAGDES from patoolslib.ta1012 where case_no=?"
Set sel_prdiag_qr = rdb.CreateQuery("echo_select17", str)

'TO GET CHEMOTHERAPY
str = "SELECT DRUG,DOSE FROM dislib.INDICATION WHERE REQNO=?"
Set SEL_CHEMO_QR = rdb.CreateQuery("ECHO_SELECT18", str)

'TO GET DOPPLER TEST RESULT
str = "SELECT REQNO,VALVE,PMV,PMG,RG FROM DISLIB.DOPREPORT WHERE REQNO=?"
Set SEL_DOPPLER_QR = rdb.CreateQuery("ECHO_SELECT19", str)

'TO GET OTHER DETAIL LIKE GENERAL AND ECHO
str = "SELECT REQNO,LVIDD,LVIDS,IVSD,LVPWD,RVIDD,LA,AO,EF,FS,EPSS,HB,ECG,XRAYCHEST,REPORT,IMPRESSION," & _
      "MITRAVALVE,AML,PMLNORMAL,SUBAPP,AROVALVE,TRIVALVE,PLUMVALVE,PLANIMETRY,PHT,EXAMDATE,PPERFORMED,REMARKS,FLD1 FROM DISLIB.ECCHOREP WHERE REQNO=?"
Set SEL_OTHER_QR = rdb.CreateQuery("ECHO_SELECT20", str)

'TO GET PREVIOUS DIAGNOSIS
str = "select prvdiagdes from patoolslib.ta1012 where case_no=?"
Set sel_prevdia_qr = rdb.CreateQuery("ECHO_select21", str)

'*******************************************************************************************
'************       ALL UPDATE QUERY            *****************************************
'*******************************************************************************************
'update reqstat if he/she successfully entered data
str = "update rislib.rrequisitn set reqnstat='T' where caseno=? and reqnno=?"
Set upd_11_qr = rdb.CreateQuery("echo_update0", str)

'to update rislib for enterby and enter by
str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
Set upd_8_qr = rdb.CreateQuery("echo_update1", str)

'update commit flag
str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set upd_6_qr = rdb.CreateQuery("echo_update2", str)

'update DOPPLER TABLE
str = "update dislib.dopreport set pmv=?,pmg=?,rg=? where reqno=? and valve=?"
Set upd_1_qr = rdb.CreateQuery("echo_update3", str)

'UPDATE ECCHOREPORT TABLE
str = "update dislib.ecchorep set LVIDD=?,LVIDS=?,IVSD=?,LVPWD=?,RVIDD=?,LA=?,AO=?,EF=?,FS=?,EPSS=?,HB=?,ECG=?,XRAYCHEST=?,REPORT=?,IMPRESSION=?,MITRAVALVE=?,AML=?,PMLNORMAL=?,SUBAPP=?,AROVALVE=?,TRIVALVE=?,PLUMVALVE=?,PLANIMETRY=?,PHT=?,examdate=?,pperformed=?,remarks=?,FLD1=null WHERE REQNO=?"
Set upd_2_qr = rdb.CreateQuery("echo_update4", str)
'*******************************************************************************************
'************       ALL INSERT QUERY            *****************************************
'*******************************************************************************************
str = "INSERT INTO DISLIB.DOPREPORT(reqno,valve,pmv,pmg,rg) VALUES(?,?,?,?,?)"
Set ins_1_qr = rdb.CreateQuery("echo_0_ins", str)

str = "INSERT INTO DISLIB.ECCHOREP(REQNO,LVIDD,LVIDS,IVSD,LVPWD,RVIDD,LA,AO,EF,FS,EPSS,HB,ECG,XRAYCHEST,REPORT,IMPRESSION,MITRAVALVE ,AML,PMLNORMAL,SUBAPP,AROVALVE,TRIVALVE,PLUMVALVE,PLANIMETRY,PHT,EXAMDATE,PPERFORMED,REMARKS,FLD1) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set ins_2_qr = rdb.CreateQuery("echo_1_ins", str)

cmd_save.Enabled = False
cmd_commit.Enabled = False
cmd_print.Enabled = False
cmd_modify.Enabled = False
Exit Sub
MSG:
    MsgBox Err.Description
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
    If Trim(USRNAME_RS("COMMIT")) = "YES" Then
        cmd_commit.Enabled = True
    Else
        cmd_commit.Enabled = False
    End If
    
    If USRNAME_RS("VIEW") = "YES" Then
        cmd_view.Enabled = True
    Else
        cmd_view.Enabled = False
    End If
    
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
'ElseIf cap = "VIEW" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'ElseIf cap = "SAVE" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = True
'    cmd_view.Enabled = True
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'ElseIf cap = "COMMIT" Then
'    cmd_save.Enabled = False
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'ElseIf cap = "print" Then
'    cmd_save.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_cancel.Enabled = True
'    CMD_NEW.Enabled = False
'ElseIf cap = "cancel" Then
'    cmd_save.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_cancel.Enabled = True
'    CMD_NEW.Enabled = False
'End If
End Sub


Private Sub STAFF()
On Error GoTo MSG
If new_flag Then                        '1
   ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = mod_reqno ' VAR_REQNNO
               Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
               SELICU_FLAG = True
            Else
               SELICU_QR(0) = mod_reqno 'VAR_REQNNO
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
            SEL_dreq_staff_qr(0) = mod_reqno
            SEL_dreq_staff_qr(1) = MaskEdBox1.Text
            Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
            If SEL_dreq_staff_rs.RowCount > 0 Then
                txt_name = SEL_dreq_staff_rs("PNAME")
                txt_sex = SEL_dreq_staff_rs("SEX")
                txt_age = SEL_dreq_staff_rs("AGE")
                'txt_age = DateDiff("YYYY", SEL_dreq_staff_rs("AGE"), Format(DATE11, "DD/MM/YYYY"))
                If Trim(SEL_dreq_staff_rs("RELATION")) = "SELF" Then
                    txt_status = "ST"
                Else
                    txt_status = "DP"
                End If
            End If
            
            getupperpanel
                     
        ElseIf VIEW_FLAG Then            '1 else
        'commented by sundeep on 220104
'''''''''            ''******ICU****************
'''''''''        SEL_dreq_staff_qr(0) = mod_reqno
'''''''''        SEL_dreq_staff_qr(1) = MaskEdBox1.FormattedText
'''''''''        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
'''''''''        If SEL_dreq_staff_rs.RowCount > 0 Then
'''''''''            txt_name = SEL_dreq_staff_rs("PNAME")
'''''''''            txt_sex = SEL_dreq_staff_rs("SEX")
'''''''''            txt_age = SEL_dreq_staff_rs("AGE")
'''''''''            If Trim(SEL_dreq_staff_rs("RELATION")) <> "SELF" Then
'''''''''                txt_status = "DP"
'''''''''            Else
'''''''''                txt_status = "ST"
'''''''''            End If
'''''''''        End If
'''''''''    REFDR_QR(0) = MaskEdBox1.FormattedText
'''''''''    REFDR_QR(1) = reqno
'''''''''    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
'''''''''    If REFDR_RS.RowCount > 0 Then
'''''''''        txt_refdoc.Text = REFDR_RS("REFBYDOC")
'''''''''    End If
'''''''''    'sel_bed_qr(0) = MaskEdBox1.FormattedText
'''''''''    'Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
'''''''''        'If sel_bed_rs.RowCount > 0 Then
''''''''''            txt_ward.Visible = True
''''''''''            txt_bedno.Visible = True
''''''''''            lbl_ward.Visible = True
''''''''''            lbl_bed.Visible = True
''''''''''            txt_ward.Text = sel_bed_rs("WARD_NAME")
''''''''''            txt_bedno.Text = sel_bed_rs("bed_no")
''''''''''            sel_bed_rs.Close
''''''''''        Else
''''''''''            txt_status.Text = "OP"
''''''''''            txt_ward.Visible = False
''''''''''            txt_bedno.Visible = False
''''''''''            lbl_ward.Visible = False
''''''''''            lbl_bed.Visible = False
''''''''''            sel_bed_rs.Close
'''''''''       ' End If
'''''''''        'insert1
''''''''' ElseIf view_flag Then
'''''''''     ''******ICU****************
'''''''''            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
'''''''''            If SELICU_FLAG = False Then
'''''''''               SELICU_QR(0) = mod_reqno ' VAR_REQNNO
'''''''''               Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
'''''''''               SELICU_FLAG = True
'''''''''            Else
'''''''''               SELICU_QR(0) = mod_reqno 'VAR_REQNNO
'''''''''               SELICU_RS.MoreResults
'''''''''               SELICU_RS.Requery
'''''''''            End If
'''''''''            If Not SELICU_RS.EOF Then
'''''''''               If Trim(SELICU_RS("ICDS2")) = "YES" Then
'''''''''                  lblicu.Visible = True
'''''''''               Else
'''''''''                  lblicu.Visible = False
'''''''''               End If
'''''''''            End If
'''''''''            ''******ICU****************
'''''''''            getupperpanel
'commented by sundeep on 220104
        'ElseIf view_flag Then
            ''******ICU****************
        SEL_dreq_staff_qr(0) = mod_reqno
        SEL_dreq_staff_qr(1) = MaskEdBox1.Text
        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
        If SEL_dreq_staff_rs.RowCount > 0 Then
            txt_name = SEL_dreq_staff_rs("PNAME")
            txt_sex = SEL_dreq_staff_rs("SEX")
            txt_age = SEL_dreq_staff_rs("AGE")
            'txt_age = DateDiff("YYYY", SEL_dreq_staff_rs("AGE"), Format(DATE11, "DD/MM/YYYY"))
            If Trim(SEL_dreq_staff_rs("RELATION")) = "SELF" Then
                txt_status = "ST"
            Else
                txt_status = "DP"
            End If
        End If
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = mod_reqno
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
    sel_bed_qr(0) = MaskEdBox1.FormattedText
    Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
        If sel_bed_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
            txt_ward.Text = sel_bed_rs("WARD_NAME")
            txt_bedno.Text = sel_bed_rs("bed_no")
            sel_bed_rs.Close
        Else
            txt_status.Text = "OP"
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
            sel_bed_rs.Close
        End If
        RETRIVE
 End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub


Private Sub disable_field()
On Error GoTo MSG
    txt_lvidd.Enabled = False
    txt_lvids.Enabled = False
    txt_ivsd.Enabled = False
    txt_lvpwd.Enabled = False
    txt_rvidd.Enabled = False
    txt_la.Enabled = False
    txt_ao.Enabled = False
    txt_ef.Enabled = False
    txt_fs.Enabled = False
    TXT_EPSS.Enabled = False
    TXT_HB.Enabled = False
    If Len(Trim(txt_drug.Text)) > 2 Then
        txt_drug.Enabled = False
        txt_dose.Enabled = False
    Else
        txt_drug.Enabled = True
        txt_dose.Enabled = True
    End If
    TXT_ECG.Enabled = False
    TXT_XRAYCHEST.Enabled = False
    txt_comments.Enabled = False
    txt_summary.Enabled = False
    cmb_mitral.Enabled = False
    cmb_aml.Enabled = False
    cmb_pml.Enabled = False
    cmb_subvalvular.Enabled = False
    cmb_arotic.Enabled = False
    Txt_tricuspid.Enabled = False
    txt_pulmonary.Enabled = False
    txt_planimetry.Enabled = False
    txt_pht.Enabled = False
    txt_1.Enabled = False
    txt_2.Enabled = False
    txt_3.Enabled = False
    'and this three text box visible false
    txt_1.Visible = False
    txt_2.Visible = False
    txt_2.Visible = False
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub enable_field()
On Error GoTo MSG
    txt_lvidd.Enabled = True
    txt_lvids.Enabled = True
    txt_ivsd.Enabled = True
    txt_lvpwd.Enabled = True
    txt_rvidd.Enabled = True
    txt_la.Enabled = True
    txt_ao.Enabled = True
    txt_ef.Enabled = True
    txt_fs.Enabled = True
    TXT_EPSS.Enabled = True
    TXT_HB.Enabled = True
    'txt_drug.Enabled = True
    'txt_dose.Enabled = True
    TXT_ECG.Enabled = True
    TXT_XRAYCHEST.Enabled = True
    txt_comments.Enabled = True
    txt_summary.Enabled = True
    cmb_mitral.Enabled = True
    cmb_aml.Enabled = True
    cmb_pml.Enabled = True
    cmb_subvalvular.Enabled = True
    cmb_arotic.Enabled = True
    Txt_tricuspid.Enabled = True
    txt_pulmonary.Enabled = True
    txt_planimetry.Enabled = True
    txt_pht.Enabled = True
    cmb_enterby.Enabled = True
    cmb_reportby.Enabled = True
    cmb_finalby.Enabled = True
    set_position
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

'common procedure for new and view mode ALONG WITH FOR CHEMOTHERAPY
Private Sub getupperpanel()
On Error GoTo MSG
        'THIS TO GET DEPTCODE AND MODALITYCODE
'        sel_micro_qr(0) = mod_reqno
'        Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
'        While Not sel_micro_rs.EOF
'            txt_microno.Text = sel_micro_rs("PATHNO")
'            'STILL TO WRITE FOR PROCEDURE HEADING
'            S1 = sel_micro_rs("METASTASIS")
        '
        '
        '
        '    sel_micro_rs.MoveNext
        'Wend
        set_position
        SEL_CHEMO_QR(0) = mod_reqno
        Set SEL_CHEMO_RS = SEL_CHEMO_QR.OpenResultset(1, 3)
        If Not SEL_CHEMO_RS.EOF Then
            txt_drug.Text = "" & SEL_CHEMO_RS("DRUG")
            txt_dose.Text = "" & SEL_CHEMO_RS("DOSE")
            SEL_CHEMO_RS.MoveNext
            SEL_CHEMO_RS.Close
        Else
             SEL_CHEMO_RS.Close
        End If
            sel_prevdia_qr(0) = MaskEdBox1.FormattedText
        Set sel_prevdia_rs = sel_prevdia_qr.OpenResultset(1, 3)
        If Not sel_prevdia_rs.EOF Then
            txt_provdiag.Text = "" & sel_prevdia_rs(0)
            sel_prevdia_rs.Close
        Else
            sel_prevdia_rs.Close
        End If
        S1 = Mid(cmb_reqno.Text, 1, 3) 'TO GET DEPTCODE
        S2 = Mid(cmb_reqno.Text, 5, 2) ' TO GET MODALITY CODE
        SEL_MODALITY_QR(0) = S1
        SEL_MODALITY_QR(1) = S2
        Set SEL_MODALITY_RS = SEL_MODALITY_QR.OpenResultset(1, 3)
        While Not SEL_MODALITY_RS.EOF
            txt_modcode.Text = Trim(SEL_MODALITY_RS("MODLITYDES"))
            SEL_MODALITY_RS.MoveNext
        Wend
        txt_deptcode.Text = S1
        txt_modcode.Text = S2
        If Opt_patient.Value Then
            SEL_detail_qr(0) = MaskEdBox1.FormattedText
            Set SEL_detail_rs = SEL_detail_qr.OpenResultset(1, 3)
            If SEL_detail_rs.RowCount = 0 Then
                SEL_detail_rs.Close
                Exit Sub
            Else
                txt_sex.Text = SEL_detail_rs("sex")
                txt_name.Text = Trim(SEL_detail_rs("first_name")) & " " & Trim(SEL_detail_rs("mid_name")) & " " & Trim(SEL_detail_rs("last_name"))
                txt_age = DateDiff("YYYY", SEL_detail_rs("BIRTH_DATE"), CStr(DATE11))
                txt_pridoc.Text = SEL_detail_rs("ud_name")
                txt_category.Text = SEL_detail_rs("pat_catg")
                SEL_detail_rs.Close
            End If
        End If
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = mod_reqno
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            If REFDR_RS.RowCount > 0 Then
            'put check here also 20/12/02
            '***********************************
                 txt_refdoc.Text = REFDR_RS("REFBYDOC")
                If Not callstatus Then
                    If REFDR_RS("COMMITFLAG") = "Y" Then
                        MsgBox "Report is Commited,You cannot Modify the Final Report"
                        cmd_modify.Enabled = False
                    End If
                End If
            End If
            sel_bed_qr(0) = MaskEdBox1.FormattedText
            Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
        If sel_bed_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
            txt_ward.Text = sel_bed_rs("WARD_NAME")
            txt_bedno.Text = sel_bed_rs("bed_no")
            txt_status.Text = "IP"
            sel_bed_rs.Close
        Else
            txt_status.Text = "OP"
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
            sel_bed_rs.Close
        End If
        If Len(Trim(txt_drug.Text)) > 2 Then
            txt_drug.Enabled = False
            txt_dose.Enabled = False
        End If
        
        'sel_proccode_qr(0) = mod_reqno
        'Set sel_proccode_rs = sel_proccode_qr.OpenResultset(1, 3)
        'If Not sel_proccode_rs.EOF Then
        '        txt_procname.Text = Trim(sel_proccode_rs(1))
        '        sel_proccode_rs.MoveNext
        '        sel_proccode_rs.Close
        ' Else
        '    sel_proccode_rs.Close
        ' End If
        ' sel_prdiag_qr(0) = MaskEdBox1.FormattedText
        ' Set sel_prdiag_rs = sel_prdiag_qr.OpenResultset(1, 3)
        ' If Not sel_prdiag_rs.EOF Then
        '        txt_procname.Text = Trim(sel_prdiag_rs(0))
        '        sel_prdiag_rs.MoveNext
        '        sel_prdiag_rs.Close
        ' Else
        '    sel_prdiag_rs.Close
        ' End If
        Exit Sub
MSG:
        MsgBox Err.Description
End Sub

Private Sub initialize_combo()
On Error GoTo MSG
cmb_subvalvular.AddItem "Normal"
cmb_subvalvular.AddItem "Thickened"
cmb_subvalvular.AddItem "Crowding"
cmb_aml.AddItem "Normal"
cmb_aml.AddItem "Thickened"
cmb_aml.AddItem "Calcified"
cmb_aml.AddItem "Restricted motion"
cmb_pml.AddItem "Normal"
cmb_pml.AddItem "Thickened"
cmb_pml.AddItem "Calcified"
cmb_pml.AddItem "Restricted motion"
cmb_mitral.AddItem "Annulus-normal"
cmb_mitral.AddItem "Annulus-calcified"
cmb_arotic.AddItem "Normal"
cmb_arotic.AddItem "Thickened"
cmb_arotic.AddItem "Calcified"
cmb_arotic.AddItem "Bicuspid"
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub set_flag(flag As String) 'used to flag status
If Trim(flag) = "NEW" Then
    new_flag = True
    VIEW_FLAG = False
    mod_flag = False
ElseIf Trim(flag) = "VIEW" Then
    new_flag = False
    VIEW_FLAG = True
    mod_flag = False
ElseIf Trim(flag) = "MOD" Then
    new_flag = False
    VIEW_FLAG = False
    mod_flag = True
Else
    new_flag = False
    VIEW_FLAG = False
    mod_flag = False
End If
End Sub

Private Sub clearupperpanel()
On Error GoTo MSG
    lblicu.Visible = False
    MaskEdBox1.Text = ""
    cmb_reqno.clear
    cmb_modality.clear
    txt_deptcode.Text = ""
    cmb_dept.clear
    txt_microno.Text = ""
    txt_name.Text = ""
    txt_age.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_status.Text = ""
    txt_ward.Text = ""
    txt_bedno.Text = ""
    txt_pridoc.Text = ""
    txt_modcode.Text = ""
    txt_refdoc.Text = ""
    'txt_lasttreat.Text = ""
Exit Sub
MSG:
        MsgBox Err.Description
End Sub

Private Sub clearlastpanel()
On Error GoTo MSG
    lblicu.Visible = False
    txt_comments.Text = ""
    cmb_enterby.Text = ""
    cmb_finalby.Text = ""
    cmb_reportby.Text = ""
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub clearlowerpanel()
On Error GoTo MSG
    lblicu.Visible = False
    txt_lvidd.Text = ""
    txt_lvids.Text = ""
    txt_ivsd.Text = ""
    txt_lvpwd.Text = ""
    txt_rvidd.Text = ""
    txt_la.Text = ""
    txt_ao.Text = ""
    txt_ef.Text = ""
    txt_fs.Text = ""
    TXT_EPSS.Text = ""
    TXT_HB.Text = ""
    txt_drug.Text = ""
    txt_dose.Text = ""
    TXT_ECG.Text = ""
    TXT_XRAYCHEST.Text = ""
    txt_comments.Text = ""
    txt_summary.Text = ""
    cmb_mitral.Text = ""
    cmb_aml.Text = ""
    cmb_pml.Text = ""
    cmb_subvalvular.Text = ""
    cmb_arotic.Text = ""
    Txt_tricuspid.Text = ""
    txt_pulmonary.Text = ""
    txt_planimetry.Text = ""
    txt_pht.Text = ""
    txt_1.Text = ""
    txt_2.Text = ""
    txt_3.Text = ""
    msflex_doppler.clear
    msk_procdate.Mask = "##/##/####"
    msk_procdate.Text = Format(DATE11, "DD/MM/YYYY")
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub


Private Sub MaskEdBox1_LostFocus()
On Error GoTo MSG:
If Trim(MaskEdBox1.Text) = "" Then
    btnenabl "INI"
    Exit Sub
End If
If Opt_patient.Value Then
   '' MaskEdBox1.Text = UCase(padzero(5, "C", MaskEdBox1.Text, 3)) '''commented by Monika on 05/12/2022
   '''' Commented PADZERO FUNCTION INSTEAD USED NEWNOPAD FUNCTION
'cno_pad = padzero(5, "C", msk_caseno.Text, 3)
    

    
  If IsNumeric(Left(Trim(MaskEdBox1.Text), 1)) Then
          Dim Ks As Integer
        Dim str1s As String
        Dim strs As String
        Dim cno_pad As String
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
    If new_flag Then
            cmb_reqno.clear
            req_qr(0) = MaskEdBox1.FormattedText
            req_qr(1) = Trim(txt_modcode.Text)
            req_qr(2) = Trim(txt_deptcode.Text)
            Set req_rs = req_qr.OpenResultset(1, 3)
        If req_rs.RowCount > 0 Then
            no_req = True
            While Not req_rs.EOF
                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
                req_rs.MoveNext
            Wend
            req_rs.Close
        Else
            no_req = False
            req_rs.Close
            MsgBox "No Requisition for this Modality", vbOKOnly
            btnenabl "cancel"
            set_flag "asd"
            Exit Sub
            End If 'end of rowcount
        cmb_reqno.SetFocus
      '  SendKeys "^{f4}"
    End If  'end of new_flag
    If VIEW_FLAG Then
            new_flag = False
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
            If SELRREQUISITN_RS.RowCount > 0 Then
                no_req = True
                cmb_reqno.clear
                While Not SELRREQUISITN_RS.EOF
                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                    SELRREQUISITN_RS.MoveNext
                Wend
                SELRREQUISITN_RS.Close
            Else
                no_req = False
                SELRREQUISITN_RS.Close
                MsgBox "No Requisition for this Modality", vbOKOnly
                btnenabl "cancel"
                set_flag "asd"
                Exit Sub
            End If
        End If
ElseIf opt_staff.Value = True Then
    If new_flag Then
        cmb_reqno.clear
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                req_qr(2) = Trim(txt_deptcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            If req_rs.RowCount > 0 Then
                no_req = True
                While Not req_rs.EOF
                    cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
                    req_rs.MoveNext
                Wend
            Else
                cmb_reqno.clear
                no_req = False
                MsgBox "No Requisition for this Modality", vbOKOnly
                btnenabl ("INI")
                Exit Sub
            End If
    End If
    If VIEW_FLAG Then
                cmb_reqno.clear
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
            'End If
            If SELRREQUISITN_RS.RowCount > 0 Then
                no_req = True
                While Not SELRREQUISITN_RS.EOF
                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                    SELRREQUISITN_RS.MoveNext
                Wend
                SELRREQUISITN_RS.Close
             Else
                no_req = False
                MsgBox " No Requisition for this Modality"
                SELRREQUISITN_RS.Close
                btnenabl ("INI")
                Exit Sub
             End If
        End If
End If
 Exit Sub
MSG:
    MsgBox Err.Description

End Sub
Private Sub getenterlist()
On Error GoTo MSG
    Dim sel_performed_qr As rdo.rdoQuery
    Dim sel_performed_rs As rdo.rdoResultset
    str = "select ccname from dislib.dissec where enterby='YES' AND MODALITY='EC' "
    Set sel_performed_qr = rdb.CreateQuery("echo_perfomedby", str)
    Set sel_performed_rs = sel_performed_qr.OpenResultset(1, 3)
    If Not sel_performed_rs.EOF Then
        cmb_enterby.clear
            While Not sel_performed_rs.EOF
                cmb_enterby.AddItem Trim(sel_performed_rs("CCNAME"))
                sel_performed_rs.MoveNext
            Wend
        sel_performed_rs.Close
        sel_performed_qr.Close
    Else
        sel_performed_rs.Close
        sel_performed_qr.Close
    End If
    For i = 0 To (cmb_enterby.ListCount - 1)
        If Trim(cmb_enterby.List(i)) = "MR. CHANDRA MOHAN" Then
                cmb_enterby.ListIndex = i
        End If
    Next i
Exit Sub
MSG:
    MsgBox Err.Description

End Sub

Private Sub GETFINALLIST()
On Error GoTo MSG
    Set sel_FINAL_rs = sel_FINAL_qr.OpenResultset(1, 3)
    If sel_FINAL_rs.RowCount > 0 Then
        cmb_finalby.clear
            While Not sel_FINAL_rs.EOF
                cmb_finalby.AddItem Trim(sel_FINAL_rs("CCNAME"))
                sel_FINAL_rs.MoveNext
            Wend
            sel_FINAL_rs.Close
    Else
        sel_FINAL_rs.Close
    End If
'    For i = 0 To (cmb_final.ListCount - 1)
'        If Trim(cmb_final.List(i)) = "Dr.Rohini Kelkar" Then
'                cmb_final.ListIndex = i
'        End If
'    Next i
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub


Private Sub GETREPORTERLIST()
On Error GoTo MSG
    Set sel_ENTER_rs = sel_ENTER_qr.OpenResultset(1, 3)
    If sel_ENTER_rs.RowCount > 0 Then
        cmb_reportby.clear
            While Not sel_ENTER_rs.EOF
                cmb_reportby.AddItem Trim(sel_ENTER_rs("CCNAME"))
                sel_ENTER_rs.MoveNext
            Wend
        sel_ENTER_rs.Close
    Else
        sel_ENTER_rs.Close
    End If
    'For i = 0 To (cmb_enterby.ListCount - 1)
    '    If Trim(cmb_enterby.List(i)) = "MR. CHANDRA MOHAN" Then
    '            cmb_enterby.ListIndex = i
    '    End If
    'Next i
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub getperformedlist(DEPTCODE As String)
On Error GoTo MSG
    Dim sel_performed_qr As rdo.rdoQuery
    Dim sel_performed_rs As rdo.rdoResultset
    str = "select USERNAME from dislib.dislnkusr where USRTYPE>=1 AND DEPT_CODE=? "
    Set sel_performed_qr = rdb.CreateQuery("echo_perfomedby", str)
    sel_performed_qr(0) = Trim(DEPTCODE)
    Set sel_performed_rs = sel_performed_qr.OpenResultset(1, 3)
    If Not sel_performed_rs.EOF Then
        cmb_procdone.clear
            While Not sel_performed_rs.EOF
                cmb_procdone.AddItem Trim(sel_performed_rs("USERNAME"))
                sel_performed_rs.MoveNext
            Wend
        sel_performed_rs.Close
        sel_performed_qr.Close
    Else
        sel_performed_rs.Close
        sel_performed_qr.Close
    End If
    
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub RETRIVE()
On Error GoTo MSG
    getupperpanel
    'GET BOTH DETAIL LIKE DOPPLER AND OTHER TEST
    SEL_DOPPLER_QR(0) = mod_reqno
    Set SEL_DOPPLER_RS = SEL_DOPPLER_QR.OpenResultset(1, 3)
    srno = 1
    If Not SEL_DOPPLER_RS.EOF Then
        While Not SEL_DOPPLER_RS.EOF
           'HERE WRITE PROGRAM TO LOAD FLEXGRID
            msflex_doppler.TextMatrix(srno, 0) = SEL_DOPPLER_RS("valve")
            msflex_doppler.TextMatrix(srno, 1) = SEL_DOPPLER_RS("pmv")
            msflex_doppler.TextMatrix(srno, 2) = SEL_DOPPLER_RS("pmg")
            msflex_doppler.TextMatrix(srno, 3) = SEL_DOPPLER_RS("rg")
            srno = srno + 1
            SEL_DOPPLER_RS.MoveNext
        Wend
            SEL_DOPPLER_RS.Close
     Else
            SEL_DOPPLER_RS.Close
     End If
    SEL_OTHER_QR(0) = mod_reqno
    Set SEL_OTHER_RS = SEL_OTHER_QR.OpenResultset(1, 3)
    If Not SEL_OTHER_RS.EOF Then
        While Not SEL_OTHER_RS.EOF
            txt_lvidd.Text = SEL_OTHER_RS("LVIDD")
            txt_lvids.Text = SEL_OTHER_RS("LVIDS")
            txt_ivsd.Text = SEL_OTHER_RS("IVSD")
            txt_lvpwd.Text = SEL_OTHER_RS("LVPWD")
            txt_rvidd.Text = SEL_OTHER_RS("RVIDD")
            txt_la.Text = SEL_OTHER_RS("LA")
            txt_ao.Text = SEL_OTHER_RS("AO")
            txt_ef.Text = SEL_OTHER_RS("EF")
            txt_fs.Text = SEL_OTHER_RS("FS")
            TXT_EPSS.Text = SEL_OTHER_RS("EPSS")
            TXT_HB.Text = SEL_OTHER_RS("HB")
            TXT_ECG.Text = SEL_OTHER_RS("ECG")
            TXT_XRAYCHEST.Text = SEL_OTHER_RS("XRAYCHEST")
            txt_comments.Text = " " & SEL_OTHER_RS("REPORT")
            txt_summary.Text = " " & SEL_OTHER_RS("IMPRESSION")
            'remember
            cmb_mitral.Text = SEL_OTHER_RS("MITRAVALVE")
            cmb_aml.Text = SEL_OTHER_RS("AML")
            cmb_pml.Text = SEL_OTHER_RS("PMLNORMAL")
            cmb_subvalvular.Text = SEL_OTHER_RS("SUBAPP")
            cmb_arotic.Text = SEL_OTHER_RS("AROVALVE")
            Txt_tricuspid.Text = SEL_OTHER_RS("TRIVALVE")
            txt_pulmonary.Text = SEL_OTHER_RS("PLUMVALVE")
            txt_planimetry.Text = SEL_OTHER_RS("PLANIMETRY")
            txt_pht.Text = SEL_OTHER_RS("PHT")
            msk_procdate.Mask = "##/##/####"
            msk_procdate.Text = Format(SEL_OTHER_RS("EXAMDATE"), "DD/MM/YYYY")
            cmb_procdone.Text = SEL_OTHER_RS("PPERFORMED")
            txt_remarks.Text = SEL_OTHER_RS("REMARKS")
            SEL_OTHER_RS.MoveNext
        Wend
        SEL_OTHER_RS.Close
    Else
        SEL_OTHER_RS.Close
    End If
    GET_ENTER_BY
    disable_field
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub insert1()
On Error GoTo MSG
    ins_2_qr(0) = mod_reqno
    ins_2_qr(1) = txt_lvidd.Text
    ins_2_qr(2) = txt_lvids.Text
    ins_2_qr(3) = txt_ivsd.Text
    ins_2_qr(4) = txt_lvpwd.Text
    ins_2_qr(5) = txt_rvidd.Text
    ins_2_qr(6) = txt_la.Text
    ins_2_qr(7) = txt_ao.Text
    ins_2_qr(8) = txt_ef.Text
    ins_2_qr(9) = txt_fs.Text
    ins_2_qr(10) = TXT_EPSS.Text
    ins_2_qr(11) = TXT_HB.Text
    If Len(TXT_ECG.Text) > 0 Then
        ins_2_qr(12) = Null
        ins_2_qr(12) = TXT_ECG.Text
    Else
        ins_2_qr(12) = Null
        ins_2_qr(12) = ""
    End If
    If Len(TXT_XRAYCHEST.Text) > 0 Then
        ins_2_qr(13) = Null
        ins_2_qr(13) = TXT_XRAYCHEST
    Else
        ins_2_qr(13) = Null
        ins_2_qr(13) = ""
    End If
    If Len(txt_comments.Text) > 0 Then
        ins_2_qr(14) = Null
        ins_2_qr(14) = txt_comments.Text
    Else
        ins_2_qr(14) = Null
        ins_2_qr(14) = ""
    End If
    If Len(txt_summary.Text) > 0 Then
        ins_2_qr(15) = Null
        ins_2_qr(15) = txt_summary.Text
    Else
        ins_2_qr(15) = Null
        ins_2_qr(15) = ""
    End If
    ins_2_qr(16) = cmb_mitral
    ins_2_qr(17) = cmb_aml
    If Len(cmb_pml.Text) > 0 Then
        ins_2_qr(18) = Null
        ins_2_qr(18) = cmb_pml.Text
    Else
        ins_2_qr(18) = Null
        ins_2_qr(18) = ""
    End If
    ins_2_qr(19) = cmb_subvalvular
    ins_2_qr(20) = cmb_arotic
    If Len(Txt_tricuspid.Text) > 0 Then
        ins_2_qr(21) = Null
        ins_2_qr(21) = Txt_tricuspid.Text
    Else
        ins_2_qr(21) = Null
        ins_2_qr(21) = ""
    End If
    If Len(txt_pulmonary.Text) > 0 Then
        ins_2_qr(22) = Null
        ins_2_qr(22) = txt_pulmonary.Text
    Else
        ins_2_qr(22) = Null
        ins_2_qr(22) = ""
    End If
    ins_2_qr(23) = txt_planimetry.Text
    ins_2_qr(24) = txt_pht.Text
    ins_2_qr(25) = msk_procdate.FormattedText
    ins_2_qr(26) = cmb_procdone.Text
    If Len(txt_remarks.Text) > 0 Then
        ins_2_qr(27) = Null
        ins_2_qr(27) = txt_remarks.Text
    Else
        ins_2_qr(27) = Null
        ins_2_qr(27) = ""
    End If
    ins_2_qr(28) = Null
    ins_2_qr.Execute
    'now insert doppler study result
    i = msflex_doppler.Rows
    srno = 1
    While Not srno = i
            ins_1_qr(0) = mod_reqno
            ins_1_qr(1) = Trim(msflex_doppler.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
            ins_1_qr(2) = Trim(msflex_doppler.TextMatrix(srno, 1)) 'get first value and insert
            ins_1_qr(3) = Trim(msflex_doppler.TextMatrix(srno, 2)) 'get second value and insert"
            ins_1_qr(4) = Trim(msflex_doppler.TextMatrix(srno, 3)) 'get third value and insert
            ins_1_qr.Execute
        srno = srno + 1
    Wend
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub MODIFY()
On Error GoTo MSG
    upd_2_qr(0) = txt_lvidd.Text
    upd_2_qr(1) = txt_lvids.Text
    upd_2_qr(2) = txt_ivsd.Text
    upd_2_qr(3) = txt_lvpwd.Text
    upd_2_qr(4) = txt_rvidd.Text
    upd_2_qr(5) = txt_la.Text
    upd_2_qr(6) = txt_ao.Text
    upd_2_qr(7) = txt_ef.Text
    upd_2_qr(8) = txt_fs.Text
    upd_2_qr(9) = TXT_EPSS.Text
    upd_2_qr(10) = TXT_HB.Text
    If Len(TXT_ECG.Text) > 0 Then
        upd_2_qr(11) = Null
        upd_2_qr(11) = TXT_ECG.Text
    Else
        upd_2_qr(11) = Null
        upd_2_qr(11) = ""
    End If
    If Len(TXT_XRAYCHEST.Text) > 0 Then
        upd_2_qr(12) = Null
        upd_2_qr(12) = TXT_XRAYCHEST
    Else
        upd_2_qr(12) = Null
        upd_2_qr(12) = ""
    End If
    If Len(txt_comments.Text) > 0 Then
        upd_2_qr(13) = Null
        upd_2_qr(13) = txt_comments.Text
    Else
        upd_2_qr(13) = Null
        upd_2_qr(13) = ""
    End If
    If Len(txt_summary.Text) > 0 Then
        upd_2_qr(14) = Null
        upd_2_qr(14) = txt_summary.Text
    Else
        upd_2_qr(14) = Null
        upd_2_qr(14) = ""
    End If
    upd_2_qr(15) = cmb_mitral.Text
    upd_2_qr(16) = cmb_aml.Text
    If Len(cmb_pml.Text) > 0 Then
        upd_2_qr(17) = Null
        upd_2_qr(17) = cmb_pml.Text
    Else
        upd_2_qr(17) = Null
        upd_2_qr(17) = ""
    End If
    upd_2_qr(18) = cmb_subvalvular.Text
    upd_2_qr(19) = cmb_arotic.Text
    If Len(Txt_tricuspid.Text) > 0 Then
        upd_2_qr(20) = Null
        upd_2_qr(20) = Txt_tricuspid.Text
    Else
        upd_2_qr(20) = Null
        upd_2_qr(20) = ""
    End If
    If Len(txt_pulmonary.Text) > 0 Then
        upd_2_qr(21) = Null
        upd_2_qr(21) = txt_pulmonary.Text
    Else
        upd_2_qr(21) = Null
        upd_2_qr(21) = ""
    End If
    upd_2_qr(22) = txt_planimetry.Text
    upd_2_qr(23) = txt_pht.Text
    upd_2_qr(24) = msk_procdate.FormattedText
    upd_2_qr(25) = cmb_procdone.Text
    If Len(txt_remarks.Text) > 0 Then
        upd_2_qr(26) = Null
        upd_2_qr(26) = txt_remarks.Text
    Else
        upd_2_qr(26) = Null
        upd_2_qr(26) = ""
    End If
    
    
'    UPD_2_QR(24) = Null
    upd_2_qr(27) = mod_reqno
    upd_2_qr.Execute
    i = msflex_doppler.Rows
    srno = 1
    While Not srno = i
            upd_1_qr(0) = Trim(msflex_doppler.TextMatrix(srno, 1)) 'get first value and insert
            upd_1_qr(1) = Trim(msflex_doppler.TextMatrix(srno, 2)) 'get second value and insert"
            upd_1_qr(2) = Trim(msflex_doppler.TextMatrix(srno, 3)) 'get third value and insert
            upd_1_qr(3) = mod_reqno
            upd_1_qr(4) = Trim(msflex_doppler.TextMatrix(srno, 0))
            upd_1_qr.Execute
        srno = srno + 1
    Wend
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub update_rislib()
On Error GoTo MSG
    upd_8_qr(0) = cmb_enterby.Text
    upd_8_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_8_qr(2) = cmb_reportby.Text
    upd_8_qr(3) = "T"
    upd_8_qr(4) = cmb_finalby.Text
    upd_8_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_8_qr(6) = Trim(mod_reqno)
    upd_8_qr.Execute
   
    'SHWETA UPDATE RTIME LOG 02-04-2013
  If UPDRTIMELOG(Trim(mod_reqno), "SAVE") = True Then
  Else
  GoTo MSG
  End If
         MsgBox "Record Has Been Saved Successfully", vbOKOnly
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub




Private Sub set_position()
On Error GoTo MSG
If new_flag Or mod_flag Or callstatus Then
    txt_1.Enabled = True
    txt_2.Enabled = True
    txt_3.Enabled = True
    msflex_doppler.col = 1
    txt_1.Width = msflex_doppler.CellWidth
    txt_1.Height = msflex_doppler.CellHeight
    txt_1.Left = msflex_doppler.CellLeft + msflex_doppler.Left
    txt_1.Top = msflex_doppler.CellTop + msflex_doppler.Top
    txt_1.Text = msflex_doppler.TextMatrix(1, 1)
    txt_1.Enabled = True
    msflex_doppler.col = 2
    txt_2.Width = msflex_doppler.CellWidth
    txt_2.Height = msflex_doppler.CellHeight
    txt_2.Left = msflex_doppler.CellLeft + msflex_doppler.Left
    txt_2.Top = msflex_doppler.CellTop + msflex_doppler.Top
    txt_2.Text = msflex_doppler.TextMatrix(1, 2)
    txt_2.Enabled = True
    msflex_doppler.col = 3
    txt_3.Width = msflex_doppler.CellWidth
    txt_3.Height = msflex_doppler.CellHeight
    txt_3.Left = msflex_doppler.CellLeft + msflex_doppler.Left
    txt_3.Top = msflex_doppler.CellTop + msflex_doppler.Top
    txt_3.Text = msflex_doppler.TextMatrix(1, 3)
    txt_3.Enabled = True
    txt_1.Visible = True
    txt_2.Visible = True
    txt_3.Visible = True
 End If
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub


Private Sub TXT_COMMENTS_GotFocus()
On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    If new_flag = True Then
        txt_comments.Text = "All cardiac chambers are normal in size." & Chr(13) & "The cardiac wall is normal in thickness and contractility." & Chr(13) & "There is no regional wall motion abnormality(RWMA) at rest." & Chr(13) & "The IAS and IVS are intact." & Chr(13) & "The cardiac valves are structurally and functionally normal." & Chr(13) & "No pericardial effusion." & Chr(13) & "There are no clots or vegetations seen." & Chr(13) & "LV systolic function normal"
    End If
    objDoc.Content = txt_comments.Text
    objDoc.Content.Font.Size = 12
    objWord.Visible = True
    objWord.Activate
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub txt_comments_lostfocus()
On Error GoTo MSG
    txt_comments.Text = objDoc.Content
    objDoc.Content = ""
    ' Modified By Pankaj Baggal 03/09/2004
    'objDoc.SaveAs "c:\DISREPORTSEC\frozen.doc"
    objDoc.Close (wdDoNotSaveChanges)
    objWord.quit
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub

Private Sub txt_summary_gotfocus()
On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    objDoc.Content = txt_summary.Text
    
    objWord.Visible = True
    objWord.Activate
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub txt_summary_lostfocus()
On Error GoTo MSG
    txt_summary.Text = objDoc.Content
    objDoc.Content = ""
    ' Modified By Pankaj Baggal 03/09/2004
    'objDoc.SaveAs "c:\DISREPORTSEC\frozen.doc"
    objDoc.Close (wdDoNotSaveChanges)
    objWord.quit
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub GET_ENTER_BY()
On Error GoTo MSG
    sel_ris_qr(0) = mod_reqno
    Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
    If Not sel_ris_rs.EOF Then
        While Not sel_ris_rs.EOF
            cmb_enterby.Text = sel_ris_rs("ENTERBY") & ""
            cmb_reportby.Text = sel_ris_rs("preportby") & ""
            cmb_finalby.Text = sel_ris_rs("freportby") & ""
            cmb_reportby.Enabled = False
            cmb_enterby.Enabled = False
            cmb_finalby.Enabled = False
            sel_ris_rs.MoveNext
        Wend
        sel_ris_rs.Close
    Else
        sel_ris_rs.Close
    End If
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub


Public Sub printehco(moda As String, mskval As String, reqval As String)
    Me.Hide
    VIEW_FLAG = True
    new_flag = False
    callstatus = True
    If IsNumeric(mskval) Then
        opt_staff.Value = True
    Else
        Opt_patient.Value = True
    End If
    txt_deptcode.Text = "CAZ"
    txt_modcode.Text = moda
    MaskEdBox1.Text = mskval
    cmb_modality_Click
    cmb_reqno.AddItem req_chglen(reqval)
    cmb_reqno.Text = req_chglen(reqval)
    cmb_reqno_Click
    cmd_print_Click
End Sub

