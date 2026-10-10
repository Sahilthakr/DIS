VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{9B8DFCFF-6726-11D6-99D4-000244193CE4}#1.0#0"; "PROJECT1.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form clreport 
   Caption         =   "Biochemistry Laboratory"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   495
      Left            =   3000
      TabIndex        =   66
      Top             =   -120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
   End
   Begin VB.Frame REPORT 
      Height          =   3735
      Left            =   0
      TabIndex        =   46
      Top             =   4080
      Width           =   11775
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
         Left            =   1920
         TabIndex        =   57
         Top             =   2880
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   54
         Top             =   3240
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
         Left            =   8025
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txt_observed 
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
         Left            =   5040
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2535
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ForeColorFixed  =   192
         FormatString    =   $"clreport.frx":0000
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
         Left            =   1080
         TabIndex        =   58
         Top             =   3000
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
         Left            =   840
         TabIndex        =   56
         Top             =   3360
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
         Left            =   6345
         TabIndex        =   55
         Top             =   3360
         Width           =   1590
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   0
      TabIndex        =   41
      Top             =   7800
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
         Left            =   360
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
         Left            =   1800
         MousePointer    =   99  'Custom
         TabIndex        =   45
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
         MousePointer    =   99  'Custom
         TabIndex        =   44
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
         Left            =   4440
         MousePointer    =   99  'Custom
         TabIndex        =   5
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
         MousePointer    =   99  'Custom
         TabIndex        =   6
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
         MousePointer    =   99  'Custom
         TabIndex        =   7
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
         Left            =   10440
         MousePointer    =   99  'Custom
         TabIndex        =   43
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
         MousePointer    =   99  'Custom
         TabIndex        =   42
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
      TabIndex        =   40
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
      TabIndex        =   39
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame common 
      Height          =   3255
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   11775
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   61
         Top             =   2880
         Width           =   270
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   60
         Top             =   2520
         Width           =   270
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   59
         Top             =   2160
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   50
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
            TabIndex        =   52
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
            TabIndex        =   51
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
         TabIndex        =   21
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
         TabIndex        =   3
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
      Begin VB.ComboBox CMB_PATHNO 
         Height          =   315
         Left            =   10200
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   68
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   10560
         TabIndex        =   64
         Top             =   2880
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
         Left            =   10560
         TabIndex        =   63
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
         Left            =   10560
         TabIndex        =   62
         Top             =   2160
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   2520
         Width           =   840
      End
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
      TabIndex        =   49
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
      TabIndex        =   48
      Top             =   600
      Width           =   420
   End
   Begin VB.Label lbl_biochemlab 
      AutoSize        =   -1  'True
      Caption         =   "CLINICAL  LABORATORY SERVICES - REPORT "
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
      Left            =   2520
      TabIndex        =   47
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "clreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Dim STR As String
''Dim testorder As String
''Dim sel_deptqr As rdoQuery
''Dim sel_deptrs As rdoResultset
''Dim SELDEPTRSFLAG1 As Boolean
''Dim sel_moadlityqr As rdoQuery
''Dim sel_modalityrs As rdoResultset
''Dim SEL_MODALITYFLAG1 As Boolean
''Dim sel_casno As rdoQuery
''Dim req_qr As rdoQuery
''Dim req_rs As rdoResultset
''Dim REQFLAG1 As Boolean
''Dim SELTA101_QR As rdoQuery
''Dim SELTA101_RS As rdoResultset
''Dim selta101flag1 As Boolean
''Dim repdr_qr As rdoQuery
''Dim repdr_rs As rdoResultset
''Dim REPDRFLAG As Boolean
''Dim REFDR_QR As rdoQuery
''Dim REFDR_RS As rdoResultset
''Dim REFDRFLAG1 As Boolean
''Dim selma104_qr As rdoQuery
''Dim selma104_rs As rdoResultset
''Dim SELMA104FLAG1 As Boolean
''Dim SELPATHDREQUISITN_QR As rdoQuery
''Dim SELPATHDREQUISITN_RS As rdoResultset
''Dim SELPATHDREQUISITNFALG1 As Boolean
''Dim selma302_qr As rdoQuery
''Dim selma302_rs As rdoResultset
''Dim SELMA302FLAG1 As Boolean
''Dim selta1012_qr As rdoQuery
''Dim selta1012_rs As rdoResultset
''Dim SELTA1012FLAG1 As Boolean
'''Dim SELEMPLIB_QR As rdoQuery
'''Dim SELEMPLIB_RS As rdoResultset
''Dim i As Integer
'''Dim insdfsrep_qr As rdoQuery
'''Dim insdgrossdesc_qr As rdoQuery
'''Dim insdfs_qr As rdoQuery
''Dim SELRREQUISITN_QR As rdoQuery
''Dim SELRREQUISITN_RS As rdoResultset
''Dim SELRREQUISITNFLAG1 As Boolean
'''Dim selempbasinfo1_qr As rdoQuery
'''Dim selempbasinfo1_rs As rdoResultset
''Dim updrrequisitn_qr As rdoQuery
''Dim updrrequisitn1_qr As rdoQuery
''Dim INSDADDENDUM_QR As rdoQuery
''Dim seldrequisitn_qr As rdoQuery
''Dim seldrequisitn_rs As rdoResultset
''Dim SELDREQUISITNFLAG1 As Boolean
''
''
''Dim NEWFLAG As Boolean
''Dim VIEWFLAG As Boolean
''Dim flagsave As Integer
''Dim ADDENDUMFALG As Boolean
''
''Dim seldissecentr1_qr As rdoQuery
''Dim seldissecentr1_rs As rdoResultset
''Dim seldissecentrflag1 As Boolean
''
'''Dim seldissecprov1_qr As rdoQuery
'''Dim seldissecprov1_rs As rdoResultset
'''Dim seldissecprovflag1 As Boolean
''
''Dim seldissecfin1_qr As rdoQuery
''Dim seldissecfin1_rs As rdoResultset
''Dim seldissecfinflag1 As Boolean
''
''Dim SELDREQSTAFFDQR As rdoQuery
''Dim SELDREQSTAFFDRS As rdoResultset
''Dim SELDREQSTAFFD1FLAG As Boolean
''Dim upddreqnoqr As rdoQuery
''Dim SELDMULPROCQR As rdoQuery
''Dim SELDMULPROCRS As rdoResultset
''Dim SELDMULPROCFLAG As Boolean
''Dim selMA101QR As rdoQuery
''Dim SELMA101RS As rdoResultset
''Dim SELMA101FLAG As Boolean
''Dim SELDTEMPLETQR As rdoQuery
''Dim SELDTEMPLETRS As rdoResultset
''Dim SELDTEMPLETFLAG As Boolean
''Dim INSDMPROCEDURQR As rdoQuery
''Dim UPDDMPROCEDURQR As rdoQuery
''Dim SELDMA101QR As rdoQuery
''Dim SELDMA101RS As rdoResultset
''Dim SELDMA101FLAG As Boolean
''Dim SELDPROCEDURQR As rdoQuery
''Dim SELDPROCEDURRS As rdoResultset
''Dim SELDPROCEDURFLAG As Boolean
''Dim SELDBCPROSUBQR As rdoQuery
''Dim SELDBCPROSUBRS As rdoResultset
''Dim SELDBCPROSUBFLAG As Boolean
''
''Dim INSDMBCCOMMQR As rdoQuery
''Dim UPDDMBCCOMMQR As rdoQuery
''Dim SELDMBCCOMMQR As rdoQuery
''Dim SELDMBCCOMMRS As rdoResultset
''Dim SELDMBCCOMMFLAG As Boolean
'''Dim updrrequisitn1_qr As rdoQuery
''Dim SELDREQNOQR As rdoQuery
''Dim SELDREQNORS As rdoResultset
''Dim SELDREQNOFLAG As Boolean
''
''Dim SELDBCALKPHSQR As rdoQuery
''Dim SELDBCALKPHSRS As rdoResultset
''Dim SELDBCALKPHSFLAG As Boolean
''Dim RNAGE As String
''Dim SS As String
''Private Sub cmb_dept_Click()
''Dim deptno As String
''Dim dept As Integer
''Dim l As Integer
'On Error GoTo msg:
''    If Trim(cmb_dept) = "" Then
''        Exit Sub
''    End If
''
''    dept = InStr(cmb_dept.Text, "$")
''    txt_deptcode.Text = Mid(cmb_dept.Text, dept + 1, 3)
''    If SEL_MODALITYFLAG1 = False Then
''        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
''        Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
''        SEL_MODALITYFLAG1 = True
''    Else
''        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
''        sel_modalityrs.MoreResults
''        sel_modalityrs.Requery
''    End If
''    If sel_modalityrs.RowCount > 0 Then
''        cmb_modality.clear
''        While Not sel_modalityrs.EOF
''            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
''            sel_modalityrs.MoveNext
''        Wend
''    End If
''Exit Sub
''msg:
''MsgBox Error(Err)
''End Sub
''
''Private Sub cmb_dept_GotFocus()
'On Error GoTo msg:
''Dim ind_cbodept As Integer
''If Trim(cmb_dept) = "" Then
''    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
''        If Trim(cmb_dept.List(ind_cbodept)) = "BIOCHEMISTRY$FEC" Then
''            cmb_dept.ListIndex = ind_cbodept
''        End If
''    Next ind_cbodept
''    End If
''Exit Sub
''msg:
''MsgBox Error(Err)
''End Sub
''Private Sub cmb_modality_click()
''Dim modality As String
''Dim moda As Integer
''Dim l As Integer
'On Error GoTo msg:
''If Trim(cmb_modality) = "" Then
''    Exit Sub
''End If
''moda = InStr(cmb_modality, "$")
''l = Len(cmb_modality.Text)
''txt_modcode.Text = Right(cmb_modality.Text, (l - moda))
''Exit Sub
''msg:
''MsgBox Error(Err)
''End Sub
''
''Private Sub cmb_modality_GotFocus()
'On Error GoTo msg:
''
''Dim ind_cbodept As Integer
''If Trim(cmb_modality) = "" Then
''    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
''        If Trim(cmb_modality.List(ind_cbodept)) = "BIOCHEMISTRY$BC" Then
''            cmb_modality.ListIndex = ind_cbodept
''        End If
''    Next ind_cbodept
''    End If
''Exit Sub
''msg:
''MsgBox Error(Err)
''End Sub
''
''Private Sub cmb_reqno_Click()
''Dim prpathno As String
''Dim ADDENDUM As String
'On Error GoTo msg:
''If Opt_patient.Value = True Then
''        If NEWFLAG = True Then
''            NEWFLAG = False
''            If selta101flag1 = False Then
''                SELTA101_QR(0) = MaskEdBox1.FormattedText
''                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
''                selta101flag1 = True
''            Else
''                SELTA101_QR(0) = MaskEdBox1.FormattedText
''                SELTA101_RS.MoreResults
''                SELTA101_RS.Requery
''            End If
''            If SELTA101_RS.RowCount = 0 Then
''                Exit Sub
''            End If
''
''            txt_sex.Text = SELTA101_RS("sex")
''            txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
''            txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
''            txt_pridoc.Text = SELTA101_RS("ud_name")
''            txt_category.Text = SELTA101_RS("pat_catg")
''
''            If REFDRFLAG1 = False Then
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
''                REFDRFLAG1 = True
''            Else
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                REFDR_RS.MoreResults
''                REFDR_RS.Requery
''            End If
''            If REFDR_RS.RowCount > 0 Then
''                txt_refdoc.Text = REFDR_RS("REFBYDOC")
''            End If
''            If SELMA104FLAG1 = False Then
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
''                SELMA104FLAG1 = True
''            Else
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                selma104_rs.MoreResults
''                selma104_rs.Requery
''            End If
''
''            If selma104_rs.RowCount > 0 Then
''                txt_ward.Visible = True
''                txt_bedno.Visible = True
''                lbl_ward.Visible = True
''                lbl_bed.Visible = True
''                txt_ward.Text = selma104_rs("WARD_NAME")
''                txt_bedno.Text = selma104_rs("bed_no")
''                txt_status.Text = "IP"
''            Else
''                txt_status.Text = "OP"
''                txt_ward.Visible = False
''                txt_bedno.Visible = False
''                lbl_ward.Visible = False
''                lbl_bed.Visible = False
''            End If
''
''            If SELPATHDREQUISITNFALG1 = False Then
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
''                SELPATHDREQUISITNFALG1 = True
''            Else
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                SELPATHDREQUISITN_RS.MoreResults
''                SELPATHDREQUISITN_RS.Requery
''            End If
''            If SELPATHDREQUISITN_RS.RowCount > 0 Then
''                txt_pathno.Text = SELPATHDREQUISITN_RS("PATHNO")
''                If SELPATHDREQUISITN_RS.EOF Then
''                    txt_relclnhis.Text = " " & SELPATHDREQUISITN_RS("NMETERIALS")
''                End If
''                'txt_refdoc.Text = SELPATHDREQUISITN_RS("REFDOC")
''
''            End If
''            If SELTA1012FLAG1 = False Then
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
''                SELTA1012FLAG1 = True
''            Else
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                selta1012_rs.MoreResults
''                selta1012_rs.Requery
''            End If
''            If selta1012_rs.RowCount > 0 Then
''               ' txt_procedure.Text = selta1012_rs("prvdiagdes")
''            End If
''
''           If SELDMULPROCFLAG = False Then
''                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
''                Set SELDMULPROCRS = SELDMULPROCQR.OpenResultset(1, 3)
''                SELDMULPROCFLAG = True
''           Else
''                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
''                SELDMULPROCRS.MoreResults
''                SELDMULPROCRS.Requery
''           End If
''           If SELDMULPROCRS.RowCount > 0 Then
''
''                MSFlexGrid1.Rows = 1
''                i = 1
''
''                While Not SELDMULPROCRS.EOF
''                   If SELDMULPROCRS("pro_code") = "F849" Or SELDMULPROCRS("PRO_CODE") = "F854" Or SELDMULPROCRS("PRO_CODE") = "F860" Or SELDMULPROCRS("PRO_CODE") = "F865" Or SELDMULPROCRS("PRO_CODE") = "F875" Or SELDMULPROCRS("PRO_CODE") = "F880" Or SELDMULPROCRS("PRO_CODE") = "F802" Then
''                       If SELDBCPROSUBFLAG = False Then
''                            SELDBCPROSUBQR(0) = SELDMULPROCRS("pro_code")
''                            Set SELDBCPROSUBRS = SELDBCPROSUBQR.OpenResultset(1, 3)
''                            SELDBCPROSUBFLAG = True
''                        Else
''                            SELDBCPROSUBQR(0) = SELDMULPROCRS("pro_code")
''                            SELDBCPROSUBRS.MoreResults
''                            SELDBCPROSUBRS.Requery
''                        End If
''                        If SELDBCPROSUBRS.RowCount > 0 Then
''                            While Not SELDBCPROSUBRS.EOF
''                                If SELMA101FLAG = False Then
''                                    selMA101QR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
''                                    SELMA101FLAG = True
''                                Else
''                                    selMA101QR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    SELMA101RS.MoreResults
''                                    SELMA101RS.Requery
''                                End If
''                                If SELDTEMPLETFLAG = False Then
''                                    SELDTEMPLETQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                                    SELDTEMPLETFLAG = True
''                                Else
''                                    SELDTEMPLETQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    SELDTEMPLETRS.MoreResults
''                                    SELDTEMPLETRS.Requery
''                                End If
''
''                                '******
''                                If SELDBCPROSUBRS("PROCODESUB") = "F869" Then
''                                    ALKAGE
''                                    If SELDBCALKPHSFLAG = False Then
''                                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                        SELDBCALKPHSQR(1) = RNAGE
''                                        SELDBCALKPHSQR(2) = SS
''                                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
''                                        SELDBCALKPHSFLAG = True
''                                    Else
''                                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                        SELDBCALKPHSQR(1) = RNAGE
''                                        SELDBCALKPHSQR(2) = SS
''                                        'Set SELBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
''                                        SELDBCALKPHSRS.MoreResults
''                                        SELDBCALKPHSRS.Requery
''                                    End If
''                                    If SELDTEMPLETRS.RowCount > 0 Then
''                                    If SELMA101RS.RowCount > 0 Then
''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
''                                        Else
''                                            SELDTEMPLETRS.MoveNext
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
''                                        End If
''                                        i = i + 1
''                                    End If
''                                    End If
''                               Else
''                                    If SELDTEMPLETRS.RowCount > 0 Then
''                                    If SELMA101RS.RowCount > 0 Then
''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        Else
''                                            SELDTEMPLETRS.MoveNext
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        End If
''                                        i = i + 1
''                                    End If
''                                    End If
''                                End If
'''                                '******
'''                                If SELDTEMPLETRS.RowCount > 0 Then
'''                                    If SELMA101RS.RowCount > 0 Then
'''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                        Else
'''                                            SELDTEMPLETRS.MoveNext
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                        End If
'''                                        i = i + 1
'''                                    End If
'''                                End If
''
''                                SELDBCPROSUBRS.MoveNext
''                            Wend
''                        End If
''                   Else
''                            If SELMA101FLAG = False Then
''                                selMA101QR(0) = SELDMULPROCRS("PRO_CODE")
''                                Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
''                                SELMA101FLAG = True
''                            Else
''                                selMA101QR(0) = SELDMULPROCRS("PRO_CODE")
''                                SELMA101RS.MoreResults
''                                SELMA101RS.Requery
''                            End If
''                            If SELDTEMPLETFLAG = False Then
''                                SELDTEMPLETQR(0) = SELDMULPROCRS("PRO_CODE")
''                                Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                                SELDTEMPLETFLAG = True
''                            Else
''                                SELDTEMPLETQR(0) = SELDMULPROCRS("PRO_CODE")
''                                SELDTEMPLETRS.MoreResults
''                                SELDTEMPLETRS.Requery
''                            End If
''                            '*******
''                                If SELDMULPROCRS("PRO_CODE") = "F869" Then
''                                    ALKAGE
''                                    If SELDBCALKPHSFLAG = False Then
''                                        SELDBCALKPHSQR(0) = SELDMULPROCRS("PRO_CODE")
''                                        SELDBCALKPHSQR(1) = RNAGE
''                                        SELDBCALKPHSQR(2) = SS
''                                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
''                                        SELDBCALKPHSFLAG = True
''                                    Else
''                                        SELDBCALKPHSQR(0) = SELDMULPROCRS("PRO_CODE")
''                                        SELDBCALKPHSQR(1) = RNAGE
''                                        SELDBCALKPHSQR(2) = SS
''                                        SELDBCALKPHSRS.MoreResults
''                                        SELDBCALKPHSRS.Requery
''                                    End If
''                                    If SELDBCALKPHSRS.RowCount > 0 Then
''                                        If SELMA101RS.RowCount > 0 Then
''                                            If SELDBCALKPHSRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDBCALKPHSRS("SEX")) = "" Then
''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
''                                            Else
''                                                SELDBCALKPHSRS.MoveNext
''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
''                                            End If
''                                            i = i + 1
''                                        End If
''                                    End If
''                                Else
''                                        If SELDTEMPLETRS.RowCount > 0 Then
''                                            If SELMA101RS.RowCount > 0 Then
''                                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                                Else
''                                                    SELDTEMPLETRS.MoveNext
''                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                                End If
''                                                i = i + 1
''                                            End If
''                                        End If
''                                End If
''                            '*******
'''                        If SELDTEMPLETRS.RowCount > 0 Then
'''                            If SELMA101RS.RowCount > 0 Then
'''                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
'''                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                Else
'''                                    SELDTEMPLETRS.MoveNext
'''                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                End If
'''                                i = i + 1
'''                            End If
'''                        End If
''                    End If
''                    SELDMULPROCRS.MoveNext
''
''                Wend
''                MSFlexGrid1.Row = 1
''                MSFlexGrid1.Col = 2
''                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
''                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''                txt_observed.Visible = True
''           End If
''
''        '    seldrequisitn_qr(0) = MaskEdBox1.FormattedText
''        '    seldrequisitn_rs.MoreResults
''        '    seldrequisitn_rs.Requery
''        '    If seldrequisitn_rs.RowCount > 0 Then
''        '     prpathno = ""
''        '     While Not seldrequisitn_rs.EOF
''        '     prpathno = prpathno & " " & seldrequisitn_rs("pathno")
''        '     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
''        '     seldrequisitn_rs.MoveNext
''        '     Wend
''        '     txt_prepathreqno.Text = prpathno
''        '    End If
''        '    txt_addendum.Enabled = False
''        '    cmb_addendumby.Enabled = False
''
''
''
''        If seldissecentrflag1 = False Then
''            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
''            seldissecentrflag1 = True
''        Else
''            seldissecentr1_rs.MoreResults
''            seldissecentr1_rs.Requery
''        End If
''        If seldissecentr1_rs.RowCount > 0 Then
''            cmb_enterby.clear
''            While Not seldissecentr1_rs.EOF
''                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
''                seldissecentr1_rs.MoveNext
''            Wend
''        End If
''
''        If seldissecfinflag = False Then
''            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
''            seldissecfinflag1 = True
''        Else
''            seldissecfin1_rs.MoreResults
''            seldissecfin1_rs.Requery
''        End If
''        If seldissecfin1_rs.RowCount > 0 Then
''            cmb_finalby.clear
''            While Not seldissecfin1_rs.EOF
''                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
''                seldissecfin1_rs.MoveNext
''            Wend
''        End If
''
''
''
''        End If
''            If VIEWFLAG = True Then
''
''
''
''            VIEWFLAG = False
''            If selta101flag1 = False Then
''                SELTA101_QR(0) = MaskEdBox1.FormattedText
''                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
''                selta101flag1 = True
''            Else
''                SELTA101_QR(0) = MaskEdBox1.FormattedText
''                SELTA101_RS.MoreResults
''                SELTA101_RS.Requery
''            End If
''            If SELTA101_RS.RowCount = 0 Then
''                Exit Sub
''            End If
''
''            txt_sex.Text = SELTA101_RS("sex")
''            txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
''            txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
''            txt_pridoc.Text = SELTA101_RS("ud_name")
''            txt_category.Text = SELTA101_RS("pat_catg")
''
''            If REFDRFLAG1 = False Then
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
''                REFDRFLAG1 = True
''            Else
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                REFDR_RS.MoreResults
''                REFDR_RS.Requery
''            End If
''            If REFDR_RS.RowCount > 0 Then
''                txt_refdoc.Text = REFDR_RS("REFBYDOC")
''                If REFDR_RS("COMMITFLAG") = "Y" Then
''                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
''                    ADDENDUMFALG = True
''                End If
''                cmb_enterby.Text = REFDR_RS("ENTERBY")
''                cmb_finalby.Text = REFDR_RS("FREPORTBY")
'''                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
'''                    cmb_finalby.Text = ""
'''                Else
'''                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
'''                End If
''            End If
''
''            If SELMA104FLAG1 = False Then
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
''                SELMA104FLAG1 = True
''            Else
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                selma104_rs.MoreResults
''                selma104_rs.Requery
''            End If
''
''
''            If selma104_rs.RowCount > 0 Then
''                txt_ward.Visible = False
''                txt_bedno.Visible = False
''                lbl_ward.Visible = False
''                lbl_bed.Visible = False
''                txt_ward.Text = selma104_rs("WARD_NAME")
''                txt_bedno.Text = selma104_rs("bed_no")
''                txt_status.Text = "IP"
''            Else
''                txt_ward.Visible = False
''                txt_bedno.Visible = False
''                lbl_ward.Visible = False
''                lbl_bed.Visible = False
''                txt_status.Text = "OP"
''            End If
''
''
''       If SELPATHDREQUISITNFALG1 = False Then
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
''                SELPATHDREQUISITNFALG1 = True
''            Else
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                SELPATHDREQUISITN_RS.MoreResults
''                SELPATHDREQUISITN_RS.Requery
''            End If
''            If SELPATHDREQUISITN_RS.RowCount > 0 Then
''                txt_pathno.Text = SELPATHDREQUISITN_RS("PATHNO")
''                If SELPATHDREQUISITN_RS.EOF = False Then
''                    txt_relclnhis.Visible = True
''
''                    txt_relclnhis.Text = " " & SELPATHDREQUISITN_RS("NMETERIALS")
''                End If
''            End If
''            If SELTA1012FLAG1 = False Then
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
''                SELTA1012FLAG1 = True
''            Else
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                selta1012_rs.MoreResults
''                selta1012_rs.Requery
''            End If
''            If selta1012_rs.RowCount > 0 Then
''                'txt_procedure.Text = selta1012_rs("prvdiagdes")
''            End If
''            '**** filling the flex grid
''            If SELDPROCEDURFLAG = False Then
''                SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
''                Set SELDPROCEDURRS = SELDPROCEDURQR.OpenResultset(1, 3)
''                SELDPROCEDURFLAG = True
''           Else
''                SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
''                SELDPROCEDURRS.MoreResults
''                SELDPROCEDURRS.Requery
''           End If
''           If SELDPROCEDURRS.RowCount > 0 Then
''
''''                MSFlexGrid1.Rows = 1
''''                i = 1
''''
''''                While Not SELDPROCEDURRS.EOF
''''                    If SELMA101FLAG = False Then
''''                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
''''                        SELMA101FLAG = True
''''                    Else
''''                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        SELMA101RS.MoreResults
''''                        SELMA101RS.Requery
''''                    End If
''''                    If SELDTEMPLETFLAG = False Then
''''                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''''                        SELDTEMPLETFLAG = True
''''                    Else
''''                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        SELDTEMPLETRS.MoreResults
''''                        SELDTEMPLETRS.Requery
''''                    End If
''''                    If SELDTEMPLETRS.RowCount > 0 Then
''''                        If SELMA101RS.RowCount > 0 Then
''''                            If SELDTEMPLETRS("SEX") = "M" Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''''                                 MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & SELDPROCEDURRS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                            Else
''''                                SELDTEMPLETRS.MoveNext
''''                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & SELDPROCEDURRS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                            End If
''''                        End If
''''
''''                            i = i + 1
''''
''''                    End If
''''                    SELDPROCEDURRS.MoveNext
''''
''''                Wend
'''''                MSFlexGrid1.Row = 1
'''''                MSFlexGrid1.Col = 2
'''''                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'''''                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''''                txt_observed.Visible = False
''                    'MSFlexGrid1.Rows = 1
''                    'MSFlexGrid1.Row = 1
''                    i = 1
''                    srno = 1
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
''
''                    If SELDTEMPLETRS.RowCount > 0 Then
''                            If SELMA101RS.RowCount > 0 Then
''                            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''                             MSFlexGrid1.Row = srno
''                             MSFlexGrid1.Col = 0
''                            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
''                                MSFlexGrid1.TextMatrix(i, 0) = i
''                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101RS("SER_DET"))
''                                If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
''                                    MSFlexGrid1.Row = srno
''                                     MSFlexGrid1.Col = 2
''                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
''                                Else
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
''
''                                      MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = RGB(100, 255, 150)
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
''
''                                       MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "" Then
''
''                                        MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                End If
''                                J = 0
''                                J = InStr(SELDPROCEDURRS("observalue"), "*")
''                                If J = 0 Then
''                                    MSFlexGrid1.TextMatrix(i, 2) = SELDPROCEDURRS("OBSERVALUE")
''                                Else
''                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, J - 1)
''                                End If
''                                MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                MSFlexGrid1.TextMatrix(i, 4) = SELDPROCEDURRS("pvalue")
''                            Else
''                                SELDTEMPLETRS.MoveNext
''                                MSFlexGrid1.TextMatrix(i, 0) = i
''                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101RS("SER_DET"))
''                                If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
''                                    MSFlexGrid1.Row = srno
''                                     MSFlexGrid1.Col = 2
''                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
''                                Else
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
''
''                                      MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = RGB(150, 255, 100)
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
''
''                                       MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "" Then
''
''                                        MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                End If
''
''                                J = 0
''                                J = InStr(SELDPROCEDURRS("observalue"), "*")
''                                If J = 0 Then
''                                    MSFlexGrid1.TextMatrix(i, 2) = Trim(SELDPROCEDURRS("OBSERVALUE"))
''                                Else
''                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, J - 1)
''                                End If
''                                MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                MSFlexGrid1.TextMatrix(i, 4) = SELDPROCEDURRS("pvalue")
''                            End If
''                        End If
''                           srno = srno + 1
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
''
''           End If
''
''           If SELDMBCCOMMFLAG = False Then
''                SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
''                Set SELDMBCCOMMRS = SELDMBCCOMMQR.OpenResultset(1, 3)
''                SELDMBCCOMMFLAG = True
''            Else
''                SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
''                SELDMBCCOMMRS.MoreResults
''                SELDMBCCOMMRS.Requery
''            End If
''            If SELDMBCCOMMRS.RowCount > 0 Then
''                If SELDMBCCOMMRS.EOF = False Then
''                    txt_comments.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
''                End If
''            End If
''            '**** filling the flex grid
'''            If SELDREQUISITNFLAG1 = False Then
'''                seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'''                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'''            Else
'''                seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'''                seldrequisitn_rs.MoreResults
'''                seldrequisitn_rs.Requery
'''            End If
'''            If seldrequisitn_rs.RowCount > 0 Then
'''            'txt_refdoc.Text = seldrequisitn_rs("REFDOC")
'''             prpathno = ""
'''             While Not seldrequisitn_rs.EOF
'''                prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'''                'CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
'''                seldrequisitn_rs.MoveNext
'''             Wend
'''                'txt_prepathreqno.Text = Trim(prpathno)
'''            End If
''
'''            If seldissecfinflag = False Then
'''            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'''            seldissecfinflag1 = True
'''        Else
'''            seldissecfin1_rs.MoreResults
'''            seldissecfin1_rs.Requery
'''        End If
'''        If seldissecfin1_rs.RowCount > 0 Then
'''            cmb_finalized.clear
'''            While Not seldissecfin1_rs.EOF
'''                cmb_finalized.AddItem Trim(seldissecfin1_rs("CCNAME"))
'''                seldissecfin1_rs.MoveNext
'''            Wend
'''            cmb_finalized.Text = REFDR_RS("FREPORTBY")
'''        End If
'''            cmb_reqno.Enabled = False
'''            cmb_dept.Enabled = False
'''            cmb_modality.Enabled = False
'''            MaskEdBox1.Enabled = False
'''
'''            'Msftest.Enabled = True
''        End If
''End If
''If opt_staff.Value = True Then
''    STAFF
''
''
''End If
''Exit Sub
''msg:
''MsgBox Err.Description
''End Sub
''
''Private Sub cmd_back_Click()
''Me.Hide
''DIS_MAIN.Show
''btnenabl ("INI")
''field_clear
''VIEWFLAG = False
''NEWFLAG = False
''BCCLEAR
''
''End Sub
''
''Public Sub btnenabl(cap As String)
''If cap = "NEW" Then
''    cmd_save.Enabled = True
''    cmd_cancel.Enabled = True
''    cmd_back.Enabled = True
''    CMD_NEW.Enabled = False
''    cmd_modify.Enabled = False
''    cmd_print.Enabled = False
''    cmd_commit.Enabled = False
''    cmd_view.Enabled = False
''
''ElseIf cap = "VIEW" Then
''    cmd_cancel.Enabled = True
''    cmd_back.Enabled = True
''    cmd_modify.Enabled = True
''    CMD_NEW.Enabled = False
''    cmd_save.Enabled = False
''    cmd_print.Enabled = True
''    cmd_commit.Enabled = False
''    cmd_view.Enabled = False
''
''ElseIf cap = "SAVE" Then
''    cmd_cancel.Enabled = True
''    cmd_back.Enabled = True
''    cmd_modify.Enabled = False
''    CMD_NEW.Enabled = False
''    cmd_save.Enabled = False
''    cmd_print.Enabled = True
''    cmd_commit.Enabled = True
''    cmd_view.Enabled = True
''
''ElseIf cap = "INI" Then
''    cmd_cancel.Enabled = False
''    cmd_back.Enabled = True
''    cmd_modify.Enabled = False
''    CMD_NEW.Enabled = True
''    cmd_save.Enabled = False
''    cmd_print.Enabled = False
''    cmd_commit.Enabled = False
''    cmd_view.Enabled = True
''
''ElseIf cap = "COMMIT" Then
''    cmd_save.Enabled = False
''    cmd_cancel.Enabled = False
''    cmd_back.Enabled = True
''    CMD_NEW.Enabled = True
''    cmd_modify.Enabled = False
''    cmd_print.Enabled = True
''    cmd_commit.Enabled = False
''    cmd_view.Enabled = True
''
''ElseIf cap = "print" Then
''     cmd_save.Enabled = False
''     cmd_back.Enabled = True
''     cmd_modify.Enabled = False
''     cmd_print.Enabled = False
''     cmd_cancel.Enabled = True
''     CMD_NEW.Enabled = False
''
''End If
''End Sub
''Private Sub cmd_commit_Click()
''Dim i As Integer
'''If Len(Trim(cmb_finalby)) = 0 Then
'''    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
'''    cmb_finalby.SetFocus
'''    Exit Sub
'''End If
''i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
''If i = vbYes Then
'On Error GoTo msg:
'' rdb.BeginTrans
''    updrrequisitn1_qr(0) = UCase(cmb_finalby.Text)
''    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    updrrequisitn1_qr(2) = "F"
''    updrrequisitn1_qr(3) = "Y"
''    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
''    updrrequisitn1_qr.Execute
''    MsgBox "Record is Committed Successfully."
''End If
''rdb.CommitTrans
''btnenabl ("INI")
''
''cmd_print.Enabled = True
''Exit Sub
''msg:
''rdb.RollbackTrans
''MsgBox Err.Description
''End Sub
''
''Private Sub cmd_modify_Click()
''Dim STR As String
''btnenabl ("NEW")
''FLDENABL ("MODIFY")
''flagsave = 2
''
''If REFDRFLAG = False Then
''    REFDR_QR(0) = MaskEdBox1.FormattedText
''    REFDR_QR(1) = reqno
''    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
''    REFDRFLAG = True
''Else
''    REFDR_QR(0) = MaskEdBox1.FormattedText
''    REFDR_QR(1) = reqno
''    REFDR_RS.MoreResults
''    REFDR_RS.Requery
''End If
''
''      If seldissecentrflag1 = False Then
''            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
''            seldissecentrflag1 = True
''        Else
''            seldissecentr1_rs.MoreResults
''            seldissecentr1_rs.Requery
''        End If
''        If seldissecentr1_rs.RowCount > 0 Then
''        STR = ""
''            STR = cmb_enterby.Text
''            cmb_enterby.clear
''            While Not seldissecentr1_rs.EOF
''                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
''                seldissecentr1_rs.MoveNext
''            Wend
''            cmb_enterby.Text = STR
''        End If
''
''        If seldissecfinflag = False Then
''            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
''            seldissecfinflag1 = True
''        Else
''            seldissecfin1_rs.MoreResults
''            seldissecfin1_rs.Requery
''        End If
''        If seldissecfin1_rs.RowCount > 0 Then
''            STR = ""
''            STR = cmb_finalby.Text
''            cmb_finalby.clear
''            While Not seldissecfin1_rs.EOF
''                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
''                seldissecfin1_rs.MoveNext
''            Wend
''            cmb_finalby.Text = STR
''        End If
''cmb_enterby.Enabled = True
''cmb_finalby.Enabled = True
''
''
''MSFlexGrid1.Row = 1
''MSFlexGrid1.Col = 2
''txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
''txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''txt_observed.Visible = True
''txt_comments.Enabled = True
''End Sub
''
''Private Sub cmd_new_Click()
'On Error GoTo msg:
''btnenabl ("NEW")
''field_clear
''TXT_CLEAR
''BCCLEAR
''
''If DATETIMEFLAG = False Then
''   datetime
''   DATETIMEFLAG = True
''Else
''    datetime
''End If
''txt_date.Text = DATE11
''txt_time.Text = time1
''If SELDEPTRSFLAG1 = False Then
''    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
''    SELDEPTRSFLAG1 = True
''Else
''    sel_deptrs.MoreResults
''    sel_deptrs.Requery
''End If
''If sel_deptrs.RowCount > 0 Then
''    sel_deptrs.MoveFirst
''    While Not sel_deptrs.EOF
''        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
''        sel_deptrs.MoveNext
''        Wend
''   End If
''Opt_patient.Value = True
''txt_observed.Visible = False
''flagsave = 1
''VIEWFLAG = 0
''NEWFLAG = True
''FLDENABL ("ENA")
''cmb_dept.SetFocus
''MaskEdBox1.Enabled = True
''Exit Sub
''msg:
''MsgBox Err.Description
''End Sub
''
''Private Sub cmd_cancel_Click()
''btnenabl ("INI")
''field_clear
''VIEWFLAG = False
''NEWFLAG = False
''BCCLEAR
''
''
''End Sub
''
''Public Sub field_clear()
''Dim obj As Object
''cmb_dept.ListIndex = -1
''cmb_modality.ListIndex = -1
''txt_deptcode.Text = " "
''cmb_reqno.ListIndex = -1
''cmb_reqno.clear
''MaskEdBox1.Mask = ""
''MaskEdBox1.Text = ""
''MaskEdBox1.Mask = ">??/#####"
''txt_name.Text = " "
''txt_sex.Text = " "
''txt_age.Text = " "
''txt_category.Text = " "
''txt_status.Text = " "
''txt_pridoc.Text = " "
''txt_ward.Text = " "
''txt_bedno.Text = " "
''txt_refdoc.Text = " "
''txt_relclnhis.Text = " "
''End Sub
''
''Public Sub TXT_CLEAR()
''Dim obj As Object
''For Each obj In biochemlab_report.Controls
''If TypeOf obj Is TextBox Then
''obj = ""
''End If
''Next
''End Sub
''
''
''Public Sub FLDENABL(EN As String)
''If EN = "DIS" Then
''        cmb_dept.Enabled = False
''        cmb_modality.Enabled = False
''        MaskEdBox1.Enabled = False
''        cmb_reqno.Enabled = False
''        txt_pathno.Enabled = False
''        txt_name.Enabled = False
''        txt_sex.Enabled = False
''        txt_age.Enabled = False
''        txt_category.Enabled = False
''        txt_status.Enabled = False
''        txt_pridoc.Enabled = False
''        txt_ward.Enabled = False
''        txt_bedno.Enabled = False
''        txt_modcode.Enabled = False
''        txt_refdoc.Enabled = False
''        txt_relclnhis.Visible = False
''        txt_observed.Visible = False
''        txt_comments.Enabled = False
''        cmb_enterby.Enabled = False
''        cmb_finalby.Enabled = False
''End If
''If EN = "ENA" Then
''        cmb_dept.Enabled = True
''        cmb_modality.Enabled = True
''        MaskEdBox1.Enabled = True
''        cmb_reqno.Enabled = True
''        txt_pathno.Enabled = False
''        txt_name.Enabled = False
''        txt_sex.Enabled = False
''        txt_age.Enabled = False
''        txt_category.Enabled = False
''        txt_status.Enabled = False
''        txt_pridoc.Enabled = False
''        txt_ward.Enabled = False
''        txt_bedno.Enabled = False
''        txt_modcode.Enabled = False
''        txt_refdoc.Enabled = False
''        txt_relclnhis.Visible = True
''        txt_observed.Visible = False
''        txt_comments.Enabled = True
''        cmb_enterby.Enabled = True
''        cmb_finalby.Enabled = True
''End If
''If EN = "MOD" Then
''        cmb_dept.Enabled = True
''        cmb_modality.Enabled = True
''        MaskEdBox1.Enabled = True
''        cmb_reqno.Enabled = True
''        txt_pathno.Enabled = True
''        txt_name.Enabled = True
''        txt_sex.Enabled = True
''        txt_age.Enabled = True
''        txt_category.Enabled = True
''        txt_status.Enabled = True
''        txt_pridoc.Enabled = True
''        txt_ward.Enabled = True
''        txt_bedno.Enabled = True
''        txt_modcode.Enabled = True
''        txt_refdoc.Enabled = True
''        txt_relclnhis.Visible = True
''        txt_observed.Visible = True
''        txt_comments.Enabled = True
''        cmb_enterby.Enabled = True
''        cmb_finalby.Enabled = True
''End If
''End Sub
''
''
''Public Sub BCCLEAR()
''TXT_CLEAR
''cmb_dept.Text = ""
''cmb_modality.Text = ""
''txt_observed.Visible = False
''MSFlexGrid1.Rows = 1
''
''End Sub
''
''Private Sub cmd_print_Click()
'On Error GoTo msg:
''
''Dim STR1 As String
''Dim STR2 As String
''Dim STR3 As String
''Dim i As Integer
'' On Error GoTo prtmsg:
''
''DoEvents
''
''Dim VAR_REQNNO As String
''
''
''        VAR_REQNNO = req_chglen(cmb_reqno.Text)
''        If REFDRFLAG1 = False Then
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
''                REFDRFLAG1 = True
''            Else
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                REFDR_RS.MoreResults
''                REFDR_RS.Requery
''            End If
''
''
''            If SELRREQUISITNFLAG1 = False Then
''                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
''                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
''                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
''                SELRREQUISITN_QR(3) = "T"
''                SELRREQUISITN_QR(4) = "F"
''                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
''                SELRREQUISITNFLAG1 = True
''            Else
''                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
''                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
''                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
''                SELRREQUISITN_QR(3) = "T"
''                SELRREQUISITN_QR(4) = "F"
''                SELRREQUISITN_RS.MoreResults
''                SELRREQUISITN_RS.Requery
''            End If
''            If SELRREQUISITN_RS.RowCount > 0 Then
''                prepath = ""
''                While Not SELRREQUISITN_RS.EOF
''                    If Trim(SELRREQUISITN_RS("reqnno")) <> req_chglen(Trim(cmb_reqno.Text)) Then
''                        prepath = prepath & req_chglen(SELRREQUISITN_RS("REQNNO")) & " "
''                    End If
''
''                    SELRREQUISITN_RS.MoveNext
''                Wend
''             End If
''             Dim pcount As Integer
''             pcount = 0
''            For i = 1 To MSFlexGrid1.Rows - 1
''                If Trim(MSFlexGrid1.TextMatrix(i, 4)) <> "" Then
''                  pcount = pcount + 1
''                End If
''            Next i
''If Opt_patient.Value = True Then
''
''    STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
''    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
''
''
''        CrystalReport1.ReportFileName = "C:\DISREPORT\DISBC.RPT"
''        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
''        If Trim(prepath) = "" Then
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
''        End If
''        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
''
''
''
''        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
''            STR1 = ""
''            STR1 = "Entered By " & Trim(cmb_enterby.Text)
''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
''         Else
''
''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
''        End If
''
''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
''        If pcount <> 0 Then
''
''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
''        Else
''            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
''        End If
''End If
'''**** STAFF
''If opt_staff.Value = True Then
''
''STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
''    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
''
''
''        CrystalReport1.ReportFileName = "C:\DISREPORT\DISBCS.RPT"
''        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
''        If Trim(prepath) = "" Then
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
''        End If
''        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
''
''
''
''        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
''            STR1 = ""
''            STR1 = "Entered By " & Trim(cmb_enterby.Text)
''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
''         Else
''
''            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
''        End If
''
''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
''        If pcount <> 0 Then
''
''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
''        Else
''            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
''        End If
''
''
''End If
'''***** STAFF
''CrystalReport1.Action = 1
''Exit Sub
''prtmsg: MsgBox Error(Err)
'''txt_msg.Text = ""
''btnenabl ("print")
''Exit Sub
''msg:
''MsgBox Error(Err)
''End Sub
''
''Private Sub cmd_save_Click()
''Dim pcount As Integer
''    pcount = 0
''    For i = 1 To MSFlexGrid1.Rows - 1
''        If Trim(MSFlexGrid1.TextMatrix(i, 4)) <> "" Then
''          pcount = pcount + 1
''        End If
''    Next i
''    If pcount <> 0 Then
''      J = MsgBox(pcount & " Panic Value, Do you want to Save? ", vbYesNo)
''        If J = vbNo Then
''         Exit Sub
''        End If
''    End If
''assign
''btnenabl ("SAVE")
'''cmd_print.Enabled = True
'''cmd_cancel.Enabled = True
''
''
'''commented by venkat
'''If pcount = 0 Then
'''    cmd_commit.Enabled = True
'''Else
'''    cmd_commit.Enabled = False
'''End If
''cmd_commit.Enabled = True
''End Sub
''
''Private Sub cmd_view_Click()
'On Error GoTo msg:
''btnenabl ("VIEW")
''FLDENABL ("VIEW")
''field_clear
''TXT_CLEAR
''If DATETIMEFLAG = False Then
''   datetime
''   DATETIMEFLAG = True
''Else
''    datetime
''End If
''txt_date.Text = DATE11
''txt_time.Text = time1
''If SEL_DEPTRSFLAG = False Then
''    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
''    SEL_DEPTFLAG = True
''Else
''    sel_deptrs.MoreResults
''    sel_deptrs.Requery
''End If
''If sel_deptrs.RowCount > 0 Then
''    sel_deptrs.MoveFirst
''    While Not sel_deptrs.EOF
''        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
''        sel_deptrs.MoveNext
''        Wend
''   End If
''
''    If seldissecentrflag1 = False Then
''            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
''            seldissecentrflag1 = True
''        Else
''            seldissecentr1_rs.MoreResults
''            seldissecentr1_rs.Requery
''        End If
''        If seldissecentr1_rs.RowCount > 0 Then
''            cmb_enterby.clear
''            While Not seldissecentr1_rs.EOF
''                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
''                seldissecentr1_rs.MoveNext
''            Wend
''        End If
''
''        If seldissecfinflag = False Then
''            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
''            seldissecfinflag1 = True
''        Else
''            seldissecfin1_rs.MoreResults
''            seldissecfin1_rs.Requery
''        End If
''        If seldissecfin1_rs.RowCount > 0 Then
''            cmb_finalby.clear
''            While Not seldissecfin1_rs.EOF
''                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
''                seldissecfin1_rs.MoveNext
''            Wend
''        End If
''
''
''cmb_dept.Enabled = True
''ADDENDUMFALG = False
''VIEWFLAG = True
''cmb_dept.SetFocus
''cmb_modality.Enabled = True
''MaskEdBox1.Enabled = True
''cmb_reqno.Enabled = True
''cmd_print.Enabled = True
''Opt_patient.Value = True
''Exit Sub
''msg:
''MsgBox Err.Description
''End Sub
''
''Private Sub Form_Load()
'On Error GoTo msg:
''btnenabl ("INI")
''NEWFLAG = False
''VIEWFLAG = False
''FLDENABL ("DIS")
''datetime
''txt_date.Text = Format(DATE11, "DD/MM/YYYY")
''txt_time.Text = time1
''If BCREPORTFLAG = False Then
''    BCREPORTFLAG = True
''''for dept selection
''STR = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='CCC' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
''Set sel_deptqr = rdb.CreateQuery("sel_deptqr1BC", STR)
''
'''Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
''
''
'''' for modality selection
''STR = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
''Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr1BC", STR)
''
''
''
'''    'for retriving reqnos
''Set req_qr = rdb.CreateQuery("reqBC", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E'AND MODALITYCD='BC' ORDER BY REQNDT desc ")
''
'''
''Set SELTA101_QR = rdb.CreateQuery("selta101_qr11BC", "select * from PATOOLSLIB.ta101 where case_no =?")
''
'''    ' FOR REFEFING DOCTOR
''    STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
''    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1BC", STR)
''
'''    ' FOR PATH NO
''    STR = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
''    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1BC", STR)
''''
'''    ' for bedno, ward retriveing
''    STR = "select * from PATOOLSLIB.ma104 where case_no =?"
''    Set selma104_qr = rdb.CreateQuery("selma104_qr11BC", STR)
'''    selma104_qr(0) = "BN/00001"
'''    Set selma104_rs = selma104_qr.OpenResultset(2, 3)
'''    'for provisional diagnosis
'''
''    STR = "select * from patoolslib.ta1012 where case_no =?"
''    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1BC", STR)
'''    selta1012_qr(0) = "BN/00001"
'''    Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)
''
''
''
''    STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
''    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1BC", STR)
''
''   STR = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
''   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1BC", STR)
'''
''   STR = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
''   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1", STR)
'''
'''   STR = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
'''   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QR1", STR)
'''
'''   STR = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
'''   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11", STR)
'''   SELDADDENDUM_QR(0) = "FSS"
'''   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'''
''   STR = "select * from dislib.drequisitn where caseno=?"
''   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1BC", STR)
''
''
''    STR = "select * from dislib.dissec where enterby='YES'AND MODALITY='BC'"
''   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1BC", STR)
''
'''   STR = "select * from dislib.dissec where PROVby='YES'AND MODALITY='SP'"
'''   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1BC", STR)
'''
''   STR = "select * from dislib.dissec where FINALby='YES'AND MODALITY='BC'"
''   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1BC", STR)
''
''    STR = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
''    Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1BC", STR)
''
''  STR = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
''  Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRBC", STR)
''
''  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE=?"
''  STR = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
''  Set selMA101QR = rdb.CreateQuery("SELMA101QRBC", STR)
''
''  STR = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET WHERE PRO_CODE=?"
''  Set SELDTEMPLETQR = rdb.CreateQuery("SELDTEMPLETQRBC", STR)
''
''  STR = "INSERT INTO DISLIB.DMPROCEDUR(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE) VALUES (?,?,?,?,?,?,?)"
''  Set INSDMPROCEDURQR = rdb.CreateQuery("INSDMPROCEDURQRBC", STR)
''
''  STR = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=? WHERE REQNO=? AND PRO_CODE=? "
''  Set UPDDMPROCEDURQR = rdb.CreateQuery("UPDDMPROCEDURQRBC", STR)
''
''  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_DET=?"
''  STR = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET=?"
''  Set SELDMA101QR = rdb.CreateQuery("SELDMA101QRBC", STR)
''
''  STR = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
''   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1bc", STR)
''
''  'STR = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
''  STR = "SELECT PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
''  Set SELDPROCEDURQR = rdb.CreateQuery("SELDMPROCEDURQRBC", STR)
''
''  STR = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
''  Set SELDBCPROSUBQR = rdb.CreateQuery("SELDBCPROSUBQRBC", STR)
''
''  STR = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES(?,?)"
''  Set INSDMBCCOMMQR = rdb.CreateQuery("INSDMBCCOMMQRBC", STR)
''
''  STR = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
''  Set UPDDMBCCOMMQR = rdb.CreateQuery("UPDDMBCCOMMQRBC", STR)
''
''  STR = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
''  Set SELDMBCCOMMQR = rdb.CreateQuery("SELDMBCCOMMQRBC", STR)
''
''  STR = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
''  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrBC", STR)
''
''  STR = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%BC%'"
''  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRBC", STR)
''
''  STR = "SELECT * FROM DISLIB.DBCALKPHS WHERE PROCODE=? AND AGE2=? AND SEX=? "
''  Set SELDBCALKPHSQR = rdb.CreateQuery("SELDBCALKPHSQRBC", STR)
''
''End If
''MSFlexGrid1.Col = 2
''txt_observed.Width = MSFlexGrid1.CellWidth
''txt_observed.Height = MSFlexGrid1.CellHeight
''txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
''txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''
''Exit Sub
''msg:
''MsgBox Err.Description
''End Sub
''
''Private Sub MaskEdBox1_LostFocus()
''Dim CASENO As String
'On Error GoTo msg:
''If Len(Trim(MaskEdBox1.Text)) = "" Then
''Exit Sub
''End If
''If Opt_patient.Value = True Then
''        CASENO = MaskEdBox1.Text
''        MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
''         If NEWFLAG = True Then
''            If REQFLAG1 = False Then
''                req_qr(0) = MaskEdBox1.FormattedText
''                Set req_rs = req_qr.OpenResultset(1, 3)
''            Else
''                req_qr(0) = MaskEdBox1.FormattedText
''                req_rs.MoreResults
''                req_rs.Requery
''          End If
''          If req_rs.RowCount > 0 Then
''                cmb_reqno.clear
''                While Not req_rs.EOF
''                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
''
''                req_rs.MoveNext
''
''                Wend
''          Else
''            cmb_reqno.clear
''
''            End If
''
'''          If SELDREQNOFLAG = False Then
'''            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'''            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'''            SELDREQNOFLAG = True
'''          Else
'''            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'''            SELDREQNORS.MoreResults
'''            SELDREQNORS.Requery
'''          End If
'''          If SELDREQNORS.RowCount > 0 Then
'''            CMB_PATHNO.clear
'''            While Not SELDREQNORS.EOF
'''                CMB_PATHNO.AddItem SELDREQNORS("PATHNO")
'''                SELDREQNORS.MoveNext
'''            Wend
'''          End If
''        End If
''
''        If VIEWFLAG = True Then
''            If SELRREQUISITNFLAG1 = False Then
''                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
''                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
''                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
''                SELRREQUISITN_QR(3) = "T"
''                SELRREQUISITN_QR(4) = "F"
''                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
''                SELRREQUISITNFLAG1 = True
''            Else
''                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
''                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
''                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
''                SELRREQUISITN_QR(3) = "T"
''                SELRREQUISITN_QR(4) = "F"
''                SELRREQUISITN_RS.MoreResults
''                SELRREQUISITN_RS.Requery
''            End If
''            If SELRREQUISITN_RS.RowCount > 0 Then
''                cmb_reqno.clear
''                While Not SELRREQUISITN_RS.EOF
''                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
''                    SELRREQUISITN_RS.MoveNext
''                Wend
''             Else
''                MsgBox " No  Requisition for this Modality"
''                btnenabl ("INI")
''                Exit Sub
''            End If
''
''        End If
''End If
''    If opt_staff.Value = True Then
''            CASENO = MaskEdBox1.Text
''        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
''         If NEWFLAG = True Then
''            If REQFLAG1 = False Then
''                req_qr(0) = MaskEdBox1.FormattedText
''                Set req_rs = req_qr.OpenResultset(1, 3)
''            Else
''                req_qr(0) = MaskEdBox1.FormattedText
''                req_rs.MoreResults
''                req_rs.Requery
''          End If
''          If req_rs.RowCount > 0 Then
''                cmb_reqno.clear
''                While Not req_rs.EOF
''                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
''
''                req_rs.MoveNext
''
''                Wend
''          Else
''            cmb_reqno.clear
''
''            End If
''
'''          If SELDREQNOFLAG = False Then
'''            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'''            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'''            SELDREQNOFLAG = True
'''          Else
'''            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'''            SELDREQNORS.MoreResults
'''            SELDREQNORS.Requery
'''          End If
'''          If SELDREQNORS.RowCount > 0 Then
'''            CMB_PATHNO.clear
'''            While Not SELDREQNORS.EOF
'''                CMB_PATHNO.AddItem SELDREQNORS("PATHNO")
'''                SELDREQNORS.MoveNext
'''            Wend
'''          End If
''        End If
''
''        If VIEWFLAG = True Then
''            If SELRREQUISITNFLAG1 = False Then
''                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
''                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
''                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
''                SELRREQUISITN_QR(3) = "T"
''                SELRREQUISITN_QR(4) = "F"
''                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
''                SELRREQUISITNFLAG1 = True
''            Else
''                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
''                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
''                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
''                SELRREQUISITN_QR(3) = "T"
''                SELRREQUISITN_QR(4) = "F"
''                SELRREQUISITN_RS.MoreResults
''                SELRREQUISITN_RS.Requery
''            End If
''            If SELRREQUISITN_RS.RowCount > 0 Then
''                cmb_reqno.clear
''                While Not SELRREQUISITN_RS.EOF
''                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
''                    SELRREQUISITN_RS.MoveNext
''                Wend
''             Else
''                MsgBox " No  Requisition for this Modality"
''                btnenabl ("INI")
''                Exit Sub
''            End If
''
''        End If
''End If
'' Exit Sub
''msg:
''MsgBox Error(Err)
''End Sub
''
''Private Sub MSFlexGrid1_Click()
''If MSFlexGrid1.Col = 2 Then
''        txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
''        txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''
''End If
''End Sub
''
''Private Sub MSFlexGrid1_GotFocus()
'''MSFlexGrid1.Col = 2
'''txt_observed.Width = MSFlexGrid1.CellWidth
'''txt_observed.Height = MSFlexGrid1.CellHeight
'''txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'''txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''End Sub
''
''Private Sub Opt_patient_Click()
''If Opt_patient.Value = True Then
''    lbl_caseno.Caption = "Case No."
''    MaskEdBox1.Mask = "??>/#####"
''    MaskEdBox1.Enabled = True
''    MaskEdBox1.SetFocus
''End If
''End Sub
''
''Private Sub opt_staff_Click()
''If opt_staff.Value = True Then
''
''    lbl_caseno.Caption = "  CC No."
''    MaskEdBox1.Mask = "########"
''    MaskEdBox1.Enabled = True
''    MaskEdBox1.SetFocus
''End If
''End Sub
''
''Private Sub txt_observed_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''    If SELDMA101FLAG = False Then
''        SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
''        Set SELDMA101RS = SELDMA101QR.OpenResultset(1, 3)
''        SELDMA101FLAG = True
''    Else
''        SELDMA101QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
''        SELDMA101RS.MoreResults
''        SELDMA101RS.Requery
''    End If
''    If SELDMA101RS.RowCount > 0 Then
''        If SELDTEMPLETFLAG = False Then
''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''            SELDTEMPLETFLAG = True
''        Else
''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''            SELDTEMPLETRS.MoreResults
''            SELDTEMPLETRS.Requery
''        End If
''
''        If SELDTEMPLETRS.RowCount > 0 Then
''            If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
''                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
''                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
''                If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC"))) Then
''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
''                Else
''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
''                End If
''                For i = 1 To 1000
''                Beep
''                Next i
''            Else
''                 If Val(Trim(txt_observed.Text)) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                    MSFlexGrid1.CellBackColor = vbWhite
''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
''                    If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
''                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
''                    Else
''                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
''                    End If
''                 Else
''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
''                 End If
''                 If Val(Trim(txt_observed.Text)) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                    MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
''                        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
''                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
''                        Else
''                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
''                        End If
''
''                 End If
''                 If Trim(txt_observed.Text) = "" Then
''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
''                    MSFlexGrid1.CellBackColor = RGB(255, 255, 255)
''
''                 End If
''            End If
''        End If
''    End If
''
''    i = MSFlexGrid1.Rows
''    If i = MSFlexGrid1.Row Then
''        cmd_save.SetFocus
''        Exit Sub
''    End If
''    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
''    End If
''    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''            txt_observed.Text = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
''
''            txt_observed.Visible = True
''            txt_observed.SetFocus
''
''            txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''
''    End If
''End If
''End Sub
''
''Private Sub txt_observed_LostFocus()
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
'''        If SELDTEMPLETFLAG = False Then
'''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
'''            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'''            SELDTEMPLETFLAG = True
'''        Else
'''            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
'''            SELDTEMPLETRS.MoreResults
'''            SELDTEMPLETRS.Requery
'''        End If
'''
'''        If SELDTEMPLETRS.RowCount > 0 Then
'''            If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
'''                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
'''                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'''                If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC"))) Then
'''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
'''                Else
'''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'''                End If
'''                For i = 1 To 1000
'''                Beep
'''                Next i
'''            Else
'''                 If Val(Trim(txt_observed.Text)) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'''                    MSFlexGrid1.CellBackColor = vbWhite
'''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'''                    If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
'''                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
'''                    Else
'''                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'''                    End If
'''                 Else
'''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
'''                 End If
'''                 If Val(Trim(txt_observed.Text)) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'''                    MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
'''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'''                        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) > Val(Trim(txt_observed.Text)) Then
'''                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txt_observed.Text
'''                        Else
'''                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'''                        End If
'''
'''                 End If
'''                 If Trim(txt_observed.Text) = "" Then
'''                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'''
'''                 End If
'''            End If
'''        End If
'''    End If
'''
'''    i = MSFlexGrid1.Rows - 1
'''    If i = MSFlexGrid1.Row Then
'''        cmd_save.SetFocus
'''        Exit Sub
'''    End If
'''    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
'''            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
'''    End If
'''    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'''            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'''            txt_observed.Text = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'''
'''            txt_observed.Visible = True
'''            txt_observed.SetFocus
'''
'''            txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'''
'''    End If
'''End If
''End Sub
''
''Public Sub assign()
'On Error GoTo msg:
''rdb.BeginTrans
''    If flagsave = 1 Then
''         For i = 1 To MSFlexGrid1.Rows - 1
''
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
''                INSDMPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
''                INSDMPROCEDURQR(1) = Trim(SELDMA101RS("SER_CODE"))
''                'INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                If SELDTEMPLETFLAG = False Then
''                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                    SELDTEMPLETFLAG = True
''                Else
''                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''                    SELDTEMPLETRS.MoreResults
''                    SELDTEMPLETRS.Requery
''                End If
''                If SELDTEMPLETRS.RowCount > 0 Then
''                    If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''
''                        Else
''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                        End If
''                        INSDMPROCEDURQR(3) = "H"
''
''                    Else
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                            INSDMPROCEDURQR(3) = "N"
''
''                        End If
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
''                                INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''                            Else
''                                INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                             End If
''                                INSDMPROCEDURQR(3) = "L"
''
''                        End If
''                        If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
''                            INSDMPROCEDURQR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                            INSDMPROCEDURQR(3) = ""
''
''                        End If
''                    End If
''                End If
''                INSDMPROCEDURQR(4) = Trim(SELDTEMPLETRS("uom"))
''                INSDMPROCEDURQR(5) = Trim(MSFlexGrid1.TextMatrix(i, 3))
''                INSDMPROCEDURQR(6) = Trim(MSFlexGrid1.TextMatrix(i, 4))
''                INSDMPROCEDURQR.Execute
''            End If
''
''
''         Next i
''
''            updrrequisitn_qr(0) = cmb_enterby.Text
''            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''            updrrequisitn_qr(2) = cmb_finalby.Text
''            updrrequisitn_qr(3) = "T"
''            updrrequisitn_qr(4) = cmb_finalby.Text
''            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
''            updrrequisitn_qr.Execute
''
''            INSDMBCCOMMQR(0) = Trim(req_chglen(cmb_reqno.Text))
''            If Len(Trim(txt_comments.Text)) > 0 Then
''                INSDMBCCOMMQR(1) = Null
''                INSDMBCCOMMQR(1) = txt_comments.Text
''            Else
''                INSDMBCCOMMQR(1) = Null
''
''            End If
''            INSDMBCCOMMQR.Execute
''
''
''
''    End If
''    If flagsave = 2 Then
''            For i = 1 To MSFlexGrid1.Rows - 1
''
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
'''
''                'UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''
''                If SELDTEMPLETFLAG = False Then
''                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                    SELDTEMPLETFLAG = True
''                Else
''                    SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
''                    SELDTEMPLETRS.MoreResults
''                    SELDTEMPLETRS.Requery
''                End If
''                If SELDTEMPLETRS.RowCount > 0 Then
''                    If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
''
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("highpanic"))) Then
''                            UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''
''                        Else
''                            UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                        End If
''
''                        UPDDMPROCEDURQR(1) = "H"
''                     Else
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                            UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                            UPDDMPROCEDURQR(1) = "N"
''                        End If
''                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
''                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("lowpanic"))) Then
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2)) & "*"
''                            Else
''                                UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                            End If
''                            UPDDMPROCEDURQR(1) = "L"
''                        End If
''                        If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
''                            UPDDMPROCEDURQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
''                            UPDDMPROCEDURQR(1) = ""
''                        End If
''                    End If
''                End If
''                 UPDDMPROCEDURQR(2) = Trim(SELDTEMPLETRS("UOM"))
''                 UPDDMPROCEDURQR(3) = Trim(MSFlexGrid1.TextMatrix(i, 3))
''                 UPDDMPROCEDURQR(4) = Trim(MSFlexGrid1.TextMatrix(i, 4))
''                 UPDDMPROCEDURQR(5) = req_chglen(cmb_reqno.Text)
''                 UPDDMPROCEDURQR(6) = Trim(SELDMA101RS("SER_CODE"))
''                UPDDMPROCEDURQR.Execute
''            End If
''
''
''         Next i
''
''            updrrequisitn_qr(0) = cmb_enterby.Text
''            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''            updrrequisitn_qr(2) = cmb_finalby.Text
''            updrrequisitn_qr(3) = "T"
''            updrrequisitn_qr(4) = cmb_finalby.Text
''            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
''            updrrequisitn_qr.Execute
''
''
''            If Len(Trim(txt_comments.Text)) > 0 Then
''                UPDDMBCCOMMQR(0) = Null
''                UPDDMBCCOMMQR(0) = txt_comments.Text
''            Else
''                UPDDMBCCOMMQR(0) = Null
''
''            End If
''            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
''            UPDDMBCCOMMQR.Execute
''
''
''    End If
''rdb.CommitTrans
''MsgBox "Record is Saved"
''btnenabl ("INI")
''txt_observed.Visible = False
''Exit Sub
''msg:
''    rdb.RollbackTrans
''MsgBox Err.Description
''End Sub
''
''Public Sub STAFF()
'' If NEWFLAG = True Then
''            NEWFLAG = False
''            If SELDREQSTAFFD1FLAG = False Then
''                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
''                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
''                Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
''            Else
''                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
''                SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
''                SELDREQSTAFFDRS.MoreResults
''                SELDREQSTAFFDRS.Requery
''
''            End If
''            If SELDREQSTAFFDRS.RowCount > 0 Then
''                txt_sex.Text = SELDREQSTAFFDRS("sex")
''                txt_name.Text = SELDREQSTAFFDRS("PNAME")
''                txt_age = SELDREQSTAFFDRS("AGE")
''                txt_pridoc.Text = ""
''                txt_category.Text = SELDREQSTAFFDRS("RELATION")
''            End If
''            If REFDRFLAG1 = False Then
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
''                REFDRFLAG1 = True
''            Else
''                REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                REFDR_RS.MoreResults
''                REFDR_RS.Requery
''            End If
''            If REFDR_RS.RowCount > 0 Then
''                txt_refdoc.Text = REFDR_RS("REFBYDOC")
''            End If
''
''            If SELMA104FLAG1 = False Then
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
''                SELMA104FLAG1 = True
''            Else
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                selma104_rs.MoreResults
''                selma104_rs.Requery
''            End If
''
''
''            If selma104_rs.RowCount > 0 Then
''                txt_ward.Visible = True
''                txt_bedno.Visible = True
''                lbl_ward.Visible = True
''                lbl_bed.Visible = True
''                txt_ward.Text = selma104_rs("WARD_NAME")
''                txt_bedno.Text = selma104_rs("bed_no")
''                txt_status.Text = "IP"
''            Else
''                txt_status.Text = "ST"
''                txt_ward.Visible = False
''                txt_bedno.Visible = False
''                lbl_ward.Visible = False
''                lbl_bed.Visible = False
''            End If
''
''
''            If SELPATHDREQUISITNFALG1 = False Then
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
''                SELPATHDREQUISITNFALG1 = True
''            Else
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                SELPATHDREQUISITN_RS.MoreResults
''                SELPATHDREQUISITN_RS.Requery
''            End If
''            If SELPATHDREQUISITN_RS.RowCount > 0 Then
''                txt_pathno.Text = SELPATHDREQUISITN_RS("PATHNO")
''
''            End If
''
''            If SELTA1012FLAG1 = False Then
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
''                SELTA1012FLAG1 = True
''            Else
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                selta1012_rs.MoreResults
''                selta1012_rs.Requery
''            End If
''            If selta1012_rs.RowCount > 0 Then
''                'txt_procedure.Text = selta1012_rs("prvdiagdes")
''            End If
''            '******
''            If SELDMULPROCFLAG = False Then
''                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
''                Set SELDMULPROCRS = SELDMULPROCQR.OpenResultset(1, 3)
''                SELDMULPROCFLAG = True
''           Else
''                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
''                SELDMULPROCRS.MoreResults
''                SELDMULPROCRS.Requery
''           End If
''           If SELDMULPROCRS.RowCount > 0 Then
''
''                MSFlexGrid1.Rows = 1
''                i = 1
''
''                While Not SELDMULPROCRS.EOF
''                   If SELDMULPROCRS("pro_code") = "F849" Or SELDMULPROCRS("PRO_CODE") = "F854" Or SELDMULPROCRS("PRO_CODE") = "F860" Or SELDMULPROCRS("PRO_CODE") = "F865" Or SELDMULPROCRS("PRO_CODE") = "F875" Or SELDMULPROCRS("PRO_CODE") = "F880" Or SELDMULPROCRS("PRO_CODE") = "F802" Then
''                       If SELDBCPROSUBFLAG = False Then
''                            SELDBCPROSUBQR(0) = SELDMULPROCRS("pro_code")
''                            Set SELDBCPROSUBRS = SELDBCPROSUBQR.OpenResultset(1, 3)
''                            SELDBCPROSUBFLAG = True
''                        Else
''                            SELDBCPROSUBQR(0) = SELDMULPROCRS("pro_code")
''                            SELDBCPROSUBRS.MoreResults
''                            SELDBCPROSUBRS.Requery
''                        End If
''                        If SELDBCPROSUBRS.RowCount > 0 Then
''                            While Not SELDBCPROSUBRS.EOF
''                                If SELMA101FLAG = False Then
''                                    selMA101QR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
''                                    SELMA101FLAG = True
''                                Else
''                                    selMA101QR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    SELMA101RS.MoreResults
''                                    SELMA101RS.Requery
''                                End If
''                                If SELDTEMPLETFLAG = False Then
''                                    SELDTEMPLETQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                                    SELDTEMPLETFLAG = True
''                                Else
''                                    SELDTEMPLETQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                    SELDTEMPLETRS.MoreResults
''                                    SELDTEMPLETRS.Requery
''                                End If
''
''                                '******
''                                If SELDBCPROSUBRS("PROCODESUB") = "F869" Then
''                                        ALKAGE
''                                    If SELDBCALKPHSFLAG = False Then
''                                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                        SELDBCALKPHSQR(1) = RNAGE
''                                        SELDBCALKPHSQR(2) = SS
''                                        Set SELBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
''                                        SELDBCALKPHSFLAG = True
''                                    Else
''                                        SELDBCALKPHSQR(0) = SELDBCPROSUBRS("PROCODESUB")
''                                        SELDBCALKPHSQR(1) = RNAGE
''                                        SELDBCALKPHSQR(2) = SS
''                                        SELBCALKPHSRS.MoreResults
''                                        SELBCALKPHSRS.Requery
''                                    End If
''                                    If SELDTEMPLETRS.RowCount > 0 Then
''                                    If SELMA101RS.RowCount > 0 Then
''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELBCALKPHSRS("RANGE"))
''                                        Else
''                                            SELDTEMPLETRS.MoveNext
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELBCALKPHSRS("RANGE"))
''                                        End If
''                                        i = i + 1
''                                    End If
''                                    End If
''                               Else
''                                    If SELDTEMPLETRS.RowCount > 0 Then
''                                    If SELMA101RS.RowCount > 0 Then
''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        Else
''                                            SELDTEMPLETRS.MoveNext
''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        End If
''                                        i = i + 1
''                                    End If
''                                    End If
''                                End If
'''                                '******
'''                                If SELDTEMPLETRS.RowCount > 0 Then
'''                                    If SELMA101RS.RowCount > 0 Then
'''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                        Else
'''                                            SELDTEMPLETRS.MoveNext
'''                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                        End If
'''                                        i = i + 1
'''                                    End If
'''                                End If
''
''                                SELDBCPROSUBRS.MoveNext
''                            Wend
''                        End If
''                   Else
''                            If SELMA101FLAG = False Then
''                                selMA101QR(0) = SELDMULPROCRS("PRO_CODE")
''                                Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
''                                SELMA101FLAG = True
''                            Else
''                                selMA101QR(0) = SELDMULPROCRS("PRO_CODE")
''                                SELMA101RS.MoreResults
''                                SELMA101RS.Requery
''                            End If
''                            If SELDTEMPLETFLAG = False Then
''                                SELDTEMPLETQR(0) = SELDMULPROCRS("PRO_CODE")
''                                Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                                SELDTEMPLETFLAG = True
''                            Else
''                                SELDTEMPLETQR(0) = SELDMULPROCRS("PRO_CODE")
''                                SELDTEMPLETRS.MoreResults
''                                SELDTEMPLETRS.Requery
''                            End If
''                            '*******
''                                If SELDMULPROCRS("PRO_CODE") = "F869" Then
''                                     ALKAGE
''                                    If SELDBCALKPHSFLAG = False Then
''                                        SELDBCALKPHSQR(0) = SELDMULPROCRS("PRO_CODE")
''                                        SELDBCALKPHSQR(1) = RNAGE
''                                        SELDBCALKPHSQR(2) = SS
''                                        Set SELBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
''                                        SELDBCALKPHSFLAG = True
''                                    Else
''                                        SELDBCALKPHSQR(0) = SELDMULPROCRS("PRO_CODE")
''                                        SELDBCALKPHSQR(1) = Trim(txt_age.Text)
''                                        SELDBCALKPHSQR(2) = RNAGE
''                                        SELDBCALKPHSRS.MoreResults
''                                        SELDBCALKPHSRS.Requery
''                                    End If
''                                    If SELDTEMPLETRS.RowCount > 0 Then
''                                        If SELMA101RS.RowCount > 0 Then
''                                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
''                                            Else
''                                                SELDTEMPLETRS.MoveNext
''                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDBCALKPHSRS("RANGE"))
''                                            End If
''                                            i = i + 1
''                                        End If
''                                    End If
''                                Else
''                                        If SELDTEMPLETRS.RowCount > 0 Then
''                                            If SELMA101RS.RowCount > 0 Then
''                                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                                Else
''                                                    SELDTEMPLETRS.MoveNext
''                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                                End If
''                                                i = i + 1
''                                            End If
''                                        End If
''                                End If
''                            '*******
'''                        If SELDTEMPLETRS.RowCount > 0 Then
'''                            If SELMA101RS.RowCount > 0 Then
'''                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
'''                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                Else
'''                                    SELDTEMPLETRS.MoveNext
'''                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & "" & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'''                                End If
'''                                i = i + 1
'''                            End If
'''                        End If
''                    End If
''                    SELDMULPROCRS.MoveNext
''
''                Wend
''                MSFlexGrid1.Row = 1
''                MSFlexGrid1.Col = 2
''                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
''                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''                txt_observed.Visible = True
''           End If
''
''        '    seldrequisitn_qr(0) = MaskEdBox1.FormattedText
''        '    seldrequisitn_rs.MoreResults
''        '    seldrequisitn_rs.Requery
''        '    If seldrequisitn_rs.RowCount > 0 Then
''        '     prpathno = ""
''        '     While Not seldrequisitn_rs.EOF
''        '     prpathno = prpathno & " " & seldrequisitn_rs("pathno")
''        '     CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
''        '     seldrequisitn_rs.MoveNext
''        '     Wend
''        '     txt_prepathreqno.Text = prpathno
''        '    End If
''        '    txt_addendum.Enabled = False
''        '    cmb_addendumby.Enabled = False
''
''
''
''        If seldissecentrflag1 = False Then
''            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
''            seldissecentrflag1 = True
''        Else
''            seldissecentr1_rs.MoreResults
''            seldissecentr1_rs.Requery
''        End If
''        If seldissecentr1_rs.RowCount > 0 Then
''            cmb_enterby.clear
''            While Not seldissecentr1_rs.EOF
''                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
''                seldissecentr1_rs.MoveNext
''            Wend
''        End If
''
''        If seldissecfinflag = False Then
''            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
''            seldissecfinflag1 = True
''        Else
''            seldissecfin1_rs.MoreResults
''            seldissecfin1_rs.Requery
''        End If
''        If seldissecfin1_rs.RowCount > 0 Then
''            cmb_finalby.clear
''            While Not seldissecfin1_rs.EOF
''                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
''                seldissecfin1_rs.MoveNext
''            Wend
''        End If
''
''            '******
''
''
''        End If
''        If VIEWFLAG = True Then
''
''
''
''            VIEWFLAG = False
''            If SELDREQSTAFFD1FLAG = False Then
''                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
''                SELDREQSTAFFDQR(1) = MaskEdBox1.FormattedText
''                Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
''            Else
''                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
''                SELDREQSTAFFDQR(1) = MaskEdBox1.FormattedText
''                SELDREQSTAFFDRS.MoreResults
''                SELDREQSTAFFDRS.Requery
''
''            End If
''            If SELDREQSTAFFDRS.RowCount > 0 Then
''                txt_sex.Text = SELDREQSTAFFDRS("sex")
''                txt_name.Text = SELDREQSTAFFDRS("PNAME")
''                txt_age = DateDiff("YYYY", SELDREQSTAFFDRS("BIRTH_DATE"), CStr(DATE11))
''                txt_pridoc.Text = ""
''                txt_category.Text = SELDREQSTAFFDRS("RELATION")
''            End If
''
''            If REFDRFLAG1 = True Then
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
''                REFDRFLAG1 = True
''            Else
''                REFDR_QR(0) = MaskEdBox1.FormattedText
''                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
''                REFDR_RS.MoreResults
''                REFDR_RS.Requery
''            End If
''
''            If REFDR_RS.RowCount > 0 Then
''                txt_refdoc.Text = REFDR_RS("REFBYDOC")
''                If REFDR_RS("COMMITFLAG") = "Y" Then
''                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
''                    ADDENDUMFALG = True
''                End If
''                cmb_enterby.Text = REFDR_RS("ENTERBY")
''                cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
''                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
''                    cmb_finalized.Text = ""
''                Else
''                    cmb_finalized.Text = REFDR_RS("FREPORTBY")
''                End If
''            End If
''
''
''
''
''            If SELMA104FLAG1 = False Then
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
''                SELMA104FLAG1 = True
''            Else
''                selma104_qr(0) = MaskEdBox1.FormattedText
''                selma104_rs.MoreResults
''                selma104_rs.Requery
''            End If
''
''
''            If selma104_rs.RowCount > 0 Then
''                txt_ward.Visible = False
''                txt_bedno.Visible = False
''                lbl_ward.Visible = False
''                lbl_bed.Visible = False
''                txt_ward.Text = selma104_rs("WARD_NAME")
''                txt_bedno.Text = selma104_rs("bed_no")
''                txt_status.Text = "IP"
''            Else
''                txt_ward.Visible = False
''                txt_bedno.Visible = False
''                lbl_ward.Visible = False
''                lbl_bed.Visible = False
''                txt_status.Text = "OP"
''            End If
''
''         If SELPATHDREQUISITNFALG1 = False Then
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
''                SELPATHDREQUISITNFALG1 = True
''            Else
''                SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''                SELPATHDREQUISITN_RS.MoreResults
''                SELPATHDREQUISITN_RS.Requery
''            End If
''            If SELPATHDREQUISITN_RS.RowCount > 0 Then
''                txt_ihmeterial.Text = SELPATHDREQUISITN_RS("NMETERIALS")
''                txt_refdoc.Text = SELPATHDREQUISITN_RS("REFDOC")
''            End If
''
''
''            If SELTA1012FLAG1 = False Then
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
''                SELTA1012FLAG1 = True
''            Else
''                selta1012_qr(0) = MaskEdBox1.FormattedText
''                selta1012_rs.MoreResults
''                selta1012_rs.Requery
''            End If
''            If selta1012_rs.RowCount > 0 Then
''                'txt_procedure.Text = selta1012_rs("prvdiagdes")
''            End If
''            '*******
''             '**** filling the flex grid
''            If SELDPROCEDURFLAG = False Then
''                SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
''                Set SELDPROCEDURRS = SELDPROCEDURQR.OpenResultset(1, 3)
''                SELDPROCEDURFLAG = True
''           Else
''                SELDPROCEDURQR(0) = req_chglen(cmb_reqno.Text)
''                SELDPROCEDURRS.MoreResults
''                SELDPROCEDURRS.Requery
''           End If
''           If SELDPROCEDURRS.RowCount > 0 Then
''
''''                MSFlexGrid1.Rows = 1
''''                i = 1
''''
''''                While Not SELDPROCEDURRS.EOF
''''                    If SELMA101FLAG = False Then
''''                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
''''                        SELMA101FLAG = True
''''                    Else
''''                        selMA101QR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        SELMA101RS.MoreResults
''''                        SELMA101RS.Requery
''''                    End If
''''                    If SELDTEMPLETFLAG = False Then
''''                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''''                        SELDTEMPLETFLAG = True
''''                    Else
''''                        SELDTEMPLETQR(0) = SELDPROCEDURRS("PRO_CODE")
''''                        SELDTEMPLETRS.MoreResults
''''                        SELDTEMPLETRS.Requery
''''                    End If
''''                    If SELDTEMPLETRS.RowCount > 0 Then
''''                        If SELMA101RS.RowCount > 0 Then
''''                            If SELDTEMPLETRS("SEX") = "M" Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''''                                 MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & SELDPROCEDURRS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                            Else
''''                                SELDTEMPLETRS.MoveNext
''''                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS("SER_DET")) & vbTab & SELDPROCEDURRS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''''                            End If
''''                        End If
''''
''''                            i = i + 1
''''
''''                    End If
''''                    SELDPROCEDURRS.MoveNext
''''
''''                Wend
'''''                MSFlexGrid1.Row = 1
'''''                MSFlexGrid1.Col = 2
'''''                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'''''                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''''                txt_observed.Visible = False
''                    'MSFlexGrid1.Rows = 1
''                    'MSFlexGrid1.Row = 1
''                    i = 1
''                    srno = 1
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
''
''                    If SELDTEMPLETRS.RowCount > 0 Then
''                            If SELMA101RS.RowCount > 0 Then
''                            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''                             MSFlexGrid1.Row = srno
''                             MSFlexGrid1.Col = 0
''                            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or Trim(SELDTEMPLETRS("SEX")) = "A" Then
''                                MSFlexGrid1.TextMatrix(i, 0) = i
''                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101RS("SER_DET"))
''                                If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
''                                    MSFlexGrid1.Row = srno
''                                     MSFlexGrid1.Col = 2
''                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
''                                Else
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
''
''                                      MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = RGB(100, 255, 150)
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
''
''                                       MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "" Then
''
''                                        MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                End If
''                                J = 0
''                                J = InStr(SELDPROCEDURRS("observalue"), "*")
''                                If J = 0 Then
''                                    MSFlexGrid1.TextMatrix(i, 2) = SELDPROCEDURRS("OBSERVALUE")
''                                Else
''                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, J - 1)
''                                End If
''                                MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                MSFlexGrid1.TextMatrix(i, 4) = SELDPROCEDURRS("pvalue")
''                            Else
''                                SELDTEMPLETRS.MoveNext
''                                MSFlexGrid1.TextMatrix(i, 0) = i
''                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101RS("SER_DET"))
''                                If Trim(SELDPROCEDURRS("obser_type")) = "H" Then
''                                    MSFlexGrid1.Row = srno
''                                     MSFlexGrid1.Col = 2
''                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
''                                Else
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "L" Then
''
''                                      MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = RGB(150, 255, 100)
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "N" Then
''
''                                       MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                    If Trim(SELDPROCEDURRS("obser_type")) = "" Then
''
''                                        MSFlexGrid1.Col = 2
''                                      MSFlexGrid1.CellBackColor = vbWhite
''                                    End If
''                                End If
''
''                                J = 0
''                                J = InStr(SELDPROCEDURRS("observalue"), "*")
''                                If J = 0 Then
''                                    MSFlexGrid1.TextMatrix(i, 2) = Trim(SELDPROCEDURRS("OBSERVALUE"))
''                                Else
''                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(SELDPROCEDURRS("OBSERVALUE"), 1, J - 1)
''                                End If
''                                MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                MSFlexGrid1.TextMatrix(i, 4) = SELDPROCEDURRS("pvalue")
''                            End If
''                        End If
''                           srno = srno + 1
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
''
''           End If
''
''           If SELDMBCCOMMFLAG = False Then
''                SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
''                Set SELDMBCCOMMRS = SELDMBCCOMMQR.OpenResultset(1, 3)
''                SELDMBCCOMMFLAG = True
''            Else
''                SELDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
''                SELDMBCCOMMRS.MoreResults
''                SELDMBCCOMMRS.Requery
''            End If
''            If SELDMBCCOMMRS.RowCount > 0 Then
''                If SELDMBCCOMMRS.EOF = False Then
''                    txt_comments.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
''                End If
''            End If
''            '**** filling the flex grid
''
''
''
''            '*******
''
''        End If
''End Sub
''
''Public Sub ALKAGE()
''If txt_age.Text > 17 Then
''RNAGE = 17
''SS = ""
''End If
'''If txt_age.Text <= 0.6 Then
'''RNAGE = 0.6
'''SS = ""
'''End If
''If txt_age.Text <= 6 And txt_age.Text > 4 Then
''RNAGE = 6
''SS = ""
''End If
''If txt_age.Text <= 1 And txt_age.Text > 0.7 Then
''RNAGE = 1
''SS = ""
''End If
''If txt_age.Text <= 3 And txt_age.Text > 1 Then
''RNAGE = 3
''SS = ""
''End If
''If (txt_age.Text <= 12 And txt_age.Text >= 7) And Trim(txt_sex.Text) = "M" Then
''RNAGE = 12
''SS = "M"
''End If
''If (txt_age.Text <= 12 And txt_age.Text >= 7) And Trim(txt_sex.Text) = "F" Then
''RNAGE = 12
''SS = "F"
''End If
''If (txt_age.Text <= 17 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "M" Then
''RNAGE = 17
''SS = "M"
''End If
''If (txt_age.Text <= 17 And txt_age.Text >= 13) And Trim(txt_sex.Text) = "F" Then
''RNAGE = 17
''SS = "F"
''End If
''End Sub
