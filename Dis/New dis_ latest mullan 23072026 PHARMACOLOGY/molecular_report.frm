VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form molecular_report 
   Caption         =   "Molecular Report"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12555
   LinkTopic       =   "Form2"
   ScaleHeight     =   8940
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   480
   End
   Begin VB.Frame common 
      Height          =   2895
      Left            =   0
      TabIndex        =   32
      Top             =   960
      Width           =   11775
      Begin VB.CommandButton CMD_EMR 
         BackColor       =   &H00FFFF00&
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
         Height          =   315
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2520
         Width           =   1215
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
         ItemData        =   "molecular_report.frx":0000
         Left            =   6240
         List            =   "molecular_report.frx":0002
         TabIndex        =   37
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txt_clndiag 
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
         Left            =   8355
         TabIndex        =   52
         Top             =   2190
         Width           =   3195
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
         TabIndex        =   51
         Top             =   2160
         Width           =   3855
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
         TabIndex        =   50
         Top             =   1800
         Width           =   1095
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
         TabIndex        =   49
         Top             =   1800
         Width           =   975
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
         TabIndex        =   48
         Top             =   1800
         Width           =   1335
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
         TabIndex        =   47
         Top             =   1800
         Width           =   1575
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
         TabIndex        =   46
         Top             =   1440
         Width           =   1095
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
         TabIndex        =   45
         Top             =   1440
         Width           =   975
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
         TabIndex        =   44
         Top             =   1440
         Width           =   735
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
         TabIndex        =   43
         Top             =   1440
         Width           =   855
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
         TabIndex        =   42
         Top             =   1080
         Width           =   9135
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
         TabIndex        =   41
         Top             =   600
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
         TabIndex        =   39
         Top             =   600
         Width           =   2775
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
         TabIndex        =   40
         Top             =   240
         Width           =   1335
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
         TabIndex        =   36
         Top             =   240
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1335
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
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Value           =   -1  'True
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
            Left            =   240
            TabIndex        =   34
            Top             =   600
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   38
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.Label lbl_clndis 
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
         Left            =   6435
         TabIndex        =   69
         Top             =   2190
         Width           =   1830
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
         TabIndex        =   68
         Top             =   2160
         Width           =   2130
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
         TabIndex        =   67
         Top             =   1800
         Width           =   1215
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
         Top             =   315
         Width           =   720
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
         TabIndex        =   65
         Top             =   1800
         Width           =   705
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
         TabIndex        =   64
         Top             =   1800
         Width           =   495
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
         TabIndex        =   63
         Top             =   1800
         Width           =   1995
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
         TabIndex        =   62
         Top             =   1440
         Width           =   555
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
         TabIndex        =   61
         Top             =   1440
         Width           =   765
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
         TabIndex        =   60
         Top             =   1440
         Width           =   345
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
         TabIndex        =   59
         Top             =   1440
         Width           =   330
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
         Left            =   1800
         TabIndex        =   58
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label lbl_cytologyno 
         AutoSize        =   -1  'True
         Caption         =   "M Path No."
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
         Left            =   9135
         TabIndex        =   57
         Top             =   630
         Width           =   975
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
         Left            =   4800
         TabIndex        =   56
         Top             =   720
         Width           =   1440
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
         Left            =   1560
         TabIndex        =   55
         Top             =   720
         Width           =   795
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
         TabIndex        =   54
         Top             =   240
         Width           =   975
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
         TabIndex        =   53
         Top             =   360
         Width           =   1830
      End
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10440
      TabIndex        =   31
      Top             =   600
      Width           =   1095
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10440
      TabIndex        =   30
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame COMMAND 
      Height          =   795
      Left            =   0
      TabIndex        =   21
      Top             =   7800
      Width           =   11775
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
         Left            =   7470
         MouseIcon       =   "molecular_report.frx":0004
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   240
         Width           =   1215
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
         MouseIcon       =   "molecular_report.frx":030E
         MousePointer    =   99  'Custom
         TabIndex        =   28
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
         Left            =   9012
         MouseIcon       =   "molecular_report.frx":0618
         MousePointer    =   99  'Custom
         TabIndex        =   27
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
         Left            =   6048
         MouseIcon       =   "molecular_report.frx":0922
         MousePointer    =   99  'Custom
         TabIndex        =   26
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
         Left            =   4626
         MouseIcon       =   "molecular_report.frx":0C2C
         MousePointer    =   99  'Custom
         TabIndex        =   25
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
         Left            =   3204
         MouseIcon       =   "molecular_report.frx":0F36
         MousePointer    =   99  'Custom
         TabIndex        =   24
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
         Left            =   1782
         MouseIcon       =   "molecular_report.frx":1240
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
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
         Left            =   360
         MouseIcon       =   "molecular_report.frx":154A
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame REPORT 
      Height          =   4125
      Left            =   15
      TabIndex        =   0
      Top             =   3660
      Width           =   11775
      Begin VB.ComboBox cmb_result 
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
         ItemData        =   "molecular_report.frx":1854
         Left            =   2400
         List            =   "molecular_report.frx":185E
         TabIndex        =   80
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nature of Material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   2400
         TabIndex        =   77
         Top             =   1035
         Width           =   6825
         Begin VB.TextBox txt_archive 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2235
            TabIndex        =   5
            Top             =   615
            Width           =   4485
         End
         Begin VB.TextBox txt_fresh 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2235
            TabIndex        =   4
            Top             =   255
            Width           =   4485
         End
         Begin VB.Label lbl_Archival 
            AutoSize        =   -1  'True
            Caption         =   "2.        Archival"
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
            Left            =   750
            TabIndex        =   79
            Top             =   645
            Width           =   1350
         End
         Begin VB.Label Label4 
            Caption         =   "1.        Fresh"
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
            Height          =   255
            Left            =   750
            TabIndex        =   78
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.ComboBox cmb_laterality 
         Height          =   315
         Left            =   9195
         TabIndex        =   3
         Top             =   660
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.ComboBox cmb_site 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "molecular_report.frx":1876
         Left            =   2400
         List            =   "molecular_report.frx":1878
         TabIndex        =   2
         Top             =   675
         Width           =   3495
      End
      Begin VB.TextBox txt_impression 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2415
         TabIndex        =   6
         Top             =   2505
         Width           =   9135
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8655
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   3285
         Width           =   2895
      End
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1530
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3690
         Width           =   3255
      End
      Begin VB.ComboBox cmb_enteryby 
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
         Left            =   1545
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   3255
         Width           =   3240
      End
      Begin VB.ComboBox cmb_addendumby 
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
         Left            =   8640
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   3705
         Width           =   2895
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2415
         TabIndex        =   7
         Top             =   2880
         Width           =   9135
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   315
         Width           =   9135
      End
      Begin VB.Label Label1 
         Caption         =   "Laterality"
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
         Left            =   8265
         TabIndex        =   76
         Top             =   720
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label3 
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
         Height          =   315
         Left            =   1380
         TabIndex        =   75
         Top             =   2550
         Width           =   915
      End
      Begin VB.Label lbl_impression 
         AutoSize        =   -1  'True
         Caption         =   "Results"
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
         Left            =   1650
         TabIndex        =   20
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label lbl_addendum 
         AutoSize        =   -1  'True
         Caption         =   "Addendum By"
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
         Index           =   0
         Left            =   7320
         TabIndex        =   19
         Top             =   3780
         Width           =   1170
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
         Left            =   6915
         TabIndex        =   18
         Top             =   3330
         Width           =   1590
      End
      Begin VB.Label lbl_Reportby 
         AutoSize        =   -1  'True
         Caption         =   "Reported by"
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
         TabIndex        =   17
         Top             =   3735
         Width           =   1050
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
         Left            =   435
         TabIndex        =   16
         Top             =   3300
         Width           =   990
      End
      Begin VB.Label lbl_addendum1 
         AutoSize        =   -1  'True
         Caption         =   "Addendum "
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
         Left            =   1440
         TabIndex        =   15
         Top             =   2865
         Width           =   960
      End
      Begin VB.Label lbl_procedure 
         AutoSize        =   -1  'True
         Caption         =   "Name of Procedure"
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
         Left            =   540
         TabIndex        =   14
         Top             =   345
         Width           =   1695
      End
      Begin VB.Label lbl_sitespeciman 
         AutoSize        =   -1  'True
         Caption         =   "Site"
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
         Left            =   1830
         TabIndex        =   13
         Top             =   720
         Width           =   345
      End
      Begin VB.Label LBL_VPP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_cyt 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - MOLECULAR PATHOLOGY"
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
      Left            =   3255
      TabIndex        =   74
      Top             =   570
      Width           =   5220
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
      TabIndex        =   73
      Top             =   600
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
      Left            =   9840
      TabIndex        =   72
      Top             =   240
      Width           =   420
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
      TabIndex        =   71
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   675
      TabIndex        =   70
      Top             =   60
      Width           =   9015
   End
End
Attribute VB_Name = "molecular_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim str As String

'SAVE THE DATE AND TIME INTO RTIMELOG
Dim UPDRTIME_QR As rdoQuery
Dim UPD_RS As rdoResultset

Dim UPDRTIMECADD_QR As rdoQuery

Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEC_RS As rdoResultset

Dim UPDRTIMEP_QR As rdoQuery
Dim UPDRTIMEM_QR As rdoQuery

Dim FSFLAG As Boolean
Dim sel_deptqr As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SEL_DEPTFLAG As Boolean
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim SEL_MODALITYFLAG As Boolean
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim REQFLAG As Boolean
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim REPDRFLAG As Boolean
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim REFDRFLAG As Boolean
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset
Dim selma104flag As Boolean
Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset
Dim SELPATHDREQUISITNFLAG As Boolean
Dim selma302_qr As rdoQuery
Dim selma302_rs As rdoResultset
Dim SELMA302FLAG As Boolean
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELTA1012FLAG As Boolean
Dim WREPORT As Application
Dim DREPORT As Object
Dim WIMPRE As Application
Dim DIMPRE As Document

Dim d_grossdesc As Object
Dim W_grossdes As Application

Dim WMETERIAL As Application
Dim DMETERIAL As Document

Dim SELEMPLIB_QR As rdoQuery
Dim SELEMPLIB_RS As rdoResultset
Dim I As Integer
Dim insdimpresion_qr As rdoQuery
Dim insdspreport_qr As rdoQuery
Dim INSDIHREPORT_QR As rdoQuery
Dim upddfsreport_qr As rdoQuery
Dim upddihreport_qr As rdoQuery
Dim upddimpresion_qr As rdoQuery
Dim upddspreport_qr As rdoQuery
'Dim upddimpresion_qr As rdoQuery
Dim updgrossdesc_qr As rdoQuery
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim selempbasinfo1_qr As rdoQuery
Dim selempbasinfo1_rs As rdoResultset
Dim SELEMPBASINFO1FLAG As Boolean
Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim UPDRREQUISITN1FLAG As Boolean
Dim INSDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUMFLAG As Boolean
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG As Boolean
Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer
Dim ADDENDUMFALG As Boolean

Dim SELMOLEPROC_QR As rdoQuery
Dim SELMOLEPROC_RS As rdoResultset
Dim SELMOLEPROC_FLAG As Boolean

Dim SELMOLEREQ_QR As rdoQuery
Dim SELMOLEREQ_RS As rdoResultset
Dim SELMOLEREQ_FLAG As Boolean

Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
Dim SELMA101_FLAG As Boolean

Dim INSFREPORT_QR As rdoQuery
Dim UPDFREPORT_QR As rdoQuery
Dim SELRREQNOQR As rdoQuery
Dim SELRREQNORS As rdoResultset
Dim SELRREQNOFLAG As Boolean
Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset
Dim seldfreportflag As Boolean
Dim SELDFREPORT_IHQR As rdoQuery
Dim SELDFREPORT_IHRS As rdoResultset
Dim SELDFREPORT_IHFLAG As Boolean
Dim upddrequisitn_qr As rdoQuery

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFOFLAG As Boolean
Dim seldgradeqr As rdoQuery
Dim seldgraders As rdoResultset
Dim SELDGRADEFLAG As Boolean
Dim seldgrade_qr As rdoQuery
Dim seldgrade_rs As rdoResultset
Dim SELDGRADE_FLAG As Boolean
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean
Dim selihreportqr As rdoQuery
Dim selihreportrs As rdoResultset
Dim SELIHREPORTFLAG As Boolean

Dim seldissecentr_qr As rdoQuery
Dim seldissecentr_rs As rdoResultset
Dim seldissecentrflag As Boolean

Dim seldissecprov_qr As rdoQuery
Dim seldissecprov_rs As rdoResultset
Dim seldissecprovflag As Boolean

Dim seldissecfin_qr As rdoQuery
Dim seldissecfin_rs As rdoResultset
Dim seldissecfinflag As Boolean

Dim seldissecadd_qr As rdoQuery
Dim seldissecadd_rs As rdoResultset
Dim seldissecaddrflag As Boolean

Dim seldissecgross_qr As rdoQuery
Dim seldissecgross_rs As rdoResultset
Dim seldissecgrossflag As Boolean

Dim SELDREQSTAFFD1QR1 As rdoQuery
Dim SELDREQSTAFFD1RS1 As rdoResultset
Dim SELDREQSTAFFD1FLAG1 As Boolean

Dim viewsitedupqr As rdoQuery
Dim viewsiteduprs As rdoResultset
Dim VIEWSITEDUPRSFLAG As Boolean

Dim SELCLIN_QR As rdoQuery
Dim SELCLIN_RS As rdoResultset
Dim SELCLINFLAG As Boolean

Dim INSDISIMP_QR As rdoQuery
Dim UPDDISIMP_QR As rdoQuery

Dim UPDDREQ_QR As rdoQuery

''REPORT CLASSIFICATION

Dim UPDIMMHISCHE_QR As rdoQuery

Dim SELIMMHISCHE_QR As rdoQuery
Dim SELIMMHISCHE_RS As rdoResultset
Dim SELIMMHISCHE_FLAG As Boolean

''REPORT CLASSIFICATION

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

''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean
Dim VAR_REQNNO As String
''*******ICU**********

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
       If SEL_MODALITYFLAG = False Then
        sel_moadlityqr(0) = Trim(txt_deptcode.Text)
        Set sel_modalityrs = sel_moadlityqr.OpenResultset(2, 3)
        SEL_MODALITYFLAG = True
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
MsgBox ERROR(Err)
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "FZZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
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
'        If seldissecentrflag = False Then
'            Set seldissecentr_rs = seldissecentr_qr.OpenResultset(1, 3)
'            seldissecentrflag = True
'        Else
'            seldissecentr_rs.MoreResults
'            seldissecentr_rs.Requery
'        End If
'        If seldissecentr_rs.RowCount > 0 Then
'            cmb_enteryby.clear
'            While Not seldissecentr_rs.EOF
'                cmb_enteryby.AddItem Trim(seldissecentr_rs("CCNAME"))
'                seldissecentr_rs.MoveNext
'            Wend
'        End If
'        If seldissecprovflag = False Then
'            Set seldissecprov_rs = seldissecprov_qr.OpenResultset(1, 3)
'            seldissecprovflag = True
'        Else
'            seldissecprov_rs.MoreResults
'            seldissecprov_rs.Requery
'        End If
'        If seldissecprov_rs.RowCount > 0 Then
'            cmb_provby.clear
'            While Not seldissecprov_rs.EOF
'                cmb_provby.AddItem Trim(seldissecprov_rs("CCNAME"))
'                seldissecprov_rs.MoveNext
'            Wend
'        End If
'        If seldissecfinflag = False Then
'            Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
'            seldissecfinflag = True
'        Else
'            seldissecfin_rs.MoreResults
'            seldissecfin_rs.Requery
'        End If
'        If seldissecfin_rs.RowCount > 0 Then
'            cmb_finalby.clear
'            While Not seldissecfin_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin_rs("CCNAME"))
'                seldissecfin_rs.MoveNext
'            Wend
'        End If
'        If seldissecgrossflag = False Then
'            Set seldissecgross_rs = seldissecgross_qr.OpenResultset(1, 3)
'            seldissecgrossflag = True
'        Else
'            seldissecgross_rs.MoreResults
'            seldissecgross_rs.Requery
'        End If
'        If seldissecgross_rs.RowCount > 0 Then
'            While Not seldissecgross_rs.EOF
'                seldissecgross_rs.MoveNext
'            Wend
'        End If
Exit Sub
MSG:
'MsgBox Error(Err)
End Sub

Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "MOLECULAR PATHOLOGY$MP" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If

Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_rpt_class_Click()
    If Trim(cmb_rpt_class.Text) <> "" Then
       If Trim(cmb_rpt_class.Text) = "Neg" Then
       
       str = "SELECT distinct USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ?  AND DEPT_CODE = 'FZZ'"
       Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1", str)
            cmb_finalby.clear
            SELFINALBY_QR(0) = 4 ' MADE 6 FROM 5
            Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
            If Not SELFINALBY_RS.EOF Then
               Do While Not SELFINALBY_RS.EOF
                cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
                SELFINALBY_RS.MoveNext
               Loop
            End If
            SELFINALBY_QR.Close
        End If
        
        If Trim(cmb_rpt_class.Text) = "Pos" Then
        
        str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE >= ? AND DEPT_CODE = 'FZZ'"
        Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1", str)
             cmb_finalby.clear
            SELFINALBY_QR(0) = 4  'MADE 8 FROM 5
            Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
            If Not SELFINALBY_RS.EOF Then
                Do While Not SELFINALBY_RS.EOF
                    cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
                    SELFINALBY_RS.MoveNext
                Loop
            End If
            SELFINALBY_QR.Close
        End If
    End If
End Sub

Private Sub cmd_back_Click()
'btnenabl ("INI")
BTNENABL1 ("INI")
FLDENABL ("DIS")
TXT_CLEAR
DIS_MAIN.show
Me.Hide
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
    'cmd_fullreport.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
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

Private Sub cmd_cancel_Click()
'btnenabl ("INI")
BTNENABL1 ("INI")
field_clear
TXT_CLEAR
ADDENDUMFALG = False
FSFLAG = False
FLDENABL ("DIS")
End Sub

Private Sub cmd_commit_Click()
Dim I As Integer

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
 If CURUSRDET_RS("usrtype") <= 6 Then 'MADE 6 FROM 2
          MsgBox "Postive Reports cannot be commited by Cyto tecnhologist"
         Exit Sub
 End If
     

''********SECURITY***************************

I = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If I = vbYes Then
On Error GoTo MSG:

 rdb.BeginTrans
    updrrequisitn1_qr(0) = Left(cmb_finalby.Text, 30)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
End If
rdb.CommitTrans
'btnenabl ("INI")

    'rtimelog
    UPDRTIMEC_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute
BTNENABL1 ("INI")
MsgBox "Record is Committed Successfully."
cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_fullreport_Click()
On Error GoTo MSG:
txt_fullreport_GotFocus
Set D1 = w.ActiveDocument
If Len(txt_addendum) > 0 Then
D1.Content = D1.Content & "ADDENDUM:" & Chr(13) & Trim(txt_addendum)
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_fullreport_LostFocus()
On Error GoTo MSG:
D1.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
'D1.SaveAs "C:\REPORT.DOC"
D1.Close (wdDoNotSaveChanges)
w.quit
'Msk_caseno.Enabled = False
'cmd_cancel.SetFocus
MSG:
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
    strLink = "http://10.80.1.9/emrtmhact/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
End Sub

Private Sub cmd_modify_Click()

''********SECURITY***************************
CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
CURUSRDET_RS.MoreResults
CURUSRDET_RS.Requery
If CURUSRDET_RS.RowCount > 0 Then
    If LOGINTYP > CURUSRDET_RS("usrtype") Then
      ''   no action as he can commit
    ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
     '   MsgBox "Sorry! This Report needs is to Modified by a Higher Level than  that of Yours"
     '   cmd_modify.SetFocus
     '   Exit Sub
    'ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
     '   If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
            '  If loginusrid <> curusrdetrs("usr_id") Then
      '      MsgBox "Sorry! This Report is not to be Finalized by You"
       '     cmd_commit.SetFocus
        '    Exit Sub
        'End If
    End If
    If CURUSRDET_RS("MODIFY") <> "YES" Then
        MsgBox "Sorry! You are not authorised to Modify the Report of this Modality"
        cmd_commit.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
    'cmd_commit.SetFocus
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_reportby.Enabled = True
    Exit Sub
End If

''********SECURITY***************************
'btnenabl ("NEW")
BTNENABL1 ("NEW")
FLDENABL ("MODIFY")
flagsave = 2
If REFDRFLAG = False Then
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    REFDRFLAG = True
Else
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    REFDR_RS.MoreResults
    REFDR_RS.Requery
End If
If REFDR_RS.RowCount > 0 Then
    If Trim(REFDR_RS("reqnstat")) = "T" And Trim(REFDR_RS("COMMITFLAG")) = "F" Then
    ADDENDUMFALG = True
    End If
End If
If ADDENDUMFALG = True Then
    txt_addendum.Enabled = True
    cmb_addendumby.Enabled = True
'    If seldissecaddrflag = False Then
'            Set seldissecadd_rs = seldissecadd_qr.OpenResultset(1, 3)
'            seldissecaddrflag = True
'        Else
'            seldissecadd_rs.MoreResults
'            seldissecadd_rs.Requery
'        End If
'        If seldissecadd_rs.RowCount > 0 Then
'            cmb_addendumby.clear
'            While Not seldissecadd_rs.EOF
'                cmb_addendumby.AddItem Trim(seldissecadd_rs("CCNAME"))
'                seldissecadd_rs.MoveNext
'            Wend
'    End If
    FLDENABL ("clindiag")
    FLDENABL ("ADDENDUM")
    flagsave = 2
End If
End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG:
'btnenabl ("NEW")
BTNENABL1 ("NEW")
field_clear
TXT_CLEAR
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
cmb_enteryby.Text = cmb_enteryby.List(0)
cmb_enteryby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enteryby.List(0)) = cmb_finalby.List(L1) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
Next L1
cmb_reportby.Text = cmb_reportby.List(0)
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
 
flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
Opt_patient.Value = True
'MaskEdBox1.Enabled = True

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
On Error GoTo prtmsg:

DoEvents

'If Trim(cmb_rpt_class.Text) = "Pos" Then
    If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
 
VAR_REQNNO = req_chglen(cmb_reqno.Text)

If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = req_chglen(cmb_reqno.Text)
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
For I = 0 To 17
    CrystalReport1.Formulas(I) = ""
Next
 str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
If Opt_patient.Value = True Then
    If ADDENDUMFALG = False Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISMP.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(cmb_site.Text) = "" Then
            CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "STL = '" & "Site  " & cmb_site.Text & "'"
        End If
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL MOLECULAR PATHOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL MOLECULAR PATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    End If
End If
If opt_staff.Value = True Then
    If ADDENDUMFALG = False Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(cmb_site.Text) = "" Then
            CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "STL = '" & "Site  " & Trim(cmb_site.Text) & "'"
        End If
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL MOLECULAR PATHOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL MOLECULAR PATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    End If
End If
If Trim(cmb_reportby.Text) <> "" And Trim(cmb_finalby.Text) <> "" And UCase(Trim(cmb_reportby.Text)) <> UCase(Trim(cmb_finalby.Text)) Then
'If Trim(cmb_rpt_class.Text) = "Pos" Then
 If Trim(cmb_reportby.Text) <> Trim(cmb_finalby.Text) Then
    CrystalReport1.Formulas(11) = "PRFN= '" & Trim(cmb_reportby.Text) & "'"
    CrystalReport1.Formulas(12) = "ENT = '" & Trim(cmb_enteryby.Text) & "'"
    CrystalReport1.Formulas(13) = "PCAP= '" & "Registrar" & " '"
    CrystalReport1.Formulas(14) = "P = '" & Trim(cmb_reportby.Text) & "'"
    CrystalReport1.Formulas(16) = "FCAP = '" & "Pathologist" & "'"
    CrystalReport1.Formulas(17) = "F = '" & Trim(cmb_finalby.Text) & "'"
    'CrystalReport1.Formulas(12) = "F= '" & Trim(cmb_finalby.Text) & "'"
 Else
    CrystalReport1.Formulas(11) = "P= '" & Trim(cmb_reportby.Text) & "'"
    'CRYSTALREPORT1.Formulas(12) = "F= '" & Trim(cmb_finalby.Text) & "'"
    CrystalReport1.Formulas(13) = "PCAP= '" & "Registrar" & " '"
    'CRYSTALREPORT1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
 End If
 
End If
If Trim(cmb_reportby.Text) <> "" And Trim(cmb_finalby.Text) = "" Then
    CrystalReport1.Formulas(11) = "PONLY= '" & Trim(cmb_reportby.Text) & "'"
    'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
    CrystalReport1.Formulas(13) = "PFCAP= '" & "Registrar" & " '"
End If
If Trim(cmb_reportby.Text) <> "" And Trim(cmb_finalby.Text) <> "" And UCase(Trim(cmb_reportby.Text)) = UCase(Trim(cmb_finalby.Text)) Then
    CrystalReport1.Formulas(11) = "PONLY= '" & Trim(cmb_reportby.Text) & "'"
    CrystalReport1.Formulas(13) = "PFCAP= '" & "Registrar" & " '"
End If


CrystalReport1.WindowState = crptMaximized
'CRYSTALREPORT1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1

'rtime log
UPDRTIMEP_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIMEP_QR(1) = Format(txt_time.Text, "HH:MM:SS")
UPDRTIMEP_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIMEP_QR.Execute
Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("print")
BTNENABL1 ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_save_Click()
On Error GoTo MSG:
Dim m As Integer
If Trim(cmb_finalby.Text) = "" Then
    MsgBox "Finalized By Field is Empty !!", vbInformation, "Molecular Pathology"
    cmb_finalby.SetFocus
    Exit Sub
End If
If Trim(cmb_reportby.Text) = "" Then
    MsgBox "Provisional By Field is Empty !!", vbInformation, "Molecular Pathology"
    cmb_reportby.SetFocus
    Exit Sub
End If
If Trim(cmb_enteryby.Text) = "" Then
    MsgBox "Entered By Field is Empty !!", vbInformation, "Molecular Pathology"
    cmb_enteryby.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_reqno)) = 0 Then
    m = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
    cmb_reqno.SetFocus
    Exit Sub
End If

    If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
        

        
'       If Len(Trim(txt_impdiag.Text)) = 0 Then
'            MsgBox "Please Enter the Impression ", vbInformation
'            txt_impdiag.SetFocus
'            Exit Sub
'       End If
'       If Len(Trim(txt_typespeciman.Text)) > 1000 Then
'            MsgBox "Nature of Meterial Max. of 1000 Characters Only.", vbInformation
'            txt_typespeciman.SetFocus
'            Exit Sub
'       End If
'       If Len(Trim(txt_addendum.Text)) > 600 Then
'            MsgBox "Addendum Max. of 600 Characters Only.", vbInformation
'            txt_addendum.SetFocus
'            Exit Sub
'       End If
'rdb.BeginTrans
       assign
      'rtimelog
UPDRTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                
UPDRTIME_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIME_QR.Execute
       
'rdb.CommitTrans
       'MsgBox "Record is saved Successfully."
  FLDENABL ("DIS")
  'btnenabl ("SAVE")
  BTNENABL1 ("SAVE")
  Exit Sub
MSG:
'rdb.RollbackTrans
MsgBox ERROR(Err)
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
'btnenabl ("VIEW")
BTNENABL1 ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR
'&&&&&&&&&&&&&&&&&&&&77 not found hence declared
Dim SEL_DEPTRSFLAG As Boolean

If SEL_DEPTRSFLAG = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SEL_DEPTRSFLAG = True
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
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
'cmd_commit.Enabled = True
cmb_dept.SetFocus
cmd_print.Enabled = True
Opt_patient.Value = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
Dim VAR1
'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
VAR1 = MaskEdBox1.FormattedText
VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
AppActivate VAR
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
If CYREPORTFLAG = False Then
    CYREPORTFLAG = True
'btnenabl ("INI")

NEWFLAG = False
viewflag = False
MaskEdBox1.Enabled = False


''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR117", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR117", str)

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set SELMA101_QR = rdb.CreateQuery("MA101QR", str)
SELMA101_FLAG = False

''********ICU***************

'THIS IS RTIMLOG
 str = "update rislib.rtimelog set ADPREPDATE=?,ADPREPTIME=? where reqnno=?"
 Set UPDRTIMECADD_QR = rdb.CreateQuery("updatetimecadcy", str)
 
 str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=? WHERE REQNNO=?"
 Set UPDRTIME_QR = rdb.CreateQuery("QWEQWEQWEcy", str)

 str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEccy", str)
 
 
 str = "UPDATE RISLIB.RTIMELOG SET PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEP_QR = rdb.CreateQuery("QWEQWEQWEPcy", str)

 str = "UPDATE RISLIB.RTIMELOG SET MODREPDATE=?,MODREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEM_QR = rdb.CreateQuery("QWEQWEQWEMcy", str)

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enteryby.AddItem LOGINUSRNAME
    cmb_enteryby.Text = Trim(LOGINUSRNAME)
    
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STRAA", str)
If Trim(USRDEPT) = "ALL" Then
    REPWTRN_QR(0) = "FZZ"
Else
    REPWTRN_QR(0) = Trim(USRDEPT)
End If
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    cmb_addendumby.clear
    cmb_finalby.clear
    cmb_reportby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 5 Then   'MADE 5 TO 8
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        If REPWTRNRS("USRTYPE") >= 3 And REPWTRNRS("USRTYPE") <> 4 Then
            cmb_reportby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
                REPWTRNRS.MoveNext
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE > ? AND DEPT_CODE = 'FZZ'"
Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1", str)
SELFINALBY_QR(0) = 5
Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
If Not SELFINALBY_RS.EOF Then
    Do While Not SELFINALBY_RS.EOF
       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
       SELFINALBY_RS.MoveNext
    Loop
End If

  SELFINALBY_QR.Close
'*********SECURITY**********************************

''update drequisitn table by slides no.
str = "UPDATE DISLIB.DREQUISITN SET LABNO=? WHERE REQNO=?"
Set UPDDREQ_QR = rdb.CreateQuery("UPDDREQ_STR", str)
   
    'To display current date
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
    
'for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' OR DEPT_CODE='IZZ')"
Set sel_deptqr = rdb.CreateQuery("selC1_deptqrcg", str)
'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
        
    ' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqrcg", str)

    'for retriving reqnos
Set req_qr = rdb.CreateQuery("1reqCG", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT DESC")

    ' for retriveing name,age etc
Set SELTA101_QR = rdb.CreateQuery("selta101_qrcg", "select * from PATOOLSLIB.ta101 where case_no =?")

    ' FOR REFEFING DOCTOR
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
Set REFDR_QR = rdb.CreateQuery("REFDR1_QRCG", str)

    'FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1_QRCG", str)

    'for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma1041_qrcg", str)
    
    'for provisional diagnosis
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("se1lta10121_qrcg", str)
    
    ' for update report
    str = "update dislib.mpreport set PROCEDURE=?,FRESH=?,ARCHIVAL=?,RESULT=?,IMPRESSION=? where reqno=?"
    Set upddfsreport_qr = rdb.CreateQuery("upddcyreport_qrcg", str)
       
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QRCG", str)
    
    ' for update enterby, provreportby, finalby
    str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?, FREPORTBY=?, FREPORTDT=? WHERE REQNNO=?"
    Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qrcg", str)
   
    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrcg", str)
    
    'for insert addendum report
    str = "INSERT INTO DISLIB.ADDENDUMCY(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
    Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUM_QRCG", str)
   
    str = "SELECT  * FROM DISLIB.ADDENDUMCY WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QRCG", str)
   
    str = "select * from dislib.drequisitn where caseno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_1qrcg", str)

    txt_addendum.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_dept.Enabled = False
    FLDENABL ("DIS")

   str = "insert into  dislib.mpreport(REQNO,PROCEDURE,FRESH,ARCHIVAL,RESULT,IMPRESSION,site) values (?,?,?,?,?,?,?)"
   Set INSFREPORT_QR = rdb.CreateQuery("INSFREPORT_QRmp", str)
  
   str = "SELECT * FROM dISLIB.DMOLEPROC WHERE REQNO=?"
   Set SELMOLEPROC_QR = rdb.CreateQuery("SELMOLEPROC_QRmp", str)
   
   str = "SELECT * FROM dISLIB.DMOLEReq WHERE REQNO=?"
   Set SELMOLEREQ_QR = rdb.CreateQuery("SELMOLEREQ_QRmp", str)

   str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
   Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQRmp", str)
   
   str = "SELECT * FROM DISLIB.mpreport WHERE REQNO=?"
   Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QRmp", str)
    
   str = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?, NMETERIALS=? WHERE REQNO=?"
   Set upddrequisitn_qr = rdb.CreateQuery("UPDSPREPORT1_QRmp", str)
    
'    STR = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
'    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO11_QR1", STR)
'    SELEMPBASINFO_QR(0) = "00000"
'    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
    Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRmp", str)
    
   str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
   Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1QRmp", str)
   
   str = "INSERT INTO RISLIB.DISIMP(REQNO,IMP)VALUES(?,?)"
   Set INSDISIMP_QR = rdb.CreateQuery("INSDISIMP_QRmp", str)
  
   str = "UPDATE RISLIB.DISIMP SET IMP=? WHERE REQNO=?"
   Set UPDDISIMP_QR = rdb.CreateQuery("UPDDISIMP_QRmp", str)
   
End If

'cmb_rpt_class.AddItem "N"
BTNENABL1 ("INI")
Exit Sub
MSG:
MsgBox Err.Description
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

Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
       '' MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) '''commented by Monika on 02/12/2022
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
        
          If REQFLAG = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
                REQFLAG = True
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
        End If
        If viewflag = True Then
            If LOGINTYP = "0" Then
                If SELRREQUISITNFLAG = False Then
                    SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                    SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                    SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                    SELRREQUISITN_QR(3) = "F"
                    SELRREQUISITN_QR(4) = "F"
                    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                    SELRREQUISITNFLAG = True
                Else
                    SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                    SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                    SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                    SELRREQUISITN_QR(3) = "F"
                    SELRREQUISITN_QR(4) = "F"
                    SELRREQUISITN_RS.MoreResults
                    SELRREQUISITN_RS.Requery
                End If
            Else
            If SELRREQUISITNFLAG = False Then
                
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
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
                field_clear
                Exit Sub
            End If
        
        End If
End If

If opt_staff.Value = True Then
    If NEWFLAG = True Then
    
       If REQFLAG = False Then
             req_qr(0) = MaskEdBox1.FormattedText
             req_qr(1) = Trim(txt_modcode.Text)
             Set req_rs = req_qr.OpenResultset(1, 3)
             REQFLAG = True
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
        End If
        If viewflag = True Then
            If LOGINTYP = "0" Then
                If SELRREQUISITNFLAG = False Then
                
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "F"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "F"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
        Else
           If SELRREQUISITNFLAG = False Then
                
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            End If
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
cmb_reqno.SetFocus
 Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
On Error GoTo MSG:
If Opt_patient.Value = True Then
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
        If SELTA101FALG = False Then
            SELTA101_QR(0) = MaskEdBox1.FormattedText
             Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            SELTA101FALG = True
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
        If REFDRFLAG = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
            If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If REFDR_RS("recordno") = "1" Then
            LBL_VPP.Caption = "VPP"
            If Trim(REFDR_RS("RECPTNO")) = "VPP" Then
                MsgBox "Payment is not Done"
                'btnenabl ("INI")
                BTNENABL1 ("INI")
                field_clear
                Exit Sub
            End If
            End If
            If REFDR_RS("recordno") = "2" Then
            LBL_VPP.Caption = "LR"
            End If
            If REFDR_RS("recordno") = "3" Then
            LBL_VPP.Caption = "BAR"
            End If
            If REFDR_RS("recordno") = "4" Then
            LBL_VPP.Caption = "DrOp"
            End If
        End If
'            If REFDR_RS.RowCount > 0 Then
'            txt_refdoc.Text = REFDR_RS("REFBYDOC")
'            End If
            If selma104flag = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                selma104flag = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
      
        If selma104_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
            txt_ward.Text = selma104_rs("WARD_NAME")
            txt_bedno.Text = selma104_rs("bed_no")
            txt_status.Text = "IP"
        Else
            txt_status.Text = "OP"
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            lbl_bed.Visible = False
        End If
      
     
        If SELPATHDREQUISITNFLAG = False Then
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
            SELPATHDREQUISITNFLAG = True
        Else
            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            selpathdrequisitn_rs.MoreResults
            selpathdrequisitn_rs.Requery
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
        End If
        
  If NEWFLAG = True Then
        NEWFLAG = False
         If SELMOLEPROC_FLAG = False Then
            SELMOLEPROC_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELMOLEPROC_RS = SELMOLEPROC_QR.OpenResultset(1, 3)
            SELMOLEPROC_FLAG = True
        Else
            SELMOLEPROC_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELMA301RS = SELMOLEPROC_QR.OpenResultset(1, 3)
            SELMOLEPROC_RS.MoreResults
            SELMOLEPROC_RS.Requery
        End If
        If SELMOLEPROC_RS.RowCount > 0 Then
           SELMOLEPROC_RS.MoveFirst
         While Not SELMOLEPROC_RS.EOF
           SELMA101_QR(0) = Trim(SELMOLEPROC_RS("PRO_CODE"))
           If SELMA101_FLAG = False Then
             Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
             SELMA101_FLAG = True
           Else
             SELMA101_RS.MoreResults
             SELMA101_RS.Requery
           End If
             If SELMA101_RS.RowCount > 0 Then
                  If Trim(txt_procedure.Text) <> "" Then
                    txt_procedure = Trim(txt_procedure.Text) & ", " & SELMA101_RS(1)
                  Else
                   txt_procedure.Text = Trim(SELMA101_RS(1))
                  End If
                    
             End If
            SELMOLEPROC_RS.MoveNext
          Wend
        End If
        
        If SELMOLEREQ_FLAG = False Then
            SELMOLEREQ_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELMOLEREQ_RS = SELMOLEREQ_QR.OpenResultset(1, 3)
            SELMOLEREQ_FLAG = True
        Else
            SELMOLEREQ_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELMA301RS = SELMOLEREQ_QR.OpenResultset(1, 3)
            SELMOLEREQ_RS.MoreResults
            SELMOLEREQ_RS.Requery
        End If
        If SELMOLEREQ_RS.RowCount > 0 Then
           txt_fresh.Text = SELMOLEREQ_RS("FRESH")
           txt_archive.Text = SELMOLEREQ_RS("ARCHIVE")
           cmb_site.Text = SELMOLEREQ_RS("SITE")
           
        End If
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = MaskEdBox1.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = MaskEdBox1.FormattedText
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
        If selta1012_rs.RowCount > 0 Then
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
''        If SELDREQUISITNFLAG = False Then
''            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
''            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''            SELDREQUISITNFLAG = True
''        Else
''            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
''            seldrequisitn_rs.MoreResults
''            seldrequisitn_rs.Requery
''        End If
''        If seldrequisitn_rs.RowCount > 0 Then
''         prpathno = ""
''         While Not seldrequisitn_rs.EOF
''          prpathno = prpathno & " " & seldrequisitn_rs("pathno")
''          seldrequisitn_rs.MoveNext
''         Wend
''        End If
''        txt_pathno.Text = Trim(prpathno)
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
        
    End If

    If viewflag = True Then
        viewflag = False
         
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = MaskEdBox1.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = MaskEdBox1.FormattedText
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
        If selta1012_rs.RowCount > 0 Then
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
        If seldfreportflag = False Then
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
            seldfreportflag = True
        Else
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDFREPORT_RS.MoreResults
            SELDFREPORT_RS.Requery
        End If
                    
        
           If SELDFREPORT_RS.RowCount > 0 And SELDFREPORT_RS.EOF = False Then
              txt_archive.Text = " " & Trim(SELDFREPORT_RS("archival"))
              txt_fresh.Text = " " & Trim(SELDFREPORT_RS("fresh"))
              cmb_result.Text = " " & Trim(SELDFREPORT_RS("result"))
              txt_procedure.Text = " " & Trim(SELDFREPORT_RS("PROCEDURE"))
              cmb_site.Text = " " & Trim(SELDFREPORT_RS("site"))
                 If IsNull(SELDFREPORT_RS("impression")) Then
                     txt_impression.Text = ""
                 Else
                     txt_impression.Text = SELDFREPORT_RS("IMPRESSION").Value
                 End If
           End If
            
        If REFDR_RS.RowCount > 0 Then
         If SELDFREPORT_RS.EOF = False Then
            cmb_enteryby.Text = Trim(REFDR_RS("ENTERBY"))
            cmb_reportby.Text = Trim(REFDR_RS("PREPORTBY"))
               If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                cmb_finalby.Text = ""
                Else
                cmb_finalby.Text = Trim(REFDR_RS("FREPORTBY"))
                End If
         End If
        Else
             cmb_finalby.Text = ""
        End If
                
        If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
        If seldrequisitn_rs.RowCount > 0 Then
         prpathno = ""
'         CMB_PATHNO.clear
         While Not seldrequisitn_rs.EOF
            prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'            CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
            seldrequisitn_rs.MoveNext
         Wend
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
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
        'End If
        cmb_reqno.Enabled = False
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
'        txt_clindiag.Enabled = False1
    End If

End If
If opt_staff.Value = True Then
    STAFF
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub
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
    lbl_caseno.Caption = "CC No."
    ' MaskEdBox1.Mask = "#######"   ' COMMENTED BY VARINDER
    MaskEdBox1.Enabled = True
    'MaskEdBox1.SetFocus
End If
End Sub

Private Sub txt_fullreport_GotFocus()
On Error GoTo m14:
w.Visible = False
Set D = w.ActiveDocument
D.Content = ""
D.Content = txt_fullreport
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
D.Content = txt_fullreport
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set D = w.ActiveDocument
D.Content = txt_fullreport.Text
w.Visible = True
w.Activate
End Sub

Private Sub txt_fullreport_LostFocus()
On Error GoTo m15:
Set D = w.ActiveDocument
txt_fullreport.Text = D.Content
'flagimp = False
D.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
D.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Public Sub assign()
On Error GoTo MSG:
Dim I As Integer
Dim MICROSCOPIC As String
rdb.BeginTrans
If flagsave = 1 Then
   flagsave = 0

''    '**** INSERTING INTO MPREPORT

        INSFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
        INSFREPORT_QR(1) = Trim(txt_procedure.Text)
        If Len(Trim(txt_fresh.Text)) > 0 Then
            INSFREPORT_QR(2) = Null
            INSFREPORT_QR(2) = txt_fresh.Text
        Else
            INSFREPORT_QR(2) = Null
        End If
        If Len(Trim(txt_archive.Text)) > 0 Then
            INSFREPORT_QR(3) = Null
            INSFREPORT_QR(3) = txt_archive.Text
        Else
            INSFREPORT_QR(3) = Null
        End If
        If Len(Trim(cmb_result.Text)) > 0 Then
            INSFREPORT_QR(4) = Null
            INSFREPORT_QR(4) = cmb_result.Text
        Else
            INSFREPORT_QR(4) = Null
        End If
        If Len(Trim(txt_impression.Text)) > 0 Then
            INSFREPORT_QR(5) = Null
            INSFREPORT_QR(5) = txt_impression.Text
        Else
            INSFREPORT_QR(5) = Null
        End If
        If Len(Trim(cmb_site.Text)) > 0 Then
            INSFREPORT_QR(6) = Null
            INSFREPORT_QR(6) = cmb_site.Text
        Else
            INSFREPORT_QR(6) = Null
        End If
            INSFREPORT_QR.Execute
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    UPDDREQ_QR(0) = ""
    UPDDREQ_QR(1) = Trim(req_chglen(cmb_reqno.Text))
    UPDDREQ_QR.Execute
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    '**** INSERTING INTO MPREPORT
    
    
    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_reportby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    
'    upddrequisitn_qr(0) = ""
'    If Len(Trim(txt_typespeciman.Text)) > 0 Then
'        upddrequisitn_qr(1) = Null
'        upddrequisitn_qr(1) = txt_typespeciman.Text
'    Else
'        upddrequisitn_qr(1) = Null
'        upddrequisitn_qr(1) = " "
'    End If
'    upddrequisitn_qr(2) = req_chglen(cmb_reqno.Text)
'    upddrequisitn_qr.Execute
    
     '**** INSERTING INTO DISIMP
        INSDISIMP_QR(0) = req_chglen(cmb_reqno.Text)
        INSDISIMP_QR(1) = Null
        INSDISIMP_QR(1) = Trim(txt_impression.Text)
        INSDISIMP_QR.Execute
    
  '**** INSERTING INTO DISIMP

End If

If flagsave = 2 Then
    flagsave = 0
    
    upddfsreport_qr(0) = Null
    upddfsreport_qr(0) = Trim(txt_procedure.Text)
    upddfsreport_qr(1) = Null
    upddfsreport_qr(1) = Trim(txt_fresh.Text)
    upddfsreport_qr(2) = " "
    upddfsreport_qr(2) = Trim(txt_archive.Text)
    upddfsreport_qr(3) = " "
    upddfsreport_qr(3) = Trim(cmb_result.Text)
    upddfsreport_qr(4) = " "
    upddfsreport_qr(4) = Trim(txt_impression.Text)
    upddfsreport_qr(5) = Trim(req_chglen(cmb_reqno.Text))
    upddfsreport_qr.Execute
    
''    upddrequisitn_qr(0) = ""
''    If Len(Trim(txt_typespeciman.Text)) > 0 Then
''        upddrequisitn_qr(1) = Null
''        upddrequisitn_qr(1) = txt_typespeciman.Text
''    Else
''        upddrequisitn_qr(1) = Null
''        upddrequisitn_qr(1) = " "
''    End If
''    upddrequisitn_qr(2) = req_chglen(cmb_reqno.Text)
''    upddrequisitn_qr.Execute
    
    updrrequisitn_qr(0) = cmb_enteryby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_reportby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
   
    '**** UPDATE THE DISIMP
        If Len(Trim(txt_impression.Text)) > 0 Then
        UPDDISIMP_QR(0) = Null
        UPDDISIMP_QR(0) = Trim(txt_impression.Text)
        UPDDISIMP_QR(1) = Trim(req_chglen(cmb_reqno.Text))
        UPDDISIMP_QR.Execute
        End If
        
    '*** UPDATE THE DISIMP
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    UPDDREQ_QR(0) = ""
    UPDDREQ_QR(1) = Trim(req_chglen(cmb_reqno.Text))
    UPDDREQ_QR.Execute
    
    ''*******UPDATE DREQUISITN BY SLIDENO.
    
    
    '**** UPDATING INTO FREPORT
    
    If ADDENDUMFALG = True Then
        ''********SECURITY***************************
        CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
        CURUSRDET_RS.MoreResults
        CURUSRDET_RS.Requery
        If CURUSRDET_RS.RowCount > 0 Then
            If LOGINTYP > CURUSRDET_RS("usrtype") Then
              ''   no action as he can commit
            ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
                MsgBox "Sorry! U'r not an Authorized person to write an Addendum !!", vbInformation, "Pathlogy"
                cmd_save.SetFocus
                Exit Sub
            End If
            If CURUSRDET_RS("ADDENDUM") <> "YES" Then
                MsgBox "Sorry! You are not authorised to Write Addendum of this Modality"
                cmd_save.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
            cmd_commit.SetFocus
            Exit Sub
        End If
''********SECURITY***************************
        If Len(Trim(cmb_addendumby.Text)) = 0 Then
            MsgBox "The Addendum by field is EMPTY.", vbCritical
            cmb_addendumby.SetFocus
            Exit Sub
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
  MsgBox "Record is Successfully Saved."
  Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
End Sub

Public Sub field_clear()
Dim obj As Object
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
cmb_reqno.ListIndex = -1
cmb_reqno.clear
cmb_addendumby.ListIndex = -1
cmb_addendumby.Text = ""
cmb_enteryby.ListIndex = -1
cmb_enteryby.Text = ""
cmb_reportby.ListIndex = -1
cmb_reportby.Text = ""
cmb_finalby.ListIndex = -1
cmb_finalby.Text = ""
cmb_site.ListIndex = -1
cmb_site.Text = ""
'cmb_laterality.ListIndex = -1
'cmb_laterality.Text = ""
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
If Opt_patient.Value = True Then
    'MaskEdBox1.Mask = ">??/#####"
ElseIf opt_staff.Value = True Then
    MaskEdBox1.Mask = ">######"
End If
LBL_VPP.Caption = ""
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In molecular_report.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    MaskEdBox1.Enabled = True
    txt_addendum.Enabled = True
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_reportby.Enabled = True
    cmb_addendumby.Enabled = True
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_dept.Enabled = True
    cmb_reqno.Enabled = True
    txt_procedure.Enabled = True
    cmb_result.Enabled = True
    txt_impression.Enabled = True
    txt_addendum.Enabled = False
End If
If str1 = "DIS" Then
    txt_addendum.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_reportby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    txt_procedure.Enabled = False
    txt_impression.Enabled = False
    txt_addendum.Enabled = False
    cmb_result.Enabled = False
End If
If str1 = "VIEW" Then
    txt_addendum.Enabled = False
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_reportby.Enabled = False
    cmb_addendumby.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_procedure.Enabled = False
    txt_impression.Enabled = False
    txt_addendum.Enabled = False
    cmb_result.Enabled = False
End If
If str1 = "MODIFY" Then
    txt_addendum.Enabled = False
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = True
    cmb_reportby.Enabled = True
    cmb_addendumby.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_procedure.Enabled = True
    txt_impression.Enabled = True
    txt_addendum.Enabled = False
    cmb_result.Enabled = True
End If
If str1 = "ADDENDUM" Then
    txt_addendum.Enabled = True
    cmb_enteryby.Enabled = True
    cmb_finalby.Enabled = False
    cmb_reportby.Enabled = False
'    cmb_fsd.Enabled = False
    cmb_reqno.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_dept.Enabled = False
    txt_procedure.Enabled = False
    txt_impression.Enabled = False
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
        cmb_addendumby.Enabled = True
    Else
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
    cmb_result.Enabled = False
 End If
End Sub

Public Sub STAFF()
On Error GoTo MSG:
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
    If SELDREQSTAFFD1FLAG1 = False Then
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
        SELDREQSTAFFD1QR1(1) = MaskEdBox1.Text
        Set SELDREQSTAFFD1RS1 = SELDREQSTAFFD1QR1.OpenResultset(1, 3)
    Else
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
        SELDREQSTAFFD1QR1(1) = MaskEdBox1.Text
        SELDREQSTAFFD1RS1.MoreResults
        SELDREQSTAFFD1RS1.Requery
    End If
    If SELDREQSTAFFD1RS1.RowCount > 0 Then
        txt_name = SELDREQSTAFFD1RS1("PNAME")
        txt_sex = SELDREQSTAFFD1RS1("SEX")
        txt_age = SELDREQSTAFFD1RS1("AGE")
        If SELDREQSTAFFD1RS1("RELATION") <> "SELF" Then
            txt_status = "DP"
        Else
            txt_status = "ST"
        End If
    End If

     If REFDRFLAG = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If

    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
  
        If selma104flag = False Then
            selma104_qr(0) = MaskEdBox1.FormattedText
            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
            selma104flag = True
        Else
            selma104_qr(0) = MaskEdBox1.FormattedText
            selma104_rs.MoreResults
            selma104_rs.Requery
        End If
  
    If selma104_rs.RowCount > 0 Then
        txt_ward.Text = selma104_rs("WARD_NAME")
        txt_bedno.Text = selma104_rs("bed_no")
        txt_status.Text = "IP"
    Else
        txt_status.Text = "OP"
    End If
 
''        If SELPATHDREQUISITNFLAG = False Then
''            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''            Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
''            SELPATHDREQUISITNFLAG = True
''        Else
''            SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
''            selpathdrequisitn_rs.MoreResults
''            selpathdrequisitn_rs.Requery
''        End If
''    '****
''        If selpathdrequisitn_rs.RowCount > 0 Then
''            txt_cytology.Text = selpathdrequisitn_rs("PATHNO")
''            If selpathdrequisitn_rs.EOF = False Then
''                txt_typespeciman.Text = " " & selpathdrequisitn_rs("NMETERIALS")
''            Else
''                txt_typespeciman.Text = ""
''            End If
''            If IsNull(Trim(selpathdrequisitn_rs("SITE"))) Then
''                'txt_sites.Text = ""
''            Else
''                If SELMOLEPROC_FLAG = False Then
''                    SELMOLEPROC_QR(0) = selpathdrequisitn_rs("SITE")
''                    Set SELMOLEPROC_RS = SELMOLEPROC_QR.OpenResultset(1, 3)
''                    SELMOLEPROC_FLAG = True
''                Else
''                    SELMOLEPROC_QR(0) = selpathdrequisitn_rs("SITE")
''                    Set SELMA301RS = SELMOLEPROC_QR.OpenResultset(1, 3)
''                    SELMOLEPROC_RS.MoreResults
''                    SELMOLEPROC_RS.Requery
''                End If
''                If SELMOLEPROC_RS.RowCount > 0 Then
''                    TXT_CSITE.Text = SELMOLEPROC_RS("SITEDESC")
''                    'txt_noofslides.Text = SELPATHDREQUISITN_RS("OUTSLIDES")
''                    Else
''                            If VIEWSITEDUPRSFLAG = False Then
''                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
''                                 Set viewsiteduprs = viewsitedupqr.OpenResultset(1, 3)
''                                 VIEWSITEDUPRSFLAG = True
''                            Else
''                                 viewsitedupqr(0) = selpathdrequisitn_rs("SITE")
''                                 viewsiteduprs.MoreResults
''                                 viewsiteduprs.Requery
''                            End If
''                            If viewsiteduprs.RowCount > 0 Then
''                                txt_sites.Text = viewsiteduprs("SITEDESC")
''                            End If
''                    End If
''            End If
''        End If

    '****
    If SELTA1012FLAG = False Then
            selta1012_qr(0) = MaskEdBox1.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = MaskEdBox1.FormattedText
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
        
        If selta1012_rs.RowCount > 0 Then
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
        If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
        If seldrequisitn_rs.RowCount > 0 Then
         prpathno = ""
         While Not seldrequisitn_rs.EOF
         prpathno = prpathno & " " & seldrequisitn_rs("pathno")
         'CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
         seldrequisitn_rs.MoveNext
         Wend
'         txt_prepathreqno.Text = prpathno
        End If
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
        

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
    If SELIMMHISCHE_FLAG = False Then
        SELIMMHISCHE_QR(0) = VAR_REQNNO
        Set SELIMMHISCHE_RS = SELIMMHISCHE_QR.OpenResultset(1, 3)
        SELIMMHISCHE_FLAG = True
    Else
        SELIMMHISCHE_QR(0) = VAR_REQNNO
        SELIMMHISCHE_RS.MoreResults
        SELIMMHISCHE_RS.Requery
    End If
    If Not SELIMMHISCHE_RS.EOF Then
        If Trim(SELIMMHISCHE_RS("IMMHISCHE")) = "P" Then
                cmb_rpt_class.Text = "Pos"
            Else
                cmb_rpt_class.Text = "Neg"
            End If
    End If
    If SELDREQSTAFFD1FLAG1 = False Then
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
        SELDREQSTAFFD1QR1(1) = Trim(MaskEdBox1.Text)
        Set SELDREQSTAFFD1RS1 = SELDREQSTAFFD1QR1.OpenResultset(1, 3)
    Else
        SELDREQSTAFFD1QR1(0) = req_chglen(cmb_reqno)
        SELDREQSTAFFD1QR1(1) = Trim(MaskEdBox1.Text)
        SELDREQSTAFFD1RS1.MoreResults
        SELDREQSTAFFD1RS1.Requery
    End If
    If SELDREQSTAFFD1RS1.RowCount > 0 Then
        txt_name = SELDREQSTAFFD1RS1("PNAME")
        txt_sex = SELDREQSTAFFD1RS1("SEX")
        txt_age = SELDREQSTAFFD1RS1("AGE")
        If Trim(SELDREQSTAFFD1RS1("RELATION")) <> "SELF" Then
            txt_status = "DP"
        Else
            txt_status = "ST"
        End If
    End If

    If REFDRFLAG = False Then
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            REFDRFLAG = True
        Else
            REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
            REFDR_QR(1) = req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If

    If REFDR_RS.RowCount > 0 Then
        '***************************************
            If REFDR_RS("RECORDNO") = 1 Then
                LBL_VPP.Caption = "VPP"
            End If
            If REFDR_RS("RECORDNO") = 2 Then
                LBL_VPP.Caption = "LR"
            End If
            If REFDR_RS("RECORDNO") = 3 Then
                LBL_VPP.Caption = "BAR"
            End If
            If REFDR_RS("RECORDNO") = 4 Then
                LBL_VPP.Caption = "DrOp"
            End If
        
      '***************************************
        txt_refdoc.Text = Trim(REFDR_RS("REFBYDOC"))
        If REFDR_RS("COMMITFLAG") = "Y" Then
            MsgBox "Final Report is Generated, You cann't Modify the Final Report,You can Add ADDENDUMS Only"
            cmd_modify.Enabled = False
            ADDENDUMFALG = True
            txt_addendum.Enabled = True
        End If
    End If
  
    If selma104flag = False Then
            selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
            selma104flag = True
        Else
            selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
            selma104_rs.MoreResults
            selma104_rs.Requery
        End If
  
  
'    If selma104_rs.RowCount > 0 Then
'        txt_ward.Text = selma104_rs("WARD_NAME")
'        txt_bedno.Text = selma104_rs("bed_no")
'        txt_status.Text = "IP"
'    Else
'        txt_status.Text = "OP"
'    End If
            
           
        End If
    
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
            selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
            selta1012_rs.MoreResults
            selta1012_rs.Requery
        
        End If
        If selta1012_rs.RowCount > 0 Then
            txt_clndiag.Text = selta1012_rs("prvdiagdes")
        End If
        If seldfreportflag = False Then
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
            seldfreportflag = True
        Else
            SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
            SELDFREPORT_RS.MoreResults
            SELDFREPORT_RS.Requery
        End If
                    
        
           If SELDFREPORT_RS.RowCount > 0 Then
               If SELDFREPORT_RS.EOF = False Then
               If IsNull(SELDFREPORT_RS("CYREPORT")) Then
                       txt_impression.Text = ""
                       Else
                       SELDFREPORT_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                        SELDFREPORT_RS.MoreResults
                        SELDFREPORT_RS.Requery
                       txt_impression.Text = SELDFREPORT_RS("CYREPORT").Value
                       End If
'
'                   txt_report.Text = SELDFREPORT_RS("CYREPORT").Value
                Else
                   txt_impression.Text = ""
                End If
                
                If SELDFREPORT_RS.EOF = False Then
                   txt_impression.Text = SELDFREPORT_RS("IMPRESSION").Value
                Else
                   txt_impression.Text = ""
                End If
            Else
                'txt_immhischerep.Text = ""
            End If
            
        If REFDR_RS.RowCount > 0 Then
         If SELDFREPORT_RS.EOF = False Then
            cmb_enteryby.Text = Trim(REFDR_RS("ENTERBY"))
            cmb_reportby.Text = Trim(REFDR_RS("PREPORTBY"))
               If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                cmb_finalby.Text = ""
                Else
                cmb_finalby.Text = Trim(REFDR_RS("FREPORTBY"))
                End If
         End If
        Else
             cmb_finalby.Text = ""
        End If
                
        If SELDREQUISITNFLAG = False Then
            seldrequisitn_qr(0) = Trim(MaskEdBox1.FormattedText)
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQUISITNFLAG = True
        Else
            seldrequisitn_qr(0) = Trim(MaskEdBox1.FormattedText)
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
        If seldrequisitn_rs.RowCount > 0 Then
         prpathno = ""
'         CMB_PATHNO.clear
         While Not seldrequisitn_rs.EOF
            prpathno = prpathno & " " & seldrequisitn_rs("pathno")
'            CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
            seldrequisitn_rs.MoveNext
         Wend
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
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
        'End If
        cmb_reqno.Enabled = False
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
'        txt_clindiag.Enabled = False1
  

 
    
    
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Timer1_Timer()
DATETIME
txt_time.Text = time1
txt_date.Text = DATE11
End Sub

Private Sub txt_addendum_GotFocus()
On Error GoTo m3:
W_grossdes.Visible = False
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = ""
d_grossdesc.Content = txt_grossdesc
W_grossdes.Caption = "Addendum"
W_grossdes.WindowState = wdWindowStateMaximize
d_grossdesc.Content = txt_grossdesc
W_grossdes.Visible = True
W_grossdes.Activate
m3:
On Error Resume Next:
Set W_grossdes = CreateObject("word.Application")
W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
W_grossdes.Caption = "Addendum"
W_grossdes.WindowState = wdWindowStateMaximize
Set d_grossdesc = W_grossdes.ActiveDocument
d_grossdesc.Content = txt_addendum.Text
W_grossdes.Visible = True
W_grossdes.Activate
End Sub

Private Sub txt_addendum_LostFocus()
On Error GoTo m4:
Set d_grossdesc = W_grossdes.ActiveDocument
txt_addendum.Text = d_grossdesc.Content
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





Private Sub txt_impression_gotfocus()
On Error GoTo m3:
WREPORT.Visible = False
Set DREPORT = WREPORT.ActiveDocument
DREPORT.Content = ""
DREPORT.Content = txt_impression
WREPORT.Caption = "Report"
WREPORT.WindowState = wdWindowStateMaximize
DREPORT.Content = txt_impression
WREPORT.Visible = True
WREPORT.Activate
m3:
On Error Resume Next:
Set WREPORT = CreateObject("word.Application")
WREPORT.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
WREPORT.Caption = "Report"
WREPORT.WindowState = wdWindowStateMaximize
Set DREPORT = WREPORT.ActiveDocument
DREPORT.Content = txt_impression.Text
WREPORT.Visible = True
WREPORT.Activate
End Sub
Private Sub txt_impression_lostfocus()
On Error GoTo m4:
Set DREPORT = WREPORT.ActiveDocument
txt_impression.Text = DREPORT.Content
DREPORT.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DREPORT.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DREPORT.Close (wdDoNotSaveChanges)
WREPORT.quit
Exit Sub
m4:
MsgBox Err.Description
End Sub



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
End Sub

