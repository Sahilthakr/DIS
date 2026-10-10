VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form micro_report 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Microbiology Laboratory"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame common 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   11775
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   9360
         TabIndex        =   70
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox cmb_bioHazard 
         Height          =   315
         ItemData        =   "micro_report.frx":0000
         Left            =   10320
         List            =   "micro_report.frx":000A
         TabIndex        =   69
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txt_modality 
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
         Left            =   6240
         TabIndex        =   65
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txt_dept 
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
         TabIndex        =   64
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox CMB_PATHNO 
         Height          =   315
         Left            =   10200
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         ItemData        =   "micro_report.frx":0017
         Left            =   6240
         List            =   "micro_report.frx":0019
         TabIndex        =   34
         Top             =   600
         Width           =   2175
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
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   1.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1215
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   32
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Opt_patient 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   975
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   1215
         End
      End
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
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   46
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BioHazard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9000
         TabIndex        =   68
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   63
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   62
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   61
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   60
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_ward 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   59
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   58
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   57
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   56
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   55
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   54
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   53
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   52
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   51
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   50
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   48
         Top             =   360
         Width           =   1830
      End
   End
   Begin VB.Frame REPORT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   11775
      Begin VB.TextBox txt_result 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   7320
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   5640
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   9960
         TabIndex        =   27
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton cmd_adenum 
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
         Left            =   10680
         TabIndex        =   24
         Top             =   3000
         Width           =   1095
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
         TabIndex        =   22
         Top             =   2760
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
         TabIndex        =   20
         Top             =   3120
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   3480
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
         Left            =   5760
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   3480
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2415
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   0
         AllowUserResizing=   1
         MousePointer    =   1
         FormatString    =   "^S.No.|Test  Name             |^Observed Value|^Result            |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Left            =   8640
         TabIndex        =   26
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lbl_addendnum 
         BackStyle       =   0  'Transparent
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
         Left            =   480
         TabIndex        =   25
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label LBL_COMMENTS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   600
         TabIndex        =   21
         Top             =   2760
         Width           =   750
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   600
         TabIndex        =   19
         Top             =   3480
         Width           =   990
      End
      Begin VB.Label lbl_final 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   4065
         TabIndex        =   18
         Top             =   3480
         Width           =   1590
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame COMMAND 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   7
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
         Left            =   240
         MouseIcon       =   "micro_report.frx":001B
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   120
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
         Left            =   1680
         MouseIcon       =   "micro_report.frx":0325
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   120
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
         MouseIcon       =   "micro_report.frx":062F
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
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
         Left            =   4680
         MouseIcon       =   "micro_report.frx":0939
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   120
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
         Left            =   6240
         MouseIcon       =   "micro_report.frx":0C43
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   120
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
         Left            =   9240
         MouseIcon       =   "micro_report.frx":0F4D
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   120
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
         Left            =   10560
         MouseIcon       =   "micro_report.frx":1257
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   120
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
         Left            =   7680
         MouseIcon       =   "micro_report.frx":1561
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   120
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
      Left            =   10800
      TabIndex        =   6
      Top             =   120
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
      Left            =   10800
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   330
      Left            =   720
      TabIndex        =   23
      Top             =   0
      Width           =   8595
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   10200
      TabIndex        =   15
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   10200
      TabIndex        =   14
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl_biochemlab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT - MICROBIOLOGY LABORATORY "
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
      TabIndex        =   13
      Top             =   480
      Width           =   6000
   End
End
Attribute VB_Name = "micro_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''' Designed and Done by KANCHAN

Dim str As String
Dim R1, C1
Dim RS As rdoResultset
Dim sel_deptqr As rdoQuery
Dim sel_deptrs As rdoResultset

Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset

Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset


Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset
Dim SELDREQSTAFFDQR As rdoQuery
Dim SELDREQSTAFFDRS As rdoResultset
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset

Dim newdmulproce_qr As rdoQuery
Dim newdmulproce_rs As rdoResultset
Dim retminterface_qr As rdoQuery
Dim retminterface_rs As rdoResultset

Dim insmicroproce_qr As rdoQuery
Dim insmicroproce_rs As rdoResultset
Dim insmicrotemp_qr As rdoQuery
Dim deletemicrotemp As rdoQuery
Dim viewmicroproce_qr As rdoQuery
Dim viewmicroproce_rs As rdoResultset

Dim proname_qr As rdoQuery
Dim proname_rs As rdoResultset
Dim procode_qr As rdoQuery
Dim procode_rs As rdoResultset

Dim savrreq_qr As rdoQuery
Dim modrreq_qr As rdoQuery
Dim modrreq_rs As rdoResultset

Dim USRNAME_QR As rdoQuery
Dim USRNAME_RS As rdoResultset
Dim REPWTRN_QR As rdoQuery
Dim REPWTRNRS As rdoResultset

'''''''''modify
Dim updatemicroproce_qr As rdoQuery
Dim testcode_qr As rdoQuery
Dim testcode_rs As rdoResultset
Dim testname As Boolean
''''

Dim UPDRTIMECADD_QR As rdoQuery
Dim UPDRTIME_QR As rdoQuery
Dim UPDRTIMEM_QR As rdoQuery
Dim UPDRTIMEP_QR As rdoQuery
Dim UPDRTIMEC_QR As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim INSDMBCCOMMQR As rdoQuery
Dim UPDDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMRS As rdoResultset

Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset
Dim UPDMINTERFACE_QR As rdoQuery


'''''''

Dim REQFLAG1 As Boolean
Dim SELDEPTRSFLAG1 As Boolean
Dim selta101flag1 As Boolean
Dim REFDRFLAG1 As Boolean
Dim SELPATHDREQUISITNFALG1 As Boolean
Dim SELDREQSTAFFD1FLAG As Boolean
Dim REFDRFLAG As Boolean
Dim SELRREQUISITNFLAG1 As Boolean

Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim MODFLAG As Boolean
Dim flagsave As Integer
Dim viewrec As Boolean
Dim newrec As Boolean
Dim pronameflag As Boolean
Dim pronameflag1 As Boolean
Dim modrreqflag As Boolean
Dim adenum_flg As Boolean
Dim proname_flg As Boolean
Dim RETMINTERFACE As Boolean
Dim SELDMBCCOMMFLAG As Boolean
Dim DATETIMEFLAG As Boolean


Dim LOGINESRNAME As String
Dim LOGINTYPE As String
Dim LOGINTYP As String 'SHWETA 06-NOV=2012
Dim LOGINUSRID As String

Dim a As Integer

Dim EN As HMSCrypt

Public Sub assign()

''''save record in DISLIB.MICROPROCE table
If NEWFLAG = True Then
    
    For i = 1 To MSFlexGrid1.Rows - 1

        If proname_flg = False Then
            procode_qr(0) = MSFlexGrid1.TextMatrix(i, 1)
            Set procode_rs = procode_qr.OpenResultset(1, 3)
            'proname_flg = True
        Else
            procode_qr(0) = MSFlexGrid1.TextMatrix(i, 1)
            procode_rs.MoreResults
            procode_rs.Requery
        End If
        
  ''********* COMMENT IF CONDITION TO INSERT ALL ROWS AT SANGRUR BY RAHUL ON DATE 12.12.2016  ***************'
'Dim RD_MICRO As rdoResultset
'Set RD_MICRO = rdb.OpenResultset("SELECT * FROM DISLIB.MICROPROCE WHERE REQNO='" & Trim(Replace(cmb_reqno.Text, "/", "")) & "'", 1, 1)
'If RD_MICRO.EOF Then
            insmicroproce_qr(0) = req_chglen(cmb_reqno.Text)
            insmicroproce_qr(1) = procode_rs(0)
           
            a = makekey
            If Len(MSFlexGrid1.TextMatrix(i, 2)) > 0 Then
            insmicroproce_qr(2) = EN.Encrypt(MSFlexGrid1.TextMatrix(i, 2), a)
            insmicroproce_qr(6) = Trim(MSFlexGrid1.TextMatrix(i, 2))
            Else
            insmicroproce_qr(6) = ""
            End If
            
            If Len(MSFlexGrid1.TextMatrix(i, 3)) > 0 Then
            insmicroproce_qr(3) = EN.Encrypt(MSFlexGrid1.TextMatrix(i, 3), a)
            insmicroproce_qr(5) = Trim(MSFlexGrid1.TextMatrix(i, 3))
            Else
            insmicroproce_qr(3) = ""
            insmicroproce_qr(5) = ""
            End If
            
            insmicroproce_qr(4) = CInt(a)
            
            insmicroproce_qr.Execute
            
                   ' End If     'AT SANGRUR BY RAHUL ON DATE 12.12.2016  ***************'


    Next i
    
            INSDMBCCOMMQR(0) = Trim(req_chglen(cmb_reqno.Text))
            
            If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                
              INSDMBCCOMMQR(1) = TXT_COMMENTS.Text
            Else
              INSDMBCCOMMQR(1) = Null

            End If
            INSDMBCCOMMQR.Execute
            
End If

If MODFLAG = True Then
For i = 1 To MSFlexGrid1.Rows - 1

        If testname = False Then
            testcode_qr(0) = MSFlexGrid1.TextMatrix(i, 1)
            Set testcode_rs = testcode_qr.OpenResultset(1, 3)
            'proname_flg = True
        Else
            testcode_qr(0) = MSFlexGrid1.TextMatrix(i, 1)
            testcode_rs.MoreResults
            testcode_rs.Requery
        End If
        
 
            a = makekey
            
            
            
            

            
            
            
            If Len(MSFlexGrid1.TextMatrix(i, 2)) > 0 Then
            updatemicroproce_qr(0) = EN.Encrypt(MSFlexGrid1.TextMatrix(i, 2), a)
            updatemicroproce_qr(4) = Trim(MSFlexGrid1.TextMatrix(i, 2))
            Else
            updatemicroproce_qr(0) = ""
            updatemicroproce_qr(4) = ""
            End If
            
            If Len(MSFlexGrid1.TextMatrix(i, 3)) > 0 Then
            updatemicroproce_qr(1) = EN.Encrypt(MSFlexGrid1.TextMatrix(i, 3), a)
            updatemicroproce_qr(3) = Trim(MSFlexGrid1.TextMatrix(i, 3))
            Else
            updatemicroproce_qr(1) = ""
            updatemicroproce_qr(3) = ""
            End If
            
            updatemicroproce_qr(2) = CInt(a)
            
            updatemicroproce_qr(5) = req_chglen(cmb_reqno.Text)
            updatemicroproce_qr(6) = testcode_rs(0)
            
            updatemicroproce_qr.Execute
            
            
           ''''UPDATE INTERFACE/MINTERFACE
           UPDMINTERFACE_QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
           UPDMINTERFACE_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 3))
           UPDMINTERFACE_QR(2) = req_chglen(cmb_reqno.Text)
           UPDMINTERFACE_QR(3) = testcode_rs(0)
           UPDMINTERFACE_QR.Execute
 
 Next i
 
 '''''''''''''''''''' DISLIB.DMBCCOMM
 If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
 Dim QR_DMBCCOMM As rdoQuery
Dim RS_DMBCCOMM As rdoResultset
str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=? "
Set QR_DMBCCOMM = rdb.CreateQuery("", str)
QR_DMBCCOMM(0) = Trim(req_chglen(cmb_reqno.Text))
Set RS_DMBCCOMM = QR_DMBCCOMM.OpenResultset(1, 3)

                If RS_DMBCCOMM.EOF Then
                
                INSDMBCCOMMQR(0) = Trim(req_chglen(cmb_reqno.Text))
                        
                        If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                            
                          INSDMBCCOMMQR(1) = TXT_COMMENTS.Text
                        Else
                          INSDMBCCOMMQR(1) = Null
                    
                        End If
                        INSDMBCCOMMQR.Execute
            
                
                
                Else
                          If Len(Trim(TXT_COMMENTS.Text)) > 0 Then
                                UPDDMBCCOMMQR(0) = Trim(TXT_COMMENTS.Text)
                            Else
                                UPDDMBCCOMMQR(0) = Null
                            End If
                            UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
                            UPDDMBCCOMMQR.Execute
                 
                
                 
                End If
End If
End If
''''''''update record in the RISLIB.RREQUISITN table

savrreq_qr(0) = cmb_enterby.Text
savrreq_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
savrreq_qr(2) = cmb_finalby.Text
savrreq_qr(3) = "T"
savrreq_qr(4) = cmb_finalby.Text
savrreq_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
savrreq_qr(6) = MaskEdBox1.FormattedText
savrreq_qr(7) = req_chglen(cmb_reqno.Text)

savrreq_qr.Execute






''''(addendum)update record in the RISLIB.RTIMELOG table

    UPDRTIMECADD_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    UPDRTIMECADD_QR(1) = time1
    UPDRTIMECADD_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMECADD_QR.Execute


MsgBox ("Record is saved.")
BTNENABL1 ("INI")
txt_observed.Visible = False
txt_result.Visible = False

End Sub

Public Sub STAFF()

  VAR_REQNNO = req_chglen(cmb_reqno.Text)
 
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
                txt_modcode.Text = REFDR_RS("MODALITYCD")
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
                    ADDENDUMFALG = True
                    cmd_adenum.Visible = True
                End If
                 cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
                'cmb_enterby.Text = REFDR_RS("ENTERBY")
                'cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
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
          
          

End Sub





'Private Sub CMD_NEW_Click()
'
'Dim L1 As Integer
'
'On Error GoTo MSG:
'btnenable1 ("NEW")
'field_clear
'
'MSFlexGrid1.clear
'
'If DATETIMEFLAG = False Then
'   datetime
'   DATETIMEFLAG = True
'Else
'    datetime
'End If
'txt_date.Text = DATE11
'txt_time.Text = time1
'
'cmb_enterby.Text = cmb_enterby.List(0)
'cmb_enterby.Enabled = False
'
'For L1 = 0 To cmb_finalby.ListCount
'    If Trim(cmb_enterby.List(0)) = Trim(cmb_finalby.List(L1)) Then
'        cmb_finalby.Text = Trim(cmb_finalby.List(L1))
'    End If
'Next L1
'
'If SELDEPTRSFLAG1 = False Then
'    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
'    SELDEPTRSFLAG1 = True
'Else
'    sel_deptrs.MoreResults
'    sel_deptrs.Requery
'End If
'
'Opt_patient.Value = True
'flagsave = 1
'viewflag = 0
'NEWFLAG = True
'
''cmb_dept.SetFocus
'MaskEdBox1.Enabled = True
'Exit Sub
'
'MSG:
'    MsgBox Err.Description
'End Sub



Private Sub cmb_reqno_Click()


If viewflag = True Then
'TO INSERT PENDING REQN AND TEST

Dim selMINTERFACEN_qr1 As rdoQuery
Dim selMINTERFACEN_rs1 As rdoResultset
str = "select * from  INTERFACE.MINTERFACE WHERE REQNO=? "
Set selMINTERFACEN_qr1 = rdb.CreateQuery("", str)
selMINTERFACEN_qr1(0) = req_chglen(cmb_reqno.Text)

Set selMINTERFACEN_rs1 = selMINTERFACEN_qr1.OpenResultset(1, 3)
If selMINTERFACEN_rs1.EOF = False Then
    While Not selMINTERFACEN_rs1.EOF
            Dim selMICROPROCEN_qr1 As rdoQuery
            Dim selMICROPROCEN_rs1 As rdoResultset
            str = "select * from  dislib.MICROPROCE WHERE REQNO=? AND PRO_CODE=? "
            Set selMICROPROCEN_qr1 = rdb.CreateQuery("", str)
            selMICROPROCEN_qr1(0) = selMINTERFACEN_rs1("REQNO")
            selMICROPROCEN_qr1(1) = selMINTERFACEN_rs1("TESTCODE")
            Set selMICROPROCEN_rs1 = selMICROPROCEN_qr1.OpenResultset(1, 3)
            If selMICROPROCEN_rs1.EOF Then
            
            insmicroproce_qr(0) = selMINTERFACEN_rs1("REQNO")
            insmicroproce_qr(1) = selMINTERFACEN_rs1("TESTCODE")
           
            a = makekey
            If Len(selMINTERFACEN_rs1("TESTRESULT")) > 0 Then
            insmicroproce_qr(2) = EN.Encrypt(selMINTERFACEN_rs1("TESTRESULT"), a)
            insmicroproce_qr(6) = Trim(selMINTERFACEN_rs1("TESTRESULT"))
            Else
            insmicroproce_qr(6) = ""
            End If
            
            If Len(selMINTERFACEN_rs1("FLD2")) > 0 Then
            insmicroproce_qr(3) = EN.Encrypt(selMINTERFACEN_rs1("FLD2"), a)
            insmicroproce_qr(5) = Trim(selMINTERFACEN_rs1("FLD2"))
            Else
            insmicroproce_qr(3) = ""
            insmicroproce_qr(5) = ""
            End If
            
            insmicroproce_qr(4) = CInt(a)
            
            insmicroproce_qr.Execute
            
            End If
            selMINTERFACEN_rs1.MoveNext
    Wend
End If
End If


If Opt_patient.Value = True Then
        
            If selta101flag1 = False Then
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101flag1 = True
            Else
                SELTA101_QR(0) = MaskEdBox1.FormattedText
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
            End If

            txt_sex.Text = SELTA101_RS("sex")
            txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
            txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            txt_pridoc.Text = SELTA101_RS("ud_name")
            txt_category.Text = SELTA101_RS("pat_catg")
            'txt_status.Text = SELTA101_rs("")

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
          txt_modcode.Text = REFDR_RS("MODALITYCD")
          txt_refdoc.Text = REFDR_RS("REFBYDOC")
          If REFDR_RS("COMMITFLAG") = "Y" Then
              MsgBox "Report is commited , You can not Modify"
              cmd_modify.Enabled = False
              ADDENDUMFALG = True
              cmd_adenum.Visible = True
          End If
          cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
          cmb_finalby.Text = "" & REFDR_RS("FREPORTBY")
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
                TXT_COMMENTS.Text = " " & Trim(SELDMBCCOMMRS("COMMENTS").Value)
            End If
        End If
        
End If


If opt_staff.Value = True Then
    STAFF
End If


''''for retriving data in grid fromtable microproce
    
    '' viewing records
    
    If viewflag = True Then
        
            If viewrec = False Then
                viewmicroproce_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                Set viewmicroproce_rs = viewmicroproce_qr.OpenResultset(1, 3)
                viewrec = True
            Else
                viewmicroproce_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                viewmicroproce_rs.MoreResults
                viewmicroproce_rs.Requery
            End If
      
        If viewmicroproce_rs.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = viewmicroproce_rs.RowCount + 1
          
               While Not viewmicroproce_rs.EOF
                
              ''to get test name
               
                If pronameflag1 = False Then
                    proname_qr(0) = viewmicroproce_rs(1)
                    Set proname_rs = proname_qr.OpenResultset(1, 3)
                    pronameflag = True
                Else
                    proname_qr(0) = viewmicroproce_rs(1)
                    proname_rs.MoreResults
                    proname_rs.Requery
                End If
          
                
                MSFlexGrid1.TextMatrix(i, 0) = i
                MSFlexGrid1.TextMatrix(i, 1) = Trim(proname_rs(1))
                If Trim(viewmicroproce_rs("FLD1")) = "" Then
              Else
                a = CInt(Trim(viewmicroproce_rs("FLD1")))
                
                If Len(Trim(viewmicroproce_rs(2))) > 0 Then
                    'decryption = Trim(viewmicroproce_rs(2))
                    MSFlexGrid1.TextMatrix(i, 2) = EN.Decrypt(Trim(viewmicroproce_rs(2)), a)
                Else
                    MSFlexGrid1.TextMatrix(i, 2) = ""
                End If
                
                If Len(Trim(viewmicroproce_rs(3))) > 0 Then
                    'decryption = Trim(viewmicroproce_rs(3))
                    MSFlexGrid1.TextMatrix(i, 3) = EN.Decrypt(Trim(viewmicroproce_rs(3)), a)
                Else
                    MSFlexGrid1.TextMatrix(i, 3) = ""
                End If
                  End If
                i = i + 1
                viewmicroproce_rs.MoveNext
            Wend
         End If
        
            txt_observed.Visible = False
            txt_result.Visible = False
    End If

    '' Adding new record
    
    If NEWFLAG = True Then
        
            
            If newrec = False Then
                newdmulproce_qr(0) = MaskEdBox1.FormattedText
                newdmulproce_qr(1) = req_chglen(Trim(cmb_reqno.Text))
                Set newdmulproce_rs = newdmulproce_qr.OpenResultset(1, 3)
                newrec = True
            Else
                newdmulproce_qr(0) = MaskEdBox1.FormattedText
                newdmulproce_qr(1) = req_chglen(Trim(cmb_reqno.Text))
                newdmulproce_rs.MoreResults
                newdmulproce_rs.Requery
            End If
                     
       Dim srno As Integer
       
        If newdmulproce_rs.RowCount > 0 Then
           
           i = 1
          
            MSFlexGrid1.Rows = newdmulproce_rs.RowCount + 1
            
                       
            While Not newdmulproce_rs.EOF
          
                If pronameflag = False Then
                    proname_qr(0) = newdmulproce_rs(3)
                    Set proname_rs = proname_qr.OpenResultset(1, 3)
                    pronameflag = True
                Else
                    proname_qr(0) = newdmulproce_rs(3)
                    proname_rs.MoreResults
                    proname_rs.Requery
                End If
          
                If RETMINTERFACE = False Then
                   retminterface_qr(0) = MaskEdBox1.FormattedText
                   retminterface_qr(1) = req_chglen(Trim(cmb_reqno.Text))
                   retminterface_qr(2) = newdmulproce_rs(3)
                   Set retminterface_rs = retminterface_qr.OpenResultset(1, 3)
                Else
                   retminterface_qr(0) = MaskEdBox1.FormattedText
                   retminterface_qr(1) = req_chglen(Trim(cmb_reqno.Text))
                   retminterface_qr(2) = newdmulproce_rs(3)
                   retminterface_rs.MoreResults
                   retminterface_rs.Requery
                End If
                    
                    
                    
                MSFlexGrid1.TextMatrix(i, 0) = Trim(i)
                MSFlexGrid1.TextMatrix(i, 1) = Trim(proname_rs(1))
                If retminterface_rs.RowCount > 0 Then
                MSFlexGrid1.TextMatrix(i, 2) = Trim(retminterface_rs("TESTRESULT"))
                MSFlexGrid1.TextMatrix(i, 3) = Trim(retminterface_rs("FLD2"))
                Else
                MSFlexGrid1.TextMatrix(i, 2) = ""
                MSFlexGrid1.TextMatrix(i, 3) = ""
                End If
                
                i = i + 1
                newdmulproce_rs.MoveNext
            Wend
         End If
    
    
    End If
    
    
    
    
 Dim selBioHazard_qr1 As rdoQuery
Dim selBioHazard_rs1 As rdoResultset
str = "select * from  dislib.biohazard where case_no=?"
Set selBioHazard_qr1 = rdb.CreateQuery("", str)
selBioHazard_qr1(0) = Trim(MaskEdBox1.FormattedText)
Set selBioHazard_rs1 = selBioHazard_qr1.OpenResultset(1, 3)
If selBioHazard_rs1.EOF Then
Else
    If selBioHazard_rs1(1) = "Y" Then
    cmb_bioHazard.Text = "Yes"
    End If
    If selBioHazard_rs1(1) = "N" Then
    cmb_bioHazard.Text = "No"
    End If
End If

    
End Sub


Private Sub cmd_back_Click()

txt_date = Format(DATE11, "dd/mm/yyyy")
txt_time = time1
Me.Hide

DIS_MAIN.show
'btnenabl ("INI")
BTNENABL1 ("INI")
field_clear
viewflag = False
NEWFLAG = False
BCCLEAR

End Sub

Private Sub cmd_cancel_Click()



BTNENABL1 ("INI")
field_clear
viewflag = False
NEWFLAG = False
BCCLEAR

txt_dept.Text = "MICROBIOLOGY"
txt_modality.Text = "SEROLOGY"
txt_deptcode.Text = "GZZ"
txt_modcode.Text = "SE"

txt_date.Text = Format(DATE11, "dd/mm/yyyy")
txt_time.Text = time1

End Sub

Private Sub cmd_commit_Click()

If cmb_bioHazard.Text <> "Yes" And cmb_bioHazard.Text <> "No" Then
MsgBox "Please select Biohazard yes or no"
cmb_bioHazard.SetFocus
Exit Sub
End If

If cmb_bioHazard.Text = "Yes" Then
Dim a
a = MsgBox("Biohazard is Yes??", vbYesNo)
If a = vbNo Then
Exit Sub
End If

End If



Dim i As Integer

txt_date.Text = Format(DATE11, "dd/mm/yyyy")
txt_time.Text = time1
Dim updatemicroproce_qr1 As rdoQuery
 str = "UPDATE DISLIB.MICROPROCE SET FLD3='C' WHERE REQNO=? AND PRO_CODE=?"
 Set updatemicroproce_qr1 = rdb.CreateQuery("", str)
''''''added by shweta 05-02-2014
For i = 1 To MSFlexGrid1.Rows - 1
testcode_qr(0) = MSFlexGrid1.TextMatrix(i, 1)
            Set testcode_rs = testcode_qr.OpenResultset(1, 3)
a = makekey





            If Len(MSFlexGrid1.TextMatrix(i, 2)) > 0 Then
            updatemicroproce_qr(0) = EN.Encrypt(MSFlexGrid1.TextMatrix(i, 2), a)
            updatemicroproce_qr(4) = Trim(MSFlexGrid1.TextMatrix(i, 2))
            Else
            updatemicroproce_qr(0) = ""
            updatemicroproce_qr(4) = ""
            End If
            
            If Len(MSFlexGrid1.TextMatrix(i, 3)) > 0 Then
            updatemicroproce_qr(1) = EN.Encrypt(MSFlexGrid1.TextMatrix(i, 3), a)
            updatemicroproce_qr(3) = Trim(MSFlexGrid1.TextMatrix(i, 3))
            Else
            updatemicroproce_qr(1) = ""
            updatemicroproce_qr(3) = ""
            End If
            
            updatemicroproce_qr(2) = CInt(a)
            
            updatemicroproce_qr(5) = req_chglen(cmb_reqno.Text)
            updatemicroproce_qr(6) = testcode_rs(0)
            
            updatemicroproce_qr.Execute
            
           
            updatemicroproce_qr1(0) = req_chglen(cmb_reqno.Text)
            updatemicroproce_qr1(1) = testcode_rs(0)
            updatemicroproce_qr1.Execute
Next

'''' till here

i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If i = vbYes Then
On Error GoTo MSG:
 
''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
    
        If (LOGINTYP > CURUSRDET_RS("usrtype") Or LOGINTYP >= 6) Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
            cmd_commit.SetFocus
            Exit Sub
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
            'If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
                '  If loginusrid <> curusrdetrs("usr_id") Then
             '   MsgBox "Sorry! This Report is not to be Finalized by You"
              '  cmd_commit.SetFocus
               ' Exit Sub
            'End If
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
    MsgBox "Record is Committed Successfully."
    
   'rtimelog
    UPDRTIMEC_QR(0) = Format(DATE11, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(time1, "HH:MM:SS")
                    
    UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute
   'rtimelog
    
End If



Dim selBioHazard_qr1 As rdoQuery
Dim selBioHazard_rs1 As rdoResultset
str = "select * from  dislib.biohazard where case_no=?"
Set selBioHazard_qr1 = rdb.CreateQuery("", str)
selBioHazard_qr1(0) = Trim(MaskEdBox1.FormattedText)
Set selBioHazard_rs1 = selBioHazard_qr1.OpenResultset(1, 3)
If selBioHazard_rs1.EOF Then


Dim BioHazard_qr1 As rdoQuery
str = "insert into dislib.biohazard(CASE_NO,BIOYN,USRID, IPADD ,SAVDATE,SAVTIME) values(?,?,?,?,?,?)"
 Set BioHazard_qr1 = rdb.CreateQuery("", str)
BioHazard_qr1(0) = Trim(MaskEdBox1.FormattedText)
BioHazard_qr1(1) = Left(Trim(cmb_bioHazard.Text), 1)
BioHazard_qr1(2) = logonfrm.usrid
BioHazard_qr1(3) = ipadd
BioHazard_qr1(4) = Format(DATE11, "YYYY-MM-DD")
BioHazard_qr1(5) = Format(time1, "HH:MM:SS")
BioHazard_qr1.Execute
End If
rdb.CommitTrans
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

Private Sub CMD_EMR_Click()
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

    MODFLAG = True

txt_date.Text = Format(DATE11, "dd/mm/yyyy")
txt_time.Text = time1

BTNENABL1 ("NEW")
FLDENABL ("MOD")

If MaskEdBox1.Text = "" Then
    MsgBox "PLEASE FETCH THE VALUES FROM VIEW THAT ARE TO BE MODIFIED"
    cmd_view.Enabled = True
    
End If
cmb_enterby.Enabled = True
cmb_finalby.Enabled = True


UPDRTIMEM_QR(0) = Format(DATE11, "YYYY-MM-DD")
UPDRTIMEM_QR(1) = Format(time1, "HH:MM:SS")
             
UPDRTIMEM_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIMEM_QR.Execute

End Sub

Private Sub cmd_new_Click()



BTNENABL1 ("NEW")

field_clear
TXT_CLEAR
BCCLEAR

NEWFLAG = True
FLDENABL ("ENA")

txt_dept.Text = "MICROBIOLOGY"
txt_modality.Text = "SEROLOGY"
txt_deptcode.Text = "GZZ"
txt_modcode.Text = "SE"

    If DATETIMEFLAG = False Then
    datetime
    DATETIMEFLAG = True
     Else
    datetime
    End If
    txt_date.Text = DATE11
    txt_time.Text = time1

End Sub
Private Sub INSMICRO()

Dim proname_flg1 As Boolean
 
 For i = 1 To MSFlexGrid1.Rows - 1

        If proname_flg1 = False Then
            procode_qr(0) = MSFlexGrid1.TextMatrix(i, 1)
            Set procode_rs = procode_qr.OpenResultset(1, 3)
            'proname_flg = True
        Else
            procode_qr(0) = MSFlexGrid1.TextMatrix(i, 1)
            procode_rs.MoreResults
            procode_rs.Requery
        End If

            insmicrotemp_qr(0) = req_chglen(cmb_reqno.Text)
            insmicrotemp_qr(1) = procode_rs(0)
           
            insmicrotemp_qr(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
            insmicrotemp_qr(3) = Trim(MSFlexGrid1.TextMatrix(i, 3))
            
            insmicrotemp_qr.Execute
            

    Next i

End Sub

Private Sub cmd_print_Click()




Dim VAR_REQNNO As String
    VAR_REQNNO = req_chglen(cmb_reqno.Text)
    
'' ************ ADD HERE TO PRINT REOPRT DIRECT FROM EMR AT SANGRUR BY RAHUL ON DATE 12.12.2016   **************'
    strLink = "http://10.80.1.9/EMRREPORTS/allreports.aspx?reqno=" & VAR_REQNNO & "  ~GUEST"
    Shell "c:\program files\internet explorer\iexplore " & strLink, vbMaximizedFocus
Exit Sub
 '' ********** TILL HERE BY RAHUL ON DATE 12.12.2016   **************'
    
    
   '' to insert records in dislib.MICROTEMP
    deletemicrotemp.Execute
    INSMICRO
    
'*****
    Dim z As Integer
    For z = 0 To 20
        CrystalReport1.Formulas(z) = ""
    Next z
'*****
           
        
If Opt_patient.Value = True Then
    
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\SEROLOGY.RPT"
    CrystalReport1.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
    CrystalReport1.Formulas(1) = "AGE='" & Trim(txt_age.Text) & "'"
    
    '***** Shraddha 11/08/10 *****
'    If VAR_REQNNO = req_chglen(cmb_reqno.List(cmb_reqno.ListCount - 1)) Then
'        CrystalReport1.Formulas(2) = "NewCap='Results of Screen Test'"
'    Else
'        CrystalReport1.Formulas(2) = "NewCap='RESULT'"
'    End If
    '*****************************
    
    GoTo display:
End If

If opt_staff.Value = True Then

    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\SEROLOGYstaff.RPT"
    CrystalReport1.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
        
    GoTo display:

End If

display:
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1

'rtime log
UPDRTIMEP_QR(0) = Format(DATE11, "YYYY-MM-DD")
UPDRTIMEP_QR(1) = Format(time1, "HH:MM:SS")
                
UPDRTIMEP_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIMEP_QR.Execute
'rtimelog


'' to delete records from dislib.MICROTEMP
deletemicrotemp.Execute


Exit Sub
prtmsg: MsgBox Error(Err)
BTNENABL1 ("print")
Exit Sub

End Sub

Private Sub cmd_save_Click()

'On Error GoTo err

txt_date.Text = Format(DATE11, "dd/mm/yyyy")
txt_time.Text = time1

Dim VAR_REQ As String
Dim i As Integer

'For i = 1 To MSFlexGrid1.Rows - 1

 '   Next i

Me.MousePointer = vbHourglass
    
    rdb.BeginTrans
    assign
    


UPDRTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
UPDRTIME_QR(1) = time1
UPDRTIME_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIME_QR.Execute

    rdb.CommitTrans
BTNENABL1 ("SAVE")

Me.MousePointer = vbArrow
Exit Sub

'err:
 '       MsgBox (err.Description)
  '      Exit Sub
End Sub

Private Sub cmd_view_Click()



BTNENABL1 ("VIEW")
FLDENABL ("VIEW")
field_clear
TXT_CLEAR

txt_dept.Text = "MICROBIOLOGY"
txt_modality.Text = "SEROLOGY"
txt_deptcode.Text = "GZZ"
txt_modcode.Text = "SE"

viewflag = True

    If DATETIMEFLAG = False Then
    datetime
    DATETIMEFLAG = True
     Else
    datetime
    End If
    txt_date.Text = DATE11
    txt_time.Text = time1
    
End Sub

Private Sub Command1_Click()
Dim UPDRREQ_QR As rdoQuery
Dim b As String
b = Trim(cmb_reqno.Text)
mod_reqno = Replace(b, "/", "") 'put change value to this var so that it used later on
str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,COMMITFLAG=? WHERE REQNNO=?"
    Set UPDRREQ_QR = rdb.CreateQuery("", str)

UPDRREQ_QR(0) = "T"
UPDRREQ_QR(1) = Null
UPDRREQ_QR(2) = mod_reqno
UPDRREQ_QR.Execute

viewflag = True
cmb_reqno_Click
BTNENABL1 ("VIEW")
FLDENABL ("VIEW")
End Sub

Private Sub Form_Load()

logonfrm.Main
datetime

'Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
'DATE11 = RS("dt1") '**add 24/11/2k
'time1 = Format(RS("TM1"), "hh:mm:ss") '***2/2/99

If DATETIMEFLAG = False Then
    datetime
    DATETIMEFLAG = True
     Else
    datetime
    End If
    txt_date.Text = DATE11
    txt_time.Text = time1


NEWFLAG = False
viewflag = False

txt_dept.Text = "MICROBIOLOGY"
txt_modality.Text = "SEROLOGY"
txt_deptcode.Text = "GZZ"
txt_modcode.Text = "SE"
FLDENABL ("DIS")

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1BIO11", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)

If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = Trim(LOGINUSRNAME)
    cmb_enterby.Text = cmb_enterby.List(0)
Else
    MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
    End
    'cmd_cancel_Click
    Exit Sub
End If

''''''''''''



' Commented and addidng or usrtype >10 for Sangrur by Rahul on 02.12.2016
'str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='GZZ' AND USRTYPE>=8 ORDER BY USERNAME"
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE( DEPT_CODE='GZZ' AND USRTYPE>=8) or USRTYPE>=9 ORDER BY USERNAME"     ' USERTYPE >=9 ADDED BY VARINDER
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
    'cmd_cancel_Click
    Exit Sub
End If


str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BIO", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

'''''''

Set req_qr = rdb.CreateQuery("", "select REQNNO from Rislib.RREQUISITN where CASENO = ? AND REQNSTAT = 'E' AND MODALITYCD = 'SE' AND DEPT_CODE='GZZ' ORDER BY REQNNO DESC ")

Set SELTA101_QR = rdb.CreateQuery("selta101_qr11BC", "select * from PATOOLSLIB.ta101 where case_no =?")

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT=? OR REQNSTAT=?) AND MODALITYCD = 'SE' AND DEPT_CODE='GZZ' ORDER BY REQNDT DESC"
Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QR1", str)

''''''''''

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
Set REFDR_QR = rdb.CreateQuery("REFDR_QR1BC", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1BC", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1BC", str)
 
str = "SELECT * FROM DISLIB.MICROPROCE WHERE REQNO=?"
Set viewmicroproce_qr = rdb.CreateQuery("viewmicroproce", str)

'STR = "SELECT * FROM DISLIB.DMULPROCE WHERE CASENO=? AND REQNO=?"
'Set newdmulproce_qr = rdb.CreateQuery("NEWDMULPROCE", STR)

str = "SELECT * FROM INTERFACE.MINTERFACE WHERE CASENO=? AND REQNO=?"
Set newdmulproce_qr = rdb.CreateQuery("NEWDMULPROCE", str)

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set proname_qr = rdb.CreateQuery("PRO_NAME", str)


'str = "insert into dislib.microproce(reqno,pro_code,observalue,result,FLD1) values (?,?,?,?,?)"
str = "insert into dislib.microproce(reqno,pro_code,observalue,result,FLD1,nrange,fld2) values (?,?,?,?,?,?,?)"
Set insmicroproce_qr = rdb.CreateQuery("insert", str)

str = "INSERT INTO DISLIB.MICROTEMP(REQNO,PRO_CODE,OBSERVALUE,RESULT) VALUES (?,?,?,?)"
Set insmicrotemp_qr = rdb.CreateQuery("MICRO_TEMP", str)

str = "UPDATE RISLIB.RREQUISITN SET ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE CASENO=? AND REQNNO=?"
Set savrreq_qr = rdb.CreateQuery("saverreq", str)

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE TRIM(SER_DET)=?"
Set procode_qr = rdb.CreateQuery("PROCODE_QR", str)

str = "SELECT * FROM INTERFACE.MINTERFACE WHERE CASENO=? AND REQNO=? AND TESTCODE=?"
Set retminterface_qr = rdb.CreateQuery("RETMINTERFACE", str)

'str = "UPDATE DISLIB.MICROPROCE SET OBSERVALUE=?, RESULT=?, FLD1=? WHERE REQNO=? AND PRO_CODE=?"
str = "UPDATE DISLIB.MICROPROCE SET OBSERVALUE=?, RESULT=?, FLD1=?,NRANGE=? ,FLD2=? WHERE REQNO=? AND PRO_CODE=?"
Set updatemicroproce_qr = rdb.CreateQuery("UPDATE", str)

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE TRIM(SER_DET)=?"
Set testcode_qr = rdb.CreateQuery("TESTCODE_QR", str)


'''''''''    RISLIB.RTIMELOG

str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1NHE", str)
    

str = "UPDATE RISLIB.RTIMELOG SET ADPREPDATE=?,ADPREPTIME=? WHERE REQNNO=?"
Set UPDRTIMECADD_QR = rdb.CreateQuery("UPDATETIMECADD", str)
  
str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=? WHERE REQNNO=?"
Set UPDRTIME_QR = rdb.CreateQuery("QWEQWEQWE", str)

str = "UPDATE RISLIB.RTIMELOG SET MODREPDATE=?,MODREPTIME=? WHERE REQNNO=?"
Set UPDRTIMEM_QR = rdb.CreateQuery("UPDRTIMEMODIFY", str)

str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEc", str)
 
str = "UPDATE RISLIB.RTIMELOG SET PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
Set UPDRTIMEP_QR = rdb.CreateQuery("QWEQWEQWEP", str)

str = "UPDATE INTERFACE.MINTERFACE SET TESTRESULT=?,FLD2=? WHERE REQNO=? AND TESTCODE=?"
Set UPDMINTERFACE_QR = rdb.CreateQuery("UPDATEMINTERFACE", str)

'''''''''   DISLIB.DMBCCOMM

str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
Set UPDDMBCCOMMQR = rdb.CreateQuery("UPDDMBCCOMMQRBC", str)

str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
Set SELDMBCCOMMQR = rdb.CreateQuery("SELDMBCCOMMQRBC", str)

str = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES(?,?)"
Set INSDMBCCOMMQR = rdb.CreateQuery("INSDMBCCOMMQRBC", str)
  
  
str = "DELETE FROM DISLIB.MICROTEMP"
Set deletemicrotemp = rdb.CreateQuery("DELETE", str)

  
Set EN = New HMSCrypt

MSFlexGrid1.col = 2
txt_observed.Width = MSFlexGrid1.CellWidth
txt_observed.Height = MSFlexGrid1.CellHeight
txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

MSFlexGrid1.col = 3
txt_result.Width = MSFlexGrid1.CellWidth
txt_result.Height = MSFlexGrid1.CellHeight
txt_result.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_result.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

BTNENABL1 ("INI")

End Sub

Private Sub MaskEdBox1_LostFocus()


Dim CASENO As String
'On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
 Exit Sub
End If

If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
       '' MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented By Monika on 05/12/2022
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
                REQFLAG1 = True
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
        End If
        
        If viewflag = True Then
        
            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = "F"
                SELRREQUISITN_QR(2) = "T"
                
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
               SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = "F"
                SELRREQUISITN_QR(2) = "T"
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
        End If
        
    
        If viewflag = True Then
        
            If SELRREQUISITNFLAG1 = False Then
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = "F"
                SELRREQUISITN_QR(2) = "T"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
               SELRREQUISITNFLAG1 = True
            Else
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = "F"
                SELRREQUISITN_QR(2) = "T"
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
                BTNENABL1 ("INI")
                Exit Sub
            End If
        End If
    End If
cmb_reqno.Enabled = True
'MSG:
'MsgBox Error(Err.Description)

End Sub

Private Sub MSFlexGrid1_Click()

R1 = MSFlexGrid1.row
C1 = MSFlexGrid1.col

If C1 = 2 Then
    txt_observed.Text = MSFlexGrid1.TextMatrix(R1, C1)
    txt_observed.Visible = True
    txt_observed.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
    txt_observed.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
    txt_observed.Width = MSFlexGrid1.CellWidth
    txt_observed.Height = MSFlexGrid1.CellHeight
    txt_observed.SelStart = 0
    txt_observed.SelLength = Len(txt_observed.Text)
    txt_observed.SetFocus
End If

If C1 = 3 Then
    txt_result.Text = MSFlexGrid1.TextMatrix(R1, C1)
    txt_result.Visible = True
    txt_result.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
    txt_result.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
    txt_result.Width = MSFlexGrid1.CellWidth
    txt_result.Height = MSFlexGrid1.CellHeight
    txt_result.SelStart = 0
    txt_result.SelLength = Len(txt_result.Text)
    txt_result.SetFocus
End If

End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
   ' MaskEdBox1.Mask = "??>/#####"
    MaskEdBox1.Enabled = True
    'AMIT
    MaskEdBox1.SetFocus
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then

    lbl_caseno.Caption = "  CC No."
    MaskEdBox1.Mask = "########"
    MaskEdBox1.Enabled = True
    'AMIT
    MaskEdBox1.SetFocus
End If
End Sub



'''''''''''''''''''''''
'Dim CASENO As String
'On Error GoTo MSG:
'If Len(Trim(MaskEdBox1.Text)) = "" Then
' Exit Sub
'End If
'If Opt_patient.Value = True Then
'        CASENO = MaskEdBox1.Text
'        MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
'         If NEWFLAG = True Then
'            If REQFLAG1 = False Then
'                req_qr(0) = MaskEdBox1.FormattedText
'                Set req_rs = req_qr.OpenResultset(1, 3)
'            Else
'                req_qr(0) = MaskEdBox1.FormattedText
'                req_rs.MoreResults
'                req_rs.Requery
'          End If
'          If req_rs.RowCount > 0 Then
'                cmb_reqno.clear
'                While Not req_rs.EOF
'                    cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
'                    req_rs.MoveNext
'                Wend
'          Else
'            cmb_reqno.clear
'          End If
'        End If
'
'        If viewflag = True Then
'            If SELRREQUISITNFLAG1 = False Then
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'                SELRREQUISITNFLAG1 = True
'            Else
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                SELRREQUISITN_RS.MoreResults
'                SELRREQUISITN_RS.Requery
'            End If
'            If SELRREQUISITN_RS.RowCount > 0 Then
'                cmb_reqno.clear
'                While Not SELRREQUISITN_RS.EOF
'                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
'                    SELRREQUISITN_RS.MoveNext
'                Wend
'             Else
'                MsgBox " No  Requisition for this Modality"
'                Exit Sub
'            End If
'
'        End If
'End If
'
'    If opt_staff.Value = True Then
'            CASENO = MaskEdBox1.Text
'        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
'         If NEWFLAG = True Then
'            If REQFLAG1 = False Then
'                req_qr(0) = MaskEdBox1.FormattedText
'                Set req_rs = req_qr.OpenResultset(1, 3)
'            Else
'                req_qr(0) = MaskEdBox1.FormattedText
'                req_rs.MoreResults
'                req_rs.Requery
'          End If
'          If req_rs.RowCount > 0 Then
'                cmb_reqno.clear
'                While Not req_rs.EOF
'                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
'
'                req_rs.MoveNext
'
'                Wend
'          Else
'            cmb_reqno.clear
'
'            End If
'        End If
'
'        If viewflag = True Then
'            If SELRREQUISITNFLAG1 = False Then
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'                SELRREQUISITNFLAG1 = True
'            Else
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                SELRREQUISITN_RS.MoreResults
'                SELRREQUISITN_RS.Requery
'            End If
'            If SELRREQUISITN_RS.RowCount > 0 Then
'                cmb_reqno.clear
'                While Not SELRREQUISITN_RS.EOF
'                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
'                    SELRREQUISITN_RS.MoveNext
'                Wend
'             Else
'                MsgBox " No  Requisition for this Modality"
'                Exit Sub
'            End If
'
'        End If
'End If
' Exit Sub
'MSG:
'MsgBox Error(Err)
'
'

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

Public Sub field_clear()
Dim obj As Object

cmb_finalby.Text = ""
cmb_enterby.Text = ""
cmb_reqno.clear

If Opt_patient.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = ">??/#####"
End If

If opt_staff.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
' MaskEdBox1.Mask = "#######"   ' COMMENTED BY VARINDER
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
TXT_COMMENTS.Text = " "
cmb_bioHazard.Text = ""
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In micro_report.Controls
If TypeOf obj Is TextBox Then
    If obj <> txt_deptcode.Text And obj <> txt_modcode.Text Then
        'If obj <> txt_modcode.Text Then
            obj = ""
        'End If
    End If
End If
Next

End Sub

Public Sub BCCLEAR()

TXT_CLEAR

MSFlexGrid1.Visible = True

MSFlexGrid1.Rows = 1

End Sub

Public Sub FLDENABL(EN As String)

If EN = "DIS" Then
        txt_dept.Enabled = False
        txt_modality.Enabled = False
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
        TXT_COMMENTS.Enabled = False
        cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
    cmb_bioHazard.Enabled = False
End If

If EN = "ENA" Then
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
        TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        cmb_bioHazard.Enabled = True
End If

If EN = "MOD" Then
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
        TXT_COMMENTS.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        cmb_bioHazard.Enabled = True
End If
End Sub

Public Function makekey()
Dim Y, m, d, hr, mn, sc As Integer
Dim myarr(1 To 20) As String
Dim STQ As String
'Dim key
'datetime
'MSgBox (Now)
key = 0

d = Val(Mid(DATE11, 1, 2))
m = Val(Mid(DATE11, 4, 2))
Y = Val(Mid(DATE11, 9, 2))
hr = Val(Mid(time1, 1, 2))
mn = Val(Mid(time1, 4, 2))
sc = Val(Mid(time1, 7, 2))

'd = Val(Mid(Now, 1, 2))
'm = Val(Mid(Now, 4, 2))
'Y = Val(Mid(Now, 9, 2))
'hr = Val(Mid(Now, 12, 2))
'mn = Val(Mid(Now, 15, 2))
'sc = Val(Mid(Now, 18, 2))
Dim MyValue
MyValue = Int((50 * rnd) + 1)  ' Generate random value between 1 and 50.
key = Y + m + d + hr + 2 * sc - mn + MyValue
Select Case key
    Case 1 To 50
        key = key + 25
    Case 1 To 100
        key = key - 30
    Case 101 To 150
        key = key - 60
    Case 151 To 200
        key = key - 80
    Case 201 To 250
        key = key - 130
    Case 251 To 253
        key = key - 140
    Case Is > 253
        key = 90
 End Select
 makekey = key
 
 End Function
Private Function transform(ByRef str As String, ByVal ikey As Integer) As String
Dim m, n, K As Integer
Dim tempstr, AA As String
Dim tempstar
'Dim key
tempstar = str
K = ikey
m = 1
n = Len(Trim(str))
While n > 0
    Mid(tempstar, m, 1) = Chr(Asc(Mid(str, m, 1)) + key + m)
    m = m + 1
    n = n - 1
Wend
transform = tempstar
End Function

Public Function DECRPT(INPSTR As String, KEYVAL As Integer)
Dim ts(1 To 80) As String
Dim OUTSTR As String
Dim m, n As Integer
Dim tempcahr
m = 1
n = Len(Trim(INPSTR))
tempcahr = Trim(INPSTR)

While n > 0
'MsgBox (Chr(Asc(Mid(Text3.Text, m, 1)) - KEY))

ts(m) = Chr(Asc(Mid(tempcahr, m, 1)) - KEYVAL - m)

m = m + 1
n = n - 1
Wend

n = 1
m = m - 1
'Text4.Text = ""
While m > 0
'Text4.Text = Text4.Text + ts(n)
OUTSTR = OUTSTR + ts(n)
n = n + 1
m = m - 1
Wend
DECRPT = OUTSTR
End Function




Private Sub TXT_COMMENTS_GotFocus()

If Trim(TXT_COMMENTS.Text) = "" Then
    TXT_COMMENTS.clear
    TXT_COMMENTS.AddItem "THIS HbsAg  RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 0.05  PATIENT VALUE:       IU/mL. PLEASE REPEAT THE SAMPLE FOR HBsAg TESTING. "
    TXT_COMMENTS.AddItem "THIS HCV RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 1.00  PATIENT VALUE:        . PLEASE REPEAT THE SAMPLE FOR HCV PCR."
    TXT_COMMENTS.AddItem "PLEASE REPEAT THE SAMPLE FOR HIV TESTING AFTER 4 WEEKS.          "
    TXT_COMMENTS.AddItem "The HBsAg value is low and requires a confirmation test. Please send a fresh sample for HBV DNA Quantitive."
    TXT_COMMENTS.AddItem "Please send a Whole blood sample for Quantitive PCR."
    '******** Following two statements are added by Shraddha Pandit ***********
    TXT_COMMENTS.AddItem "Please send EDATA sample for Quantitative PCR."
    TXT_COMMENTS.AddItem "Please send EDATA sample for Qualitative PCR."
    TXT_COMMENTS.AddItem "HBsAg TEST HAS BEEN PERFORMED AT NICHOLAS PIRAMAL LABORATORIES." 'Shraddha 03/08/10
     '*****************Following statements are added by varinder
    TXT_COMMENTS.AddItem "Test Performed by Chemiluminescent Immunoassay."
     TXT_COMMENTS.AddItem "HIV testing is performed by two additional methods - enzyme linked fluorescent assay and Immunochromatography."
Else
    TXT_COMMENTS.clear
    TXT_COMMENTS.AddItem "THIS HbsAg  RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 0.05  PATIENT VALUE:       IU/mL. PLEASE REPEAT THE SAMPLE FOR HBsAg TESTING. "
    TXT_COMMENTS.AddItem "THIS HCV RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 1.00 PATIENT VALUE:        . PLEASE REPEAT THE SAMPLE FOR HCV PCR."
    TXT_COMMENTS.AddItem "PLEASE REPEAT THE SAMPLE FOR HIV TESTING AFTER 4 WEEKS.          "
    TXT_COMMENTS.AddItem "The HBsAg value is low and requires a confirmation test. Please send a fresh sample for HBV DNA Quantitive."
    TXT_COMMENTS.AddItem "Please send a Whole blood sample for Quantitive PCR."
    '******** Following two statements are added by Shraddha Pandit ***********
    TXT_COMMENTS.AddItem "Please send EDATA sample for Quantitative PCR."
    TXT_COMMENTS.AddItem "Please send EDATA sample for Qualitative PCR."
    TXT_COMMENTS.AddItem "SPECIMENS WITH CONCENTRATION VALUES <6.0 AU/mL ARE CONSIDERED NONREACTIVE FOR IgG ANTIBODIES TO CMV."
    TXT_COMMENTS.AddItem "SPECIMENS WITH CONCENTRATION VALUES >=6.0 AU/mL ARE CONSIDERED REACTIVE FOR IgG ANTIBODIES TO CMV. THIS INDICATES PAST OR ACUTE INFECTION."
    TXT_COMMENTS.AddItem "HBsAg TEST HAS BEEN PERFORMED AT NICHOLAS PIRAMAL LABORATORIES." 'Shraddha 03/08/10
     '*****************Following statements are added by varinder
    TXT_COMMENTS.AddItem "Test Performed by Chemiluminescent Immunoassay."
     TXT_COMMENTS.AddItem "HIV testing is performed by two additional methods - enzyme linked fluorescent assay and Immunochromatography."
End If

End Sub
Private Sub txt_observed_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text

End If

End Sub


Private Sub txt_result_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = txt_result.Text
    End If

End Sub
