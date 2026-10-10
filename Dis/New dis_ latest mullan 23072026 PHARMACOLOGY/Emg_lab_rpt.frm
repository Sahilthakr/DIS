VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form emg_lab_rpt 
   Caption         =   "Emergency Laboratory Report"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   120
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame common 
      Height          =   3135
      Left            =   0
      TabIndex        =   18
      Top             =   960
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
         Height          =   375
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9960
         TabIndex        =   62
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9960
         TabIndex        =   61
         Top             =   2520
         Width           =   270
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9960
         TabIndex        =   60
         Top             =   2880
         Width           =   270
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   51
         Top             =   720
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
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1095
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
            TabIndex        =   52
            Top             =   600
            Width           =   1095
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   2160
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
      Begin VB.Label LBL_CBG 
         Caption         =   "CAPILLARY BLOOD GAS REPORT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   3120
         TabIndex        =   78
         Top             =   2640
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label LBL_VBG 
         Caption         =   "VENOUS BLOOD GAS REPORT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   3120
         TabIndex        =   76
         Top             =   2640
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label4 
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
         TabIndex        =   65
         Top             =   2160
         Width           =   945
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
         Left            =   10440
         TabIndex        =   64
         Top             =   2520
         Width           =   900
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
         Left            =   10440
         TabIndex        =   63
         Top             =   2880
         Width           =   1140
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label LBL_ABG 
         Caption         =   "ARTERIAL BLOOD GAS REPORT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   3120
         TabIndex        =   77
         Top             =   2640
         Visible         =   0   'False
         Width           =   3975
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10440
      TabIndex        =   17
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10440
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame COMMAND 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton cmd_addendum 
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
         Height          =   390
         Left            =   10410
         TabIndex        =   73
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   360
         MouseIcon       =   "Emg_lab_rpt.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   1616
         MouseIcon       =   "Emg_lab_rpt.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   2872
         MouseIcon       =   "Emg_lab_rpt.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   4128
         MouseIcon       =   "Emg_lab_rpt.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   5384
         MouseIcon       =   "Emg_lab_rpt.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   7896
         MouseIcon       =   "Emg_lab_rpt.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   9152
         MouseIcon       =   "Emg_lab_rpt.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   270
         Width           =   1080
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
         Height          =   390
         Left            =   6640
         MouseIcon       =   "Emg_lab_rpt.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame REPORT 
      Height          =   3735
      Left            =   0
      TabIndex        =   10
      Top             =   4095
      Width           =   12855
      Begin VB.ComboBox CMB_MACHINE 
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
         ItemData        =   "Emg_lab_rpt.frx":1850
         Left            =   10200
         List            =   "Emg_lab_rpt.frx":1860
         TabIndex        =   79
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   9135
         TabIndex        =   72
         Top             =   3300
         Width           =   2295
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
         Height          =   300
         Left            =   7140
         TabIndex        =   70
         Top             =   2745
         Width           =   4275
      End
      Begin VB.ComboBox TXT_COMMENTS 
         Height          =   315
         ItemData        =   "Emg_lab_rpt.frx":187F
         Left            =   945
         List            =   "Emg_lab_rpt.frx":1881
         TabIndex        =   66
         Top             =   2790
         Width           =   4935
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
         Left            =   5370
         TabIndex        =   59
         Top             =   3270
         Width           =   2400
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
         Left            =   1200
         TabIndex        =   57
         Top             =   3240
         Width           =   2550
      End
      Begin VB.TextBox txt_comments1 
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
         Left            =   945
         TabIndex        =   54
         Top             =   2790
         Width           =   4920
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
         Left            =   4440
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   192
         FormatString    =   $"Emg_lab_rpt.frx":1883
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
      Begin VB.Label Label6 
         Caption         =   "Addendum BY"
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
         Left            =   7830
         TabIndex        =   71
         Top             =   3315
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   6120
         TabIndex        =   69
         Top             =   2805
         Width           =   915
      End
      Begin VB.Label lbl_finalby 
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
         Left            =   3795
         TabIndex        =   58
         Top             =   3345
         Width           =   1575
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Analysed by"
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
         TabIndex        =   56
         Top             =   3360
         Width           =   1035
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   55
         Top             =   2895
         Width           =   750
      End
   End
   Begin VB.Label LBL_CAPTIONS 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   80
      Top             =   120
      Width           =   255
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
      Left            =   1320
      TabIndex        =   68
      Top             =   0
      Width           =   8655
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
      TabIndex        =   67
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Location"
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
      Height          =   195
      Left            =   6600
      TabIndex        =   50
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label lbl_biochemlab 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - EMERGENCY LABORATORY SERVICES "
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
      Left            =   2280
      TabIndex        =   49
      Top             =   600
      Width           =   7050
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
      Top             =   720
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
      TabIndex        =   47
      Top             =   360
      Width           =   420
   End
   Begin VB.Label lbl_rwd 
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
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "emg_lab_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ADDED BY MILIND DT. 03/03/06
Dim userid_qr As rdoQuery
Dim userid_rs As rdoResultset
Dim userdesc_qr As rdoQuery
Dim userdesc_rs As rdoResultset
Dim adn_QR As rdoQuery
Dim adn_rs As rdoResultset

'upto here

Dim SELBG_QR As rdoQuery
Dim SELBG_RS As rdoResultset

Dim str As String
'Dim FSFLAG As Boolean
Dim sel_deptqr As rdoQuery
'Dim UPDSPREPORT_QR As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SEL_DEPTFLAG As Boolean
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim SEL_MODALITYFLAG As Boolean
'Dim sel_casno As rdoQuery
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim REQFLAG As Boolean
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim SELTA101FLAG As Boolean
'Dim repdr_qr As rdoQuery
'Dim repdr_rs As rdoResultset
'Dim REPDRFLAG As Boolean
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim REFDRFLAG As Boolean
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset
Dim selma104flag As Boolean
'Dim SELPATHDREQUISITN_QR As rdoQuery
'Dim SELPATHDREQUISITN_RS As rdoResultset
'Dim SELPATHDREQUISITNFLAG As Boolean
'Dim selma302_qr As rdoQuery
'Dim selma302_rs As rdoResultset
'Dim SELMA302FLAG As Boolean

'Dim SELDSPREPORTFLAG As Boolean
'Dim selta1012_qr As rdoQuery
'Dim selta1012_rs As rdoResultset
'Dim SELTA1012FLAG As Boolean


'Dim SELEMPLIB_QR As rdoQuery
'Dim SELEMPLIB_RS As rdoResultset

Dim seldissecentr1_qr As rdoQuery
Dim seldissecentr1_rs As rdoResultset
Dim seldissecentrflag1 As Boolean

Dim seldissecfin1_qr As rdoQuery
Dim seldissecfin1_rs As rdoResultset
Dim seldissecfinflag1 As Boolean

Dim i As Integer

Dim SELDBCOSMOQR As rdoQuery
Dim SELDBCOSMORS As rdoResultset
Dim SELDBCOSMOFLAG As Boolean
Dim RNAGE As String
Dim SS As String
Dim OSMOLV As Integer
Dim OSMOHV As Integer

Dim INSDPRO As rdoQuery

Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
'Dim selempbasinfo1_qr As rdoQuery
'Dim selempbasinfo1_rs As rdoResultset
'Dim SELEMPBASINFO1FLAG As Boolean
Dim updrrequisitn_qr As rdoQuery

Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer


Dim SELDREQSTAFFD1QR1 As rdoQuery
Dim SELDREQSTAFFD1RS1 As rdoResultset
Dim SELDREQSTAFFD1FLAG1 As Boolean

Dim DMULPROCE_QR As rdoQuery
Dim DMULPROCE_RS As rdoResultset
Dim DMULPROCEFLAG As Boolean

Dim DMPROCEDUR_QR As rdoQuery
Dim DMPROCEDUR_RS As rdoResultset
Dim DMPROCEDURFLAG As Boolean

Dim UPDDMPROCEDUR_QR As rdoQuery
Dim UPDDMPROCEDUR_RS As rdoResultset
Dim UPDDMPROCEDURFLAG As Boolean

Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
Dim SELMA101FLAG As Boolean

Dim SELMA101_QR1 As rdoQuery
Dim SELMA101_RS1 As rdoResultset
Dim SELMA101FLAG1 As Boolean

Dim MODIFYFLAG As Boolean
Dim SELDTEMPLETQR As rdoQuery
Dim SELDTEMPLETRS As rdoResultset
Dim SELDTEMPLETFLAG As Boolean

Dim SELDMA101QR As rdoQuery
Dim SELDMA101RS As rdoResultset
Dim SELDMA101FLAG As Boolean

Dim SELDBPROSUBQR As rdoQuery
Dim SELDBPROSUBRS As rdoResultset
Dim SELDBPROSUBFLAG As Boolean
Dim INSDMBCCOMMQR As rdoQuery
Dim UPDDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMRS As rdoResultset
Dim SELDMBCCOMMFLAG As Boolean
Dim SELELPROSUBQR As rdoQuery
Dim SELELPROSUBRS As rdoResultset
Dim SELELPROSUBFLAG As Boolean
Dim updrrequisitn1_qr As rdoQuery

Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset
Dim selpathdrequisitnrsflag As Boolean

''''******** interface.dinterface
Dim seldinterface_qr As rdoQuery
Dim seldinterface_rs As rdoResultset
Dim seldinterfaceflag As Boolean


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
Dim REPWTRN_QR As rdoQuery
Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

'addendum
Dim insertaddendum_qr As rdoQuery
Dim selectaddendum_qr As rdoQuery
Dim selectaddendum_rs As rdoResultset
Dim updateaddendum_qr As rdoQuery
Dim ADDENDUMFALG As Boolean

'RTIMELOG
Dim UPDRTIMECADD_QR As rdoQuery
Dim UPDRTIME_QR As rdoQuery
Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEP_QR As rdoQuery
Dim UPDRTIMEM_QR As rdoQuery

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset
Dim VAR_REQNNO As String

'***** Shraddha 16/12/10 *****
Dim SEL_DINT_QR As rdoQuery
Dim SEL_DINT_RS As rdoResultset
Dim SEL_DTEMPLET1_QR As rdoQuery
Dim SEL_DTEMPLET1_RS As rdoResultset

'*****************************

''****** SECURITY ***********
Public Sub SELECTDINTERFACE()
i = 1

                                If seldinterfaceflag = False Then
                                    seldinterface_qr(0) = VAR_REQNNO
                                  '  seldinterface_qr(1) = Trim(SELELPROSUBRS("PROCODESUB"))
                                    Set seldinterface_rs = seldinterface_qr.OpenResultset(1, 3)
                                    seldinterfaceflag = True
                                Else
                                    seldinterface_qr(0) = VAR_REQNNO
                                   ' seldinterface_qr(1) = Trim(SELELPROSUBRS("PROCODESUB"))
                                    seldinterface_rs.MoreResults
                                    seldinterface_rs.Requery
                                End If

                            
                            If seldinterface_rs.RowCount > 0 Then
                            
                            If Trim(seldinterface_rs("MACHINENO")) = "" Then 'Shraddha 17/12/10 *****
                            
                            '***** 17/12/10
                            While Not seldinterface_rs.EOF

                                If SELMA101FLAG = False Then
                                    SELMA101_QR(0) = Trim(seldinterface_rs("TESTCODE"))
                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                    SELMA101FLAG = True
                                Else
                                    SELMA101_QR(0) = Trim(seldinterface_rs("TESTCODE"))
                                    SELMA101_RS.MoreResults
                                    SELMA101_RS.Requery
                                End If

                                'If SELDTEMPLETRS.RowCount > 0 Then
                                    '***** 13/01/11 Commented : Shraddha  *****
'                                    If seldinterface_rs("TESTCODE") = "F998" Then
'
'                                        OSMOAGE
'
'                                        If SELDBCOSMOFLAG = False Then
'                                            SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
'                                            SELDBCOSMOQR(1) = RNAGE
'                                            SELDBCOSMOQR(2) = SS
'                                            Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
'                                            SELDBCOSMOFLAG = True
'                                        Else
'                                            SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
'                                            SELDBCOSMOQR(1) = RNAGE
'                                            SELDBCOSMOQR(2) = SS
'                                            SELDBCOSMORS.MoreResults
'                                            SELDBCOSMORS.Requery
'                                        End If
'
'                                        RNAGE = SELDBCOSMORS("RANGE")
'
'                                        OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
'                                        OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
'
'
'                                        If SELMA101_RS.RowCount > 0 Then
'
'
'                                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
'                                                MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                            Else
'                                                SELDTEMPLETRS.MoveNext
'                                                MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                            End If
'                                            I = I + 1
'                                        End If
'
'
'                                    Else
                                    
                                    '***** Shraddha 07/10/2011 *****
'                                    SEL_DTEMPLET1_QR(0) = Trim(seldinterface_rs("TESTCODE"))
'                                    SEL_DTEMPLET1_QR(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5)
'                                    Set SEL_DTEMPLET1_RS = SEL_DTEMPLET1_QR.OpenResultset(1, 3)
                                    '*******************************
                                    
                                        If SELMA101_RS.RowCount > 0 Then
                                            'If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab '& Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                            'Else
                                             '   SELDTEMPLETRS.MoveNext
                                             '   MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                            'End If
                                        End If
                                        
                                    'End If 'Commented 13/01/11

                                'Else
                                 '   MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                'End If
                                i = i + 1
                                
                                seldinterface_rs.MoveNext
                                Wend
                            '***** 17/12/10
                            
                            'Commented by Shraddha on 17/12/10
'                                 While Not seldinterface_rs.EOF
'                                        If SELDTEMPLETFLAG = False Then
'                                            SELDTEMPLETQR(0) = Trim(seldinterface_rs("testcode"))
'                                            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                                            SELDTEMPLETFLAG = True
'                                        Else
'                                            SELDTEMPLETQR(0) = Trim(seldinterface_rs("TESTCODE"))
'                                            SELDTEMPLETRS.MoreResults
'                                            SELDTEMPLETRS.Requery
'                                        End If
'
'                                        If SELMA101FLAG = False Then
'                                            SELMA101_QR(0) = Trim(seldinterface_rs("TESTCODE"))
'                                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                                            SELMA101FLAG = True
'                                        Else
'                                            SELMA101_QR(0) = Trim(seldinterface_rs("TESTCODE"))
'                                            SELMA101_RS.MoreResults
'                                            SELMA101_RS.Requery
'                                        End If
'    ''''''''KANCHAN
'
'                                        If SELDTEMPLETRS.RowCount > 0 Then
'
'                                            If seldinterface_rs("TESTCODE") = "F998" Then
'
'                                                OSMOAGE
'
'                                                If SELDBCOSMOFLAG = False Then
'                                                    SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
'                                                    SELDBCOSMOQR(1) = RNAGE
'                                                    SELDBCOSMOQR(2) = SS
'                                                    Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
'                                                    SELDBCOSMOFLAG = True
'                                                Else
'                                                    SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
'                                                    SELDBCOSMOQR(1) = RNAGE
'                                                    SELDBCOSMOQR(2) = SS
'                                                    SELDBCOSMORS.MoreResults
'                                                    SELDBCOSMORS.Requery
'                                                End If
'
'                                                RNAGE = SELDBCOSMORS("RANGE")
'
'                                                OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
'                                                OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
'
'
'                                                If SELMA101_RS.RowCount > 0 Then
'                                                    If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
'                                                        MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                                    Else
'                                                        SELDTEMPLETRS.MoveNext
'                                                        MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                                    End If
'                                                    I = I + 1
'                                                End If
'                                            Else
'                                                If SELMA101_RS.RowCount > 0 Then
'                                                    If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
'                                                        MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                                    Else
'                                                        SELDTEMPLETRS.MoveNext
'                                                        MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                                    End If
'                                                End If
'                                            End If
'
'                                        Else
'                                                'SELDTEMPLETRS.MoveNext
'                                                MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                        End If
'                                             I = I + 1
'                                        'End If
'                                'End If
'                                seldinterface_rs.MoveNext
'                                Wend

                            '***** Shraddha 17/12/10 *****
                            Else
                                While Not seldinterface_rs.EOF
                                    If SELDTEMPLETFLAG = False Then
                                        SELDTEMPLETQR(0) = Trim(seldinterface_rs("testcode"))
                                        SELDTEMPLETQR(1) = Trim(seldinterface_rs("machineno"))
                                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                                        SELDTEMPLETFLAG = True
                                    Else
                                        SELDTEMPLETQR(0) = Trim(seldinterface_rs("TESTCODE"))
                                        SELDTEMPLETQR(1) = Trim(seldinterface_rs("machineno"))
                                        SELDTEMPLETRS.MoreResults
                                        SELDTEMPLETRS.Requery
                                    End If
                                        
                                    If SELMA101FLAG = False Then
                                        SELMA101_QR(0) = Trim(seldinterface_rs("TESTCODE"))
                                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                        SELMA101FLAG = True
                                    Else
                                        SELMA101_QR(0) = Trim(seldinterface_rs("TESTCODE"))
                                        SELMA101_RS.MoreResults
                                        SELMA101_RS.Requery
                                    End If

    
                                    If SELDTEMPLETRS.RowCount > 0 Then
        
                                        If seldinterface_rs("TESTCODE") = "F998" Then
                
                                            OSMOAGE
        
                                            If SELDBCOSMOFLAG = False Then
                                                SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
                                                SELDBCOSMOQR(1) = RNAGE
                                                SELDBCOSMOQR(2) = SS
                                                Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
                                                SELDBCOSMOFLAG = True
                                            Else
                                                SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
                                                SELDBCOSMOQR(1) = RNAGE
                                                SELDBCOSMOQR(2) = SS
                                                SELDBCOSMORS.MoreResults
                                                SELDBCOSMORS.Requery
                                            End If
    
                                            RNAGE = SELDBCOSMORS("RANGE")
        
                                            OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
                                            OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)

    
                                            If SELMA101_RS.RowCount > 0 Then
                                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
                                                    'MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
                                                    'Commented and added by Shraddha on 17/12/10
                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE & vbTab & vbTab & Trim(seldinterface_rs("machineno"))
                                                Else
                                                    SELDTEMPLETRS.MoveNext
                                                    'MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
                                                    'Commented and added by Shraddha on 17/12/10
                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE & vbTab & vbTab & Trim(seldinterface_rs("machineno"))
                                                End If
                                                i = i + 1
                                            End If
                                        Else
                                            If SELMA101_RS.RowCount > 0 Then
                                                If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
                                                    'Commented and added by Shraddha on 17/12/10
                                                    'MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & Trim(seldinterface_rs("machineno"))
                                                    
                                                Else
                                                    
                                                    'MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                                    'Commented and added by Shraddha on 17/12/10
                                                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & Trim(seldinterface_rs("machineno"))
                                                    SELDTEMPLETRS.MoveNext
                                                End If
                                            End If
                                        End If
                                                    
                                    Else
                                        'MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                        'Commented and added by Shraddha on 17/12/10
                                        If Not SELDTEMPLETRS.EOF Then
                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & Trim(seldinterface_rs("machineno"))
                                        End If
                                        
                                    End If
                                    
                                    i = i + 1
                                       
                                    seldinterface_rs.MoveNext
                                Wend
                            End If
                            '*****************************
                        End If

End Sub


Public Sub btnenabl(cap As String)
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
'    CMD_NEW.Enabled = True
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

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_dept.Enabled = True
    MaskEdBox1.Enabled = True
    MSFlexGrid1.Enabled = True
    txt_comments.Enabled = True
    cmb_enterby.Enabled = True
    cmb_finalby.Enabled = True
End If
If str1 = "DIS" Then
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_dept.Enabled = False
    MaskEdBox1.Enabled = False
    MSFlexGrid1.Enabled = False
    txt_comments.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
End If
If str1 = "VIEW" Then
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_dept.Enabled = True
    MaskEdBox1.Enabled = True
    MSFlexGrid1.Enabled = False
    txt_comments.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
End If
If str1 = "MODIFY" Then
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_dept.Enabled = False
    MaskEdBox1.Enabled = False
    txt_comments.Enabled = True
    cmb_enterby.Enabled = True
    cmb_finalby.Enabled = True
    MSFlexGrid1.Enabled = True
End If
End Sub

Private Sub cmb_machine_Click()
Dim num As Integer
Dim sc As String

cmb_machine.Visible = False
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = Trim(cmb_machine.Text)
num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)

     If SELDTEMPLETFLAG1 = False Then
         SELDTEMPLETQR(0) = sc
         SELDTEMPLETQR(1) = Trim(cmb_machine.Text)
         Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
         SELDTEMPLETFLAG1 = True
     Else
         SELDTEMPLETQR(0) = sc
         SELDTEMPLETQR(1) = Trim(cmb_machine.Text)
         SELDTEMPLETRS.MoreResults
         SELDTEMPLETRS.Requery
     End If
     
     If SELDTEMPLETRS.RowCount > 0 Then
     
     SEX = Trim(txt_sex.Text)
                    If Trim(sc) = "F957" Then
                       If Trim(txt_age.Text) <= 15 Then
                          SEX = "C"
                       Else
                          SEX = "A"
                       End If
                    End If
     
        If SELDTEMPLETRS("SEX") = Trim(SEX) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = RNAGE1
        Else
        
            If SELDTEMPLETRS.RowCount > 1 Then 'SHWETA 10-08-2011
            SELDTEMPLETRS.MoveNext
            End If
            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = RNAGE1
        End If
     Else
           'MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = ""
           
     'NEW REFERANCE RANGE DATA
                Dim RS_VWREFRNG As rdoResultset
                 Set RS_VWREFRNG = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & sc & "' ", 1, 1)
        
                
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
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & sc & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range

                        If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Or Left(Trim(SELDBCALKPHSGGTRS!range), 1) = ">" Then
                        Else
                        ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        End If
                    End If
                End If
       
             ' RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = RNAGE1
      
           
     End If
     
     '***** Shraddha 17/12/10 *****
     If MSFlexGrid1.row = 1 Then
     If MSFlexGrid1.TextMatrix(1, 5) <> "" Then
        If seldinterfaceflag = False Then
           seldinterface_qr(0) = VAR_REQNNO
           Set seldinterface_rs = seldinterface_qr.OpenResultset(1, 3)
           seldinterfaceflag = True
        Else
           seldinterface_qr(0) = VAR_REQNNO
           seldinterface_rs.MoreResults
           seldinterface_rs.Requery
        End If
     
            i = 1
            While Not seldinterface_rs.EOF
            If MSFlexGrid1.Rows <= i Then 'SHWETA 17-OCT-2012
            GoTo XX
           End If  '
                MSFlexGrid1.TextMatrix(i, 5) = MSFlexGrid1.TextMatrix(1, 5)
              
                '*****
                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = Trim(MSFlexGrid1.TextMatrix(1, 5))
                num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
                sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
                
                     If SELDTEMPLETFLAG1 = False Then
                         SELDTEMPLETQR(0) = sc
                         SELDTEMPLETQR(1) = Trim(MSFlexGrid1.TextMatrix(1, 5))
                         Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                         SELDTEMPLETFLAG1 = True
                     Else
                         SELDTEMPLETQR(0) = sc
                         SELDTEMPLETQR(1) = Trim(MSFlexGrid1.TextMatrix(1, 5))
                         SELDTEMPLETRS.MoreResults
                         SELDTEMPLETRS.Requery
                     End If
                     
                     If SELDTEMPLETRS.RowCount > 0 Then
                     
                     SEX = Trim(txt_sex.Text)
                    If Trim(sc) = "F957" Then
                       If Trim(txt_age.Text) <= 15 Then
                          SEX = "C"
                       Else
                          SEX = "A"
                       End If
                    End If
     
                     
                        If SELDTEMPLETRS("SEX") = Trim(SEX) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
                            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                            MSFlexGrid1.TextMatrix(i, 3) = RNAGE1
                        Else
                            If sc = "F987" Or Trim(sc) = "F957" Then
                            SELDTEMPLETRS.MoveNext
                                                       
                            End If
                            RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                            MSFlexGrid1.TextMatrix(i, 3) = RNAGE1
                            SELDTEMPLETRS.MoveNext
                        End If
                     Else
                           MSFlexGrid1.TextMatrix(i, 3) = ""
                           
                           
                            'NEW REFERANCE RANGE DATA 'ADDED ON 23-03-2015
                Dim RS_VWREFRNG1 As rdoResultset
                 Set RS_VWREFRNG1 = rdb.OpenResultset("SELECT * FROM DISLIb.VW_BIOREFR WHERE PROCODE='" & sc & "' ", 1, 1)
        
                
                If Not RS_VWREFRNG1.EOF Then
                Dim agecal1 As Double
                agecal1 = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                txt_age.Text = txt_age ' DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                If Trim(txt_age.Text) <= 1 Then
                Dim age31 As Double
                 age31 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                age31 = "." & age3 '(age Mod 12)
                If age31 = ".0" Then
                    age31 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                    age31 = ".0" & (age3 Mod 30)
                End If
                txt_age.Text = Trim(age31)
            End If
                   ' ALKAGEO 'FOR F876
                    Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & sc & "' AND (AGE_1)<" & (txt_age.Text) & " AND AGE_2>=" & (txt_age.Text) & " and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                   ' Set SELDBCALKPHSGGTRS = rdb.OpenResultset("SELECT * FROM DISLIB.DBCALKPHS2 WHERE PROCODE='" & sc & "' and sex = '" & Trim(txt_sex.Text) & "' ", 1, 1)
                    If Not SELDBCALKPHSGGTRS.EOF Then
                    RNAGE1 = SELDBCALKPHSGGTRS!range

                        If Left(Trim(SELDBCALKPHSGGTRS!range), 1) = "<" Or Left(Trim(SELDBCALKPHSGGTRS!range), 1) = ">" Then
                        Else
                        ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
                        End If
                        ' RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                    End If
                End If
       
             
             MSFlexGrid1.TextMatrix(i, 3) = RNAGE1
                           
                     End If
                '*****
                seldinterface_rs.MoveNext
                i = i + 1
               
            Wend
XX:
     End If
     End If
     '*****************************
     
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
    MsgBox Err.Description
End Sub

Private Sub cmb_modality_GotFocus()

On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "EMERGENCY LAB$EL" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
End If
Exit Sub

MSG:
MsgBox Error(Err)

End Sub

Private Sub cmb_reqno_Click()

'On Error GoTo MSG:
Dim ST, SEX As String
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
                  lbl_rwd.Visible = False
                  LBL_CAPTIONS.Visible = False
               ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                  lbl_rwd.Visible = True
                  lblicu.Visible = False
                  LBL_CAPTIONS.Visible = False
            '***** Shraddha 05/12/2011 *****
                ElseIf Trim(SELICU_RS("ICDS2")) = "PR" Then
                      lbl_rwd.Visible = False
                      lblicu.Visible = False
                      LBL_CAPTIONS.Caption = "PROTOCOL"
                      LBL_CAPTIONS.Visible = True
                ElseIf Trim(SELICU_RS("ICDS2")) = "CM" Then
                      lbl_rwd.Visible = False
                      lblicu.Visible = False
                      LBL_CAPTIONS.Caption = "CASUALTY MAIN"
                      LBL_CAPTIONS.Visible = True
                ElseIf Trim(SELICU_RS("ICDS2")) = "CG" Then
                      lbl_rwd.Visible = False
                      lblicu.Visible = False
                      LBL_CAPTIONS.Caption = "CASUALTY GJB"
                      LBL_CAPTIONS.Visible = True
               End If
            'End If
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
            
'            If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
'                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)
'                   Dim N1 As Integer
'                   N1 = YEAR1 / 12
'                   txt_age.Text = N1
'               End If
'
'            If Trim(txt_age.Text) = 0 Then
'                  Dim age As String
'                  age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                  age = "." & (age Mod 12)
'                     If age = ".0" Then
'                        age = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        age = ".0" & (age) ' Mod 30)
'                     End If
'                  txt_age.Text = Trim(age)
'            End If
            
            txt_pridoc.Text = SELTA101_RS("ud_name")
            txt_category.Text = SELTA101_RS("pat_catg")
        
            If REFDRFLAG = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
            If REFDR_RS.RowCount > 0 Then
                txt_refdoc.Text = REFDR_RS("REFBYDOC")
            End If
            If selpathdrequisitnrsflag = False Then
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                selpathdrequisitnrsflag = True
            Else
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = Trim(selpathdrequisitn_rs("pathno"))
            End If

'            If selma104flag = False Then
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'                selma104flag = True
'            Else
'                selma104_qr(0) = MaskEdBox1.FormattedText
'                selma104_rs.MoreResults
'                selma104_rs.Requery
'            End If
      
            If selpathdrequisitn_rs.RowCount > 0 Then
                If REFDR_RS("patnstat") = "IP" Then
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
        
            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F995" Then
                LBL_VBG.Visible = True
            Else
                LBL_VBG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F901" Then
                LBL_ABG.Visible = True
            Else
                LBL_ABG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F902" Then
                LBL_CBG.Visible = True
            Else
                LBL_CBG.Visible = False
            End If
            
            SELECTDINTERFACE
'*********** kANCHAN
'            If DMULPROCE_RS.RowCount > 0 Then
'                i = 1
'                MSFlexGrid1.Rows = 1
'                '******
'                While Not DMULPROCE_RS.EOF
'                'JUSTIN 31-08-2004
'                    'If DMULPROCE_RS("PRO_CODE") = "F981" Or DMULPROCE_RS("PRO_CODE") = "F875" Or DMULPROCE_RS("PRO_CODE") = "F901" Or DMULPROCE_RS("PRO_CODE") = "F903" Or DMULPROCE_RS("PRO_CODE") = "F908" Then
'                    If DMULPROCE_RS("PRO_CODE") = "F981" Or DMULPROCE_RS("PRO_CODE") = "F875" Or DMULPROCE_RS("PRO_CODE") = "F903" Or DMULPROCE_RS("PRO_CODE") = "F908" Or DMULPROCE_RS("PRO_CODE") = "F949" Or DMULPROCE_RS("PRO_CODE") = "F901" Or DMULPROCE_RS("PRO_CODE") = "F902" Or DMULPROCE_RS("PRO_CODE") = "F995" Or DMULPROCE_RS("PRO_CODE") = "F022" Then
'                    ''   F901 IS EXCLUDED FROM THE IF AS IT REFERS TO ARTERIAL BLOOD GASES WHICH IS NOT A GROUP SERVICS
'                        If SELELPROSUBFLAG = False Then
'                            SELELPROSUBQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            Set SELELPROSUBRS = SELELPROSUBQR.OpenResultset(1, 3)
'                            SELELPROSUBFLAG = True
'                        Else
'                            SELELPROSUBQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            SELELPROSUBRS.MoreResults
'                            SELELPROSUBRS.Requery
'                        End If
'                        If SELELPROSUBRS.RowCount > 0 Then
'                            While Not SELELPROSUBRS.EOF
'                                If SELMA101FLAG = False Then
'                                    SELMA101_QR(0) = Trim(SELELPROSUBRS("PROCODESUB"))
'                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                                    SELMA101FLAG = True
'                                Else
'                                    SELMA101_QR(0) = Trim(SELELPROSUBRS("PROCODESUB"))
'                                    SELMA101_RS.MoreResults
'                                    SELMA101_RS.Requery
'                                End If
'                                 If SELDTEMPLETFLAG = False Then
'                                    SELDTEMPLETQR(0) = Trim(SELELPROSUBRS("PROCODESUB"))
'                                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                                    SELDTEMPLETFLAG = True
'                                Else
'                                    SELDTEMPLETQR(0) = SELELPROSUBRS("PROCODESUB")
'                                    SELDTEMPLETRS.MoreResults
'                                    SELDTEMPLETRS.Requery
'                                End If
'
'                                '''''''''' KANCHAN & PRATS 08/01/2008
'                                ''*** CODE TO RETRIVE OBSERVED VALUE FROM INTERFACE.DINTERFACE
'
'                                If seldinterfaceflag = False Then
'                                    seldinterface_qr(0) = VAR_REQNNO
'                                    seldinterface_qr(1) = Trim(SELELPROSUBRS("PROCODESUB"))
'                                    Set seldinterface_rs = seldinterface_qr.OpenResultset(1, 3)
'                                    seldinterfaceflag = True
'                                Else
'                                    seldinterface_qr(0) = VAR_REQNNO
'                                    seldinterface_qr(1) = Trim(SELELPROSUBRS("PROCODESUB"))
'                                    seldinterface_rs.MoreResults
'                                    seldinterface_rs.Requery
'                                End If
'
'
'                                'for alkaline phosphate changes in range according to age
'                                SEX = Trim(txt_sex.Text)
'                                If Trim(SELELPROSUBRS("PROCODESUB")) = "F957" Then
'                                   If Trim(txt_age.Text) <= 15 Then
'                                      SEX = "C"
'                                   Else
'                                      SEX = "A"
'                                   End If
'                                End If
'                                If SELDTEMPLETRS.RowCount > 0 Then
'                                    If SELMA101_RS.RowCount > 0 Then
'                                        If SELDTEMPLETRS("SEX") = SEX Or Trim(SELDTEMPLETRS("SEX")) = "" Then
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                        Else
'                                            SELDTEMPLETRS.MoveNext
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                        End If
'                                        i = i + 1
'                                    End If
'                                End If
'                                SELELPROSUBRS.MoveNext
'                            Wend
'                        End If
'
''                    DMULPROCE_RS.MoveNext
''                    Wend
'
'                Else
'
'                        If SELMA101FLAG = False Then
'                            SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                            SELMA101FLAG = True
'                        Else
'                            SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            SELMA101_RS.MoreResults
'                            SELMA101_RS.Requery
'                        End If
'                         If SELDTEMPLETFLAG = False Then
'                            SELDTEMPLETQR(0) = DMULPROCE_RS("PRO_CODE")
'                            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                            SELDTEMPLETFLAG = True
'                        Else
'                            SELDTEMPLETQR(0) = DMULPROCE_RS("PRO_CODE")
'                            SELDTEMPLETRS.MoreResults
'                            SELDTEMPLETRS.Requery
'                        End If
'
'
'                        '''''''''' KANCHAN & PRATS 08/01/2008
'                        ''*** CODE TO RETRIVE OBSERVED VALUE FROM INTERFACE.DINTERFACE
'
'                                If seldinterfaceflag = False Then
'                                    seldinterface_qr(0) = VAR_REQNNO
'                                    seldinterface_qr(1) = Trim(DMULPROCE_RS("PRO_CODE"))
'                                    Set seldinterface_rs = seldinterface_qr.OpenResultset(1, 3)
'                                    seldinterfaceflag = True
'                                Else
'                                    seldinterface_qr(0) = VAR_REQNNO
'                                    seldinterface_qr(1) = Trim(DMULPROCE_RS("PRO_CODE"))
'                                    seldinterface_rs.MoreResults
'                                    seldinterface_rs.Requery
'                                End If
'
'                        If seldinterface_rs.RowCount > 0 Then
'
'                            While Not seldinterface_rs.EOF
'                                    If SELDTEMPLETFLAG = False Then
'                                        SELDTEMPLETQR(0) = seldinterface_rs("testcode")
'                                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                                        SELDTEMPLETFLAG = True
'                                    Else
'                                        SELDTEMPLETQR(0) = seldinterface_rs("TESTCODE")
'                                        SELDTEMPLETRS.MoreResults
'                                        SELDTEMPLETRS.Requery
'                                    End If
'''''''''KANCHAN
'
'                                    If SELDTEMPLETRS.RowCount > 0 Then
'
'                                    If seldinterface_rs("TESTCODE") = "F998" Then
'
'                                        OSMOAGE
'
'                                        If SELDBCOSMOFLAG = False Then
'                                            SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
'                                            SELDBCOSMOQR(1) = RNAGE
'                                            SELDBCOSMOQR(2) = SS
'                                            Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
'                                            SELDBCOSMOFLAG = True
'                                        Else
'                                            SELDBCOSMOQR(0) = seldinterface_rs("TESTCODE")
'                                            SELDBCOSMOQR(1) = RNAGE
'                                            SELDBCOSMOQR(2) = SS
'                                            SELDBCOSMORS.MoreResults
'                                            SELDBCOSMORS.Requery
'                                        End If
'
'                                        RNAGE = SELDBCOSMORS("RANGE")
'
'                                        OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
'                                        OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
'
'
'                                        If SELMA101_RS.RowCount > 0 Then
'                                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
'                                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                        Else
'                                            SELDTEMPLETRS.MoveNext
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                        End If
'                                        i = i + 1
'                                    End If
'                                  Else
'                                        If SELMA101_RS.RowCount > 0 Then
'                                            If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                        Else
'                                            SELDTEMPLETRS.MoveNext
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                        End If
'                                        i = i + 1
'                                    End If
'                            End If
'                        End If
'                     'End If
'
'
'
'
'
''''''KANCHAN
'
'                            Wend
'                        End If
'********** KANCHAN
                        
'                        If PC = "F999" Then
'                            OSMOAGE
'
'                            If SELDBCALKPHSGGTFLAG = False Then
'                                    SELDBCALKPHSGGTQR(0) = DMULPROCE_RS("PRO_CODE")
'                                    SELDBCALKPHSGGTQR(1) = RNAGE
'                                    SELDBCALKPHSGGTQR(2) = SS
'                                    Set SELDBCALKPHSGGTRS = SELDBCALKPHSGGTQR.OpenResultset(1, 3)
'                                    SELDBCALKPHSFLAG = True
'                            Else
'                                    SELDBCALKPHSGGTQR(0) = DMULPROCE_RS("PRO_CODE")
'                                    SELDBCALKPHSGGTQR(1) = RNAGE
'                                    SELDBCALKPHSGGTQR(2) = SS
'                                    SELDBCALKPHSGGTRS.MoreResults
'                                    SELDBCALKPHSGGTRS.Requery
'                            End If
'
'
'                            RANGE = SELDBCALKPHSGGTRS!RANGE
'
'                            ALKLV = Left(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
'                            ALKHV = Right(Trim(SELDBCALKPHSGGTRS("RANGE")), 3)
'                        End If

            '''' COMMENTED BY KANCHAN
                        
'                        If SELDTEMPLETRS.RowCount > 0 Then
'
'                            If DMULPROCE_RS("PRO_CODE") = "F998" Then
'
'                                OSMOAGE
'
'                                If SELDBCOSMOFLAG = False Then
'                                        SELDBCOSMOQR(0) = DMULPROCE_RS("PRO_CODE")
'                                        SELDBCOSMOQR(1) = RNAGE
'                                        SELDBCOSMOQR(2) = SS
'                                        Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
'                                        SELDBCOSMOFLAG = True
'                                Else
'                                        SELDBCOSMOQR(0) = DMULPROCE_RS("PRO_CODE")
'                                        SELDBCOSMOQR(1) = RNAGE
'                                        SELDBCOSMOQR(2) = SS
'                                        SELDBCOSMORS.MoreResults
'                                        SELDBCOSMORS.Requery
'                                End If
'
'                                RNAGE = SELDBCOSMORS("RANGE")
'
'                                OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
'                                OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
'
'
'                                    If SELMA101_RS.RowCount > 0 Then
'                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                        Else
'                                            SELDTEMPLETRS.MoveNext
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & RNAGE
'                                        End If
'                                        i = i + 1
'                                    End If
'                            Else
'                                    If SELMA101_RS.RowCount > 0 Then
'                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                        Else
'                                            SELDTEMPLETRS.MoveNext
'                                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & Trim(seldinterface_rs("TESTRESULT")) & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                        End If
'                                        i = i + 1
'                                    End If
'                            End If
'                        End If
'                     End If
'                    DMULPROCE_RS.MoveNext
'                   Wend
                   
                ''''' kanchan
                '******
'                While Not DMULPROCE_RS.EOF
'                    If SELMA101FLAG = False Then
'                        SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                        SELMA101FLAG = True
'                    Else
'                        SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                        SELMA101_RS.MoreResults
'                        SELMA101_RS.Requery
'                    End If
'                     If SELDTEMPLETFLAG = False Then
'                        SELDTEMPLETQR(0) = DMULPROCE_RS("PRO_CODE")
'                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                        SELDTEMPLETFLAG = True
'                    Else
'                        SELDTEMPLETQR(0) = DMULPROCE_RS("PRO_CODE")
'                        SELDTEMPLETRS.MoreResults
'                        SELDTEMPLETRS.Requery
'                    End If
'                    If SELDTEMPLETRS.RowCount > 0 Then
'                        If SELMA101_RS.RowCount > 0 Then
'                            If SELDTEMPLETRS("SEX") = "M" Or Trim(SELDTEMPLETRS("SEX")) = "" Then
'                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                            Else
'                                SELDTEMPLETRS.MoveNext
'                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                            End If
'                            i = i + 1
'                        End If
'                    End If
'                DMULPROCE_RS.MoveNext
'                Wend
                
            End If
''            MSFlexGrid1.row = 1
  '          MSFlexGrid1.Col = 2
        '    txt_observed.Visible = True
            
'            If seldissecentrflag1 = False Then
'                Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'                seldissecentrflag1 = True
'            Else
'                seldissecentr1_rs.MoreResults
'                seldissecentr1_rs.Requery
'            End If
'            If seldissecentr1_rs.RowCount > 0 Then
'                ST = cmb_enterby.Text
'                cmb_enterby.clear
'                While Not seldissecentr1_rs.EOF
'                    cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                    seldissecentr1_rs.MoveNext
'                Wend
'                cmb_enterby.Text = ST
'                ST = ""
'            End If

'            If seldissecfinflag = False Then
'                Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'                seldissecfinflag1 = True
'            Else
'                seldissecfin1_rs.MoreResults
'                seldissecfin1_rs.Requery
'            End If
'            If seldissecfin1_rs.RowCount > 0 Then
'                ST = cmb_finalby.Text
'                cmb_finalby.clear
'                While Not seldissecfin1_rs.EOF
'                    cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                    seldissecfin1_rs.MoveNext
'                Wend
'                cmb_finalby.Text = ST
'            End If
         
      End If
      
       If viewflag = True Then
          'viewflag = False
        ' txt_observed.Visible = False
        ''******ICU****************
        '***** Added by Shraddha on 21/05/10 *****
        If opt_staff.Value = True Then
            STAFF
        Else
        '***************
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
                  lbl_rwd.Visible = True
                  lblicu.Visible = False
               '***** Shraddha 05/12/2011 *****
                ElseIf Trim(SELICU_RS("ICDS2")) = "PR" Then
                      lbl_rwd.Visible = False
                      lblicu.Visible = False
                      LBL_CAPTIONS.Caption = "PROTOCOL"
                      LBL_CAPTIONS.Visible = True
                ElseIf Trim(SELICU_RS("ICDS2")) = "CM" Then
                      lbl_rwd.Visible = False
                      lblicu.Visible = False
                      LBL_CAPTIONS.Caption = "CASUALTY MAIN"
                      LBL_CAPTIONS.Visible = True
                ElseIf Trim(SELICU_RS("ICDS2")) = "CG" Then
                      lbl_rwd.Visible = False
                      lblicu.Visible = False
                      LBL_CAPTIONS.Caption = "CASUALTY GJB"
                      LBL_CAPTIONS.Visible = True
               End If
               End If
            End If
            ''******ICU****************
        If SELTA101FALG = False Then
            SELTA101_QR(0) = MaskEdBox1.FormattedText
            Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            SELTA101FALG = True
        Else
            SELTA101_QR(0) = MaskEdBox1.FormattedText
            Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            SELTA101_RS.MoreResults
            SELTA101_RS.Requery
        End If
           
        If SELTA101_RS.RowCount = 0 Then
            Exit Sub
        End If
    
        txt_sex.Text = SELTA101_RS("sex")
        txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
        txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
        
'        If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
'                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)
'                   Dim N As Integer
'                   N = YEAR1 / 12
'                   txt_age.Text = N
'        End If
'
'        If Trim(txt_age.Text) = 0 Then
'                  Dim age1 As String
'                  age1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                  age1 = "." & (age1 Mod 12)
'                     If age1 = ".0" Then
'                        age1 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        age1 = ".0" & (age1) ' Mod 30)
'                     End If
'                  txt_age.Text = Trim(age1)
'        End If
        
        txt_pridoc.Text = SELTA101_RS("ud_name")
        txt_category.Text = SELTA101_RS("pat_catg")
        
            If selpathdrequisitnrsflag = False Then
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                selpathdrequisitnrsflag = True
            Else
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = Trim(selpathdrequisitn_rs("pathno"))
            End If
        
        If REFDRFLAG = False Then
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        Else
            REFDR_QR(0) = MaskEdBox1.FormattedText
            REFDR_QR(1) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
            REFDR_RS.MoreResults
            REFDR_RS.Requery
        End If
    
        If REFDR_RS.RowCount > 0 Then
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If REFDR_RS.EOF = False Then
            cmb_enterby.Text = " " & Trim(REFDR_RS("enterby"))
            End If
            If REFDR_RS.EOF = False Then
                cmb_finalby.Text = " " & Trim(REFDR_RS("freportby"))
            End If
            If REFDR_RS("COMMITFLAG") = "Y" Then
                MsgBox "Final Report is Generated, You cann't Modify the Final Report."
                cmd_modify.Enabled = False
                ADDENDUMFALG = True
                cmd_addendum.Visible = True
            End If
       End If
    
             
'        If selma104flag = False Then
'            selma104_qr(0) = MaskEdBox1.FormattedText
'            Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'            selma104flag = True
'        Else
'            selma104_qr(0) = MaskEdBox1.FormattedText
'            selma104_rs.MoreResults
'            selma104_rs.Requery
'        End If
      
      
        If REFDR_RS.RowCount > 0 Then
           If REFDR_RS("patnstat") = "IP" Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                lbl_ward.Visible = True
                lbl_bed.Visible = True
                txt_ward.Text = selpathdrequisitn_rs("WARD_NAME")
                txt_bedno.Text = selpathdrequisitn_rs("bedno")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_status.Text = "OP"
            End If
       End If
            
            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F995" Then
                LBL_VBG.Visible = True
            Else
                LBL_VBG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F901" Then
                LBL_ABG.Visible = True
            Else
                LBL_ABG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F902" Then
                LBL_CBG.Visible = True
            Else
                LBL_CBG.Visible = False
            End If

           If DMPROCEDURFLAG = False Then
                DMPROCEDUR_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMPROCEDUR_RS = DMPROCEDUR_QR.OpenResultset(1, 3)
                DMPROCEDURFLAG = True
            Else
                DMPROCEDUR_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMPROCEDUR_RS.MoreResults
                DMPROCEDUR_RS.Requery
            End If
        
                        
            If DMPROCEDUR_RS.RowCount > 0 Then
                i = 1
               ' MSFlexGrid1.Rows = 1
                'While Not DMPROCEDUR_RS.EOF
'                    If SELMA101FLAG = False Then
'                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
'                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                        SELMA101FLAG = True
'                    Else
'                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
'                        SELMA101_RS.MoreResults
'                        SELMA101_RS.Requery
'                    End If
'                    If SELDTEMPLETFLAG = False Then
'                        SELDTEMPLETQR(0) = DMPROCEDUR_RS("PRO_CODE")
'                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                        SELDTEMPLETFLAG = True
'                    Else
'                        SELDTEMPLETQR(0) = DMPROCEDUR_RS("PRO_CODE")
'                        SELDTEMPLETRS.MoreResults
'                        SELDTEMPLETRS.Requery
'                    End If
'                    If SELDTEMPLETRS.RowCount > 0 Then
'                        If SELMA101_RS.RowCount > 0 Then
'                            If SELDTEMPLETRS("sex") = "M" Or Trim(SELDTEMPLETRS("SEX")) = "" Then
'                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                            Else
'                               SELDTEMPLETRS.MoveNext
'                               MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                            End If
'                            i = i + 1
'                        End If
'                    End If
' ***** COLOR CODING

                    i = 1
                    srno = 1
                   
                    
                While Not DMPROCEDUR_RS.EOF
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                                     
                    
                    If SELDTEMPLETFLAG = False Then
                        SELDTEMPLETQR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        SELDTEMPLETQR(1) = Trim(DMPROCEDUR_RS("remarks")) 'Shraddha 31/12/10
                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                        SELDTEMPLETFLAG = True
                    Else
                        SELDTEMPLETQR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        SELDTEMPLETQR(1) = Trim(DMPROCEDUR_RS("remarks")) 'Shraddha 31/12/10
                        SELDTEMPLETRS.MoreResults
                        SELDTEMPLETRS.Requery
                    End If
                    '***** Shraddha 31/12/10 *****
                    If Trim(DMPROCEDUR_RS("remarks")) = "" Or IsNull(DMPROCEDUR_RS("remarks")) = True Then
                        MsgBox "m/c name not entered"
                        Exit Sub
                    End If
                    '*****************************
                    'for alkaline phosphate changes in range according to age
                    SEX = Trim(txt_sex.Text)
                    If Trim(DMPROCEDUR_RS("PRO_CODE")) = "F957" Then
                       If Trim(txt_age.Text) <= 15 Then
                          SEX = "C"
                       Else
                          SEX = "A"
                       End If
                    End If
                    
                    
                    If SELDTEMPLETRS.RowCount > 0 Then
                                                    
                            If SELMA101_RS.RowCount > 0 Then
                            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                             MSFlexGrid1.row = srno
                             MSFlexGrid1.col = 0
                            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
                            If SELDTEMPLETRS("SEX") = SEX Or Trim(SELDTEMPLETRS("SEX")) = "" Then
                                MSFlexGrid1.TextMatrix(i, 0) = i
                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101_RS("SER_DET")) & "$" & Trim(SELMA101_RS("SER_code"))
                                If Trim(DMPROCEDUR_RS("obser_type")) = "H" Then
                                    MSFlexGrid1.row = srno
                                     MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                                Else
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "L" Then
                                      
                                      MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = RGB(100, 255, 150)
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "N" Then
                                     
                                       MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "" Then
                                      
                                        MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                End If
                                j = 0
                                j = InStr(DMPROCEDUR_RS("OBSERVALUE"), "*")
                                If j = 0 Then
                                    MSFlexGrid1.TextMatrix(i, 2) = DMPROCEDUR_RS("OBSERVALUE")
                                Else
                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                    MSFlexGrid1.TextMatrix(i, 4) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                End If
                                MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                
                                If DMPROCEDUR_RS("PRO_CODE") = "F998" Then

                                    OSMOAGE

                                    If SELDBCOSMOFLAG = False Then
                                            SELDBCOSMOQR(0) = DMPROCEDUR_RS("PRO_CODE")
                                            SELDBCOSMOQR(1) = RNAGE
                                            SELDBCOSMOQR(2) = SS
                                            Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
                                            SELDBCOSMOFLAG = True
                                    Else
                                            SELDBCOSMOQR(0) = DMPROCEDUR_RS("PRO_CODE")
                                            SELDBCOSMOQR(1) = RNAGE
                                            SELDBCOSMOQR(2) = SS
                                            SELDBCOSMORS.MoreResults
                                            SELDBCOSMORS.Requery
                                    End If

                                    RNAGE = SELDBCOSMORS!range

                                    OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
                                    OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
                                End If

                                If Trim(DMPROCEDUR_RS("PRO_CODE")) = "F998" Then
                                    MSFlexGrid1.TextMatrix(i, 3) = OSMOLV & "-" & OSMOHV
                                Else
                                    MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                End If
                                
                            Else
                                SELDTEMPLETRS.MoveNext
                                MSFlexGrid1.TextMatrix(i, 0) = i
                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101_RS("SER_DET")) & "$" & Trim(SELMA101_RS("SER_CODE"))
                                If Trim(DMPROCEDUR_RS("obser_type")) = "H" Then
                                    MSFlexGrid1.row = srno
                                     MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                                Else
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "L" Then
                                      
                                      MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = RGB(150, 255, 100)
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "N" Then
                                     
                                       MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "" Then
                                      
                                        MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                End If
                                j = 0
                                j = InStr(DMPROCEDUR_RS("OBSERVALUE"), "*")
                                If j = 0 Then
                                    MSFlexGrid1.TextMatrix(i, 2) = DMPROCEDUR_RS("OBSERVALUE")
                                Else
                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                    MSFlexGrid1.TextMatrix(i, 4) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                End If
                                
                                If DMPROCEDUR_RS("PRO_CODE") = "F998" Then
                                
                                    OSMOAGE
        
                                    If SELDBCOSMOFLAG = False Then
                                            SELDBCOSMOQR(0) = DMPROCEDUR_RS("PRO_CODE")
                                            SELDBCOSMOQR(1) = RNAGE
                                            SELDBCOSMOQR(2) = SS
                                            Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
                                            SELDBCOSMOFLAG = True
                                    Else
                                            SELDBCOSMOQR(0) = DMPROCEDUR_RS("PRO_CODE")
                                            SELDBCOSMOQR(1) = RNAGE
                                            SELDBCOSMOQR(2) = SS
                                            SELDBCOSMORS.MoreResults
                                            SELDBCOSMORS.Requery
                                    End If
            
                                    RNAGE = SELDBCOSMORS!range
        
                                    OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
                                    OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
                                End If

                                If Trim(DMPROCEDUR_RS("PRO_CODE")) = "F998" Then
                                    MSFlexGrid1.TextMatrix(i, 3) = OSMOLV & "-" & OSMOHV
                                Else
                                    SELDTEMPLETRS.MoveFirst
                                    'i = SELDTEMPLETRS.RowCount
                                    MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                End If
                                
                                    
                            End If
                        End If
                        
                        '***** Shraddha 30/12/10 *****
                        If Trim(DMPROCEDUR_RS("remarks")) <> "" Or IsNull(DMPROCEDUR_RS("remarks")) = False Then
                            MSFlexGrid1.TextMatrix(i, 5) = Trim(DMPROCEDUR_RS("remarks"))
                        End If
                        '*****************************
                           srno = srno + 1
                            i = i + 1

                    End If
                    DMPROCEDUR_RS.MoveNext
'                    DMULPROCE_RS.MoveNext

                Wend
'                MSFlexGrid1.Row = 1
'                MSFlexGrid1.Col = 2
'                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                txt_observed.Visible = False




'***** COLOR CODING
'                    DMPROCEDUR_RS.MoveNext
'                Wend
            End If
            viewflag = False 'Shraddha 21/05/10
      End If
      If SELDMBCCOMMFLAG = False Then
        SELDMBCCOMMQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
        Set SELDMBCCOMMRS = SELDMBCCOMMQR.OpenResultset(1, 3)
        SELDMBCCOMMFLAG = True
      Else
        SELDMBCCOMMQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
        SELDMBCCOMMRS.MoreResults
        SELDMBCCOMMRS.Requery
      End If
      If SELDMBCCOMMRS.RowCount > 0 Then
            If SELDMBCCOMMRS.EOF = False Then
                txt_comments.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
            End If
      End If
      MSFlexGrid1.Enabled = True
      
      'addendum
       selectaddendum_qr(0) = req_chglen(cmb_reqno.Text)
       If Addendum_Flag = False Then
          Set selectaddendum_rs = selectaddendum_qr.OpenResultset(2, 3)
          Addendum_Flag = True
       Else
          selectaddendum_rs.MoreResults
          selectaddendum_rs.Requery
       End If
      
      If selectaddendum_rs.RowCount > 0 Then
         selectaddendum_rs.MoveFirst
         While Not selectaddendum_rs.EOF
           txt_addendum.Text = txt_addendum.Text & "  " & Trim(selectaddendum_rs("add_max_ln"))
           selectaddendum_rs.MoveNext
         Wend
      End If
      
      End If 'Shraddha 21/05/10
'End If




'***** Commented by Shraddha on 21/05/10 *****




'For showing previous values
Call prev_val
'*******************************
If opt_staff.Value = True Then
    STAFF
End If
'*********************************************
Exit Sub:
'MSG:
 '   MsgBox Err.Description

End Sub

Private Sub cmb_reqno_LostFocus()
'*******Shraddha B PANDIT********
Dim C_QR As rdoQuery
Dim C_RS As rdoResultset
Dim sql As String
sql = "select REQNO from dislib.dmulproce where reqno=? AND PRO_CODE IN ('F903','F904','F905','F906','F907','F911','F912')"
Set C_QR = rdb.CreateQuery("Cqr1", sql)
C_QR(0) = req_chglen(cmb_reqno.Text)
Set C_RS = C_QR.OpenResultset(1, 1)
If C_RS.RowCount > 0 Then
    
    'STR = "SELECT USERNAME FROM DISLIB.REPWRITERS,DISLIB.DSECURITY "
    'STR = STR + " Where USER_ID = USR_ID And USRTYPE = '2' and dept_code='FEC' AND (USRTYPE='8' AND USER_ID LIKE 'A4%') AND DEPT_CODE='FEC'"
    'STR = STR + " Where USER_ID = USR_ID And USRTYPE = '2' and dept_code='FEC'"
    'STR = STR + " Where USER_ID = USR_ID And MODXX = 'A' AND dept_code='FEC'"
    'Next Query statement added on 03/11/2008
    str = "SELECT USERNAME FROM DISLIB.DSECURITY, DISLIB.REPWRITERS "
    str = str + "Where USER_ID = USR_ID And (USRTYPE = '2' or USRTYPE = '8') "
    str = str + "AND DEPT_CODE='FEC' AND USER_ID LIKE 'A4%' AND MODXX=6"
    
    Set REPWTRN_QR = rdb.CreateQuery("", str)
    Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    
    If REPWTRNRS.RowCount > 0 Then
        cmb_finalby.clear
        REPWTRNRS.MoveFirst
        While Not REPWTRNRS.EOF
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            REPWTRNRS.MoveNext
        Wend
        cmb_finalby.Text = cmb_finalby.List(0)
    Else
        MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
        cmd_cancel_Click
        Exit Sub
    End If
Else
    str = "SELECT USERNAME FROM DISLIB.REPWRITERS,DISLIB.DSECURITY "
    str = str + " Where USER_ID = USR_ID And (USRTYPE = '8' OR USRTYPE='9') and dept_code='FEC' and (USR_ID LIKE 'A2%' OR USR_ID LIKE 'A4%') AND MODXX='6'"
    'STR = STR + " Where USER_ID = USR_ID And (USRTYPE = '8' OR USRTYPE='9') and dept_code='FEC' and (USER_ID LIKE '%4')"
    Set REPWTRN_QR = rdb.CreateQuery("", str)
    Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    
    If REPWTRNRS.RowCount > 0 Then
        cmb_finalby.clear
        REPWTRNRS.MoveFirst
        While Not REPWTRNRS.EOF
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            REPWTRNRS.MoveNext
        Wend
        cmb_finalby.Text = cmb_finalby.List(0)
    Else
        MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
        cmd_cancel_Click
        Exit Sub
    End If
End If
'****************************************************
'***** Added on 05/11/08 by Shraddha Pandit as told by Mrs. Surekha M'am *****
Dim z As Integer
For z = 0 To cmb_finalby.ListCount
    If Trim(cmb_enterby.List(0)) = cmb_finalby.List(z) Then
        cmb_finalby.Text = cmb_finalby.List(z)
    End If
Next z
'*****************************************************************************
End Sub

Private Sub cmd_addendum_Click()
  cmd_save.Enabled = True
  txt_addendum.Enabled = True
End Sub

Private Sub cmd_back_Click()
DIS_MAIN.show
cmd_cancel_Click
Me.Hide
End Sub

Private Sub cmd_cancel_Click()
btnenabl ("INI")
clear
LBL_CAPTIONS.Caption = ""
cmb_machine.Visible = False
txt_observed.Visible = False

FLDENABL ("DIS")
flagsave = 0
End Sub

Private Sub cmd_commit_Click()
Dim i As Integer
'If Len(Trim(cmb_finalby)) = 0 Then
'    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
'    cmb_finalby.SetFocus
'    Exit Sub
'End If
i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If i = vbYes Then
On Error GoTo MSG:

''********SECURITY***************************

 
'STR = "select * from DISLIB.Dislnkusr where username = ?"
'Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11MICRO", STR)
'CURUSRDET_QR(0) = "ABC"
'Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
 
 
    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
        'COMMENTED BY SHWETA 22-08-2013 AS TOLD BY KALPANA MAM REPORTS TO BE COMMITED BY AT TIME OF SAVING
            'MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
           ' cmd_commit.SetFocus
            'Exit Sub
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
            If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
                '  If loginusrid <> curusrdetrs("usr_id") Then
                'MsgBox "Sorry! This Report is not to be Finalized by You"
                'cmd_commit.SetFocus
                'Exit Sub
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
    updrrequisitn1_qr(0) = UCase(Left(cmb_finalby.Text, 30))
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    
    'RTIMELOG 18-01-05
        UPDRTIMEC_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDRTIMEC_QR(1) = time1 'txt_time.Text
        UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
        UPDRTIMEC_QR.Execute
   
    MsgBox "Record is Committed Successfully."
    
End If
rdb.CommitTrans
btnenabl ("COMMIT")

cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description


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
btnenabl ("NEW")
FLDENABL ("MODIFY")
'ADDED BY AMARJEET ON 08-11-2002 FOR NOT TO ALLOW MODIFICATION IF COMMITED
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
            If REFDR_RS("COMMITFLAG") = "Y" Then
                txt_observed.Visible = False
            Else
                txt_observed.Visible = True
            End If
    End If
'MSFlexGrid1.row = 1 'abd
'MSFlexGrid1.Col = 2
'txt_observed.Text = MSFlexGrid1.TextMatrix(1, 2)
MODIFYFLAG = True
               ' MSFlexGrid1.Rows = 1
'                For i = 1 To MSFlexGrid1.Rows - 1
'                UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 1)
'                UPDDMPROCEDUR_QR(1) = req_chglen(cmb_reqno.Text)
'                 UPDDMPROCEDUR_QR(2) = MSFlexGrid1.TextMatrix(i, 2)
'
'                'i = i + 1
'               ' UPDDMPROCEDUR_RS.MoveNext
'               ' Wend
'            'End If
''UPDDMPROCEDUR_QR(1) = req_chglen(cmb_reqno.Text)
' UPDDMPROCEDUR_QR.Execute
'Next i

'If seldissecentrflag1 = False Then
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
MSFlexGrid1.col = 2
txt_observed.Width = MSFlexGrid1.CellWidth
txt_observed.Height = MSFlexGrid1.CellHeight
txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
End Sub
Private Sub cmd_new_Click()
Dim L1 As Integer
btnenabl ("NEW")
FLDENABL ("ENA")
clear
LBL_CAPTIONS.Caption = ""
    If DATETIMEFLAG = False Then
    datetime
    DATETIMEFLAG = True
     Else
    datetime
    End If
    txt_date.Text = DATE11
    txt_time.Text = time1
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Enabled = False
    For L1 = 0 To cmb_finalby.ListCount
        If Trim(cmb_enterby.List(0)) = cmb_finalby.List(L1) Then
            cmb_finalby.Text = cmb_finalby.List(L1)
        End If
    Next L1
    'cmb_finalby.Text = cmb_finalby.List(0)
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
'cmb_modality.ListIndex = 0
cmb_dept.SetFocus
Opt_patient.Value = True
MaskEdBox1.Enabled = True
NEWFLAG = True
flagsave = 1
On Error GoTo NXT
MaskEdBox1.SetFocus
NXT:
End Sub

Private Sub cmd_print_Click()
'On Error GoTo MSG:


Dim n, d As String

'RTIMELOG 18-01-05
        UPDRTIMEP_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDRTIMEP_QR(1) = time1 'txt_time.Text
        UPDRTIMEP_QR(2) = req_chglen(cmb_reqno.Text)
        UPDRTIMEP_QR.Execute
        
Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim i As Integer
 'On Error GoTo prtmsg:

DoEvents

'******************************* for displaying designation ************'
'********** added by milind on 03-03-06
userid_qr(0) = Trim(cmb_finalby.Text)
Set userid_rs = userid_qr.OpenResultset(1, 3)
n = userid_rs(0)
userdesc_qr(0) = n
Set userdesc_rs = userdesc_qr.OpenResultset(1, 3)
d = userdesc_rs(0)
'************upto here**********



    VAR_REQNNO = req_chglen(cmb_reqno.Text)

        If REFDRFLAG1 = False Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = VAR_REQNNO
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = VAR_REQNNO
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
            

'    If SELDREQUISITNFLAG = False Then
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'            SELDREQUISITNFLAG = True
'        Else
'            seldrequisitn_qr(0) = MaskEdBox1.FormattedText
'            seldrequisitn_rs.MoreResults
'            seldrequisitn_rs.Requery
'        End If
        If SELRREQUISITN_RS.RowCount > 0 Then
            prepath = ""

            While Not SELRREQUISITN_RS.EOF
                If Trim(SELRREQUISITN_RS("reqnno")) <> VAR_REQNNO Then
                    prepath = prepath & req_chglen(SELRREQUISITN_RS("reqnno")) & " "
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
 
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
If Opt_patient.Value = True Then
      '  CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSPFINAL.RPT"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.RPT"
   ''**********ICU********************
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
        
'        If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
'                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)
'                   Dim N As Integer
'                   N = YEAR1 / 12
'                   'txt_age.Text = N
            
'        End If
'
'        If Trim(txt_age.Text) = 0 Then
'                  Dim age1 As String
'                  age1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                  age1 = "." & (age1 Mod 12)
'                     If age1 = ".0" Then
'                        age1 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        age1 = ".0" & (age1) ' Mod 30)
'                     End If
'                  'txt_age.Text = Trim(age1)
'                  CrystalReport1.Formulas(2) = "dob = '" & age1 & "'"
'        End If
        'ADDED BY MILIND FOR DISPLAYING ARTERIAL OR VENOUS OR CAPILLARY
        'BLOOD GAS REPORT AS A HEADING IN REPORT ON 29-03-2007
        
        SELBG_QR(0) = VAR_REQNNO
        Set SELBG_RS = SELBG_QR.OpenResultset(1, 1)
        
        If SELBG_RS.RowCount > 0 Then
            
                If Trim(SELBG_RS("PRO_CODE")) = "F901" And SELBG_RS.RowCount = 2 Then
                    
                    SELBG_RS.MoveNext
                    
                    If Trim(SELBG_RS("PRO_CODE")) = "F022" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "ARTERIAL BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                        
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F995" And SELBG_RS.RowCount = 2 Then
                    
                    SELBG_RS.MoveNext
                    
                    If Trim(SELBG_RS("PRO_CODE")) = "F022" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "VENOUS BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                                        
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F902" And SELBG_RS.RowCount = 2 Then
                    
                    SELBG_RS.MoveNext
                    
                    If Trim(SELBG_RS("PRO_CODE")) = "F022" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "CAPILLARY BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                        
                End If
                
                If Trim(SELBG_RS("PRO_CODE")) = "F901" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "ARTERIAL BLOOD GAS" & "'"
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F995" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "VENOUS BLOOD GAS" & "'"
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F902" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "CAPILLARY BLOOD GAS" & "'"
'                Else
'                    CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                End If
        
        End If
        'UPTO HERE
        
        
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
              '  CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL1.RPT"
              CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.RPT"
            End If
        End If
    End If
      
      ''**********ICU********************
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
'        End If
'        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                
'        Dim selproc_qr As rdoQuery
'        Dim selproc_rs As rdoResultset
'        'Dim n
'        STR = "select pro_code from dislib.dmprocedur where reqno=? order by pro_code"
'        Set selproc_qr = rdb.CreateQuery("selproc_qr", STR)
'        selproc_qr(0) = req_chglen(cmb_reqno.Text)
'        Set selproc_rs = selproc_qr.OpenResultset(1, 3)
'
'        Dim selserdet_qr As rdoQuery
'        Dim selserdet_rs As rdoResultset
'
'        STR = "select ser_det from patoolslib.ma101 where ser_code=?"
'        Set selserdet_qr = rdb.CreateQuery("selserdet_qr", STR)
'
'        While selproc_rs.EOF = False
'            selserdet_qr(0) = selproc_rs(0)
'            Set selserdet_rs = selserdet_qr.OpenResultset(1, 3)
'            selproc_rs.MoveNext
'                While selserdet_rs.EOF = False
'                    CrystalReport1.Formulas(18) = "AS= '" & selserdet_rs("ser_det") & "'"
'                    selserdet_rs.MoveNext
'                Wend
'        Wend
        
                
'        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
'            'CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(cmb_enterby.Text)
'            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'            'CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
  
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If pcount <> 0 Then
        
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        End If
            CrystalReport1.Formulas(17) = "DESG = '" & d & "'"
           ' CrystalReport1.Formulas(35) = "DOB = '" & txt_Age & "'"
            
        '***** Shraddha 22/10/10 *****
        Dim SELDISWARD_QR As rdoQuery
        Dim SELDISWARD_RS As rdoResultset
        str = "SELECT * FROM DISLIB.DIS_WARD WHERE IPADD=?"
        Set SELDISWARD_QR = rdb.CreateQuery("", str)
        
        SELDISWARD_QR(0) = Trim(ipadd)
        Set SELDISWARD_RS = SELDISWARD_QR.OpenResultset(1, 3)
        
        If Not SELDISWARD_RS.EOF Then
            CrystalReport1.Formulas(17) = "WARDLOC = '" & Trim(SELDISWARD_RS("LOCATION")) & "'"
        Else
               CrystalReport1.Formulas(17) = "WARDLOC = ''"
        End If
        SELDISWARD_RS.Close
        
        
        CrystalReport1.Formulas(18) = "ipaddress = '" & ipadd & "'"
        '*****************************
        
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
End If
If opt_staff.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISELS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                
        'ADDED BY MILIND FOR DISPLAYING ARTERIAL OR VENOUS OR CAPILLARY
        'BLOOD GAS REPORT AS A HEADING IN REPORT ON 29-03-2007
        
        SELBG_QR(0) = VAR_REQNNO
        Set SELBG_RS = SELBG_QR.OpenResultset(1, 1)
        
        If SELBG_RS.RowCount > 0 Then
            
                If Trim(SELBG_RS("PRO_CODE")) = "F901" And SELBG_RS.RowCount = 2 Then
                    
                    SELBG_RS.MoveNext
                    
                    If Trim(SELBG_RS("PRO_CODE")) = "F022" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "ARTERIAL BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                        
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F995" And SELBG_RS.RowCount = 2 Then
                    
                    SELBG_RS.MoveNext
                    
                    If Trim(SELBG_RS("PRO_CODE")) = "F022" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "VENOUS BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                                        
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F902" And SELBG_RS.RowCount = 2 Then
                    
                    SELBG_RS.MoveNext
                    
                    If Trim(SELBG_RS("PRO_CODE")) = "F022" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "CAPILLARY BLOOD GAS" & "'"
                    Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                    End If
                        
                End If
                
                If Trim(SELBG_RS("PRO_CODE")) = "F901" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "ARTERIAL BLOOD GAS" & "'"
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F995" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "VENOUS BLOOD GAS" & "'"
                ElseIf Trim(SELBG_RS("PRO_CODE")) = "F902" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "CAPILLARY BLOOD GAS" & "'"
'                Else
'                    CrystalReport1.Formulas(5) = "PROFIN = '" & "NO" & "'"
                End If
        
        End If
        'UPTO HERE

        
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            'CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            'CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
  
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        If pcount <> 0 Then
        
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        End If
            CrystalReport1.Formulas(17) = "DESG = '" & d & "'"
            
            CrystalReport1.Formulas(18) = "ipaddress = '" & ipadd & "'"
        
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
    
End If
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
    pcount = 0
    For i = 1 To MSFlexGrid1.Rows - 1
        If Trim(MSFlexGrid1.TextMatrix(i, 4)) <> "" Then
          pcount = pcount + 1
        End If
    Next i
    If pcount <> 0 Then
      j = MsgBox(pcount & " Panic Value, Do you want to Save? ", vbYesNo)
        If j = vbNo Then
         Exit Sub
        End If
    End If
    
    '***** Shraddha 31/12/10 *****
    If Trim(MSFlexGrid1.TextMatrix(1, 5)) = "" Then
        MsgBox "Please select machine name"
        Exit Sub
    End If
    '*****************************
    
assign
btnenabl ("SAVE")
'cmb_finalby.clear
End Sub

Private Sub cmd_view_Click()
btnenabl ("VIEW")
FLDENABL ("VIEW")
txt_observed.Visible = False
clear
LBL_CAPTIONS.Caption = ""
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
If sel_deptrs.RowCount > 0 Then
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
        Wend
End If
cmb_dept.SetFocus
Opt_patient.Value = True
viewflag = True
flagsave = 2
'cmb_modality.ListIndex = 0
End Sub



Private Sub MSFlexGrid1_Click()
If MSFlexGrid1.col = 2 Then
        txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
        txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        txt_observed.Visible = True
        txt_observed.Enabled = True
End If
If MSFlexGrid1.col = 5 Then
        cmb_machine.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        cmb_machine.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
        cmb_machine.Visible = True
End If
End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
   ' MaskEdBox1.Mask = "??>/#####"
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    'MaskEdBox1.Mask = "#######"                ' COMMENTED BY VARINDER
End If
End Sub
Private Sub Form_Load()
On Error GoTo MSG:
datetime
LBL_CAPTIONS.Caption = ""
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
txt_addendum.Enabled = False
cmd_addendum.Visible = False
ADDENDUMFALG = False

cmb_machine.Visible = False

'cmb_finalby.clear 'Shraddha 05/11/08

'********************displaying designation in crystal report*********'
'******************added by milind on 03-03-06
str = "select usr_id from dislib.dislnkusr where username=?"
Set userid_qr = rdb.CreateQuery("userid_qr", str)

str = "select user_des from dislib.dsecurity where usr_id=?"
Set userdesc_qr = rdb.CreateQuery("userdesc_qr", str)

'*******************************'
'************upto here*********************

If ELREPORTFLAG = False Then
    ELREPORTFLAG = True


'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? AND MODMO='6'"
Set USRNAME_QR = rdb.CreateQuery("USRNAME1EL", str)
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

'----- Shraddha Pandit 14/11/08 -----
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11MICRO", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'------------------------------------

'*********SECURITY**********************************
''ADDED BY MONO FOR MALE FEMALE

'str = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET WHERE PRO_CODE=? AND SEX=?"
'Commented and added by Shraddha on 14/12/10
str = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND MACHINNA=?"
Set SELDTEMPLETQR = rdb.CreateQuery("SELDTEMPLETQRBCFMEL", str)

''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR1111", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR1111", str)


''********ICU***************

str = "SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SELBG_QR = rdb.CreateQuery("SELBG_QR", str)

str = "SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE=? AND AGE2=? AND SEX=? "
Set SELDBCOSMOQR = rdb.CreateQuery("SELDBCOSMOQR", str)

str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FED' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel1_deptqrTUMREP122", str)

'for retriving reqnos
Set req_qr = rdb.CreateQuery("1reqelrep", "select * from rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT desc ")



' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodalityE WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqr11elrep", str)

' for retriveing name,age etc
Set SELTA101_QR = rdb.CreateQuery("selta101_qr11elREP", "select * from PATOOLSLIB.ta101 where case_no =?")
  
  ' for bedno, ward retriveing
   str = "select * from PATOOLSLIB.ma104 where case_no =?"
   Set selma104_qr = rdb.CreateQuery("selma1041_qr11elREP", str)
   
  ''''******** interface.dinterface
  'STR = "SELECT * FROM INTERFACE.DINTERFACE WHERE REQNO=? AND TESTCODE=?"
  str = "SELECT * FROM INTERFACE.DINTERFACE WHERE REQNO=?"
  'Commented and added by Shraddha on 30/06/2011
  'str = "SELECT * FROM INTERFACE.DINTERFACE WHERE REQNO=?"
  str = "SELECT * FROM INTERFACE.DINTERFACE A, DISLIB.ELTEST B WHERE A.TESTCODE=B.TESTCODE AND REQNO=? ORDER BY B.SRNO"
  Set seldinterface_qr = rdb.CreateQuery("SELDINTERFACE", str)
  
' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR1_QR1elREP", str)
' FOR REFEFING DOCTOR
    str = "SELECT pro_code FROM DISLIB.DMULPROCE WHERE REQNO=? "
    Set DMULPROCE_QR = rdb.CreateQuery("DMULPROCE_QR1elREP", str)
    
    'str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
    'Commented ab=nd added by Shraddha on 30/12/10
    str = "SELECT *  FROM dislib.DMPROCEDUR A, DISLIB.ELTEST B WHERE A.PRO_CODE=B.TESTCODE AND reqno=? ORDER BY B.SRNO "
    Set DMPROCEDUR_QR = rdb.CreateQuery("DMPROCEDUR_QR1elREP", str)
    
'FOR INSERTING
    'str = "insert into  dislib.DMPROCEDUR(reqno,PRO_CODE,OBSERVALUE,obser_type,uom,nrange,pvalue) values (?,?,?,?,?,?,?)"
    'Commented ab=nd added by Shraddha on 30/12/10
    str = "insert into  dislib.DMPROCEDUR(reqno,PRO_CODE,OBSERVALUE,obser_type,uom,nrange,pvalue,remarks) values (?,?,?,?,?,?,?,?)"
    Set INSDPRO = rdb.CreateQuery("INSDPRO1elREPEL", str)
    
    
'FOR UPDATING
   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qrel", str)
   
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR11elREP", str)
    
    'str = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=? WHERE REQNO = ? AND PRO_CODE =?"
    'REMARKS
    'Commented and added by Shraddha on 30/12/10
    str = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=?, REMARKS=? WHERE REQNO = ? AND PRO_CODE =?"
    Set UPDDMPROCEDUR_QR = rdb.CreateQuery("UPDDMPROCEDUR_QR1elREP", str)

'FOR RETRIEVING THE PROCEDURE NAME
  ' STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE = ?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE = ?"
  Set SELMA101_QR = rdb.CreateQuery(SELMA101_QRelREP, str)
    
'FOR RETRIEVING THE PROCEDURE CODE
  ' STR = "SELECT * FROM DISLIB.MA101 WHERE SER_DET = ?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE = ?"
  Set SELMA101_QR1 = rdb.CreateQuery(SELMA101_QR1elREP, str)
  
'  '***** Added by Shraddha on 16/12/10 *****
'  str = "SELECT * FROM interface.dinterface where reqno=? and TESTCODE=?"
'  Set SEL_DINT_QR = rdb.CreateQuery("", str)
'
  str = "SELECT * FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND MACHINNA=?"
  Set SEL_DTEMPLET1_QR = rdb.CreateQuery("", str)
'  '*****************************************
   
  str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
  Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1QR11elREP", str)
   
  'str = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET WHERE PRO_CODE=?"
  'Commented and added by Shraddha on 14/12/10
'  str = "SELECT SEX,LOWVALUE,HIGHVALUE,UOM,LOWPANIC,HIGHPANIC FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=?"
'  Set SELDTEMPLETQR = rdb.CreateQuery("SELDTEMPLETQREL", str)
    
  str = "SELECT * FROM DISLIB.DELPROSUB WHERE BCPROCODE=?"
  Set SELDBCPROSUBQR = rdb.CreateQuery("SELDBCPROSUBQREL", str)
  
  str = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES(?,?)"
  Set INSDMBCCOMMQR = rdb.CreateQuery("INSDMBCCOMMQRBEL", str)
  
  str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
  Set UPDDMBCCOMMQR = rdb.CreateQuery("UPDDMBCCOMMQREL", str)
  
  str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
  Set SELDMBCCOMMQR = rdb.CreateQuery("SELDMBCCOMMQREL", str)
  
  str = "SELECT * FROM DISLIB.DELPROSUB WHERE PROCODE=?"
  Set SELELPROSUBQR = rdb.CreateQuery("SELELPROSUBQREL", str)
  
    str = "select * from dislib.dissec where enterby='YES'AND MODALITY='EL' "
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1el", str)
   

   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='EL'"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1el", str)
   
 str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrEL", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
  Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1el", str)
    
  
  'ADDENDUM
  str = "UPDATE DISLIB.ADDENDUMCY SET ADDENDNO=?,ADDEND_DT=?,ADDEND_TM=?,USER_ID=?,ADD_MAX_LN=? WHERE REQNO=?"
  Set updateaddendum_qr = rdb.CreateQuery("UPDATEADDENDUMel", str)
  
  str = "INSERT INTO DISLIB.ADDENDUMCY VALUES(?,?,?,?,?,?)"
  Set insertaddendum_qr = rdb.CreateQuery("INSERTADDENDUMel", str)
  
  str = "SELECT * FROM DISLIB.ADDENDUMCY WHERE REQNO=?"
  Set selectaddendum_qr = rdb.CreateQuery("SELECTADDENDel", str)
  
  'THIS IS RTIMLOG
 
 str = "update rislib.rtimelog set ADPREPDATE=?,ADPREPTIME=? where reqnno=?"
 Set UPDRTIMECADD_QR = rdb.CreateQuery("updatetimecadd1el", str)
 
 str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=? WHERE REQNNO=?"
 Set UPDRTIME_QR = rdb.CreateQuery("QWEQWEQWEel", str)

 str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEcel", str)
 
 
 str = "UPDATE RISLIB.RTIMELOG SET PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEP_QR = rdb.CreateQuery("QWEQWEQWEPel", str)

 str = "UPDATE RISLIB.RTIMELOG SET MODREPDATE=?,MODREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEM_QR = rdb.CreateQuery("QWEQWEQWEMel", str)
    
       
End If
btnenabl ("INI")
FLDENABL ("DIS")

MSFlexGrid1.col = 2
txt_observed.Width = MSFlexGrid1.CellWidth
txt_observed.Height = MSFlexGrid1.CellHeight
txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

Exit Sub
MSG:
MsgBox Err.Description
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
        cmb_modality.ListIndex = 0
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
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "FED" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub
Public Sub clear()
lblicu.Visible = False
lbl_rwd.Visible = False
LBL_ABG.Visible = False
LBL_CBG.Visible = False
LBL_VBG.Visible = False
cmb_dept.clear
cmb_modality.clear
cmb_reqno.clear
txt_age.Text = " "
'txt_amtchg.Text = " "
txt_bedno.Text = " "
MaskEdBox1.Text = " "
txt_category.Text = " "
txt_date.Text = " "
txt_deptcode.Text = " "
'txt_Location.Text = " "
txt_modcode.Text = " "
txt_name.Text = " "
txt_pathno.Text = " "
txt_pridoc.Text = " "
'txt_recpt.Text = " "
txt_refdoc.Text = " "
txt_sex.Text = " "
txt_status.Text = " "
txt_time.Text = " "
txt_ward.Text = " "
txt_comments.Text = ""
txt_addendum.Text = ""
cmb_finalby.clear
'cmb_enterby.Text = ""
'cmb_finalby.Text = ""
'For i = 1 To MSFlexGrid1.Row
'For J = 0 To MSFlexGrid1.Col
'MSFlexGrid1.TextMatrix(i, J) = " "
'Next J
'Next i
MSFlexGrid1.Rows = 1
txt_observed.Text = " "
NEWFLAG = False
viewflag = False
End Sub

Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
        MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
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
              MsgBox "NO REQUISITION IS PENDING IN NEW FOR THIS CASE NO"
              End If
     End If
     
              If viewflag = True Then
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
    MsgBox Err.Description

End Sub

Private Sub MSFlexGrid1LeaveCell()
 MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
End Sub

Private Sub Timer1_Timer()
'txt_time.Text = Format(Now, "hh:mm:ss")
End Sub

Private Sub TXT_COMMENTS_GotFocus()
txt_comments.clear
txt_comments.AddItem "#Rechecked"

txt_comments.AddItem "#No Sample"
txt_comments.AddItem "#Sample ? Pl. repeat the sample."
txt_comments.AddItem "#Machine not working"
txt_comments.AddItem "#Serum haemolysed, please repeat the sample."
txt_comments.AddItem "#Serum highly haemolysed, please repeat the sample."
txt_comments.AddItem "#Serum lipaemic"
txt_comments.AddItem "#Serum highly lipaemic"
'***** Commented by Shraddha on 24/07/10 *****
'TXT_COMMENTS.AddItem "#Report from Biochemistry"
'TXT_COMMENTS.AddItem "#O.R.- out of range(above 2 minutes);PTC=   ,PTTC=  ,Pl. repeat the sample"
'TXT_COMMENTS.AddItem "#At  00:00   Am/Pm  "
'TXT_COMMENTS.AddItem "#Report Attached."
'*********************************************
txt_comments.AddItem "#Clotted Sample, please repeat the sample."
txt_comments.AddItem "#Wrong Requisition"
txt_comments.AddItem "#PTC=  sec,PTTC=  sec."
txt_comments.AddItem "#Out of Range (High/Low)sample ?,Please repeat the sample"

'TXT_COMMENTS.AddItem "Electrolytes reported from E-Lab"
txt_comments.AddItem "#Second Sample Received at - "

txt_comments.AddItem "#Analysed on AU400"
txt_comments.AddItem "#Analysed on DYTONA"

End Sub

Private Sub txt_observed_KeyPress(KeyAscii As Integer)

Dim SEX As String

If KeyAscii = 13 Then

    num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
    sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)
    
    If SELMA101FLAG1 = False Then
        SELMA101_QR1(0) = sc 'Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
        Set SELMA101_RS1 = SELMA101_QR1.OpenResultset(1, 3)
        SELMA101FLAG1 = True
    Else
        SELMA101_QR1(0) = sc 'Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
        SELMA101_RS1.MoreResults
        SELMA101_RS1.Requery
    End If
    
    If SELMA101_RS1.RowCount > 0 Then
        If SELDTEMPLETFLAG = False Then
            SELDTEMPLETQR(0) = Trim(SELMA101_RS1("SER_CODE"))
            SELDTEMPLETQR(1) = "DYTON"
            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
            SELDTEMPLETFLAG = True
        Else
            SELDTEMPLETQR(0) = Trim(SELMA101_RS1("SER_CODE"))
            SELDTEMPLETRS.MoreResults
            SELDTEMPLETRS.Requery
        End If
        
        'for alkaline phosphate changes in range according to age
        SEX = Trim(txt_sex.Text)
        If Trim(SELMA101_RS1("SER_CODE")) = "F957" Then
           If Trim(txt_age.Text) <= 15 Then
              SEX = "C"
           Else
              SEX = "A"
           End If
        End If
                
        If SELDTEMPLETRS.RowCount > 0 Then
            If Trim(SELDTEMPLETRS("sex")) = SEX Or Trim(SELDTEMPLETRS("sex")) = "" Or IsNull(SELDTEMPLETRS("sex")) Then
            Else
            
            If SELDTEMPLETRS.RowCount > 1 Then 'SHWETA 10-08-2012
            
            
                SELDTEMPLETRS.MoveNext
                End If
            End If
            
            low = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 1, 3)
            high = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 5, 7)
            
            '''''''' BY SHRIKANT ''''''''''''''''''
            If Trim(SELMA101_RS1("SER_CODE")) = "F955" Then
                    Dim SS As Variant
                    Dim v As Integer
                    SS = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 2)) - Val(txt_observed)
                    v = InStr(SS, ".")
                    If v = 0 Then
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row + 1, 2) = SS & ".0"
                        'txt_observed = SS & ".0"
                    Else
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row + 1, 2) = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 2)) - Val(txt_observed)
                        'txt_observed = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 2)) - Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
                    End If
            End If
            ''''''''''''''''''''''''''''''''''''''''''
            
            If SELMA101_RS1("SER_CODE") = "F998" Then
                If Val(Trim(txt_observed.Text)) > high Then
                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                    If (Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC")))) And Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) <> 0 Then
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                    Else
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                    End If
                    For i = 1 To 1000
                    Beep
                    Next i
                Else
                    If Val(Trim(txt_observed.Text)) >= low Then
                        MSFlexGrid1.CellBackColor = vbWhite
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                        If (Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC")))) And Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) <> 0 Then
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                        Else
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                        End If
                    Else
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = ""
                    End If
                    
                    If Val(Trim(txt_observed.Text)) < low Then
                        MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) >= Val(Trim(txt_observed.Text)) Then
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                        Else
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                        End If
                    End If
                    If Trim(txt_observed.Text) = "" Then
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                    End If
                End If
            Else
                    If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                        
                            MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                            If (Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC")))) And Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) <> 0 Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                            Else
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                            End If
                            
                            If Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) = 0 Then
                                MSFlexGrid1.CellBackColor = vbWhite
                            End If
                            
'                        For i = 1 To 1000
'                            Beep
'                        Next i
                    Else
                         If Val(Trim(txt_observed.Text)) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                
                                MSFlexGrid1.CellBackColor = vbWhite
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                                
                                If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) >= Val(Trim(txt_observed.Text)) And Val(Trim(txt_observed.Text)) <> -0 Then
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                                Else
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                                End If
                                If Val(Trim(txt_observed.Text)) < 0 Then
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                                End If
                                
                            
                         Else
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = ""
                         End If
                         If Val(Trim(txt_observed.Text)) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                            
                                MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                                If (Val(Trim(SELDTEMPLETRS("LOWPANIC"))) >= Val(Trim(txt_observed.Text))) And (Val(Trim(SELDTEMPLETRS("LOWVALUE"))) And Val(Trim(SELDTEMPLETRS("HIGHVALUE")))) Then
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = txt_observed.Text
                                Else
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
                                End If
                             
                                If Val(Trim(SELDTEMPLETRS("LOWVALUE"))) = 0 And Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) = 0 Then
                                    MSFlexGrid1.CellBackColor = vbWhite
                                End If
                                                                                                 
                         End If
                         If Trim(txt_observed.Text) = "" Then
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
                         End If
                    End If
            End If
            End If
    End If

    i = MSFlexGrid1.Rows
    If i = MSFlexGrid1.row Then
        cmd_save.SetFocus
        Exit Sub
    End If

    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
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
'    If SELMA101FLAG1 = False Then
'        SELMA101_QR1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
'        Set SELMA101_RS1 = SELMA101_QR1.OpenResultset(1, 3)
'        SELMA101FLAG1 = True
'    Else
'        SELMA101_QR1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
'        SELMA101_RS1.MoreResults
'        SELMA101_RS1.Requery
'    End If
'    If SELMA101_RS1.RowCount > 0 Then
'        If SELDTEMPLETFLAG = False Then
'            SELDTEMPLETQR(0) = Trim(SELMA101_RS1("SER_CODE"))
'            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'            SELDTEMPLETFLAG = True
'        Else
'            SELDTEMPLETQR(0) = Trim(SELMA101_RS1("SER_CODE"))
'            SELDTEMPLETRS.MoreResults
'            SELDTEMPLETRS.Requery
'        End If
'
'        If SELDTEMPLETRS.RowCount > 0 Then
'            If Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
'                MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'                If (Val(Trim(txt_observed.Text)) > Val(Trim(SELDTEMPLETRS("HighPANIC")))) And Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) <> 0 Then
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
'                 Else
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
'                 End If
'                 If Val(Trim(txt_observed.Text)) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
'                    MSFlexGrid1.CellBackColor = RGB(150, 255, 200)
'                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'                        If Val(Trim(SELDTEMPLETRS("LOWPANIC"))) >= Val(Trim(txt_observed.Text)) Then
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
'    i = MSFlexGrid1.Rows
'    If i = MSFlexGrid1.Row Then
'        cmd_save.SetFocus
'        Exit Sub
'    End If
'
'
'
' MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
'    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
'            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
'    End If
'    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'            txt_observed.Text = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'
'            txt_observed.Visible = True
'            txt_observed.setfocus
'
'            txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'
'    End If
End Sub


Public Sub STAFF()
'On Error GoTo msg:
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
                  lbl_rwd.Visible = False
               ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                  lbl_rwd.Visible = True
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
            If selpathdrequisitnrsflag = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(cmb_reqno.Text)
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                selpathdrequisitnrsflag = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(cmb_reqno.Text)
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = Trim(selpathdrequisitn_rs("pathno"))
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
  
  
'    If selma104_rs.RowCount > 0 Then
'        txt_ward.Text = selma104_rs("WARD_NAME")
'        txt_bedno.Text = selma104_rs("bed_no")
'        txt_status.Text = "IP"
'    Else
'        txt_status.Text = "OP"
'    End If
    
    '*****
    
    
            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
        
                       
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F995" Then
                LBL_VBG.Visible = True
            Else
                LBL_VBG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F901" Then
                LBL_ABG.Visible = True
            Else
                LBL_ABG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F902" Then
                LBL_CBG.Visible = True
            Else
                LBL_CBG.Visible = False
            End If

            '***** Shraddha 20/12/10 *****
            
''            If DMULPROCE_RS.RowCount > 0 Then
''                I = 1
''                MSFlexGrid1.Rows = 1
''
''                ''**********NEW***********************
''                While Not DMULPROCE_RS.EOF
''                    'If DMULPROCE_RS("PRO_CODE") = "F860" Or DMULPROCE_RS("PRO_CODE") = "F875" Or DMULPROCE_RS("PRO_CODE") = "F901" Or DMULPROCE_RS("PRO_CODE") = "F903" Or DMULPROCE_RS("PRO_CODE") = "F908" Then
''                    If DMULPROCE_RS("PRO_CODE") = "F860" Or DMULPROCE_RS("PRO_CODE") = "F875" Or DMULPROCE_RS("PRO_CODE") = "F903" Or DMULPROCE_RS("PRO_CODE") = "F908" Or DMULPROCE_RS("PRO_CODE") = "F865" Or DMULPROCE_RS("PRO_CODE") = "F981" Or DMULPROCE_RS("PRO_CODE") = "F901" Or DMULPROCE_RS("PRO_CODE") = "F902" Or DMULPROCE_RS("PRO_CODE") = "F995" Or DMULPROCE_RS("PRO_CODE") = "F022" Then
''                    ''   F901 IS EXCLUDED FROM THE IF AS IT REFERS TO ARTERIAL BLOOD GASES WHICH IS NOT A GROUP SERVICS
''
''                        If SELELPROSUBFLAG = False Then
''                            SELELPROSUBQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
''                            Set SELELPROSUBRS = SELELPROSUBQR.OpenResultset(1, 3)
''                            SELELPROSUBFLAG = True
''                        Else
''                            SELELPROSUBQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
''                            SELELPROSUBRS.MoreResults
''                            SELELPROSUBRS.Requery
''                        End If
''                        If SELELPROSUBRS.RowCount > 0 Then
''                            While Not SELELPROSUBRS.EOF
''
''                                If SELMA101FLAG = False Then
''                                    SELMA101_QR(0) = Trim(SELELPROSUBRS("PROCODESUB"))
''                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                                    SELMA101FLAG = True
''                                Else
''                                    SELMA101_QR(0) = Trim(SELELPROSUBRS("PROCODESUB"))
''                                    SELMA101_RS.MoreResults
''                                    SELMA101_RS.Requery
''                                End If
''                                 If SELDTEMPLETFLAG = False Then
''                                    SELDTEMPLETQR(0) = Trim(SELELPROSUBRS("PROCODESUB"))
''                                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                                    SELDTEMPLETFLAG = True
''                                Else
''                                    SELDTEMPLETQR(0) = SELELPROSUBRS("PROCODESUB")
''                                    SELDTEMPLETRS.MoreResults
''                                    SELDTEMPLETRS.Requery
''                                End If
''                                If SELDTEMPLETRS.RowCount > 0 Then
''                                    If SELMA101_RS.RowCount > 0 Then
''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Then
''                                            MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        Else
''                                            SELDTEMPLETRS.MoveNext
''                                            MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        End If
''                                        I = I + 1
''                                    End If
''                                End If
''
''                                SELELPROSUBRS.MoveNext
''                            Wend
''                        End If
''
''
''
''                Else
''
''                        If SELMA101FLAG = False Then
''                            SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
''                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                            SELMA101FLAG = True
''                        Else
''                            SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
''                            SELMA101_RS.MoreResults
''                            SELMA101_RS.Requery
''                        End If
''                         If SELDTEMPLETFLAG = False Then
''                            SELDTEMPLETQR(0) = DMULPROCE_RS("PRO_CODE")
''                            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
''                            SELDTEMPLETFLAG = True
''                        Else
''                            SELDTEMPLETQR(0) = DMULPROCE_RS("PRO_CODE")
''                            SELDTEMPLETRS.MoreResults
''                            SELDTEMPLETRS.Requery
''                        End If
''
''                        If SELDTEMPLETRS.RowCount > 0 Then
''
''                            If DMULPROCE_RS("PRO_CODE") = "F998" Then
''
''                                OSMOAGE
''
''                                If SELDBCOSMOFLAG = False Then
''                                        SELDBCOSMOQR(0) = DMULPROCE_RS("PRO_CODE")
''                                        SELDBCOSMOQR(1) = RNAGE
''                                        SELDBCOSMOQR(2) = SS
''                                        Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
''                                        SELDBCOSMOFLAG = True
''                                Else
''                                        SELDBCOSMOQR(0) = DMULPROCE_RS("PRO_CODE")
''                                        SELDBCOSMOQR(1) = RNAGE
''                                        SELDBCOSMOQR(2) = SS
''                                        SELDBCOSMORS.MoreResults
''                                        SELDBCOSMORS.Requery
''                                End If
''
''                                RNAGE1 = SELDBCOSMORS!Range
''
''                                OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
''                                OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
''
''
''                                    If SELMA101_RS.RowCount > 0 Then
''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
''                                            MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & " " & vbTab & RNAGE1
''                                        Else
''                                            SELDTEMPLETRS.MoveNext
''                                            MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & " " & vbTab & RNAGE1
''                                        End If
''                                        I = I + 1
''                                    End If
''                            Else
''                                    If SELMA101_RS.RowCount > 0 Then
''                                        If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(Trim(SELDTEMPLETRS("SEX"))) Then
''                                            MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        Else
''                                            SELDTEMPLETRS.MoveNext
''                                            MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(SELMA101_RS(0)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
''                                        End If
''                                        I = I + 1
''                                    End If
''                            End If
''
''
''                        End If
''                     End If
''                    DMULPROCE_RS.MoveNext
''                   Wend
''
''                ''*********NEW************************
''
''
''
''
''            End If
'**** Commented on 20/12/10 *****
SELECTDINTERFACE
            MSFlexGrid1.Rows = seldinterface_rs.RowCount + 1
            'MSFlexGrid1.Cols = 2
            txt_observed.Visible = True
            

    
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
                   lbl_rwd.Visible = True
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
                If SELDREQSTAFFD1RS1("RELATION") <> "STAFF" Then
                    txt_status = "DP"
                Else
                    txt_status = "ST"
                End If
        End If
        
            If selpathdrequisitnrsflag = False Then
                SELPATHDREQUISITN_QR(0) = req_chglen(cmb_reqno.Text)
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                selpathdrequisitnrsflag = True
            Else
                SELPATHDREQUISITN_QR(0) = req_chglen(cmb_reqno.Text)
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = Trim(selpathdrequisitn_rs("pathno"))
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
            If REFDR_RS.EOF = False Then
            cmb_enterby.Text = " " & Trim(REFDR_RS("enterby"))
            End If
            If REFDR_RS.EOF = False Then
                cmb_finalby.Text = " " & Trim(REFDR_RS("freportby"))
            End If
            
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
  
  
'        If selma104_rs.RowCount > 0 Then
'            txt_ward.Text = selma104_rs("WARD_NAME")
'            txt_bedno.Text = selma104_rs("bed_no")
'            txt_status.Text = "IP"
'        Else
'            txt_status.Text = "OP"
'        End If
        '*******
        
            If DMPROCEDURFLAG = False Then
                 DMPROCEDUR_QR(0) = req_chglen(cmb_reqno.Text)
                Set DMPROCEDUR_RS = DMPROCEDUR_QR.OpenResultset(1, 3)
                DMPROCEDURFLAG = True
            Else
                DMPROCEDUR_QR(0) = req_chglen(cmb_reqno.Text)
                DMPROCEDUR_RS.MoreResults
                DMPROCEDUR_RS.Requery
            End If
            
            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F995" Then
                LBL_VBG.Visible = True
            Else
                LBL_VBG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F901" Then
                LBL_ABG.Visible = True
            Else
                LBL_ABG.Visible = False
            End If
            
            If Trim(DMULPROCE_RS("PRO_CODE")) = "F902" Then
                LBL_CBG.Visible = True
            Else
                LBL_CBG.Visible = False
            End If

            If DMPROCEDUR_RS.RowCount > 0 Then
                i = 1
               ' MSFlexGrid1.Rows = 1
                'While Not DMPROCEDUR_RS.EOF
'                    If SELMA101FLAG = False Then
'                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
'                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                        SELMA101FLAG = True
'                    Else
'                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
'                        SELMA101_RS.MoreResults
'                        SELMA101_RS.Requery
'                    End If
'                    If SELDTEMPLETFLAG = False Then
'                        SELDTEMPLETQR(0) = DMPROCEDUR_RS("PRO_CODE")
'                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                        SELDTEMPLETFLAG = True
'                    Else
'                        SELDTEMPLETQR(0) = DMPROCEDUR_RS("PRO_CODE")
'                        SELDTEMPLETRS.MoreResults
'                        SELDTEMPLETRS.Requery
'                    End If
'                    If SELDTEMPLETRS.RowCount > 0 Then
'                        If SELMA101_RS.RowCount > 0 Then
'                            If SELDTEMPLETRS("sex") = "M" Or Trim(SELDTEMPLETRS("SEX")) = "" Then
'                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                            Else
'                               SELDTEMPLETRS.MoveNext
'                               MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                            End If
'                            i = i + 1
'                        End If
'                    End If
' ***** COLOR CODING

                    i = 1
                    srno = 1
                While Not DMPROCEDUR_RS.EOF
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = DMPROCEDUR_RS("PRO_CODE")
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = DMPROCEDUR_RS("PRO_CODE")
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    seldinterface_qr(0) = VAR_REQNNO
                    Set seldinterface_rs = seldinterface_qr.OpenResultset(1, 3)
                    If SELDTEMPLETFLAG = False Then
                        SELDTEMPLETQR(0) = DMPROCEDUR_RS("PRO_CODE")
                        SELDTEMPLETQR(1) = "" 'Trim(seldinterface_rs("machineno"))
                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                        SELDTEMPLETFLAG = True
                    Else
                        SELDTEMPLETQR(0) = DMPROCEDUR_RS("PRO_CODE")
                        SELDTEMPLETRS.MoreResults
                        SELDTEMPLETRS.Requery
                    End If
                    
                    If SELDTEMPLETRS.RowCount > 0 Then
                            If SELMA101_RS.RowCount > 0 Then
                            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                             MSFlexGrid1.row = srno
                             MSFlexGrid1.col = 0
                            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
                            If SELDTEMPLETRS("SEX") = "M" Or Trim(SELDTEMPLETRS("SEX")) = "" Then
                                MSFlexGrid1.TextMatrix(i, 0) = i
                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101_RS("SER_DET"))
                                If Trim(DMPROCEDUR_RS("obser_type")) = "H" Then
                                    MSFlexGrid1.row = srno
                                     MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                                Else
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "L" Then
                                      
                                      MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = RGB(100, 255, 150)
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "N" Then
                                     
                                       MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "" Then
                                      
                                        MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                End If
                                j = 0
                                j = InStr(DMPROCEDUR_RS("OBSERVALUE"), "*")
                                If j = 0 Then
                                    MSFlexGrid1.TextMatrix(i, 2) = DMPROCEDUR_RS("OBSERVALUE")
                                Else
                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                    MSFlexGrid1.TextMatrix(i, 4) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                End If
                                MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                            Else
                                SELDTEMPLETRS.MoveNext
                                MSFlexGrid1.TextMatrix(i, 0) = i
                                MSFlexGrid1.TextMatrix(i, 1) = Trim(SELMA101_RS("SER_DET"))
                                If Trim(DMPROCEDUR_RS("obser_type")) = "H" Then
                                    MSFlexGrid1.row = srno
                                     MSFlexGrid1.col = 2
                                    MSFlexGrid1.CellBackColor = RGB(255, 100, 100)
                                Else
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "L" Then
                                      
                                      MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = RGB(150, 255, 100)
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "N" Then
                                     
                                       MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                    If Trim(DMPROCEDUR_RS("obser_type")) = "" Then
                                      
                                        MSFlexGrid1.col = 2
                                      MSFlexGrid1.CellBackColor = vbWhite
                                    End If
                                End If
                                j = 0
                                j = InStr(DMPROCEDUR_RS("OBSERVALUE"), "*")
                                If j = 0 Then
                                    MSFlexGrid1.TextMatrix(i, 2) = DMPROCEDUR_RS("OBSERVALUE")
                                Else
                                    MSFlexGrid1.TextMatrix(i, 2) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                    MSFlexGrid1.TextMatrix(i, 4) = Mid(DMPROCEDUR_RS("OBSERVALUE"), 1, j - 1)
                                End If
                                    SELDTEMPLETRS.MoveFirst
                                    MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                If DMULPROCE_RS("PRO_CODE") = "F998" Then
                                
                                    OSMOAGE
        
                                    If SELDBCOSMOFLAG = False Then
                                            SELDBCOSMOQR(0) = DMULPROCE_RS("PRO_CODE")
                                            SELDBCOSMOQR(1) = RNAGE
                                            SELDBCOSMOQR(2) = SS
                                            Set SELDBCOSMORS = SELDBCOSMOQR.OpenResultset(1, 3)
                                            SELDBCOSMOFLAG = True
                                    Else
                                            SELDBCOSMOQR(0) = DMULPROCE_RS("PRO_CODE")
                                            SELDBCOSMOQR(1) = RNAGE
                                            SELDBCOSMOQR(2) = SS
                                            SELDBCOSMORS.MoreResults
                                            SELDBCOSMORS.Requery
                                    End If
            
                                    RNAGE1 = SELDBCOSMORS!range
        
                                    OSMOLV = Left(Trim(SELDBCOSMORS("RANGE")), 3)
                                    OSMOHV = Right(Trim(SELDBCOSMORS("RANGE")), 3)
                                End If

                                If Trim(DMPROCEDUR_RS("PRO_CODE")) = "F998" Then
                                    MSFlexGrid1.TextMatrix(i, 3) = OSMOLV & "-" & OSMOHV
                                Else
                                    MSFlexGrid1.TextMatrix(i, 3) = Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                                End If

                            End If
                        End If
                        
                        '***** Shraddha 30/12/10 *****
                        If Trim(DMPROCEDUR_RS("remarks")) <> "" Or IsNull(DMPROCEDUR_RS("remarks")) = False Then
                            MSFlexGrid1.TextMatrix(i, 5) = Trim(DMPROCEDUR_RS("remarks"))
                        End If
                        '*****************************
                        
                        
                           srno = srno + 1
                            i = i + 1

                    End If
                    DMPROCEDUR_RS.MoveNext

                Wend
'                MSFlexGrid1.Row = 1
'                MSFlexGrid1.Col = 2
'                txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                txt_observed.Visible = False




'***** COLOR CODING
'                    DMPROCEDUR_RS.MoveNext
'                Wend
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
                txt_comments.Text = " " & SELDMBCCOMMRS("COMMENTS").Value
            End If
      End If
        
        
        '*******
           
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
    i = 1
    While i < MSFlexGrid1.Rows
    
        If Trim(DMPROCEDUR_RS("PRO_CODE")) = Right(Trim(MSFlexGrid1.TextMatrix(i, 1)), 4) Then
            'MsgBox "match"
            MSFlexGrid1.TextMatrix(i, 6) = Trim(DMPROCEDUR_RS("OBSERVALUE"))
        End If
        i = i + 1
    Wend
    DMPROCEDUR_RS.MoveNext
Wend

End If
End Sub

Public Sub assign()
'On Error GoTo MSG:
rdb.BeginTrans
If flagsave = 1 Then
        '**** inserting into dmprocedur
        'MSFlexGrid1.row = 1
        
        'num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
        'sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)
        
        For i = 1 To MSFlexGrid1.Rows - 1
                
                num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
                sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
                
                If SELMA101FLAG1 = False Then
                    SELMA101_QR1(0) = sc 'Trim(MSFlexGrid1.TextMatrix(i, 1))
                    Set SELMA101_RS1 = SELMA101_QR1.OpenResultset(1, 3)
                    SELMA101FLAG1 = True
                Else
                    SELMA101_QR1(0) = sc 'Trim(MSFlexGrid1.TextMatrix(i, 1))
                    SELMA101_RS1.MoreResults
                    SELMA101_RS1.Requery
                    
                End If
                If SELMA101_RS1.RowCount > 0 Then
                    INSDPRO(0) = req_chglen(cmb_reqno.Text)
                    INSDPRO(1) = SELMA101_RS1("SER_CODE")
'                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                
                    If SELDTEMPLETFLAG = False Then
                        SELDTEMPLETQR(0) = SELMA101_RS1("ser_code")
                        Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                        
                    Else
                        SELDTEMPLETQR(0) = SELMA101_RS1("ser_code")
                        'SELDTEMPLETQR(1)=
                        SELDTEMPLETRS.MoreResults
                        SELDTEMPLETRS.Requery
                        
                    End If
                    
                'for alkaline phosphate changes in range according to age
                SEX = Trim(txt_sex.Text)
                If Trim(SELMA101_RS1("ser_code")) = "F957" Then
                   If Trim(txt_age.Text) <= 15 Then
                      SEX = "C"
                   Else
                      SEX = "A"
                   End If
                End If
                
                
                'If Trim(sc) = "F869" Then
                If Trim(SELMA101_RS1("ser_code")) = "F998" Then
                
                '***** Shraddha 07/10/2011 *****
                SEL_DTEMPLET1_QR(0) = "F998"
                SEL_DTEMPLET1_QR(1) = MSFlexGrid1.TextMatrix(i, 5)
                Set SEL_DTEMPLET1_RS = SEL_DTEMPLET1_QR.OpenResultset(1, 3)
                '*******************************

                            'If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(OSMOHV) Then
                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(SEL_DTEMPLET1_RS("HIGHVALUE")) Then
                                INSDPRO(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                INSDPRO(3) = "H"

                            Else
                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(SEL_DTEMPLET1_RS("LOWVALUE")) Then
                                    INSDPRO(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                    INSDPRO(3) = "N"

                                End If
                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(SEL_DTEMPLET1_RS("LOWVALUE")) Then
                                        INSDPRO(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                        INSDPRO(3) = "L"

                                End If
                                If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
                                    INSDPRO(2) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                    INSDPRO(3) = ""

                                End If
                            End If

                Else
                         If SELDTEMPLETRS.RowCount > 0 Then
                             If Trim(SELDTEMPLETRS("SEX")) = SEX Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(SELDTEMPLETRS("SEX")) Then
                             Else
                                If SELDTEMPLETRS.RowCount > 1 Then 'SHWETA 10-08-2012
                               
                                 SELDTEMPLETRS.MoveNext
                                 End If
                             End If
                             If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                                    
                                 If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("highpanic"))) Then
                                     INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2) & "*"
                                 Else
                                     INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                                 End If
                                     INSDPRO(3) = "H"
                                 If Val(Trim(SELDTEMPLETRS("LOWVALUE"))) = 0 And Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) = 0 Then
                                     INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                                     INSDPRO(3) = "N"
                                 End If
                             Else
                                 If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                     INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                                     INSDPRO(3) = "N"
                                 End If
                                  If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                     If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWPANIC"))) Then
                                         INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2) & "*"
                                     Else
                                         INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                                     End If
                                     INSDPRO(3) = "L"
                                  End If
                                  
                                  If Val(Trim(SELDTEMPLETRS("LOWVALUE"))) = 0 And Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) = 0 Then
                                     INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                                     INSDPRO(3) = "N"
                                  End If
                                  
                                  If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
                                     INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                                     INSDPRO(3) = ""
                                  End If
                             End If
                        Else
                        'changes done by shweta 23-03-2013
                        
                            INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2) ' ""
                            INSDPRO(3) = ""
                         End If
'                             INSDPRO(4) = Trim(SELDTEMPLETRS("UOM"))
'                             INSDPRO(5) = Trim(MSFlexGrid1.TextMatrix(i, 3))
'                             INSDPRO(6) = Trim(MSFlexGrid1.TextMatrix(i, 4))
'                             INSDPRO.Execute
                        End If
                             If SELDTEMPLETRS.EOF Then
                                INSDPRO(4) = ""
                             Else
                                INSDPRO(4) = Trim(SELDTEMPLETRS("UOM"))
                             End If
                             INSDPRO(5) = Trim(MSFlexGrid1.TextMatrix(i, 3))
                             INSDPRO(6) = Trim(MSFlexGrid1.TextMatrix(i, 4))
                             INSDPRO(7) = Trim(MSFlexGrid1.TextMatrix(i, 5)) 'Shraddha 30/12/10
                             INSDPRO.Execute
                        
                End If
        Next i
        '**** inserting into dmprocedur
        '**** UPDATE THE RREQUISITN
                 updrrequisitn_qr(0) = cmb_enterby.Text
                 updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                 updrrequisitn_qr(2) = UCase(cmb_finalby.Text)
                 updrrequisitn_qr(3) = "T"
                 updrrequisitn_qr(4) = UCase(cmb_finalby.Text)
                 updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                 updrrequisitn_qr(6) = req_chglen(cmb_reqno.Text)
                 updrrequisitn_qr.Execute
       '**** UPDATE THE RREQUISITN
       
       '**** INSERTING INTO COMMENTS
               INSDMBCCOMMQR(0) = req_chglen(cmb_reqno.Text)
               If Len(Trim(txt_comments.Text)) > 0 Then
                INSDMBCCOMMQR(1) = Null
                INSDMBCCOMMQR(1) = txt_comments
               Else
                INSDMBCCOMMQR(1) = Null
               End If
               INSDMBCCOMMQR.Execute
       
       '**** INSERTING INTO COMMENTS
       
       
       'RTIME LOG 18-01-05
        UPDRTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDRTIME_QR(1) = txt_time
        UPDRTIME_QR(2) = req_chglen(cmb_reqno.Text)
        UPDRTIME_QR.Execute
       
    
End If
If flagsave = 2 Then
        
         'num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
         'sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)
               
         For i = 1 To MSFlexGrid1.Rows - 1
'                UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
            
            num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
            sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
            
            If SELMA101FLAG1 = False Then
                SELMA101_QR1(0) = sc 'MSFlexGrid1.TextMatrix(i, 1)
                Set SELMA101_RS1 = SELMA101_QR1.OpenResultset(1, 3)
                SELMA101FLAG1 = True
            Else
                SELMA101_QR1(0) = sc 'MSFlexGrid1.TextMatrix(i, 1)
                SELMA101_RS1.MoreResults
                SELMA101_RS1.Requery
            End If
            If SELMA101_RS1.RowCount > 0 Then
                If SELDTEMPLETFLAG = False Then
                    SELDTEMPLETQR(0) = SELMA101_RS1("SER_CODE")
                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                    SELDTEMPLETFLAG = True
                Else
                    SELDTEMPLETQR(0) = SELMA101_RS1("SER_CODE")
                    SELDTEMPLETRS.MoreResults
                    SELDTEMPLETRS.Requery
                End If
                
                If Trim(SELMA101_RS1("ser_code")) = "F998" Then

                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(OSMOHV) Then
                                UPDDMPROCEDUR_QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                UPDDMPROCEDUR_QR(1) = "H"

                            Else
                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(OSMOLV) Then
                                    UPDDMPROCEDUR_QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                    UPDDMPROCEDUR_QR(1) = "N"

                                End If
                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(OSMOLV) Then
                                        UPDDMPROCEDUR_QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                        UPDDMPROCEDUR_QR(1) = "L"

                                End If
                                If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
                                    UPDDMPROCEDUR_QR(0) = Trim(MSFlexGrid1.TextMatrix(i, 2))
                                    UPDDMPROCEDUR_QR(1) = ""

                                End If
                            End If

                Else
                    If SELDTEMPLETRS.RowCount > 0 Then
                             If Trim(SELDTEMPLETRS("SEX")) = SEX Or Trim(SELDTEMPLETRS("SEX")) = "" Or IsNull(SELDTEMPLETRS("SEX")) Then
                             Else
                             
                                If SELDTEMPLETRS.RowCount > 1 Then
                                 SELDTEMPLETRS.MoveNext
                                 End If
                                
                              End If
                        
                        
                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) Then
                            
                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS("HIGHPANIC"))) Then
                                UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2) & "*"
                            Else
                                UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                            End If
                            UPDDMPROCEDUR_QR(1) = "H"
                            
                            If Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) = 0 Then
                                UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                UPDDMPROCEDUR_QR(1) = "N"
                            End If
                            
                        Else
                        
'                           If SELDTEMPLETRS("SEX") = Trim(txt_sex.Text) Then
'                           Else
'                                SELDTEMPLETRS.MoveNext
'                           End If ''''KANCHAN
                            
                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                UPDDMPROCEDUR_QR(1) = "N"
                            End If
                            
                            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWVALUE"))) Then
                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS("LOWPANIC"))) Then
                                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2) & "*"
                                Else
                                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                End If
                                 UPDDMPROCEDUR_QR(1) = "L"
                                
                                If Val(Trim(SELDTEMPLETRS("LOWVALUE"))) = 0 And Val(Trim(SELDTEMPLETRS("HIGHVALUE"))) = 0 Then
                                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                    UPDDMPROCEDUR_QR(1) = "N"
                                End If
                              
                            End If
                            
                            If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
                                 UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                 UPDDMPROCEDUR_QR(1) = ""
                            End If
                        End If
                        
                    End If
                End If
                UPDDMPROCEDUR_QR(2) = Trim(SELDTEMPLETRS("UOM"))
                UPDDMPROCEDUR_QR(3) = Trim(MSFlexGrid1.TextMatrix(i, 3))
                UPDDMPROCEDUR_QR(4) = Trim(MSFlexGrid1.TextMatrix(i, 4))
                UPDDMPROCEDUR_QR(5) = Trim(MSFlexGrid1.TextMatrix(i, 5)) 'Shraddha 30/12/10
                UPDDMPROCEDUR_QR(6) = req_chglen(cmb_reqno.Text)
                UPDDMPROCEDUR_QR(7) = Trim(SELMA101_RS1("SER_CODE"))
                UPDDMPROCEDUR_QR.Execute
            
            
            
            End If
                
        Next i
                 updrrequisitn_qr(0) = cmb_enterby.Text
                 updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                 updrrequisitn_qr(2) = UCase(cmb_finalby.Text)
                 updrrequisitn_qr(3) = "T"
                 updrrequisitn_qr(4) = UCase(cmb_finalby.Text)
                 updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                 updrrequisitn_qr(6) = req_chglen(cmb_reqno.Text)
                 updrrequisitn_qr.Execute
                 
                 If Len(Trim(txt_comments.Text)) > 0 Then
                     UPDDMBCCOMMQR(0) = Null
                     UPDDMBCCOMMQR(0) = txt_comments.Text
                 Else
                        UPDDMBCCOMMQR(0) = Null
                 End If
                 UPDDMBCCOMMQR(1) = Trim(req_chglen(cmb_reqno.Text))
                 UPDDMBCCOMMQR.Execute
                 
                 
                 'RTIMELOG 18-01-05
                 UPDRTIMEM_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                 UPDRTIMEM_QR(1) = time1 'txt_time.Text
                 UPDRTIMEM_QR(2) = req_chglen(cmb_reqno.Text)
                 UPDRTIMEM_QR.Execute
                 
                 
End If



 If ADDENDUMFALG = True Then
      ADDENDUMFALG = False
       'Dim num As Integer

   
   If Trim(txt_addendum.Text) <> "" Then
        selectaddendum_qr(0) = req_chglen(cmb_reqno.Text)
        If Addendum_Flag = False Then
            Set selectaddendum_rs = selectaddendum_qr.OpenResultset(2, 3)
            Addendum_Flag = True
        Else
            selectaddendum_rs.MoreResults
            selectaddendum_rs.Requery
        End If
      
        If selectaddendum_rs.RowCount > 0 Then
            num = selectaddendum_rs("addend_no")
            num = num + 1
        Else
            num = 1
        End If
      
        insertaddendum_qr(0) = req_chglen(cmb_reqno.Text)
        insertaddendum_qr(1) = num
        insertaddendum_qr(2) = GET_REVDATE(Format(DATE11, "dd/mm/yyyy"))
        insertaddendum_qr(3) = time1 'txt_time.Text
        insertaddendum_qr(4) = Trim(cmb_addendumby.Text)
        insertaddendum_qr(5) = Trim(txt_addendum.Text)
        insertaddendum_qr.Execute
        
         'RTIMELOG 18-01-05
         UPDRTIMECADD_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
         UPDRTIMECADD_QR(1) = time1 'txt_time.Text
         UPDRTIMECADD_QR(2) = req_chglen(cmb_reqno.Text)
         UPDRTIMECADD_QR.Execute

   End If
  End If
  'added by shweta 23-04-2013 as emr evaluation list
cmd_commit_Click
'''till here
rdb.CommitTrans    'check this  Commit
MsgBox "Record is Saved Successfully."
Exit Sub
'MSG:
 'rdb.RollbackTrans
 'MsgBox Err.Description
End Sub


Public Sub OSMOAGE()

SS = Trim(txt_sex.Text)

If (txt_age.Text <= 60 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 60
    SS = "M"
End If

If (txt_age.Text <= 60 And txt_age.Text >= 1) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 60
    SS = "F"
End If

If (txt_age.Text <= 100 And txt_age.Text >= 60) And Trim(txt_sex.Text) = "M" Then
    RNAGE = 100
    SS = "M"
End If


If (txt_age.Text <= 100 And txt_age.Text >= 60) And Trim(txt_sex.Text) = "F" Then
    RNAGE = 100
    SS = "F"
End If


End Sub
