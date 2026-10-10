VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_ClinicalPharmaReport 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Clinical Pharmacology - Report"
   ClientHeight    =   11025
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15480
   LinkTopic       =   "Form2"
   ScaleHeight     =   11025
   ScaleWidth      =   15480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox TXT_CPCOMMENTS 
      Height          =   315
      ItemData        =   "frm_ClinicalPharmaReport.frx":0000
      Left            =   6480
      List            =   "frm_ClinicalPharmaReport.frx":0007
      TabIndex        =   51
      Top             =   7320
      Width           =   5535
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox cmb_checkedby 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   45
      Top             =   7800
      Width           =   3735
   End
   Begin VB.ComboBox cmb_finalby 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   42
      Top             =   8280
      Width           =   3735
   End
   Begin VB.ComboBox cmb_enterby 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   41
      Top             =   7320
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   12975
      Begin VB.ComboBox cmb_CA 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_ClinicalPharmaReport.frx":000E
         Left            =   9840
         List            =   "frm_ClinicalPharmaReport.frx":001B
         TabIndex        =   48
         Text            =   "cmb_CA"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txt_DMG 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9000
         TabIndex        =   24
         Top             =   240
         Width           =   3735
      End
      Begin VB.ComboBox cmb_reqno 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2520
         TabIndex        =   23
         Top             =   840
         Width           =   6375
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2520
         TabIndex        =   22
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3480
         TabIndex        =   21
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5040
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   0
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.TextBox txt_nmaterial 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7560
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_CA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CA Type"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9120
         TabIndex        =   49
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lbl_dmg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DMG"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   8520
         TabIndex        =   34
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1680
         TabIndex        =   33
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4920
         TabIndex        =   32
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1920
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1800
         TabIndex        =   30
         Top             =   1320
         Width           =   615
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
         Left            =   3120
         TabIndex        =   29
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4200
         TabIndex        =   28
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Specimen "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6720
         TabIndex        =   26
         Top             =   1320
         Width           =   840
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   8760
      Width           =   11775
      Begin VB.CommandButton cmd_checked 
         BackColor       =   &H0080C0FF&
         Caption         =   "Do Check"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   4800
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":002C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   300
         Width           =   1100
      End
      Begin VB.CommandButton CMD_NEW 
         BackColor       =   &H0080C0FF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   120
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":0336
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   300
         Width           =   900
      End
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H0080C0FF&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   1200
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":0640
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   300
         Width           =   900
      End
      Begin VB.CommandButton cmd_modify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   2280
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":094A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Width           =   1000
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save Entry"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   3480
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":0C54
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1100
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   8520
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":0F5E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   900
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H0080C0FF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   7440
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":1268
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   900
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H0080C0FF&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   9600
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":1572
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   900
      End
      Begin VB.CommandButton cmd_commit 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Do Commit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   6120
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":187C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   300
         Width           =   1100
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H0080C0FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   10680
         MouseIcon       =   "frm_ClinicalPharmaReport.frx":1B86
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.Frame REPORT 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   13695
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   36
         Top             =   240
         Width           =   270
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   35
         Top             =   240
         Width           =   270
      End
      Begin VB.TextBox txt_observed 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ForeColorFixed  =   192
         AllowUserResizing=   1
         FormatString    =   $"frm_ClinicalPharmaReport.frx":1E90
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6840
         TabIndex        =   40
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   38
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSWinsockLib.Winsock WIN_CLIENT 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblorg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   53
      Top             =   0
      Width           =   13515
   End
   Begin VB.Label lblcomment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5520
      TabIndex        =   52
      Top             =   7320
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checked By"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   360
      TabIndex        =   46
      Top             =   7800
      Width           =   930
   End
   Begin VB.Label lbl_final 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finalized By"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   360
      TabIndex        =   44
      Top             =   8280
      Width           =   945
   End
   Begin VB.Label lbl_enterby 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entered  By"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   360
      TabIndex        =   43
      Top             =   7320
      Width           =   945
   End
   Begin VB.Label lbl_biochemlab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT - CLINICAL PHARMACOLOGY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   5985
   End
End
Attribute VB_Name = "frm_ClinicalPharmaReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ICASENO, IREQNO

Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer

Dim R1, C1, srno

Dim PC As String
Dim str, range, CheckedByDesg, FinalByDesg As String
Dim TEST_DET, NU As String

Dim j As Integer
Dim i As Integer

Dim SEL_NEWRREQ_QR As rdoQuery
Dim SEL_NEWRREQ_RS As rdoResultset

Dim SEL_TA101_QR As rdoQuery
Dim SEL_TA101_RS As rdoResultset

Dim SEL_DREQUISITN_QR As rdoQuery
Dim SEL_DREQUISITN_RS As rdoResultset

Dim SEL_ELEMPINFO_QR As rdoQuery
Dim SEL_ELEMPINFO_RS As rdoResultset

Dim SEL_EMDESGMAS_QR As rdoQuery
Dim SEL_EMDESGMAS_RS As rdoResultset

Dim upd_drequisitn_qr As rdoQuery
Dim UPD_DREQUISITN_CH_QR As rdoQuery

Dim UPD_RTIMESAVE_QR As rdoQuery

Dim UPD_RTIMECOMMIT_QR As rdoQuery

Dim sel_rreq_qr As rdoQuery
Dim sel_rreq_rs As rdoResultset

Dim SEL_VIEWRREQ_QR As rdoQuery
Dim SEL_VIEWRREQ_RS As rdoResultset

Dim UPD_RREQSAVE_QR As rdoQuery
Dim UPD_RREQCOMMIT_QR As rdoQuery

Dim SELRREQUISITN_prevcp_QR As rdoQuery
Dim SELRREQUISITN_prevcp_RS As rdoResultset

Dim SEL_DREQSTAFFD_QR As rdoQuery
Dim SEL_DREQSTAFFD_RS As rdoResultset

Dim SEL_DMULPROCE_QR As rdoQuery
Dim SEL_DMULPROCE_RS As rdoResultset

Dim SEL_MA101_QR As rdoQuery
Dim SEL_MA101_RS As rdoResultset

Dim SEL_DTEMPLET_QR As rdoQuery
Dim SEL_DTEMPLET_RS As rdoResultset

Dim SEL_DTEMPLET607_QR As rdoQuery
Dim SEL_DTEMPLET607_RS As rdoResultset

Dim INS_DMPROCEDUR_QR As rdoQuery
Dim UPD_DMPROCEDUR_QR As rdoQuery

Dim SEL_DMPROCEDUR_QR As rdoQuery
Dim SEL_DMPROCEDUR_RS As rdoResultset

'********************** Added By Divyanshu for Comments add on 03/12/2025
Dim INS_DMBCCOMM_QR As rdoQuery
Dim UPD_DMBCCOMM_qr As rdoQuery

Dim SEL_DMBCCOMM_QR As rdoQuery
Dim SEL_DMBCCOMM_RS As rdoResultset

Dim CPCOMMENTLIST_qr As rdoQuery
Dim CPCOMMENTLIST_rs As rdoResultset
'**********************Till Here By Divyanshu for Comments add on 03/12/2025

Dim SEL_DISLNKUSRBYID_QR As rdoQuery
Dim SEL_DISLNKUSRBYID_RS As rdoResultset

Dim SEL_DISLNKUSRBYNAME_QR As rdoQuery
Dim SEL_DISLNKUSRBYNAME_RS As rdoResultset

Dim obserValNotEntered As Boolean

Dim USRTYPE As String

Dim LSPOS As Integer
Dim GRPOS As Integer



Private Sub cmb_CA_Click()
'If NEWFLAG = True Then
    ICASENO = MaskEdBox1.FormattedText
    IREQNO = Trim(cmb_reqno.Text)
    Dim TEST_DET1, RANGE1
    MSFlexGrid1.Rows = 1
    i = 1
        
    PC = "T607"
    
    SEL_MA101_QR(0) = PC
    Set SEL_MA101_RS = SEL_MA101_QR.OpenResultset(1, 3)
    TEST_DET = Trim(SEL_MA101_RS!SER_DET)
    
    lbl_CA.Visible = True
    cmb_CA.Visible = True
    
    If Trim(cmb_CA.Text) = "" Then
        MsgBox "Please select CA Type."
        cmd_save.Enabled = False
        Exit Sub
    Else
        cmd_save.Enabled = True
        SEL_DTEMPLET607_QR(0) = Trim(PC)
        SEL_DTEMPLET607_QR(1) = Left(Trim(cmb_CA.Text), 1)
        Set SEL_DTEMPLET607_RS = SEL_DTEMPLET607_QR.OpenResultset(1, 3)
        range = Trim(SEL_DTEMPLET607_RS!LOWVALUE) & "-" & Trim(SEL_DTEMPLET607_RS!HIGHVALUE)
    End If
    
    
    MSFlexGrid1.AddItem i & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & vbTab & range & vbTab & Trim(SEL_DTEMPLET607_RS!UOM) & vbTab & Trim(SEL_DTEMPLET607_RS!LOWPANIC) & vbTab & Trim(SEL_DTEMPLET607_RS!HIGHPANIC)
    
    
         
       
  
'End If
End Sub



Private Sub cmd_checked_Click()
On Error GoTo MSG:
Dim i As Integer

If Trim(cmb_checkedby.Text) = "" Then
    MsgBox "Please select Checked By", vbInformation
    Exit Sub
End If

'***
obserValNotEntered = False
If MSFlexGrid1.Rows = 1 Then
    obserValNotEntered = True
End If
For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
        obserValNotEntered = True
        Exit For
    End If
Next i
If obserValNotEntered = True Then
    MsgBox "No values"
    Exit Sub
End If
'***

i = MsgBox("Mark Checked?", vbYesNo)
If i = vbYes Then
 
''********SECURITY***************************
    SEL_DISLNKUSRBYNAME_QR(0) = Trim(cmb_checkedby.Text)
    Set SEL_DISLNKUSRBYNAME_RS = SEL_DISLNKUSRBYNAME_QR.OpenResultset(1, 3)
    
    If Not SEL_DISLNKUSRBYNAME_RS.EOF Then
    
        If (USRTYPE > 9) Then
          ''   no action as he can check
        ElseIf (USRTYPE < 8) Then
            MsgBox "Not authorized for checked by"
            Exit Sub
        ElseIf USRTYPE = Trim(SEL_DISLNKUSRBYNAME_RS("usrtype")) Then
            ''   no action as he can check
        End If
        If SEL_DISLNKUSRBYNAME_RS("COMMIT") <> "YES" Then
            MsgBox "Not authorized for checked by"
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security detials of this user is not available in DISLNKUSR"
        'cmd_commit.SetFocus
        Exit Sub
    End If
''********SECURITY***************************

'***** EMP
    CheckedByDesg = ""
    SEL_ELEMPINFO_QR(0) = Right(Trim(UID), 6)
    Set SEL_ELEMPINFO_RS = SEL_ELEMPINFO_QR.OpenResultset(1, 3)
        
    If Not SEL_ELEMPINFO_RS.EOF Then
        
        SEL_EMDESGMAS_QR(0) = Trim(SEL_ELEMPINFO_RS("DESGCODE"))
        Set SEL_EMDESGMAS_RS = SEL_EMDESGMAS_QR.OpenResultset(1, 3)
        
        If Not SEL_EMDESGMAS_RS.EOF Then
            CheckedByDesg = Trim(SEL_EMDESGMAS_RS("DESGSHDESC"))
        End If
    End If
'*****

 rdb.BeginTrans
 'update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=CURRENT DATE,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
    UPD_RREQSAVE_QR(0) = cmb_enterby.Text
    UPD_RREQSAVE_QR(1) = cmb_checkedby.Text
    UPD_RREQSAVE_QR(2) = "T"
    UPD_RREQSAVE_QR(3) = Trim(cmb_reqno.Text)
    UPD_RREQSAVE_QR.Execute
                    
    'UPDATE RISLIB.RTIMELOG SET SAVREPDATE=current date,SAVREPTIME=current time WHERE REQNNO=?"
    UPD_RTIMESAVE_QR(0) = Trim(cmb_reqno.Text)
    UPD_RTIMESAVE_QR.Execute
    
    
    'UPDATE DISLIB.DREQUISITN SET PREPATHNO=?,SPIHREQNO=? WHERE REQNO=?
    UPD_DREQUISITN_CH_QR(0) = Trim(CheckedByDesg)
    UPD_DREQUISITN_CH_QR(1) = Right(Trim(UID), 6)
    UPD_DREQUISITN_CH_QR(2) = Trim(cmb_reqno.Text)
    UPD_DREQUISITN_CH_QR.Execute
    
    '*****
  
    MsgBox "Record Marked as Checked Successfully."
End If
rdb.CommitTrans


BTNENABL1 ("CHECKED")
cmd_print.Enabled = True

Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub CMD_EMR_Click()
If Len(Trim(MaskEdBox1.Text)) > 0 Then
   Dim VAR1
   VAR1 = MaskEdBox1.Text
   Dim strLink As String
   strLink = "http://intranet.mpmmcc.tmc.gov.in/emrtest_rb/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
End Sub

Private Sub Form_Load()
'Main
''Commented By Monika
'lblorg.Caption = G_OrgFullName
' If G_UnitCode = "KM" Then
'    frm_ClinicalPharmaReport.BackColor = RGB(255, 255, 192)
' ElseIf G_UnitCode = "KR" Then
'    frm_ClinicalPharmaReport.BackColor = RGB(255, 216, 177)
' End If
'''''''''''''''''''''''''''
flagsave = 0
NEWFLAG = False
viewflag = False

cmb_CA.Visible = False
lbl_CA.Visible = False

cmb_CA.Text = cmb_CA.List(0)

BTNENABL1 ("INI")
FLDENABL1 ("INI")
DATETIME

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set sel_rreq_qr = rdb.CreateQuery("", str)

str = "select REQNNO from Rislib.RREQUISITN where CASENO = ? AND REQNSTAT = 'E' AND MODALITYCD = 'CP' ORDER BY REQNNO DESC "
Set SEL_NEWRREQ_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
Set SEL_VIEWRREQ_QR = rdb.CreateQuery("", str)

str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=CURRENT DATE,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
Set UPD_RREQSAVE_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=CURRENT DATE, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set UPD_RREQCOMMIT_QR = rdb.CreateQuery("", str)

'*********************
str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=current date,SAVREPTIME=current time WHERE REQNNO=?"
Set UPD_RTIMESAVE_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=current date,COMREPTIME=current time WHERE REQNNO=?"
Set UPD_RTIMECOMMIT_QR = rdb.CreateQuery("", str)
'*********************

str = "SELECT NMETERIALS FROM DISLIB.DREQUISITN WHERE REQNO=? "
Set SEL_DREQUISITN_QR = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.DREQUISITN SET OPPROCEDUR=?,DRECEIVED=? WHERE REQNO=? "
Set upd_drequisitn_qr = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.DREQUISITN SET PREPATHNO=?,SPIHREQNO=? WHERE REQNO=? "
Set UPD_DREQUISITN_CH_QR = rdb.CreateQuery("", str)

'*********************
str = "SELECT * FROM EMPLIB.ELEMPINFO WHERE EMPCODE=?"
Set SEL_ELEMPINFO_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM EMPLIB.EMDESGMAS WHERE DESGCODE=?"
Set SEL_EMDESGMAS_QR = rdb.CreateQuery("", str)
'*********************

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set SEL_MA101_QR = rdb.CreateQuery("", str)
'*********************

str = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE=? "
Set SEL_DTEMPLET_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE=? AND SEX=?"
Set SEL_DTEMPLET607_QR = rdb.CreateQuery("", str)
'*********************

str = "select * from PATOOLSLIB.ta101 where case_no =?"
Set SEL_TA101_QR = rdb.CreateQuery("", str)
'*********************

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
Set SEL_DREQSTAFFD_QR = rdb.CreateQuery("", str)

'-------------------------Added By divyanshu on 05/06/2021------------
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND REQNSTAT=? ORDER BY REQNDT DESC "
Set SELRREQUISITN_prevcp_QR = rdb.CreateQuery("", str)
'------------------------------Till here-----------------------------
'*********************

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SEL_DMULPROCE_QR = rdb.CreateQuery("", str)

'*********************

'str = "INSERT INTO DISLIB.DMPROCEDUR(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,ENTRYBY,PVALUE) VALUES (?,?,?,?,?,?,?,?)"
str = "INSERT INTO DISLIB.DMPROCEDUR(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,ENTRYBY,PVALUE,REMARKS) VALUES (?,?,?,?,?,?,?,?,?)"
Set INS_DMPROCEDUR_QR = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,REMARKS=? WHERE REQNO=? AND PRO_CODE=? "
Set UPD_DMPROCEDUR_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set SEL_DMPROCEDUR_QR = rdb.CreateQuery("", str)
'*********************
'********************** Added By Divyanshu for Comments add on 03/12/2025
str = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES (?,?)"
Set INS_DMBCCOMM_QR = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=? "
Set UPD_DMBCCOMM_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
Set SEL_DMBCCOMM_QR = rdb.CreateQuery("", str)

str = "select TESTNAME from dislib.FECCOMMENT WHERE MODALITY='CP'"
Set CPCOMMENTLIST_qr = rdb.CreateQuery("CPCOMLST", str)
Set CPCOMMENTLIST_rs = CPCOMMENTLIST_qr.OpenResultset(1, 3)
If CPCOMMENTLIST_rs.EOF = False Then
CPCOMMENTLIST_rs.MoveFirst
While Not CPCOMMENTLIST_rs.EOF
TXT_CPCOMMENTS.AddItem (Trim(CPCOMMENTLIST_rs(0)))
CPCOMMENTLIST_rs.MoveNext
Wend

End If
'**********************Till Here By Divyanshu for Comments add on 03/12/2025
'**********************

str = "select * from DISLIB.Dislnkusr where username = ?"
Set SEL_DISLNKUSRBYNAME_QR = rdb.CreateQuery("", str)

'*************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)

If Not USRNAME_RS.EOF Then
    USRTYPE = Trim(USRNAME_RS("USRTYPE"))
    If (Trim(USRTYPE) = "4" Or Trim(USRTYPE) = "8" Or Trim(USRTYPE) = "6" Or Trim(USRTYPE) = "7") Then
        cmb_enterby.AddItem Trim(USRNAME_RS("USERNAME"))
        cmb_enterby.Text = Trim(USRNAME_RS("USERNAME"))
    End If
    
    If (Trim(USRTYPE) = "9" Or Trim(USRTYPE) = "8") And (Trim(USRNAME_RS("DEPT_CODE")) = "CPH" Or Trim(USRNAME_RS("DEPT_CODE")) = "FEC") Then
        cmb_checkedby.AddItem Trim(USRNAME_RS("USERNAME"))
        cmb_checkedby.Text = Trim(USRNAME_RS("USERNAME"))
    End If
    
    'If Trim(DESGCODE) = "1689" Or Trim(MODMR) = "C" Then
    'If (Trim(USRTYPE) = "9" Or Trim(USRTYPE) = "8") And Trim(USRNAME_RS("MODMR")) = "C" Then ''COMMENTED BY MONIKA
    If (Trim(USRTYPE) = "9" Or Trim(USRTYPE) = "8") And Trim(USRNAME_RS("MODMR")) = "7" Then  'ADDED BY MONIKA
        cmb_finalby.AddItem Trim(USRNAME_RS("USERNAME"))
        cmb_finalby.Text = Trim(USRNAME_RS("USERNAME"))
    End If
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_checkedby.Text = cmb_checkedby.List(0)
    cmb_finalby.Text = cmb_finalby.List(0)
Else
    MsgBox "The Details of this User are not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If


MSFlexGrid1.col = 2
txt_observed.Width = MSFlexGrid1.CellWidth
txt_observed.Height = MSFlexGrid1.CellHeight
txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

cmd_cancel_Click
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If Trim(MaskEdBox1.Text) = "" Or MaskEdBox1.Text = "__/_____" Or MaskEdBox1.Text = "_______/_______" Then
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        'MaskEdBox1.Mask = ">##?####/######"
    Else
        'MaskEdBox1.Mask = ">??/#####"
    End If
End If
KeyAscii = KeyAscii

End Sub
Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String

If Len(Trim(MaskEdBox1.Text)) = "" Then
    MsgBox "Please enter CaseNo/CCNo"
    Exit Sub
End If

CASENO = MaskEdBox1.Text
If Opt_patient.Value = True Then
    'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
    MaskEdBox1.Text = NEWNOpad(CASENO)
End If
If NEWFLAG = True Then
    'select REQNNO from Rislib.RREQUISITN where CASENO = ? AND REQNSTAT = 'E' AND MODALITYCD = 'CP' ORDER BY REQNNO DESC "
    SEL_NEWRREQ_QR(0) = MaskEdBox1.FormattedText
    Set SEL_NEWRREQ_RS = SEL_NEWRREQ_QR.OpenResultset(1, 3)
    If Not SEL_NEWRREQ_RS.EOF Then
        cmb_reqno.clear
        While Not SEL_NEWRREQ_RS.EOF
            cmb_reqno.AddItem Trim(SEL_NEWRREQ_RS("reqNno"))
            SEL_NEWRREQ_RS.MoveNext
        Wend
    Else
        cmb_reqno.clear
    End If
End If

If viewflag = True Then
'SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC
    SEL_VIEWRREQ_QR(0) = MaskEdBox1.FormattedText
    SEL_VIEWRREQ_QR(1) = "CP"
    SEL_VIEWRREQ_QR(2) = "CPH"
    SEL_VIEWRREQ_QR(3) = "T"
    SEL_VIEWRREQ_QR(4) = "F"
    Set SEL_VIEWRREQ_RS = SEL_VIEWRREQ_QR.OpenResultset(1, 3)
    
    If Not SEL_VIEWRREQ_RS.EOF Then
        cmb_reqno.clear
        While Not SEL_VIEWRREQ_RS.EOF
            cmb_reqno.AddItem Trim(SEL_VIEWRREQ_RS("REQNNO"))
            SEL_VIEWRREQ_RS.MoveNext
        Wend
    Else
        MsgBox "No Requisitions"
        BTNENABL1 ("INI")
        cmd_cancel_Click
        Exit Sub
    End If
End If

End Sub
Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
    'MaskEdBox1.Mask = "??>/#####"
    MaskEdBox1.Enabled = True
    lbl_dmg.Visible = True
    txt_DMG.Visible = True
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    MaskEdBox1.Mask = "######"
    MaskEdBox1.Enabled = True
    lbl_dmg.Visible = False
    txt_DMG.Visible = False
End If
End Sub

Private Sub cmb_reqno_Click()
'cmb_reqno.Enabled = False
Dim entr As String
Dim str, str1 As String

SEL_DREQUISITN_QR(0) = Trim(cmb_reqno.Text)
Set SEL_DREQUISITN_RS = SEL_DREQUISITN_QR.OpenResultset(1, 1)
txt_nmaterial.Text = SEL_DREQUISITN_RS("NMETERIALS")

'***** DMG

sel_rreq_qr(0) = Trim(cmb_reqno.Text)
Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 1)
txt_DMG.Text = sel_rreq_rs("REFBYDOC")
txt_category.Text = sel_rreq_rs("PAT_CATG")
'*****

If Opt_patient.Value = True Then
    SEL_TA101_QR(0) = MaskEdBox1.FormattedText
    Set SEL_TA101_RS = SEL_TA101_QR.OpenResultset(1, 3)

    If SEL_TA101_RS.EOF Then
        MsgBox "Registration details Not found in HBCH Database"
        Exit Sub
    End If
    
    txt_sex.Text = SEL_TA101_RS("sex")
    txt_name.Text = Trim(SEL_TA101_RS("first_name")) & " " & Trim(SEL_TA101_RS("mid_name")) & " " & Trim(SEL_TA101_RS("last_name"))
    txt_age.Text = DateDiff("yyyy", SEL_TA101_RS("BIRTH_DATE"), CStr(DATE11))
        
ElseIf opt_staff.Value = True Then
    
    SEL_DREQSTAFFD_QR(0) = Trim(cmb_reqno.Text)
    SEL_DREQSTAFFD_QR(1) = Trim(MaskEdBox1.FormattedText)
    Set SEL_DREQSTAFFD_RS = SEL_DREQSTAFFD_QR.OpenResultset(1, 3)
    
    If Not SEL_DREQSTAFFD_RS.EOF Then
        txt_sex.Text = Trim(SEL_DREQSTAFFD_RS("sex"))
        txt_name.Text = Trim(SEL_DREQSTAFFD_RS("PNAME"))
        txt_age.Text = Trim(SEL_DREQSTAFFD_RS("AGE"))
        txt_category.Text = Trim(SEL_DREQSTAFFD_RS("RELATION"))
    End If
End If

MSFlexGrid1.Visible = True
MSFlexGrid1.clear
MSFlexGrid1.FormatString = "Sr.No.|^Test Name                                       |^Observed Value|^Reference Range |^ Unit       |^ LLQ |^ULQ|^PrevVal|^PreValDt  "

If NEWFLAG = True Then
    ICASENO = MaskEdBox1.FormattedText
    IREQNO = Trim(cmb_reqno.Text)
    Dim TEST_DET1, RANGE1
    MSFlexGrid1.Rows = 1
    i = 1
    
    SEL_DMULPROCE_QR(0) = Trim(cmb_reqno.Text)
    Set SEL_DMULPROCE_RS = SEL_DMULPROCE_QR.OpenResultset(1, 3)

    If Not SEL_DMULPROCE_RS.EOF Then
        Do Until SEL_DMULPROCE_RS.EOF
            PC = SEL_DMULPROCE_RS!PRO_CODE
            
            SEL_MA101_QR(0) = PC
            Set SEL_MA101_RS = SEL_MA101_QR.OpenResultset(1, 3)
            TEST_DET = Trim(SEL_MA101_RS!SER_DET)
            
            If Trim(PC) = "T607" Then
                lbl_CA.Visible = True
                cmb_CA.Visible = True
                If Trim(cmb_CA.Text) = "" Then
                    MsgBox "Please select CA Type."
                    cmd_save.Enabled = False
                    Exit Sub
                Else
                    cmd_save.Enabled = True
                    SEL_DTEMPLET607_QR(0) = Trim(PC)
                    SEL_DTEMPLET607_QR(1) = Left(Trim(cmb_CA.Text), 1)
                    Set SEL_DTEMPLET607_RS = SEL_DTEMPLET607_QR.OpenResultset(1, 3)
                    range = Trim(SEL_DTEMPLET607_RS!LOWVALUE) & "-" & Trim(SEL_DTEMPLET607_RS!HIGHVALUE)
                End If
            Else
                lbl_CA.Visible = False
                cmb_CA.Visible = False
                SEL_DTEMPLET_QR(0) = PC
                Set SEL_DTEMPLET_RS = SEL_DTEMPLET_QR.OpenResultset(1, 3)
                range = Trim(SEL_DTEMPLET_RS!LOWVALUE) & "-" & Trim(SEL_DTEMPLET_RS!HIGHVALUE)
            
            End If
            MSFlexGrid1.AddItem i & vbTab & TEST_DET & "$" & Trim(PC) & vbTab & vbTab & range & vbTab & Trim(SEL_DTEMPLET_RS!UOM) & vbTab & Trim(SEL_DTEMPLET_RS!LOWPANIC) & vbTab & Trim(SEL_DTEMPLET_RS!HIGHPANIC)
            i = i + 1
    
            SEL_DMULPROCE_RS.MoveNext
        Loop
    End If
End If

If viewflag = True Then

    sel_rreq_qr(0) = Trim(cmb_reqno.Text)
    Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)

    If Not sel_rreq_rs.EOF Then
        cmb_enterby.Text = Trim(sel_rreq_rs("ENTERBY"))
        If Trim(sel_rreq_rs("PREPORTBY")) <> "" Then
            cmb_checkedby.Text = Trim(sel_rreq_rs("PREPORTBY"))
            MsgBox "Report is already marked as checked."
            cmd_checked.Enabled = False
        Else
            cmb_checkedby.Text = cmb_checkedby.List(0)
            cmd_checked.Enabled = True
        End If
        
        If sel_rreq_rs("COMMITFLAG") = "Y" Then
            MsgBox "Report is commited, it can not be modified"
            cmd_commit.Enabled = False
            CMD_MODIFY.Enabled = False
            If IsNull(sel_rreq_rs("FREPORTBY")) = False Then
                cmb_finalby.Text = Trim(sel_rreq_rs("FREPORTBY"))
            End If
        Else
            If (Trim(USRTYPE) = 4 Or Trim(USRTYPE) = 8) Then
                CMD_MODIFY.Enabled = True
            Else
                CMD_MODIFY.Enabled = False
            End If
        End If
        
    End If
    
    SEL_DMPROCEDUR_QR(0) = Trim(cmb_reqno.Text)
    Set SEL_DMPROCEDUR_RS = SEL_DMPROCEDUR_QR.OpenResultset(1, 3)

    If Not SEL_DMPROCEDUR_RS.EOF Then
        i = 1
        srno = 1
        While Not SEL_DMPROCEDUR_RS.EOF
            SEL_MA101_QR(0) = SEL_DMPROCEDUR_RS("PRO_CODE")
            Set SEL_MA101_RS = SEL_MA101_QR.OpenResultset(1, 3)
            'SEL_DTEMPLET_QR(0) = SEL_DMPROCEDUR_RS("PRO_CODE")
            'Set SEL_DTEMPLET_RS = SEL_DTEMPLET_QR.OpenResultset(1, 3)
            
          If SEL_DMPROCEDUR_RS("PRO_CODE") = "T607" Then
            lbl_CA.Visible = True
            cmb_CA.Visible = True
            
            cmb_CA.Text = Trim(SEL_DMPROCEDUR_RS("REMARKS"))
        Else
        
                lbl_CA.Visible = False
            cmb_CA.Visible = False
            
            cmb_CA.Text = ""
          End If
            
'            '-----
'            If Trim(SEL_DMPROCEDUR_RS("PRO_CODE")) = "T607" Then
'                SEL_DTEMPLET607_QR(0) = Trim(PC)
'                SEL_DTEMPLET607_QR(1) = Left(Trim(cmb_CA.Text), 1)
'                Set SEL_DTEMPLET607_RS = SEL_DTEMPLET607_QR.OpenResultset(1, 3)
'
'                If Not SEL_DTEMPLET607_RS.EOF Then
'                    INS_DMPROCEDUR_QR(4) = Trim(SEL_DTEMPLET607_RS("UOM")) 'UOM
'                End If
'            Else
'                SEL_DTEMPLET_QR(0) = sc
'                Set SEL_DTEMPLET_RS = SEL_DTEMPLET_QR.OpenResultset(1, 3)
'
'                If Not SEL_DTEMPLET_RS.EOF Then
'                    INS_DMPROCEDUR_QR(4) = Trim(SEL_DTEMPLET_RS("UOM")) 'UOM
'                End If
'            End If
'            '-----
            

            'If Not SEL_DTEMPLET_RS.EOF Then
                If Not SEL_MA101_RS.EOF Then
                    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                    MSFlexGrid1.row = srno
                    MSFlexGrid1.col = 0
                                        
                    MSFlexGrid1.TextMatrix(i, 0) = i
                    MSFlexGrid1.TextMatrix(i, 1) = Trim(SEL_MA101_RS("SER_DET")) & "$" & Trim(SEL_DMPROCEDUR_RS("PRO_CODE"))
                    If Trim(SEL_DMPROCEDUR_RS("obser_type")) = "H" Then
                        MSFlexGrid1.row = srno
                        MSFlexGrid1.col = 2
                        MSFlexGrid1.CellBackColor = vbRed
                    Else
                        If Trim(SEL_DMPROCEDUR_RS("obser_type")) = "L" Then
                            MSFlexGrid1.col = 2
                            MSFlexGrid1.CellBackColor = vbGreen
                        End If
                        If Trim(SEL_DMPROCEDUR_RS("obser_type")) = "N" Then
                            MSFlexGrid1.col = 2
                            MSFlexGrid1.CellBackColor = vbWhite
                        End If
                    End If
                
                   
                    MSFlexGrid1.TextMatrix(i, 2) = Trim(SEL_DMPROCEDUR_RS("OBSERVALUE"))
                    MSFlexGrid1.TextMatrix(i, 3) = Trim(SEL_DMPROCEDUR_RS("nrange")) 'Trim(SEL_DTEMPLET_RS("LOWVALUE")) & "-" & Trim(SEL_DTEMPLET_RS("HIGHVALUE"))
                    MSFlexGrid1.TextMatrix(i, 4) = Trim(SEL_DMPROCEDUR_RS("UOM"))
                    
                End If
                If IsNull(SEL_DMPROCEDUR_RS("ENTRYBY")) = False Then
                    MSFlexGrid1.TextMatrix(i, 5) = Trim(SEL_DMPROCEDUR_RS("ENTRYBY"))
                End If
                If IsNull(SEL_DMPROCEDUR_RS("PVALUE")) = False Then
                    MSFlexGrid1.TextMatrix(i, 6) = Trim(SEL_DMPROCEDUR_RS("PVALUE"))
                End If
                srno = srno + 1
                i = i + 1
            'End If
            SEL_DMPROCEDUR_RS.MoveNext
        Wend
        txt_observed.Visible = False
    End If
    
    '********************** Added By Divyanshu for Comments add on 03/12/2025
    
    SEL_DMBCCOMM_QR(0) = Trim(cmb_reqno.Text)
    Set SEL_DMBCCOMM_RS = SEL_DMBCCOMM_QR.OpenResultset(1, 3)

    If Not SEL_DMBCCOMM_RS.EOF Then
        TXT_CPCOMMENTS.Text = SEL_DMBCCOMM_RS("COMMENTS")
    End If
    
    '********************** Till Here By Divyanshu for Comments add on 03/12/2025

End If
    Call prev_val
End Sub


Private Sub cmd_save_Click()

If Trim(cmb_enterby.Text) = "" Then
    MsgBox "Please select Entered By", vbInformation
    Exit Sub
End If

obserValNotEntered = False
If MSFlexGrid1.Rows = 1 Then
    obserValNotEntered = True
End If
For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
        obserValNotEntered = True
        Exit For
    End If
Next i
If obserValNotEntered = True Then
    MsgBox "Please enter observed value"
    Exit Sub
End If

assign

BTNENABL1 ("SAVE")

End Sub

Public Sub assign()
On Error GoTo MSG:
Dim NU As Integer
Dim num As Integer
Dim sc As String
Dim count, low, high, W1, obsValue

rdb.BeginTrans
If flagsave = 1 Then

    For i = 1 To MSFlexGrid1.Rows - 1
        num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
        sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
        
        'DMPROCEDUR(REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,ENTRYBY,PVALUE"
        INS_DMPROCEDUR_QR(0) = Trim(cmb_reqno.Text) 'REQNO
        INS_DMPROCEDUR_QR(1) = Trim(sc) 'PRO_CODE
        INS_DMPROCEDUR_QR(2) = Trim(MSFlexGrid1.TextMatrix(i, 2)) 'OBSERVALUE
        
        '***
        
        
        
        
    
        NU = InStr(Trim(MSFlexGrid1.TextMatrix(i, 3)), "-")
        low = Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), 1, NU - 1)
        high = Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), NU + 1)
        
        obsValue = Trim(MSFlexGrid1.TextMatrix(i, 2))
        
        '***
        If Trim(high) = "NA" Then
            If InStr(Trim(obsValue), "<") = 0 And InStr(Trim(obsValue), ">") = 0 Then
                If Val(obsValue) >= Val(low) Then
                    INS_DMPROCEDUR_QR(3) = "N"
                End If
                If Val(obsValue) < Val(low) Then
                    INS_DMPROCEDUR_QR(3) = "L"
                End If
            ElseIf InStr(Trim(obsValue), "<") > 0 Then
                If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) >= Val(low) Then
                    INS_DMPROCEDUR_QR(3) = "N"
                End If
                If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) < Val(low) Then
                    INS_DMPROCEDUR_QR(3) = "L"
                End If
                
            ElseIf InStr(Trim(obsValue), ">") > 0 Then
                'MsgBox "Greater than sign (>) is not applicable"
                'Exit Sub
            End If
        Else
            '***
            If InStr(Trim(obsValue), "<") = 0 And InStr(Trim(obsValue), ">") = 0 Then
                '***
                If Val(obsValue) > Val(high) Then
                    INS_DMPROCEDUR_QR(3) = "H"
                Else
                    If Val(obsValue) >= Val(low) Then
                        INS_DMPROCEDUR_QR(3) = "N"
                    End If
                    If Val(obsValue) < Val(low) Then
                        INS_DMPROCEDUR_QR(3) = "L"
                    End If
                End If
                '***
            ElseIf InStr(Trim(obsValue), "<") > 0 Then
                '***
                If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) >= Val(low) Then
                    INS_DMPROCEDUR_QR(3) = "N"
                End If
                If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) < Val(low) Then
                    INS_DMPROCEDUR_QR(3) = "L"
                End If
                '***
                
            ElseIf InStr(Trim(txt_observed.Text), ">") > 0 Then
                '###
                If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) > Val(high) Then
                    INS_DMPROCEDUR_QR(3) = "H"
                Else
                    If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) >= Val(low) Then
                        INS_DMPROCEDUR_QR(3) = "N"
                    End If
                End If
                '###
            End If
            
            '***
        End If
        
        
        'INS_DMPROCEDUR_QR(4) = Trim(SEL_DTEMPLET_RS("UOM")) 'UOM
        
        If Trim(sc) = "T607" Then
            SEL_DTEMPLET607_QR(0) = Trim(PC)
            SEL_DTEMPLET607_QR(1) = Left(Trim(cmb_CA.Text), 1)
            Set SEL_DTEMPLET607_RS = SEL_DTEMPLET607_QR.OpenResultset(1, 3)
            
            If Not SEL_DTEMPLET607_RS.EOF Then
                INS_DMPROCEDUR_QR(4) = Trim(SEL_DTEMPLET607_RS("UOM")) 'UOM
            End If
        Else
            SEL_DTEMPLET_QR(0) = sc
            Set SEL_DTEMPLET_RS = SEL_DTEMPLET_QR.OpenResultset(1, 3)
            
            If Not SEL_DTEMPLET_RS.EOF Then
                INS_DMPROCEDUR_QR(4) = Trim(SEL_DTEMPLET_RS("UOM")) 'UOM
            End If
        End If
        
        INS_DMPROCEDUR_QR(5) = Trim(MSFlexGrid1.TextMatrix(i, 3)) 'NRANGE
        INS_DMPROCEDUR_QR(6) = Trim(MSFlexGrid1.TextMatrix(i, 5)) 'ENTERBY
        INS_DMPROCEDUR_QR(7) = Trim(MSFlexGrid1.TextMatrix(i, 6)) 'PVALUE
        
        If Trim(sc) = "T607" Then
            INS_DMPROCEDUR_QR(8) = Trim(cmb_CA.Text)
        Else
            INS_DMPROCEDUR_QR(8) = ""
        End If
        INS_DMPROCEDUR_QR.Execute
        
        
    Next i
    
    '********************** Added By Divyanshu for Comments add on 03/12/2025
    
    INS_DMBCCOMM_QR(0) = Trim(cmb_reqno.Text)  'REQNO
    INS_DMBCCOMM_QR(1) = Trim(TXT_CPCOMMENTS.Text)  'Comments
    INS_DMBCCOMM_QR.Execute
    
    '********************** Till Here By Divyanshu for Comments add on 03/12/2025
    
    'update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=CURRENT DATE,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
    UPD_RREQSAVE_QR(0) = cmb_enterby.Text
    UPD_RREQSAVE_QR(1) = cmb_checkedby.Text
    UPD_RREQSAVE_QR(2) = "T"
    UPD_RREQSAVE_QR(3) = Trim(cmb_reqno.Text)
    UPD_RREQSAVE_QR.Execute
    
    'UPDATE RISLIB.RTIMELOG SET SAVREPDATE=current date,SAVREPTIME=current time WHERE REQNNO=?"
    UPD_RTIMESAVE_QR(0) = Trim(cmb_reqno.Text)
    UPD_RTIMESAVE_QR.Execute
    
End If

If flagsave = 2 Then
'UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=? WHERE REQNO=? AND PRO_CODE=? "
    
    For i = 1 To MSFlexGrid1.Rows - 1
        obsValue = Trim(MSFlexGrid1.TextMatrix(i, 2))
        num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
        sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
        
        SEL_DTEMPLET_QR(0) = Trim(sc)
        Set SEL_DTEMPLET_RS = SEL_DTEMPLET_QR.OpenResultset(1, 3)

        If Not SEL_DTEMPLET_RS.EOF Then
            NU = InStr(Trim(MSFlexGrid1.TextMatrix(i, 3)), "-")
            low = Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), 1, NU - 1)
            high = Mid(Trim(MSFlexGrid1.TextMatrix(i, 3)), NU + 1)
           
            UPD_DMPROCEDUR_QR(0) = Trim(obsValue)
            
            '***
            If Trim(high) = "NA" Then
                If InStr(Trim(obsValue), "<") = 0 And InStr(Trim(obsValue), ">") = 0 Then
                    If Val(obsValue) >= Val(low) Then
                        UPD_DMPROCEDUR_QR(1) = "N"
                    End If
                    If Val(obsValue) < Val(low) Then
                        UPD_DMPROCEDUR_QR(1) = "L"
                    End If
                ElseIf InStr(Trim(obsValue), "<") > 0 Then
                    '^^^
                    If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) >= Val(low) Then
                        UPD_DMPROCEDUR_QR(1) = "N"
                    End If
                    If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) < Val(low) Then
                        UPD_DMPROCEDUR_QR(1) = "L"
                    End If
                    '^^^
                End If
                
            Else
            '***
                If InStr(Trim(obsValue), "<") = 0 And InStr(Trim(obsValue), ">") = 0 Then
                    '---
                    If Val(obsValue) > Val(high) Then
                        UPD_DMPROCEDUR_QR(1) = "H"
                    Else
                        If Val(obsValue) >= Val(low) Then
                            UPD_DMPROCEDUR_QR(1) = "N"
                        End If
                        If Val(obsValue) < Val(low) Then
                            UPD_DMPROCEDUR_QR(1) = "L"
                        End If
                    End If
                    '---
                '$$$
                ElseIf InStr(Trim(txt_observed.Text), "<") > 0 Then
                    '---
                    If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) >= Val(low) Then
                        UPD_DMPROCEDUR_QR(1) = "N"
                    End If
                    If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) < Val(low) Then
                        UPD_DMPROCEDUR_QR(1) = "L"
                    End If
                    '---
                    
                ElseIf InStr(Trim(txt_observed.Text), ">") > 0 Then
                    '---
                    If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) > Val(high) Then
                        UPD_DMPROCEDUR_QR(1) = "H"
                    Else
                        If Val(Right(Trim(obsValue), Len(Trim(obsValue)) - 1)) >= Val(low) Then
                            UPD_DMPROCEDUR_QR(1) = "N"
                        End If
                    End If
                    '---
                End If
                '$$$
                
            End If
        End If
            
        If Trim(sc) = "T607" Then
            UPD_DMPROCEDUR_QR(2) = Trim(cmb_CA.Text)
        Else
            UPD_DMPROCEDUR_QR(2) = ""
        End If
            
        UPD_DMPROCEDUR_QR(3) = Trim(cmb_reqno.Text)
        UPD_DMPROCEDUR_QR(4) = Trim(sc)
        UPD_DMPROCEDUR_QR.Execute
        
    Next i
    
    '********************** Added By Divyanshu for Comments add on 03/12/2025
    
    UPD_DMBCCOMM_qr(0) = Trim(TXT_CPCOMMENTS.Text)
    UPD_DMBCCOMM_qr(1) = Trim(cmb_reqno.Text)
    UPD_DMBCCOMM_qr.Execute
    
    '********************** Till Here By Divyanshu for Comments add on 03/12/2025
    
        
    'update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=CURRENT DATE,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
    UPD_RREQSAVE_QR(0) = cmb_enterby.Text
    UPD_RREQSAVE_QR(1) = cmb_checkedby.Text
    UPD_RREQSAVE_QR(2) = "T"
    UPD_RREQSAVE_QR(3) = Trim(cmb_reqno.Text)
    UPD_RREQSAVE_QR.Execute
    
    'UPDATE RISLIB.RTIMELOG SET SAVREPDATE=current date,SAVREPTIME=current time WHERE REQNNO=?"
    UPD_RTIMESAVE_QR(0) = Trim(cmb_reqno.Text)
    UPD_RTIMESAVE_QR.Execute

End If
rdb.CommitTrans
MsgBox "Record Saved Successfully"
BTNENABL1 ("INI")
txt_observed.Visible = False
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description

End Sub

Private Sub cmd_view_Click()

BTNENABL1 ("VIEW")
FLDENABL1 ("VIEW")

NEWFLAG = False
viewflag = True

'*****
MSFlexGrid1.FormatString = "Sr.No.|^Test Name                                       |^Observed Value|^Reference Range |^ Unit       |^ LLQ |^ULQ  "


End Sub
Public Sub prev_val()

SELRREQUISITN_prevcp_QR(0) = MaskEdBox1.FormattedText
SELRREQUISITN_prevcp_QR(1) = "CP"
SELRREQUISITN_prevcp_QR(2) = "CPH"
SELRREQUISITN_prevcp_QR(3) = "F"
Set SELRREQUISITN_prevcp_RS = SELRREQUISITN_prevcp_QR.OpenResultset(1, 3)


If Not SELRREQUISITN_prevcp_RS.EOF Then


'*********************************************

Dim DMPROCEDUR_QR As rdoQuery
Dim DMPROCEDUR_RS As rdoResultset
str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO = ? "
Set DMPROCEDUR_QR = rdb.CreateQuery("", str)
DMPROCEDUR_QR(0) = Trim(SELRREQUISITN_prevcp_RS("reqnno")) 'req_chglen(cmb_reqno.List(1))
Set DMPROCEDUR_RS = DMPROCEDUR_QR.OpenResultset(1, 3)
While Not DMPROCEDUR_RS.EOF
    i = 1
    While i < MSFlexGrid1.Rows
    
        If Trim(DMPROCEDUR_RS("PRO_CODE")) = Right(Trim(MSFlexGrid1.TextMatrix(i, 1)), 4) Then
            'MsgBox "match"
            MSFlexGrid1.TextMatrix(i, 7) = Trim(DMPROCEDUR_RS("OBSERVALUE"))
            
            '***********Added by divyanshu for showing the previous value date in the grid on 13/07/2020********
            
            Dim RREQN_QR As rdoQuery
            Dim RREQN_RS As rdoResultset
            str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO = ? "
            Set RREQN_QR = rdb.CreateQuery("", str)
            RREQN_QR(0) = Trim(SELRREQUISITN_prevcp_RS("reqnno"))
            Set RREQN_RS = RREQN_QR.OpenResultset(1, 3)
            If Not RREQN_RS.EOF Then
               MSFlexGrid1.TextMatrix(i, 8) = Trim(RREQN_RS("REQNDT"))
            End If
            
            '*********************************Till here by divyanshu on 13/07/2020*******************
        End If
        i = i + 1
    Wend
    DMPROCEDUR_RS.MoveNext
Wend
End If
End Sub
Private Sub cmd_back_Click()
BTNENABL1 ("INI")
FLDENABL1 ("INI")
field_clear
viewflag = False
NEWFLAG = False
BCCLEAR
Unload Me
'Me.Hide
DIS_MAIN.show
End Sub

Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    CMD_MODIFY.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_MODIFY.Enabled = True
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_MODIFY.Enabled = False
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = True
    cmd_view.Enabled = True
    
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    CMD_MODIFY.Enabled = False
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
    CMD_MODIFY.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
    
ElseIf cap = "print" Then
     cmd_save.Enabled = False
     cmd_back.Enabled = True
     CMD_MODIFY.Enabled = False
     cmd_print.Enabled = False
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False
     
End If
End Sub
Private Sub cmd_commit_Click()
On Error GoTo MSG:
Dim i As Integer

If Trim(cmb_finalby.Text) = "" Then
    MsgBox "Please select Finalized By", vbInformation
    Exit Sub
End If

'***
obserValNotEntered = False
If MSFlexGrid1.Rows = 1 Then
    obserValNotEntered = True
End If
For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(i, 2)) = "" Then
        obserValNotEntered = True
        Exit For
    End If
Next i
If obserValNotEntered = True Then
    MsgBox "No values"
    Exit Sub
End If
'***

i = MsgBox(" Do you want to Commit the Record?", vbYesNo)
If i = vbYes Then
 
''********SECURITY***************************
    SEL_DISLNKUSRBYNAME_QR(0) = Trim(cmb_finalby.Text)
    Set SEL_DISLNKUSRBYNAME_RS = SEL_DISLNKUSRBYNAME_QR.OpenResultset(1, 3)
    
    If Not SEL_DISLNKUSRBYNAME_RS.EOF Then
    
        If (USRTYPE > SEL_DISLNKUSRBYNAME_RS("usrtype") Or USRTYPE >= 6) Then
          ''   no action as he can commit
        ElseIf USRTYPE < SEL_DISLNKUSRBYNAME_RS("usrtype") Then
            MsgBox "Sorry! This Report needs  is to Committed by a Higher Level than  that of Yours"
            'cmd_commit.SetFocus
            Exit Sub
        ElseIf USRTYPE = SEL_DISLNKUSRBYNAME_RS("usrtype") Then
            
        End If
        If SEL_DISLNKUSRBYNAME_RS("COMMIT") <> "YES" Then
            MsgBox "Sorry! You are not authorized to Commit the Report"
            'cmd_commit.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security detials of this user is not available in DISLNKUSR"
        'cmd_commit.SetFocus
        Exit Sub
    End If
''********SECURITY***************************

'***** EMP
FinalByDesg = ""
SEL_ELEMPINFO_QR(0) = Right(Trim(UID), 6)
Set SEL_ELEMPINFO_RS = SEL_ELEMPINFO_QR.OpenResultset(1, 3)
    
If Not SEL_ELEMPINFO_RS.EOF Then
    
    SEL_EMDESGMAS_QR(0) = Trim(SEL_ELEMPINFO_RS("DESGCODE"))
    Set SEL_EMDESGMAS_RS = SEL_EMDESGMAS_QR.OpenResultset(1, 3)
    
    If Not SEL_EMDESGMAS_RS.EOF Then
        FinalByDesg = Trim(SEL_EMDESGMAS_RS("DESGSHDESC"))
    End If
End If
'*****

 rdb.BeginTrans
 'UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=CURRENT DATE, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    UPD_RREQCOMMIT_QR(0) = UCase(cmb_finalby.Text)
    UPD_RREQCOMMIT_QR(1) = "F"
    UPD_RREQCOMMIT_QR(2) = "Y"
    UPD_RREQCOMMIT_QR(3) = Trim(cmb_reqno.Text)
    UPD_RREQCOMMIT_QR.Execute
                    
    'UPDATE RISLIB.RTIMELOG SET COMREPDATE=current date,COMREPTIME=current time WHERE REQNNO=?"
    UPD_RTIMECOMMIT_QR(0) = Trim(cmb_reqno.Text)
    UPD_RTIMECOMMIT_QR.Execute
    
    'UPDATE DISLIB.DREQUISITN SET OPPROCEDUR=? WHERE REQNO=?
    'UPDATE DISLIB.DREQUISITN SET OPPROCEDUR=?,DRECEIVED=? WHERE REQNO=? "
    upd_drequisitn_qr(0) = Trim(FinalByDesg)
    upd_drequisitn_qr(1) = Right(Trim(UID), 6)
    upd_drequisitn_qr(2) = Trim(cmb_reqno.Text)
    upd_drequisitn_qr.Execute
  
    'Send SMS to Patient
'    If Opt_patient.Value = True Then
'        sendRMSG Trim(cmb_reqno.Text), MaskEdBox1.FormattedText, "F", WIN_CLIENT.LocalIP
'    End If
    '*****
  
    MsgBox "Record Committed Successfully."
End If
rdb.CommitTrans
'----------------------Added By Divyanshu for the message to patient on 20/11/2025 ----------------
Dim ipadres As String
Dim msgrep As String
Dim crtmsgrep As String
ipadres = WIN_CLIENT.LocalIP
'msgrep = SendFinalRepMsg(Replace(Trim(cmb_reqno.Text), "/", ""), Trim(MaskEdBox1.FormattedText), "Clinical Pharmocology", ipadres, UID) '''COMMENTED BY MONIKA 21/07/2026
'----------------------------Till Here By Divyanshu for message on 20/11/2025 ---------------------

BTNENABL1 ("COMMIT")
cmd_print.Enabled = True

Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_exit_Click()
End
End Sub

Private Sub cmd_modify_Click()
BTNENABL1 ("NEW")
FLDENABL1 ("MODIFY")

flagsave = 2
End Sub

Private Sub cmd_new_Click()

BTNENABL1 ("NEW")
FLDENABL1 ("NEW")
'**********

TXT_CLEAR
BCCLEAR
MSFlexGrid1.clear

DATETIME
txt_observed.Visible = False
flagsave = 1
viewflag = False
NEWFLAG = True

End Sub

Private Sub cmd_cancel_Click()
BTNENABL1 ("INI")
FLDENABL1 ("INI")

field_clear
viewflag = False
NEWFLAG = False
flagsave = 0
BCCLEAR

End Sub

Public Sub field_clear()
Dim obj As Object

cmb_reqno.clear

If Opt_patient.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = ">??/#####"
End If

If opt_staff.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
MaskEdBox1.Mask = "#######"
End If
TXT_CPCOMMENTS.Text = ""
txt_name.Text = " "
txt_sex.Text = " "
txt_age.Text = " "
txt_category.Text = " "
lbl_CA.Visible = False
cmb_CA.Visible = False
cmb_CA.Text = ""
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In biochemlab_report.Controls
If TypeOf obj Is TextBox Then
       obj = ""
End If
Next

End Sub

Public Sub BCCLEAR()
TXT_CLEAR
txt_observed.Visible = False
MSFlexGrid1.Visible = True
MSFlexGrid1.Rows = 1
End Sub
Public Sub FLDENABL1(str1 As String)
Dim obj As Object
If str1 = "INI" Then

    MaskEdBox1.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    cmb_reqno.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
    Opt_patient.Value = False
    opt_staff.Value = False
    
    MaskEdBox1.Text = ""
    txt_name.Text = ""
    txt_age.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_nmaterial.Text = ""
    txt_DMG.Text = ""

    cmb_reqno.clear
    
End If

If str1 = "NEW" Then

    MaskEdBox1.Enabled = False
    Opt_patient.Enabled = True
    opt_staff.Enabled = True
    cmb_reqno.Enabled = True
    cmb_enterby.Enabled = False
    'cmb_finalby.Enabled = True
    cmb_finalby.Enabled = False
    Opt_patient.Value = False
    opt_staff.Value = False
    
    MaskEdBox1.Text = ""
End If

If str1 = "VIEW" Then

    MaskEdBox1.Enabled = False
    Opt_patient.Enabled = True
    opt_staff.Enabled = True
    cmb_reqno.Enabled = True
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
    Opt_patient.Value = False
    opt_staff.Value = False
        
    MaskEdBox1.Text = ""
End If
If str1 = "MODIFY" Then

    MaskEdBox1.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    cmb_reqno.Enabled = False
    cmb_enterby.Enabled = False
    
    cmb_finalby.Enabled = False
End If

End Sub


Private Sub cmd_print_Click()

DoEvents

Dim strLink As String
strLink = "https://intranet.hbchrcm.tmc.gov.in/emrreports/allreports.aspx?reqno=" & Trim(cmb_reqno.Text) & "~guest"
Call Shell("C:\Program Files\Internet Explorer\iexplore " & strLink, vbMaximizedFocus)

End Sub


Private Sub MSFlexGrid1_Click()
txt_observed.Visible = False
If flagsave = 1 Then
    txt_observed.Visible = True
ElseIf flagsave = 2 Then
    txt_observed.Visible = True
End If

If txt_observed.Visible = False Then
    Exit Sub
End If

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
End If
End Sub

Private Sub txt_observed_KeyPress(KeyAscii As Integer)

Dim NU As Integer
Dim num As Integer
Dim sc As String
Dim low, high As Variant

num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)

If KeyAscii <> 13 And KeyAscii <> 46 And KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 60 And KeyAscii <> 62 Then
    MsgBox "Please enter valid value"
    txt_observed.Text = ""
    Exit Sub
End If

If KeyAscii = 13 Then
    If Trim(txt_observed.Text) = "" Then
        MsgBox "Please enter observed value"
        Exit Sub
    End If
    SEL_DTEMPLET_QR(0) = sc
    Set SEL_DTEMPLET_RS = SEL_DTEMPLET_QR.OpenResultset(1, 3)

    NU = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), "-")
    low = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), 1, NU - 1)
    high = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)), NU + 1)
    
    If Trim(high) = "NA" Then
        If InStr(Trim(txt_observed.Text), "<") = 0 And InStr(Trim(txt_observed.Text), ">") = 0 Then
            If Val(Trim(txt_observed.Text)) >= Val(low) Then
                MSFlexGrid1.CellBackColor = vbWhite
            End If
            If Val(Trim(txt_observed.Text)) < Val(low) Then
                MSFlexGrid1.CellBackColor = vbGreen
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
        ElseIf InStr(Trim(txt_observed.Text), "<") > 0 Then
            LSPOS = InStr(Trim(txt_observed.Text), "<")
            If LSPOS <> 1 Then
                MsgBox "Less than sign (<) should be at beginning"
                Exit Sub
            Else
                If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) >= Val(low) Then
                    MSFlexGrid1.CellBackColor = vbWhite
                End If
                If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) < Val(low) Then
                    MSFlexGrid1.CellBackColor = vbGreen
                End If
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
        ElseIf InStr(Trim(txt_observed.Text), ">") > 0 Then
            MsgBox "Greater than sign (>) is not applicable"
            Exit Sub
        End If
        
    Else
        If InStr(Trim(txt_observed.Text), "<") = 0 And InStr(Trim(txt_observed.Text), ">") = 0 Then
            If Val(Trim(txt_observed.Text)) > Val(high) Then
                MSFlexGrid1.CellBackColor = vbRed
                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
            Else
                If Val(Trim(txt_observed.Text)) >= Val(low) Then
                    MSFlexGrid1.CellBackColor = vbWhite
                End If
                If Val(Trim(txt_observed.Text)) < Val(low) Then
                    MSFlexGrid1.CellBackColor = vbGreen
                End If
                MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
            End If
        ElseIf InStr(Trim(txt_observed.Text), "<") > 0 Then
            LSPOS = InStr(Trim(txt_observed.Text), "<")
            If LSPOS <> 1 Then
                MsgBox "Less than sign (<) should be at beginning"
                Exit Sub
            Else
                '***
                If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) > Val(high) Then
                    MsgBox "Less than sign (<) is not applicable"
                    Exit Sub
                    'MSFlexGrid1.CellBackColor = vbRed
                    'MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
                Else
                    If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) >= Val(low) Then
                        MSFlexGrid1.CellBackColor = vbWhite
                    End If
                    If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) < Val(low) Then
                        MSFlexGrid1.CellBackColor = vbGreen
                    End If
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
                End If
                '***
            End If
        ElseIf InStr(Trim(txt_observed.Text), ">") > 0 Then
            GRPOS = InStr(Trim(txt_observed.Text), ">")
            If GRPOS <> 1 Then
                MsgBox "Greater than sign (>) should be at beginning"
                Exit Sub
            Else
                '***
                If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) > Val(high) Then
                    MSFlexGrid1.CellBackColor = vbRed
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
                Else
                    If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) >= Val(low) Then
                        MSFlexGrid1.CellBackColor = vbWhite
                    End If
                    If Val(Right(Trim(txt_observed.Text), Len(Trim(txt_observed.Text)) - 1)) < Val(low) Then
                        MsgBox "Greater than sign (>) is not applicable"
                        Exit Sub
                        'MSFlexGrid1.CellBackColor = vbGreen
                    End If
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = Trim(txt_observed.Text)
                End If
                '***
            End If
        End If
    End If
    
    
    
    
    
    i = MSFlexGrid1.Rows - 1
    If i = MSFlexGrid1.row Then
        txt_observed.Visible = False
        Exit Sub
    End If
    If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1 'srno
    End If
    If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
        txt_observed.Visible = False
    End If
    If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        txt_observed.Text = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
        txt_observed.Visible = True
        txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    End If
End If
End Sub

Private Sub BTNENABL1(cap As String)
If cap = "NEW" Then
    CMD_NEW.Enabled = False
    cmd_view.Enabled = False
    CMD_MODIFY.Enabled = False
    cmd_save.Enabled = True
    cmd_checked.Enabled = False
    cmd_commit.Enabled = False
    cmd_print.Enabled = False
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_exit.Enabled = True
    
ElseIf cap = "VIEW" Then
    CMD_NEW.Enabled = False
    cmd_view.Enabled = False
    If USRTYPE = 4 Then
       CMD_MODIFY.Enabled = True
    Else
        CMD_MODIFY.Enabled = True
    End If
    cmd_save.Enabled = False
    If Trim(USRNAME_RS("COMMIT")) = "YES" And (Trim(USRNAME_RS("DEPT_CODE")) = "CPH" Or Trim(USRNAME_RS("DEPT_CODE")) = "FEC") Then
        cmd_checked.Enabled = True
    Else
        cmd_checked.Enabled = False
    End If
    'If Trim(DESGCODE) = "1689" Or Trim(MODMR) = "C" Then
    ''COMMENTED BY MONIKA TO CHANGE MODMR CONDITION'''
'    If (Trim(USRTYPE) = "9" Or Trim(USRTYPE) = "8") And Trim(USRNAME_RS("MODMR")) = "C" Then
If (Trim(USRTYPE) = "9" Or Trim(USRTYPE) = "8") And Trim(USRNAME_RS("MODMR")) = "7" Then
        cmd_commit.Enabled = True
    Else
        cmd_commit.Enabled = False
    End If
    cmd_print.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_exit.Enabled = True
    
ElseIf cap = "SAVE" Then
    CMD_NEW.Enabled = False
    cmd_view.Enabled = False
    CMD_MODIFY.Enabled = False
    cmd_save.Enabled = False
    
    cmd_checked.Enabled = False

    cmd_commit.Enabled = False
    
    cmd_print.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_exit.Enabled = True
    
ElseIf cap = "INI" Then
    If (Trim(USRTYPE) = 4 Or Trim(USRTYPE) = 8 Or Trim(USRTYPE) = 7 Or Trim(USRTYPE) = 6) Then
        CMD_NEW.Enabled = True
    ElseIf Trim(USRTYPE) = 9 Then
        CMD_NEW.Enabled = False
    Else
        CMD_NEW.Enabled = False
    End If
    cmd_view.Enabled = True
    CMD_MODIFY.Enabled = False
    cmd_save.Enabled = False
    cmd_checked.Enabled = False
    cmd_commit.Enabled = False
    cmd_print.Enabled = False
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_exit.Enabled = True
    
ElseIf cap = "COMMIT" Then
    CMD_NEW.Enabled = False
    cmd_view.Enabled = False
    CMD_MODIFY.Enabled = False
    cmd_save.Enabled = False
    cmd_checked.Enabled = False
    cmd_commit.Enabled = False
    cmd_print.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_exit.Enabled = True
    
ElseIf cap = "CHECKED" Then
    CMD_NEW.Enabled = False
    cmd_view.Enabled = False
    CMD_MODIFY.Enabled = False
    cmd_save.Enabled = False
    cmd_checked.Enabled = False
    cmd_commit.Enabled = False
    cmd_print.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_exit.Enabled = True
    
ElseIf cap = "print" Then
    CMD_NEW.Enabled = False
    cmd_view.Enabled = False
    CMD_MODIFY.Enabled = False
    cmd_save.Enabled = False
    cmd_checked.Enabled = False
    cmd_commit.Enabled = False
    cmd_print.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_exit.Enabled = True
End If
End Sub


