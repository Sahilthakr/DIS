VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form imhreport 
   Caption         =   "Report - ImmunohistoChemistry"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CRYSTALREPORT1 
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame REPORT 
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   0
      TabIndex        =   55
      Top             =   3840
      Width           =   11775
      Begin VB.ComboBox cmb_grade 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "IMHREPORT.frx":0000
         Left            =   2400
         List            =   "IMHREPORT.frx":0002
         TabIndex        =   79
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txt_impfindiag 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2430
         TabIndex        =   72
         Top             =   1200
         Width           =   3855
      End
      Begin VB.ComboBox CMB_N 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "IMHREPORT.frx":0004
         Left            =   8790
         List            =   "IMHREPORT.frx":0006
         TabIndex        =   71
         Top             =   1560
         Width           =   585
      End
      Begin VB.ComboBox cmb_m 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "IMHREPORT.frx":0008
         Left            =   9750
         List            =   "IMHREPORT.frx":000A
         TabIndex        =   70
         Text            =   "cmb_m"
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox cmb_t 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "IMHREPORT.frx":000C
         Left            =   6720
         List            =   "IMHREPORT.frx":000E
         TabIndex        =   69
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txt_size 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7710
         TabIndex        =   68
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txt_ihmeterial 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   66
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txt_report 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   720
         Width           =   6975
      End
      Begin VB.ComboBox cmb_finalized 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7200
         TabIndex        =   8
         Top             =   3360
         Width           =   2175
      End
      Begin VB.ComboBox cmb_provisionalby 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox cmb_enterby 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "IMHREPORT.frx":0010
         Left            =   2520
         List            =   "IMHREPORT.frx":0012
         TabIndex        =   6
         Top             =   3000
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid Msftest 
         Height          =   495
         Left            =   840
         TabIndex        =   56
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   10485760
         ForeColorFixed  =   255
         AllowUserResizing=   3
         FormatString    =   "S.No  |^          Test Order          |^             Test Result               |^               Remarks         "
      End
      Begin VB.TextBox txt_positive 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   840
         TabIndex        =   57
         Top             =   2520
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txt_remarks 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   480
         TabIndex        =   58
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label LBL_VPP 
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   10200
         TabIndex        =   81
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_grade 
         AutoSize        =   -1  'True
         Caption         =   "Grade"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   80
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label lbl_impression 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   78
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6480
         TabIndex        =   77
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8670
         TabIndex        =   76
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "M"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9510
         TabIndex        =   75
         Top             =   1680
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7350
         TabIndex        =   74
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label LBL_CM 
         Caption         =   "cm"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8310
         TabIndex        =   73
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label LBL_IHMETERIAL 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Meterial"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   67
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label LBL_IHREPORT 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1680
         TabIndex        =   65
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lbl_finalizedby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalised By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5520
         TabIndex        =   61
         Top             =   3480
         Width           =   1590
      End
      Begin VB.Label lbl_provisionalby 
         AutoSize        =   -1  'True
         Caption         =   "Provisional By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5880
         TabIndex        =   60
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label LBL_ENTEREDBY 
         AutoSize        =   -1  'True
         Caption         =   "Entered By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1320
         TabIndex        =   59
         Top             =   3120
         Width           =   945
      End
   End
   Begin VB.Frame common 
      Height          =   3015
      Left            =   0
      TabIndex        =   24
      Top             =   840
      Width           =   11775
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   0
         TabIndex        =   62
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   0
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   0
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   22
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
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
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   53
         Top             =   2520
         Width           =   1830
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   42
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   41
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   40
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   39
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   37
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   36
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   35
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   34
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Path No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   32
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   31
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   30
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   1830
      End
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame COMMAND 
      Height          =   855
      Left            =   0
      TabIndex        =   25
      Top             =   7680
      Width           =   11775
      Begin VB.CommandButton cmd_fullreport 
         Caption         =   "Full&Report"
         Height          =   375
         Left            =   10080
         MouseIcon       =   "IMHREPORT.frx":0014
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   6360
         MouseIcon       =   "IMHREPORT.frx":031E
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   8880
         MouseIcon       =   "IMHREPORT.frx":0628
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   7680
         MouseIcon       =   "IMHREPORT.frx":0932
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5160
         MouseIcon       =   "IMHREPORT.frx":0C3C
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3960
         MouseIcon       =   "IMHREPORT.frx":0F46
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "IMHREPORT.frx":1250
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1560
         MouseIcon       =   "IMHREPORT.frx":155A
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         MouseIcon       =   "IMHREPORT.frx":1864
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label lbl_cap 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - IMUNOHISTOCHEMISTRY"
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
      Left            =   3480
      TabIndex        =   44
      Top             =   480
      Width           =   5115
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9840
      TabIndex        =   27
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9840
      TabIndex        =   26
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "imhreport"
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
Dim seldfsreport_qr As rdoQuery
Dim seldfsreport_rs As rdoResultset
Dim seldihreport_qr As rdoQuery
Dim seldihreport_rs As rdoResultset
Dim DIHREPORTFLAG1 As Boolean
Dim seldimpresion_qr As rdoQuery
Dim seldimpresion_rs As rdoResultset
Dim SELDIMPRESIONFLAG1 As Boolean
'Dim seldgrossdesc_qr As rdoQuery
'Dim seldgrossdesc_rs As rdoResultset
'Dim seldspreport_qr As rdoQuery
'Dim seldspreport_rs As rdoResultset
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELTA1012FLAG1 As Boolean
Dim w_immun As Application
'Dim d_immun As Document
Dim d_immun As Object
Dim WIHMET As Application
'Dim DIHMET As Document
Dim DIHMET As Object
Dim WIMP As Application
'Dim DIMP As Document
Dim DIMP As Object
'Dim SELEMPLIB_QR As rdoQuery
'Dim SELEMPLIB_RS As rdoResultset
Dim I As Integer
'Dim insdfsrep_qr As rdoQuery
'Dim insdgrossdesc_qr As rdoQuery
'Dim insdfs_qr As rdoQuery
Dim insdimpresion_qr As rdoQuery
'Dim insdspreport_qr As rdoQuery
Dim INSDIHREPORT_QR As rdoQuery
'Dim upddfsreport_qr As rdoQuery
Dim upddihreport_qr As rdoQuery
Dim upddimpresion_qr As rdoQuery
'Dim upddspreport_qr As rdoQuery
'Dim upddimpresion_qr As rdoQuery
'Dim updgrossdesc_qr As rdoQuery
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean
'Dim selempbasinfo1_qr As rdoQuery
'Dim selempbasinfo1_rs As rdoResultset

Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim INSDADDENDUM_QR As rdoQuery
'Dim SELDADDENDUM_QR As rdoQuery
'Dim SELDADDENDUM_RS As rdoResultset
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG1 As Boolean

Dim insdfreport_qr As rdoQuery
Dim upddfreport_qr As rdoQuery
Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset
Dim seldfreportflag As Boolean
Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer
Dim ADDENDUMFALG As Boolean

Dim seldissecentr1_qr As rdoQuery
Dim seldissecentr1_rs As rdoResultset
Dim seldissecentrflag1 As Boolean

Dim seldissecprov1_qr As rdoQuery
Dim seldissecprov1_rs As rdoResultset
Dim seldissecprovflag1 As Boolean

Dim seldissecfin1_qr As rdoQuery
Dim seldissecfin1_rs As rdoResultset
Dim seldissecfinflag1 As Boolean

Dim SELDREQSTAFFDQR As rdoQuery
Dim SELDREQSTAFFDRS As rdoResultset
Dim SELDREQSTAFFD1FLAG As Boolean
Dim upddreqnoqr As rdoQuery

Dim SELDGRAGEQR As rdoQuery
Dim seldgraders As rdoResultset
Dim SELDGRADEFLAG As Boolean

Dim INSDISIMP_QR As rdoQuery
Dim UPDDISIMP_QR As rdoQuery


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
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub



Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "IMMUNO HISTOCHEMISTRY$IH" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_back_Click()
TXT_CLEAR
btnenabl ("INI")
FLDENABL ("DIS")
DIS_MAIN.show
imhreport.Hide

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
    cmd_fullreport.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    cmd_fullreport.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = True
    cmd_view.Enabled = True
    cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
    cmd_fullreport.Enabled = False
ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    CMD_NEW.Enabled = True
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = True
    cmd_fullreport.Enabled = False
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
On Error GoTo MSG:
btnenabl ("INI")
field_clear
TXT_CLEAR
opt_patient.Value = False
opt_staff.Value = False
'txt_positive.Visible = False
'txt_remarks.Visible = False
testorder = ""
FLDENABL ("DIS")
'Msftest.Rows = 1
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_commit_Click()
On Error GoTo MSG:

 rdb.BeginTrans
    updrrequisitn1_qr(0) = Left(cmb_finalized.Text, 30)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
rdb.CommitTrans
btnenabl ("INI")
MsgBox "Record is Committed Successfully."
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_fullreport_Click()
'txt_fullreport_GotFocus
'Set d1 = w.ActiveDocument
'If Len(txt_frosecrep) > 0 Then
'd1.Content = d1.Content & "FROZEN:" & Chr(13) & Trim(txt_frosecrep)
'End If
'If Len(txt_grossdesc.Text) > 0 Then
'd1.Content = d1.Content & "GROSS DESC: " & Chr(13) & Trim(txt_grossdesc)
'End If
'If Len(txt_immhischerep) > 0 Then
'd1.Content = d1.Content & "IMMUNOHISTO:" & Chr(13) & Trim(txt_immhischerep)
'End If
'If Len(txt_microdesc) > 0 Then
'd1.Content = d1.Content & "MICROSCOPICDESC:" & Chr(13) & Trim(txt_microdesc)
'End If
'If Len(txt_impfindiag) > 0 Then
'd1.Content = d1.Content & "IMPRESSION:" & Chr(13) & Trim(txt_impfindiag)
'End If
'If Len(txt_addendum) > 0 Then
'd1.Content = d1.Content & "ADDENDUM:" & Chr(13) & Trim(txt_addendum)
'End If
End Sub

Private Sub cmd_fullreport_LostFocus()
'On Error GoTo msg:
'
'd1.Content = ""
'd1.SaveAs "C:\REPORT.DOC"
'd1.Close
'w.Quit
''Msk_caseno.Enabled = False
''cmd_cancel.SetFocus
'MSG:
End Sub

Private Sub cmd_modify_Click()
Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim STR4 As String
btnenabl ("NEW")
FLDENABL ("MODIFY")
flagsave = 2
        If SELDGRADEFLAG = False Then
                Set seldgraders = SELDGRAGEQR.OpenResultset(1, 3)
                SELDGRADEFLAG = True
            Else
                seldgraders.MoreResults
                seldgraders.Requery
            End If
            str1 = cmb_grade.Text
            str2 = cmb_t.Text
            STR4 = cmb_m.Text
            STR3 = CMB_N.Text
            If seldgraders.RowCount > 0 Then
                cmb_grade.clear
                While Not seldgraders.EOF
                    cmb_grade.AddItem seldgraders("gradedesc") & "$" & seldgraders("gradecode")
                    seldgraders.MoveNext
                Wend
            End If
            cmb_grade.Text = str1
            cmb_t.clear
            cmb_m.clear
            CMB_N.clear
        For I = 0 To 5
            cmb_t.AddItem I
            cmb_m.AddItem I
            CMB_N.AddItem I
        Next I
            cmb_m.AddItem "X"
            CMB_N.AddItem "X"
            cmb_t = str2
            cmb_m = STR4
            CMB_N = STR3
            
'txt_positive.Visible = True
'txt_remarks.Visible = True


End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG:
field_clear
btnenabl ("NEW")
If SELDEPTRSFLAG1 = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SELDEPTRSFLAG1 = True
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
opt_patient.Value = True
testorder = ""
txt_positive.Enabled = True
txt_remarks.Enabled = True
flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()


On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
 On Error GoTo prtmsg:
'txt_msg.Text = "PRINTING..."
DoEvents
    str = ""
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & TIME11
Dim VAR_REQNNO As String


VAR_REQNNO = req_chglen(cmb_reqno.Text)

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
            prepath = ""
           
            While Not seldrequisitn_rs.EOF
                If Trim(seldrequisitn_rs("pathno")) <> Trim(TXT_PATH.Text) Then
                    prepath = prepath & seldrequisitn_rs("pathno")
                End If
                seldrequisitn_rs.MoveNext
            Wend
            
        End If


CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"   'AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISIH.RPT"
CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & txt_prepathreqno & "'"
            If REFDRFLAG1 = True Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = MaskEdBox1.FormattedText
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
If REFDR_RS.RowCount > 0 Then

    If Trim(REFDR_RS("REQNSTAT")) = "T" Then
        CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "PROVISIONAL HISTOPATH & IHC REPORT" & "'"
        CRYSTALREPORT1.Formulas(4) = "ENT = '" & "Enter by " & cmb_enterby.Text & "'"
        'CRYSTALREPORT1.Formulas(3) = "ENT = '" & "ENTERED BY " & cmb_enterby & "'"
    Else
        CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "FINAL HISTOPATH & IHC REPORT" & "'"
        CRYSTALREPORT1.Formulas(4) = "ENT = '" & "Enter by " & cmb_enterby.Text & "'"
    End If
    
End If
        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
        I = InStr(cmb_grade.Text, "$")
        If I = 0 Then
            If Len(Trim(cmb_grade.Text)) <> 0 Then
                CRYSTALREPORT1.Formulas(5) = "GRADE = '" & "GRADE " & cmb_grade.Text & " '"
            Else
                CRYSTALREPORT1.Formulas(5) = "GRADE = '" & "" & " '"
            End If
            
        Else
            CRYSTALREPORT1.Formulas(5) = "GRADE = '" & "GRADE " & Mid(cmb_grade.Text, 1, I - 1) & " '"
        End If
        If Trim(txt_size.Text) = "" Then
            CRYSTALREPORT1.Formulas(6) = "TSIZE = '" & "" & " '"
        Else
            CRYSTALREPORT1.Formulas(6) = "TSIZE = '" & "TSIZE  " & txt_size & " '"
        End If
        If Trim(cmb_t.Text) = "" Then
            CRYSTALREPORT1.Formulas(7) = "T = '" & "" & " '"
        Else
            CRYSTALREPORT1.Formulas(7) = "T = '" & "T " & cmb_t.Text & " '"
        End If
        If Trim(CMB_N.Text) = "" Then
            CRYSTALREPORT1.Formulas(8) = "N = '" & "" & " '"
        Else
            CRYSTALREPORT1.Formulas(8) = "N = '" & "N " & CMB_N & " '"
        End If
        If Trim(cmb_m.Text) = "" Then
            CRYSTALREPORT1.Formulas(9) = "M = '" & "" & " '"
        Else
            CRYSTALREPORT1.Formulas(9) = "M = '" & "M " & cmb_m & " '"
            
        End If
        If Trim(prepath) = "" Then
            CRYSTALREPORT1.Formulas(10) = "PREPATH = '" & "" & " '"
        Else
            CRYSTALREPORT1.Formulas(10) = "PREPATH = '" & "Previous Path No. " & Left(prepath, 100) & " '"
            
        End If
        If Trim(txt_impfindiag) = "" Then
            CRYSTALREPORT1.Formulas(11) = "IMP = '" & "" & " '"
        Else
            CRYSTALREPORT1.Formulas(11) = "IMP = '" & "Impression" & " '"
            
        End If
        CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.Action = 1


'    STR1 = ip
'    STR2 = date11
'    STR3 = time1
'    STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
'    STR1 = USR_NAME & Space(3) & STR1
'    crp_rep.Formulas(3) = "us_id = '" & STR1 & "'"
''    crp_rep.Formulas(2) = "INOUT = '" & txt_status & "'"
''    crp_rep.Formulas(3) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
''
''    crp_rep.Formulas(4) = "ward_bed = '" & txt_ward.Text & "'"
''    crp_rep.Formulas(5) = "req_placed = '" & cmb_dept & "'"
''    crp_rep.Formulas(6) = "name='" & txt_name.Text & "'"
''    crp_rep.Formulas(7) = "sex='" & txt_sex.Text & "'"
''    crp_rep.Formulas(8) = "age='" & txt_age.Text & "'"
''    crp_rep.Formulas(9) = "primarydrser='" & txt_pridoc & " '"
''    crp_rep.Formulas(10) = "refdr='" & txt_refdoc & " ' "
''    crp_rep.Formulas(11) = "provdiagby='" & txt_procedure & " ' "
    
    
    

'If addendum_flag = True Then
'cr_rep.ReportFileName = "C:\RISREPORTS\crpt_rep2.rpt"
'Else
'cr_rep.ReportFileName = "C:\RISREPORTS\crpt_rep1.rpt"
'End If

''crp_rep.ReportFileName = "C:\SPREPORT\REQ_REP.RPT"
''crp_rep.WindowState = crptMaximized
''crp_rep.Action = 1

Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
btnenabl ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)

End Sub

Private Sub cmd_save_Click()

On Error GoTo MSG:


Dim m As Integer
If Len(Trim(cmb_reqno)) = 0 Then
    m = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
    cmb_reqno.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_enterby)) = 0 Then
    m = MsgBox("The Enter by field is EMPTY.", vbCritical)
    cmb_enterby.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_provisionalby)) = 0 Then
    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
    cmb_provisionalby.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_finalized)) = 0 Then
    m = MsgBox("The Provisional By field is EMPTY.", vbCritical)
    cmb_finalized.SetFocus
    Exit Sub
End If
If Len(Trim(txt_report.Text)) > 7000 Then
    MsgBox " Immunohistochemistry Max. of 7000 Characters Only.", vbInformation
    txt_report.SetFocus
    Exit Sub
End If
If Len(Trim(txt_ihmeterial.Text)) > 1000 Then
    MsgBox " Nature of Material  Max. of 1000 Characters Only.", vbInformation
    txt_ihmeterial.SetFocus
    Exit Sub
End If
If Len(Trim(txt_impfindiag.Text)) > 2048 Then
    MsgBox " Impression  Max. of 2048 Characters Only.", vbInformation
    txt_impfindiag.SetFocus
    Exit Sub
End If
If Len(Trim(txt_impfindiag.Text)) = 0 Then
    MsgBox " Please Enter Impression. ", vbInformation
    txt_impfindiag.SetFocus
    Exit Sub
End If
   
    While InStr(txt_report, Chr(9)) > 0
            Dim beforst As String, afterst As String
            Dim K As Long
            K = InStr(txt_report, Chr(9))
            beforst = Mid(txt_report, 1, K - 1) & " "
            afterst = Mid(txt_freport, K + 1)
            txt_report = beforst & afterst
       Wend
        While InStr(txt_report, Chr(146)) > 0

            K = InStr(txt_report, Chr(146))
            beforst = Mid(txt_report, 1, K - 1) & "`"
            afterst = Mid(txt_report, K + 1)
            txt_report = beforst & afterst
       Wend
       While InStr(txt_ihmeterial, Chr(9)) > 0
            
            K = InStr(txt_ihmeterial, Chr(9))
            beforst = Mid(txt_ihmeterial, 1, K - 1) & " "
            afterst = Mid(txt_ihmeterial, K + 1)
            txt_ihmeterial = beforst & afterst
       Wend
        While InStr(txt_ihmeterial, Chr(146)) > 0

            K = InStr(txt_ihmeterial, Chr(146))
            beforst = Mid(txt_ihmeterial, 1, K - 1) & "`"
            afterst = Mid(txt_ihmeterial, K + 1)
            txt_ihmeterial = beforst & afterst
       Wend
       While InStr(txt_impfindiag, Chr(9)) > 0
            
            K = InStr(txt_impfindiag, Chr(9))
            beforst = Mid(txt_impfindiag, 1, K - 1) & " "
            afterst = Mid(txt_impfindiag, K + 1)
            txt_impfindiag = beforst & afterst
       Wend
        While InStr(txt_impfindiag, Chr(146)) > 0

            K = InStr(txt_impfindiag, Chr(146))
            beforst = Mid(txt_impfindiag, 1, K - 1) & "`"
            afterst = Mid(txt_impfindiag, K + 1)
            txt_impfindiag = beforst & afterst
       Wend
       
'If Len(Trim(txt_report.Text)) > 7000 Then
'    MsgBox " Immunohistochemistry Max. of 7000 Characters Only.", vbInformation
'    txt_report.SetFocus
'    Exit Sub
'End If
'If Len(Trim(txt_ihmeterial.Text)) > 1000 Then
'    MsgBox " Nature of Material  Max. of 1000 Characters Only.", vbInformation
'    txt_ihmeterial.SetFocus
'    Exit Sub
'End If
'If Len(Trim(txt_impfindiag.Text)) > 2048 Then
'    MsgBox " Impression  Max. of 2048 Characters Only.", vbInformation
'    txt_impfindiag.SetFocus
'    Exit Sub
'End If
'If Len(Trim(txt_impfindiag.Text)) = 0 Then
'    MsgBox " Please Enter Impression. ", vbInformation
'    txt_impfindiag.SetFocus
'    Exit Sub
'End If
       
'rdb.BeginTrans
       assign
'rdb.CommitTrans
       'MsgBox "Record is saved Successfully."
  FLDENABL ("DIS")
  btnenabl ("SAVE")
  Exit Sub
MSG:
rdb.RollbackTrans
MsgBox ERROR(Err)

End Sub

Private Sub cmd_view_Click()
btnenabl ("VIEW")
FLDENABL ("VIEW")
field_clear
If SELDEPTRSFLAG1 = False Then
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    SELDEPTRSFLAG1 = True
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
opt_patient.Value = True
cmb_dept.Enabled = True
viewflag = True
cmb_dept.SetFocus
cmd_print.Enabled = True
txt_positive.Visible = False
txt_remarks.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
btnenabl ("INI")
NEWFLAG = False
viewflag = False
FLDENABL ("DIS")
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
If IHREPORTFLAG = False Then
    IHREPORTFLAG = True
''for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr11", str)

'Set sel_deptrs = sel_deptqr.OpenResultset(2, 3)
   

'' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr11", str)

'sel_moadlityqr(0) = "FZZ"
'Set sel_modalityrs = sel_moadlityqr.OpenResultset(2, 3)

'    'for retriving reqnos
Set req_qr = rdb.CreateQuery("req", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' ORDER BY REQNDT DESC")
'req_qr(0) = "aaa"
'Set req_rs = req_qr.OpenResultset(1, 3)
'
'' FOR ENTER BY, FINALIZED BY
'Set SELEMPLIB_QR = rdb.CreateQuery("SELEMPLIB_QR1", "SELECT * FROM EMPLIB.EMPBASINFO WHERE TITLE='DR.'")
''SELEMPLIB_QR(0) = "AA"
'Set SELEMPLIB_RS = SELEMPLIB_QR.OpenResultset(2, 3)
'SELEMPLIB_RS.MoveFirst
'While Not SELEMPLIB_RS.EOF
'cmb_enteryby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
'cmb_finalby.AddItem SELEMPLIB_RS("EmpTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
'cmb_provby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
'cmb_addendumby.AddItem SELEMPLIB_RS("EMPTITLE") & "$" & SELEMPLIB_RS("EMPCODE")
'cmb_grossby.AddItem SELEMPLIB_RS("emptitle") & "$" & SELEMPLIB_RS("empcode")
'SELEMPLIB_RS.MoveNext
'Wend
'
'
'' for retriveing name,age etc
'
Set SELTA101_QR = rdb.CreateQuery("selta101_qr111", "select * from PATOOLSLIB.ta101 where case_no =?")
'selta101_qr(0) = "aa"
'Set selta101_rs = selta101_qr.OpenResultset(2, 3)
''for view purpose
'STR = "select * from emplib.empbasinfo where empcode =?"
'Set selempbasinfo1_qr = rdb.CreateQuery("selempbasinfo1_qr1", STR)
'selempbasinfo1_qr(0) = "00001"
'Set selempbasinfo1_rs = selempbasinfo1_qr.OpenResultset(2, 3)
'
'
''' for name of  the reporting doctor
'STR = "select emptitle from emplib.empbasinfo"
'Set repdr_qr = rdb.CreateQuery("repdr_qr10", STR)
'Set repdr_rs = repdr_qr.OpenResultset(2, 3)
'repdr_rs.MoveFirst
'    While Not repdr_rs.EOF
'        cmb_reportdoc.AddItem repdr_rs("emptitle")
'        repdr_rs.MoveNext
'    Wend
'    ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1", str)
'    REFDR_QR(0) = "BN/00001"
'    REFDR_QR(1) = "FZZSP010000"
'    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
'    ' FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1", str)
'    SELPATHDREQUISITN_QR(0) = "SSSSS"
'    'SELPATHDREQUISITN_QR(1) = "FZZSP010000"
'    Set SELPATHDREQUISITN_RS = SELPATHDREQUISITN_QR.OpenResultset(2, 3)
''
'    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11", str)
'    selma104_qr(0) = "BN/00001"
'    Set selma104_rs = selma104_qr.OpenResultset(2, 3)
'    'for provisional diagnosis
'
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1", str)
'    selta1012_qr(0) = "BN/00001"
'    Set selta1012_rs = selta1012_qr.OpenResultset(2, 3)

    str = "select * from dislib.dihreport where reqno=?"
    Set seldihreport_qr = rdb.CreateQuery("seldihreport_1qr", str)
'    seldihreport_qr(0) = "zzzzzz"
'    Set seldihreport_rs = seldihreport_qr.OpenResultset(2, 3)


'
    str = "update dislib.dihreport set TESTRESULT=?,REMARKS=? where reqno=? AND TESTORDER=?"
    Set upddihreport_qr = rdb.CreateQuery("upddihreport_qr11", str)

    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1", str)
'    SELRREQUISITN_QR(0) = "BN/00000"
'    SELRREQUISITN_QR(1) = "IH"
'    SELRREQUISITN_QR(2) = "FZZ"
'    SELRREQUISITN_QR(3) = "T"
'    SELRREQUISITN_QR(4) = "F"
'    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'
'
'
   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1", str)
'
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1", str)
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
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1", str)
'   seldrequisitn_qr(0) = "BN/00000"
'   Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'    lbl_Grade.Visible = True
'    cmb_grade.Visible = True
'    txt_addendum.Enabled = False
'    cmb_addendumby.Enabled = False
'    cmb_dept.Enabled = False
'   FLDENABL ("DIS")
'   txt_fullreport.Visible = False

    str = "insert into dislib.dfreport(reqno,ihreport,impression) values(?,?,?)"
    Set insdfreport_qr = rdb.CreateQuery("insdfreport_qr1", str)
    
    str = "update dislib.dfreport set ihreport=?,impression=? where reqno=?"
    Set upddfreport_qr = rdb.CreateQuery("upddfreport_qr1", str)
    
    str = "select * from dislib.dfreport where reqno=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("seldfreport_qr1", str)
    
    str = "select * from dislib.dissec where enterby='YES'AND MODALITY='SP'"
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1", str)
   
   str = "select * from dislib.dissec where PROVby='YES'AND MODALITY='SP'"
   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1", str)
   
   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='SP'"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1", str)
    
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1", str)
    
    str = "UPDATE DISLIB.DREQUISITN SET NMETERIALS=? WHERE REQNO=?"
    Set upddreqnoqr = rdb.CreateQuery("upddreqnoqrIH", str)
    
    str = "SELECT * FROM DISLIB.DIMPRESION WHERE REQNO=?"
    Set seldimpresion_qr = rdb.CreateQuery("seldimpresion_qrIH", str)
    
    str = "INSERT INTO DISLIB.DIMPRESION (REQNO,T,TSIZE,N,M,GRADEDESC,GRADECODE) VALUES(?,?,?,?,?,?,?)"
    Set insdimpresion_qr = rdb.CreateQuery("insdimpresion_qrIH", str)
    
    str = "UPDATE DISLIB.DIMPRESION SET T=?,TSIZE=?,N=?,M=?,GRADEDESC=?,GRADECODE=? WHERE REQNO=?"
    Set upddimpresion_qr = rdb.CreateQuery("upddimpresion_qrIH", str)
    
    str = "SELECT * FROM DISLIB.DGRADE"
    Set SELDGRAGEQR = rdb.CreateQuery("SELDGRAGEQRIH", str)
    
    str = "INSERT INTO RISLIB.DISIMP(REQNO,IMP)VALUES(?,?)"
    Set INSDISIMP_QR = rdb.CreateQuery("INSDISIMP_QRIH", str)
  
    str = "UPDATE RISLIB.DISIMP SET IMP=? WHERE REQNO=?"
    Set UPDDISIMP_QR = rdb.CreateQuery("UPDDISIMP_QRIH", str)
End If
'Msftest.Col = 2
'txt_positive.Width = Msftest.CellWidth
'txt_positive.Height = Msftest.CellHeight
'txt_positive.Left = Msftest.CellLeft + Msftest.Left
'txt_positive.Top = Msftest.CellTop + Msftest.Top
'Msftest.Col = 3
'txt_remarks.Width = Msftest.CellWidth
'txt_remarks.Height = Msftest.CellHeight
'txt_remarks.Left = Msftest.CellLeft + Msftest.Left
'txt_remarks.Top = Msftest.CellTop + Msftest.Top
'txt_positive.Enabled = False
'txt_remarks.Enabled = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Private Sub Form_Unload(cancel As Integer)
Dim RS As rdoResultset
Dim RQ As rdoQuery
For Each RS In rdb.rdoResultsets
    RS.Close
Next

For Each RQ In rdb.rdoQueries
    RQ.Close
Next
End Sub




Private Sub Label1_Click()

End Sub

Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If
If opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
        MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
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
MsgBox ERROR(Err)
End Sub

Private Sub cmb_reqno_Click()
Dim prpathno As String
Dim ADDENDUM As String
On Error GoTo MSG:
If opt_patient.Value = True Then
        If NEWFLAG = True Then
            NEWFLAG = False
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
                If REFDR_RS("recordno") = "1" Then
                LBL_VPP.Caption = "VPP"
                If Trim(REFDR_RS("RECPTNO")) = "VPP" Then
                    MsgBox "Payment is not done"
                    btnenabl ("INI")
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
            
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
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
                TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF = False Then
                    txt_ihmeterial.Text = " " & selpathdrequisitn_rs("NMETERIALS").Value
                Else
                    txt_ihmeterial.Text = ""
                End If
                
                
                'txt_refdoc.Text = SELPATHDREQUISITN_RS("REFDOC")
        '        If IsNull(Trim(SELPATHDREQUISITN_RS("SITE"))) Then
        '            txt_sites.Text = ""
        '        Else
        '            txt_sites.Text = SELPATHDREQUISITN_RS("SITE")
        '        End If
        '        If IsNull(Trim(SELPATHDREQUISITN_RS("LATERALITS"))) Then
        '            txt_laterality.Text = ""
        '        Else
        '            txt_laterality.Text = SELPATHDREQUISITN_RS("LATERALITS")
        '        End If
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
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
            If SELDGRADEFLAG = False Then
                Set seldgraders = SELDGRAGEQR.OpenResultset(1, 3)
                SELDGRADEFLAG = True
            Else
                seldgraders.MoreResults
                seldgraders.Requery
            End If
            If seldgraders.RowCount > 0 Then
                cmb_grade.clear
                While Not seldgraders.EOF
                    cmb_grade.AddItem seldgraders("gradedesc") & "$" & seldgraders("gradecode")
                    seldgraders.MoveNext
                Wend
            End If
            
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
            If DIHREPORTFLAG1 = False Then
                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
                Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
                DIHREPORTFLAG1 = True
            Else
                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
                seldihreport_rs.MoreResults
                seldihreport_rs.Requery
            End If
            If seldihreport_rs.RowCount > 0 Then
                While Not seldihreport_rs.EOF
                    testorder = testorder & Chr(13) & Trim(seldihreport_rs("testorder"))
                    seldihreport_rs.MoveNext
                Wend
                txt_report.Text = Trim(testorder)
            End If
'            I = 1
'            Msftest.Rows = 1
'            If seldihreport_rs.RowCount > 0 Then
'                While Not seldihreport_rs.EOF
'                    Msftest.AddItem I & vbTab & Trim(seldihreport_rs("testorder"))
'                    seldihreport_rs.MoveNext
'                    I = I + 1
'                Wend
'            End If
            cmb_t.clear
            cmb_m.clear
            CMB_N.clear
        For I = 0 To 5
            cmb_t.AddItem I
            cmb_m.AddItem I
            CMB_N.AddItem I
        Next I
            cmb_m.AddItem "X"
            CMB_N.AddItem "X"
        If seldissecentrflag1 = False Then
            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
            seldissecentrflag1 = True
        Else
            seldissecentr1_rs.MoreResults
            seldissecentr1_rs.Requery
        End If
        If seldissecentr1_rs.RowCount > 0 Then
            cmb_enterby.clear
            While Not seldissecentr1_rs.EOF
                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
                seldissecentr1_rs.MoveNext
            Wend
        End If
        If seldissecprovflag1 = False Then
            Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
            seldissecprovflag1 = True
        Else
            seldissecprov1_rs.MoreResults
            seldissecprov1_rs.Requery
        End If
        If seldissecprov1_rs.RowCount > 0 Then
            cmb_provisionalby.clear
            While Not seldissecprov1_rs.EOF
                cmb_provisionalby.AddItem Trim(seldissecprov1_rs("CCNAME"))
                seldissecprov1_rs.MoveNext
            Wend
        End If
        If seldissecfinflag = False Then
            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
            seldissecfinflag1 = True
        Else
            seldissecfin1_rs.MoreResults
            seldissecfin1_rs.Requery
        End If
        If seldissecfin1_rs.RowCount > 0 Then
            cmb_finalized.clear
            While Not seldissecfin1_rs.EOF
                cmb_finalized.AddItem Trim(seldissecfin1_rs("CCNAME"))
                seldissecfin1_rs.MoveNext
            Wend
        End If
            
            
            
        End If
            If viewflag = True Then
        
           
            
            viewflag = False
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
                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalized.Text = ""
                Else
                    cmb_finalized.Text = REFDR_RS("FREPORTBY")
                End If
                    If REFDR_RS("recordno") = "1" Then
                    LBL_VPP.Caption = "VPP"
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
            
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_status.Text = "OP"
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
                TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
                If selpathdrequisitn_rs.EOF = False Then
                      txt_ihmeterial.Text = selpathdrequisitn_rs("NMETERIALS")
                Else
                    txt_ihmeterial.Text = " "
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
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
        
            If seldfreportflag = False Then
                SELDFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
                seldfreportflag = True
            Else
                SELDFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                SELDFREPORT_RS.MoreResults
                SELDFREPORT_RS.Requery
            End If
            If SELDFREPORT_RS.RowCount > 0 Then
                If SELDFREPORT_RS.EOF = False Then
                    txt_report.Text = SELDFREPORT_RS("ihreport").Value
                Else
                    txt_report.Text = ""
                End If
                If SELDFREPORT_RS.EOF = False Then
                    txt_impfindiag.Text = " " & SELDFREPORT_RS("iMPRESSION").Value
                Else
                    txt_impfindiag.Text = ""
                End If
            
            End If
            If SELDIMPRESIONFLAG1 = False Then
                seldimpresion_qr(0) = req_chglen(cmb_reqno.Text)
                Set seldimpresion_rs = seldimpresion_qr.OpenResultset(1, 3)
                SELDIMPRESIONFLAG1 = True
            Else
                seldimpresion_qr(0) = req_chglen(cmb_reqno.Text)
                seldimpresion_rs.MoreResults
                seldimpresion_rs.Requery
            End If
            If seldimpresion_rs.RowCount > 0 Then
                If IsNull(Trim(seldimpresion_rs("gradedesc"))) Then
                    
                Else
                cmb_grade.AddItem seldimpresion_rs("gradedesc")
                cmb_grade.Text = seldimpresion_rs("gradedesc")
                End If
                cmb_t.Text = seldimpresion_rs("t")
                txt_size.Text = seldimpresion_rs("tsize")
                CMB_N.Text = seldimpresion_rs("n")
                cmb_m.AddItem seldimpresion_rs("m")
                cmb_m.Text = seldimpresion_rs("m")
            End If
            
'            If DIHREPORTFLAG1 = False Then
'                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
'                Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
'                DIHREPORTFLAG1 = True
'            Else
'                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
'                seldihreport_rs.MoreResults
'                seldihreport_rs.Requery
'            End If
'            Msftest.Rows = 1
'            If seldihreport_rs.RowCount > 0 Then
'            I = 1
'                While Not seldihreport_rs.EOF
'                    Msftest.AddItem I & vbTab & Trim(seldihreport_rs("TESTORDER")) & vbTab & Trim(seldihreport_rs("TESTRESULT")) & vbTab & Trim(seldihreport_rs("REMARKS"))
'                    I = I + 1
'                    seldihreport_rs.MoveNext
'                Wend
'            End If
        
            If SELDREQUISITNFLAG1 = False Then
                seldrequisitn_qr(0) = MaskEdBox1.FormattedText
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            Else
                seldrequisitn_qr(0) = MaskEdBox1.FormattedText
                seldrequisitn_rs.MoreResults
                seldrequisitn_rs.Requery
            End If
            If seldrequisitn_rs.RowCount > 0 Then
            'txt_refdoc.Text = seldrequisitn_rs("REFDOC")
             prpathno = ""
             While Not seldrequisitn_rs.EOF
                prpathno = prpathno & " " & seldrequisitn_rs("pathno")
                'CMB_PATHNO.AddItem seldrequisitn_rs("PATHNO") & "$" & seldrequisitn_rs("REQNO")
                seldrequisitn_rs.MoveNext
             Wend
                'txt_prepathreqno.Text = Trim(prpathno)
            End If
                
            If seldissecfinflag = False Then
            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
            seldissecfinflag1 = True
        Else
            seldissecfin1_rs.MoreResults
            seldissecfin1_rs.Requery
        End If
        If seldissecfin1_rs.RowCount > 0 Then
            cmb_finalized.clear
            While Not seldissecfin1_rs.EOF
                cmb_finalized.AddItem Trim(seldissecfin1_rs("CCNAME"))
                seldissecfin1_rs.MoveNext
            Wend
            cmb_finalized.Text = REFDR_RS("FREPORTBY")
        End If
            cmb_reqno.Enabled = False
            cmb_dept.Enabled = False
            cmb_modality.Enabled = False
            MaskEdBox1.Enabled = False
            
            'Msftest.Enabled = True
        End If
End If
If opt_staff.Value = True Then
     If NEWFLAG = True Then
            NEWFLAG = False
            If SELDREQSTAFFD1FLAG = False Then
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = MaskEdBox1.FormattedText
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
                txt_age = DateDiff("YYYY", SELDREQSTAFFDRS("BIRTH_DATE"), CStr(DATE11))
                txt_pridoc.Text = ""
                txt_category.Text = SELDREQSTAFFDRS("RELATION")
            End If
            If REFDRFLAG1 = True Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
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
          
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
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
                TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
        '        If IsNull(Trim(SELPATHDREQUISITN_RS("SITE"))) Then
        '            txt_sites.Text = ""
        '        Else
        '            txt_sites.Text = SELPATHDREQUISITN_RS("SITE")
        '        End If
        '        If IsNull(Trim(SELPATHDREQUISITN_RS("LATERALITS"))) Then
        '            txt_laterality.Text = ""
        '        Else
        '            txt_laterality.Text = SELPATHDREQUISITN_RS("LATERALITS")
        '        End If
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
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
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
'            If DIHREPORTFLAG1 = False Then
'                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
'                Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
'                DIHREPORTFLAG1 = True
'            Else
'                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
'                seldihreport_rs.MoreResults
'                seldihreport_rs.Requery
'            End If
'            I = 1
'            Msftest.Rows = 1
'            If seldihreport_rs.RowCount > 0 Then
'                While Not seldihreport_rs.EOF
'                    Msftest.AddItem I & vbTab & Trim(seldihreport_rs("testorder"))
'                    seldihreport_rs.MoveNext
'                    I = I + 1
'                Wend
'            End If
            If seldissecentrflag1 = False Then
            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
            seldissecentrflag1 = True
        Else
            seldissecentr1_rs.MoreResults
            seldissecadd1_rs.Requery
        End If
        If seldissecentr1_rs.RowCount > 0 Then
            cmb_enterby.clear
            While Not seldissecentr1_rs.EOF
                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
                seldissecentr1_rs.MoveNext
            Wend
        End If
        If seldissecprovflag1 = False Then
            Set seldissecprov1_rs = seldissecprov1_qr.OpenResultset(1, 3)
            seldissecprovflag1 = True
        Else
            seldissecprov1_rs.MoreResults
            seldissecprov1_rs.Requery
        End If
        If seldissecprov1_rs.RowCount > 0 Then
            cmb_provisionalby.clear
            While Not seldissecprov_rs.EOF
                cmb_provisionalby.AddItem Trim(seldissecprov1_rs("CCNAME"))
                seldissecprov_rs.MoveNext
            Wend
        End If
        If seldissecfinflag = False Then
            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
            seldissecfinflag1 = True
        Else
            seldissecfin1_rs.MoreResults
            seldissecfin1_rs.Requery
        End If
        If seldissecfin1_rs.RowCount > 0 Then
            cmb_finalized.clear
            While Not seldissecfin1_rs.EOF
                cmb_finalized.AddItem Trim(seldissecfin1_rs("CCNAME"))
                seldissecfin_rs.MoveNext
            Wend
        End If
            
        
        End If
        If viewflag = True Then
        
           
            
            viewflag = False
            If SELDREQSTAFFD1FLAG = False Then
                SELDREQSTAFFDQR(0) = req_chglen(cmb_reqno)
                SELDREQSTAFFDQR(1) = MaskEdBox1.FormattedText
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
                txt_age = DateDiff("YYYY", SELDREQSTAFFDRS("BIRTH_DATE"), CStr(DATE11))
                txt_pridoc.Text = ""
                txt_category.Text = SELDREQSTAFFDRS("RELATION")
            End If
        
            If REFDRFLAG1 = True Then
                REFDR_QR(0) = MaskEdBox1.FormattedText
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
                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
                    ADDENDUMFALG = True
                End If
                cmb_enterby.Text = REFDR_RS("ENTERBY")
                cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalized.Text = ""
                Else
                    cmb_finalized.Text = REFDR_RS("FREPORTBY")
                End If
            End If
            
            
            
          
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = MaskEdBox1.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = MaskEdBox1.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
                txt_status.Text = "OP"
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
                txt_ihmeterial.Text = selpathdrequisitn_rs("NMETERIALS")
                txt_refdoc.Text = selpathdrequisitn_rs("REFDOC")
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
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
        
            If seldfreportflag = False Then
                SELDFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
                seldfreportflag = True
            Else
                SELDFREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                SELDFREPORT_RS.MoreResults
                SELDFREPORT_RS.Requery
            End If
            If SELDFREPORT_RS.RowCount > 0 Then
                If SELDFREPORT_RS.EOF = False Then
                    txt_report.Text = SELDFREPORT_RS("ihreport").Value
                End If
            Else
                txt_report.Text = ""
            End If
            
        
''            If DIHREPORTFLAG1 = False Then
''                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
''                Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
''                DIHREPORTFLAG1 = True
''            Else
''                seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
''                seldihreport_rs.MoreResults
''                seldihreport_rs.Requery
''            End If
''            Msftest.Rows = 1
''            If seldihreport_rs.RowCount > 0 Then
''            I = 1
''                While Not seldihreport_rs.EOF
''                    Msftest.AddItem I & vbTab & Trim(seldihreport_rs("TESTORDER")) & vbTab & Trim(seldihreport_rs("TESTRESULT")) & vbTab & Trim(seldihreport_rs("REMARKS"))
''                    I = I + 1
''                    seldihreport_rs.MoveNext
''                Wend
''            End If
        
                   
            If SELDREQUISITNFLAG1 = False Then
                seldrequisitn_qr(0) = MaskEdBox1.FormattedText
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
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
                'txt_prepathreqno.Text = Trim(prpathno)
            End If
                
            
            cmb_reqno.Enabled = False
            cmb_dept.Enabled = False
            cmb_modality.Enabled = False
            MaskEdBox1.Enabled = False
            
            Msftest.Enabled = True
        End If

End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub










Private Sub txt_fullreport_GotFocus()
'On Error GoTo m14:
'w.Visible = False
'Set d = w.ActiveDocument
'd.Content = ""
'd.Content = txt_fullreport
'w.Caption = "Full Report"
'w.WindowState = wdWindowStateMaximize
'd.Content = txt_fullreport
'w.Visible = True
'w.Activate
'm14:
On Error Resume Next:
'Set w = CreateObject("word.Application")
'w.Documents.Add ("c:\risreports\REPORT.dot")
'w.Caption = "Full report"
'w.WindowState = wdWindowStateMaximize
'Set d = w.ActiveDocument
'd.Content = txt_fullreport.Text
'w.Visible = True
'w.Activate
End Sub

Private Sub txt_fullreport_LostFocus()
'On Error GoTo m15:
'Set d = w.ActiveDocument
'txt_fullreport.Text = d.Content
''flagimp = False
'd.Content = ""
'd.SaveAs "C:\fronzen.DOC"
'd.Close
'w.Quit
''W1.Visible = False
'Exit Sub
'm15:
End Sub


















Public Sub assign()
On Error GoTo MSG:
Dim I As Integer
rdb.BeginTrans
Dim MICROSCOPIC As String
If flagsave = 1 Then
   flagsave = 0
'    For I = 0 To Msftest.Rows - 1
'        upddihreport_qr(0) = Msftest.TextMatrix(I, 2)
'        upddihreport_qr(1) = Msftest.TextMatrix(I, 3)
'        upddihreport_qr(2) = req_chglen(cmb_reqno.Text)
'        upddihreport_qr(3) = Msftest.TextMatrix(I, 1)
'        upddihreport_qr.Execute
'    Next I
    insdfreport_qr(0) = req_chglen(cmb_reqno.Text)
    If Len(Trim(txt_report.Text)) > 0 Then
        insdfreport_qr(1).Value = Null
        insdfreport_qr(1) = Trim(txt_report.Text)
    Else
        insdfreport_qr(1).Value = Null
        insdfreport_qr(1) = " "
    End If
    If Len(Trim(txt_impfindiag.Text)) > 0 Then
        insdfreport_qr(2).Value = Null
        insdfreport_qr(2) = Trim(txt_impfindiag.Text)
    Else
        insdfreport_qr(2).Value = Null
        insdfreport_qr(2) = " "
    End If
    insdfreport_qr.Execute
    
    updrrequisitn_qr(0) = cmb_enterby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provisionalby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalized.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    If Len(Trim(txt_ihmeterial.Text)) > 0 Then
        upddreqnoqr(0) = Null
        upddreqnoqr(0) = Trim(txt_ihmeterial.Text)
    Else
        upddreqnoqr(0) = Null
        upddreqnoqr(0) = " "
    End If
    upddreqnoqr(1) = Trim(req_chglen(cmb_reqno.Text))
    upddreqnoqr.Execute
    
    insdimpresion_qr(0) = req_chglen(cmb_reqno.Text)
    insdimpresion_qr(1) = Trim(cmb_t.Text)
    insdimpresion_qr(2) = Trim(txt_size.Text)
    insdimpresion_qr(3) = Trim(CMB_N.Text)
    insdimpresion_qr(4) = Trim(cmb_m.Text)
    I = InStr(cmb_grade.Text, "$")
    If I = 0 Then
        insdimpresion_qr(5) = ""
        insdimpresion_qr(6) = ""
    Else
        insdimpresion_qr(5) = Mid(cmb_grade.Text, 1, I - 1)
        insdimpresion_qr(6) = Mid(cmb_grade.Text, I + 1, 2)
    End If
    insdimpresion_qr.Execute
    
     '**** INSERTING INTO DISIMP
        If Len(Trim(txt_impfindiag.Text)) > 0 Then
        INSDISIMP_QR(0) = req_chglen(cmb_reqno.Text)
        INSDISIMP_QR(1) = Null
        INSDISIMP_QR(1) = Trim(txt_impfindiag.Text)
        INSDISIMP_QR.Execute
        End If
  '**** INSERTING INTO DISIMP
    
End If
      
    
If flagsave = 2 Then
    flagsave = 0
 
'    For I = 0 To Msftest.Rows - 1
'        upddihreport_qr(0) = Msftest.TextMatrix(I, 2)
'        upddihreport_qr(1) = Msftest.TextMatrix(I, 3)
'        upddihreport_qr(2) = req_chglen(cmb_reqno.Text)
'        upddihreport_qr(3) = Msftest.TextMatrix(I, 1)
'        upddihreport_qr.Execute
'    Next I
    If Len(Trim(txt_report.Text)) > 0 Then
        upddfreport_qr(0).Value = Null
        upddfreport_qr(0) = Trim(txt_report)
    Else
        upddfreport_qr(0).Value = Null
        upddfreport_qr(0) = " "
    End If
    If Len(Trim(txt_impfindiag.Text)) > 0 Then
        upddfreport_qr(1).Value = Null
        upddfreport_qr(1) = Trim(txt_impfindiag)
    Else
        upddfreport_qr(1).Value = Null
        upddfreport_qr(1) = " "
    End If
    upddfreport_qr(2) = req_chglen(cmb_reqno.Text)
    upddfreport_qr.Execute
    
    updrrequisitn_qr(0) = cmb_enterby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = cmb_provisionalby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalized.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
         
    If Len(Trim(txt_ihmeterial.Text)) > 0 Then
        upddreqnoqr(0) = Null
        upddreqnoqr(0) = Trim(txt_ihmeterial.Text)
    Else
        upddreqnoqr(0) = Null
        upddreqnoqr(0) = " "
    End If
    upddreqnoqr(1) = Trim(req_chglen(cmb_reqno.Text))
    upddreqnoqr.Execute
    
    upddimpresion_qr(0) = cmb_t.Text
    upddimpresion_qr(1) = txt_size.Text
    upddimpresion_qr(2) = CMB_N.Text
    upddimpresion_qr(3) = cmb_m.Text
    I = InStr(cmb_reqno.Text, "$")
    If I = 0 Then
        If Len(Trim(cmb_grade.Text)) > 0 Then
            upddimpresion_qr(4) = cmb_grade.Text
            upddimpresion_qr(5) = 0
        Else
            upddimpresion_qr(4) = ""
            upddimpresion_qr(5) = ""
        End If
    Else
        upddimpresion_qr(4) = Mid(cmb_grade.Text, 1, I - 1)
        upddimpresion_qr(5) = Mid(cmb_grade.Text, I + 1, 2)
    End If
    upddimpresion_qr(6) = req_chglen(cmb_reqno.Text)
    upddimpresion_qr.Execute
    
        '**** UPDATE THE DISIMP
        If Len(Trim(txt_impfindiag.Text)) > 0 Then
        UPDDISIMP_QR(0) = Null
        UPDDISIMP_QR(0) = Trim(txt_impfindiag.Text)
        UPDDISIMP_QR(1) = req_chglen(cmb_reqno.Text)
        UPDDISIMP_QR.Execute
        End If
        
        '*** UPDATE THE DISIMP
    
End If
rdb.CommitTrans
MsgBox "Record is saved Successfully."
Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
    
    
End Sub

Public Sub field_clear()
Dim obj As Object
On Error GoTo MSG:
cmb_dept.ListIndex = -1

cmb_modality.ListIndex = -1

cmb_reqno.ListIndex = -1

MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
MaskEdBox1.Mask = ">??/#####"
MaskEdBox1.Text = ""
txt_impfindiag.Text = ""
cmb_t.Text = ""
txt_size.Text = ""
CMB_N.Text = ""
cmb_m.ListIndex = -1
cmb_grade.Text = ""
cmb_grade.ListIndex = -1
LBL_VPP.Caption = ""
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In imhreport.Controls
If TypeOf obj Is TextBox Then
obj = ""
cmb_enterby.Text = " "
cmb_finalized.Text = " "
cmb_provisionalby.Text = " "
End If
Next
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then

    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_reportdoc.Enabled = True
    cmb_dept.Enabled = True
    MaskEdBox1.Enabled = True
'    txt_positive.Enabled = True
'    txt_remarks.Enabled = True
    Msftest.Enabled = True
    txt_procedure.Enabled = False
    cmb_enterby.Enabled = True
    cmb_provisionalby.Enabled = True
    cmb_finalized.Enabled = True
    txt_report.Enabled = True
    txt_ihmeterial.Enabled = True
    txt_impfindiag.Enabled = True
    cmb_t.Enabled = True
    txt_size.Enabled = True
    CMB_N.Enabled = True
    cmb_m.Enabled = True
    cmb_grade.Enabled = True
End If
If str1 = "DIS" Then

    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_reportdoc.Enabled = False
    cmb_dept.Enabled = False
    MaskEdBox1.Enabled = False
    txt_positive.Enabled = False
    txt_remarks.Enabled = False
    Msftest.Enabled = False
    txt_procedure.Enabled = False
    cmb_enterby.Enabled = False
    cmb_provisionalby.Enabled = False
    cmb_finalized.Enabled = False
    txt_report.Enabled = False
    txt_ihmeterial.Enabled = False
    txt_impfindiag.Enabled = False
    cmb_t.Enabled = False
    txt_size.Enabled = False
    CMB_N.Enabled = False
    cmb_m.Enabled = False
    cmb_grade.Enabled = False
End If
If str1 = "VIEW" Then

    cmb_modality.Enabled = False
    cmb_reportdoc.Enabled = False
    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
    txt_positive.Enabled = FASLE
    txt_remarks.Enabled = False
    Msftest.Enabled = False
    txt_procedure.Enabled = False
    cmb_enterby.Enabled = False
    cmb_provisionalby.Enabled = False
    cmb_finalized.Enabled = False
    txt_report.Enabled = False
    txt_ihmeterial.Enabled = False
    txt_impfindiag.Enabled = False
    cmb_t.Enabled = False
    txt_size.Enabled = False
    CMB_N.Enabled = False
    cmb_m.Enabled = False
    cmb_grade.Enabled = False
End If
If str1 = "MODIFY" Then

    cmb_reqno.Enabled = True
    cmb_modality.Enabled = True
    MaskEdBox1.Enabled = True
    cmb_dept.Enabled = True
'    txt_positive.Enabled = True
'    txt_remarks.Enabled = True
    Msftest.Enabled = True
    txt_procedure.Enabled = False
    cmb_enterby.Enabled = True
    cmb_provisionalby.Enabled = True
    cmb_finalized.Enabled = True
    txt_report.Enabled = True
    txt_ihmeterial.Enabled = True
    txt_impfindiag.Enabled = True
    cmb_t.Enabled = True
    txt_size.Enabled = True
    CMB_N.Enabled = True
    cmb_m.Enabled = True
    cmb_grade.Enabled = True
End If
If str1 = "ADDENDUM" Then

End If
End Sub

Private Sub Msftest_Click()
'        txt_positive = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
'        txt_positive.SetFocus
'        txt_positive.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'        If Trim(MSFlexGrid1.Text) = "COLLECTED" Then
'        chk_collected.Value = 1
'        Else
'        chk_collected.Value = 0
'        End If
End Sub

Private Sub Opt_patient_Click()
lbl_caseno.Caption = "Case No."
MaskEdBox1.Mask = "??>/#####"
End Sub

Private Sub opt_staff_Click()
lbl_caseno.Caption = "CC No."
MaskEdBox1.Mask = "######"
End Sub

Private Sub txt_ihmeterial_GotFocus()
On Error GoTo m7:
WIHMET.Visible = False
Set DIHMET = WIHMET.ActiveDocument
DIHMET.Content.Text = ""
DIHMET.Content.Text = txt_ihmeterial.Text
WIHMET.Caption = "Nature of Meterial"
WIHMET.WindowState = wdWindowStateMaximize
DIHMET.Content.Text = txt_ihmeterial.Text
WIHMET.Visible = True
WIHMET.Activate
m7:
On Error Resume Next:
Set WIHMET = CreateObject("word.Application")
WIHMET.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
WIHMET.Caption = "Nature of Meterial"
WIHMET.WindowState = wdWindowStateMaximize
Set DIHMET = WIHMET.ActiveDocument
DIHMET.Content.Text = txt_ihmeterial.Text
WIHMET.Visible = True
End Sub

Private Sub txt_ihmeterial_LostFocus()
On Error GoTo m8:
Set DIHMET = WIHMET.ActiveDocument
txt_ihmeterial.Text = DIHMET.Content
txt_ihmeterial.Text = Trim(txt_ihmeterial.Text)
'flagimp = False
DIHMET.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DIHMET.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DIHMET.Close (wdDoNotSaveChanges)
WIHMET.quit
'W1.Visible = False
Exit Sub
m8:
MsgBox Err.Description
End Sub

Private Sub txt_impfindiag_GotFocus()
On Error GoTo m7:
WIMP.Visible = False
Set DIMP = WIMP.ActiveDocument
DIMP.Content.Text = ""
DIMP.Content.Text = txt_impfindiag.Text
WIMP.Caption = "Impression"
WIMP.WindowState = wdWindowStateMaximize
DIMP.Content.Text = txt_impfindiag.Text
WIMP.Visible = True
WIMP.Activate
m7:
On Error Resume Next:
Set WIMP = CreateObject("word.Application")
WIMP.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
WIMP.Caption = "Impression"
WIMP.WindowState = wdWindowStateMaximize
Set DIMP = WIMP.ActiveDocument
DIMP.Content.Text = txt_impfindiag.Text
WIMP.Visible = True
End Sub

Private Sub txt_impfindiag_LostFocus()
On Error GoTo m8:
Set DIMP = WIMP.ActiveDocument
txt_impfindiag.Text = DIMP.Content
'flagimp = False
DIMP.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DIMP.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DIMP.Close (wdDoNotSaveChanges)
WIMP.quit
'W1.Visible = False
Exit Sub
m8:
MsgBox Err.Description
End Sub

'Private Sub txt_positive_Change()
'Msftest.TextMatrix(Msftest.Row, 2) = txt_positive.Text
'End Sub

'Private Sub txt_positive_GotFocus()
'txt_positive.Text = Msftest.TextMatrix(Msftest.Row, 2)
'End Sub

'Private Sub txt_positive_LostFocus()
'    txt_remarks.SetFocus
'End Sub

'Private Sub txt_remarks_Change()
'Msftest.TextMatrix(Msftest.Row, 3) = txt_remarks.Text
'End Sub
'
'Private Sub txt_remarks_GotFocus()
'txt_remarks.Text = Msftest.TextMatrix(Msftest.Row, 3)
'End Sub

'Private Sub txt_remarks_LostFocus()
'On Error GoTo msg:
'If Len(Trim(txt_positive)) = 0 Then
'     cmd_save.SetFocus
'     Exit Sub
'
'End If
'
'If Msftest.Row = Msftest.Rows - 1 Then
'        Msftest.Rows = Msftest.Rows + 1
'        Msftest.TextMatrix(Msftest.Rows - 1, 0) = Msftest.Rows - 1
'End If
'If Msftest.Row < Msftest.Rows - 1 Then
'        Msftest.Row = Msftest.Row + 1
'        txt_remarks = Msftest.TextMatrix(Msftest.Row, 3)
'        txt_positive = Msftest.TextMatrix(Msftest.Row, 2)
'        txt_positive.Visible = True
'        txt_remarks.Visible = True
'        txt_positive.SetFocus
'        txt_remarks.Top = Msftest.CellTop + Msftest.Top
'        txt_positive.Top = Msftest.CellTop + Msftest.Top
'End If
'Exit Sub
'MSG:
'MsgBox Err.Description
'End Sub


Private Sub txt_report_GotFocus()
On Error GoTo m7:
w_immun.Visible = False
Set d_immun = w_immun.ActiveDocument
d_immun.Content.Text = ""
d_immun.Content.Text = txt_report.Text
w_immun.Caption = "Immuno"
w_immun.WindowState = wdWindowStateMaximize
d_immun.Content.Text = txt_report.Text
w_immun.Visible = True
w_immun.Activate
m7:
On Error Resume Next:
Set w_immun = CreateObject("word.Application")
w_immun.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w_immun.Caption = "Immuno"
w_immun.WindowState = wdWindowStateMaximize
Set d_immun = w_immun.ActiveDocument
d_immun.Content.Text = txt_report.Text
w_immun.Visible = True
End Sub

Private Sub txt_report_LostFocus()
On Error GoTo m8:
Set d_immun = w_immun.ActiveDocument
txt_report.Text = d_immun.Content
'flagimp = False
d_immun.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_immun.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_immun.Close (wdDoNotSaveChanges)
w_immun.quit
'W1.Visible = False
Exit Sub
m8:
MsgBox Err.Description
End Sub
